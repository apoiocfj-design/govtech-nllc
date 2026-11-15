# -*- coding: utf-8 -*-
"""
MÓDULO DE GERAÇÃO DE RELATÓRIOS OFICIAIS (WORD .DOCX E PDF A4)
Padrão Marinha do Brasil e Instrução Normativa SEGES/ME nº 65/2021 (Lei nº 14.133/2021)
Gera em memória (BytesIO) relatórios completos e prontos para instrução processual.
"""

import io
import os
from datetime import datetime
from typing import Dict, Any

# python-docx
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn

# reportlab
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import cm, mm
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import (
    SimpleDocTemplate,
    Paragraph,
    Spacer,
    Table,
    TableStyle,
    Image as RLImage,
    KeepTogether,
    HRFlowable,
)

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
BRASAO_PATH = os.path.join(BASE_DIR, "brasao_oficial.png")
if not os.path.exists(BRASAO_PATH):
    BRASAO_PATH = os.path.join(BASE_DIR, "..", "frontend", "brasao_transp_0.png")


def _format_brl(val: float) -> str:
    if val is None:
        return "R$ 0,00"
    return f"R$ {val:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def _get_mes_extenso(mes: int) -> str:
    meses = [
        "janeiro", "fevereiro", "março", "abril", "maio", "junho",
        "julho", "agosto", "setembro", "outubro", "novembro", "dezembro"
    ]
    return meses[mes - 1] if 1 <= mes <= 12 else ""


# =============================================================================
# 1. GERADOR DO RELATÓRIO EM MICROSOFT WORD (.DOCX)
# =============================================================================

def set_cell_background(cell, fill_hex):
    """Define a cor de fundo de uma célula da tabela no python-docx."""
    tcPr = cell._tc.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>')
    tcPr.append(shd)


def set_cell_margins(cell, top=100, bottom=100, left=150, right=150):
    """Define margens internas (padding) para a célula."""
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = parse_xml(
        f'<w:tcMar {nsdecls("w")}>'
        f'<w:top w:w="{top}" w:type="dxa"/>'
        f'<w:bottom w:w="{bottom}" w:type="dxa"/>'
        f'<w:left w:w="{left}" w:type="dxa"/>'
        f'<w:right w:w="{right}" w:type="dxa"/>'
        f'</w:tcMar>'
    )
    tcPr.append(tcMar)


def build_in65_docx_bytes(
    report: Dict[str, Any],
    quantity: int,
    om_name: str = "CAPITANIA FLUVIAL DE JUAZEIRO",
    elaborador_nome: str = "ENCARREGADO DO SETOR DE OBTENÇÃO",
    ordenador_nome: str = "ORDENADOR DE DESPESAS",
) -> bytes:
    """Gera o arquivo Word (.docx) formal com Brasão Oficial e padrão visual Marinha do Brasil."""
    doc = docx.Document()

    # Configuração de Margens da Página (2 cm nas laterais, 2 cm superior/inferior)
    for section in doc.sections:
        section.top_margin = Inches(0.8)
        section.bottom_margin = Inches(0.8)
        section.left_margin = Inches(0.8)
        section.right_margin = Inches(0.8)

    # CABEÇALHO OFICIAL: BRASÃO
    if os.path.exists(BRASAO_PATH):
        p_img = doc.add_paragraph()
        p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_img.paragraph_format.space_before = Pt(0)
        p_img.paragraph_format.space_after = Pt(4)
        run_img = p_img.add_run()
        run_img.add_picture(BRASAO_PATH, width=Inches(1.05))

    # TEXTO DO CABEÇALHO MILITAR / GOVERNAMENTAL
    p_head = doc.add_paragraph()
    p_head.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_head.paragraph_format.line_spacing = 1.05
    p_head.paragraph_format.space_after = Pt(12)

    lines = [
        ("REPÚBLICA FEDERATIVA DO BRASIL", 10.5, True),
        ("MINISTÉRIO DA DEFESA", 10, True),
        ("COMANDO DA MARINHA", 10, True),
        (om_name.upper(), 9.5, False),
        ("SETOR DE OBTENÇÃO E CONTRATAÇÕES", 9, False),
    ]
    for text, sz, bold in lines:
        r = p_head.add_run(f"{text}\n")
        r.font.name = "Calibri"
        r.font.size = Pt(sz)
        r.font.bold = bold
        r.font.color.rgb = RGBColor(15, 23, 42)

    # TÍTULO PRINCIPAL DO DOCUMENTO
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_title.paragraph_format.space_before = Pt(4)
    p_title.paragraph_format.space_after = Pt(4)
    r_tit = p_title.add_run("RELATÓRIO DE PESQUISA DE PREÇOS\n")
    r_tit.font.name = "Calibri"
    r_tit.font.size = Pt(12)
    r_tit.font.bold = True
    r_tit.font.color.rgb = RGBColor(15, 23, 42)

    r_sub = p_title.add_run(
        "Conforme o Art. 23 da Lei nº 14.133/2021 e a Instrução Normativa SEGES/ME nº 65/2021"
    )
    r_sub.font.name = "Calibri"
    r_sub.font.size = Pt(9.5)
    r_sub.font.italic = True
    r_sub.font.color.rgb = RGBColor(71, 85, 105)

    doc.add_paragraph().paragraph_format.space_after = Pt(4)

    # DADOS BÁSICOS
    item_code = report.get("catmat_catser_code", "—")
    item_desc = report.get("item_description", "—")
    unit_str = report.get("target_supply_unit", "UN")
    unit_price = float(report.get("final_estimated_unit_price", 0.0))
    total_price = quantity * unit_price
    sm = report.get("sanitized_metrics", {})

    # SEÇÃO 1: CARACTERIZAÇÃO DO ITEM DEMANDADO
    p_s1 = doc.add_paragraph()
    p_s1.paragraph_format.space_before = Pt(8)
    p_s1.paragraph_format.space_after = Pt(4)
    r_s1 = p_s1.add_run("1. IDENTIFICAÇÃO DO OBJETO E PREÇO ESTIMADO")
    r_s1.font.bold = True
    r_s1.font.size = Pt(10.5)

    t_item = doc.add_table(rows=5, cols=2)
    t_item.alignment = WD_TABLE_ALIGNMENT.CENTER
    t_item.autofit = False

    t_data = [
        ("Código do Catálogo (CATMAT/CATSER):", f"{item_code}"),
        ("Descrição Detalhada do Objeto:", f"{item_desc}"),
        ("Unidade de Fornecimento & Quantidade:", f"{unit_str} | Quantidade solicitada: {quantity}"),
        ("Preço Unitário Estimado Homologável:", f"{_format_brl(unit_price)}"),
        ("Valor Total Estimado da Contratação:", f"{_format_brl(total_price)}"),
    ]

    col_widths = (Inches(2.4), Inches(4.5))
    for r_idx, (k, v) in enumerate(t_data):
        row = t_item.rows[r_idx]
        c0, c1 = row.cells[0], row.cells[1]
        c0.width, c1.width = col_widths[0], col_widths[1]
        set_cell_margins(c0, 80, 80, 100, 100)
        set_cell_margins(c1, 80, 80, 100, 100)
        set_cell_background(c0, "F1F5F9")
        set_cell_background(c1, "FFFFFF" if r_idx < 4 else "FEF3C7")

        p0 = c0.paragraphs[0]
        p0.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        r0 = p0.add_run(k)
        r0.font.bold = True
        r0.font.size = Pt(9)
        r0.font.name = "Calibri"

        p1 = c1.paragraphs[0]
        p1.alignment = WD_ALIGN_PARAGRAPH.LEFT
        r1 = p1.add_run(v)
        r1.font.bold = (r_idx >= 3)
        r1.font.size = Pt(9.5 if r_idx < 4 else 10.5)
        r1.font.name = "Calibri"
        if r_idx == 4:
            r1.font.color.rgb = RGBColor(146, 64, 14)

    # SEÇÃO 2: FONTES CONSULTADAS E COTAÇÕES
    p_s2 = doc.add_paragraph()
    p_s2.paragraph_format.space_before = Pt(12)
    p_s2.paragraph_format.space_after = Pt(4)
    r_s2 = p_s2.add_run("2. COTAÇÕES COLETADAS NAS COMPRAS PÚBLICAS SIMILARES (Art. 5º da IN 65/2021)")
    r_s2.font.bold = True
    r_s2.font.size = Pt(10.5)

    valid_quotes = report.get("valid_quotes", [])
    quotes_to_show = valid_quotes[:10]

    t_quotes = doc.add_table(rows=1 + len(quotes_to_show), cols=5)
    t_quotes.alignment = WD_TABLE_ALIGNMENT.CENTER
    q_widths = (Inches(0.6), Inches(1.5), Inches(1.3), Inches(2.2), Inches(1.3))

    q_headers = ["Nº", "UASG / Órgão", "Compra nº", "Fornecedor Vencedor", "Preço Unit."]
    hdr_row = t_quotes.rows[0]
    for c_idx, h in enumerate(q_headers):
        cell = hdr_row.cells[c_idx]
        cell.width = q_widths[c_idx]
        set_cell_background(cell, "1E293B")
        set_cell_margins(cell, 80, 80, 80, 80)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(h)
        r.font.bold = True
        r.font.size = Pt(8.5)
        r.font.color.rgb = RGBColor(255, 255, 255)

    for idx, q in enumerate(quotes_to_show):
        row = t_quotes.rows[1 + idx]
        vals = [
            f"{idx + 1:02d}",
            f"{q.get('buyer_uasg', '')} - {q.get('buyer_uf', '')}\n{q.get('buyer_org_name', '')[:25]}",
            f"{q.get('purchase_number_year', '')}\n({q.get('purchase_date', '')})",
            f"{q.get('supplier_name', '')[:35]}\nCNPJ: {q.get('supplier_cnpj', '')}",
            _format_brl(float(q.get("unit_price", 0.0))),
        ]
        bg = "FFFFFF" if idx % 2 == 0 else "F8FAFC"
        for c_idx, val in enumerate(vals):
            cell = row.cells[c_idx]
            cell.width = q_widths[c_idx]
            set_cell_background(cell, bg)
            set_cell_margins(cell, 60, 60, 60, 60)
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.RIGHT if c_idx == 4 else (WD_ALIGN_PARAGRAPH.CENTER if c_idx in (0, 2) else WD_ALIGN_PARAGRAPH.LEFT)
            r = p.add_run(val)
            r.font.size = Pt(8)
            r.font.name = "Calibri"
            if c_idx == 4:
                r.font.bold = True

    # SEÇÃO 3: SANEAMENTO ESTATÍSTICO
    p_s3 = doc.add_paragraph()
    p_s3.paragraph_format.space_before = Pt(12)
    p_s3.paragraph_format.space_after = Pt(4)
    r_s3 = p_s3.add_run("3. SANEAMENTO ESTATÍSTICO E MEMÓRIA DE CÁLCULO (Art. 6º da IN 65/2021)")
    r_s3.font.bold = True
    r_s3.font.size = Pt(10.5)

    t_stat = doc.add_table(rows=4, cols=4)
    t_stat.alignment = WD_TABLE_ALIGNMENT.CENTER
    s_widths = (Inches(1.7), Inches(1.7), Inches(1.8), Inches(1.7))

    stat_grid = [
        ("Cotações Válidas Auditadas:", str(sm.get("sample_size", len(valid_quotes))), "Outliers Expurgados:", str(len(report.get("purged_quotes", [])))),
        ("Média Saneada:", _format_brl(sm.get("mean")), "Mediana Saneada:", _format_brl(sm.get("median"))),
        ("Desvio Padrão (s):", _format_brl(sm.get("std_deviation")), "Coef. Variação (CV):", f"{sm.get('cv_percent', 0.0):.2f}%"),
        ("Série Homogênea (CV ≤ 25%):", "SIM" if report.get("is_homogeneous") else "NÃO (Heterogênea)", "Critério Adotado:", report.get("adopted_parameter", "MEDIANA")),
    ]

    for r_idx, row_vals in enumerate(stat_grid):
        row = t_stat.rows[r_idx]
        for c_idx, val in enumerate(row_vals):
            cell = row.cells[c_idx]
            cell.width = s_widths[c_idx]
            is_label = (c_idx % 2 == 0)
            set_cell_background(cell, "F1F5F9" if is_label else "FFFFFF")
            set_cell_margins(cell, 60, 60, 80, 80)
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.RIGHT if is_label else WD_ALIGN_PARAGRAPH.LEFT
            r = p.add_run(val)
            r.font.size = Pt(8.5)
            r.font.name = "Calibri"
            r.font.bold = is_label or (c_idx == 3 and r_idx == 3)

    # SEÇÃO 4: JUSTIFICATIVA METODOLÓGICA
    p_s4 = doc.add_paragraph()
    p_s4.paragraph_format.space_before = Pt(12)
    p_s4.paragraph_format.space_after = Pt(4)
    r_s4 = p_s4.add_run("4. JUSTIFICATIVA PARA A METODOLOGIA ADOTADA (Art. 6º, §4º da IN 65/2021)")
    r_s4.font.bold = True
    r_s4.font.size = Pt(10.5)

    p_just = doc.add_paragraph()
    p_just.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p_just.paragraph_format.line_spacing = 1.15
    p_just.paragraph_format.space_after = Pt(12)
    r_just = p_just.add_run(report.get("legal_methodology_justification", ""))
    r_just.font.size = Pt(9)
    r_just.font.name = "Calibri"
    r_just.font.italic = True

    # SEÇÃO 5: ASSINATURAS E FECHAMENTO
    today = datetime.now()
    mes_nome = _get_mes_extenso(today.month)
    data_str = f"Juazeiro, BA, {today.day} de {mes_nome} de {today.year}."

    p_dt = doc.add_paragraph()
    p_dt.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p_dt.paragraph_format.space_before = Pt(16)
    p_dt.paragraph_format.space_after = Pt(28)
    r_dt = p_dt.add_run(data_str)
    r_dt.font.size = Pt(9.5)
    r_dt.font.name = "Calibri"

    t_sig = doc.add_table(rows=1, cols=2)
    t_sig.alignment = WD_TABLE_ALIGNMENT.CENTER
    c_sig0, c_sig1 = t_sig.rows[0].cells[0], t_sig.rows[0].cells[1]
    c_sig0.width, c_sig1.width = Inches(3.4), Inches(3.4)

    p_sig0 = c_sig0.paragraphs[0]
    p_sig0.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_sig0.add_run("_________________________________________\n").font.size = Pt(9)
    r_s0_1 = p_sig0.add_run(f"{elaborador_nome}\n")
    r_s0_1.font.bold = True
    r_s0_1.font.size = Pt(9)
    r_s0_2 = p_sig0.add_run("Elaborador da Pesquisa de Preços")
    r_s0_2.font.size = Pt(8.5)
    r_s0_2.font.color.rgb = RGBColor(100, 116, 139)

    p_sig1 = c_sig1.paragraphs[0]
    p_sig1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_sig1.add_run("_________________________________________\n").font.size = Pt(9)
    r_s1_1 = p_sig1.add_run(f"{ordenador_nome}\n")
    r_s1_1.font.bold = True
    r_s1_1.font.size = Pt(9)
    r_s1_2 = p_sig1.add_run("Ratificador da Pesquisa de Preços")
    r_s1_2.font.size = Pt(8.5)
    r_s1_2.font.color.rgb = RGBColor(100, 116, 139)

    # NOTA DE RODAPÉ
    p_ft = doc.add_paragraph()
    p_ft.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_ft.paragraph_format.space_before = Pt(24)
    r_ft = p_ft.add_run(
        "Documento emitido eletronicamente via Sistema GovTech NLLC • Amparado na Lei nº 14.133/2021 e IN SEGES/ME nº 65/2021"
    )
    r_ft.font.size = Pt(7.5)
    r_ft.font.color.rgb = RGBColor(148, 163, 184)

    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()


# =============================================================================
# 2. GERADOR DO RELATÓRIO EM PDF OFICIAL (A4 / REPORTLAB)
# =============================================================================

def build_in65_pdf_bytes(
    report: Dict[str, Any],
    quantity: int,
    om_name: str = "CAPITANIA FLUVIAL DE JUAZEIRO",
    elaborador_nome: str = "ENCARREGADO DO SETOR DE OBTENÇÃO",
    ordenador_nome: str = "ORDENADOR DE DESPESAS",
) -> bytes:
    """Gera o arquivo PDF A4 de alta qualidade visual com Brasão Oficial e diagramação formal."""
    buf = io.BytesIO()
    doc = SimpleDocTemplate(
        buf,
        pagesize=A4,
        leftMargin=1.5 * cm,
        rightMargin=1.5 * cm,
        topMargin=1.2 * cm,
        bottomMargin=1.2 * cm,
    )

    styles = getSampleStyleSheet()

    # Estilos customizados
    st_header_main = ParagraphStyle(
        "NavalHeaderMain",
        parent=styles["Normal"],
        fontName="Helvetica-Bold",
        fontSize=10,
        leading=12,
        alignment=1,  # Centralizado
        textColor=colors.HexColor("#0f172a"),
    )

    st_header_sub = ParagraphStyle(
        "NavalHeaderSub",
        parent=styles["Normal"],
        fontName="Helvetica",
        fontSize=8.5,
        leading=11,
        alignment=1,
        textColor=colors.HexColor("#334155"),
    )

    st_doc_title = ParagraphStyle(
        "DocTitle",
        parent=styles["Normal"],
        fontName="Helvetica-Bold",
        fontSize=11.5,
        leading=14,
        alignment=1,
        textColor=colors.HexColor("#0f172a"),
    )

    st_doc_sub = ParagraphStyle(
        "DocSub",
        parent=styles["Normal"],
        fontName="Helvetica-Oblique",
        fontSize=8,
        leading=10,
        alignment=1,
        textColor=colors.HexColor("#64748b"),
    )

    st_sec_title = ParagraphStyle(
        "SecTitle",
        parent=styles["Normal"],
        fontName="Helvetica-Bold",
        fontSize=9,
        leading=12,
        textColor=colors.HexColor("#0f172a"),
    )

    st_cell_hdr = ParagraphStyle(
        "CellHdr",
        parent=styles["Normal"],
        fontName="Helvetica-Bold",
        fontSize=7.5,
        leading=9,
        textColor=colors.white,
        alignment=1,
    )

    st_cell_txt = ParagraphStyle(
        "CellTxt",
        parent=styles["Normal"],
        fontName="Helvetica",
        fontSize=7.5,
        leading=9,
        textColor=colors.HexColor("#1e293b"),
    )

    st_cell_txt_center = ParagraphStyle(
        "CellTxtCenter",
        parent=st_cell_txt,
        alignment=1,
    )

    st_cell_bold = ParagraphStyle(
        "CellBold",
        parent=styles["Normal"],
        fontName="Helvetica-Bold",
        fontSize=7.5,
        leading=9,
        textColor=colors.HexColor("#0f172a"),
    )

    st_cell_bold_right = ParagraphStyle(
        "CellBoldRight",
        parent=st_cell_bold,
        alignment=2,
    )

    st_just = ParagraphStyle(
        "JustTxt",
        parent=styles["Normal"],
        fontName="Helvetica-Oblique",
        fontSize=8,
        leading=11,
        alignment=4,  # Justificado
        textColor=colors.HexColor("#1e293b"),
    )

    st_sig_title = ParagraphStyle(
        "SigTitle",
        parent=styles["Normal"],
        fontName="Helvetica-Bold",
        fontSize=8,
        leading=10,
        alignment=1,
    )

    st_sig_sub = ParagraphStyle(
        "SigSub",
        parent=styles["Normal"],
        fontName="Helvetica",
        fontSize=7.5,
        leading=9,
        alignment=1,
        textColor=colors.HexColor("#64748b"),
    )

    st_footer = ParagraphStyle(
        "FooterTxt",
        parent=styles["Normal"],
        fontName="Helvetica",
        fontSize=7,
        leading=8,
        alignment=1,
        textColor=colors.HexColor("#94a3b8"),
    )

    story = []

    # 1. BRASÃO OFICIAL
    if os.path.exists(BRASAO_PATH):
        story.append(RLImage(BRASAO_PATH, width=2.4 * cm, height=2.4 * cm))
        story.append(Spacer(1, 1.5 * mm))

    # 2. CABEÇALHO MILITAR
    story.append(Paragraph("REPÚBLICA FEDERATIVA DO BRASIL", st_header_main))
    story.append(Paragraph("MINISTÉRIO DA DEFESA • COMANDO DA MARINHA", st_header_main))
    story.append(Paragraph(f"{om_name.upper()} • SETOR DE OBTENÇÃO", st_header_sub))
    story.append(Spacer(1, 3 * mm))

    # TÍTULO FORMAL
    story.append(Paragraph("RELATÓRIO DE PESQUISA DE PREÇOS", st_doc_title))
    story.append(
        Paragraph(
            "Conforme Lei nº 14.133/2021 e Instrução Normativa SEGES/ME nº 65/2021",
            st_doc_sub,
        )
    )
    story.append(Spacer(1, 3 * mm))
    story.append(HRFlowable(width="100%", thickness=0.75, color=colors.HexColor("#cbd5e1"), spaceBefore=0, spaceAfter=4))

    # DADOS BÁSICOS
    item_code = report.get("catmat_catser_code", "—")
    item_desc = report.get("item_description", "—")
    unit_str = report.get("target_supply_unit", "UN")
    unit_price = float(report.get("final_estimated_unit_price", 0.0))
    total_price = quantity * unit_price
    sm = report.get("sanitized_metrics", {})
    valid_quotes = report.get("valid_quotes", [])

    # SEÇÃO 1: OBJETO
    story.append(Paragraph("1. IDENTIFICAÇÃO DO OBJETO E VALOR ESTIMADO", st_sec_title))
    story.append(Spacer(1, 1.5 * mm))

    item_table_data = [
        [Paragraph("Código do Catálogo:", st_cell_bold_right), Paragraph(str(item_code), st_cell_txt)],
        [Paragraph("Descrição do Objeto:", st_cell_bold_right), Paragraph(item_desc, st_cell_txt)],
        [Paragraph("Unidade & Quantidade:", st_cell_bold_right), Paragraph(f"{unit_str} | Quantidade solicitada: {quantity}", st_cell_txt)],
        [Paragraph("Preço Unit. Homologável:", st_cell_bold_right), Paragraph(_format_brl(unit_price), st_cell_bold)],
        [Paragraph("Valor Total Estimado:", st_cell_bold_right), Paragraph(f"<b>{_format_brl(total_price)}</b>", st_cell_bold)],
    ]
    t_item = Table(item_table_data, colWidths=[5 * cm, 13 * cm])
    t_item.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (0, -1), colors.HexColor("#f1f5f9")),
        ("BACKGROUND", (1, 4), (1, 4), colors.HexColor("#fef3c7")),
        ("ALIGN", (0, 0), (0, -1), "RIGHT"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#e2e8f0")),
        ("TOPPADDING", (0, 0), (-1, -1), 3),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
    ]))
    story.append(t_item)
    story.append(Spacer(1, 3 * mm))

    # SEÇÃO 2: COTAÇÕES (TOP 6)
    story.append(Paragraph("2. COTAÇÕES VÁLIDAS AUDITADAS (Art. 5º da IN 65/2021)", st_sec_title))
    story.append(Spacer(1, 1.5 * mm))

    quotes_data = [
        [
            Paragraph("Nº", st_cell_hdr),
            Paragraph("Órgão / UASG / UF", st_cell_hdr),
            Paragraph("Nº Compra / Data", st_cell_hdr),
            Paragraph("Fornecedor Vencedor & CNPJ", st_cell_hdr),
            Paragraph("Preço Unit.", st_cell_hdr),
        ]
    ]

    for idx, q in enumerate(valid_quotes[:7]):
        quotes_data.append([
            Paragraph(f"{idx+1:02d}", st_cell_txt_center),
            Paragraph(f"UASG {q.get('buyer_uasg', '')} ({q.get('buyer_uf', '')})<br/>{q.get('buyer_org_name', '')[:25]}", st_cell_txt),
            Paragraph(f"{q.get('purchase_number_year', '')}<br/>{q.get('purchase_date', '')}", st_cell_txt_center),
            Paragraph(f"{q.get('supplier_name', '')[:35]}<br/>CNPJ: {q.get('supplier_cnpj', '')}", st_cell_txt),
            Paragraph(_format_brl(float(q.get("unit_price", 0.0))), st_cell_bold_right),
        ])

    t_quotes = Table(quotes_data, colWidths=[1 * cm, 4.5 * cm, 3.2 * cm, 6.5 * cm, 2.8 * cm])
    t_quotes.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1e293b")),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#cbd5e1")),
        ("TOPPADDING", (0, 0), (-1, -1), 2.5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 2.5),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#f8fafc")]),
    ]))
    story.append(t_quotes)
    story.append(Spacer(1, 3 * mm))

    # SEÇÃO 3: SANEAMENTO ESTATÍSTICO
    story.append(Paragraph("3. SANEAMENTO ESTATÍSTICO E MEMÓRIA DE CÁLCULO (Art. 6º da IN 65/2021)", st_sec_title))
    story.append(Spacer(1, 1.5 * mm))

    stat_data = [
        [
            Paragraph("Amostras Válidas:", st_cell_bold_right),
            Paragraph(str(sm.get("sample_size", len(valid_quotes))), st_cell_txt),
            Paragraph("Outliers Expurgados:", st_cell_bold_right),
            Paragraph(str(len(report.get("purged_quotes", []))), st_cell_txt),
        ],
        [
            Paragraph("Média Saneada:", st_cell_bold_right),
            Paragraph(_format_brl(sm.get("mean")), st_cell_txt),
            Paragraph("Mediana Saneada:", st_cell_bold_right),
            Paragraph(_format_brl(sm.get("median")), st_cell_txt),
        ],
        [
            Paragraph("Desvio Padrão (s):", st_cell_bold_right),
            Paragraph(_format_brl(sm.get("std_deviation")), st_cell_txt),
            Paragraph("Coef. Variação (CV):", st_cell_bold_right),
            Paragraph(f"{sm.get('cv_percent', 0.0):.2f}%", st_cell_txt),
        ],
        [
            Paragraph("Série Homogênea:", st_cell_bold_right),
            Paragraph("SIM (CV ≤ 25%)" if report.get("is_homogeneous") else "NÃO (Heterogênea)", st_cell_txt),
            Paragraph("Critério Adotado:", st_cell_bold_right),
            Paragraph(f"<b>{report.get('adopted_parameter', 'MEDIANA')}</b>", st_cell_bold),
        ],
    ]
    t_stat = Table(stat_data, colWidths=[4.5 * cm, 4.5 * cm, 4.5 * cm, 4.5 * cm])
    t_stat.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (0, -1), colors.HexColor("#f1f5f9")),
        ("BACKGROUND", (2, 0), (2, -1), colors.HexColor("#f1f5f9")),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#e2e8f0")),
        ("TOPPADDING", (0, 0), (-1, -1), 2.5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 2.5),
    ]))
    story.append(t_stat)
    story.append(Spacer(1, 3 * mm))

    # SEÇÃO 4: JUSTIFICATIVA JURÍDICA
    story.append(Paragraph("4. JUSTIFICATIVA METODOLÓGICA (Art. 6º da IN 65/2021)", st_sec_title))
    story.append(Spacer(1, 1 * mm))
    story.append(Paragraph(report.get("legal_methodology_justification", ""), st_just))
    story.append(Spacer(1, 4 * mm))

    # SEÇÃO 5: ASSINATURAS
    today = datetime.now()
    mes_nome = _get_mes_extenso(today.month)
    data_str = f"Juazeiro, BA, {today.day} de {mes_nome} de {today.year}."

    sig_data = [
        [Paragraph(f"<i>{data_str}</i>", st_sig_sub), Paragraph("", st_sig_sub)],
        [Spacer(1, 6 * mm), Spacer(1, 6 * mm)],
        [
            Paragraph("________________________________________", st_sig_title),
            Paragraph("________________________________________", st_sig_title),
        ],
        [
            Paragraph(f"<b>{elaborador_nome}</b>", st_sig_title),
            Paragraph(f"<b>{ordenador_nome}</b>", st_sig_title),
        ],
        [
            Paragraph("Elaborador da Pesquisa de Preços", st_sig_sub),
            Paragraph("Ratificador da Pesquisa de Preços", st_sig_sub),
        ],
    ]
    t_sig = Table(sig_data, colWidths=[9 * cm, 9 * cm])
    t_sig.setStyle(TableStyle([
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 1),
        ("TOPPADDING", (0, 0), (-1, -1), 1),
    ]))
    story.append(KeepTogether(t_sig))
    story.append(Spacer(1, 4 * mm))
    story.append(
        Paragraph(
            "Documento emitido eletronicamente via GovTech NLLC • Trilha de Auditoria Auditável no PNCP / Compras.gov.br",
            st_footer,
        )
    )

    doc.build(story)
    return buf.getvalue()
