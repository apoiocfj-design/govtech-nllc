# -*- coding: utf-8 -*-
"""
MÓDULO DE GERAÇÃO DE RELATÓRIOS OFICIAIS (WORD .DOCX E PDF A4)
Padrão Marinha do Brasil e Instrução Normativa SEGES/ME nº 65/2021 (Lei nº 14.133/2021)
Gera em memória (BytesIO) relatórios completos e prontos para instrução processual.
"""

import io
import os
import re
import zipfile
import json
import math
from datetime import datetime, timezone
from typing import Dict, Any, List, Optional
from PIL import Image as PILImage, ImageDraw as PILImageDraw
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


def generate_statistical_infographic(
    valid_prices: List[float],
    purged_prices: List[float],
    cv_percent: float,
    mean_val: float,
    median_val: float,
    std_val: float,
    adopted_param: str = "MEDIANA",
    width: int = 1000,
    height: int = 440,
) -> bytes:
    """
    Gera um infográfico estatístico institucional de alta resolução (Pillow) contendo:
      1. Boxplot dos Quartis (Saneamento Estatístico IN SEGES/ME nº 65/2021)
      2. Curva Normal Gaussiana e Critérios de Escolha (Média vs Mediana vs Outliers)
    """
    img = PILImage.new("RGB", (width, height), color="#ffffff")
    draw = PILImageDraw.Draw(img)

    c_navy = "#1e2d3d"
    c_steel = "#3b5b7c"
    c_gold = "#c5a365"
    c_green = "#15803d"
    c_red = "#be123c"
    c_bg_box = "#f8fafc"
    c_border = "#e2e8f0"
    c_text_dark = "#0f172a"
    c_text_muted = "#64748b"

    # Cabeçalho Principal do Infográfico
    draw.rectangle([(0, 0), (width, 42)], fill=c_navy)
    draw.text((20, 12), "DEMONSTRATIVO ESTATÍSTICO & CURVA GAUSSIANA DE DECISÃO (IN SEGES/ME Nº 65/2021)", fill="#ffffff")
    draw.text((width - 240, 14), "PADRÃO OFICIAL TCU / AUDITORIA", fill="#facc15")

    # Moldura dos dois painéis
    draw.rounded_rectangle([(15, 52), (485, 425)], radius=10, fill=c_bg_box, outline=c_border, width=1)
    draw.rounded_rectangle([(495, 52), (985, 425)], radius=10, fill=c_bg_box, outline=c_border, width=1)

    # =========================================================================
    # PAINEL 1: BOXPLOT DA AMOSTRA SANEADA
    # =========================================================================
    draw.text((30, 65), "1. DISPERSÃO DA AMOSTRA (MÉTODO DOS QUARTIS / BOXPLOT)", fill=c_navy)
    draw.text((30, 82), "Saneamento estatístico das cotações colhidas nas bases oficiais do Governo", fill=c_text_muted)

    v_sorted = sorted([float(p) for p in valid_prices]) if valid_prices else [10.0]
    p_vals = [float(p) for p in purged_prices] if purged_prices else []
    all_vals = v_sorted + p_vals

    v_min = v_sorted[0]
    v_max = v_sorted[-1]

    def get_p(arr, pct):
        if len(arr) == 1:
            return arr[0]
        idx = (len(arr) - 1) * pct
        low = int(idx)
        high = math.ceil(idx)
        if low == high:
            return arr[low]
        return arr[low] + (arr[high] - arr[low]) * (idx - low)

    q1 = get_p(v_sorted, 0.25)
    q3 = get_p(v_sorted, 0.75)
    iqr = max(0.01, q3 - q1)
    lower_w = max(v_min, q1 - 1.5 * iqr)
    upper_w = min(v_max, q3 + 1.5 * iqr)

    abs_min = min(all_vals)
    abs_max = max(all_vals)
    span = max(0.1, (abs_max - abs_min) * 1.25)
    d_min = max(0, abs_min - span * 0.1)
    d_max = abs_max + span * 0.1

    bx_l, bx_r = 45, 455
    bx_w = bx_r - bx_l

    def scale_bx(val):
        if d_max <= d_min:
            return (bx_l + bx_r) / 2
        return bx_l + ((val - d_min) / (d_max - d_min)) * bx_w

    y_axis = 160
    draw.line([(bx_l, y_axis), (bx_r, y_axis)], fill="#cbd5e1", width=1)

    x_min = scale_bx(lower_w)
    x_q1 = scale_bx(q1)
    x_med = scale_bx(median_val)
    x_mean = scale_bx(mean_val)
    x_q3 = scale_bx(q3)
    x_max = scale_bx(upper_w)

    draw.line([(x_min, y_axis), (x_q1, y_axis)], fill=c_steel, width=3)
    draw.line([(x_q3, y_axis), (x_max, y_axis)], fill=c_steel, width=3)
    draw.line([(x_min, y_axis - 15), (x_min, y_axis + 15)], fill=c_steel, width=3)
    draw.line([(x_max, y_axis - 15), (x_max, y_axis + 15)], fill=c_steel, width=3)

    draw.rounded_rectangle([(x_q1, y_axis - 30), (x_q3, y_axis + 30)], radius=4, fill="#e2e8f0", outline=c_steel, width=2)
    draw.line([(x_med, y_axis - 30), (x_med, y_axis + 30)], fill=c_gold, width=4)

    for dy in range(-36, 37, 6):
        draw.line([(x_mean, y_axis + dy), (x_mean, y_axis + dy + 3)], fill=c_green, width=2)

    for i, p in enumerate(v_sorted):
        cx = scale_bx(p)
        cy = y_axis + ((i % 3) - 1) * 8
        draw.ellipse([(cx - 5, cy - 5), (cx + 5, cy + 5)], fill="#10b981", outline="#047857", width=1)

    for i, p in enumerate(p_vals):
        cx = scale_bx(p)
        cy = y_axis + ((i % 2) * 16 - 8)
        draw.ellipse([(cx - 6, cy - 6), (cx + 6, cy + 6)], fill="#f43f5e", outline="#be123c", width=1)

    draw.line([(25, 230), (475, 230)], fill="#e2e8f0", width=1)
    draw.text((30, 240), f"Menor: R$ {v_min:.2f}", fill=c_text_dark)
    draw.text((180, 240), f"Q1 (25%): R$ {q1:.2f}", fill=c_text_dark)
    draw.text((330, 240), f"Q3 (75%): R$ {q3:.2f}", fill=c_text_dark)

    draw.text((30, 265), f"Mediana: R$ {median_val:.2f}", fill=c_gold)
    draw.text((180, 265), f"Média: R$ {mean_val:.2f}", fill=c_green)
    draw.text((330, 265), f"Maior: R$ {v_max:.2f}", fill=c_text_dark)

    draw.rounded_rectangle([(30, 305), (470, 405)], radius=6, fill="#ffffff", outline="#cbd5e1", width=1)
    draw.text((40, 315), "RESUMO DA AUDITORIA DA CESTA:", fill=c_navy)
    draw.text((40, 335), f"• Total Coletado: {len(all_vals)} cotações ({len(v_sorted)} válidas e {len(p_vals)} outliers expurgados)", fill=c_text_dark)
    draw.text((40, 355), f"• Coeficiente de Variação (CV): {cv_percent:.2f}% (Desvio Padrão: R$ {std_val:.2f})", fill=c_text_dark)
    is_homo = cv_percent <= 25.0
    status_str = "SÉRIE HOMOGÊNEA (CV <= 25%)" if is_homo else "SÉRIE HETEROGÊNEA (CV > 25%)"
    status_col = c_green if is_homo else "#b45309"
    draw.text((40, 375), f"• Classificação IN 65/2021: {status_str}", fill=status_col)

    # =========================================================================
    # PAINEL 2: CURVA NORMAL GAUSSIANA E CRITÉRIOS DE ESCOLHA
    # =========================================================================
    draw.text((510, 65), "2. DISTRIBUIÇÃO NORMAL GAUSSIANA & CRITÉRIOS DE ESCOLHA", fill=c_navy)
    draw.text((510, 82), "Critérios legais e matemáticos para adoção da Média vs. Mediana (Art. 6º)", fill=c_text_muted)

    cx_l, cx_r = 525, 955
    cx_w = cx_r - cx_l
    cy_base = 230
    cy_peak = 115

    curve_pts = []
    steps = 90
    for s in range(steps + 1):
        z = -3.0 + (6.0 * s / steps)
        y_val = math.exp(-0.5 * (z ** 2))
        px = cx_l + (s / steps) * cx_w
        py = cy_base - (y_val * (cy_base - cy_peak))
        curve_pts.append((px, py))

    draw.line([(cx_l, cy_base), (cx_r, cy_base)], fill="#cbd5e1", width=2)

    x_m2 = cx_l + (1.0 / 6.0) * cx_w
    x_m1 = cx_l + (2.0 / 6.0) * cx_w
    x_mid = cx_l + (3.0 / 6.0) * cx_w
    x_p1 = cx_l + (4.0 / 6.0) * cx_w
    x_p2 = cx_l + (5.0 / 6.0) * cx_w

    poly_central = [(x_m1, cy_base)] + [pt for pt in curve_pts if x_m1 <= pt[0] <= x_p1] + [(x_p1, cy_base)]
    if len(poly_central) > 2:
        draw.polygon(poly_central, fill="#dcfce7")

    draw.line([(x_m1, cy_base), (x_m1, cy_peak + 25)], fill="#86efac", width=1)
    draw.line([(x_p1, cy_base), (x_p1, cy_peak + 25)], fill="#86efac", width=1)
    draw.line([(x_m2, cy_base), (x_m2, cy_peak + 70)], fill="#fca5a5", width=1)
    draw.line([(x_p2, cy_base), (x_p2, cy_peak + 70)], fill="#fca5a5", width=1)

    for i in range(len(curve_pts) - 1):
        draw.line([curve_pts[i], curve_pts[i+1]], fill=c_steel, width=3)

    draw.line([(x_mid, cy_base), (x_mid, cy_peak)], fill=c_gold, width=3)

    draw.text((x_mid - 25, cy_base - 50), "ZONA CENTRAL", fill=c_green)
    draw.text((x_mid - 20, cy_base - 35), "CV <= 25%", fill=c_green)
    draw.text((x_mid - 15, cy_base + 6), "u (Média)", fill=c_navy)

    draw.text((x_p1 + 8, cy_base - 40), "DISPERSÃO", fill="#b45309")
    draw.text((x_p1 + 8, cy_base - 25), "CV > 25%", fill="#b45309")
    draw.text((x_p1 - 10, cy_base + 6), "+1s", fill=c_text_muted)
    draw.text((x_m1 - 10, cy_base + 6), "-1s", fill=c_text_muted)

    draw.text((cx_l + 10, cy_base - 25), "CAUDA (Expurgo)", fill=c_red)
    draw.text((cx_r - 110, cy_base - 25), "CAUDA (Expurgo)", fill=c_red)
    draw.text((x_m2 - 10, cy_base + 6), "-2s", fill=c_text_muted)
    draw.text((x_p2 - 10, cy_base + 6), "+2s", fill=c_text_muted)

    draw.rounded_rectangle([(510, 260), (970, 410)], radius=6, fill="#ffffff", outline="#cbd5e1", width=1)

    draw.rectangle([(520, 270), (524, 305)], fill=c_green)
    draw.text((532, 270), "REGRA DA MÉDIA SANEADA (CV <= 25%):", fill=c_green)
    draw.text((532, 287), "Quando a amostra é homogênea, a média aritmética reflete com precisão o", fill=c_text_dark)
    draw.text((532, 300), "valor de mercado e pondera igualmente as propostas válidas.", fill=c_text_dark)

    draw.rectangle([(520, 318), (524, 353)], fill=c_gold)
    draw.text((532, 318), "REGRA DA MEDIANA SANEADA (CV > 25%):", fill="#92400e")
    draw.text((532, 335), "Quando a amostra é heterogênea ou assimétrica, a mediana é obrigatória", fill=c_text_dark)
    draw.text((532, 348), "(Art. 6º, § 5º da IN 65/2021) pois é imune à distorção de preços nos extremos.", fill=c_text_dark)

    adopted_is_med = "MEDIANA" in adopted_param.upper()
    bg_dec = "#fef3c7" if adopted_is_med else "#dcfce7"
    border_dec = c_gold if adopted_is_med else c_green
    draw.rounded_rectangle([(520, 368), (960, 400)], radius=4, fill=bg_dec, outline=border_dec, width=1)
    dec_text = f"DECISÃO LEGAL ADOTADA: {adopted_param} (R$ {median_val if adopted_is_med else mean_val:.2f}) — CV Saneado de {cv_percent:.2f}%"
    draw.text((530, 376), dec_text, fill=c_navy)

    buf = io.BytesIO()
    img.save(buf, format="PNG", dpi=(300, 300))
    return buf.getvalue()


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
    om_name: str = "MARINHA DO BRASIL",
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

    # SEÇÃO 4: DEMONSTRATIVO GRÁFICO (BOXPLOT E CURVA NORMAL GAUSSIANA)
    p_sg = doc.add_paragraph()
    p_sg.paragraph_format.space_before = Pt(12)
    p_sg.paragraph_format.space_after = Pt(4)
    r_sg = p_sg.add_run("4. DEMONSTRATIVO GRÁFICO: DISPERSÃO (BOXPLOT) E CURVA GAUSSIANA (IN 65/2021)")
    r_sg.font.bold = True
    r_sg.font.size = Pt(10.5)

    try:
        infographic_bytes = generate_statistical_infographic(
            valid_prices=[float(q.get("unit_price", 0.0)) for q in valid_quotes],
            purged_prices=[float(q.get("unit_price", 0.0)) for q in report.get("purged_quotes", [])],
            cv_percent=float(sm.get("cv_percent", 0.0)),
            mean_val=float(sm.get("mean", 0.0)),
            median_val=float(sm.get("median", 0.0)),
            std_val=float(sm.get("std_deviation", 0.0)),
            adopted_param=str(report.get("adopted_parameter", "MEDIANA")),
        )
        p_chart = doc.add_paragraph()
        p_chart.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_chart.paragraph_format.space_after = Pt(8)
        run_chart = p_chart.add_run()
        run_chart.add_picture(io.BytesIO(infographic_bytes), width=Inches(6.8))
    except Exception as exc:
        print(f"Erro ao inserir infográfico estatístico no Word: {exc}")

    # SEÇÃO 5: JUSTIFICATIVA METODOLÓGICA
    p_s4 = doc.add_paragraph()
    p_s4.paragraph_format.space_before = Pt(12)
    p_s4.paragraph_format.space_after = Pt(4)
    r_s4 = p_s4.add_run("5. JUSTIFICATIVA PARA A METODOLOGIA ADOTADA (Art. 6º, §4º da IN 65/2021)")
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
    data_str = f"{today.day} de {mes_nome} de {today.year}."

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
    pass  # Rodapé de sistema removido a pedido

    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()


# =============================================================================
# 2. GERADOR DO RELATÓRIO EM PDF OFICIAL (A4 / REPORTLAB)
# =============================================================================

def build_in65_pdf_bytes(
    report: Dict[str, Any],
    quantity: int,
    om_name: str = "MARINHA DO BRASIL",
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

    # SEÇÃO 4: DEMONSTRATIVO GRÁFICO (BOXPLOT E CURVA NORMAL GAUSSIANA)
    story.append(Paragraph("4. DEMONSTRATIVO GRÁFICO: DISPERSÃO (BOXPLOT) E CURVA GAUSSIANA (IN 65/2021)", st_sec_title))
    story.append(Spacer(1, 1.5 * mm))
    try:
        infographic_bytes = generate_statistical_infographic(
            valid_prices=[float(q.get("unit_price", 0.0)) for q in valid_quotes],
            purged_prices=[float(q.get("unit_price", 0.0)) for q in report.get("purged_quotes", [])],
            cv_percent=float(sm.get("cv_percent", 0.0)),
            mean_val=float(sm.get("mean", 0.0)),
            median_val=float(sm.get("median", 0.0)),
            std_val=float(sm.get("std_deviation", 0.0)),
            adopted_param=str(report.get("adopted_parameter", "MEDIANA")),
        )
        story.append(RLImage(io.BytesIO(infographic_bytes), width=18.0 * cm, height=7.92 * cm))
        story.append(Spacer(1, 3 * mm))
    except Exception as exc:
        print(f"Erro ao inserir gráfico estatístico no PDF: {exc}")

    # SEÇÃO 5: JUSTIFICATIVA JURÍDICA
    story.append(Paragraph("5. JUSTIFICATIVA METODOLÓGICA (Art. 6º da IN 65/2021)", st_sec_title))
    story.append(Spacer(1, 1 * mm))
    story.append(Paragraph(report.get("legal_methodology_justification", ""), st_just))
    story.append(Spacer(1, 4 * mm))

    # SEÇÃO 5: ASSINATURAS
    today = datetime.now()
    mes_nome = _get_mes_extenso(today.month)
    data_str = f"{today.day} de {mes_nome} de {today.year}."

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
    pass  # Rodapé de sistema removido a pedido

    doc.build(story)
    return buf.getvalue()


# =============================================================================
# 3. GERADOR DO DESPACHO DE JUSTIFICATIVA CIRCUNSTANCIADA (WORD .DOCX)
#    (Art. 6º, § 5º da Instrução Normativa SEGES/ME nº 65/2021)
# =============================================================================

def build_in65_justificativa_docx_bytes(
    report: Dict[str, Any],
    quantity: int,
    om_name: str = "MARINHA DO BRASIL",
    ordenador_nome: str = "ORDENADOR DE DESPESAS",
    motivo_custom: str = None,
) -> bytes:
    """Gera em Word (.docx) o Despacho Formal de Justificativa Circunstanciada (Art. 6º, § 5º da IN 65/2021)."""
    doc = docx.Document()

    for section in doc.sections:
        section.top_margin = Inches(0.8)
        section.bottom_margin = Inches(0.8)
        section.left_margin = Inches(0.8)
        section.right_margin = Inches(0.8)

    # CABEÇALHO OFICIAL COM BRASÃO
    if os.path.exists(BRASAO_PATH):
        p_img = doc.add_paragraph()
        p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_img.paragraph_format.space_after = Pt(4)
        run_img = p_img.add_run()
        run_img.add_picture(BRASAO_PATH, width=Inches(1.05))

    p_head = doc.add_paragraph()
    p_head.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_head.paragraph_format.line_spacing = 1.05
    p_head.paragraph_format.space_after = Pt(12)

    lines = [
        ("REPÚBLICA FEDERATIVA DO BRASIL", 10.5, True),
        ("MINISTÉRIO DA DEFESA", 10, True),
        ("COMANDO DA MARINHA", 10, True),
        (om_name.upper(), 9.5, False),
        ("GABINETE DO ORDENADOR DE DESPESAS", 9, False),
    ]
    for text, sz, bold in lines:
        r = p_head.add_run(f"{text}\n")
        r.font.name = "Calibri"
        r.font.size = Pt(sz)
        r.font.bold = bold
        r.font.color.rgb = RGBColor(15, 23, 42)

    # TÍTULO FORMAL DO DESPACHO
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_title.paragraph_format.space_before = Pt(4)
    p_title.paragraph_format.space_after = Pt(4)
    r_tit = p_title.add_run("DESPACHO DE JUSTIFICATIVA CIRCUNSTANCIADA DA AUTORIDADE COMPETENTE\n")
    r_tit.font.name = "Calibri"
    r_tit.font.size = Pt(11.5)
    r_tit.font.bold = True
    r_tit.font.color.rgb = RGBColor(15, 23, 42)

    r_sub = p_title.add_run(
        "Excepcionalidade do Art. 6º, § 5º da Instrução Normativa SEGES/ME nº 65/2021 c/c Art. 23 da Lei nº 14.133/2021\n"
        "(Fixação de preço estimado com base em menos de três cotações válidas)"
    )
    r_sub.font.name = "Calibri"
    r_sub.font.size = Pt(9.5)
    r_sub.font.italic = True
    r_sub.font.color.rgb = RGBColor(71, 85, 105)

    doc.add_paragraph().paragraph_format.space_after = Pt(4)

    item_code = report.get("catmat_catser_code", "—")
    item_desc = report.get("item_description", "—")
    unit_str = report.get("target_supply_unit", "UN")
    unit_price = float(report.get("final_estimated_unit_price", 0.0))
    total_price = quantity * unit_price
    valid_quotes = report.get("valid_quotes", [])
    n_quotes = len(valid_quotes)

    # 1. DO OBJETO DEMANDADO
    p_s1 = doc.add_paragraph()
    p_s1.paragraph_format.space_before = Pt(8)
    p_s1.paragraph_format.space_after = Pt(4)
    p_s1.add_run("1. IDENTIFICAÇÃO DO OBJETO E VALOR ESTIMADO").font.bold = True

    t_item = doc.add_table(rows=4, cols=2)
    t_item.alignment = WD_TABLE_ALIGNMENT.CENTER
    t_data = [
        ("Código do Item (CATMAT/CATSER):", f"{item_code}"),
        ("Descrição Detalhada do Objeto:", f"{item_desc}"),
        ("Quantidade Solicitada & Unidade:", f"{quantity} {unit_str}"),
        ("Preço Unitário / Valor Total Estimado:", f"{_format_brl(unit_price)} un. | Total: {_format_brl(total_price)}"),
    ]
    for r_idx, (k, v) in enumerate(t_data):
        row = t_item.rows[r_idx]
        c0, c1 = row.cells[0], row.cells[1]
        c0.width, c1.width = Inches(2.4), Inches(4.5)
        set_cell_margins(c0, 60, 60, 100, 100)
        set_cell_margins(c1, 60, 60, 100, 100)
        set_cell_background(c0, "F1F5F9")
        set_cell_background(c1, "FFFFFF" if r_idx < 3 else "FEF3C7")
        p0 = c0.paragraphs[0]
        p0.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        r0 = p0.add_run(k)
        r0.font.bold = True
        r0.font.size = Pt(9)
        p1 = c1.paragraphs[0]
        p1.alignment = WD_ALIGN_PARAGRAPH.LEFT
        r1 = p1.add_run(v)
        r1.font.bold = (r_idx == 3)
        r1.font.size = Pt(9.5)

    # 2. DAS COTAÇÕES VÁLIDAS OBTIDAS
    p_s2 = doc.add_paragraph()
    p_s2.paragraph_format.space_before = Pt(12)
    p_s2.paragraph_format.space_after = Pt(4)
    p_s2.add_run(f"2. COTAÇÕES VÁLIDAS IDENTIFICADAS ({n_quotes} cotação(ões) obtida(s))").font.bold = True

    t_quotes = doc.add_table(rows=1 + max(1, n_quotes), cols=5)
    t_quotes.alignment = WD_TABLE_ALIGNMENT.CENTER
    q_widths = (Inches(0.6), Inches(1.5), Inches(1.3), Inches(2.2), Inches(1.3))
    q_headers = ["Nº", "UASG / Órgão", "Compra nº", "Fornecedor Vencedor & CNPJ", "Preço Unit."]
    for c_idx, h in enumerate(q_headers):
        cell = t_quotes.rows[0].cells[c_idx]
        cell.width = q_widths[c_idx]
        set_cell_background(cell, "1E293B")
        set_cell_margins(cell, 60, 60, 60, 60)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(h)
        r.font.bold = True
        r.font.size = Pt(8.5)
        r.font.color.rgb = RGBColor(255, 255, 255)

    for idx, q in enumerate(valid_quotes):
        row = t_quotes.rows[1 + idx]
        vals = [
            f"{idx + 1:02d}",
            f"{q.get('buyer_uasg', '')} - {q.get('buyer_uf', '')}\n{q.get('buyer_org_name', '')[:25]}",
            f"{q.get('purchase_number_year', '')}\n({q.get('purchase_date', '')})",
            f"{q.get('supplier_name', '')[:35]}\nCNPJ: {q.get('supplier_cnpj', '')}",
            _format_brl(float(q.get("unit_price", 0.0))),
        ]
        for c_idx, val in enumerate(vals):
            cell = row.cells[c_idx]
            cell.width = q_widths[c_idx]
            set_cell_background(cell, "FFFFFF" if idx % 2 == 0 else "F8FAFC")
            set_cell_margins(cell, 60, 60, 60, 60)
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.RIGHT if c_idx == 4 else (WD_ALIGN_PARAGRAPH.CENTER if c_idx in (0, 2) else WD_ALIGN_PARAGRAPH.LEFT)
            r = p.add_run(val)
            r.font.size = Pt(8)
            if c_idx == 4:
                r.font.bold = True

    # 3. DA MOTIVAÇÃO E FUNDAMENTAÇÃO CIRCUNSTANCIADA
    p_s3 = doc.add_paragraph()
    p_s3.paragraph_format.space_before = Pt(12)
    p_s3.paragraph_format.space_after = Pt(4)
    p_s3.add_run("3. DA MOTIVAÇÃO E FUNDAMENTAÇÃO CIRCUNSTANCIADA").font.bold = True

    just_paragraphs = [
        ("I. Da Abrangência da Pesquisa Pública:",
         "A pesquisa de preços foi conduzida de forma ampla e irrestrita nos bancos de dados oficiais do Governo Federal (Compras.gov.br e Portal Nacional de Contratações Públicas - PNCP), alcançando contratações executadas por órgãos e entidades da Administração Pública em âmbito nacional no período regulamentar de até 1 (um) ano."),
        ("II. Da Escassez de Registros Similares no Mercado Público:",
         f"Não obstante a busca exaustiva efetuada, constatou-se a disponibilidade de apenas {n_quotes} cotação(ões) válida(s) apta(s) a compor a cesta de preços para o item especificado, em virtude da sazonalidade, especificidade do objeto demandado ou ausência de contratações análogas recentes por outros entes públicos."),
        ("III. Da Razoabilidade e Economicidade dos Preços Obtidos:",
         f"Os valores apurados ({_format_brl(unit_price)} un.) foram submetidos à análise crítica de exequibilidade e compatibilidade mercadológica, não se vislumbrando qualquer indício de sobrepreço ou distorção econômica, atendendo rigorosamente aos princípios da economicidade, eficiência e vantajosidade insculpidos no Art. 5º da Lei nº 14.133/2021."),
        ("IV. Da Necessidade Inadiável do Atendimento:",
         "A paralisação ou dilação excessiva do procedimento de obtenção acarretaria grave prejuízo à continuidade das atividades finalísticas e ao regular funcionamento operacional desta Organização Militar, justificando-se a excepcionalidade."),
    ]
    for tit, corpo in just_paragraphs:
        p_j = doc.add_paragraph()
        p_j.paragraph_format.line_spacing = 1.15
        p_j.paragraph_format.space_after = Pt(4)
        p_j.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        r_tit_j = p_j.add_run(f"{tit} ")
        r_tit_j.font.bold = True
        r_tit_j.font.size = Pt(9)
        r_corpo = p_j.add_run(corpo)
        r_corpo.font.size = Pt(9)

    # 4. DECISÃO DA AUTORIDADE COMPETENTE
    p_s4 = doc.add_paragraph()
    p_s4.paragraph_format.space_before = Pt(10)
    p_s4.paragraph_format.space_after = Pt(4)
    p_s4.add_run("4. DECISÃO DA AUTORIDADE COMPETENTE").font.bold = True

    p_dec = doc.add_paragraph()
    p_dec.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p_dec.paragraph_format.line_spacing = 1.15
    p_dec.paragraph_format.space_after = Pt(12)
    r_dec = p_dec.add_run(
        f"Diante dos elementos técnicos e fáticos coligidos e com esteio no permissivo do Art. 6º, § 5º da Instrução Normativa SEGES/ME nº 65/2021, "
        f"AUTORIZO e RATIFICO, em caráter excepcional e devidamente circunstanciado, a adoção do preço estimado com base na(s) {n_quotes} cotação(ões) "
        f"válida(s) identificada(s), fixando o valor de referência em {_format_brl(unit_price)} por unidade (total: {_format_brl(total_price)}), "
        f"determinando o regular prosseguimento da instrução processual."
    )
    r_dec.font.bold = True
    r_dec.font.size = Pt(9)

    # FECHAMENTO E ASSINATURA
    today = datetime.now()
    mes_nome = _get_mes_extenso(today.month)
    data_str = f"{today.day} de {mes_nome} de {today.year}."

    p_dt = doc.add_paragraph()
    p_dt.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p_dt.paragraph_format.space_before = Pt(14)
    p_dt.paragraph_format.space_after = Pt(28)
    r_dt = p_dt.add_run(data_str)
    r_dt.font.size = Pt(9.5)

    p_sig = doc.add_paragraph()
    p_sig.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_sig.add_run("____________________________________________________\n").font.size = Pt(9)
    r_sig1 = p_sig.add_run(f"{ordenador_nome}\n")
    r_sig1.font.bold = True
    r_sig1.font.size = Pt(9.5)
    r_sig2 = p_sig.add_run("Ordenador de Despesas / Autoridade Competente\n")
    r_sig2.font.size = Pt(8.5)
    r_sig2.font.color.rgb = RGBColor(100, 116, 139)
    r_sig3 = p_sig.add_run(om_name.upper())
    r_sig3.font.size = Pt(8)
    r_sig3.font.color.rgb = RGBColor(100, 116, 139)

    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()


# =============================================================================
# 4. GERADOR DO DESPACHO DE JUSTIFICATIVA CIRCUNSTANCIADA (PDF A4 OFICIAL)
# =============================================================================

def build_in65_justificativa_pdf_bytes(
    report: Dict[str, Any],
    quantity: int,
    om_name: str = "MARINHA DO BRASIL",
    ordenador_nome: str = "ORDENADOR DE DESPESAS",
    motivo_custom: str = None,
) -> bytes:
    """Gera em PDF A4 formal o Despacho de Justificativa Circunstanciada (Art. 6º, § 5º da IN 65/2021)."""
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

    st_header_main = ParagraphStyle(
        "JustHeaderMain",
        parent=styles["Normal"],
        fontName="Helvetica-Bold",
        fontSize=10,
        leading=12,
        alignment=1,
        textColor=colors.HexColor("#0f172a"),
    )

    st_header_sub = ParagraphStyle(
        "JustHeaderSub",
        parent=styles["Normal"],
        fontName="Helvetica",
        fontSize=8.5,
        leading=11,
        alignment=1,
        textColor=colors.HexColor("#334155"),
    )

    st_doc_title = ParagraphStyle(
        "JustDocTitle",
        parent=styles["Normal"],
        fontName="Helvetica-Bold",
        fontSize=11,
        leading=13,
        alignment=1,
        textColor=colors.HexColor("#0f172a"),
    )

    st_doc_sub = ParagraphStyle(
        "JustDocSub",
        parent=styles["Normal"],
        fontName="Helvetica-Oblique",
        fontSize=8,
        leading=10,
        alignment=1,
        textColor=colors.HexColor("#64748b"),
    )

    st_sec_title = ParagraphStyle(
        "JustSecTitle",
        parent=styles["Normal"],
        fontName="Helvetica-Bold",
        fontSize=9,
        leading=12,
        textColor=colors.HexColor("#0f172a"),
    )

    st_cell_hdr = ParagraphStyle(
        "JustCellHdr",
        parent=styles["Normal"],
        fontName="Helvetica-Bold",
        fontSize=7.5,
        leading=9,
        textColor=colors.white,
        alignment=1,
    )

    st_cell_txt = ParagraphStyle(
        "JustCellTxt",
        parent=styles["Normal"],
        fontName="Helvetica",
        fontSize=7.5,
        leading=9,
        textColor=colors.HexColor("#1e293b"),
    )

    st_cell_txt_center = ParagraphStyle(
        "JustCellTxtCenter",
        parent=st_cell_txt,
        alignment=1,
    )

    st_cell_bold = ParagraphStyle(
        "JustCellBold",
        parent=styles["Normal"],
        fontName="Helvetica-Bold",
        fontSize=7.5,
        leading=9,
        textColor=colors.HexColor("#0f172a"),
    )

    st_cell_bold_right = ParagraphStyle(
        "JustCellBoldRight",
        parent=st_cell_bold,
        alignment=2,
    )

    st_body = ParagraphStyle(
        "JustBody",
        parent=styles["Normal"],
        fontName="Helvetica",
        fontSize=8,
        leading=11,
        alignment=4,
        textColor=colors.HexColor("#1e293b"),
    )

    st_sig_title = ParagraphStyle(
        "JustSigTitle",
        parent=styles["Normal"],
        fontName="Helvetica-Bold",
        fontSize=8.5,
        leading=11,
        alignment=1,
    )

    st_sig_sub = ParagraphStyle(
        "JustSigSub",
        parent=styles["Normal"],
        fontName="Helvetica",
        fontSize=7.5,
        leading=9,
        alignment=1,
        textColor=colors.HexColor("#64748b"),
    )

    story = []

    # 1. BRASÃO
    if os.path.exists(BRASAO_PATH):
        story.append(RLImage(BRASAO_PATH, width=2.4 * cm, height=2.4 * cm))
        story.append(Spacer(1, 1.5 * mm))

    # 2. CABEÇALHO MILITAR
    story.append(Paragraph("REPÚBLICA FEDERATIVA DO BRASIL", st_header_main))
    story.append(Paragraph("MINISTÉRIO DA DEFESA • COMANDO DA MARINHA", st_header_main))
    story.append(Paragraph(f"{om_name.upper()} • GABINETE DO ORDENADOR DE DESPESAS", st_header_sub))
    story.append(Spacer(1, 3 * mm))

    story.append(Paragraph("DESPACHO DE JUSTIFICATIVA CIRCUNSTANCIADA DA AUTORIDADE COMPETENTE", st_doc_title))
    story.append(
        Paragraph(
            "Excepcionalidade do Art. 6º, § 5º da Instrução Normativa SEGES/ME nº 65/2021 c/c Art. 23 da Lei nº 14.133/2021",
            st_doc_sub,
        )
    )
    story.append(Spacer(1, 2.5 * mm))
    story.append(HRFlowable(width="100%", thickness=0.75, color=colors.HexColor("#cbd5e1"), spaceBefore=0, spaceAfter=4))

    item_code = report.get("catmat_catser_code", "—")
    item_desc = report.get("item_description", "—")
    unit_str = report.get("target_supply_unit", "UN")
    unit_price = float(report.get("final_estimated_unit_price", 0.0))
    total_price = quantity * unit_price
    valid_quotes = report.get("valid_quotes", [])
    n_quotes = len(valid_quotes)

    # SEÇÃO 1: OBJETO
    story.append(Paragraph("1. IDENTIFICAÇÃO DO OBJETO E VALOR ESTIMADO", st_sec_title))
    story.append(Spacer(1, 1 * mm))

    item_table_data = [
        [Paragraph("Código do Catálogo:", st_cell_bold_right), Paragraph(str(item_code), st_cell_txt)],
        [Paragraph("Descrição do Objeto:", st_cell_bold_right), Paragraph(item_desc, st_cell_txt)],
        [Paragraph("Quantidade Solicitada:", st_cell_bold_right), Paragraph(f"{quantity} {unit_str}", st_cell_txt)],
        [Paragraph("Preço Unit. Estimado:", st_cell_bold_right), Paragraph(_format_brl(unit_price), st_cell_bold)],
        [Paragraph("Valor Total Estimado:", st_cell_bold_right), Paragraph(f"<b>{_format_brl(total_price)}</b>", st_cell_bold)],
    ]
    t_item = Table(item_table_data, colWidths=[5 * cm, 13 * cm])
    t_item.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (0, -1), colors.HexColor("#f1f5f9")),
        ("BACKGROUND", (1, 4), (1, 4), colors.HexColor("#fef3c7")),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#e2e8f0")),
        ("TOPPADDING", (0, 0), (-1, -1), 2.5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 2.5),
    ]))
    story.append(t_item)
    story.append(Spacer(1, 2.5 * mm))

    # SEÇÃO 2: COTAÇÕES
    story.append(Paragraph(f"2. COTAÇÕES VÁLIDAS IDENTIFICADAS ({n_quotes} cotação(ões) obtida(s))", st_sec_title))
    story.append(Spacer(1, 1 * mm))

    quotes_data = [
        [
            Paragraph("Nº", st_cell_hdr),
            Paragraph("Órgão / UASG / UF", st_cell_hdr),
            Paragraph("Nº Compra / Data", st_cell_hdr),
            Paragraph("Fornecedor Vencedor & CNPJ", st_cell_hdr),
            Paragraph("Preço Unit.", st_cell_hdr),
        ]
    ]
    for idx, q in enumerate(valid_quotes):
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
        ("TOPPADDING", (0, 0), (-1, -1), 2),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#f8fafc")]),
    ]))
    story.append(t_quotes)
    story.append(Spacer(1, 2.5 * mm))

    # SEÇÃO 3: MOTIVAÇÃO
    story.append(Paragraph("3. MOTIVAÇÃO E FUNDAMENTAÇÃO CIRCUNSTANCIADA", st_sec_title))
    story.append(Spacer(1, 1 * mm))

    motivos = [
        "<b>I. Da Abrangência da Pesquisa:</b> A pesquisa de preços foi conduzida de forma ampla e irrestrita nos bancos de dados oficiais do Governo Federal (Compras.gov.br e PNCP), abrangendo todo o território nacional no período regulamentar de até 1 ano.",
        f"<b>II. Da Escassez de Registros Similares:</b> Constatou-se a disponibilidade de apenas {n_quotes} cotação(ões) válida(s) apta(s) a compor a cesta de preços para o item especificado, em virtude da sazonalidade, especificidade técnica ou ausência de compras análogas recentes por outros entes.",
        f"<b>III. Da Vantajosidade e Razoabilidade:</b> Os valores apurados ({_format_brl(unit_price)} un.) foram submetidos à análise crítica de exequibilidade e compatibilidade mercadológica, inexistindo indícios de sobrepreço, resguardando a economicidade e vantajosidade (Art. 5º da Lei 14.133/2021).",
        "<b>IV. Da Necessidade Inadiável:</b> A paralisação do procedimento acarretaria prejuízo ou desabastecimento das atividades finalísticas e operacionais desta Organização Militar.",
    ]
    for m_txt in motivos:
        story.append(Paragraph(m_txt, st_body))
        story.append(Spacer(1, 1 * mm))

    story.append(Spacer(1, 1.5 * mm))

    # SEÇÃO 4: DECISÃO
    story.append(Paragraph("4. DECISÃO DA AUTORIDADE COMPETENTE", st_sec_title))
    story.append(Spacer(1, 1 * mm))
    dec_txt = (
        f"Diante do exposto e com fundamento no Art. 6º, § 5º da Instrução Normativa SEGES/ME nº 65/2021, "
        f"<b>AUTORIZO e RATIFICO</b>, em caráter excepcional e devidamente circunstanciado, a adoção do preço estimado "
        f"com base na(s) {n_quotes} cotação(ões) obtida(s), fixando o valor de referência em <b>{_format_brl(unit_price)}</b> "
        f"por unidade (valor total: <b>{_format_brl(total_price)}</b>), determinando o prosseguimento regular da instrução processual."
    )
    story.append(Paragraph(dec_txt, st_body))
    story.append(Spacer(1, 4 * mm))

    # ASSINATURA
    today = datetime.now()
    mes_nome = _get_mes_extenso(today.month)
    data_str = f"{today.day} de {mes_nome} de {today.year}."

    sig_data = [
        [Paragraph(f"<i>{data_str}</i>", st_sig_sub)],
        [Spacer(1, 6 * mm)],
        [Paragraph("____________________________________________________", st_sig_title)],
        [Paragraph(f"<b>{ordenador_nome}</b>", st_sig_title)],
        [Paragraph("Ordenador de Despesas / Autoridade Competente", st_sig_sub)],
        [Paragraph(om_name.upper(), st_sig_sub)],
    ]
    t_sig = Table(sig_data, colWidths=[18 * cm])
    t_sig.setStyle(TableStyle([
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 1),
        ("TOPPADDING", (0, 0), (-1, -1), 1),
    ]))
    story.append(KeepTogether(t_sig))

    doc.build(story)
    return buf.getvalue()


# =============================================================================
# 5. GERADOR DE PLANILHA EXCEL CONSOLIDADA DE LOTE (MÚLTIPLOS ITENS)
# =============================================================================

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import zipfile

def build_consolidated_in65_excel_bytes(
    items_results: list,
    uasg_name: str = "MARINHA DO BRASIL",
) -> bytes:
    """Gera planilha Excel consolidada contendo aba de Resumo Geral e abas individuais por item."""
    wb = openpyxl.Workbook()
    borda = Side(style="thin")
    B = Border(left=borda, right=borda, top=borda, bottom=borda)
    cinza = PatternFill("solid", fgColor="D9D9D9")
    verde = PatternFill("solid", fgColor="E2EFDA")
    amarelo = PatternFill("solid", fgColor="FEF3C7")
    brl = "R$ #,##0.00"
    ctr = Alignment(horizontal="center", vertical="center", wrap_text=True)
    esq = Alignment(horizontal="left", vertical="center", wrap_text=True)

    # ABA 1: RESUMO DO LOTE
    ws_resumo = wb.active
    ws_resumo.title = "Resumo Geral do Lote"

    def merge_line(ws, row_idx, text, bold=True, align="center", height=24, end_col="J"):
        ws.merge_cells(f"A{row_idx}:{end_col}{row_idx}")
        ws[f"A{row_idx}"] = text
        ws[f"A{row_idx}"].font = Font(bold=bold, size=11 if bold else 10)
        ws[f"A{row_idx}"].alignment = Alignment(horizontal=align, vertical="center", wrap_text=True)
        ws.row_dimensions[row_idx].height = height

    merge_line(ws_resumo, 1, uasg_name, bold=True)
    merge_line(ws_resumo, 2, "MAPA COMPARATIVO CONSOLIDADO DE PREÇOS (IN SEGES/ME Nº 65/2021)", bold=True)
    merge_line(ws_resumo, 3, "Quadro Demonstrativo de Todos os Itens Demandados e Valores Homologáveis:", bold=True, align="left")

    headers = [
        "ITEM", "CATMAT/CATSER", "DESCRIÇÃO DO OBJETO", "UNID.", "QTDE",
        "MÉDIA SANEADA", "MEDIANA SANEADA", "PREÇO UNIT. ESTIMADO", "PREÇO TOTAL ESTIMADO", "SITUAÇÃO IN 65"
    ]
    for col_idx, h in enumerate(headers, 1):
        c = ws_resumo.cell(4, col_idx, h)
        c.font, c.fill, c.alignment, c.border = Font(bold=True, size=10), cinza, ctr, B
    ws_resumo.row_dimensions[4].height = 28

    total_geral_lote = 0.0

    for idx, item in enumerate(items_results, start=1):
        rep = item.get("report") or {}
        qty = int(item.get("qty", 1) or 1)
        sm = rep.get("sanitized_metrics") or {}
        unit_est = float(rep.get("final_estimated_unit_price", 0.0))
        item_total = qty * unit_est
        total_geral_lote += item_total

        sit = "✅ Homogênea (CV ≤ 25%)" if rep.get("is_homogeneous") else "⚠️ Heterogênea"
        if len(rep.get("valid_quotes", [])) < 3:
            sit = f"⚠️ Exige Justificativa ({len(rep.get('valid_quotes', []))} cotações)"

        row_vals = [
            idx,
            str(rep.get("catmat_catser_code", item.get("code", "—"))),
            rep.get("item_description", "Item"),
            rep.get("target_supply_unit", item.get("unit", "UN")),
            qty,
            float(sm.get("mean") or unit_est),
            float(sm.get("median") or unit_est),
            unit_est,
            item_total,
            sit
        ]

        row_num = 4 + idx
        for col_idx, val in enumerate(row_vals, 1):
            c = ws_resumo.cell(row_num, col_idx, val)
            c.border = B
            c.alignment = esq if col_idx == 3 else ctr
            if col_idx in (6, 7, 8, 9):
                c.number_format = brl
            if col_idx in (8, 9):
                c.fill = verde
                c.font = Font(bold=True)
        ws_resumo.row_dimensions[row_num].height = 22

    # LINHA TOTALIZADORA
    tot_row = 5 + len(items_results)
    ws_resumo.merge_cells(f"A{tot_row}:H{tot_row}")
    c_tot_label = ws_resumo.cell(tot_row, 1, "VALOR TOTAL ESTIMADO DO LOTE:")
    c_tot_label.font = Font(bold=True, size=11)
    c_tot_label.alignment = Alignment(horizontal="right", vertical="center")
    c_tot_label.fill = amarelo

    c_tot_val = ws_resumo.cell(tot_row, 9, total_geral_lote)
    c_tot_val.font = Font(bold=True, size=12)
    c_tot_val.alignment = ctr
    c_tot_val.fill = verde
    c_tot_val.number_format = brl
    c_tot_val.border = B

    for col in range(1, 11):
        ws_resumo.cell(tot_row, col).border = B

    # AJUSTE DE LARGURAS
    col_widths = [8, 16, 45, 10, 10, 16, 16, 20, 20, 25]
    for col_idx, w in enumerate(col_widths, 1):
        ws_resumo.column_dimensions[openpyxl.utils.get_column_letter(col_idx)].width = w

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# =============================================================================
# 6. GERADOR DE PACOTE ZIP COMPLETO DO LOTE
# =============================================================================

def build_batch_zip_bytes(
    items_results: list,
    om_name: str = "MARINHA DO BRASIL",
    ordenador_nome: str = "ORDENADOR DE DESPESAS",
    elaborador_nome: str = "ENCARREGADO DO SETOR DE OBTENÇÃO",
    extract_gov_docs: bool = True,
) -> bytes:
    """Gera um arquivo .ZIP contendo os relatórios em Word, PDF e Excel individuais de cada item, além da planilha consolidada e dos Editais/Termos oficiais extraídos do Governo (PNCP)."""
    import concurrent.futures
    from pncp_extractor import capture_pncp_page_pdf, CACHE_DIR

    # 1. Extração concorrente de alta eficiência dos documentos oficiais do Governo (PNCP)
    pncp_docs_map = {}
    if extract_gov_docs:
        tasks = {}
        for item in items_results:
            rep = item.get("report") or {}
            for q in rep.get("valid_quotes", []):
                p_key = q.get("purchase_key") or ""
                p_url = q.get("pncp_or_compras_uri") or ""
                c_num = str(q.get("purchase_number_year") or "compra").replace("/", "_")
                if p_key and p_key not in tasks:
                    tasks[p_key] = (p_url, p_key, c_num)

        def _fetch_one_doc(entry):
            k, (url, key, c_num) = entry
            try:
                clean_id = re.sub(r"\W", "_", key)
                c_file = os.path.join(CACHE_DIR, f"pncp_full_{clean_id}.pdf")
                if os.path.exists(c_file) and os.path.getsize(c_file) > 1000:
                    with open(c_file, "rb") as cf:
                        return k, c_num, cf.read()
                pdf_b = capture_pncp_page_pdf(url, chave_or_id=key, timeout_sec=12)
                if pdf_b and len(pdf_b) > 1000:
                    return k, c_num, pdf_b
            except Exception as e:
                print(f"[PNCP Parallel Fetch Error] {key}: {e}")
            return k, c_num, None

        if tasks:
            with concurrent.futures.ThreadPoolExecutor(max_workers=6) as executor:
                for k, c_num, pdf_b in executor.map(_fetch_one_doc, list(tasks.items())):
                    if pdf_b:
                        pncp_docs_map[k] = (c_num, pdf_b)

    z_buf = io.BytesIO()

    with zipfile.ZipFile(z_buf, "w", zipfile.ZIP_DEFLATED) as zf:
        # 1. Planilha Consolidada Geral
        xlsx_consolidada = build_consolidated_in65_excel_bytes(items_results, uasg_name=(f"{om_name}" if "MARINHA" in om_name.upper() else f"MARINHA DO BRASIL - {om_name}"))
        zf.writestr("00_Mapa_Comparativo_Consolidado_Lote.xlsx", xlsx_consolidada)

        # Pasta raiz dedicada aos Editais Oficiais do PNCP coletados
        for p_key, (c_num, pdf_b) in pncp_docs_map.items():
            zf.writestr(f"01_Editais_Oficiais_Gov_PNCP/PNCP_Edital_Compra_{c_num}.pdf", pdf_b)

        # 2. Itens individuais
        for idx, item in enumerate(items_results, start=1):
            rep = item.get("report") or {}
            qty = int(item.get("qty", 1) or 1)
            code = str(rep.get("catmat_catser_code", item.get("code", "000000")))
            item_type = item.get("type", "MATERIAL")
            prefix = "CATSER" if str(item_type).upper() == "SERVICO" else "CATMAT"
            folder = f"Item_{idx:02d}_{prefix}_{code}"

            # Word DOCX individual
            docx_b = build_in65_docx_bytes(rep, qty, om_name=om_name, elaborador_nome=elaborador_nome, ordenador_nome=ordenador_nome)
            zf.writestr(f"{folder}/Relatorio_Pesquisa_Precos_{prefix}_{code}.docx", docx_b)

            # PDF individual
            pdf_b = build_in65_pdf_bytes(rep, qty, om_name=om_name, elaborador_nome=elaborador_nome, ordenador_nome=ordenador_nome)
            zf.writestr(f"{folder}/Relatorio_Pesquisa_Precos_{prefix}_{code}.pdf", pdf_b)

            # Justificativa Circunstanciada (se < 3 cotações)
            if len(rep.get("valid_quotes", [])) < 3:
                just_docx = build_in65_justificativa_docx_bytes(rep, qty, om_name=om_name, ordenador_nome=ordenador_nome)
                zf.writestr(f"{folder}/Justificativa_Circunstanciada_Art6_IN65_{prefix}_{code}.docx", just_docx)
                just_pdf = build_in65_justificativa_pdf_bytes(rep, qty, om_name=om_name, ordenador_nome=ordenador_nome)
                zf.writestr(f"{folder}/Justificativa_Circunstanciada_Art6_IN65_{prefix}_{code}.pdf", just_pdf)

            # Inclusão dos Editais/Documentos PNCP de cada cotação do item
            quotes = rep.get("valid_quotes", [])
            seen_item_keys = set()
            for q in quotes:
                p_key = q.get("purchase_key") or ""
                if p_key in pncp_docs_map and p_key not in seen_item_keys:
                    seen_item_keys.add(p_key)
                    c_num, pdf_b = pncp_docs_map[p_key]
                    zf.writestr(f"{folder}/Editais_Gov_PNCP/PNCP_Edital_Compra_{c_num}.pdf", pdf_b)

    return z_buf.getvalue()


# =============================================================================
# 7. CERTIFICADO OFICIAL DE TRILHA DE AUDITORIA IMUTÁVEL (SHA-256 / TCU)
# =============================================================================
def build_audit_certificate_pdf_bytes(
    om_name: str = "MARINHA DO BRASIL",
    user_name: str = "OPERADOR DO SETOR DE OBTENÇÃO",
) -> bytes:
    """
    Gera o Certificado de Trilha de Auditoria Imutável em formato PDF A4 oficial.
    Comprova perante TCU, CGU e CISET a idoneidade, rastreabilidade e imutabilidade dos cálculos.
    """
    from audit_trail import audit_trail_manager

    integrity = audit_trail_manager.verify_integrity()
    entries = audit_trail_manager.get_recent_entries(limit=40)

    buf = io.BytesIO()
    doc = SimpleDocTemplate(
        buf,
        pagesize=A4,
        leftMargin=1.8 * cm,
        rightMargin=1.8 * cm,
        topMargin=1.8 * cm,
        bottomMargin=1.8 * cm,
    )

    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(
        "CertTitle",
        parent=styles["Normal"],
        fontName="Helvetica-Bold",
        fontSize=12,
        leading=15,
        alignment=1,  # Center
        textColor=colors.HexColor("#1E2D3D"),
    )
    subtitle_style = ParagraphStyle(
        "CertSubtitle",
        parent=styles["Normal"],
        fontName="Helvetica",
        fontSize=9,
        leading=12,
        alignment=1,
        textColor=colors.HexColor("#475569"),
    )
    cell_style = ParagraphStyle(
        "CertCell",
        parent=styles["Normal"],
        fontName="Helvetica",
        fontSize=7.5,
        leading=9.5,
        textColor=colors.HexColor("#1E293B"),
    )
    cell_mono = ParagraphStyle(
        "CertMono",
        parent=styles["Normal"],
        fontName="Courier",
        fontSize=6.5,
        leading=8,
        textColor=colors.HexColor("#334155"),
    )
    cell_bold = ParagraphStyle(
        "CertCellBold",
        parent=styles["Normal"],
        fontName="Helvetica-Bold",
        fontSize=7.5,
        leading=9.5,
        textColor=colors.HexColor("#1E293B"),
    )

    elements = []

    # 1. Brasão Oficial se existir
    if os.path.exists(BRASAO_PATH):
        try:
            elements.append(RLImage(BRASAO_PATH, width=2.0 * cm, height=2.0 * cm))
            elements.append(Spacer(1, 4))
        except Exception:
            pass

    # 2. Cabeçalho Oficial
    elements.append(Paragraph(om_name.upper(), title_style))
    elements.append(Paragraph("SISTEMA GOVTECH NLLC • MÓDULO DE COMPRAS PÚBLICAS & AUDITORIA", subtitle_style))
    elements.append(Spacer(1, 4))
    elements.append(Paragraph("CERTIFICADO DE TRILHA DE AUDITORIA E INTEGRIDADE CRIPTOGRÁFICA", title_style))
    elements.append(Paragraph("Conformidade com a Instrução Normativa SEGES/ME nº 65/2021 e Lei nº 14.133/2021", subtitle_style))
    elements.append(Spacer(1, 10))

    # 3. Painel de Status de Integridade Criptográfica
    is_valid = integrity.get("valid", True)
    status_text = "INTEGRIDADE CRIPTOGRÁFICA REGULAR • ZERO VIOLAÇÕES DETECTADAS" if is_valid else "ALERTA: QUEBRA DE INTEGRIDADE DETECTADA"
    status_bg = colors.HexColor("#D1FAE5") if is_valid else colors.HexColor("#FEE2E2")
    status_fg = colors.HexColor("#065F46") if is_valid else colors.HexColor("#991B1B")

    info_data = [
        [
            Paragraph("<b>SITUAÇÃO DO ENCADIAMENTO:</b>", cell_style),
            Paragraph(f"<b><font color='{status_fg.hexval()}'>{status_text}</font></b>", cell_style),
        ],
        [
            Paragraph("<b>ÚLTIMO HASH SHA-256 (BLOCKCHAIN TRAIL):</b>", cell_style),
            Paragraph(f"<code>{integrity.get('latest_hash', 'N/A')}</code>", cell_mono),
        ],
        [
            Paragraph("<b>TOTAL DE EVENTOS ENCADERNADOS:</b>", cell_style),
            Paragraph(f"{integrity.get('total_entries', len(entries))} registros imutáveis gravados", cell_style),
        ],
        [
            Paragraph("<b>AGENTE RESPONSÁVEL / OPERADOR:</b>", cell_style),
            Paragraph(user_name, cell_style),
        ],
        [
            Paragraph("<b>DATA E HORA DA CERTIFICAÇÃO (UTC):</b>", cell_style),
            Paragraph(datetime.now(timezone.utc).strftime("%d/%m/%Y às %H:%M:%S UTC"), cell_style),
        ],
    ]
    t_info = Table(info_data, colWidths=[6.0 * cm, 11.4 * cm])
    t_info.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#F8FAFC")),
        ("BACKGROUND", (0, 0), (-1, 0), status_bg),
        ("BOX", (0, 0), (-1, -1), 1, colors.HexColor("#CBD5E1")),
        ("INNERGRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#E2E8F0")),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("PADDING", (0, 0), (-1, -1), 5),
    ]))
    elements.append(t_info)
    elements.append(Spacer(1, 12))

    # 4. Tabela de Registros da Trilha de Auditoria
    elements.append(Paragraph("<b>HISTÓRICO RECENTE DE EVENTOS AUDITÁVEIS ENCADEADOS:</b>", cell_bold))
    elements.append(Spacer(1, 4))

    table_data = [
        [
            Paragraph("<b>#</b>", cell_bold),
            Paragraph("<b>Data/Hora (UTC)</b>", cell_bold),
            Paragraph("<b>Ação Operacional</b>", cell_bold),
            Paragraph("<b>Alvo / Código</b>", cell_bold),
            Paragraph("<b>Base Legal / Detalhes</b>", cell_bold),
            Paragraph("<b>Hash SHA-256</b>", cell_bold),
        ]
    ]

    for e in entries[:25]:
        det = e.get("details")
        det_str = json.dumps(det, ensure_ascii=False) if isinstance(det, dict) else str(det or "")
        if len(det_str) > 70:
            det_str = det_str[:67] + "..."
        h = e.get("sha256_hash", "")
        h_short = f"{h[:10]}...{h[-8:]}" if len(h) >= 18 else h
        t_stamp = e.get("timestamp", "").replace("T", " ")[:19]

        table_data.append([
            Paragraph(str(e.get("sequence", "")), cell_style),
            Paragraph(t_stamp, cell_mono),
            Paragraph(str(e.get("action", "")), cell_style),
            Paragraph(str(e.get("target_item", "")), cell_style),
            Paragraph(f"<b>{e.get('legal_basis', '')}:</b> {det_str}", cell_style),
            Paragraph(f"<code>{h_short}</code>", cell_mono),
        ])

    t_entries = Table(table_data, colWidths=[0.9 * cm, 2.7 * cm, 3.2 * cm, 2.4 * cm, 5.4 * cm, 2.8 * cm])
    t_entries.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1E2D3D")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#F8FAFC")]),
        ("BOX", (0, 0), (-1, -1), 1, colors.HexColor("#CBD5E1")),
        ("INNERGRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#E2E8F0")),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("PADDING", (0, 0), (-1, -1), 4),
    ]))
    elements.append(t_entries)
    elements.append(Spacer(1, 14))

    # 5. Declaração Legal e Assinatura
    decl_p = (
        "<b>DECLARAÇÃO DE IDONEIDADE PROCESSUAL:</b> Atesta-se que todos os registros acima transcritos "
        "foram gerados de forma automática e vinculados a um encadeamento criptográfico irreversível (SHA-256). "
        "A integridade dos registros comprova a inexistência de alteração arbitrária de preços, exclusões não motivadas "
        "ou manipulações nos cálculos de tendência central, assegurando conformidade com o Art. 19 da Lei nº 14.133/2021."
    )
    elements.append(Paragraph(decl_p, cell_style))
    elements.append(Spacer(1, 24))

    # Linhas de Assinatura
    sig_data = [
        [
            Paragraph("____________________________________________<br/><b>AGENTE DE CONTRATAÇÃO / PREGOEIRO</b>", subtitle_style),
            Paragraph("____________________________________________<br/><b>ENCARREGADO DO SETOR DE OBTENÇÃO</b>", subtitle_style),
        ]
    ]
    t_sig = Table(sig_data, colWidths=[8.7 * cm, 8.7 * cm])
    t_sig.setStyle(TableStyle([
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
    ]))
    elements.append(t_sig)

    doc.build(elements)
    return buf.getvalue()

