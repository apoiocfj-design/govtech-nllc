"""
Gerador de Ofício Formal de Manifestação de Interesse (IRP) / Pedido de Adesão à Ata SRP (Carona)
nos termos do Art. 86, § 2º da Lei nº 14.133/2021 e Decreto Federal nº 11.462/2023.
Emite versões idênticas em Word (.docx) e PDF (.pdf) com cabeçalho oficial padronizado.
"""

from __future__ import annotations

import io
from datetime import datetime
from typing import Any, Dict, List

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Cm, Pt, RGBColor
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_JUSTIFY, TA_LEFT, TA_RIGHT
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import cm
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle


def _fmt_brl(val: float) -> str:
    try:
        v = float(val or 0.0)
    except Exception:
        v = 0.0
    s = f"{v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return f"R$ {s}"


def _build_oficio_context(payload: Dict[str, Any]) -> Dict[str, Any]:
    now_str = datetime.now().strftime("%d/%m/%Y")
    om_solicitante = (payload.get("om_solicitante") or "ORGANIZAÇÃO MILITAR / ÓRGÃO SOLICITANTE").strip()
    uasg_solicitante = (payload.get("uasg_solicitante") or "").strip()
    ordenador = (payload.get("ordenador_nome") or "ORDENADOR DE DESPESAS").strip()
    elaborador = (payload.get("elaborador_nome") or "ENCARREGADO DO SETOR DE OBTENÇÃO").strip()

    orgao_gerenciador = (payload.get("orgao_nome") or "ÓRGÃO GERENCIADOR").strip()
    cnpj_gerenciador = (payload.get("cnpj") or "").strip()
    unidade_gerenciadora = (payload.get("unidade_nome") or "").strip()
    uf_gerenciador = (payload.get("uf") or "").strip()
    municipio_gerenciador = (payload.get("municipio") or "").strip()
    doc_kind = str(payload.get("resolved_kind") or payload.get("document_type") or "irp").lower()
    ano = str(payload.get("ano") or datetime.now().year)
    seq = str(payload.get("sequencial") or "")
    pncp_id = (payload.get("numero_controle_pncp") or f"{cnpj_gerenciador}-1-{seq}/{ano}").strip()
    objeto = (payload.get("objeto") or "Registro de Preços de interesse da Administração.").strip()
    link_pncp = (payload.get("link_oficial_pncp") or "").strip()
    itens = payload.get("itens") or []

    is_carona = doc_kind in {"ata", "arp", "edital"}
    titulo_doc = (
        "OFÍCIO DE PEDIDO DE ADESÃO À ATA DE REGISTRO DE PREÇOS (CARONA)"
        if is_carona
        else "OFÍCIO DE MANIFESTAÇÃO DE INTERESSE EM INTENÇÃO DE REGISTRO DE PREÇOS (IRP)"
    )
    assunto = (
        f"Solicitação de autorização para adesão (órgão não participante) à Ata/SRP — ID PNCP {pncp_id}"
        if is_carona
        else f"Manifestação de interesse como órgão participante na IRP nº {seq}/{ano} — ID PNCP {pncp_id}"
    )

    total_estimado = 0.0
    norm_itens: List[Dict[str, Any]] = []
    for idx, it in enumerate(itens[:40], start=1):
        qtd = float(it.get("quantidade") or 1.0)
        vu = float(it.get("valor_unitario_estimado") or 0.0)
        vt = qtd * vu
        total_estimado += vt
        norm_itens.append({
            "num": it.get("numero_item") or idx,
            "desc": str(it.get("descricao") or f"Item {idx}")[:220],
            "unid": str(it.get("unidade_fornecimento") or "UN"),
            "qtd": qtd,
            "vu": vu,
            "vt": vt,
        })

    return {
        "now_str": now_str,
        "om_solicitante": om_solicitante,
        "uasg_solicitante": uasg_solicitante,
        "ordenador": ordenador,
        "elaborador": elaborador,
        "orgao_gerenciador": orgao_gerenciador,
        "cnpj_gerenciador": cnpj_gerenciador,
        "unidade_gerenciadora": unidade_gerenciadora,
        "uf_gerenciador": uf_gerenciador,
        "municipio_gerenciador": municipio_gerenciador,
        "pncp_id": pncp_id,
        "objeto": objeto,
        "link_pncp": link_pncp,
        "is_carona": is_carona,
        "titulo_doc": titulo_doc,
        "assunto": assunto,
        "itens": norm_itens,
        "total_estimado": total_estimado,
    }


def build_irp_oficio_docx_bytes(payload: Dict[str, Any]) -> bytes:
    """Gera o Ofício de Manifestação de Interesse / Pedido de Adesão (Carona) em Word (.docx)."""
    ctx = _build_oficio_context(payload)
    doc = Document()

    for sec in doc.sections:
        sec.top_margin = Cm(2.0)
        sec.bottom_margin = Cm(2.0)
        sec.left_margin = Cm(2.5)
        sec.right_margin = Cm(2.0)

    style = doc.styles["Normal"]
    style.font.name = "Calibri"
    style.font.size = Pt(10.5)

    p_hdr = doc.add_paragraph()
    p_hdr.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r1 = p_hdr.add_run(f"{ctx['om_solicitante'].upper()}\n")
    r1.bold = True
    r1.font.size = Pt(12)
    r1.font.color.rgb = RGBColor(30, 45, 61)
    if ctx["uasg_solicitante"]:
        r2 = p_hdr.add_run(f"UASG: {ctx['uasg_solicitante']}  |  ")
        r2.font.size = Pt(9.5)
    r3 = p_hdr.add_run("PLANEJAMENTO DA CONTRATAÇÃO — LEI Nº 14.133/2021")
    r3.font.size = Pt(9.5)
    r3.font.color.rgb = RGBColor(90, 100, 115)

    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    rt = p_title.add_run(ctx["titulo_doc"])
    rt.bold = True
    rt.font.size = Pt(11.5)
    rt.font.color.rgb = RGBColor(30, 45, 61)

    p_date = doc.add_paragraph()
    p_date.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p_date.add_run(f"Em {ctx['now_str']}.")

    p_dest = doc.add_paragraph()
    p_dest.add_run("Ao Senhor(a) Ordenador de Despesas / Autoridade Competente\n").bold = True
    p_dest.add_run(f"{ctx['orgao_gerenciador']}\n")
    if ctx["unidade_gerenciadora"]:
        p_dest.add_run(f"Unidade Gerenciadora: {ctx['unidade_gerenciadora']}\n")
    loc = f"{ctx['municipio_gerenciador']}/{ctx['uf_gerenciador']}".strip("/")
    p_dest.add_run(f"CNPJ: {ctx['cnpj_gerenciador']}  |  Localidade: {loc or 'Brasil'}\n")

    p_assunto = doc.add_paragraph()
    p_assunto.add_run("Assunto: ").bold = True
    p_assunto.add_run(ctx["assunto"])

    p_ref = doc.add_paragraph()
    p_ref.add_run("Referência Legal: ").bold = True
    p_ref.add_run("Art. 86, § 2º e § 3º da Lei nº 14.133/2021 e Decreto Federal nº 11.462/2023.")

    p_body1 = doc.add_paragraph()
    p_body1.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    if ctx["is_carona"]:
        p_body1.add_run(
            f"1. Cumprimentando-o(a) cordialmente, solicito a Vossa Senhoria autorização para que "
            f"{ctx['om_solicitante']} realize ADESÃO (na condição de órgão não participante) à Ata de Registro "
            f"de Preços vinculada à contratação identificada no Portal Nacional de Contratações Públicas (PNCP) "
            f"sob o nº {ctx['pncp_id']}, cujo objeto é: \"{ctx['objeto']}\"."
        )
    else:
        p_body1.add_run(
            f"1. Cumprimentando-o(a) cordialmente, manifesto formalmente o interesse de {ctx['om_solicitante']} "
            f"em figurar como ÓRGÃO PARTICIPANTE na Intenção de Registro de Preços (IRP) identificada no PNCP "
            f"sob o nº {ctx['pncp_id']}, cujo objeto é: \"{ctx['objeto']}\"."
        )

    p_body2 = doc.add_paragraph()
    p_body2.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p_body2.add_run(
        "2. A presente solicitação fundamenta-se na vantajosidade econômica e operacional da contratação "
        "compartilhada, em observância aos princípios da eficiência, economicidade e planejamento "
        "estabelecidos na Lei nº 14.133/2021 e no Decreto nº 11.462/2023, comprometendo-se este Órgão "
        "a observar integralmente os limites quantitativos legais e as condições pactuadas no instrumento convocatório."
    )

    if ctx["itens"]:
        p_it = doc.add_paragraph()
        p_it.add_run("3. Relação de Itens e Quantitativos Estimados:").bold = True
        tbl = doc.add_table(rows=1, cols=5)
        tbl.style = "Table Grid"
        hdrs = ["Item", "Descrição do Objeto", "Unid.", "Qtd.", "Valor Unit. Est."]
        for i, h in enumerate(hdrs):
            cell = tbl.rows[0].cells[i]
            cell.text = h
            for run in cell.paragraphs[0].runs:
                run.bold = True
                run.font.size = Pt(9)
        for it in ctx["itens"]:
            row = tbl.add_row().cells
            row[0].text = str(it["num"])
            row[1].text = it["desc"]
            row[2].text = it["unid"]
            row[3].text = f"{it['qtd']:g}"
            row[4].text = _fmt_brl(it["vu"]) if it["vu"] > 0 else "A cotar"
            for c in row:
                for run in c.paragraphs[0].runs:
                    run.font.size = Pt(8.5)

        if ctx["total_estimado"] > 0:
            p_tot = doc.add_paragraph()
            p_tot.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            r_tot = p_tot.add_run(f"Valor Total Estimado da Demanda: {_fmt_brl(ctx['total_estimado'])}")
            r_tot.bold = True

    if ctx["link_pncp"]:
        p_lnk = doc.add_paragraph()
        p_lnk.add_run("Link Oficial no PNCP: ").bold = True
        p_lnk.add_run(ctx["link_pncp"])

    p_close = doc.add_paragraph()
    p_close.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p_close.add_run(
        "4. Colocando-nos à disposição para quaisquer esclarecimentos adicionais, renovo protestos "
        "de elevada estima e distinta consideração."
    )

    doc.add_paragraph("\n")
    p_sig = doc.add_paragraph()
    p_sig.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_sig.add_run("_______________________________________________________\n")
    rs1 = p_sig.add_run(f"{ctx['ordenador']}\n")
    rs1.bold = True
    p_sig.add_run(f"{ctx['om_solicitante']}\n")
    p_sig.add_run(f"Elaborado por: {ctx['elaborador']}").font.size = Pt(8.5)

    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()


def build_irp_oficio_pdf_bytes(payload: Dict[str, Any]) -> bytes:
    """Gera o Ofício de Manifestação de Interesse / Pedido de Adesão (Carona) em PDF pronto para SEI."""
    ctx = _build_oficio_context(payload)
    buf = io.BytesIO()
    doc = SimpleDocTemplate(
        buf,
        pagesize=A4,
        leftMargin=2.2 * cm,
        rightMargin=2.0 * cm,
        topMargin=2.0 * cm,
        bottomMargin=2.0 * cm,
    )
    styles = getSampleStyleSheet()
    st_hdr = ParagraphStyle("Hdr", parent=styles["Normal"], fontName="Helvetica-Bold", fontSize=11, alignment=TA_CENTER, textColor=colors.HexColor("#1e2d3d"))
    st_sub = ParagraphStyle("Sub", parent=styles["Normal"], fontName="Helvetica", fontSize=8.5, alignment=TA_CENTER, textColor=colors.HexColor("#5a6473"))
    st_title = ParagraphStyle("TitleDoc", parent=styles["Normal"], fontName="Helvetica-Bold", fontSize=10.5, alignment=TA_CENTER, spaceBefore=10, spaceAfter=10, textColor=colors.HexColor("#1e2d3d"))
    st_right = ParagraphStyle("Right", parent=styles["Normal"], fontName="Helvetica", fontSize=9.5, alignment=TA_RIGHT, spaceAfter=8)
    st_body = ParagraphStyle("Body", parent=styles["Normal"], fontName="Helvetica", fontSize=9.5, leading=13.5, alignment=TA_JUSTIFY, spaceAfter=7)
    st_left = ParagraphStyle("Left", parent=styles["Normal"], fontName="Helvetica", fontSize=9.5, leading=13, alignment=TA_LEFT, spaceAfter=6)
    st_cell = ParagraphStyle("Cell", parent=styles["Normal"], fontName="Helvetica", fontSize=8, leading=10)

    story: List[Any] = [
        Paragraph(ctx["om_solicitante"].upper(), st_hdr),
        Paragraph("PLANEJAMENTO DA CONTRATAÇÃO — LEI Nº 14.133/2021 & DECRETO Nº 11.462/2023", st_sub),
        Paragraph(ctx["titulo_doc"], st_title),
        Paragraph(f"Em {ctx['now_str']}.", st_right),
        Paragraph(
            f"<b>Ao Senhor(a) Ordenador de Despesas / Autoridade Competente</b><br/>"
            f"{ctx['orgao_gerenciador']}<br/>"
            f"{('Unidade Gerenciadora: ' + ctx['unidade_gerenciadora'] + '<br/>') if ctx['unidade_gerenciadora'] else ''}"
            f"CNPJ: {ctx['cnpj_gerenciador']} | Localidade: {ctx['municipio_gerenciador']}/{ctx['uf_gerenciador']}",
            st_left,
        ),
        Paragraph(f"<b>Assunto:</b> {ctx['assunto']}", st_left),
        Paragraph("<b>Referência Legal:</b> Art. 86, § 2º e § 3º da Lei nº 14.133/2021 e Decreto Federal nº 11.462/2023.", st_left),
        Spacer(1, 4),
    ]

    if ctx["is_carona"]:
        story.append(Paragraph(
            f"1. Cumprimentando-o(a) cordialmente, solicito a Vossa Senhoria autorização para que "
            f"<b>{ctx['om_solicitante']}</b> realize <b>ADESÃO (órgão não participante / carona)</b> à Ata de Registro "
            f"de Preços vinculada à contratação identificada no PNCP sob o nº <b>{ctx['pncp_id']}</b>, "
            f"cujo objeto é: \"{ctx['objeto']}\".",
            st_body,
        ))
    else:
        story.append(Paragraph(
            f"1. Cumprimentando-o(a) cordialmente, manifesto formalmente o interesse de <b>{ctx['om_solicitante']}</b> "
            f"em figurar como <b>ÓRGÃO PARTICIPANTE</b> na Intenção de Registro de Preços (IRP) identificada no PNCP "
            f"sob o nº <b>{ctx['pncp_id']}</b>, cujo objeto é: \"{ctx['objeto']}\".",
            st_body,
        ))

    story.append(Paragraph(
        "2. A presente solicitação fundamenta-se na vantajosidade econômica e operacional da contratação "
        "compartilhada, em observância aos princípios da eficiência, economicidade e planejamento "
        "estabelecidos na Lei nº 14.133/2021 e no Decreto nº 11.462/2023, comprometendo-se este Órgão "
        "a observar integralmente os limites quantitativos legais e as condições pactuadas no instrumento convocatório.",
        st_body,
    ))

    if ctx["itens"]:
        story.append(Paragraph("<b>3. Relação de Itens e Quantitativos Estimados:</b>", st_left))
        tdata = [["Item", "Descrição do Objeto", "Unid.", "Qtd.", "Valor Unit."]]
        for it in ctx["itens"]:
            tdata.append([
                str(it["num"]),
                Paragraph(it["desc"], st_cell),
                it["unid"],
                f"{it['qtd']:g}",
                _fmt_brl(it["vu"]) if it["vu"] > 0 else "A cotar",
            ])
        t = Table(tdata, colWidths=[1.1 * cm, 9.6 * cm, 1.5 * cm, 1.6 * cm, 2.8 * cm], repeatRows=1)
        t.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1e2d3d")),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
            ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
            ("FONTSIZE", (0, 0), (-1, 0), 8),
            ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#cbd5e1")),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("ALIGN", (0, 0), (0, -1), "CENTER"),
            ("ALIGN", (2, 0), (4, -1), "RIGHT"),
            ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#f8fafc")]),
        ]))
        story.append(t)
        story.append(Spacer(1, 6))
        if ctx["total_estimado"] > 0:
            story.append(Paragraph(f"<b>Valor Total Estimado da Demanda: {_fmt_brl(ctx['total_estimado'])}</b>", st_right))

    if ctx["link_pncp"]:
        story.append(Paragraph(f"<b>Link Oficial no PNCP:</b> {ctx['link_pncp']}", st_left))

    story.append(Paragraph(
        "4. Colocando-nos à disposição para quaisquer esclarecimentos adicionais, renovo protestos "
        "de elevada estima e distinta consideração.",
        st_body,
    ))
    story.append(Spacer(1, 22))
    story.append(Paragraph("_______________________________________________________", st_hdr))
    story.append(Paragraph(f"<b>{ctx['ordenador']}</b>", st_hdr))
    story.append(Paragraph(ctx["om_solicitante"], st_sub))
    story.append(Paragraph(f"Elaborado por: {ctx['elaborador']}", st_sub))

    doc.build(story)
    return buf.getvalue()
