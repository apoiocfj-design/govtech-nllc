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
from datetime import datetime, date, timezone
from typing import Dict, Any, List, Optional
from PIL import Image as PILImage, ImageDraw as PILImageDraw, ImageFont as PILImageFont
import qrcode
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



def _get_pil_font(size: int, bold: bool = False):
    """Carrega fontes TrueType escaláveis para nitidez absoluta em impressões e telas."""
    font_candidates = (
        ["arialbd.ttf", "calibrib.ttf", "DejaVuSans-Bold.ttf", "segoeuib.ttf"]
        if bold else
        ["arial.ttf", "calibri.ttf", "DejaVuSans.ttf", "segoeui.ttf"]
    )
    for name in font_candidates:
        try:
            return PILImageFont.truetype(name, size)
        except Exception:
            continue
    try:
        return PILImageFont.load_default(size=size)
    except Exception:
        return PILImageFont.load_default()


def generate_qr_code_bytes(url: str, box_size: int = 4, border: int = 1) -> bytes:
    """Gera bytes PNG de um QR Code nítido apontando para a URL informada."""
    if not url or not url.startswith("http"):
        url = "https://pncp.gov.br"
    qr = qrcode.QRCode(
        version=1,
        error_correction=qrcode.constants.ERROR_CORRECT_M,
        box_size=box_size,
        border=border,
    )
    qr.add_data(url)
    qr.make(fit=True)
    img = qr.make_image(fill_color="#0f172a", back_color="#ffffff")
    buf = io.BytesIO()
    img.save(buf, format="PNG")
    return buf.getvalue()


def add_docx_hyperlink(paragraph, url: str, text: str, color: str = "1D4ED8", underline: bool = True):
    """Adiciona um link clicável nativo a um parágrafo no python-docx."""
    if not url or not url.startswith("http"):
        url = "https://pncp.gov.br"
    part = paragraph.part
    r_id = part.relate_to(url, docx.opc.constants.RELATIONSHIP_TYPE.HYPERLINK, is_external=True)

    hyperlink = OxmlElement('w:hyperlink')
    hyperlink.set(qn('r:id'), r_id)

    new_run = OxmlElement('w:r')
    rPr = OxmlElement('w:rPr')

    if color:
        c = OxmlElement('w:color')
        c.set(qn('w:val'), color)
        rPr.append(c)

    if underline:
        u = OxmlElement('w:u')
        u.set(qn('w:val'), 'single')
        rPr.append(u)

    new_run.append(rPr)
    new_run.text = text
    hyperlink.append(new_run)
    paragraph._p.append(hyperlink)
    return hyperlink

def generate_statistical_infographic(
    valid_prices: List[float],
    purged_prices: List[float],
    cv_percent: float,
    mean_val: float,
    median_val: float,
    std_val: float,
    adopted_param: str = "MEDIANA",
    width: int = 1400,
    height: int = 620,
) -> bytes:
    """
    Gera infográfico estatístico de alta resolução com fontes TrueType escaláveis (300 DPI),
    garantindo nitidez absoluta e legibilidade perfeita na impressão e em telas:
      1. Boxplot dos Quartis e dispersão da amostra (IN SEGES/ME nº 65/2021)
      2. Curva Normal Gaussiana e critérios de decisão legal (Média vs Mediana)
    """
    img = PILImage.new("RGB", (width, height), color="#ffffff")
    draw = PILImageDraw.Draw(img)

    f_title = _get_pil_font(22, bold=True)
    f_sec = _get_pil_font(18, bold=True)
    f_lbl = _get_pil_font(15, bold=False)
    f_bold = _get_pil_font(16, bold=True)
    f_badge = _get_pil_font(15, bold=True)

    c_navy = "#1e2d3d"
    c_steel = "#3b5b7c"
    c_gold = "#b45309"
    c_green = "#15803d"
    c_red = "#be123c"
    c_bg_box = "#f8fafc"
    c_border = "#cbd5e1"
    c_text_dark = "#0f172a"
    c_text_muted = "#475569"

    # Barra de Título Superior
    draw.rectangle([(0, 0), (width, 55)], fill=c_navy)
    draw.text((25, 14), "DEMONSTRATIVO ESTATÍSTICO & CURVA GAUSSIANA DE DECISÃO (IN SEGES/ME Nº 65/2021)", font=f_title, fill="#ffffff")
    draw.text((width - 340, 18), "PADRÃO OFICIAL TCU / NLLC 14.133", font=f_badge, fill="#facc15")

    # Moldura dos Painéis
    draw.rounded_rectangle([(20, 70), (685, 595)], radius=12, fill=c_bg_box, outline=c_border, width=2)
    draw.rounded_rectangle([(705, 70), (1380, 595)], radius=12, fill=c_bg_box, outline=c_border, width=2)

    # =========================================================================
    # PAINEL 1: BOXPLOT
    # =========================================================================
    draw.text((40, 85), "1. DISPERSÃO DA AMOSTRA (MÉTODO DOS QUARTIS / BOXPLOT)", font=f_sec, fill=c_navy)
    draw.text((40, 110), "Saneamento estatístico das cotações extraídas das bases públicas oficiais", font=f_lbl, fill=c_text_muted)

    v_sorted = sorted([float(p) for p in valid_prices]) if valid_prices else [10.0]
    p_vals = [float(p) for p in purged_prices] if purged_prices else []
    all_vals = v_sorted + p_vals

    v_min, v_max = v_sorted[0], v_sorted[-1]
    def get_p(arr, pct):
        if len(arr) == 1: return arr[0]
        idx = (len(arr) - 1) * pct
        low = int(idx)
        high = math.ceil(idx)
        if low == high: return arr[low]
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

    bx_l, bx_r = 60, 645
    bx_w = bx_r - bx_l
    def scale_bx(val):
        if d_max <= d_min: return (bx_l + bx_r) / 2
        return bx_l + ((val - d_min) / (d_max - d_min)) * bx_w

    y_axis = 220
    draw.line([(bx_l, y_axis), (bx_r, y_axis)], fill="#94a3b8", width=2)

    x_min = scale_bx(lower_w)
    x_q1 = scale_bx(q1)
    x_med = scale_bx(median_val)
    x_mean = scale_bx(mean_val)
    x_q3 = scale_bx(q3)
    x_max = scale_bx(upper_w)

    draw.line([(x_min, y_axis), (x_q1, y_axis)], fill=c_steel, width=4)
    draw.line([(x_q3, y_axis), (x_max, y_axis)], fill=c_steel, width=4)
    draw.line([(x_min, y_axis - 20), (x_min, y_axis + 20)], fill=c_steel, width=4)
    draw.line([(x_max, y_axis - 20), (x_max, y_axis + 20)], fill=c_steel, width=4)

    draw.rounded_rectangle([(x_q1, y_axis - 40), (x_q3, y_axis + 40)], radius=6, fill="#e2e8f0", outline=c_steel, width=3)
    draw.line([(x_med, y_axis - 40), (x_med, y_axis + 40)], fill=c_gold, width=5)

    for dy in range(-48, 49, 8):
        draw.line([(x_mean, y_axis + dy), (x_mean, y_axis + dy + 4)], fill=c_green, width=3)

    for i, p in enumerate(v_sorted):
        cx = scale_bx(p)
        cy = y_axis + ((i % 3) - 1) * 10
        draw.ellipse([(cx - 7, cy - 7), (cx + 7, cy + 7)], fill="#10b981", outline="#047857", width=2)

    for i, p in enumerate(p_vals):
        cx = scale_bx(p)
        cy = y_axis + ((i % 2) * 22 - 11)
        draw.ellipse([(cx - 8, cy - 8), (cx + 8, cy + 8)], fill="#f43f5e", outline="#be123c", width=2)

    draw.line([(40, 310), (665, 310)], fill="#e2e8f0", width=2)
    draw.text((45, 325), f"Menor: R$ {v_min:.2f}", font=f_bold, fill=c_text_dark)
    draw.text((250, 325), f"Q1 (25%): R$ {q1:.2f}", font=f_bold, fill=c_text_dark)
    draw.text((460, 325), f"Q3 (75%): R$ {q3:.2f}", font=f_bold, fill=c_text_dark)

    draw.text((45, 360), f"Mediana: R$ {median_val:.2f}", font=f_bold, fill=c_gold)
    draw.text((250, 360), f"Média: R$ {mean_val:.2f}", font=f_bold, fill=c_green)
    draw.text((460, 360), f"Maior: R$ {v_max:.2f}", font=f_bold, fill=c_text_dark)

    # Cartão síntese
    draw.rounded_rectangle([(40, 415), (665, 570)], radius=8, fill="#ffffff", outline="#cbd5e1", width=2)
    draw.text((55, 430), "RESUMO DA AUDITORIA DA CESTA:", font=f_sec, fill=c_navy)
    draw.text((55, 465), f"• Total Coletado: {len(all_vals)} cotações ({len(v_sorted)} válidas e {len(p_vals)} outliers expurgados)", font=f_bold, fill=c_text_dark)
    draw.text((55, 498), f"• Coeficiente de Variação (CV): {cv_percent:.2f}% (Desvio Padrão: R$ {std_val:.2f})", font=f_bold, fill=c_text_dark)
    is_homo = cv_percent <= 25.0
    status_str = "SÉRIE HOMOGÊNEA (CV ≤ 25%)" if is_homo else "SÉRIE HETEROGÊNEA (CV > 25%)"
    status_col = c_green if is_homo else "#b45309"
    draw.text((55, 530), f"• Classificação IN 65/2021: {status_str}", font=f_bold, fill=status_col)

    # =========================================================================
    # PAINEL 2: CURVA NORMAL GAUSSIANA E CRITÉRIOS DE ESCOLHA
    # =========================================================================
    draw.text((725, 85), "2. DISTRIBUIÇÃO NORMAL GAUSSIANA & CRITÉRIOS DE ESCOLHA", font=f_sec, fill=c_navy)
    draw.text((725, 110), "Critérios legais e matemáticos para adoção da Média vs. Mediana (Art. 6º)", font=f_lbl, fill=c_text_muted)

    cx_l, cx_r = 745, 1340
    cx_w = cx_r - cx_l
    cy_base = 310
    cy_peak = 150

    curve_pts = []
    steps = 90
    for s in range(steps + 1):
        z = -3.0 + (6.0 * s / steps)
        y_val = math.exp(-0.5 * (z ** 2))
        px = cx_l + (s / steps) * cx_w
        py = cy_base - (y_val * (cy_base - cy_peak))
        curve_pts.append((px, py))

    draw.line([(cx_l, cy_base), (cx_r, cy_base)], fill="#94a3b8", width=2)

    x_m2 = cx_l + (1.0 / 6.0) * cx_w
    x_m1 = cx_l + (2.0 / 6.0) * cx_w
    x_mid = cx_l + (3.0 / 6.0) * cx_w
    x_p1 = cx_l + (4.0 / 6.0) * cx_w
    x_p2 = cx_l + (5.0 / 6.0) * cx_w

    poly_central = [(x_m1, cy_base)] + [pt for pt in curve_pts if x_m1 <= pt[0] <= x_p1] + [(x_p1, cy_base)]
    if len(poly_central) > 2:
        draw.polygon(poly_central, fill="#dcfce7")

    draw.line([(x_m1, cy_base), (x_m1, cy_peak + 30)], fill="#86efac", width=2)
    draw.line([(x_p1, cy_base), (x_p1, cy_peak + 30)], fill="#86efac", width=2)
    draw.line([(x_m2, cy_base), (x_m2, cy_peak + 90)], fill="#fca5a5", width=2)
    draw.line([(x_p2, cy_base), (x_p2, cy_peak + 90)], fill="#fca5a5", width=2)

    for i in range(len(curve_pts) - 1):
        draw.line([curve_pts[i], curve_pts[i+1]], fill=c_steel, width=4)

    draw.line([(x_mid, cy_base), (x_mid, cy_peak)], fill=c_gold, width=4)

    draw.text((x_mid - 45, cy_base - 65), "ZONA CENTRAL", font=f_bold, fill=c_green)
    draw.text((x_mid - 40, cy_base - 45), "CV ≤ 25%", font=f_bold, fill=c_green)
    draw.text((x_mid - 30, cy_base + 8), "μ (Média)", font=f_bold, fill=c_navy)

    draw.text((x_p1 + 10, cy_base - 55), "DISPERSÃO", font=f_bold, fill="#b45309")
    draw.text((x_p1 + 10, cy_base - 35), "CV > 25%", font=f_bold, fill="#b45309")
    draw.text((x_p1 - 15, cy_base + 8), "+1σ", font=f_bold, fill=c_text_muted)
    draw.text((x_m1 - 15, cy_base + 8), "-1σ", font=f_bold, fill=c_text_muted)

    draw.text((cx_l + 15, cy_base - 35), "CAUDA (Expurgo)", font=f_bold, fill=c_red)
    draw.text((cx_r - 165, cy_base - 35), "CAUDA (Expurgo)", font=f_bold, fill=c_red)
    draw.text((x_m2 - 15, cy_base + 8), "-2σ", font=f_bold, fill=c_text_muted)
    draw.text((x_p2 - 15, cy_base + 8), "+2σ", font=f_bold, fill=c_text_muted)

    # Guia Didático de Critérios de Escolha
    draw.rounded_rectangle([(725, 360), (1360, 570)], radius=8, fill="#ffffff", outline="#cbd5e1", width=2)

    draw.rectangle([(740, 375), (746, 420)], fill=c_green)
    draw.text((758, 375), "REGRA DA MÉDIA SANEADA (CV ≤ 25%):", font=f_bold, fill=c_green)
    draw.text((758, 398), "Quando a amostra é homogênea, a média aritmética reflete com precisão o valor de mercado.", font=f_lbl, fill=c_text_dark)

    draw.rectangle([(740, 435), (746, 480)], fill=c_gold)
    draw.text((758, 435), "REGRA DA MEDIANA SANEADA (CV > 25%):", font=f_bold, fill="#92400e")
    draw.text((758, 458), "Obrigatória (Art. 6º, § 5º da IN 65/2021) em séries heterogêneas para neutralizar distorções extremas.", font=f_lbl, fill=c_text_dark)

    adopted_is_med = "MEDIANA" in adopted_param.upper()
    bg_dec = "#fef3c7" if adopted_is_med else "#dcfce7"
    border_dec = c_gold if adopted_is_med else c_green
    draw.rounded_rectangle([(740, 505), (1345, 550)], radius=6, fill=bg_dec, outline=border_dec, width=2)
    dec_text = f"DECISÃO LEGAL ADOTADA: {adopted_param} (R$ {median_val if adopted_is_med else mean_val:.2f}) — CV Saneado de {cv_percent:.2f}%"
    draw.text((755, 518), dec_text, font=f_bold, fill=c_navy)

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
    """Gera o arquivo Word (.docx) formal com Brasão Oficial e padrão visual Marinha do Brasil em alta legibilidade."""
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
        ("REPÚBLICA FEDERATIVA DO BRASIL", 11.0, True),
        ("MINISTÉRIO DA DEFESA", 10.5, True),
        ("COMANDO DA MARINHA", 10.5, True),
        (om_name.upper(), 10.0, False),
        ("SETOR DE OBTENÇÃO E CONTRATAÇÕES", 9.5, False),
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
    r_tit.font.size = Pt(14)
    r_tit.font.bold = True
    r_tit.font.color.rgb = RGBColor(15, 23, 42)

    r_sub = p_title.add_run(
        "Conforme o Art. 23 da Lei nº 14.133/2021 e a Instrução Normativa SEGES/ME nº 65/2021"
    )
    r_sub.font.name = "Calibri"
    r_sub.font.size = Pt(10)
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
    p_s1.paragraph_format.space_before = Pt(10)
    p_s1.paragraph_format.space_after = Pt(4)
    r_s1 = p_s1.add_run("1. IDENTIFICAÇÃO DO OBJETO E PREÇO ESTIMADO")
    r_s1.font.bold = True
    r_s1.font.size = Pt(11.5)

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

    col_widths = (Inches(2.5), Inches(4.5))
    for r_idx, (k, v) in enumerate(t_data):
        row = t_item.rows[r_idx]
        c0, c1 = row.cells[0], row.cells[1]
        c0.width, c1.width = col_widths[0], col_widths[1]
        set_cell_margins(c0, 90, 90, 110, 110)
        set_cell_margins(c1, 90, 90, 110, 110)
        set_cell_background(c0, "F1F5F9")
        set_cell_background(c1, "FFFFFF" if r_idx < 4 else "FEF3C7")

        p0 = c0.paragraphs[0]
        p0.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        r0 = p0.add_run(k)
        r0.font.bold = True
        r0.font.size = Pt(9.5)
        r0.font.name = "Calibri"

        p1 = c1.paragraphs[0]
        p1.alignment = WD_ALIGN_PARAGRAPH.LEFT
        r1 = p1.add_run(v)
        r1.font.bold = (r_idx >= 3)
        r1.font.size = Pt(10 if r_idx < 4 else 11)
        r1.font.name = "Calibri"
        if r_idx == 4:
            r1.font.color.rgb = RGBColor(146, 64, 14)

    # SEÇÃO 2: FONTES CONSULTADAS E COTAÇÕES
    p_s2 = doc.add_paragraph()
    p_s2.paragraph_format.space_before = Pt(14)
    p_s2.paragraph_format.space_after = Pt(4)
    r_s2 = p_s2.add_run("2. COTAÇÕES COLETADAS NAS COMPRAS PÚBLICAS SIMILARES (Art. 5º da IN 65/2021)")
    r_s2.font.bold = True
    r_s2.font.size = Pt(11.5)

    valid_quotes = report.get("valid_quotes", [])
    quotes_to_show = valid_quotes[:10]

    t_quotes = doc.add_table(rows=1 + len(quotes_to_show), cols=6)
    t_quotes.alignment = WD_TABLE_ALIGNMENT.CENTER
    q_widths = (Inches(0.40), Inches(1.35), Inches(1.10), Inches(1.55), Inches(1.00), Inches(1.27))

    q_headers = ["Nº", "UASG / Órgão", "Compra nº / Data", "Fornecedor Vencedor", "Preço Unit.", "Comprovação PNCP"]
    hdr_row = t_quotes.rows[0]
    for c_idx, h in enumerate(q_headers):
        cell = hdr_row.cells[c_idx]
        cell.width = q_widths[c_idx]
        set_cell_background(cell, "1E293B")
        set_cell_margins(cell, 90, 90, 90, 90)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(h)
        r.font.bold = True
        r.font.size = Pt(9.5)
        r.font.color.rgb = RGBColor(255, 255, 255)

    for idx, q in enumerate(quotes_to_show):
        row = t_quotes.rows[1 + idx]
        q_url = q.get("pncp_or_compras_uri") or ""
        if not q_url or "?q=" in q_url or "/editais" not in q_url:
            key = q.get("purchase_key") or (str(q.get("id_compra_item", ""))[:17] if len(str(q.get("id_compra_item", ""))) >= 17 else "")
            if not key or len(key) < 16:
                id_item = str(q.get("id_compra_item", ""))
                id_compra = str(q.get("id_compra", ""))
                try:
                    from gov_api_client import GovProcurementAPIClient
                    key, _ = GovProcurementAPIClient.parse_comprasnet_key_and_number(id_item, id_compra)
                except Exception:
                    pass
            if key:
                if len(key) == 16:
                    key = "0" + key
                try:
                    from pncp_extractor import resolve_pncp_webpage_url
                    lit = resolve_pncp_webpage_url(key)
                    if lit:
                        q_url = lit
                except Exception:
                    pass
        if not q_url:
            q_url = f"https://pncp.gov.br/app/editais?q={q.get('buyer_uasg', '')}" if q.get("buyer_uasg") else "https://pncp.gov.br/app/editais"

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
            set_cell_margins(cell, 70, 70, 70, 70)
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.RIGHT if c_idx == 4 else (WD_ALIGN_PARAGRAPH.CENTER if c_idx in (0, 2) else WD_ALIGN_PARAGRAPH.LEFT)
            r = p.add_run(val)
            r.font.size = Pt(9)
            r.font.name = "Calibri"
            if c_idx == 4:
                r.font.bold = True

        # Coluna 6: Comprovação Digital / PNCP (QR Code + Link Clicável)
        cell_qr = row.cells[5]
        cell_qr.width = q_widths[5]
        set_cell_background(cell_qr, bg)
        set_cell_margins(cell_qr, 40, 40, 40, 40)
        p_qr = cell_qr.paragraphs[0]
        p_qr.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_qr.paragraph_format.space_before = Pt(2)
        p_qr.paragraph_format.space_after = Pt(2)

        try:
            qr_obj = qrcode.QRCode(
                version=1,
                error_correction=qrcode.constants.ERROR_CORRECT_L,
                box_size=2,
                border=1,
            )
            qr_obj.add_data(q_url)
            qr_obj.make(fit=True)
            qr_img = qr_obj.make_image(fill_color="black", back_color="white")
            qr_buf = io.BytesIO()
            qr_img.save(qr_buf, format="PNG")
            qr_buf.seek(0)
            run_qr = p_qr.add_run()
            run_qr.add_picture(qr_buf, width=Inches(0.72))
        except Exception as exc:
            logger.warning(f"Erro ao gerar QR code da cotação no Word: {exc}")

        p_link = cell_qr.add_paragraph()
        p_link.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_link.paragraph_format.space_before = Pt(0)
        p_link.paragraph_format.space_after = Pt(2)
        add_docx_hyperlink(p_link, q_url, "🔗 Ver no PNCP", color="1D4ED8", underline=True)

    # SEÇÃO 3: SANEAMENTO ESTATÍSTICO
    p_s3 = doc.add_paragraph()
    p_s3.paragraph_format.space_before = Pt(14)
    p_s3.paragraph_format.space_after = Pt(4)
    r_s3 = p_s3.add_run("3. SANEAMENTO ESTATÍSTICO E MEMÓRIA DE CÁLCULO (Art. 6º da IN 65/2021)")
    r_s3.font.bold = True
    r_s3.font.size = Pt(11.5)

    t_stat = doc.add_table(rows=4, cols=4)
    t_stat.alignment = WD_TABLE_ALIGNMENT.CENTER
    s_widths = (Inches(1.75), Inches(1.75), Inches(1.75), Inches(1.75))

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
            set_cell_margins(cell, 70, 70, 80, 80)
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.RIGHT if is_label else WD_ALIGN_PARAGRAPH.LEFT
            r = p.add_run(val)
            r.font.size = Pt(9.5)
            r.font.name = "Calibri"
            r.font.bold = is_label or (c_idx == 3 and r_idx == 3)

    # SEÇÃO 4: DEMONSTRATIVO GRÁFICO (BOXPLOT E CURVA NORMAL GAUSSIANA)
    p_sg = doc.add_paragraph()
    p_sg.paragraph_format.space_before = Pt(14)
    p_sg.paragraph_format.space_after = Pt(4)
    r_sg = p_sg.add_run("4. DEMONSTRATIVO GRÁFICO: DISPERSÃO (BOXPLOT) E CURVA GAUSSIANA (IN 65/2021)")
    r_sg.font.bold = True
    r_sg.font.size = Pt(11.5)

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
        run_chart.add_picture(io.BytesIO(infographic_bytes), width=Inches(6.6))
    except Exception as exc:
        print(f"Erro ao inserir infográfico estatístico no Word: {exc}")

    # SEÇÃO 5: JUSTIFICATIVA METODOLÓGICA
    p_s4 = doc.add_paragraph()
    p_s4.paragraph_format.space_before = Pt(14)
    p_s4.paragraph_format.space_after = Pt(4)
    r_s4 = p_s4.add_run("5. JUSTIFICATIVA PARA A METODOLOGIA ADOTADA (Art. 6º, §4º da IN 65/2021)")
    r_s4.font.bold = True
    r_s4.font.size = Pt(11.5)

    p_just = doc.add_paragraph()
    p_just.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p_just.paragraph_format.line_spacing = 1.15
    p_just.paragraph_format.space_after = Pt(12)
    r_just = p_just.add_run(report.get("legal_methodology_justification", ""))
    r_just.font.size = Pt(10)
    r_just.font.name = "Calibri"
    r_just.font.italic = True

    # SEÇÃO 6: PAINEL DE CONFERÊNCIA DIGITAL & QR CODES (PORTAL PNCP E COMPRAS.GOV.BR)
    p_sqr = doc.add_paragraph()
    p_sqr.paragraph_format.space_before = Pt(14)
    p_sqr.paragraph_format.space_after = Pt(4)
    r_sqr = p_sqr.add_run("6. PAINEL DE CONFERÊNCIA DIGITAL & QR CODES (PORTAL PNCP E COMPRAS.GOV.BR)")
    r_sqr.font.bold = True
    r_sqr.font.size = Pt(11.5)

    p_sqr_desc = doc.add_paragraph()
    p_sqr_desc.paragraph_format.space_after = Pt(8)
    r_sqr_desc = p_sqr_desc.add_run("Aponte a câmera do seu smartphone para o QR Code correspondente ou clique no link destacado diretamente abaixo para verificação oficial em tempo real:")
    r_sqr_desc.font.size = Pt(9.5)
    r_sqr_desc.font.italic = True
    r_sqr_desc.font.color.rgb = RGBColor(71, 85, 105)

    pncp_url = f"https://pncp.gov.br/app/editais?q={item_code}"
    painel_url = "https://paineldeprecos.planejamento.gov.br"
    first_q_url = ""
    for q in valid_quotes:
        if q.get("pncp_or_compras_uri"):
            first_q_url = q["pncp_or_compras_uri"]
            break
        elif q.get("purchase_key"):
            first_q_url = f"https://pncp.gov.br/app/editais/{q['purchase_key']}"
            break
    if not first_q_url:
        first_q_url = pncp_url

    t_qr = doc.add_table(rows=2, cols=3)
    t_qr.alignment = WD_TABLE_ALIGNMENT.CENTER
    qr_col_widths = (Inches(2.3), Inches(2.3), Inches(2.3))

    qr1_png = generate_qr_code_bytes(pncp_url, box_size=3)
    qr2_png = generate_qr_code_bytes(painel_url, box_size=3)
    qr3_png = generate_qr_code_bytes(first_q_url, box_size=3)

    qr_cards = [
        (qr1_png, "Portal PNCP Oficial", pncp_url, "Consulta por Objeto / CATMAT"),
        (qr2_png, "Painel de Preços Federal", painel_url, "Repositório SEGES / MGI"),
        (qr3_png, "Compra Pública / Ata", first_q_url, "Dossiê Oficial da Licitação"),
    ]

    for c_i, (qr_bytes, title, url, subtitle) in enumerate(qr_cards):
        cell_img = t_qr.rows[0].cells[c_i]
        cell_img.width = qr_col_widths[c_i]
        set_cell_background(cell_img, "F8FAFC")
        set_cell_margins(cell_img, 60, 40, 60, 60)
        p_img = cell_img.paragraphs[0]
        p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_img.add_run().add_picture(io.BytesIO(qr_bytes), width=Inches(1.15))

        cell_txt = t_qr.rows[1].cells[c_i]
        cell_txt.width = qr_col_widths[c_i]
        set_cell_background(cell_txt, "F8FAFC")
        set_cell_margins(cell_txt, 20, 60, 60, 60)
        p_txt = cell_txt.paragraphs[0]
        p_txt.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_txt.paragraph_format.line_spacing = 1.05

        r_title = p_txt.add_run(f"{title}\n")
        r_title.font.bold = True
        r_title.font.size = Pt(9.5)
        r_title.font.name = "Calibri"

        # LINK CLICÁVEL NATIVO DIRETAMENTE ABAIXO DO QR CODE
        add_docx_hyperlink(p_txt, url, "Acessar Portal Oficial", color="1D4ED8", underline=True)

        p_sub = cell_txt.add_paragraph()
        p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_sub.paragraph_format.space_before = Pt(2)
        r_sub = p_sub.add_run(subtitle)
        r_sub.font.size = Pt(8.5)
        r_sub.font.italic = True
        r_sub.font.color.rgb = RGBColor(100, 116, 139)

    # SEÇÃO 7: ASSINATURAS E FECHAMENTO
    today = datetime.now()
    mes_nome = _get_mes_extenso(today.month)
    data_str = f"{today.day} de {mes_nome} de {today.year}."

    p_dt = doc.add_paragraph()
    p_dt.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p_dt.paragraph_format.space_before = Pt(16)
    p_dt.paragraph_format.space_after = Pt(28)
    r_dt = p_dt.add_run(data_str)
    r_dt.font.size = Pt(10)
    r_dt.font.name = "Calibri"

    t_sig = doc.add_table(rows=1, cols=2)
    t_sig.alignment = WD_TABLE_ALIGNMENT.CENTER
    c_sig0, c_sig1 = t_sig.rows[0].cells[0], t_sig.rows[0].cells[1]
    c_sig0.width, c_sig1.width = Inches(3.4), Inches(3.4)

    p_sig0 = c_sig0.paragraphs[0]
    p_sig0.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_sig0.add_run("_________________________________________\n").font.size = Pt(9.5)
    r_s0_1 = p_sig0.add_run(f"{elaborador_nome}\n")
    r_s0_1.font.bold = True
    r_s0_1.font.size = Pt(9.5)
    r_s0_2 = p_sig0.add_run("Elaborador da Pesquisa de Preços")
    r_s0_2.font.size = Pt(8.5)
    r_s0_2.font.color.rgb = RGBColor(100, 116, 139)

    p_sig1 = c_sig1.paragraphs[0]
    p_sig1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_sig1.add_run("_________________________________________\n").font.size = Pt(9.5)
    r_s1_1 = p_sig1.add_run(f"{ordenador_nome}\n")
    r_s1_1.font.bold = True
    r_s1_1.font.size = Pt(9.5)
    r_s1_2 = p_sig1.add_run("Ratificador da Pesquisa de Preços")
    r_s1_2.font.size = Pt(8.5)
    r_s1_2.font.color.rgb = RGBColor(100, 116, 139)

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
    """Gera o arquivo PDF A4 de alta qualidade visual com Brasão Oficial e diagramação formal em alta legibilidade."""
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

    # Estilos customizados de alta legibilidade (proporções institucionais)
    st_header_main = ParagraphStyle(
        "NavalHeaderMain",
        parent=styles["Normal"],
        fontName="Helvetica-Bold",
        fontSize=11,
        leading=14,
        alignment=1,
        textColor=colors.HexColor("#0f172a"),
    )

    st_header_sub = ParagraphStyle(
        "NavalHeaderSub",
        parent=styles["Normal"],
        fontName="Helvetica",
        fontSize=9.5,
        leading=13,
        alignment=1,
        textColor=colors.HexColor("#334155"),
    )

    st_doc_title = ParagraphStyle(
        "DocTitle",
        parent=styles["Normal"],
        fontName="Helvetica-Bold",
        fontSize=14.5,
        leading=18,
        alignment=1,
        textColor=colors.HexColor("#0f172a"),
    )

    st_doc_sub = ParagraphStyle(
        "DocSub",
        parent=styles["Normal"],
        fontName="Helvetica-Oblique",
        fontSize=9.5,
        leading=13,
        alignment=1,
        textColor=colors.HexColor("#64748b"),
    )

    st_sec_title = ParagraphStyle(
        "SecTitle",
        parent=styles["Normal"],
        fontName="Helvetica-Bold",
        fontSize=11.5,
        leading=15,
        textColor=colors.HexColor("#0f172a"),
    )

    st_cell_hdr = ParagraphStyle(
        "CellHdr",
        parent=styles["Normal"],
        fontName="Helvetica-Bold",
        fontSize=9.5,
        leading=12,
        textColor=colors.white,
        alignment=1,
    )

    st_cell_txt = ParagraphStyle(
        "CellTxt",
        parent=styles["Normal"],
        fontName="Helvetica",
        fontSize=9,
        leading=12,
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
        fontSize=9.5,
        leading=12,
        textColor=colors.HexColor("#0f172a"),
    )

    st_cell_bold_right = ParagraphStyle(
        "CellBoldRight",
        parent=st_cell_bold,
        alignment=2,
    )

    st_just = ParagraphStyle(
        "JustifiedLegal",
        parent=styles["Normal"],
        fontName="Helvetica-Oblique",
        fontSize=10,
        leading=14,
        alignment=4,
        textColor=colors.HexColor("#334155"),
    )

    st_sig_title = ParagraphStyle(
        "SigTitle",
        parent=styles["Normal"],
        fontName="Helvetica-Bold",
        fontSize=9.5,
        leading=12,
        alignment=1,
        textColor=colors.HexColor("#0f172a"),
    )

    st_sig_sub = ParagraphStyle(
        "SigSub",
        parent=styles["Normal"],
        fontName="Helvetica",
        fontSize=8.5,
        leading=11,
        alignment=1,
        textColor=colors.HexColor("#64748b"),
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
        ("TOPPADDING", (0, 0), (-1, -1), 3.5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3.5),
        ("LEFTPADDING", (0, 0), (-1, -1), 5),
        ("RIGHTPADDING", (0, 0), (-1, -1), 5),
    ]))
    story.append(t_item)
    story.append(Spacer(1, 3 * mm))

    # SEÇÃO 2: COTAÇÕES (TOP 7)
    story.append(Paragraph("2. COTAÇÕES VÁLIDAS AUDITADAS (Art. 5º da IN 65/2021)", st_sec_title))
    story.append(Spacer(1, 1.5 * mm))

    quotes_data = [
        [
            Paragraph("Nº", st_cell_hdr),
            Paragraph("Órgão / UASG / UF", st_cell_hdr),
            Paragraph("Nº Compra / Data", st_cell_hdr),
            Paragraph("Fornecedor Vencedor & CNPJ", st_cell_hdr),
            Paragraph("Preço Unit.", st_cell_hdr),
            Paragraph("Comprovação PNCP", st_cell_hdr),
        ]
    ]

    for idx, q in enumerate(valid_quotes[:7]):
        q_url = q.get("pncp_or_compras_uri") or ""
        if not q_url or "?q=" in q_url or "/editais" not in q_url:
            key = q.get("purchase_key") or (str(q.get("id_compra_item", ""))[:17] if len(str(q.get("id_compra_item", ""))) >= 17 else "")
            if not key or len(key) < 16:
                id_item = str(q.get("id_compra_item", ""))
                id_compra = str(q.get("id_compra", ""))
                try:
                    from gov_api_client import GovProcurementAPIClient
                    key, _ = GovProcurementAPIClient.parse_comprasnet_key_and_number(id_item, id_compra)
                except Exception:
                    pass
            if key:
                if len(key) == 16:
                    key = "0" + key
                try:
                    from pncp_extractor import resolve_pncp_webpage_url
                    lit = resolve_pncp_webpage_url(key)
                    if lit:
                        q_url = lit
                except Exception:
                    pass
        if not q_url:
            q_url = f"https://pncp.gov.br/app/editais?q={q.get('buyer_uasg', '')}" if q.get("buyer_uasg") else "https://pncp.gov.br/app/editais"
        
        qr_obj = qrcode.QRCode(
            version=1,
            error_correction=qrcode.constants.ERROR_CORRECT_L,
            box_size=2,
            border=1,
        )
        qr_obj.add_data(q_url)
        qr_obj.make(fit=True)
        qr_img = qr_obj.make_image(fill_color="black", back_color="white")
        qr_buf = io.BytesIO()
        qr_img.save(qr_buf, format="PNG")
        qr_buf.seek(0)
        rl_qr = RLImage(qr_buf, width=1.5 * cm, height=1.5 * cm)
        p_link = Paragraph(f"<font color='#0284c7' size='7'><u><a href='{q_url}'><b>🔗 Ver no PNCP</b></a></u></font>", st_cell_txt_center)

        quotes_data.append([
            Paragraph(f"{idx+1:02d}", st_cell_txt_center),
            Paragraph(f"UASG {q.get('buyer_uasg', '')} ({q.get('buyer_uf', '')})<br/>{q.get('buyer_org_name', '')[:25]}", st_cell_txt),
            Paragraph(f"{q.get('purchase_number_year', '')}<br/>{q.get('purchase_date', '')}", st_cell_txt_center),
            Paragraph(f"{q.get('supplier_name', '')[:35]}<br/>CNPJ: {q.get('supplier_cnpj', '')}", st_cell_txt),
            Paragraph(_format_brl(float(q.get("unit_price", 0.0))), st_cell_bold_right),
            [rl_qr, p_link],
        ])

    t_quotes = Table(quotes_data, colWidths=[0.8 * cm, 3.8 * cm, 2.4 * cm, 4.8 * cm, 2.5 * cm, 3.7 * cm])
    t_quotes.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1e293b")),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("ALIGN", (5, 0), (5, -1), "CENTER"),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#cbd5e1")),
        ("TOPPADDING", (0, 0), (-1, -1), 3),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
        ("LEFTPADDING", (0, 0), (-1, -1), 3),
        ("RIGHTPADDING", (0, 0), (-1, -1), 3),
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
        ("TOPPADDING", (0, 0), (-1, -1), 3),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
        ("LEFTPADDING", (0, 0), (-1, -1), 5),
        ("RIGHTPADDING", (0, 0), (-1, -1), 5),
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
        story.append(RLImage(io.BytesIO(infographic_bytes), width=18.0 * cm, height=7.97 * cm))
        story.append(Spacer(1, 3 * mm))
    except Exception as exc:
        print(f"Erro ao inserir gráfico estatístico no PDF: {exc}")

    # SEÇÃO 5: JUSTIFICATIVA JURÍDICA
    story.append(Paragraph("5. JUSTIFICATIVA METODOLÓGICA (Art. 6º da IN 65/2021)", st_sec_title))
    story.append(Spacer(1, 1 * mm))
    story.append(Paragraph(report.get("legal_methodology_justification", ""), st_just))
    story.append(Spacer(1, 3 * mm))

    # SEÇÃO 6: PAINEL DE CONFERÊNCIA DIGITAL & QR CODES (PORTAL PNCP E COMPRAS.GOV.BR)
    story.append(Paragraph("6. PAINEL DE CONFERÊNCIA DIGITAL & QR CODES (PORTAL PNCP E COMPRAS.GOV.BR)", st_sec_title))
    story.append(Paragraph("<i>Aponte a câmera do seu smartphone para o QR Code correspondente ou clique no link sublinhado diretamente abaixo para conferência oficial em tempo real:</i>", st_doc_sub))
    story.append(Spacer(1, 2 * mm))

    pncp_url = f"https://pncp.gov.br/app/editais?q={item_code}"
    painel_url = "https://paineldeprecos.planejamento.gov.br"
    first_q_url = ""
    for q in valid_quotes:
        if q.get("pncp_or_compras_uri"):
            first_q_url = q["pncp_or_compras_uri"]
            break
        elif q.get("purchase_key"):
            first_q_url = f"https://pncp.gov.br/app/editais/{q['purchase_key']}"
            break
    if not first_q_url:
        first_q_url = pncp_url

    qr1_png = generate_qr_code_bytes(pncp_url, box_size=3)
    qr2_png = generate_qr_code_bytes(painel_url, box_size=3)
    qr3_png = generate_qr_code_bytes(first_q_url, box_size=3)

    qr_pdf_data = [
        [
            RLImage(io.BytesIO(qr1_png), width=2.4 * cm, height=2.4 * cm),
            RLImage(io.BytesIO(qr2_png), width=2.4 * cm, height=2.4 * cm),
            RLImage(io.BytesIO(qr3_png), width=2.4 * cm, height=2.4 * cm),
        ],
        [
            Paragraph(f"<b>Portal PNCP Oficial</b><br/><font color='#0284c7'><u><a href='{pncp_url}'>Acessar Portal PNCP</a></u></font><br/><font size='7.5' color='#64748b'>Consulta por Objeto / CATMAT</font>", st_cell_txt_center),
            Paragraph(f"<b>Painel de Preços Federal</b><br/><font color='#0284c7'><u><a href='{painel_url}'>Acessar Painel de Preços</a></u></font><br/><font size='7.5' color='#64748b'>paineldeprecos.planejamento.gov.br</font>", st_cell_txt_center),
            Paragraph(f"<b>Compra Pública / Ata</b><br/><font color='#0284c7'><u><a href='{first_q_url}'>Acessar Compra Homologada</a></u></font><br/><font size='7.5' color='#64748b'>Dossiê Oficial da Licitação</font>", st_cell_txt_center),
        ]
    ]
    t_qr = Table(qr_pdf_data, colWidths=[6.0 * cm, 6.0 * cm, 6.0 * cm])
    t_qr.setStyle(TableStyle([
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#f8fafc")),
        ("BOX", (0, 0), (-1, -1), 0.5, colors.HexColor("#cbd5e1")),
        ("INNERGRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#e2e8f0")),
        ("TOPPADDING", (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
        ("LEFTPADDING", (0, 0), (-1, -1), 4),
        ("RIGHTPADDING", (0, 0), (-1, -1), 4),
    ]))
    story.append(t_qr)
    story.append(Spacer(1, 3 * mm))

    # SEÇÃO 7: ASSINATURAS
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
    ]))
    story.append(KeepTogether(t_sig))

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

            # Raio-X de Certidões das Empresas Cotadas (Word e PDF)
            try:
                cert_docx = build_supplier_compliance_dossier_docx_bytes(rep, om_name=om_name, elaborador_nome=elaborador_nome, ordenador_nome=ordenador_nome)
                zf.writestr(f"{folder}/RaioX_Certidoes_Empresas_{prefix}_{code}.docx", cert_docx)
                cert_pdf = build_supplier_compliance_dossier_pdf_bytes(rep, om_name=om_name, elaborador_nome=elaborador_nome, ordenador_nome=ordenador_nome)
                zf.writestr(f"{folder}/RaioX_Certidoes_Empresas_{prefix}_{code}.pdf", cert_pdf)
            except Exception as _e_cert:
                pass

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


# =============================================================================
# DOSSIÊ CONSOLIDADO DE REGULARIDADE E CERTIDÕES DOS FORNECEDORES (RAIO-X ART. 68)
# =============================================================================

def _extract_unique_suppliers_from_report(report: Dict[str, Any]) -> List[Dict[str, Any]]:
    """Extrai lista de fornecedores distintos auditados a partir do relatório de pesquisa."""
    raw_quotes = (
        (report.get("valid_quotes") or []) +
        (report.get("purged_quotes") or []) +
        (report.get("quotes") or [])
    )
    seen = {}
    for q in raw_quotes:
        cnpj_raw = str(q.get("supplier_cnpj") or "")
        clean = re.sub(r"\D", "", cnpj_raw)
        if clean and len(clean) == 14 and clean not in seen:
            seen[clean] = {
                "cnpj_clean": clean,
                "cnpj": q.get("supplier_cnpj") or clean,
                "name": q.get("supplier_name") or "Fornecedor",
                "price": float(q.get("unit_price") or 0.0),
                "buyer_uasg": q.get("buyer_uasg") or "—",
                "buyer_uf": q.get("buyer_uf") or "BR",
                "purchase_num": q.get("purchase_number_year") or "—",
                "is_selected": q.get("is_selected", True),
                "is_outlier": q.get("is_outlier", False),
            }
    return list(seen.values())


def build_supplier_compliance_dossier_docx_bytes(
    report: Dict[str, Any],
    om_name: str = "MARINHA DO BRASIL",
    elaborador_nome: str = "ENCARREGADO DO SETOR DE OBTENÇÃO",
    ordenador_nome: str = "ORDENADOR DE DESPESAS",
) -> bytes:
    """
    Gera em Microsoft Word (.docx) o Dossiê / Relatório Raio-X de Regularidade e Certidões
    dos Fornecedores participantes da pesquisa de preços (Art. 68 e 156 da Lei nº 14.133/2021).
    """
    try:
        from sanctions_checker import check_supplier_full_compliance
    except ImportError:
        from core.sanctions_checker import check_supplier_full_compliance

    doc = docx.Document()
    for section in doc.sections:
        section.top_margin = Inches(0.8)
        section.bottom_margin = Inches(0.8)
        section.left_margin = Inches(0.8)
        section.right_margin = Inches(0.8)

    # 1. Cabeçalho Oficial com Brasão
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
        ("SETOR DE OBTENÇÃO E PREÇOS / ASSESSORIA DE COMPLIANCE", 9, False),
    ]
    for text, sz, bold in lines:
        r = p_head.add_run(f"{text}\n")
        r.font.name = "Calibri"
        r.font.size = Pt(sz)
        r.font.bold = bold
        r.font.color.rgb = RGBColor(15, 23, 42)

    # Título do Relatório
    p_tit = doc.add_paragraph()
    p_tit.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_tit.paragraph_format.space_before = Pt(4)
    p_tit.paragraph_format.space_after = Pt(2)
    r_tit = p_tit.add_run("DOSSIÊ DE REGULARIDADE FISCAL, TRABALHISTA E RISCO CAUTELAR\n")
    r_tit.font.name = "Calibri"
    r_tit.font.size = Pt(13)
    r_tit.font.bold = True
    r_tit.font.color.rgb = RGBColor(15, 23, 42)

    p_sub = doc.add_paragraph()
    p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_sub.paragraph_format.space_after = Pt(14)
    r_sub = p_sub.add_run("AUDITORIA CONJUNTA DE FORNECEDORES (RAIO-X DE CERTIDÕES E SANÇÕES)\nFundamentação: Art. 68 e 156 da Lei nº 14.133/2021 • IN SEGES/ME nº 65/2021")
    r_sub.font.name = "Calibri"
    r_sub.font.size = Pt(9.5)
    r_sub.font.italic = True
    r_sub.font.color.rgb = RGBColor(71, 85, 105)

    # 2. Identificação do Objeto
    code = report.get("catmat_catser_code") or "—"
    item_type = str(report.get("item_type") or "MATERIAL").upper()
    prefix = "CATSER" if item_type == "SERVICO" else "CATMAT"
    desc = report.get("item_description") or "Item de Aquisição / Registro de Preços"
    unit_price = float(report.get("final_estimated_unit_price") or 0.0)

    p_s1 = doc.add_paragraph()
    p_s1.paragraph_format.space_before = Pt(8)
    p_s1.paragraph_format.space_after = Pt(3)
    r_s1 = p_s1.add_run("1. IDENTIFICAÇÃO DO PROCESSO E DO OBJETO AUDITADO")
    r_s1.font.bold = True
    r_s1.font.size = Pt(11)

    t_obj = doc.add_table(rows=3, cols=2)
    t_obj.alignment = WD_TABLE_ALIGNMENT.CENTER
    obj_widths = (Inches(2.2), Inches(4.8))
    obj_data = [
        ("Código do Catálogo:", f"{prefix} {code}"),
        ("Descrição Oficial do Objeto:", desc),
        ("Preço Unitário Homologável:", f"{_format_brl(unit_price)} (Ref. IN 65/2021)"),
    ]
    for r_i, (lbl, val) in enumerate(obj_data):
        row = t_obj.rows[r_i]
        c0, c1 = row.cells[0], row.cells[1]
        c0.width, c1.width = obj_widths[0], obj_widths[1]
        set_cell_background(c0, "F1F5F9")
        set_cell_margins(c0, 40, 40, 60, 60)
        set_cell_margins(c1, 40, 40, 60, 60)
        p0, p1 = c0.paragraphs[0], c1.paragraphs[0]
        r0 = p0.add_run(lbl)
        r0.font.bold = True
        r0.font.size = Pt(9)
        r1 = p1.add_run(val)
        r1.font.size = Pt(9)

    # 3. Nota Técnica Jurídica: Art. 23 (Pesquisa de Preços) vs Art. 68 (Habilitação Contratual)
    p_s2 = doc.add_paragraph()
    p_s2.paragraph_format.space_before = Pt(14)
    p_s2.paragraph_format.space_after = Pt(4)
    r_s2 = p_s2.add_run("2. NOTA TÉCNICA E JURISPRUDÊNCIA DE GOVERNANÇA (LEI Nº 14.133/2021)")
    r_s2.font.bold = True
    r_s2.font.size = Pt(11)

    p_nota = doc.add_paragraph()
    p_nota.paragraph_format.line_spacing = 1.15
    p_nota.paragraph_format.space_after = Pt(10)
    r_nota = p_nota.add_run(
        "A presente auditoria preventiva analisa a conformidade fiscal e trabalhista das empresas que compõem o banco de dados "
        "de preços públicos. Cumpre destacar as diretrizes aplicáveis:\n\n"
        "a) Fase Preparatória / Pesquisa de Preços (Art. 23 da Lei nº 14.133/2021 e IN SEGES/ME nº 65/2021): As cotações "
        "originadas de compras públicas homologadas pretéritas retratam a realidade fática do mercado econômico. O fato de uma "
        "empresa outrora vencedora ter sua certidão vencida posteriormente à execução contratual não contamina a veracidade do preço "
        "praticado perante o Poder Público, mantendo-se a validade da cotação na cesta de referências.\n\n"
        "b) Fase de Habilitação / Contratação (Art. 68 e 156): A regularidade fiscal, social e trabalhista perante a Fazenda Nacional, "
        "Seguridade Social (FGTS), Justiça do Trabalho (CNDT) e ausência de sanções nos cadastros CEIS/CNEP/TCU são requisitos cogentes "
        "e eliminatórios, devendo ser integralmente comprovados pela licitante adjudicatária no momento da contratação."
    )
    r_nota.font.size = Pt(9)
    r_nota.font.color.rgb = RGBColor(51, 65, 85)

    # 4. Quadro Resumo de Fornecedores Auditados
    suppliers = _extract_unique_suppliers_from_report(report)

    p_s3 = doc.add_paragraph()
    p_s3.paragraph_format.space_before = Pt(12)
    p_s3.paragraph_format.space_after = Pt(4)
    r_s3 = p_s3.add_run(f"3. QUADRO RESUMO DE REGULARIDADE DOS FORNECEDORES ({len(suppliers)} EMPRESA(S) AUDITADA(S))")
    r_s3.font.bold = True
    r_s3.font.size = Pt(11)

    t_sum = doc.add_table(rows=1 + len(suppliers), cols=5)
    t_sum.alignment = WD_TABLE_ALIGNMENT.CENTER
    sum_widths = (Inches(0.4), Inches(2.8), Inches(1.2), Inches(1.4), Inches(1.2))
    sum_headers = ["Nº", "Fornecedor / CNPJ", "Preço Cotado", "Regularidade Fiscal", "Sanções Ativas"]
    for c_i, h in enumerate(sum_headers):
        cell = t_sum.rows[0].cells[c_i]
        cell.width = sum_widths[c_i]
        set_cell_background(cell, "1E293B")
        set_cell_margins(cell, 60, 60, 50, 50)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(h)
        r.font.bold = True
        r.font.size = Pt(8.5)
        r.font.color.rgb = RGBColor(255, 255, 255)

    compliance_results = {}
    for idx, s in enumerate(suppliers):
        c_res = check_supplier_full_compliance(s["cnpj_clean"], s["name"])
        compliance_results[s["cnpj_clean"]] = c_res
        row = t_sum.rows[1 + idx]
        bg = "FFFFFF" if idx % 2 == 0 else "F8FAFC"

        has_sanctions = c_res.get("has_sanctions", False)
        has_expired = c_res.get("has_expired_certidoes", False)

        status_fiscal_txt = "REGULAR" if not has_expired else "CERTIDÃO VENCIDA"
        status_sancao_txt = "NADA CONSTA" if not has_sanctions else "SANÇÃO ATIVA"

        vals = [
            f"{idx + 1:02d}",
            f"{s['name'][:35]}\nCNPJ: {s['cnpj']}",
            _format_brl(s["price"]),
            status_fiscal_txt,
            status_sancao_txt,
        ]
        for c_i, v in enumerate(vals):
            cell = row.cells[c_i]
            cell.width = sum_widths[c_i]
            set_cell_background(cell, bg)
            set_cell_margins(cell, 50, 50, 50, 50)
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.RIGHT if c_i == 2 else (WD_ALIGN_PARAGRAPH.CENTER if c_i in (0, 3, 4) else WD_ALIGN_PARAGRAPH.LEFT)
            r = p.add_run(v)
            r.font.size = Pt(8.5)
            if c_i == 3:
                r.font.bold = True
                r.font.color.rgb = RGBColor(22, 101, 52) if not has_expired else RGBColor(180, 83, 9)
            elif c_i == 4:
                r.font.bold = True
                r.font.color.rgb = RGBColor(22, 101, 52) if not has_sanctions else RGBColor(190, 18, 60)

    # 5. Fichas Detalhadas das 7 Certidões por Empresa
    p_s4 = doc.add_paragraph()
    p_s4.paragraph_format.space_before = Pt(16)
    p_s4.paragraph_format.space_after = Pt(4)
    r_s4 = p_s4.add_run("4. RAIO-X INDIVIDUAL DE CERTIDÕES E CONSULTAS EM TEMPO REAL")
    r_s4.font.bold = True
    r_s4.font.size = Pt(11)

    for idx, s in enumerate(suppliers):
        c_res = compliance_results.get(s["cnpj_clean"]) or check_supplier_full_compliance(s["cnpj_clean"], s["name"])
        dossie = c_res.get("dossie_certidoes", [])

        # Bloco da Empresa
        p_emp = doc.add_paragraph()
        p_emp.paragraph_format.space_before = Pt(10)
        p_emp.paragraph_format.space_after = Pt(2)
        r_emp = p_emp.add_run(f"4.{idx + 1} Fornecedor: {s['name']} (CNPJ: {s['cnpj']})")
        r_emp.font.bold = True
        r_emp.font.size = Pt(10)
        r_emp.font.color.rgb = RGBColor(30, 41, 59)

        # Tabela das 7 Certidões
        t_cert = doc.add_table(rows=1 + len(dossie), cols=5)
        t_cert.alignment = WD_TABLE_ALIGNMENT.CENTER
        cert_widths = (Inches(2.5), Inches(1.5), Inches(0.9), Inches(1.1), Inches(1.0))
        cert_headers = ["Certidão / Registro Obrigatório", "Órgão Emissor", "Situação", "Validade / Controle", "Verificação"]
        for c_i, h in enumerate(cert_headers):
            cell = t_cert.rows[0].cells[c_i]
            cell.width = cert_widths[c_i]
            set_cell_background(cell, "334155")
            set_cell_margins(cell, 40, 40, 40, 40)
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r = p.add_run(h)
            r.font.bold = True
            r.font.size = Pt(8)
            r.font.color.rgb = RGBColor(255, 255, 255)

        for c_idx, cert in enumerate(dossie):
            row = t_cert.rows[1 + c_idx]
            bg = "FFFFFF" if c_idx % 2 == 0 else "F8FAFC"
            is_venc = cert.get("is_vencida", False)
            status_txt = cert.get("status", "REGULAR")

            vals = [
                cert.get("nome", ""),
                cert.get("orgao", ""),
                status_txt,
                f"{cert.get('data_validade', '')}\n{cert.get('codigo_controle', '')}",
            ]
            for col_i, v in enumerate(vals):
                cell = row.cells[col_i]
                cell.width = cert_widths[col_i]
                set_cell_background(cell, bg)
                set_cell_margins(cell, 40, 40, 40, 40)
                p = cell.paragraphs[0]
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER if col_i in (2, 3) else WD_ALIGN_PARAGRAPH.LEFT
                r = p.add_run(v)
                r.font.size = Pt(7.5)
                if col_i == 2:
                    r.font.bold = True
                    r.font.color.rgb = RGBColor(190, 18, 60) if is_venc else RGBColor(22, 101, 52)

            # Coluna 5: QR Code + Link Clicável
            cell_qr = row.cells[4]
            cell_qr.width = cert_widths[4]
            set_cell_background(cell_qr, bg)
            set_cell_margins(cell_qr, 30, 30, 30, 30)
            p_qr = cell_qr.paragraphs[0]
            p_qr.alignment = WD_ALIGN_PARAGRAPH.CENTER
            link_url = cert.get("link_oficial") or "https://portaldatransparencia.gov.br"
            try:
                qr_bytes = generate_qr_code_bytes(link_url, box_size=2)
                p_qr.add_run().add_picture(io.BytesIO(qr_bytes), width=Inches(0.55))
            except Exception:
                pass
            p_lnk = cell_qr.add_paragraph()
            p_lnk.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p_lnk.paragraph_format.space_before = Pt(0)
            add_docx_hyperlink(p_lnk, link_url, "🔗 Validar", color="0284C7", underline=True)

    # 6. Fechamento e Assinaturas
    p_sig_block = doc.add_paragraph()
    p_sig_block.paragraph_format.space_before = Pt(20)
    p_sig_block.paragraph_format.space_after = Pt(14)
    mes_ext = _get_mes_extenso(date.today().month)
    r_dt = p_sig_block.add_run(f"Juazeiro - BA, {date.today().day} de {mes_ext} de {date.today().year}.\n\n")
    r_dt.font.size = Pt(9.5)

    t_sigs = doc.add_table(rows=1, cols=2)
    t_sigs.alignment = WD_TABLE_ALIGNMENT.CENTER
    s_widths = (Inches(3.5), Inches(3.5))
    sigs = [
        (elaborador_nome.upper(), "ENCARREGADO DO SETOR DE OBTENÇÃO / AGENTE DE CONTRATAÇÃO"),
        (ordenador_nome.upper(), "ORDENADOR DE DESPESAS / AUTORIDADE COMPETENTE"),
    ]
    for i, (nome_cargo, cargo_tit) in enumerate(sigs):
        cell = t_sigs.rows[0].cells[i]
        cell.width = s_widths[i]
        set_cell_margins(cell, 80, 80, 40, 40)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.line_spacing = 1.1
        r_line = p.add_run("_________________________________________\n")
        r_line.font.bold = True
        r_n = p.add_run(f"{nome_cargo}\n")
        r_n.font.bold = True
        r_n.font.size = Pt(9)
        r_c = p.add_run(cargo_tit)
        r_c.font.size = Pt(8)
        r_c.font.color.rgb = RGBColor(71, 85, 105)

    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()


def build_supplier_compliance_dossier_pdf_bytes(
    report: Dict[str, Any],
    om_name: str = "MARINHA DO BRASIL",
    elaborador_nome: str = "ENCARREGADO DO SETOR DE OBTENÇÃO",
    ordenador_nome: str = "ORDENADOR DE DESPESAS",
) -> bytes:
    """
    Gera em PDF A4 oficial o Dossiê / Relatório Raio-X de Regularidade e Certidões
    dos Fornecedores participantes da pesquisa de preços (Art. 68 e 156 da Lei nº 14.133/2021).
    """
    try:
        from sanctions_checker import check_supplier_full_compliance
    except ImportError:
        from core.sanctions_checker import check_supplier_full_compliance

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
    title_style = ParagraphStyle("DossierTitle", parent=styles["Normal"], fontName="Helvetica-Bold", fontSize=10.5, leading=12.5, alignment=1, textColor=colors.HexColor("#0f172a"))
    sub_style = ParagraphStyle("DossierSub", parent=styles["Normal"], fontName="Helvetica", fontSize=8, leading=10, alignment=1, textColor=colors.HexColor("#475569"))
    sec_title = ParagraphStyle("DossierSec", parent=styles["Normal"], fontName="Helvetica-Bold", fontSize=9.5, leading=12, textColor=colors.HexColor("#1e293b"))
    body_style = ParagraphStyle("DossierBody", parent=styles["Normal"], fontName="Helvetica", fontSize=7.5, leading=10, textColor=colors.HexColor("#334155"))
    cell_hdr = ParagraphStyle("DossierCellHdr", parent=styles["Normal"], fontName="Helvetica-Bold", fontSize=7, leading=8.5, alignment=1, textColor=colors.white)
    cell_txt = ParagraphStyle("DossierCellTxt", parent=styles["Normal"], fontName="Helvetica", fontSize=6.8, leading=8.2, textColor=colors.HexColor("#0f172a"))
    cell_center = ParagraphStyle("DossierCellCenter", parent=styles["Normal"], fontName="Helvetica", fontSize=6.8, leading=8.2, alignment=1, textColor=colors.HexColor("#0f172a"))
    cell_right = ParagraphStyle("DossierCellRight", parent=styles["Normal"], fontName="Helvetica-Bold", fontSize=7, leading=8.5, alignment=2, textColor=colors.HexColor("#0f172a"))

    elements = []

    # 1. Brasão e Cabeçalho
    if os.path.exists(BRASAO_PATH):
        elements.append(RLImage(BRASAO_PATH, width=2.2 * cm, height=2.2 * cm))
        elements.append(Spacer(1, 3))

    header_p = (
        "<b>REPÚBLICA FEDERATIVA DO BRASIL</b><br/>"
        "<b>MINISTÉRIO DA DEFESA • COMANDO DA MARINHA</b><br/>"
        f"<b>{om_name.upper()}</b><br/>"
        "<font size='6.5' color='#475569'>SETOR DE OBTENÇÃO E PREÇOS • ASSESSORIA DE COMPLIANCE</font>"
    )
    elements.append(Paragraph(header_p, title_style))
    elements.append(Spacer(1, 6))

    tit_p = (
        "<b>DOSSIÊ CONSOLIDADO DE REGULARIDADE FISCAL, TRABALHISTA E RISCO CAUTELAR</b><br/>"
        "<font size='7' color='#475569'><i>Auditoria Conjunta de Fornecedores da Cesta de Preços — Art. 68 e 156 da Lei nº 14.133/2021</i></font>"
    )
    elements.append(Paragraph(tit_p, title_style))
    elements.append(Spacer(1, 8))

    # 2. Dados do Objeto
    code = report.get("catmat_catser_code") or "—"
    item_type = str(report.get("item_type") or "MATERIAL").upper()
    prefix = "CATSER" if item_type == "SERVICO" else "CATMAT"
    desc = report.get("item_description") or "Item de Pesquisa"
    unit_price = float(report.get("final_estimated_unit_price") or 0.0)

    t_obj = Table([
        [Paragraph(f"<b>Catálogo:</b> {prefix} {code}", cell_txt), Paragraph(f"<b>Preço Estimado (IN 65):</b> {_format_brl(unit_price)}", cell_txt)],
        [Paragraph(f"<b>Objeto:</b> {desc}", cell_txt), Paragraph("<b>Fundamentação:</b> Lei 14.133/2021 e IN SEGES 65/2021", cell_txt)]
    ], colWidths=[11.5 * cm, 6.5 * cm])
    t_obj.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#f8fafc")),
        ("BOX", (0, 0), (-1, -1), 0.5, colors.HexColor("#cbd5e1")),
        ("INNERGRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#e2e8f0")),
        ("PADDING", (0, 0), (-1, -1), 3),
    ]))
    elements.append(t_obj)
    elements.append(Spacer(1, 6))

    # 3. Nota Técnica de Governança
    nota_p = (
        "<b>DIRETRIZ DE GOVERNANÇA (Art. 23 vs Art. 68):</b> Na pesquisa de preços preparatória, cotações extraídas do Compras.gov.br e PNCP "
        "de contratações anteriores representam a realidade histórica de mercado, permanecendo válidas na amostra. "
        "A exigência de regularidade fiscal, social e trabalhista plena (Art. 68) e ausência de sanções impeditivas (Art. 156) "
        "será exigida formalmente da licitante vencedora na fase de habilitação/contratação."
    )
    elements.append(Paragraph(nota_p, body_style))
    elements.append(Spacer(1, 8))

    # 4. Tabela Resumo das Empresas
    suppliers = _extract_unique_suppliers_from_report(report)
    elements.append(Paragraph(f"<b>3. Quadro Resumo de Fornecedores ({len(suppliers)} empresa(s) auditada(s))</b>", sec_title))
    elements.append(Spacer(1, 4))

    sum_data = [
        [
            Paragraph("Nº", cell_hdr),
            Paragraph("Fornecedor / CNPJ", cell_hdr),
            Paragraph("Preço Cotado", cell_hdr),
            Paragraph("Regularidade Fiscal", cell_hdr),
            Paragraph("Sanções Ativas", cell_hdr),
        ]
    ]

    compliance_results = {}
    for idx, s in enumerate(suppliers):
        c_res = check_supplier_full_compliance(s["cnpj_clean"], s["name"])
        compliance_results[s["cnpj_clean"]] = c_res
        has_sanctions = c_res.get("has_sanctions", False)
        has_expired = c_res.get("has_expired_certidoes", False)

        cor_fisc = "#166534" if not has_expired else "#b45309"
        cor_sanc = "#166534" if not has_sanctions else "#be123c"

        sum_data.append([
            Paragraph(f"{idx+1:02d}", cell_center),
            Paragraph(f"<b>{s['name'][:35]}</b><br/>CNPJ: {s['cnpj']}", cell_txt),
            Paragraph(_format_brl(s["price"]), cell_right),
            Paragraph(f"<font color='{cor_fisc}'><b>{'REGULAR' if not has_expired else 'CERTIDÃO VENCIDA'}</b></font>", cell_center),
            Paragraph(f"<font color='{cor_sanc}'><b>{'NADA CONSTA' if not has_sanctions else 'SANÇÃO ATIVA'}</b></font>", cell_center),
        ])

    t_sum = Table(sum_data, colWidths=[1.0 * cm, 7.5 * cm, 3.0 * cm, 3.5 * cm, 3.0 * cm])
    t_sum.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1e293b")),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#cbd5e1")),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("PADDING", (0, 0), (-1, -1), 3),
    ]))
    elements.append(t_sum)
    elements.append(Spacer(1, 10))

    # 5. Fichas Individuais das 7 Certidões
    elements.append(Paragraph("<b>4. Raio-X Individual de Certidões e Registros Oficiais</b>", sec_title))
    elements.append(Spacer(1, 5))

    for idx, s in enumerate(suppliers[:8]):
        c_res = compliance_results.get(s["cnpj_clean"]) or check_supplier_full_compliance(s["cnpj_clean"], s["name"])
        dossie = c_res.get("dossie_certidoes", [])

        elements.append(Paragraph(f"<b>4.{idx+1} {s['name']} — CNPJ: {s['cnpj']}</b>", sec_title))
        elements.append(Spacer(1, 3))

        cert_data = [
            [
                Paragraph("Certidão / Registro Oficial", cell_hdr),
                Paragraph("Órgão", cell_hdr),
                Paragraph("Situação", cell_hdr),
                Paragraph("Validade / Controle", cell_hdr),
                Paragraph("Comprovação", cell_hdr),
            ]
        ]
        for c in dossie:
            is_v = c.get("is_vencida", False)
            st_col = "#be123c" if is_v else "#166534"
            st_txt = c.get("status", "REGULAR")
            lnk = c.get("link_oficial") or "https://portaldatransparencia.gov.br"

            try:
                qr_b = generate_qr_code_bytes(lnk, box_size=2)
                rl_qr = RLImage(io.BytesIO(qr_b), width=1.1 * cm, height=1.1 * cm)
            except Exception:
                rl_qr = Paragraph("QR", cell_center)

            p_lnk = Paragraph(f"<font color='#0284c7' size='6'><u><a href='{lnk}'><b>🔗 Validar</b></a></u></font>", cell_center)

            cert_data.append([
                Paragraph(c.get("nome", ""), cell_txt),
                Paragraph(c.get("orgao", ""), cell_txt),
                Paragraph(f"<font color='{st_col}'><b>{st_txt}</b></font>", cell_center),
                Paragraph(f"{c.get('data_validade', '')}<br/><font size='5.5' color='#64748b'>{c.get('codigo_controle', '')}</font>", cell_center),
                [rl_qr, p_lnk],
            ])

        t_c = Table(cert_data, colWidths=[6.5 * cm, 3.8 * cm, 2.7 * cm, 3.0 * cm, 2.0 * cm])
        t_c.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#334155")),
            ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#cbd5e1")),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("ALIGN", (4, 1), (4, -1), "CENTER"),
            ("PADDING", (0, 0), (-1, -1), 2.5),
        ]))
        elements.append(t_c)
        elements.append(Spacer(1, 6))

    # 6. Assinaturas
    elements.append(Spacer(1, 10))
    mes_ext = _get_mes_extenso(date.today().month)
    elements.append(Paragraph(f"Juazeiro - BA, {date.today().day} de {mes_ext} de {date.today().year}.", sub_style))
    elements.append(Spacer(1, 12))

    sig_data = [
        [
            Paragraph(f"____________________________________________<br/><b>{elaborador_nome.upper()}</b><br/><font size='6.5'>ENCARREGADO DO SETOR DE OBTENÇÃO</font>", sub_style),
            Paragraph(f"____________________________________________<br/><b>{ordenador_nome.upper()}</b><br/><font size='6.5'>ORDENADOR DE DESPESAS</font>", sub_style),
        ]
    ]
    t_sig = Table(sig_data, colWidths=[9.0 * cm, 9.0 * cm])
    t_sig.setStyle(TableStyle([
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
    ]))
    elements.append(t_sig)

    doc.build(elements)
    return buf.getvalue()


