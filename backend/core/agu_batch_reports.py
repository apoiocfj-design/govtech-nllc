# -*- coding: utf-8 -*-
"""
MÓDULO DE RELATÓRIOS CONSOLIDADOS PADRÃO AGU / SEGES (IN 65/2021 & LEI 14.133/2021)
===================================================================================
Contém:
1. Conversor financeiro oficial para valor por extenso em Português do Brasil.
2. Gerador da Nota Técnica AGU Oficial (.DOCX e .PDF) seguindo fielmente o modelo
   padronizado SEGES/AGU (Seções I a VI, Tabela de Consulta Direta, Tabela de Expurgos § 4.2,
   Memória de Cálculo com valor por extenso § 5.1 e Identificação dos Agentes § VI).
3. Gerador da Justificativa Consolidada do Art. 6º, § 5º da IN 65/2021 e Janela 13-24M
   (.DOCX e .PDF) — elimina dezenas de despachos redundantes por item.
4. Gerador de Relatório Detalhado Contínuo padrão Compras.gov.br agrupado por Categoria
   (Classe CNBS / PDM) e separado por janela temporal (<=12M vs 13-24M) — ultra-veloz para
   lotes de 50, 150 ou 300+ itens sem gargalos de memória ou timeout.
"""

from __future__ import annotations

import io
import re
from collections import defaultdict
from datetime import datetime
from typing import Any, Dict, List, Optional, Tuple
from xml.sax.saxutils import escape as xml_escape

import docx
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn
from docx.shared import Cm, Inches, Pt, RGBColor

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import cm
from reportlab.platypus import (
    KeepTogether,
    PageBreak,
    Paragraph,
    SimpleDocTemplate,
    Spacer,
    Table,
    TableStyle,
)

# =============================================================================
# UTILITÁRIOS DE FORMATAÇÃO E VALOR POR EXTENSO (PT-BR OFICIAL)
# =============================================================================

_MESES_PT = [
    "", "janeiro", "fevereiro", "março", "abril", "maio", "junho",
    "julho", "agosto", "setembro", "outubro", "novembro", "dezembro",
]


def _fmt_brl(val: float) -> str:
    try:
        v = float(val or 0.0)
    except Exception:
        v = 0.0
    return f"R$ {v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def _data_extenso_hoje() -> str:
    hoje = datetime.now()
    return f"{hoje.day} de {_MESES_PT[hoje.month]} de {hoje.year}"


def _int_to_pt_br(n: int) -> str:
    """Converte inteiro não-negativo (até 999 trilhões) para extenso oficial PT-BR."""
    if n == 0:
        return "zero"

    unidades = ["", "um", "dois", "três", "quatro", "cinco", "seis", "sete", "oito", "nove"]
    especiais = [
        "dez", "onze", "doze", "treze", "quatorze", "quinze",
        "dezesseis", "dezessete", "dezoito", "dezenove",
    ]
    dezenas = [
        "", "", "vinte", "trinta", "quarenta", "cinquenta",
        "sessenta", "setenta", "oitenta", "noventa",
    ]
    centenas = [
        "", "cento", "duzentos", "trezentos", "quatrocentos", "quinhentos",
        "seiscentos", "setecentos", "oitocentos", "novecentos",
    ]

    def _centena(num: int) -> str:
        if num == 0:
            return ""
        if num == 100:
            return "cem"
        c = num // 100
        resto = num % 100
        parts = []
        if c > 0:
            parts.append(centenas[c])
        if 10 <= resto <= 19:
            parts.append(especiais[resto - 10])
        else:
            d = resto // 10
            u = resto % 10
            if d > 0:
                parts.append(dezenas[d])
            if u > 0:
                parts.append(unidades[u])
        return " e ".join(parts)

    escalas = [
        ("", ""),
        ("mil", "mil"),
        ("milhão", "milhões"),
        ("bilhão", "bilhões"),
        ("trilhão", "trilhões"),
    ]

    grupos = []
    temp = n
    idx = 0
    while temp > 0 and idx < len(escalas):
        g = temp % 1000
        if g > 0:
            grupos.append((g, idx))
        temp //= 1000
        idx += 1

    grupos.reverse()
    partes: List[str] = []
    for pos, (g, esc_idx) in enumerate(grupos):
        sing, plur = escalas[esc_idx]
        if esc_idx == 1 and g == 1:
            txt = "mil"
        else:
            base = _centena(g)
            suf = sing if g == 1 else plur
            txt = f"{base} {suf}".strip()
        if pos > 0:
            # Conector oficial PT-BR: "e" quando último grupo < 100 ou centena exata
            if pos == len(grupos) - 1 and (g < 100 or g % 100 == 0):
                partes.append(f"e {txt}")
            else:
                partes.append(f", {txt}")
        else:
            partes.append(txt)

    return "".join(partes).replace(" ,", ",").strip()


def brl_por_extenso(valor: float) -> str:
    """Converte valor monetário em Reais para texto por extenso (ex: 'mil reais e cinquenta centavos')."""
    try:
        v = round(abs(float(valor or 0.0)), 2)
    except Exception:
        v = 0.0
    inteiro = int(v)
    centavos = int(round((v - inteiro) * 100))

    partes: List[str] = []
    if inteiro > 0:
        ext_int = _int_to_pt_br(inteiro)
        # "um milhão de reais" quando múltiplo exato de milhão/bilhão/trilhão
        if inteiro >= 1_000_000 and inteiro % 1_000_000 == 0:
            moeda = "de reais"
        else:
            moeda = "real" if inteiro == 1 else "reais"
        partes.append(f"{ext_int} {moeda}")

    if centavos > 0:
        ext_cent = _int_to_pt_br(centavos)
        c_lbl = "centavo" if centavos == 1 else "centavos"
        partes.append(f"{ext_cent} {c_lbl}")

    if not partes:
        return "zero reais"
    return " e ".join(partes)


def resolve_item_category(item: Dict[str, Any]) -> Tuple[str, str]:
    """
    Resolve a Categoria (Classe CNBS ou PDM) de um item do lote para agrupamento inteligente.
    Retorna (slug_arquivo, nome_legivel_categoria).
    """
    rep = item.get("report") or {}
    cat = (
        str(rep.get("nome_classe") or "").strip()
        or str(rep.get("nome_pdm") or "").strip()
        or str(item.get("nome_classe") or "").strip()
        or str(item.get("nome_pdm") or "").strip()
    )
    if not cat:
        # Tenta extrair das cotações internas
        for q in (rep.get("valid_quotes") or []) + (rep.get("purged_quotes") or []):
            cat = str(q.get("nome_classe") or q.get("nome_pdm") or "").strip()
            if cat:
                break
    if not cat:
        itype = str(rep.get("item_type") or item.get("type") or "MATERIAL").strip().upper()
        cat = "SERVIÇOS GERAIS" if itype == "SERVICO" else "MATERIAIS DIVERSOS"

    clean_slug = re.sub(r"[^A-Za-z0-9]+", "_", cat.upper())[:42].strip("_") or "GERAL"
    return clean_slug, cat.upper()


def is_item_extended_24m(item: Dict[str, Any]) -> bool:
    """Identifica se o item utilizou a janela excepcional de 13 a 24 meses."""
    rep = item.get("report") or {}
    if rep.get("is_extended_24m") is True or item.get("is_extended_24m") is True:
        return True
    try:
        if int(rep.get("search_window_months") or item.get("search_window_months") or 12) > 12:
            return True
    except Exception:
        pass
    for q in (rep.get("valid_quotes") or []):
        if q.get("is_extended_24m"):
            return True
    return False


def _set_cell_shading(cell, fill_hex: str) -> None:
    tcPr = cell._tc.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>')
    tcPr.append(shd)


def _set_table_borders(table, color: str = "CBD5E1") -> None:
    tblPr = table._tbl.tblPr
    borders = parse_xml(
        f'<w:tblBorders {nsdecls("w")}>'
        f'<w:top w:val="single" w:sz="4" w:space="0" w:color="{color}"/>'
        f'<w:left w:val="single" w:sz="4" w:space="0" w:color="{color}"/>'
        f'<w:bottom w:val="single" w:sz="4" w:space="0" w:color="{color}"/>'
        f'<w:right w:val="single" w:sz="4" w:space="0" w:color="{color}"/>'
        f'<w:insideH w:val="single" w:sz="4" w:space="0" w:color="{color}"/>'
        f'<w:insideV w:val="single" w:sz="4" w:space="0" w:color="{color}"/>'
        f'</w:tblBorders>'
    )
    tblPr.append(borders)


def _safe_xml(text: Any, max_len: Optional[int] = None) -> str:
    s = str(text if text is not None else "")
    if max_len is not None and len(s) > max_len:
        s = s[: max_len - 1] + "…"
    return xml_escape(s)


def _categorize_purge_reason(quote: Dict[str, Any]) -> str:
    st = str(quote.get("purge_status") or "").upper()
    if "INEXEQUIVEL" in st:
        return "Inexequível"
    if "ACIMA" in st or "SOBREPRECO" in st or "ELEVADO" in st:
        return "Excessivamente elevado"
    if "UNIDADE" in st:
        return "Inconsistente (Unidade divergente)"
    if "MANUAL" in st:
        return "Desconsiderado na análise crítica"
    return "Inconsistente / Outlier estatístico"


def _compute_batch_stats(items_results: List[Dict[str, Any]]) -> Dict[str, Any]:
    total_geral = 0.0
    total_valid_quotes = 0
    total_purged_quotes = 0
    items_12m: List[Tuple[int, Dict[str, Any]]] = []
    items_24m: List[Tuple[int, Dict[str, Any]]] = []
    items_under_3: List[Tuple[int, Dict[str, Any]]] = []
    homogeneos = 0
    heterogeneos = 0
    incisos_usados = set()
    purged_rows: List[Dict[str, Any]] = []
    categorias: Dict[str, Dict[str, Any]] = {}

    for idx, it in enumerate(items_results, start=1):
        rep = it.get("report") or {}
        qty = float(it.get("qty", 1) or 1)
        u_est = float(rep.get("final_estimated_unit_price") or 0.0)
        tot = qty * u_est
        total_geral += tot

        vqs = rep.get("valid_quotes") or []
        pqs = rep.get("purged_quotes") or []
        total_valid_quotes += len(vqs)
        total_purged_quotes += len(pqs)

        if rep.get("is_homogeneous"):
            homogeneos += 1
        else:
            heterogeneos += 1

        if len(vqs) < 3:
            items_under_3.append((idx, it))

        is_24 = is_item_extended_24m(it)
        if is_24:
            items_24m.append((idx, it))
        else:
            items_12m.append((idx, it))

        _, cat_name = resolve_item_category(it)
        if cat_name not in categorias:
            categorias[cat_name] = {"count": 0, "total": 0.0, "count_12m": 0, "count_24m": 0}
        categorias[cat_name]["count"] += 1
        categorias[cat_name]["total"] += tot
        if is_24:
            categorias[cat_name]["count_24m"] += 1
        else:
            categorias[cat_name]["count_12m"] += 1

        for q in vqs:
            src = str(q.get("source_inciso") or "Inciso I - Compras.gov.br")
            if "Inciso II" in src:
                incisos_usados.add("Inciso II (contratações similares da Administração Pública — Módulo Contratos)")
            elif "Inciso III" in src:
                incisos_usados.add("Inciso III (mídia especializada / tabelas de referência)")
            elif "Inciso IV" in src:
                incisos_usados.add("Inciso IV (pesquisa direta com fornecedores)")
            else:
                incisos_usados.add("Inciso I (composição de custos / sistemas oficiais de governo — Compras.gov.br e PNCP)")

        code = str(rep.get("catmat_catser_code") or it.get("code") or "—")
        for pq in pqs:
            purged_rows.append({
                "item_idx": idx,
                "code": code,
                "fonte": (
                    f"Item {idx} ({code}) | {pq.get('source_inciso') or 'Compras.gov.br'} | "
                    f"UASG {pq.get('buyer_uasg') or '—'} ({pq.get('purchase_number_year') or '—'}) | "
                    f"{str(pq.get('supplier_name') or 'Fornecedor')[:45]}"
                ),
                "preco": float(pq.get("unit_price") or 0.0),
                "caracterizacao": _categorize_purge_reason(pq),
                "justificativa": str(pq.get("purge_justification") or "Outlier estatístico conforme Art. 6º IN 65/2021."),
            })

    if not incisos_usados:
        incisos_usados.add("Inciso I (sistemas oficiais de governo — Compras.gov.br e PNCP)")

    return {
        "total_items": len(items_results),
        "total_geral": total_geral,
        "total_valid_quotes": total_valid_quotes,
        "total_purged_quotes": total_purged_quotes,
        "items_12m": items_12m,
        "items_24m": items_24m,
        "items_under_3": items_under_3,
        "homogeneos": homogeneos,
        "heterogeneos": heterogeneos,
        "incisos_usados": sorted(incisos_usados),
        "purged_rows": purged_rows,
        "categorias": categorias,
    }


# =============================================================================
# 1. NOTA TÉCNICA OFICIAL AGU (MODELO EXATO SEGES/AGU) — .DOCX E .PDF
# =============================================================================

def build_agu_nota_tecnica_docx_bytes(
    items_results: List[Dict[str, Any]],
    om_name: str = "MARINHA DO BRASIL",
    ordenador_nome: str = "ORDENADOR DE DESPESAS",
    elaborador_nome: str = "ENCARREGADO DO SETOR DE OBTENÇÃO",
    elaborador_matricula: str = "A INFORMAR",
    cidade: str = "Rio de Janeiro - RJ",
    nota_num: str = "",
    objeto_geral: str = "",
    direct_suppliers: Optional[List[Dict[str, str]]] = None,
) -> bytes:
    """
    Gera a NOTA TÉCNICA DE PESQUISA DE PREÇOS em Word (.DOCX) seguindo rigorosamente
    a estrutura oficial do modelo da Advocacia-Geral da União (AGU) / SEGES (Guia de
    Padronização dos Procedimentos de Contratação — IN SEGES/ME nº 65/2021 e Lei nº 14.133/2021).
    """
    stats = _compute_batch_stats(items_results)
    ano_atual = datetime.now().year
    nt_label = nota_num.strip() or f"____/{ano_atual}"

    if not objeto_geral.strip():
        cat_list = list(stats["categorias"].keys())[:4]
        cat_str = ", ".join(cat_list) if cat_list else "materiais e/ou serviços demandados"
        objeto_geral = (
            f"aquisição/contratação de lote composto por {stats['total_items']} item(ns) "
            f"enquadrado(s) nas categorias {cat_str}, para atendimento das necessidades administrativas e "
            f"operacionais de {om_name.upper()}"
        )

    doc = docx.Document()
    for s in doc.sections:
        s.top_margin = Cm(2.0)
        s.bottom_margin = Cm(2.0)
        s.left_margin = Cm(2.2)
        s.right_margin = Cm(2.0)

    normal_style = doc.styles["Normal"]
    normal_style.font.name = "Calibri"
    normal_style.font.size = Pt(10.5)

    def _add_heading(text: str):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(10)
        p.paragraph_format.space_after = Pt(4)
        p.paragraph_format.keep_with_next = True
        r = p.add_run(text)
        r.bold = True
        r.font.size = Pt(11)
        r.font.color.rgb = RGBColor(15, 23, 42)
        return p

    def _add_body(text: str, bold_prefix: str = "", justify: bool = True, italic: bool = False):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY if justify else WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.space_after = Pt(4)
        p.paragraph_format.line_spacing = 1.15
        if bold_prefix:
            rb = p.add_run(bold_prefix)
            rb.bold = True
            rb.font.size = Pt(10.5)
        r = p.add_run(text)
        r.italic = italic
        r.font.size = Pt(10.5)
        return p

    # Cabeçalho Institucional
    p_hdr = doc.add_paragraph()
    p_hdr.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_hdr.paragraph_format.space_after = Pt(2)
    r_om = p_hdr.add_run(f"{om_name.upper()}\n")
    r_om.bold = True
    r_om.font.size = Pt(11.5)
    r_sub = p_hdr.add_run("PESQUISA DE PREÇOS — LEI Nº 14.133/2021 E IN SEGES/ME Nº 65/2021\n")
    r_sub.font.size = Pt(9)
    r_sub.font.color.rgb = RGBColor(71, 85, 105)

    p_nt = doc.add_paragraph()
    p_nt.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_nt.paragraph_format.space_before = Pt(6)
    p_nt.paragraph_format.space_after = Pt(10)
    r_nt = p_nt.add_run(f"NOTA TÉCNICA nº {nt_label}")
    r_nt.bold = True
    r_nt.font.size = Pt(13)
    r_nt.font.color.rgb = RGBColor(30, 45, 61)

    # SEÇÃO I
    _add_heading("I - OBJETO DA CONTRATAÇÃO")
    _add_body(f"O objeto da presente contratação é {objeto_geral}.")

    # SEÇÃO II
    _add_heading("II - FONTES CONSULTADAS")
    incisos_txt = "; ".join(stats["incisos_usados"])
    _add_body(
        f"Para a definição do valor estimado da contratação foram utilizados os parâmetros dos "
        f"{incisos_txt}, nos termos do art. 5º da Instrução Normativa SEGES/ME nº 65, de 7 de julho de 2021 "
        f"(norma regente sob a Lei nº 14.133/2021, em evolução procedimental à IN SEGES/ME nº 73/2020), "
        f"com extração automatizada diretamente das bases oficiais de Dados Abertos do Compras.gov.br "
        f"(Módulos Pesquisa de Preço, ARP e Contratos) e validação de links perante o Portal Nacional de "
        f"Contratações Públicas (PNCP).",
        bold_prefix="2.1. ",
    )
    _add_body(
        "Foram priorizadas a consulta aos sistemas oficiais de governo e às contratações similares feitas "
        "pela Administração Pública, em conformidade com o artigo 5º, § 1º, da IN SEGES/ME nº 65/2021. "
        f"Do universo de {stats['total_items']} item(ns) pesquisado(s), {len(stats['items_12m'])} item(ns) "
        f"obtiveram convergência de preços homologados na janela prioritária de até 12 (doze) meses.",
        bold_prefix="2.2. ",
    )
    if stats["items_24m"]:
        codes_24 = ", ".join(
            f"Item {idx} (Cód. {(it.get('report') or {}).get('catmat_catser_code') or it.get('code')})"
            for idx, it in stats["items_24m"][:15]
        )
        if len(stats["items_24m"]) > 15:
            codes_24 += f" e mais {len(stats['items_24m']) - 15} item(ns)"
        _add_body(
            f"Excepcionalmente, para {len(stats['items_24m'])} item(ns) sem compras homologadas nos últimos "
            f"12 (doze) meses ({codes_24}), o motor de busca ampliou automaticamente a janela temporal para "
            f"até 24 (vinte e quatro) meses, segregando tais registros em pasta apartada no dossiê do processo "
            f"('02_Pesquisas_13_a_24_Meses_Excepcional'), conforme diretrizes do Caderno de Logística de Pesquisa "
            f"de Preços (SEGES/MGI), recomendando-se avaliação de atualização monetária pelo IPCA/IBGE.",
            bold_prefix="2.2.1. ",
        )

    _add_body(
        "Na consulta direta com fornecedores (Art. 5º, inciso IV da IN SEGES/ME nº 65/2021), "
        "registra-se o seguinte quadro de controle:",
        bold_prefix="2.3. ",
    )

    t_forn = doc.add_table(rows=1, cols=3)
    t_forn.alignment = WD_TABLE_ALIGNMENT.CENTER
    _set_table_borders(t_forn)
    f_hdrs = ["Fornecedor", "Apresentou resposta?", "Justificativa para escolha"]
    for ci, h in enumerate(f_hdrs):
        cell = t_forn.rows[0].cells[ci]
        cell.text = h
        _set_cell_shading(cell, "1E2D3D")
        for p in cell.paragraphs:
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for r in p.runs:
                r.bold = True
                r.font.size = Pt(9)
                r.font.color.rgb = RGBColor(255, 255, 255)

    if direct_suppliers:
        for sup in direct_suppliers:
            rc = t_forn.add_row().cells
            rc[0].text = str(sup.get("fornecedor") or "—")
            rc[1].text = str(sup.get("respondeu") or "N")
            rc[2].text = str(sup.get("justificativa") or "—")
            for c in rc:
                for p in c.paragraphs:
                    for r in p.runs:
                        r.font.size = Pt(8.5)
    else:
        rc = t_forn.add_row().cells
        rc[0].text = "Não aplicável na etapa inicial (fontes oficiais Incisos I e II suficientes)"
        rc[1].text = "N/A"
        rc[2].text = (
            "Priorizados os parâmetros dos incisos I e II do art. 5º da IN 65/2021 (Compras.gov.br e Contratos Públicos). "
            "Caso venham a ser juntadas cotações diretas complementares para itens com baixa liquidez, preencher este quadro."
        )
        for c in rc:
            for p in c.paragraphs:
                for r in p.runs:
                    r.font.size = Pt(8.5)

    doc.add_paragraph().paragraph_format.space_after = Pt(2)

    if not stats["items_under_3"]:
        _add_body(
            "O preço estimado para a contratação considerou o mínimo de três cotações válidas em todos os "
            f"{stats['total_items']} item(ns) do lote, nos termos do artigo 6º, § 5º, da IN SEGES/ME nº 65/2021.",
            bold_prefix="2.4. ",
        )
    else:
        under_list = ", ".join(
            f"Item {idx} ({len((it.get('report') or {}).get('valid_quotes') or [])} cot.)"
            for idx, it in stats["items_under_3"][:20]
        )
        if len(stats["items_under_3"]) > 20:
            under_list += f" e outros {len(stats['items_under_3']) - 20} item(ns)"
        _add_body(
            f"Em {stats['total_items'] - len(stats['items_under_3'])} item(ns), o preço estimado considerou o "
            f"mínimo de três cotações válidas (art. 6º, § 5º da IN 65/2021). Todavia, em {len(stats['items_under_3'])} "
            f"item(ns) ({under_list}), não foi possível obter o mínimo de três preços válidos nas bases governamentais "
            f"em razão de especificidade técnica, baixa frequência licitatória ou saneamento de outliers, constando "
            f"justificativa circunstanciada consolidada em documento apartado ('03_Justificativa_Consolidada_Art6_Par5') "
            f"para apreciação e aprovação da autoridade competente.",
            bold_prefix="2.4. ",
        )

    # SEÇÃO III
    _add_heading("III - SÉRIE DE PREÇOS COLETADOS")
    _add_body(
        f"Foram coletadas e submetidas à análise crítica {stats['total_valid_quotes'] + stats['total_purged_quotes']} "
        f"cotações oficiais nas bases governamentais, das quais {stats['total_valid_quotes']} compuseram validamente "
        f"a cesta de preços e {stats['total_purged_quotes']} foram desconsideradas por divergência de unidade de "
        f"fornecimento, inexequibilidade ou sobrepreço estatístico. A integralidade das empresas e preços coletados "
        f"consta da aba 'Mapa Geral - Todas Cotações' do arquivo '00_Mapa_Comparativo_Geral_Completo.xlsx' e nos "
        f"Relatórios Detalhados agrupados por Categoria (Classe CNBS/PDM) integrantes deste dossiê:"
    )

    # Tabela resumo por categoria
    t_cat = doc.add_table(rows=1, cols=5)
    t_cat.alignment = WD_TABLE_ALIGNMENT.CENTER
    _set_table_borders(t_cat)
    for ci, h in enumerate(["Categoria / Classe CNBS", "Itens", "Até 12M", "13-24M", "Subtotal Estimado"]):
        cell = t_cat.rows[0].cells[ci]
        cell.text = h
        _set_cell_shading(cell, "1E2D3D")
        for p in cell.paragraphs:
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for r in p.runs:
                r.bold = True
                r.font.size = Pt(8.5)
                r.font.color.rgb = RGBColor(255, 255, 255)

    for cat_name, cinfo in sorted(stats["categorias"].items()):
        rc = t_cat.add_row().cells
        rc[0].text = cat_name[:60]
        rc[1].text = str(cinfo["count"])
        rc[2].text = str(cinfo["count_12m"])
        rc[3].text = str(cinfo["count_24m"])
        rc[4].text = _fmt_brl(cinfo["total"])
        for ci, c in enumerate(rc):
            for p in c.paragraphs:
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT if ci == 0 else (WD_ALIGN_PARAGRAPH.RIGHT if ci == 4 else WD_ALIGN_PARAGRAPH.CENTER)
                for r in p.runs:
                    r.font.size = Pt(8.5)

    # SEÇÃO IV
    _add_heading("IV - METODOLOGIA PARA OBTENÇÃO DO PREÇO ESTIMADO")
    _add_body(
        f"A obtenção do preço estimado deu-se de forma individualizada por item com base no critério estatístico "
        f"objetivo de dispersão do Coeficiente de Variação (CV = Desvio Padrão / Média × 100), adotando-se a "
        f"MÉDIA ARITMÉTICA para as séries homogêneas (CV ≤ 25,00% — {stats['homogeneos']} item(ns)) e a MEDIANA "
        f"para as séries heterogêneas (CV > 25,00% — {stats['heterogeneos']} item(ns)), em estrita consonância "
        f"com o art. 6º da IN SEGES/ME nº 65/2021 e com a jurisprudência do Tribunal de Contas da União (TCU).",
        bold_prefix="4.1. ",
    )
    _add_body(
        "Dentro dos preços coletados, foram desconsiderados aqueles inexequíveis, inconsistentes ou "
        "excessivamente elevados, conforme demonstrativo abaixo (limite amostral exibido até 60 registros "
        "mais representativos; relação integral na aba 'Expurgos e Inconsistências' do Mapa Geral Excel):",
        bold_prefix="4.2. ",
    )

    t_exp = doc.add_table(rows=1, cols=4)
    t_exp.alignment = WD_TABLE_ALIGNMENT.CENTER
    _set_table_borders(t_exp)
    for ci, h in enumerate(["Fonte", "Preço", "Caracterização", "Justificativa"]):
        cell = t_exp.rows[0].cells[ci]
        cell.text = h
        _set_cell_shading(cell, "7F1D1D")
        for p in cell.paragraphs:
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for r in p.runs:
                r.bold = True
                r.font.size = Pt(8.5)
                r.font.color.rgb = RGBColor(255, 255, 255)

    if stats["purged_rows"]:
        for row_e in stats["purged_rows"][:60]:
            rc = t_exp.add_row().cells
            rc[0].text = row_e["fonte"][:95]
            rc[1].text = _fmt_brl(row_e["preco"])
            rc[2].text = row_e["caracterizacao"]
            rc[3].text = row_e["justificativa"][:140]
            for ci, c in enumerate(rc):
                for p in c.paragraphs:
                    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT if ci == 1 else WD_ALIGN_PARAGRAPH.LEFT
                    for r in p.runs:
                        r.font.size = Pt(7.8)
    else:
        rc = t_exp.add_row().cells
        rc[0].text = "Não houve desconsideração de preços na presente pesquisa."
        rc[1].text = "—"
        rc[2].text = "—"
        rc[3].text = "Todas as cotações coletadas apresentaram homogeneidade e compatibilidade com o objeto."
        for c in rc:
            for p in c.paragraphs:
                for r in p.runs:
                    r.font.size = Pt(8)

    # SEÇÃO V
    _add_heading("V - MEMÓRIA DE CÁLCULO E CONCLUSÃO")
    valor_extenso = brl_por_extenso(stats["total_geral"])
    _add_body(
        f"O preço estimado da contratação é {_fmt_brl(stats['total_geral'])} ({valor_extenso}), "
        f"conforme memória de cálculo resumida abaixo:",
        bold_prefix="5.1. ",
    )

    t_mem = doc.add_table(rows=1, cols=8)
    t_mem.alignment = WD_TABLE_ALIGNMENT.CENTER
    _set_table_borders(t_mem)
    mem_hdrs = ["Item", "Cód.", "Descrição", "Janela", "Unid.", "Qtde", "Unit. Est.", "Total Est."]
    for ci, h in enumerate(mem_hdrs):
        cell = t_mem.rows[0].cells[ci]
        cell.text = h
        _set_cell_shading(cell, "1E2D3D")
        for p in cell.paragraphs:
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for r in p.runs:
                r.bold = True
                r.font.size = Pt(8)
                r.font.color.rgb = RGBColor(255, 255, 255)

    for idx, it in enumerate(items_results, start=1):
        rep = it.get("report") or {}
        qty = float(it.get("qty", 1) or 1)
        u_est = float(rep.get("final_estimated_unit_price") or 0.0)
        tot = qty * u_est
        rc = t_mem.add_row().cells
        rc[0].text = str(idx)
        rc[1].text = str(rep.get("catmat_catser_code") or it.get("code") or "—")
        rc[2].text = str(rep.get("item_description") or it.get("desc") or "Item")[:65]
        rc[3].text = "13-24M" if is_item_extended_24m(it) else "≤12M"
        rc[4].text = str(rep.get("target_supply_unit") or it.get("unit") or "UN")
        rc[5].text = f"{qty:g}"
        rc[6].text = _fmt_brl(u_est)
        rc[7].text = _fmt_brl(tot)
        if idx % 2 == 0:
            for c in rc:
                _set_cell_shading(c, "F8FAFC")
        for ci, c in enumerate(rc):
            for p in c.paragraphs:
                p.alignment = (
                    WD_ALIGN_PARAGRAPH.LEFT if ci == 2
                    else (WD_ALIGN_PARAGRAPH.RIGHT if ci in (6, 7) else WD_ALIGN_PARAGRAPH.CENTER)
                )
                for r in p.runs:
                    r.font.size = Pt(7.5)
                    if ci in (6, 7):
                        r.bold = True

    tot_rc = t_mem.add_row().cells
    tot_rc[0].text = "TOTAL"
    tot_rc[2].text = f"VALOR GLOBAL ESTIMADO ({stats['total_items']} ITENS)"
    tot_rc[7].text = _fmt_brl(stats["total_geral"])
    for c in tot_rc:
        _set_cell_shading(c, "FEF3C7")
        for p in c.paragraphs:
            for r in p.runs:
                r.bold = True
                r.font.size = Pt(8)

    doc.add_paragraph().paragraph_format.space_after = Pt(2)
    _add_body(
        "Após a realização de pesquisa de preços em conformidade com a IN SEGES/ME nº 65/2021, "
        "certifica-se que o preço estimado para a presente contratação é compatível com os praticados no mercado.",
        bold_prefix="5.2. ",
    )

    # SEÇÃO VI
    _add_heading("VI - IDENTIFICAÇÃO DOS AGENTES RESPONSÁVEIS PELA PESQUISA DE PREÇOS")
    _add_body(
        f"A presente pesquisa de preços foi conduzida por: {elaborador_nome}, matrícula nº {elaborador_matricula}.",
        bold_prefix="6.1. ",
    )
    p_dt = doc.add_paragraph()
    p_dt.paragraph_format.space_before = Pt(8)
    p_dt.paragraph_format.space_after = Pt(18)
    p_dt.add_run(f"{cidade}, {_data_extenso_hoje()}.").font.size = Pt(10.5)

    t_sig = doc.add_table(rows=1, cols=2)
    t_sig.alignment = WD_TABLE_ALIGNMENT.CENTER
    c0, c1 = t_sig.rows[0].cells
    p0 = c0.paragraphs[0]
    p0.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p0.add_run("________________________________________\n").font.size = Pt(9)
    r0 = p0.add_run(f"{elaborador_nome}\n")
    r0.bold = True
    r0.font.size = Pt(9.5)
    p0.add_run(f"Agente Responsável pela Pesquisa de Preços\nMatrícula: {elaborador_matricula}").font.size = Pt(8.5)

    p1 = c1.paragraphs[0]
    p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p1.add_run("________________________________________\n").font.size = Pt(9)
    r1 = p1.add_run(f"{ordenador_nome}\n")
    r1.bold = True
    r1.font.size = Pt(9.5)
    p1.add_run("Autoridade Competente / Ordenador de Despesas").font.size = Pt(8.5)

    p_src = doc.add_paragraph()
    p_src.paragraph_format.space_before = Pt(12)
    r_src = p_src.add_run(
        "Fonte do Modelo: Guia de Padronização dos Procedimentos de Contratação — AGU / SEGES "
        "(https://www.gov.br/compras/pt-br/acesso-a-informacao/manuais/manual-fase-interna/guia-de-padronizacao-dos-procedimentos-de-contratacao.pdf)"
    )
    r_src.italic = True
    r_src.font.size = Pt(7.5)
    r_src.font.color.rgb = RGBColor(100, 116, 139)

    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()


def build_agu_nota_tecnica_pdf_bytes(
    items_results: List[Dict[str, Any]],
    om_name: str = "MARINHA DO BRASIL",
    ordenador_nome: str = "ORDENADOR DE DESPESAS",
    elaborador_nome: str = "ENCARREGADO DO SETOR DE OBTENÇÃO",
    elaborador_matricula: str = "A INFORMAR",
    cidade: str = "Rio de Janeiro - RJ",
    nota_num: str = "",
    objeto_geral: str = "",
    direct_suppliers: Optional[List[Dict[str, str]]] = None,
) -> bytes:
    """
    Gera a NOTA TÉCNICA OFICIAL AGU em PDF A4 Retrato (pronta para juntada direta no SEI).
    """
    stats = _compute_batch_stats(items_results)
    ano_atual = datetime.now().year
    nt_label = nota_num.strip() or f"____/{ano_atual}"
    if not objeto_geral.strip():
        cat_list = list(stats["categorias"].keys())[:4]
        cat_str = ", ".join(cat_list) if cat_list else "materiais e/ou serviços demandados"
        objeto_geral = (
            f"aquisição/contratação de lote composto por {stats['total_items']} item(ns) "
            f"enquadrado(s) nas categorias {cat_str}, para atendimento das necessidades administrativas e "
            f"operacionais de {om_name.upper()}"
        )

    buf = io.BytesIO()
    doc = SimpleDocTemplate(
        buf,
        pagesize=A4,
        leftMargin=1.8 * cm,
        rightMargin=1.8 * cm,
        topMargin=1.6 * cm,
        bottomMargin=1.6 * cm,
    )
    styles = getSampleStyleSheet()
    st_om = ParagraphStyle("NTOm", parent=styles["Normal"], fontName="Helvetica-Bold", fontSize=10.5, leading=13, alignment=1, textColor=colors.HexColor("#0f172a"))
    st_sub = ParagraphStyle("NTSub", parent=styles["Normal"], fontName="Helvetica", fontSize=8, leading=10, alignment=1, textColor=colors.HexColor("#475569"))
    st_nt = ParagraphStyle("NTTitle", parent=styles["Normal"], fontName="Helvetica-Bold", fontSize=12.5, leading=15, alignment=1, textColor=colors.HexColor("#1e2d3d"), spaceBefore=6, spaceAfter=8)
    st_sec = ParagraphStyle("NTSec", parent=styles["Normal"], fontName="Helvetica-Bold", fontSize=10, leading=13, textColor=colors.HexColor("#1e2d3d"), spaceBefore=9, spaceAfter=4, keepWithNext=True)
    st_body = ParagraphStyle("NTBody", parent=styles["Normal"], fontName="Helvetica", fontSize=9.2, leading=13, alignment=4, textColor=colors.HexColor("#0f172a"), spaceAfter=4)
    st_hdr = ParagraphStyle("NTHdr", parent=styles["Normal"], fontName="Helvetica-Bold", fontSize=7.2, leading=9, alignment=1, textColor=colors.white)
    st_cell = ParagraphStyle("NTCell", parent=styles["Normal"], fontName="Helvetica", fontSize=7.0, leading=8.8, textColor=colors.HexColor("#0f172a"))
    st_ctr = ParagraphStyle("NTCtr", parent=styles["Normal"], fontName="Helvetica", fontSize=7.0, leading=8.8, alignment=1, textColor=colors.HexColor("#0f172a"))
    st_rgt = ParagraphStyle("NTRgt", parent=styles["Normal"], fontName="Helvetica-Bold", fontSize=7.0, leading=8.8, alignment=2, textColor=colors.HexColor("#0f172a"))

    el = [
        Paragraph(_safe_xml(om_name.upper()), st_om),
        Paragraph("PESQUISA DE PREÇOS — LEI Nº 14.133/2021 E IN SEGES/ME Nº 65/2021", st_sub),
        Paragraph(f"NOTA TÉCNICA nº {_safe_xml(nt_label)}", st_nt),
        Paragraph("I - OBJETO DA CONTRATAÇÃO", st_sec),
        Paragraph(f"O objeto da presente contratação é {_safe_xml(objeto_geral)}.", st_body),
        Paragraph("II - FONTES CONSULTADAS", st_sec),
        Paragraph(
            f"<b>2.1.</b> Para a definição do valor estimado da contratação foram utilizados os parâmetros dos "
            f"{_safe_xml('; '.join(stats['incisos_usados']))}, nos termos do art. 5º da Instrução Normativa SEGES/ME "
            f"nº 65/2021 (em evolução procedimental à IN SEGES/ME nº 73/2020), extraídos diretamente dos Dados Abertos "
            f"do Compras.gov.br e validados perante o PNCP.",
            st_body,
        ),
        Paragraph(
            f"<b>2.2.</b> Foram priorizadas a consulta aos sistemas oficiais de governo e às contratações similares feitas "
            f"pela Administração Pública (art. 5º, § 1º da IN 65/2021). Dos {stats['total_items']} item(ns) pesquisado(s), "
            f"{len(stats['items_12m'])} convergiram na janela prioritária de até 12 meses"
            + (
                f" e {len(stats['items_24m'])} exigiram busca ampliada excepcional de 13 a 24 meses (segregados em pasta apartada no dossiê)."
                if stats["items_24m"]
                else "."
            ),
            st_body,
        ),
        Paragraph("<b>2.3.</b> Na consulta direta com fornecedores, registra-se o seguinte controle:", st_body),
    ]

    forn_rows = [[Paragraph("Fornecedor", st_hdr), Paragraph("Respondeu?", st_hdr), Paragraph("Justificativa para escolha", st_hdr)]]
    if direct_suppliers:
        for sup in direct_suppliers:
            forn_rows.append([
                Paragraph(_safe_xml(sup.get("fornecedor") or "—"), st_cell),
                Paragraph(_safe_xml(sup.get("respondeu") or "N"), st_ctr),
                Paragraph(_safe_xml(sup.get("justificativa") or "—"), st_cell),
            ])
    else:
        forn_rows.append([
            Paragraph("Priorizadas fontes oficiais (Incisos I e II — Compras.gov.br e Contratos)", st_cell),
            Paragraph("N/A", st_ctr),
            Paragraph("Consulta direta dispensada nesta etapa por suficiência das fontes oficiais (Art. 5º, § 1º IN 65/2021).", st_cell),
        ])
    t_f = Table(forn_rows, colWidths=[6.2 * cm, 2.3 * cm, 8.9 * cm], repeatRows=1)
    t_f.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1e2d3d")),
        ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#cbd5e1")),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING", (0, 0), (-1, -1), 3),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
    ]))
    el.append(t_f)
    el.append(Spacer(1, 0.2 * cm))

    if not stats["items_under_3"]:
        el.append(Paragraph(
            f"<b>2.4.</b> O preço estimado considerou o mínimo de três cotações válidas em todos os {stats['total_items']} "
            f"item(ns), nos termos do artigo 6º, § 5º da IN SEGES/ME nº 65/2021.",
            st_body,
        ))
    else:
        el.append(Paragraph(
            f"<b>2.4.</b> Em {len(stats['items_under_3'])} item(ns) não foi possível obter o mínimo de três cotações "
            f"válidas nas bases oficiais, constando a justificativa circunstanciada consolidada em peça própria "
            f"('03_Justificativa_Consolidada_Art6_Par5') para aprovação da autoridade competente (art. 6º, § 5º).",
            st_body,
        ))

    # Seção III
    el.append(Paragraph("III - SÉRIE DE PREÇOS COLETADOS", st_sec))
    el.append(Paragraph(
        f"Foram analisadas {stats['total_valid_quotes'] + stats['total_purged_quotes']} cotações oficiais "
        f"({stats['total_valid_quotes']} válidas e {stats['total_purged_quotes']} expurgadas). A relação integral "
        f"de empresas e preços consta na aba 'Mapa Geral - Todas Cotações' do Excel e nos Relatórios Detalhados "
        f"agrupados por Categoria (Classe CNBS/PDM):",
        st_body,
    ))
    cat_rows = [[
        Paragraph("Categoria / Classe CNBS", st_hdr),
        Paragraph("Itens", st_hdr),
        Paragraph("≤ 12M", st_hdr),
        Paragraph("13-24M", st_hdr),
        Paragraph("Subtotal Estimado", st_hdr),
    ]]
    for cname, ci in sorted(stats["categorias"].items()):
        cat_rows.append([
            Paragraph(_safe_xml(cname, 65), st_cell),
            Paragraph(str(ci["count"]), st_ctr),
            Paragraph(str(ci["count_12m"]), st_ctr),
            Paragraph(str(ci["count_24m"]), st_ctr),
            Paragraph(_fmt_brl(ci["total"]), st_rgt),
        ])
    t_c = Table(cat_rows, colWidths=[9.6 * cm, 1.5 * cm, 1.5 * cm, 1.6 * cm, 3.2 * cm], repeatRows=1)
    t_c.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1e2d3d")),
        ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#cbd5e1")),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING", (0, 0), (-1, -1), 2.5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 2.5),
    ]))
    el.append(t_c)

    # Seção IV
    el.append(Paragraph("IV - METODOLOGIA PARA OBTENÇÃO DO PREÇO ESTIMADO", st_sec))
    el.append(Paragraph(
        f"<b>4.1.</b> A obtenção do preço estimado observou o Coeficiente de Variação (CV): adotou-se a <b>MÉDIA</b> "
        f"para séries homogêneas (CV ≤ 25% — {stats['homogeneos']} item(ns)) e a <b>MEDIANA</b> para séries "
        f"heterogêneas (CV &gt; 25% — {stats['heterogeneos']} item(ns)), conforme art. 6º da IN SEGES/ME nº 65/2021.",
        st_body,
    ))
    el.append(Paragraph(
        "<b>4.2.</b> Dentro dos preços coletados, foram desconsiderados aqueles inexequíveis, inconsistentes ou "
        "excessivamente elevados, conforme quadro abaixo:",
        st_body,
    ))
    exp_rows = [[
        Paragraph("Fonte", st_hdr),
        Paragraph("Preço", st_hdr),
        Paragraph("Caracterização", st_hdr),
        Paragraph("Justificativa", st_hdr),
    ]]
    if stats["purged_rows"]:
        for rexp in stats["purged_rows"][:50]:
            exp_rows.append([
                Paragraph(_safe_xml(rexp["fonte"], 85), st_cell),
                Paragraph(_fmt_brl(rexp["preco"]), st_rgt),
                Paragraph(_safe_xml(rexp["caracterizacao"]), st_ctr),
                Paragraph(_safe_xml(rexp["justificativa"], 125), st_cell),
            ])
    else:
        exp_rows.append([
            Paragraph("Sem expurgos registrados.", st_cell),
            Paragraph("—", st_ctr),
            Paragraph("—", st_ctr),
            Paragraph("Todas as cotações coletadas compuseram a amostra válida.", st_cell),
        ])
    t_e = Table(exp_rows, colWidths=[5.4 * cm, 2.2 * cm, 3.0 * cm, 6.8 * cm], repeatRows=1)
    t_e.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#7f1d1d")),
        ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#cbd5e1")),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING", (0, 0), (-1, -1), 2),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
    ]))
    el.append(t_e)

    # Seção V
    el.append(Paragraph("V - MEMÓRIA DE CÁLCULO E CONCLUSÃO", st_sec))
    valor_extenso = brl_por_extenso(stats["total_geral"])
    el.append(Paragraph(
        f"<b>5.1.</b> O preço estimado da contratação é <b>{_fmt_brl(stats['total_geral'])}</b> "
        f"({_safe_xml(valor_extenso)}), conforme memória de cálculo abaixo:",
        st_body,
    ))
    mem_rows = [[
        Paragraph("#", st_hdr),
        Paragraph("Cód.", st_hdr),
        Paragraph("Descrição do Item", st_hdr),
        Paragraph("Jan.", st_hdr),
        Paragraph("Un.", st_hdr),
        Paragraph("Qtd", st_hdr),
        Paragraph("Unit. Est.", st_hdr),
        Paragraph("Total Est.", st_hdr),
    ]]
    for idx, it in enumerate(items_results, start=1):
        rep = it.get("report") or {}
        qty = float(it.get("qty", 1) or 1)
        u_est = float(rep.get("final_estimated_unit_price") or 0.0)
        tot = qty * u_est
        mem_rows.append([
            Paragraph(str(idx), st_ctr),
            Paragraph(_safe_xml(rep.get("catmat_catser_code") or it.get("code") or "—"), st_ctr),
            Paragraph(_safe_xml(rep.get("item_description") or it.get("desc") or "Item", 68), st_cell),
            Paragraph("13-24M" if is_item_extended_24m(it) else "≤12M", st_ctr),
            Paragraph(_safe_xml(rep.get("target_supply_unit") or it.get("unit") or "UN"), st_ctr),
            Paragraph(f"{qty:g}", st_ctr),
            Paragraph(_fmt_brl(u_est), st_rgt),
            Paragraph(_fmt_brl(tot), st_rgt),
        ])
    mem_rows.append([
        Paragraph("<b>TOTAL</b>", st_ctr),
        "",
        Paragraph(f"<b>VALOR GLOBAL ESTIMADO ({stats['total_items']} ITENS)</b>", st_cell),
        "", "", "", "",
        Paragraph(f"<b>{_fmt_brl(stats['total_geral'])}</b>", st_rgt),
    ])
    t_m = Table(mem_rows, colWidths=[0.9 * cm, 1.7 * cm, 6.7 * cm, 1.3 * cm, 1.1 * cm, 1.1 * cm, 2.2 * cm, 2.4 * cm], repeatRows=1)
    t_m_style = [
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1e2d3d")),
        ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#cbd5e1")),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING", (0, 0), (-1, -1), 2),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
        ("BACKGROUND", (0, -1), (-1, -1), colors.HexColor("#fef3c7")),
        ("SPAN", (2, -1), (6, -1)),
    ]
    for r_i in range(1, len(mem_rows) - 1):
        if r_i % 2 == 0:
            t_m_style.append(("BACKGROUND", (0, r_i), (-1, r_i), colors.HexColor("#f8fafc")))
    t_m.setStyle(TableStyle(t_m_style))
    el.append(t_m)
    el.append(Spacer(1, 0.2 * cm))
    el.append(Paragraph(
        "<b>5.2.</b> Após a realização de pesquisa de preços em conformidade com a IN SEGES/ME nº 65/2021, "
        "certifica-se que o preço estimado para a presente contratação é compatível com os praticados no mercado.",
        st_body,
    ))

    # Seção VI
    sig_block = [
        Paragraph("VI - IDENTIFICAÇÃO DOS AGENTES RESPONSÁVEIS PELA PESQUISA DE PREÇOS", st_sec),
        Paragraph(
            f"<b>6.1.</b> A presente pesquisa de preços foi conduzida por: <b>{_safe_xml(elaborador_nome)}</b>, "
            f"matrícula nº {_safe_xml(elaborador_matricula)}.",
            st_body,
        ),
        Paragraph(f"{_safe_xml(cidade)}, {_data_extenso_hoje()}.", st_body),
        Spacer(1, 0.5 * cm),
        Table(
            [[
                Paragraph(f"________________________________________<br/><b>{_safe_xml(elaborador_nome)}</b><br/>Agente Responsável pela Pesquisa de Preços", st_ctr),
                Paragraph(f"________________________________________<br/><b>{_safe_xml(ordenador_nome)}</b><br/>Autoridade Competente / Ordenador de Despesas", st_ctr),
            ]],
            colWidths=[8.7 * cm, 8.7 * cm],
        ),
        Spacer(1, 0.25 * cm),
        Paragraph(
            "Fonte do Modelo: Guia de Padronização dos Procedimentos de Contratação — AGU / SEGES.",
            st_sub,
        ),
    ]
    el.append(KeepTogether(sig_block))

    doc.build(el)
    return buf.getvalue()


# =============================================================================
# 2. JUSTIFICATIVA CONSOLIDADA (ART. 6º § 5º & JANELA 13-24M) — .DOCX E .PDF
# =============================================================================

def build_consolidated_justificativa_docx_bytes(
    items_results: List[Dict[str, Any]],
    om_name: str = "MARINHA DO BRASIL",
    ordenador_nome: str = "ORDENADOR DE DESPESAS",
    elaborador_nome: str = "ENCARREGADO DO SETOR DE OBTENÇÃO",
) -> bytes:
    """
    Consolida em UM ÚNICO despacho oficial todos os itens que possuam menos de 3 cotações
    válidas (Art. 6º, § 5º da IN 65/2021) e/ou janela excepcional de 13 a 24 meses.
    """
    flagged: List[Tuple[int, Dict[str, Any], List[str]]] = []
    for idx, it in enumerate(items_results, start=1):
        rep = it.get("report") or {}
        reasons = []
        n_valid = len(rep.get("valid_quotes") or [])
        if n_valid < 3:
            lbl_v = "1 cotação válida" if n_valid == 1 else f"{n_valid} cotações válidas"
            reasons.append(f"Apenas {lbl_v} (< 3 — Art. 6º, § 5º IN 65/2021)")
        if is_item_extended_24m(it):
            reasons.append("Janela temporal ampliada (13 a 24 meses — ausência de contratações em 12M)")
        if reasons:
            flagged.append((idx, it, reasons))

    doc = docx.Document()
    for s in doc.sections:
        s.top_margin = Cm(2.0)
        s.bottom_margin = Cm(2.0)
        s.left_margin = Cm(2.0)
        s.right_margin = Cm(2.0)

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(
        f"{om_name.upper()}\n"
        f"DESPACHO DE JUSTIFICATIVA CIRCUNSTANCIADA CONSOLIDADA\n"
        f"(Art. 6º, § 5º da IN SEGES/ME nº 65/2021 e Janela Excepcional 13-24 Meses)"
    )
    r.bold = True
    r.font.size = Pt(11.5)

    p2 = doc.add_paragraph()
    p2.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p2.add_run(
        f"Em observância ao disposto no art. 23 da Lei nº 14.133/2021, no art. 6º, § 5º da IN SEGES/ME nº 65/2021 "
        f"e no Caderno de Logística de Pesquisa de Preços (SEGES/MGI), consolida-se no presente despacho a "
        f"justificativa técnica referente aos {len(flagged)} item(ns) do processo que apresentaram amostra inferior "
        f"a 3 (três) preços válidos e/ou demandaram ampliação da busca para o período de 13 a 24 meses, evitando "
        f"redundância processual nos autos:"
    ).font.size = Pt(10)

    tbl = doc.add_table(rows=1, cols=6)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    _set_table_borders(tbl)
    for ci, h in enumerate(["Item", "Código", "Descrição", "Cotações", "Unit. Est.", "Enquadramento / Justificativa"]):
        cell = tbl.rows[0].cells[ci]
        cell.text = h
        _set_cell_shading(cell, "1E2D3D")
        for pr in cell.paragraphs:
            pr.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for rn in pr.runs:
                rn.bold = True
                rn.font.size = Pt(8)
                rn.font.color.rgb = RGBColor(255, 255, 255)

    for idx, it, reasons in flagged:
        rep = it.get("report") or {}
        rc = tbl.add_row().cells
        rc[0].text = str(idx)
        rc[1].text = str(rep.get("catmat_catser_code") or it.get("code") or "—")
        rc[2].text = str(rep.get("item_description") or it.get("desc") or "Item")[:65]
        rc[3].text = str(len(rep.get("valid_quotes") or []))
        rc[4].text = _fmt_brl(float(rep.get("final_estimated_unit_price") or 0.0))
        rc[5].text = "; ".join(reasons)
        for c in rc:
            for pr in c.paragraphs:
                for rn in pr.runs:
                    rn.font.size = Pt(7.8)

    p_end = doc.add_paragraph()
    p_end.paragraph_format.space_before = Pt(12)
    p_end.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p_end.add_run(
        "Submete-se o presente despacho consolidado à aprovação expressa da Autoridade Competente, "
        "atestando-se a regularidade da pesquisa exaustiva realizada junto aos sistemas oficiais do Governo Federal."
    ).font.size = Pt(10)

    p_sig = doc.add_paragraph()
    p_sig.paragraph_format.space_before = Pt(18)
    p_sig.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_sig.add_run(
        f"Em {_data_extenso_hoje()}.\n\n\n"
        f"________________________________________        ________________________________________\n"
        f"{elaborador_nome}                                {ordenador_nome}\n"
        f"Responsável pela Pesquisa                        Autoridade Competente (Aprovação)"
    ).font.size = Pt(9)

    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()


def build_consolidated_justificativa_pdf_bytes(
    items_results: List[Dict[str, Any]],
    om_name: str = "MARINHA DO BRASIL",
    ordenador_nome: str = "ORDENADOR DE DESPESAS",
    elaborador_nome: str = "ENCARREGADO DO SETOR DE OBTENÇÃO",
) -> bytes:
    """Versão PDF do Despacho de Justificativa Consolidada (Art. 6º § 5º e Janela 13-24M)."""
    flagged = []
    for idx, it in enumerate(items_results, start=1):
        rep = it.get("report") or {}
        reasons = []
        n_valid = len(rep.get("valid_quotes") or [])
        if n_valid < 3:
            lbl_v = "1 cotação válida" if n_valid == 1 else f"{n_valid} cotações válidas"
            reasons.append(f"{lbl_v} (< 3 — Art. 6º § 5º)")
        if is_item_extended_24m(it):
            reasons.append("Janela ampliada 13-24 meses (sem registros em 12M)")
        if reasons:
            flagged.append((idx, it, reasons))

    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4, leftMargin=1.6 * cm, rightMargin=1.6 * cm, topMargin=1.5 * cm, bottomMargin=1.5 * cm)
    styles = getSampleStyleSheet()
    st_t = ParagraphStyle("JTitle", parent=styles["Normal"], fontName="Helvetica-Bold", fontSize=11, leading=14, alignment=1, textColor=colors.HexColor("#1e2d3d"))
    st_b = ParagraphStyle("JBody", parent=styles["Normal"], fontName="Helvetica", fontSize=9, leading=12.5, alignment=4, spaceBefore=6, spaceAfter=6)
    st_h = ParagraphStyle("JHdr", parent=styles["Normal"], fontName="Helvetica-Bold", fontSize=7.2, leading=9, alignment=1, textColor=colors.white)
    st_c = ParagraphStyle("JCell", parent=styles["Normal"], fontName="Helvetica", fontSize=7.0, leading=8.5)
    st_ctr = ParagraphStyle("JCtr", parent=styles["Normal"], fontName="Helvetica", fontSize=7.0, leading=8.5, alignment=1)
    st_rgt = ParagraphStyle("JRgt", parent=styles["Normal"], fontName="Helvetica-Bold", fontSize=7.0, leading=8.5, alignment=2)

    el = [
        Paragraph(_safe_xml(om_name.upper()), st_t),
        Paragraph("DESPACHO DE JUSTIFICATIVA CIRCUNSTANCIADA CONSOLIDADA<br/>Art. 6º, § 5º da IN SEGES/ME nº 65/2021 e Janela Excepcional 13-24 Meses", st_t),
        Paragraph(
            f"Consolida-se neste despacho único a justificativa técnica de {len(flagged)} item(ns) que apresentaram "
            f"menos de 3 (três) preços válidos e/ou necessitaram de ampliação temporal para 13 a 24 meses, "
            f"assegurando economicidade processual e plena rastreabilidade para aprovação da Autoridade Competente.",
            st_b,
        ),
    ]
    rows = [[
        Paragraph("#", st_h), Paragraph("Código", st_h), Paragraph("Descrição", st_h),
        Paragraph("Vál.", st_h), Paragraph("Unit. Est.", st_h), Paragraph("Enquadramento / Justificativa", st_h),
    ]]
    for idx, it, reasons in flagged:
        rep = it.get("report") or {}
        rows.append([
            Paragraph(str(idx), st_ctr),
            Paragraph(_safe_xml(rep.get("catmat_catser_code") or it.get("code") or "—"), st_ctr),
            Paragraph(_safe_xml(rep.get("item_description") or it.get("desc") or "Item", 60), st_c),
            Paragraph(str(len(rep.get("valid_quotes") or [])), st_ctr),
            Paragraph(_fmt_brl(float(rep.get("final_estimated_unit_price") or 0.0)), st_rgt),
            Paragraph(_safe_xml("; ".join(reasons)), st_c),
        ])
    t = Table(rows, colWidths=[0.9 * cm, 1.8 * cm, 5.5 * cm, 1.0 * cm, 2.2 * cm, 6.4 * cm], repeatRows=1)
    t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1e2d3d")),
        ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#cbd5e1")),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING", (0, 0), (-1, -1), 2.5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 2.5),
    ]))
    el.append(t)
    el.append(Spacer(1, 0.5 * cm))
    el.append(KeepTogether([
        Paragraph(f"Em {_data_extenso_hoje()}. Submete-se à aprovação expressa da Autoridade Competente.", st_b),
        Spacer(1, 0.4 * cm),
        Table([[
            Paragraph(f"______________________________________<br/><b>{_safe_xml(elaborador_nome)}</b><br/>Responsável pela Pesquisa", st_ctr),
            Paragraph(f"______________________________________<br/><b>{_safe_xml(ordenador_nome)}</b><br/>Autoridade Competente (Aprovação)", st_ctr),
        ]], colWidths=[8.9 * cm, 8.9 * cm]),
    ]))
    doc.build(el)
    return buf.getvalue()


# =============================================================================
# 3. RELATÓRIO DETALHADO PADRÃO COMPRAS.GOV.BR POR CATEGORIA / JANELA (PDF)
# =============================================================================

def build_category_detailed_report_pdf_bytes(
    indexed_items: List[Tuple[int, Dict[str, Any]]],
    category_label: str,
    window_label: str = "Até 12 Meses (Padrão IN 65/2021)",
    om_name: str = "MARINHA DO BRASIL",
    elaborador_nome: str = "ENCARREGADO DO SETOR DE OBTENÇÃO",
    ordenador_nome: str = "ORDENADOR DE DESPESAS",
) -> bytes:
    """
    Gera o Relatório Detalhado Contínuo padrão Compras.gov.br (inspirado no Relatório
    Detalhado Oficial do Módulo Pesquisa de Preços / SEGES), condensando todos os itens
    de uma mesma Categoria (Classe CNBS/PDM) ou janela temporal em um único PDF auditável,
    com consolidação estatística, listagem completa de TODAS as cotações (válidas + expurgadas),
    links clicáveis para o PNCP e Memória de Cálculo oficial ao final.
    """
    buf = io.BytesIO()
    doc = SimpleDocTemplate(
        buf,
        pagesize=landscape(A4),
        leftMargin=1.0 * cm,
        rightMargin=1.0 * cm,
        topMargin=1.0 * cm,
        bottomMargin=1.0 * cm,
    )
    styles = getSampleStyleSheet()
    st_top = ParagraphStyle("CatTop", parent=styles["Normal"], fontName="Helvetica-Bold", fontSize=11, leading=13.5, alignment=1, textColor=colors.HexColor("#0f172a"))
    st_sub = ParagraphStyle("CatSub", parent=styles["Normal"], fontName="Helvetica", fontSize=8, leading=10, alignment=1, textColor=colors.HexColor("#475569"))
    st_item_hdr = ParagraphStyle("CatItemHdr", parent=styles["Normal"], fontName="Helvetica-Bold", fontSize=8.5, leading=10.5, textColor=colors.white)
    st_meta = ParagraphStyle("CatMeta", parent=styles["Normal"], fontName="Helvetica", fontSize=7.2, leading=9.2, textColor=colors.HexColor("#1e293b"))
    st_th = ParagraphStyle("CatTH", parent=styles["Normal"], fontName="Helvetica-Bold", fontSize=6.6, leading=8, alignment=1, textColor=colors.white)
    st_td = ParagraphStyle("CatTD", parent=styles["Normal"], fontName="Helvetica", fontSize=6.4, leading=7.8, textColor=colors.HexColor("#0f172a"))
    st_td_c = ParagraphStyle("CatTDC", parent=styles["Normal"], fontName="Helvetica", fontSize=6.4, leading=7.8, alignment=1, textColor=colors.HexColor("#0f172a"))
    st_td_r = ParagraphStyle("CatTDR", parent=styles["Normal"], fontName="Helvetica-Bold", fontSize=6.4, leading=7.8, alignment=2, textColor=colors.HexColor("#0f172a"))
    st_formula = ParagraphStyle("CatForm", parent=styles["Normal"], fontName="Helvetica", fontSize=7.5, leading=10, textColor=colors.HexColor("#334155"))

    total_cat = sum(
        float(it.get("qty", 1) or 1) * float((it.get("report") or {}).get("final_estimated_unit_price") or 0.0)
        for _, it in indexed_items
    )

    elements: List[Any] = [
        Paragraph(f"{_safe_xml(om_name.upper())} — RELATÓRIO DETALHADO DE PESQUISA DE PREÇOS", st_top),
        Paragraph(
            f"<b>Categoria / Classe CNBS:</b> {_safe_xml(category_label)} &nbsp;|&nbsp; "
            f"<b>Janela Temporal:</b> {_safe_xml(window_label)} &nbsp;|&nbsp; "
            f"<b>Itens:</b> {len(indexed_items)} &nbsp;|&nbsp; "
            f"<b>Subtotal Estimado:</b> {_fmt_brl(total_cat)}",
            st_sub,
        ),
        Spacer(1, 0.25 * cm),
    ]

    for item_idx, it in indexed_items:
        rep = it.get("report") or {}
        qty = float(it.get("qty", 1) or 1)
        code = str(rep.get("catmat_catser_code") or it.get("code") or "—")
        desc = str(rep.get("item_description") or it.get("desc") or "Item demandado")
        unit = str(rep.get("target_supply_unit") or it.get("unit") or "UN")
        sm = rep.get("sanitized_metrics") or {}
        u_est = float(rep.get("final_estimated_unit_price") or 0.0)
        tot = qty * u_est
        adopted = str(rep.get("adopted_parameter") or rep.get("final_metric_used") or "MEDIANA")
        cv_val = sm.get("cv_percent")
        cv_str = f"{float(cv_val):.2f}%" if cv_val is not None else "0.00%"
        pdm = str(rep.get("nome_pdm") or "—")
        win_tag = "13-24 MESES (EXCEPCIONAL)" if is_item_extended_24m(it) else "ATÉ 12 MESES (IN 65/2021)"

        # Banner do Item
        banner = Table(
            [[Paragraph(f"ITEM {item_idx:03d} — CÓD. {_safe_xml(code)} | {_safe_xml(desc, 130)}", st_item_hdr)]],
            colWidths=[27.7 * cm],
        )
        banner.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#1e2d3d")),
            ("TOPPADDING", (0, 0), (-1, -1), 4),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
            ("LEFTPADDING", (0, 0), (-1, -1), 6),
        ]))

        # Strip de Consolidação Estatística (espelho do Compras.gov.br Relatório Detalhado)
        consol_data = [[
            Paragraph(f"<b>Unid./Qtd:</b> {_safe_xml(unit)} / {qty:g}<br/><b>PDM:</b> {_safe_xml(pdm, 34)}", st_meta),
            Paragraph(f"<b>Menor:</b> {_fmt_brl(sm.get('min_price') or u_est)}<br/><b>Maior:</b> {_fmt_brl(sm.get('max_price') or u_est)}", st_meta),
            Paragraph(f"<b>Média:</b> {_fmt_brl(sm.get('mean') or u_est)}<br/><b>Mediana:</b> {_fmt_brl(sm.get('median') or u_est)}", st_meta),
            Paragraph(f"<b>Desvio Padrão:</b> {_fmt_brl(sm.get('std_deviation') or 0)}<br/><b>CV%:</b> {cv_str}", st_meta),
            Paragraph(f"<b>Método:</b> {_safe_xml(adopted)}<br/><b>Janela:</b> {win_tag}", st_meta),
            Paragraph(f"<b>Unit. Est.:</b> <font color='#15803d'>{_fmt_brl(u_est)}</font><br/><b>Total Est.:</b> <font color='#15803d'>{_fmt_brl(tot)}</font>", st_meta),
        ]]
        t_consol = Table(consol_data, colWidths=[4.8 * cm, 4.1 * cm, 4.3 * cm, 4.3 * cm, 5.2 * cm, 5.0 * cm])
        t_consol.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#f1f5f9")),
            ("BOX", (0, 0), (-1, -1), 0.5, colors.HexColor("#94a3b8")),
            ("INNERGRID", (0, 0), (-1, -1), 0.3, colors.HexColor("#cbd5e1")),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("TOPPADDING", (0, 0), (-1, -1), 3),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
        ]))

        # Tabela de TODAS as cotações (Válidas + Expurgadas) — padrão Compras.gov.br
        q_rows = [[
            Paragraph("#", st_th),
            Paragraph("Status", st_th),
            Paragraph("Fonte / Inciso", st_th),
            Paragraph("Fornecedor / CNPJ", st_th),
            Paragraph("Marca", st_th),
            Paragraph("Órgão / UASG / UF", st_th),
            Paragraph("ID Compra / Link PNCP", st_th),
            Paragraph("Data", st_th),
            Paragraph("Unid / Qtd", st_th),
            Paragraph("Preço Unit.", st_th),
        ]]

        all_quotes = [(True, q) for q in (rep.get("valid_quotes") or [])] + [(False, q) for q in (rep.get("purged_quotes") or [])]
        for q_i, (is_valid, q) in enumerate(all_quotes, start=1):
            st_lbl = "<font color='#15803d'><b>COMPÕE</b></font>" if is_valid else f"<font color='#b91c1c'><b>EXPURGADA</b></font><br/>{_safe_xml(_categorize_purge_reason(q), 22)}"
            src = _safe_xml(str(q.get("source_inciso") or "Inciso I - Compras.gov.br").replace(" - Compras.gov.br", ""), 26)
            forn = f"<b>{_safe_xml(q.get('supplier_name') or '—', 40)}</b><br/>{_safe_xml(q.get('supplier_cnpj') or '—')}"
            marca = _safe_xml(q.get("brand") or "—", 22)
            esf = f" ({q.get('esfera')})" if q.get("esfera") else ""
            org = f"{_safe_xml(q.get('buyer_org_name') or q.get('nome_uasg') or '—', 38)}<br/>UASG {_safe_xml(q.get('buyer_uasg') or '—')} / {_safe_xml(q.get('buyer_uf') or '—')}{_safe_xml(esf)}"
            uri = str(q.get("pncp_or_compras_uri") or "").strip()
            pkey_d = re.sub(r"\D", "", str(q.get("purchase_key") or ""))
            if len(pkey_d) == 16:
                pkey_d = "0" + pkey_d
            compras_uri = str(q.get("comprasnet_uri") or "").strip()
            if not compras_uri and len(pkey_d) >= 17:
                compras_uri = f"https://cnetmobile.estaleiro.serpro.gov.br/comprasnet-web/public/compras/acompanhamento-compra?compra={pkey_d[:17]}"
            if ("?q=" in uri or not uri) and compras_uri:
                uri = compras_uri
            id_txt = _safe_xml(q.get("pncp_control_number") or q.get("purchase_number_year") or q.get("purchase_key") or q.get("id_compra_item") or "—", 32)
            if uri.startswith("http"):
                id_cell = Paragraph(f'<link href="{xml_escape(uri)}" color="#1d4ed8"><u>{id_txt}</u></link>', st_td_c)
            else:
                id_cell = Paragraph(id_txt, st_td_c)
            dt_str = _safe_xml(str(q.get("purchase_date") or "—")[:10])
            cap = f" ({q.get('capacidade_unidade')})" if q.get("capacidade_unidade") else ""
            uq_str = f"{_safe_xml(q.get('supply_unit') or unit)}{_safe_xml(cap)}<br/>Qtd: {float(q.get('quantity_homologated') or 0):g}"
            p_unit = _fmt_brl(float(q.get("unit_price") or 0.0))

            q_rows.append([
                Paragraph(str(q_i), st_td_c),
                Paragraph(st_lbl, st_td_c),
                Paragraph(src, st_td_c),
                Paragraph(forn, st_td),
                Paragraph(marca, st_td_c),
                Paragraph(org, st_td),
                id_cell,
                Paragraph(dt_str, st_td_c),
                Paragraph(uq_str, st_td_c),
                Paragraph(p_unit, st_td_r),
            ])

        if len(q_rows) == 1:
            q_rows.append([Paragraph("—", st_td_c)] * 3 + [Paragraph("Sem cotações registradas.", st_td)] + [Paragraph("—", st_td_c)] * 6)

        t_q = Table(
            q_rows,
            colWidths=[0.8 * cm, 2.2 * cm, 2.4 * cm, 5.2 * cm, 2.1 * cm, 5.1 * cm, 3.2 * cm, 1.9 * cm, 2.3 * cm, 2.5 * cm],
            repeatRows=1,
        )
        t_q_style = [
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#334155")),
            ("GRID", (0, 0), (-1, -1), 0.3, colors.HexColor("#cbd5e1")),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("TOPPADDING", (0, 0), (-1, -1), 2),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
        ]
        for r_i in range(1, len(q_rows)):
            is_val_row = all_quotes[r_i - 1][0] if (r_i - 1) < len(all_quotes) else True
            if not is_val_row:
                t_q_style.append(("BACKGROUND", (0, r_i), (-1, r_i), colors.HexColor("#fef2f2")))
            elif r_i % 2 == 0:
                t_q_style.append(("BACKGROUND", (0, r_i), (-1, r_i), colors.HexColor("#f8fafc")))
        t_q.setStyle(TableStyle(t_q_style))

        elements.append(KeepTogether([banner, t_consol]))
        elements.append(t_q)
        elements.append(Spacer(1, 0.35 * cm))

    # Memória de Cálculo Oficial (Art. 3º, VII da IN 65/2021)
    elements.append(KeepTogether([
        Paragraph("<b>MEMÓRIA DE CÁLCULO (Art. 3º, inciso VII da IN SEGES/ME nº 65/2021)</b>", st_meta),
        Paragraph(
            "• <b>Média Aritmética (μ):</b> Soma de todos os preços válidos dividida pelo número de cotações válidas (n). "
            "Adotada automaticamente quando o Coeficiente de Variação (CV) ≤ 25,00% (série homogênea).<br/>"
            "• <b>Mediana:</b> Valor central da série ordenada de preços válidos (ou média dos dois termos centrais se n for par). "
            "Adotada automaticamente quando CV &gt; 25,00% (série heterogênea).<br/>"
            "• <b>Desvio Padrão Amostral (s) e Coeficiente de Variação (CV):</b> s = √[ Σ(xi - μ)² / (n - 1) ] e CV = (s / μ) × 100.<br/>"
            "• <b>Saneamento de Outliers:</b> Duplo estágio — Barreira Robusta por Mediana (inexequível &lt; 25% ou excessivo &gt; 400% da mediana bruta) "
            "seguido de corte por Desvio Padrão (μ ± 1·s) quando CV &gt; 25%.<br/>"
            f"<i>Relatório gerado eletronicamente em {_data_extenso_hoje()} — Responsável: {_safe_xml(elaborador_nome)} | Autoridade: {_safe_xml(ordenador_nome)}.</i>",
            st_formula,
        ),
    ]))

    doc.build(elements)
    return buf.getvalue()


# =============================================================================
# 4. RAIO-X CONSOLIDADO DE CERTIDÕES DAS EMPRESAS DO LOTE EM WORD (.DOCX)
# =============================================================================

def build_consolidated_suppliers_dossier_docx_bytes(
    items_results: List[Dict[str, Any]],
    om_name: str = "MARINHA DO BRASIL",
    ordenador_nome: str = "ORDENADOR DE DESPESAS",
    elaborador_nome: str = "ENCARREGADO DO SETOR DE OBTENÇÃO",
) -> bytes:
    """
    Gera o Raio-X Consolidado de Certidões de TODOS os fornecedores únicos (CNPJ)
    presentes nas cotações do lote em formato Word (.DOCX) Paisagem A4 editável.
    Fundamento: Art. 68 e Art. 156 da Lei nº 14.133/2021.
    """
    try:
        from sanctions_checker import check_supplier_full_compliance
    except ImportError:
        from core.sanctions_checker import check_supplier_full_compliance

    unique_suppliers: Dict[str, Dict[str, Any]] = {}
    for idx, item in enumerate(items_results, start=1):
        rep = item.get("report") or {}
        for q in (rep.get("valid_quotes") or []):
            cnpj_raw = str(q.get("supplier_cnpj") or q.get("cnpj") or "").strip()
            clean = re.sub(r"\D", "", cnpj_raw)
            if len(clean) == 14 and clean not in unique_suppliers:
                unique_suppliers[clean] = {
                    "cnpj_raw": cnpj_raw,
                    "name": str(q.get("supplier_name") or q.get("fornecedor") or "Fornecedor Cotado").strip(),
                    "items": [idx],
                }
            elif len(clean) == 14:
                if len(unique_suppliers[clean]["items"]) < 10 and idx not in unique_suppliers[clean]["items"]:
                    unique_suppliers[clean]["items"].append(idx)

    doc = docx.Document()
    for s in doc.sections:
        s.orientation = docx.enum.section.WD_ORIENT.LANDSCAPE
        s.page_width = Inches(11.69)
        s.page_height = Inches(8.27)
        s.top_margin = Cm(1.3)
        s.bottom_margin = Cm(1.3)
        s.left_margin = Cm(1.3)
        s.right_margin = Cm(1.3)

    p_hdr = doc.add_paragraph()
    p_hdr.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_hdr.paragraph_format.space_after = Pt(4)
    r1 = p_hdr.add_run(f"{om_name.upper()}\n")
    r1.bold = True
    r1.font.size = Pt(11.5)
    r2 = p_hdr.add_run("RAIO-X CONSOLIDADO DE CERTIDÕES E REGULARIDADE DAS EMPRESAS COTADAS\n")
    r2.bold = True
    r2.font.size = Pt(12.5)
    r2.font.color.rgb = RGBColor(30, 45, 61)
    r3 = p_hdr.add_run(
        f"Art. 68 e Art. 156 da Lei nº 14.133/2021 • {len(unique_suppliers)} empresa(s) única(s) auditada(s) em {len(items_results)} item(ns) • Emitido em {_data_extenso_hoje()}"
    )
    r3.font.size = Pt(8.5)
    r3.font.color.rgb = RGBColor(71, 85, 105)

    p_note = doc.add_paragraph()
    p_note.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p_note.paragraph_format.space_after = Pt(6)
    rn = p_note.add_run(
        "ORIENTAÇÃO NORMATIVA: A pesquisa de preços (Art. 23 da Lei nº 14.133/2021 e IN SEGES/ME nº 65/2021) não sofre "
        "impedimento legal na formação da cesta de referência caso alguma empresa cotada apresente certidão vencida; "
        "contudo, o presente Raio-X acompanha os autos para transparência plena e controle prévio de habilitação (Art. 68) "
        "e sanções impeditivas (CEIS/CNEP/TCU — Art. 156) na fase de contratação."
    )
    rn.font.size = Pt(8.5)
    rn.italic = True

    tbl = doc.add_table(rows=1, cols=10)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    _set_table_borders(tbl)
    hdrs = ["#", "CNPJ", "RAZÃO SOCIAL / FORNECEDOR", "ITENS", "RFB / PGFN", "FGTS (CRF)", "CNDT (TST)", "CEIS / CNEP / TCU", "SICAF", "SITUAÇÃO"]
    for ci, h in enumerate(hdrs):
        cell = tbl.rows[0].cells[ci]
        cell.text = h
        _set_cell_shading(cell, "1E2D3D")
        for p in cell.paragraphs:
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for r in p.runs:
                r.bold = True
                r.font.size = Pt(7.8)
                r.font.color.rgb = RGBColor(255, 255, 255)

    reg_count = 0
    irreg_count = 0
    for s_idx, (clean_cnpj, info) in enumerate(list(unique_suppliers.items())[:500], start=1):
        comp = check_supplier_full_compliance(clean_cnpj, info["name"])
        cert_map = {c.get("id"): c for c in (comp.get("dossie_certidoes") or [])}
        sanctions_ok = "NADA CONSTA" if not comp.get("has_sanctions") else "IMPEDIDO"
        is_ok = bool(comp.get("is_eligible"))
        if is_ok:
            reg_count += 1
        else:
            irreg_count += 1
        sit = "REGULAR" if is_ok else "ATENÇÃO / IRREGULAR"

        def _fmt_cert(cid: str) -> str:
            cobj = cert_map.get(cid) or {}
            st = str(cobj.get("status") or "REGULAR")
            val = str(cobj.get("data_validade") or "")
            return f"{st}\n(val. {val})" if val and st == "REGULAR" else st

        rc = tbl.add_row().cells
        vals = [
            str(s_idx),
            comp.get("cnpj") or clean_cnpj,
            info["name"][:58],
            ", ".join(str(x) for x in info["items"]),
            _fmt_cert("receita_federal"),
            _fmt_cert("fgts"),
            _fmt_cert("cndt"),
            sanctions_ok,
            str((cert_map.get("sicaf") or {}).get("status") or "HABILITADO").replace("_", " "),
            sit,
        ]
        for ci, v in enumerate(vals):
            rc[ci].text = v
            if not is_ok:
                _set_cell_shading(rc[ci], "FEE2E2")
            elif s_idx % 2 == 0:
                _set_cell_shading(rc[ci], "F8FAFC")
            for p in rc[ci].paragraphs:
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT if ci == 2 else WD_ALIGN_PARAGRAPH.CENTER
                for r in p.runs:
                    r.font.size = Pt(7.3)
                    if ci == 9:
                        r.bold = True

    if not unique_suppliers:
        rc = tbl.add_row().cells
        rc[2].text = "Nenhum CNPJ identificado nas cotações válidas do lote."

    p_sum = doc.add_paragraph()
    p_sum.paragraph_format.space_before = Pt(8)
    rs = p_sum.add_run(
        f"Síntese do Raio-X: {len(unique_suppliers)} empresa(s) única(s) auditada(s) — "
        f"{reg_count} REGULAR(ES) | {irreg_count} com ALERTA/PENDÊNCIA.\n"
        f"Responsável pela Pesquisa: {elaborador_nome} | Autoridade Competente: {ordenador_nome}."
    )
    rs.bold = True
    rs.font.size = Pt(8.5)

    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()


# =============================================================================
# 5. MANUAL DIDÁTICO DO PROCESSO (ESTILO MANUAL OFICIAL — PDF E TXT)
# =============================================================================

def build_process_manual_txt(
    items_results: List[Dict[str, Any]],
    om_name: str = "MARINHA DO BRASIL",
    n_12m: int = 0,
    n_24m: int = 0,
    n_cats_12m: int = 0,
    n_cats_24m: int = 0,
    n_unique_suppliers: int = 0,
) -> str:
    """Versão texto estruturada em formato de Manual Didático para abertura rápida."""
    stats = _compute_batch_stats(items_results)
    total_extenso = brl_por_extenso(stats["total_geral"])
    lines = [
        "╔══════════════════════════════════════════════════════════════════════════════╗",
        "║         MANUAL DIDÁTICO DO PROCESSO DE PESQUISA DE PREÇOS (IN 65/2021)       ║",
        "║      Lei Federal nº 14.133/2021 • IN SEGES/ME nº 65/2021 • Modelo AGU       ║",
        "╚══════════════════════════════════════════════════════════════════════════════╝",
        "",
        f"  Órgão / OM ...............: {om_name.upper()}",
        f"  Data de Emissão ..........: {_data_extenso_hoje()}",
        f"  Total de Itens ...........: {stats['total_items']} item(ns)",
        f"  Valor Global Estimado ....: {_fmt_brl(stats['total_geral'])} ({total_extenso})",
        f"  Janela Padrão (≤ 12M) ....: {n_12m} item(ns) em {n_cats_12m} categoria(s) CNBS",
        f"  Janela Excepcional (13-24M): {n_24m} item(ns) em {n_cats_24m} categoria(s) CNBS",
        f"  Itens com < 3 cotações ...: {len(stats['items_under_3'])} item(ns) (Art. 6º § 5º)",
        f"  Empresas Únicas Auditadas : {n_unique_suppliers} CNPJ(s)",
        "",
        "────────────────────────────────────────────────────────────────────────────────",
        "  PARTE 1 — COMO ESTÁ ORGANIZADO ESTE PACOTE (.ZIP)",
        "────────────────────────────────────────────────────────────────────────────────",
        "",
        "  [RAIZ DO PACOTE — PEÇAS MESTRES PARA JUNTADA DIRETA NO SEI]",
        "    • 00_MANUAL_DIDATICO_DO_PROCESSO_LEIA_ME.pdf",
        "      Versão visual completa deste guia passo a passo para a equipe de licitação.",
        "",
        "    • 00_Mapa_Comparativo_Geral_Completo.xlsx",
        "      Planilha Mestre em Excel com 4 abas integradas:",
        "        1) Resumo Geral do Lote (totais, médias, medianas, CV% e janela temporal);",
        "        2) Mapa Geral - Todas Cotações (100% das empresas e cotações de todos os itens,",
        "           com Id contratação PNCP e links diretos clicáveis PNCP + Compras.gov.br);",
        "        3) Expurgos e Inconsistências (espelho da Seção 4.2 da Nota Técnica AGU);",
        "        4) Resumo por Categoria (totalização por Classe CNBS/PDM).",
        "",
        "    • 01_Nota_Tecnica_AGU_IN65_Pesquisa_Precos (.docx e .pdf)",
        "      Nota Técnica fiel ao modelo oficial da AGU/SEGES (Seções I a VI, valor por",
        "      extenso, quadro de consulta direta, tabela de expurgos e assinaturas).",
        "",
        "    • 02_RaioX_Certidoes_Empresas_Consolidado (.docx e .pdf)",
        "      Raio-X de regularidade fiscal, trabalhista e sancionatória (RFB/PGFN, FGTS,",
        "      CNDT, CEIS/CNEP/TCU e SICAF — Art. 68 e 156 NLLC) de todas as empresas do lote.",
        "",
        "    • 03_Justificativa_Consolidada_Art6_Par5 (.docx e .pdf)",
        "      Despacho único consolidando justificativa para itens com menos de 3 cotações",
        "      válidas (Art. 6º, § 5º) e/ou que utilizaram a janela excepcional de 13 a 24 meses.",
        "",
        "    • 04_Comprovantes_PNCP_Contratacoes_QRCode.pdf",
        "      Prova documental oficial (1 página por contratação única, espelho do portal PNCP",
        "      com QR Code, Id contratação PNCP, valores estimado/homologado e links diretos).",
        "",
        "  [PASTAS SEGREGADAS POR JANELA TEMPORAL — ZERO REDUNDÂNCIA]",
        f"    • 01_Pesquisas_Ate_12_Meses_IN65/ ({n_12m} item(ns))",
        f"    • 02_Pesquisas_13_a_24_Meses_Excepcional/ ({n_24m} item(ns))",
        "      Dentro de cada pasta você encontrará:",
        "        ├─ Caderno_Resumo_Executivo (.pdf e .docx) -> Quadro geral daquela janela;",
        "        ├─ Relatorio_Detalhado_Completo (.pdf)     -> Modelo Compras.gov.br contínuo;",
        "        ├─ Por_Categoria/                          -> 1 PDF detalhado por Classe CNBS;",
        "        └─ Relatorios_CATMAT_CATSER/               -> Relatório oficial de 2 páginas",
        "           (Brasão, QR Codes individuais, Boxplot e Curva Gaussiana) de CADA ITEM!",
        "",
        "────────────────────────────────────────────────────────────────────────────────",
        "  PARTE 2 — PASSO A PASSO DIDÁTICO PARA MONTAGEM DO PROCESSO NO SEI",
        "────────────────────────────────────────────────────────────────────────────────",
        "",
        "  PASSO 1 — Conferir e numerar a Nota Técnica AGU:",
        "            Abra '01_Nota_Tecnica_AGU_IN65_Pesquisa_Precos.docx', informe o nº da",
        "            Nota Técnica, matrícula do agente e ajuste o objeto se necessário.",
        "            Depois junte o PDF oficial assinado no SEI.",
        "",
        "  PASSO 2 — Juntar o Mapa Comparativo Geral:",
        "            Anexe '00_Mapa_Comparativo_Geral_Completo.xlsx' como planilha de",
        "            memória de cálculo (a aba 'Mapa Geral - Todas Cotações' prova todas",
        "            as empresas consultadas sem necessidade de múltiplas planilhas).",
        "",
        "  PASSO 3 — Juntar os Relatórios CATMAT/CATSER e por Categoria (≤ 12 Meses):",
        "            Utilize os relatórios da pasta '01_Pesquisas_Ate_12_Meses_IN65/'.",
        "            Cada item possui seu Relatório CATMAT/CATSER com Boxplot, Curva",
        "            Gaussiana e QR Codes individuais na subpasta 'Relatorios_CATMAT_CATSER/'.",
        "",
        "  PASSO 4 — Tratar separadamente os itens de 13 a 24 Meses (se houver):",
        "            Os itens da pasta '02_Pesquisas_13_a_24_Meses_Excepcional/' não tiveram",
        "            compras nos últimos 12 meses. Avalie a conveniência de atualização",
        "            monetária pelo IPCA/IBGE e colha anuência na Justificativa Consolidada.",
        "",
        "  PASSO 5 — Juntar Comprovantes PNCP e Raio-X das Empresas:",
        "            Anexe '04_Comprovantes_PNCP_Contratacoes_QRCode.pdf' (prova cabal de",
        "            que as compras existem no PNCP) e '02_RaioX_Certidoes_Empresas_Consolidado'.",
        "",
        "  PASSO 6 — Colher Assinaturas:",
        "            Submeta a Nota Técnica AGU e a Justificativa Consolidada (Art. 6º § 5º)",
        "            à ratificação da Autoridade Competente / Ordenador de Despesas.",
        "",
        "────────────────────────────────────────────────────────────────────────────────",
        "  PARTE 3 — CHECKLIST RÁPIDO DE CONFORMIDADE (TCU / CGU / IN 65/2021)",
        "────────────────────────────────────────────────────────────────────────────────",
        "  [ ] Nota Técnica AGU conferida, numerada e assinada (Seções I a VI).",
        "  [ ] Mapa Comparativo Geral (.xlsx) anexado com todas as empresas e expurgos.",
        "  [ ] Relatórios CATMAT/CATSER (com gráficos e QR Codes) verificados.",
        "  [ ] Comprovantes PNCP (com QR Code e Id contratação PNCP) juntados aos autos.",
        "  [ ] Raio-X de Certidões das Empresas (Art. 68 NLLC) anexado para ciência.",
        "  [ ] Justificativa Art. 6º § 5º / Janela 13-24M aprovada pela Autoridade (se aplicável).",
        "════════════════════════════════════════════════════════════════════════════════",
    ]
    return "\n".join(lines)


def build_process_manual_pdf_bytes(
    items_results: List[Dict[str, Any]],
    om_name: str = "MARINHA DO BRASIL",
    ordenador_nome: str = "ORDENADOR DE DESPESAS",
    elaborador_nome: str = "ENCARREGADO DO SETOR DE OBTENÇÃO",
    n_12m: int = 0,
    n_24m: int = 0,
    n_cats_12m: int = 0,
    n_cats_24m: int = 0,
    n_unique_suppliers: int = 0,
) -> bytes:
    """
    Gera o Manual Didático Visual do Processo (`00_MANUAL_DIDATICO_DO_PROCESSO_LEIA_ME.pdf`)
    em estilo guia institucional prático, passo a passo para instrução no SEI.
    """
    stats = _compute_batch_stats(items_results)
    total_extenso = brl_por_extenso(stats["total_geral"])

    buf = io.BytesIO()
    doc = SimpleDocTemplate(
        buf,
        pagesize=A4,
        leftMargin=1.6 * cm,
        rightMargin=1.6 * cm,
        topMargin=1.4 * cm,
        bottomMargin=1.4 * cm,
    )
    styles = getSampleStyleSheet()
    st_tag = ParagraphStyle("ManTag", parent=styles["Normal"], fontName="Helvetica-Bold", fontSize=8, leading=10, textColor=colors.HexColor("#d8be86"))
    st_hero = ParagraphStyle("ManHero", parent=styles["Normal"], fontName="Helvetica-Bold", fontSize=14.5, leading=18, textColor=colors.white)
    st_hsub = ParagraphStyle("ManHSub", parent=styles["Normal"], fontName="Helvetica", fontSize=8.5, leading=11.5, textColor=colors.HexColor("#cbd5e1"))
    st_sec = ParagraphStyle("ManSec", parent=styles["Normal"], fontName="Helvetica-Bold", fontSize=10.5, leading=13.5, textColor=colors.HexColor("#1e293b"), spaceBefore=10, spaceAfter=4, keepWithNext=True)
    st_body = ParagraphStyle("ManBody", parent=styles["Normal"], fontName="Helvetica", fontSize=8.6, leading=12.2, alignment=4, textColor=colors.HexColor("#334155"), spaceAfter=4)
    st_card_lbl = ParagraphStyle("ManCLbl", parent=styles["Normal"], fontName="Helvetica-Bold", fontSize=7.0, leading=9, alignment=1, textColor=colors.HexColor("#475569"))
    st_card_val = ParagraphStyle("ManCVal", parent=styles["Normal"], fontName="Helvetica-Bold", fontSize=11, leading=13.5, alignment=1, textColor=colors.HexColor("#0f172a"))
    st_th = ParagraphStyle("ManTH", parent=styles["Normal"], fontName="Helvetica-Bold", fontSize=7.5, leading=9.5, alignment=1, textColor=colors.white)
    st_td = ParagraphStyle("ManTD", parent=styles["Normal"], fontName="Helvetica", fontSize=7.6, leading=10.0, textColor=colors.HexColor("#1e293b"))
    st_step_num = ParagraphStyle("ManSNum", parent=styles["Normal"], fontName="Helvetica-Bold", fontSize=9.5, leading=11.5, alignment=1, textColor=colors.white)
    st_step_txt = ParagraphStyle("ManSTxt", parent=styles["Normal"], fontName="Helvetica", fontSize=8.2, leading=11.2, textColor=colors.HexColor("#1e293b"))

    el: List[Any] = []

    # Capa / Banner estilo Manual Oficial
    hero_tbl = Table(
        [[
            Paragraph(
                f"<font color='#d8be86'><b>MANUAL DIDÁTICO DO PROCESSO • GUIA DE INSTRUÇÃO SEI</b></font><br/>"
                f"<font size='14' color='#ffffff'><b>DOSSIÊ CONSOLIDADO DE PESQUISA DE PREÇOS</b></font><br/>"
                f"<font size='8.5' color='#cbd5e1'>{_safe_xml(om_name.upper())} &nbsp;|&nbsp; Lei nº 14.133/2021, IN SEGES/ME nº 65/2021 e Modelo AGU &nbsp;|&nbsp; Emitido em {_data_extenso_hoje()}</font>",
                st_hsub,
            )
        ]],
        colWidths=[17.8 * cm],
    )
    hero_tbl.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#1e293b")),
        ("TOPPADDING", (0, 0), (-1, -1), 12),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 12),
        ("LEFTPADDING", (0, 0), (-1, -1), 12),
        ("RIGHTPADDING", (0, 0), (-1, -1), 12),
    ]))
    el.append(hero_tbl)
    el.append(Spacer(1, 0.3 * cm))

    # Seção 1: Painel Executivo do Lote (4 Cards)
    cards = [[
        Paragraph(f"TOTAL DE ITENS<br/><font size='12' color='#0f172a'><b>{stats['total_items']}</b></font><br/>{n_cats_12m + n_cats_24m} categoria(s)", st_card_lbl),
        Paragraph(f"VALOR GLOBAL ESTIMADO<br/><font size='10.5' color='#15803d'><b>{_fmt_brl(stats['total_geral'])}</b></font><br/>Referência IN 65/2021", st_card_lbl),
        Paragraph(f"JANELA TEMPORAL<br/><font size='10' color='#0f172a'><b>{n_12m} (≤12M) | {n_24m} (13-24M)</b></font><br/>Pastas segregadas", st_card_lbl),
        Paragraph(f"AUDITORIA & EMPRESAS<br/><font size='10' color='#0f172a'><b>{n_unique_suppliers} CNPJ(s)</b></font><br/>{len(stats['items_under_3'])} item(ns) &lt; 3 cot.", st_card_lbl),
    ]]
    t_cards = Table(cards, colWidths=[4.1 * cm, 4.9 * cm, 4.5 * cm, 4.3 * cm])
    t_cards.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#f8fafc")),
        ("BOX", (0, 0), (-1, -1), 0.7, colors.HexColor("#cbd5e1")),
        ("INNERGRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#cbd5e1")),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING", (0, 0), (-1, -1), 6),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
    ]))
    el.append(t_cards)
    el.append(Spacer(1, 0.15 * cm))
    el.append(Paragraph(
        f"<b>Valor Global por Extenso:</b> {_safe_xml(total_extenso)}. &nbsp; "
        f"<b>Responsável Técnico:</b> {_safe_xml(elaborador_nome)} &nbsp;|&nbsp; <b>Autoridade:</b> {_safe_xml(ordenador_nome)}.",
        st_body,
    ))

    # Seção 2: Mapa Visual de Pastas e Arquivos
    el.append(Paragraph("1. MAPA DIDÁTICO DOS ARQUIVOS E PASTAS DESTE PACOTE (.ZIP)", st_sec))
    map_rows = [
        [Paragraph("Arquivo / Pasta", st_th), Paragraph("O que contém e para que serve no processo", st_th), Paragraph("Base Legal", st_th)],
        [
            Paragraph("<b>00_Mapa_Comparativo_Geral_Completo.xlsx</b>", st_td),
            Paragraph("Planilha Excel completa com 4 abas: <i>Resumo Geral</i>, <b>Mapa Geral - Todas Cotações</b> (100% das empresas de todos os itens com Id PNCP e links diretos), <i>Expurgos</i> e <i>Resumo por Categoria</i>.", st_td),
            Paragraph("Art. 23 NLLC<br/>IN 65/2021", st_td),
        ],
        [
            Paragraph("<b>01_Nota_Tecnica_AGU_IN65_Pesquisa_Precos</b><br/>(.docx e .pdf)", st_td),
            Paragraph("Nota Técnica fiel ao modelo oficial AGU/SEGES (Seções I a VI), contendo fontes consultadas, tabela de preços desconsiderados (§ 4.2), valor por extenso (§ 5.1) e assinaturas (§ VI).", st_td),
            Paragraph("Modelo AGU<br/>IN 65/2021", st_td),
        ],
        [
            Paragraph("<b>02_RaioX_Certidoes_Empresas_Consolidado</b><br/>(.docx e .pdf)", st_td),
            Paragraph("Raio-X único de regularidade fiscal, trabalhista e sancionatória (RFB/PGFN, FGTS, CNDT, CEIS/CNEP/TCU e SICAF) de todas as empresas cotadas no lote.", st_td),
            Paragraph("Art. 68 e 156<br/>Lei 14.133/21", st_td),
        ],
        [
            Paragraph("<b>03_Justificativa_Consolidada_Art6_Par5</b><br/>(.docx e .pdf)", st_td),
            Paragraph("Despacho único que justifica todos os itens com menos de 3 cotações válidas e/ou pesquisa ampliada (13–24 meses), evitando dezenas de despachos repetitivos.", st_td),
            Paragraph("Art. 6º, § 5º<br/>IN 65/2021", st_td),
        ],
        [
            Paragraph("<b>04_Comprovantes_PNCP_Contratacoes_QRCode.pdf</b>", st_td),
            Paragraph("Prova documental das compras públicas utilizadas: 1 página estilo portal PNCP por contratação única, com <b>QR Code</b>, <b>Id contratação PNCP</b>, valores estimado/homologado e links diretos.", st_td),
            Paragraph("Art. 174 NLLC<br/>PNCP Oficial", st_td),
        ],
        [
            Paragraph("<b>01_Pesquisas_Ate_12_Meses_IN65/</b><br/><i>(e 02_Pesquisas_13_a_24_Meses_Excepcional/)</i>", st_td),
            Paragraph("Pastas separadas por janela temporal contendo: <b>Caderno Executivo</b> (.pdf/.docx), <b>Relatório Detalhado Completo</b> (padrão Compras.gov.br), <b>Por_Categoria/</b> (agrupado por Classe CNBS) e <b>Relatorios_CATMAT_CATSER/</b> (relatório oficial de 2 páginas com Brasão, QR Codes, Boxplot e Curva Gaussiana de cada item!).", st_td),
            Paragraph("Art. 5º e 6º<br/>IN 65/2021", st_td),
        ],
    ]
    t_map = Table(map_rows, colWidths=[5.4 * cm, 9.8 * cm, 2.6 * cm], repeatRows=1)
    t_map_style = [
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#233242")),
        ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#cbd5e1")),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("TOPPADDING", (0, 0), (-1, -1), 3.5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3.5),
    ]
    for r_i in range(1, len(map_rows)):
        if r_i % 2 == 0:
            t_map_style.append(("BACKGROUND", (0, r_i), (-1, r_i), colors.HexColor("#f8fafc")))
    t_map.setStyle(TableStyle(t_map_style))
    el.append(t_map)

    # Seção 3: Passo a Passo Didático para Instrução no SEI
    el.append(Paragraph("2. PASSO A PASSO PRÁTICO PARA INSTRUIR O PROCESSO NO SEI", st_sec))
    steps_data = [
        ("PASSO 1", "<b>Revisar e numerar a Nota Técnica AGU:</b> Abra o arquivo <i>01_Nota_Tecnica_AGU_IN65_Pesquisa_Precos.docx</i>, preencha o número da Nota Técnica e a matrícula do agente na Seção VI (se ainda não preenchidos). Converta/junte o PDF no SEI para assinatura."),
        ("PASSO 2", "<b>Juntar o Mapa Comparativo Geral (.XLSX):</b> Anexe <i>00_Mapa_Comparativo_Geral_Completo.xlsx</i>. A aba <i>Mapa Geral - Todas Cotações</i> reúne 100% das empresas e links diretos de todos os itens em um único arquivo auditável."),
        ("PASSO 3", "<b>Juntar os Relatórios CATMAT/CATSER e por Categoria (≤ 12 Meses):</b> Na pasta <i>01_Pesquisas_Ate_12_Meses_IN65/</i>, utilize o <i>Relatorio_Detalhado_Completo_12M.pdf</i> (ou os arquivos em <i>Por_Categoria/</i>) e os relatórios individuais de 2 páginas em <i>Relatorios_CATMAT_CATSER/</i>."),
        ("PASSO 4", "<b>Tratar a pasta de 13 a 24 Meses (quando existente):</b> Os itens da pasta <i>02_Pesquisas_13_a_24_Meses_Excepcional/</i> não possuíam compras homologadas nos últimos 12 meses. Avalie a aplicação de atualização monetária (IPCA/IBGE) e mantenha a justificativa registrada."),
        ("PASSO 5", "<b>Anexar Comprovantes PNCP e Raio-X das Empresas:</b> Junte <i>04_Comprovantes_PNCP_Contratacoes_QRCode.pdf</i> (prova cabal da existência de cada compra no PNCP) e <i>02_RaioX_Certidoes_Empresas_Consolidado.pdf</i> (controle prévio de regularidade fiscal/sancionatória)."),
        ("PASSO 6", "<b>Aprovação da Autoridade Competente:</b> Caso existam itens com menos de 3 cotações ou janela de 13–24 meses, submeta <i>03_Justificativa_Consolidada_Art6_Par5</i> junto com a Nota Técnica AGU à ratificação do Ordenador de Despesas."),
    ]
    step_rows = []
    for tag, desc_html in steps_data:
        step_rows.append([Paragraph(tag, st_step_num), Paragraph(desc_html, st_step_txt)])
    t_steps = Table(step_rows, colWidths=[2.2 * cm, 15.6 * cm])
    t_steps.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (0, -1), colors.HexColor("#34495e")),
        ("BACKGROUND", (1, 0), (1, -1), colors.HexColor("#f8fafc")),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#cbd5e1")),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING", (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
        ("LEFTPADDING", (0, 0), (-1, -1), 6),
        ("RIGHTPADDING", (0, 0), (-1, -1), 6),
    ]))
    el.append(t_steps)

    # Seção 4: Glossário Estatístico + Checklist Final
    el.append(KeepTogether([
        Paragraph("3. GLOSSÁRIO RÁPIDO IN 65/2021 &amp; CHECKLIST FINAL DE CONFERÊNCIA", st_sec),
        Paragraph(
            "• <b>Série Homogênea (CV ≤ 25%):</b> baixa dispersão entre os preços válidos → adota-se a <b>MÉDIA ARITMÉTICA</b>.<br/>"
            "• <b>Série Heterogênea (CV &gt; 25%):</b> dispersão elevada → adota-se a <b>MEDIANA</b> (protege contra distorções).<br/>"
            "• <b>Saneamento de Outliers:</b> expurgo automático em duplo estágio (barreira robusta pela mediana + corte μ ± 1σ), com registro nominal na Seção 4.2 da Nota Técnica AGU e na aba <i>Expurgos</i> do Excel.<br/>"
            "• <b>Certidões na Pesquisa de Preços:</b> não há vedação legal para cotar preços de empresas com certidão vencida na formação da cesta; a regularidade fiscal/trabalhista (Art. 68) é exigida na fase de habilitação/contratação, razão pela qual o Raio-X acompanha os autos como ferramenta gerencial preventiva.",
            st_body,
        ),
        Table(
            [[Paragraph(
                "<b>CHECKLIST FINAL ANTES DE ENCAMINHAR O PROCESSO:</b><br/>"
                "☐ Nota Técnica AGU numerada e assinada &nbsp;&nbsp; ☐ Mapa Geral Excel (.xlsx) juntado &nbsp;&nbsp; "
                "☐ Relatórios CATMAT/CATSER conferidos<br/>"
                "☐ Comprovantes PNCP (QR Code) juntados &nbsp;&nbsp; ☐ Raio-X de Certidões anexado &nbsp;&nbsp; "
                "☐ Justificativa Art. 6º § 5º / 13-24M ratificada",
                st_td,
            )]],
            colWidths=[17.8 * cm],
            style=TableStyle([
                ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#fef3c7")),
                ("BOX", (0, 0), (-1, -1), 0.8, colors.HexColor("#d97706")),
                ("TOPPADDING", (0, 0), (-1, -1), 6),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
                ("LEFTPADDING", (0, 0), (-1, -1), 8),
                ("RIGHTPADDING", (0, 0), (-1, -1), 8),
            ]),
        ),
    ]))

    doc.build(el)
    return buf.getvalue()

