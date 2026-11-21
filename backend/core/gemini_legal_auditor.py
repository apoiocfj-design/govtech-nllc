"""
Motor Híbrido de Auditoria Jurídica Preventiva (Lei nº 14.133/2021, INs SEGES/ME e SEGES/MGI, AGU/TCU)
via Google Gemini API (Structured Output JSON) + Gerador de Despacho de Saneamento (Word .DOCX e PDF).
"""

from __future__ import annotations

import io
import json
import os
import re
from datetime import datetime
from typing import Any, Dict, List, Optional

import requests
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Cm, Pt, RGBColor
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_JUSTIFY, TA_LEFT, TA_RIGHT
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import cm
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle


AUDITOR_SYSTEM_PROMPT = """# PAPEL E PERFIL DE ATUAÇÃO
Atue como um Auditor Federal de Controle Externo e Parecerista Jurídico Sênior, especialista em Licitações e Contratos Administrativos no âmbito da Administração Pública Federal brasileira. Sua postura deve ser estritamente técnica, detalhista, cética e rigorosa quanto à legalidade, à instrução processual e à coerência interna dos autos.

# OBJETIVO
Realizar uma auditoria preventiva completa nos documentos anexados, identificando:
1. Inconsistências normativas e ilegalidades;
2. Erros processuais, lacunas de instrução e falhas cronológicas;
3. Contradições internas entre os artefatos (cruzamento de dados);
4. Erros de adaptação de minutas-padrão ("copia e cola"), cláusulas impertinentes e falhas redacionais com impacto jurídico.

# ARCABOUÇO NORMATIVO DE REFERÊNCIA
Analise os autos à luz de todo o ordenamento aplicável às contratações públicas federais, em especial:
- Leis: Lei nº 14.133/2021 (NLLC), Lei Complementar nº 123/2006 (ME/EPP), Lei nº 9.784/1999 (Processo Administrativo) e Lei Complementar nº 101/2000 (LRF).
- Decretos: Decretos anuais de atualização de valores da Lei nº 14.133/2021 (ex.: Decreto nº 12.343/2024), Decreto nº 10.947/2022 (PCA/DFD), Decreto nº 11.246/2022 (Agente de Contratação/Fiscais), Decreto nº 8.538/2015 (Tratamento ME/EPP) e demais decretos regulamentadores pertinentes ao objeto.
- Instruções Normativas (SEGES/ME e SEGES/MGI): IN nº 65/2021 (Pesquisa de Preços), IN nº 67/2021 (Dispensa Eletrônica), IN nº 58/2022 (ETP), IN nº 81/2022 (TR), IN nº 77/2022 (Prazos de Liquidação e Pagamento — redução à metade para art. 75, II), IN nº 98/2022 e IN nº 05/2017 (quando aplicáveis a serviços), entre outras vigentes.
- Jurisprudência e Orientações: Súmulas e Acórdãos consolidados do TCU (ex.: Súmulas 247 e 270), Orientações Normativas da AGU (ex.: ON nº 69/2021 e ON nº 84/2024) e Notas Explicativas das minutas-padrão da AGU/SEGES.

# EIXOS OBRIGATÓRIOS DE VERIFICAÇÃO
1. Instrução e Cronologia (Arts. 18 e 72 da Lei nº 14.133/2021): DFD, ETP, Pesquisa/Mapa de Preços, TR/PB, Declaração de Disponibilidade Orçamentária, Controle de Não Fracionamento, Parecer Jurídico ou Certidão de Dispensa (ON AGU 69/2021), Portarias de Designação e Autorização da Autoridade; ordem lógica de datas/horários.
2. Objeto e Especificação Técnica: Cruze códigos CATMAT/CATSER do DFD com a descrição do ETP e do TR. Aponte divergências de descrição, unidades de medida, duplicidade de códigos, indicação indevida de marca/modelo sem justificativa (art. 41, I) ou especificações genéricas/restritivas.
3. Planejamento e Valores: Vinculação ao PCA, metodologia da pesquisa de preços (IN 65/2021), data-base do orçamento estimado (não pode ser futura), enquadramento no limite legal atualizado e comprovação de não fracionamento (art. 75, § 1º).
4. Regras da Disputa e Execução: Exclusividade ME/EPP (art. 48, I LC 123/2006), critério de julgamento e parcelamento, habilitação proporcional (art. 70, III e art. 20 IN 67/2021), prazos de recebimento, liquidação e pagamento (art. 7º, § 2º IN 77/2022), vigência, garantia, sanções administrativas e hipóteses de extinção contratual.
5. Coerência Interdocumental: Numerações de processo/aviso, prazos, percentuais de multas, quantitativos e valores entre todos os arquivos anexados.

# REGRAS DE RIGOR E CITAÇÃO OBRIGATÓRIA
- Nunca aponte um erro de forma vaga: indique exatamente o documento, a página (se houver) e o item/cláusula.
- Para CADA inconsistência, cite o dispositivo legal exato e o extrato/síntese fiel do comando normativo.
- Rigor anti-alucinação: baseie-se estritamente nos documentos fornecidos. Se um documento obrigatório não constar, declare "Ausente nos autos".
- Sempre que o erro puder ser corrigido diretamente no texto da minuta, preencha 'trecho_original' com a frase EXATA (literal) do documento e 'trecho_sugerido' com a redação corrigida pronta para substituição.

# FORMATO DE SAÍDA (JSON ESTRITO)
Retorne EXCLUSIVAMENTE um objeto JSON válido com a seguinte estrutura:
{
  "parecer_preventivo": {
    "veredito": "SOBRESTAR_PARA_SANEAMENTO" | "APTO_COM_RESSALVAS" | "APTO_PARA_PUBLICACAO",
    "resumo_executivo": "Texto objetivo de 2 a 4 frases sintetizando a situação jurídica do processo."
  },
  "achados": [
    {
      "id": "AUD-01",
      "eixo": "NORMATIVA" | "PROCESSUAL" | "CONTRADICAO_INTERNA",
      "criticidade": "CRITICA" | "ALTA" | "MEDIA" | "BAIXA",
      "titulo": "Título curto e técnico do achado",
      "localizacao": "Documento / Item / Cláusula / Página",
      "descricao_erro": "Descrição objetiva e detalhada do erro encontrado",
      "base_legal": "Lei/Decreto/IN/Súmula/ON com artigo, parágrafo e inciso exatos",
      "extrato_norma": "Síntese fiel do comando normativo que ampara a correção",
      "acao_corretiva": "Ação corretiva exata a ser adotada",
      "setor_responsavel": "Setor Requisitante | Seção de Licitações | Ordenador de Despesas | Assessoria Jurídica",
      "trecho_original": "Trecho literal exato do documento a ser substituído (ou string vazia se for ausência de documento/erro estrutural)",
      "trecho_sugerido": "Redação substitutiva corrigida (ou string vazia se não aplicável)"
    }
  ],
  "roteiro_saneamento": [
    "1. Passo cronológico exato...",
    "2. Passo cronológico exato..."
  ]
}
"""


def _extract_json_from_text(raw: str) -> Dict[str, Any]:
    raw_clean = (raw or "").strip()
    if raw_clean.startswith("```"):
        raw_clean = re.sub(r"^```(?:json)?\s*", "", raw_clean)
        raw_clean = re.sub(r"\s*```$", "", raw_clean)
    try:
        return json.loads(raw_clean)
    except Exception:
        start = raw_clean.find("{")
        end = raw_clean.rfind("}")
        if start != -1 and end != -1 and end > start:
            return json.loads(raw_clean[start : end + 1])
        raise


def run_gemini_deep_audit(
    combined_text: str,
    context_note: str = "",
    api_key_override: str = "",
    model_name: str = "gemini-2.5-flash",
) -> Dict[str, Any]:
    """
    Executa a Auditoria Jurídica Profunda via API oficial do Google Gemini (REST)
    usando Structured Output JSON. Retorna status='SKIPPED' se não houver chave configurada.
    """
    api_key = (api_key_override or os.environ.get("GEMINI_API_KEY") or "").strip()
    if not api_key:
        return {
            "status": "SKIPPED",
            "reason": "GEMINI_API_KEY não configurada no servidor nem informada na tela.",
            "achados": [],
            "roteiro_saneamento": [],
        }

    # Limita texto muito longo de forma segura (Gemini 2.5 suporta 1M tokens; 320k chars ~ 80k tokens)
    text_slice = (combined_text or "").strip()[:320000]
    if len(text_slice) < 40:
        return {
            "status": "SKIPPED",
            "reason": "Texto insuficiente para auditoria semântica.",
            "achados": [],
            "roteiro_saneamento": [],
        }

    user_prompt = (
        f"CONTEXTO ADICIONAL DO USUÁRIO: {context_note.strip() or 'Processo de contratação pública federal sob a Lei nº 14.133/2021.'}\n\n"
        f"=== AUTOS / DOCUMENTOS ANEXADOS PARA AUDITORIA ===\n{text_slice}"
    )

    url = f"https://generativelanguage.googleapis.com/v1beta/models/{model_name}:generateContent?key={api_key}"
    payload = {
        "system_instruction": {"parts": [{"text": AUDITOR_SYSTEM_PROMPT}]},
        "contents": [{"role": "user", "parts": [{"text": user_prompt}]}],
        "generationConfig": {
            "temperature": 0.15,
            "responseMimeType": "application/json",
        },
    }

    try:
        resp = requests.post(url, json=payload, timeout=90)
        if resp.status_code != 200:
            return {
                "status": "ERROR",
                "reason": f"Gemini API HTTP {resp.status_code}: {resp.text[:200]}",
                "achados": [],
                "roteiro_saneamento": [],
            }
        data = resp.json()
        candidates = data.get("candidates") or []
        if not candidates:
            return {
                "status": "ERROR",
                "reason": "Resposta vazia da API Gemini.",
                "achados": [],
                "roteiro_saneamento": [],
            }
        parts = candidates[0].get("content", {}).get("parts", [])
        raw_text = "".join(p.get("text", "") for p in parts if isinstance(p, dict))
        parsed = _extract_json_from_text(raw_text)

        achados = parsed.get("achados") or []
        roteiro = parsed.get("roteiro_saneamento") or []
        parecer = parsed.get("parecer_preventivo") or {
            "veredito": "SOBRESTAR_PARA_SANEAMENTO" if achados else "APTO_PARA_PUBLICACAO",
            "resumo_executivo": "Auditoria preventiva concluída.",
        }

        return {
            "status": "SUCCESS",
            "model": model_name,
            "parecer_preventivo": parecer,
            "achados": achados,
            "roteiro_saneamento": roteiro,
        }
    except Exception as exc:
        return {
            "status": "ERROR",
            "reason": f"Falha na chamada à API Gemini: {exc}",
            "achados": [],
            "roteiro_saneamento": [],
        }


def build_despacho_saneamento_docx_bytes(payload: Dict[str, Any]) -> bytes:
    """Gera o Despacho de Saneamento + Checklist Ponto a Ponto em Word (.docx) pronto para o SEI."""
    om_name = (payload.get("om_name") or "ORGANIZAÇÃO MILITAR / ÓRGÃO PÚBLICO").strip()
    processo_ref = (payload.get("processo_ref") or payload.get("filename") or "Processo Administrativo").strip()
    parecer = payload.get("parecer_preventivo") or {}
    achados = payload.get("achados") or []
    roteiro = payload.get("roteiro_saneamento") or []
    now_str = datetime.now().strftime("%d/%m/%Y às %H:%M")

    doc = Document()
    for sec in doc.sections:
        sec.top_margin = Cm(2.0)
        sec.bottom_margin = Cm(2.0)
        sec.left_margin = Cm(2.2)
        sec.right_margin = Cm(2.0)

    doc.styles["Normal"].font.name = "Calibri"
    doc.styles["Normal"].font.size = Pt(10)

    p_hdr = doc.add_paragraph()
    p_hdr.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r1 = p_hdr.add_run(f"{om_name.upper()}\n")
    r1.bold = True
    r1.font.size = Pt(12)
    r1.font.color.rgb = RGBColor(30, 45, 61)
    r2 = p_hdr.add_run("DESPACHO DE SANEAMENTO PREVENTIVO — AUDITORIA JURÍDICA (LEI Nº 14.133/2021)")
    r2.bold = True
    r2.font.size = Pt(10)
    r2.font.color.rgb = RGBColor(140, 105, 45)

    p_meta = doc.add_paragraph()
    p_meta.add_run("Processo / Autos Analisados: ").bold = True
    p_meta.add_run(f"{processo_ref}\n")
    p_meta.add_run("Data/Hora da Auditoria: ").bold = True
    p_meta.add_run(f"{now_str}\n")
    p_meta.add_run("Veredito Preventivo: ").bold = True
    p_meta.add_run(str(parecer.get("veredito") or "SOBRESTAR PARA SANEAMENTO").replace("_", " "))

    if parecer.get("resumo_executivo"):
        p_res = doc.add_paragraph()
        p_res.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p_res.add_run("1. PARECER PREVENTIVO: ").bold = True
        p_res.add_run(str(parecer.get("resumo_executivo")))

    # Painel Resumo
    p_p = doc.add_paragraph()
    p_p.add_run(f"2. ACHADOS DETALHADOS E CHECKLIST DE CORREÇÃO ({len(achados)} apontamento(s))").bold = True

    eixo_labels = [
        ("NORMATIVA", "2.1 Inconsistências Normativas (Lei nº 14.133/2021 e INs)"),
        ("PROCESSUAL", "2.2 Erros Processuais e Falhas de Instrução (Art. 72)"),
        ("CONTRADICAO_INTERNA", "2.3 Contradições Internas e Revisão Redacional"),
    ]
    for eixo_key, eixo_title in eixo_labels:
        subset = [a for a in achados if str(a.get("eixo") or "NORMATIVA").upper() == eixo_key]
        if not subset:
            continue
        p_sub = doc.add_paragraph()
        rs = p_sub.add_run(eixo_title)
        rs.bold = True
        rs.font.color.rgb = RGBColor(30, 45, 61)

        tbl = doc.add_table(rows=1, cols=5)
        tbl.style = "Table Grid"
        for i, h in enumerate(["ID / Risco", "Localização", "Erro Identificado", "Base Legal & Extrato", "Ação Corretiva / Setor"]):
            c = tbl.rows[0].cells[i]
            c.text = h
            for run in c.paragraphs[0].runs:
                run.bold = True
                run.font.size = Pt(8.5)
        for idx, a in enumerate(subset, start=1):
            row = tbl.add_row().cells
            row[0].text = f"{a.get('id') or f'AUD-{idx:02d}'}\n[{a.get('criticidade') or 'ALTA'}]"
            row[1].text = str(a.get("localizacao") or "Autos")
            row[2].text = f"{a.get('titulo') or ''}\n{a.get('descricao_erro') or ''}".strip()
            row[3].text = f"{a.get('base_legal') or ''}\n{a.get('extrato_norma') or ''}".strip()
            row[4].text = f"{a.get('acao_corretiva') or ''}\n(Setor: {a.get('setor_responsavel') or 'Licitações'})".strip()
            for c in row:
                for run in c.paragraphs[0].runs:
                    run.font.size = Pt(8)

    if roteiro:
        p_r = doc.add_paragraph()
        p_r.add_run("\n3. ROTEIRO CRONOLÓGICO DE SANEAMENTO").bold = True
        for step in roteiro:
            ps = doc.add_paragraph(str(step))
            ps.paragraph_format.space_after = Pt(3)

    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()


def build_despacho_saneamento_pdf_bytes(payload: Dict[str, Any]) -> bytes:
    """Gera o Despacho de Saneamento Preventivo em PDF pronto para juntada no SEI."""
    om_name = (payload.get("om_name") or "ORGANIZAÇÃO MILITAR / ÓRGÃO PÚBLICO").strip()
    processo_ref = (payload.get("processo_ref") or payload.get("filename") or "Processo Administrativo").strip()
    parecer = payload.get("parecer_preventivo") or {}
    achados = payload.get("achados") or []
    roteiro = payload.get("roteiro_saneamento") or []
    now_str = datetime.now().strftime("%d/%m/%Y às %H:%M")

    buf = io.BytesIO()
    doc = SimpleDocTemplate(
        buf,
        pagesize=A4,
        leftMargin=1.6 * cm,
        rightMargin=1.6 * cm,
        topMargin=1.8 * cm,
        bottomMargin=1.8 * cm,
    )
    styles = getSampleStyleSheet()
    st_hdr = ParagraphStyle("H", parent=styles["Normal"], fontName="Helvetica-Bold", fontSize=11, alignment=TA_CENTER, textColor=colors.HexColor("#1e2d3d"))
    st_sub = ParagraphStyle("S", parent=styles["Normal"], fontName="Helvetica-Bold", fontSize=9, alignment=TA_CENTER, textColor=colors.HexColor("#9a7632"), spaceAfter=8)
    st_sec = ParagraphStyle("Sec", parent=styles["Normal"], fontName="Helvetica-Bold", fontSize=9.5, textColor=colors.HexColor("#1e2d3d"), spaceBefore=8, spaceAfter=4)
    st_body = ParagraphStyle("B", parent=styles["Normal"], fontName="Helvetica", fontSize=8.5, leading=12, alignment=TA_JUSTIFY, spaceAfter=4)
    st_cell = ParagraphStyle("C", parent=styles["Normal"], fontName="Helvetica", fontSize=7.2, leading=9.2)

    story: List[Any] = [
        Paragraph(om_name.upper(), st_hdr),
        Paragraph("DESPACHO DE SANEAMENTO PREVENTIVO — AUDITORIA JURÍDICA (LEI Nº 14.133/2021)", st_sub),
        Paragraph(
            f"<b>Autos Analisados:</b> {processo_ref} &nbsp;|&nbsp; <b>Data:</b> {now_str}<br/>"
            f"<b>Veredito Preventivo:</b> {str(parecer.get('veredito') or 'SOBRESTAR PARA SANEAMENTO').replace('_', ' ')}",
            st_body,
        ),
    ]
    if parecer.get("resumo_executivo"):
        story.append(Paragraph(f"<b>1. PARECER PREVENTIVO:</b> {parecer.get('resumo_executivo')}", st_body))

    story.append(Paragraph(f"2. ACHADOS DETALHADOS ({len(achados)} apontamento(s))", st_sec))
    if achados:
        tdata = [["ID / Risco", "Localização", "Erro Identificado", "Base Legal & Extrato", "Ação Corretiva"]]
        for idx, a in enumerate(achados, start=1):
            tdata.append([
                Paragraph(f"<b>{a.get('id') or f'AUD-{idx:02d}'}</b><br/>{a.get('criticidade') or 'ALTA'}", st_cell),
                Paragraph(str(a.get("localizacao") or "Autos"), st_cell),
                Paragraph(f"<b>{a.get('titulo') or ''}</b><br/>{a.get('descricao_erro') or ''}", st_cell),
                Paragraph(f"<b>{a.get('base_legal') or ''}</b><br/>{a.get('extrato_norma') or ''}", st_cell),
                Paragraph(f"{a.get('acao_corretiva') or ''}<br/><i>({a.get('setor_responsavel') or 'Licitações'})</i>", st_cell),
            ])
        t = Table(tdata, colWidths=[1.8 * cm, 2.7 * cm, 4.6 * cm, 4.5 * cm, 4.2 * cm], repeatRows=1)
        t.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1e2d3d")),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
            ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
            ("FONTSIZE", (0, 0), (-1, 0), 7.5),
            ("GRID", (0, 0), (-1, -1), 0.35, colors.HexColor("#cbd5e1")),
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#f8fafc")]),
        ]))
        story.append(t)

    if roteiro:
        story.append(Spacer(1, 6))
        story.append(Paragraph("3. ROTEIRO CRONOLÓGICO DE SANEAMENTO", st_sec))
        for step in roteiro:
            story.append(Paragraph(str(step), st_body))

    doc.build(story)
    return buf.getvalue()
