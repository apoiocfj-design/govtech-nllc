"""
Motor Híbrido de Auditoria Jurídica Preventiva (Lei nº 14.133/2021, INs SEGES/ME e SEGES/MGI, AGU/TCU)
via Google Gemini API (`gemini-3.8-flash` / fallback automático) + Gerador do Checklist Detalhado de
Conformidade Normativa e Saneamento Processual (Word .DOCX e PDF) no padrão oficial da Marinha / AGU.
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
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
from docx.shared import Cm, Pt, RGBColor
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_JUSTIFY, TA_LEFT, TA_RIGHT
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import cm
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle


AUDITOR_SYSTEM_PROMPT = """# PAPEL E PERFIL DE ATUAÇÃO
Atue como um Auditor Federal de Controle Externo (TCU/CGU) e Parecerista Jurídico Sênior (CJU/AGU), especialista em Licitações e Contratos Administrativos no âmbito da Administração Pública Federal brasileira. Sua postura deve ser estritamente técnica, exaustiva, detalhista, cética e rigorosa quanto à legalidade, à instrução processual e à coerência interna dos autos.

# OBJETIVO
Realizar uma auditoria preventiva COMPLETA e EXAUSTIVA nos documentos anexados (identificando de 10 a 35+ apontamentos quando houver falhas nos autos, sem omitir nenhum erro), cobrindo:
1. Inconsistências normativas e ilegalidades (Eixo 1 — IDs N-01, N-02, N-03...);
2. Erros processuais, lacunas de instrução, documentos ausentes e falhas cronológicas (Eixo 2 — IDs P-01, P-02, P-03...);
3. Contradições internas entre os artefatos, erros de adaptação de minutas-padrão ("copia e cola"), cláusulas impertinentes (ex.: obras/engenharia em compra de bens), lacunas em branco após dois-pontos, saltos de numeração e falhas redacionais (Eixo 3 — IDs R-01, R-02, R-03...).

# ARCABOUÇO NORMATIVO DE REFERÊNCIA
Analise os autos à luz de todo o ordenamento aplicável às contratações públicas federais:
- Leis: Lei nº 14.133/2021 (NLLC), Lei Complementar nº 123/2006 (ME/EPP — art. 48, I exclusividade até R$ 80.000,00), Lei nº 9.784/1999 (Processo Administrativo), Lei Complementar nº 101/2000 (LRF — art. 16) e Lei nº 14.195/2021 (art. 41 — extinção da EIRELI / substituição por SLU).
- Decretos: Decretos anuais de atualização de valores da Lei nº 14.133/2021 (Decreto nº 12.343/2024 / Decreto nº 12.807/2025), Decreto nº 10.947/2022 (PCA/DFD — arts. 15 e 16), Decreto nº 11.246/2022 (Agente de Contratação e Fiscais), Decreto nº 11.462/2023 (SRP) e Decreto nº 8.538/2015 (art. 6º — ME/EPP).
- Instruções Normativas (SEGES/ME e SEGES/MGI): IN nº 65/2021 (Pesquisa de Preços), IN nº 67/2021 (Dispensa Eletrônica — art. 4º, § 2º não fracionamento; art. 20 habilitação simplificada), IN nº 58/2022 (ETP), IN nº 81/2022 (TR), IN nº 77/2022 (art. 7º, § 2º — redução pela metade dos prazos de liquidação E de pagamento para 5 dias úteis nas despesas até o limite do art. 75, II), IN nº 98/2022 e IN nº 05/2017.
- Jurisprudência e Orientações: Súmulas do TCU (Súmula 247 — parcelamento; Súmula 270 — vedação de marca sem justificativa técnica), Orientações Normativas da AGU (ON nº 69/2021 — dispensa de parecer jurídico em pequeno valor; ON nº 84/2024) e Notas Explicativas das minutas-padrão da AGU/SEGES.

# EIXOS OBRIGATÓRIOS DE VERIFICAÇÃO (VERIFIQUE TODOS SEM EXCEÇÃO)
1. Instrução e Cronologia (Arts. 18 e 72 da Lei nº 14.133/2021):
   - Verifique se TODOS os documentos obrigatórios constam nos autos: DFD, ETP, Pesquisa de Preços e Mapa Comparativo assinado (art. 72, II e IN 65/2021), TR, Declaração Formal de Disponibilidade Orçamentária (art. 72, IV), Certidão/Controle de Não Fracionamento de Despesa (art. 75, § 1º e art. 4º, § 2º IN 67/2021), Parecer Jurídico ou Certidão de Dispensa de Análise Jurídica (art. 72, III, art. 53, § 5º e ON AGU 69/2021), Portarias de Designação de Agente de Contratação/Fiscal (art. 7º e art. 117) e Autorização da Autoridade. Se qualquer um não estiver presente no texto dos autos, gere um achado CRÍTICO no eixo PROCESSUAL declarando "Ausente nos autos".
   - Compare as datas e horários de conclusão/assinatura de todas as peças (ex.: Aviso concluído antes da Solicitação de Abertura ou antes do TR; data-base do orçamento estimado fixada no futuro em relação à elaboração do TR; uso prematuro do art. 72, VIII na solicitação inicial).
2. Objeto, CATMAT/CATSER e Especificação Técnica:
   - Cruze item a item os códigos CATMAT/CATSER e descrições do DFD com a tabela do ETP e do TR. Aponte TODAS as divergências de especificação (ex.: binóculo termal vs óptico 7x50, bolsa de morteiro vs bolsa estanque, luva nitrílica/motociclismo vs neoprene/couro), divergências de unidade de fornecimento (Embalagem vs PAR vs UN) e códigos CATMAT duplicados para itens distintos. Preencha também a lista 'divergencias_catmat' quando houver divergências entre DFD e ETP/TR!
   - Aponte indicação de marca/modelo (ex.: "Mark V") sem a expressão "ou equivalente/similar/de melhor qualidade" e sem justificativa técnica (art. 41, I da Lei 14.133/2021 e Súmula TCU 270).
   - Aponte descrições genéricas no TR sem parâmetros técnicos mínimos (dimensões, lúmens, grau IP, certificação Ex/INMETRO, tamanhos).
3. Regras da Disputa, Pagamento, Habilitação e Execução:
   - Verifique o prazo de pagamento na IN 77/2022 (art. 7º, § 2º): para contratações do art. 75, II, tanto a liquidação quanto o pagamento devem ser reduzidos à metade (5 dias úteis, e não 10 dias úteis).
   - Verifique fundamento legal errado em cláusula de extinção (ex.: citar art. 75, VIII emergencial quando o processo é art. 75, II baixo valor).
   - Verifique omissão de participação exclusiva ME/EPP no corpo do Aviso (item 3) quando a capa marca "[SIM]" e o valor é <= R$ 80.000,00 (art. 48, I LC 123/2006).
   - Verifique exigências excessivas de habilitação (atestado de capacidade técnica / fiscal estadual completa em compra de pequeno valor e entrega imediata — art. 70, III NLLC e art. 20 IN 67/2021).
4. Contradições Internas e Revisão Redacional ("Copia e Cola"):
   - Numerações divergentes para o mesmo Aviso (ex.: cabeçalho vs capa vs preâmbulo).
   - Contradição de vigência (ex.: Solicitação fala em 12 meses e TR fixa 60 dias com entrega única).
   - Indecisão sobre parcelamento no ETP vs julgamento por item no Aviso.
   - Divergência de prazo de aceite da Nota de Empenho (ex.: 2 dias úteis no Aviso vs 5 dias no Anexo I do TR, ou erro "dias dias").
   - Contradição no prazo de recebimento definitivo (ex.: item 8.3 fixa 5 dias úteis e item 8.4 amplia para 15 dias úteis no art. 75, II).
   - Divergência de percentuais de multas entre Aviso (ex.: 5%) e TR (ex.: 0,1%/dia e 10% a 15%).
   - Erro entre algarismo e extenso no intervalo mínimo de lances (ex.: "R$ 0,10 (um centavos)").
   - Remissão vazia no TR dispensando garantia "pelas razões do ETP" quando o ETP não trata de garantia.
   - Cláusulas impertinentes mantidas de modelo-padrão (ex.: item 6.12 de obras e serviços de engenharia com 75%/85%; dotação de exercícios subsequentes em entrega imediata) e cláusulas truncadas terminando em dois-pontos (:) ou saltos de numeração.

# REGRAS DE RIGOR E CITAÇÃO DIRETA OBRIGATÓRIA
- Em 'extrato_norma', forneça SEMPRE a CITAÇÃO DIRETA (entre aspas) ou transcrição fiel do texto do artigo/inciso/parágrafo da Lei/IN/Decreto que fundamenta o erro.
- Em 'texto_apos_correcao', escreva EXATAMENTE como a cláusula, item ou providência deve ficar redigida após a correção.
- Em 'trecho_original', quando o erro corresponder a uma frase presente no texto dos documentos anexados, copie um trecho literal curto (de 30 a 160 caracteres) idêntico ao texto original para permitir o destaque automático em vermelho/amarelo no documento!

# FORMATO DE SAÍDA (JSON ESTRITO)
Retorne EXCLUSIVAMENTE um objeto JSON válido com esta estrutura:
{
  "metadados_processo": {
    "numero_processo": "Ex: 63436.000936/2026-88 (ou Não informado)",
    "modalidade_fundamento": "Ex: Dispensa Eletrônica — Art. 75, II, Lei nº 14.133/2021",
    "uasg_orgao": "Ex: 782340 — Capitania Fluvial de Juazeiro",
    "valor_total_estimado": "Ex: R$ 5.040,00",
    "objeto_contratacao": "Resumo do objeto",
    "artefatos_analisados": "Ex: Solicitação de Abertura, ETP, DFD, TR e Aviso"
  },
  "parecer_preventivo": {
    "veredito": "SOBRESTAR_PARA_SANEAMENTO" | "APTO_COM_RESSALVAS" | "APTO_PARA_PUBLICACAO",
    "resumo_executivo": "Parecer técnico robusto sintetizando a quantidade de apontamentos críticos/altos/médios/baixos, os principais riscos e a recomendação de sustação antes da publicação no PNCP/Compras.gov.br."
  },
  "achados": [
    {
      "id": "N-01",
      "eixo": "NORMATIVA" | "PROCESSUAL" | "CONTRADICAO_INTERNA",
      "criticidade": "CRITICA" | "ALTA" | "MEDIA" | "BAIXA",
      "titulo": "Título técnico e direto do erro",
      "localizacao": "Artefato(s) e Item/Cláusula/Página exatos (ex: TR 42/2026 - Item 8.20)",
      "descricao_erro": "Descrição detalhada e objetiva do que está errado no documento/processo",
      "base_legal": "Dispositivo legal exato (ex: Art. 7º, § 2º, da IN SEGES/ME nº 77/2022)",
      "extrato_norma": "Citação direta entre aspas do comando normativo (ex: \"Art. 7º ... § 2º No caso de despesas cujos valores não ultrapassem o limite de que trata o inciso II do art. 75 da Lei nº 14.133, de 2021, os prazos previstos nos incisos I e II do caput serão reduzidos pela metade.\")",
      "acao_corretiva": "Orientação clara da providência corretiva",
      "texto_apos_correcao": "Redação exata de como o texto/cláusula deve ficar após a correção",
      "setor_responsavel": "Setor Requisitante | Subseção de Intendência / Licitações | Ordenador de Despesas | Setor Financeiro",
      "trecho_original": "Trecho literal exato (30 a 160 caracteres) presente no documento para destaque automático, ou string vazia se for documento ausente",
      "trecho_sugerido": "Novo texto substituto para o trecho_original"
    }
  ],
  "divergencias_catmat": [
    {
      "item": "2",
      "descricao_etp_tr": "Descrição real no ETP/TR",
      "catmat_atual": "Código CATMAT",
      "descricao_dfd_catalogo": "Descrição divergente puxada no DFD",
      "unidade_dfd_vs_tr": "Ex: Embalagem vs. PAR",
      "correcao_necessaria": "Substituir por CATMAT compatível e uniformizar unidade"
    }
  ],
  "roteiro_saneamento": [
    {
      "etapa": "1ª Etapa",
      "atividade": "Descrição detalhada da atividade cronológica",
      "itens_atendidos": "N-01, N-02, N-03",
      "setor_executor": "Setor Requisitante"
    }
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
    model_name: str = "gemini-3.8-flash",
) -> Dict[str, Any]:
    """
    Executa a Auditoria Jurídica Profunda via API oficial do Google Gemini (REST)
    com fallback automático entre modelos ativos (`gemini-3.8-flash`, `gemini-2.5-flash`).
    """
    api_key = (api_key_override or os.environ.get("GEMINI_API_KEY") or "").strip()
    if not api_key:
        return {
            "status": "SKIPPED",
            "reason": "GEMINI_API_KEY não configurada no servidor nem informada na tela.",
            "achados": [],
            "divergencias_catmat": [],
            "roteiro_saneamento": [],
        }

    text_slice = (combined_text or "").strip()[:350000]
    if len(text_slice) < 40:
        return {
            "status": "SKIPPED",
            "reason": "Texto insuficiente para auditoria semântica.",
            "achados": [],
            "divergencias_catmat": [],
            "roteiro_saneamento": [],
        }

    user_prompt = (
        f"CONTEXTO ADICIONAL DO PROCESSO: {context_note.strip() or 'Processo de contratação pública federal sob a Lei nº 14.133/2021 (verifique todos os artefatos anexados, cruzamento DFD x ETP x TR x Aviso, prazos da IN 77/2022, documentos obrigatórios do art. 72 e contradições internas).'}\n\n"
        f"=== TEXTO INTEGRAL DOS AUTOS / DOCUMENTOS ANEXADOS PARA AUDITORIA ===\n{text_slice}"
    )

    candidate_models = []
    for m in [model_name, "gemini-3.8-flash", "gemini-2.5-flash", "gemini-2.0-flash"]:
        m_clean = (m or "").strip()
        if m_clean and m_clean not in candidate_models:
            candidate_models.append(m_clean)

    payload = {
        "system_instruction": {"parts": [{"text": AUDITOR_SYSTEM_PROMPT}]},
        "contents": [{"role": "user", "parts": [{"text": user_prompt}]}],
        "generationConfig": {
            "temperature": 0.1,
            "responseMimeType": "application/json",
        },
    }

    last_error = ""
    for m_try in candidate_models:
        url = f"https://generativelanguage.googleapis.com/v1beta/models/{m_try}:generateContent?key={api_key}"
        try:
            resp = requests.post(url, json=payload, timeout=110)
            if resp.status_code == 404:
                last_error = f"Model {m_try} HTTP 404"
                continue
            if resp.status_code != 200:
                last_error = f"Gemini API ({m_try}) HTTP {resp.status_code}: {resp.text[:220]}"
                continue

            data = resp.json()
            candidates = data.get("candidates") or []
            if not candidates:
                last_error = f"Resposta vazia do modelo {m_try}."
                continue

            parts = candidates[0].get("content", {}).get("parts", [])
            raw_text = "".join(p.get("text", "") for p in parts if isinstance(p, dict))
            parsed = _extract_json_from_text(raw_text)

            achados = parsed.get("achados") or []
            divergencias_catmat = parsed.get("divergencias_catmat") or []
            roteiro = parsed.get("roteiro_saneamento") or []
            metadados = parsed.get("metadados_processo") or {}
            parecer = parsed.get("parecer_preventivo") or {
                "veredito": "SOBRESTAR_PARA_SANEAMENTO" if achados else "APTO_PARA_PUBLICACAO",
                "resumo_executivo": "Auditoria preventiva concluída.",
            }

            return {
                "status": "SUCCESS",
                "model": m_try,
                "metadados_processo": metadados,
                "parecer_preventivo": parecer,
                "achados": achados,
                "divergencias_catmat": divergencias_catmat,
                "roteiro_saneamento": roteiro,
            }
        except Exception as exc:
            last_error = f"Falha no modelo {m_try}: {exc}"

    return {
        "status": "ERROR",
        "reason": last_error or "Não foi possível completar a auditoria com os modelos Gemini.",
        "achados": [],
        "divergencias_catmat": [],
        "roteiro_saneamento": [],
    }


def merge_deep_audit_into_findings(
    audited_doc: Dict[str, Any],
    deep_audit: Dict[str, Any],
) -> None:
    """
    Injeta TODOS os achados da Auditoria Profunda (Gemini) dentro da lista principal
    de apontamentos (`audited_doc["findings"]`) e mapeia cada achado ao elemento/parágrafo
    correspondente da minuta para destaque amarelo/vermelho no editor lateral.
    """
    if not deep_audit or deep_audit.get("status") != "SUCCESS":
        return
    achados = deep_audit.get("achados") or []
    if not achados:
        return

    try:
        from core.audited_document_engine import compute_word_diff
    except ImportError:
        try:
            from audited_document_engine import compute_word_diff
        except ImportError:
            compute_word_diff = None

    elements = audited_doc.get("elements") or []
    existing_findings: List[Dict[str, Any]] = list(audited_doc.get("findings") or [])
    existing_titles = {str(f.get("title") or "").strip().lower() for f in existing_findings}

    for idx, a in enumerate(achados, start=1):
        if not isinstance(a, dict):
            continue
        titulo = str(a.get("titulo") or f"Apontamento {idx}").strip()
        if titulo.lower() in existing_titles:
            continue

        eixo = str(a.get("eixo") or "NORMATIVA").upper()
        crit = str(a.get("criticidade") or "ALTA").upper()
        loc = str(a.get("localizacao") or f"Item {idx}").strip()
        desc_erro = str(a.get("descricao_erro") or "").strip()
        base_legal = str(a.get("base_legal") or "Lei nº 14.133/2021").strip()
        extrato = str(a.get("extrato_norma") or "").strip()
        acao = str(a.get("acao_corretiva") or "").strip()
        texto_corrigido = str(a.get("texto_apos_correcao") or a.get("trecho_sugerido") or acao).strip()
        setor = str(a.get("setor_responsavel") or "Seção de Licitações").strip()
        trecho_orig = str(a.get("trecho_original") or "").strip()
        trecho_sug = str(a.get("trecho_sugerido") or texto_corrigido).strip()

        # Tenta localizar o elemento no documento que contém o trecho_original
        matched_el_idx: Optional[int] = None
        orig_full_text = trecho_orig
        revised_full_text = trecho_sug

        if trecho_orig and len(trecho_orig) >= 12:
            needle = trecho_orig.lower()
            for el_i, el in enumerate(elements):
                if el.get("type") in ("CLAUSE", "PARAGRAPH"):
                    el_txt = str(el.get("original_text") or el.get("text") or "")
                    if needle in el_txt.lower():
                        matched_el_idx = el_i
                        orig_full_text = el_txt
                        pattern = re.compile(re.escape(trecho_orig), re.IGNORECASE)
                        revised_full_text = pattern.sub(trecho_sug or texto_corrigido, el_txt, count=1)
                        break

        if matched_el_idx is not None and compute_word_diff and orig_full_text != revised_full_text:
            preview_diff = compute_word_diff(orig_full_text, revised_full_text)
            target_type = "element"
        else:
            target_type = "append_clause" if eixo == "PROCESSUAL" else "element"
            preview_diff = []
            if trecho_orig:
                preview_diff.append({"type": "removed", "text": trecho_orig})
            if texto_corrigido:
                preview_diff.append({"type": "added", "text": texto_corrigido})

        legal_combined = f"{base_legal}" + (f" — Citação da Norma: {extrato}" if extrato else "")
        explanation_combined = (
            f"[{eixo}] {desc_erro}"
            + (f" | Ação Corretiva: {acao}" if acao else "")
            + (f" | Responsável: {setor}" if setor else "")
        )

        existing_findings.append({
            "id": str(a.get("id") or f"AUD_{idx:02d}"),
            "type": "OMISSAO_OBRIGATORIA" if eixo == "PROCESSUAL" else "VULNERABILIDADE_JURIDICA",
            "finding_type": eixo,
            "eixo": eixo,
            "severity": crit,
            "title": titulo,
            "section_title": titulo,
            "clause_reference": loc,
            "original_clause": loc,
            "target_type": target_type,
            "element_idx": matched_el_idx,
            "docx_p_idx": elements[matched_el_idx].get("docx_p_idx") if matched_el_idx is not None else None,
            "original_text": orig_full_text or f"(Situação atual em {loc}: {desc_erro})",
            "original_excerpt": trecho_orig or desc_erro,
            "suggested_text": revised_full_text if matched_el_idx is not None else texto_corrigido,
            "agu_standard_clause": texto_corrigido,
            "explanation": explanation_combined,
            "reason": desc_erro,
            "legal_foundation": legal_combined,
            "extrato_norma": extrato,
            "acao_corretiva": acao,
            "texto_apos_correcao": texto_corrigido,
            "setor_responsavel": setor,
            "preview_diff": preview_diff,
            "status": "pending",
            "approved": False,
        })

    audited_doc["findings"] = existing_findings
    summary = audited_doc.get("audit_summary") or {}
    summary["findings"] = existing_findings
    summary["total_findings"] = len(existing_findings)
    # Ajusta score de conformidade proporcionalmente à gravidade dos achados
    crit_count = sum(1 for f in existing_findings if str(f.get("severity")).upper() == "CRITICA")
    alta_count = sum(1 for f in existing_findings if str(f.get("severity")).upper() == "ALTA")
    media_count = sum(1 for f in existing_findings if str(f.get("severity")).upper() not in ("CRITICA", "ALTA"))
    penalty = crit_count * 7.0 + alta_count * 4.0 + media_count * 2.0
    summary["compliance_score"] = round(max(12.0, 100.0 - penalty), 1)
    audited_doc["audit_summary"] = summary


def _set_cell_bg(cell: Any, hex_color: str) -> None:
    shading = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{hex_color.lstrip("#")}"/>')
    cell._tc.get_or_add_tcPr().append(shading)


def build_despacho_saneamento_docx_bytes(payload: Dict[str, Any]) -> bytes:
    """
    Gera o CHECKLIST DETALHADO DE CONFORMIDADE NORMATIVA E SANEAMENTO PROCESSUAL em Word (.docx)
    com layout executivo completo:
    - Cabeçalho da OM + Metadados do Processo + Parecer Preventivo;
    - 0. Painel Resumo Quantitativo por Eixo e Criticidade;
    - 1. Inconsistências Normativas (com colunas: ID | Artefato/Item | O Que Está Errado | Norma + Citação Direta | Como Deve Ficar o Texto Após Correção / Ação | Responsável / Prioridade);
    - Tabela 1.1 de Divergências de CATMAT e Unidades (quando houver);
    - 2. Erros Processuais e de Instrução (Art. 72);
    - 3. Contradições Internas e Revisão Redacional;
    - 4. Roteiro Sequencial Recomendado para Saneamento;
    - 5. Termo de Conferência e Ateste de Saneamento (3 assinaturas).
    """
    om_name = (payload.get("om_name") or "MARINHA DO BRASIL").strip()
    processo_ref = (payload.get("processo_ref") or payload.get("filename") or "Processo Administrativo").strip()
    meta = payload.get("metadados_processo") or {}
    parecer = payload.get("parecer_preventivo") or {}
    achados = payload.get("achados") or []
    div_catmat = payload.get("divergencias_catmat") or []
    roteiro = payload.get("roteiro_saneamento") or []
    now_str = datetime.now().strftime("%d/%m/%Y às %H:%M")

    doc = Document()
    for sec in doc.sections:
        sec.top_margin = Cm(1.5)
        sec.bottom_margin = Cm(1.5)
        sec.left_margin = Cm(1.5)
        sec.right_margin = Cm(1.5)

    doc.styles["Normal"].font.name = "Calibri"
    doc.styles["Normal"].font.size = Pt(9.5)

    # BANNER DE CABEÇALHO E METADADOS DO PROCESSO
    hdr_tbl = doc.add_table(rows=4, cols=4)
    hdr_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    hdr_tbl.style = "Table Grid"

    c_top = hdr_tbl.cell(0, 0)
    c_top.merge(hdr_tbl.cell(0, 3))
    _set_cell_bg(c_top, "1E2D3D")
    p0 = c_top.paragraphs[0]
    p0.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_om = p0.add_run(f"{om_name.upper()}\n")
    r_om.bold = True
    r_om.font.size = Pt(11)
    r_om.font.color.rgb = RGBColor(255, 255, 255)
    r_t = p0.add_run("CHECKLIST DETALHADO DE CONFORMIDADE NORMATIVA E SANEAMENTO PROCESSUAL\n")
    r_t.bold = True
    r_t.font.size = Pt(10.5)
    r_t.font.color.rgb = RGBColor(245, 208, 125)
    r_sub = p0.add_run("Auditoria Jurídica Preventiva — Lei nº 14.133/2021, IN SEGES/ME nº 65/2021, IN nº 67/2021, IN nº 77/2022, AGU e TCU")
    r_sub.font.size = Pt(8.5)
    r_sub.font.color.rgb = RGBColor(215, 225, 235)

    meta_rows = [
        (
            "Processo Administrativo:",
            str(meta.get("numero_processo") or processo_ref),
            "Modalidade / Fundamento:",
            str(meta.get("modalidade_fundamento") or "Lei nº 14.133/2021"),
        ),
        (
            "Unidade Gestora (UASG):",
            str(meta.get("uasg_orgao") or om_name),
            "Valor Total Estimado:",
            str(meta.get("valor_total_estimado") or "Conforme autos"),
        ),
        (
            "Objeto da Contratação:",
            str(meta.get("objeto_contratacao") or "Conforme Termo de Referência"),
            "Artefatos Analisados:",
            str(meta.get("artefatos_analisados") or f"{processo_ref} (Em {now_str})"),
        ),
    ]
    for r_idx, (k1, v1, k2, v2) in enumerate(meta_rows, start=1):
        row_cells = hdr_tbl.rows[r_idx].cells
        _set_cell_bg(row_cells[0], "F1F5F9")
        _set_cell_bg(row_cells[2], "F1F5F9")
        row_cells[0].text = k1
        row_cells[1].text = v1
        row_cells[2].text = k2
        row_cells[3].text = v2
        for ci, c in enumerate(row_cells):
            for run in c.paragraphs[0].runs:
                run.font.size = Pt(8.5)
                if ci in (0, 2):
                    run.bold = True

    # BOX PARECER PREVENTIVO
    veredito_raw = str(parecer.get("veredito") or "SOBRESTAR_PARA_SANEAMENTO").replace("_", " ")
    p_box = doc.add_table(rows=1, cols=1)
    p_box.style = "Table Grid"
    c_p = p_box.cell(0, 0)
    _set_cell_bg(c_p, "FEF2F2" if "SOBRESTAR" in veredito_raw else "FFFBEB")
    pp = c_p.paragraphs[0]
    rp1 = pp.add_run(f"PARECER PREVENTIVO — {veredito_raw}\n")
    rp1.bold = True
    rp1.font.size = Pt(9.5)
    rp1.font.color.rgb = RGBColor(153, 27, 27)
    rp2 = pp.add_run(str(parecer.get("resumo_executivo") or "Recomenda-se o saneamento integral dos apontamentos antes da publicação no PNCP."))
    rp2.font.size = Pt(8.8)

    # 0. PAINEL RESUMO DE APONTAMENTOS
    p_sec0 = doc.add_paragraph()
    p_sec0.paragraph_format.space_before = Pt(8)
    r_s0 = p_sec0.add_run("0. PAINEL RESUMO DE APONTAMENTOS")
    r_s0.bold = True
    r_s0.font.size = Pt(10.5)
    r_s0.font.color.rgb = RGBColor(30, 45, 61)

    eixos_cfg = [
        ("NORMATIVA", "1. Inconsistências Normativas (Lei 14.133/2021 e INs)", "Risco de nulidade, entrega de bem errado e violação a INs"),
        ("PROCESSUAL", "2. Erros Processuais e Instrução (Art. 72 da Lei 14.133/2021)", "Instrução incompleta perante órgãos de controle (TCU/CGU)"),
        ("CONTRADICAO_INTERNA", "3. Contradições Internas e Revisão Redacional", "Insegurança jurídica na disputa, sanções e execução contratual"),
    ]
    t_res = doc.add_table(rows=1, cols=7)
    t_res.style = "Table Grid"
    for i, h in enumerate(["Eixo de Verificação", "Crítica", "Alta", "Média", "Baixa", "Total", "Impacto Principal"]):
        c = t_res.rows[0].cells[i]
        _set_cell_bg(c, "1E2D3D")
        c.text = h
        for run in c.paragraphs[0].runs:
            run.bold = True
            run.font.size = Pt(8.2)
            run.font.color.rgb = RGBColor(255, 255, 255)

    tot_c = tot_a = tot_m = tot_b = 0
    for ek, elbl, eimp in eixos_cfg:
        sub = [a for a in achados if str(a.get("eixo") or "NORMATIVA").upper() == ek]
        nc = sum(1 for a in sub if str(a.get("criticidade") or "").upper() == "CRITICA")
        na = sum(1 for a in sub if str(a.get("criticidade") or "").upper() == "ALTA")
        nm = sum(1 for a in sub if str(a.get("criticidade") or "").upper() == "MEDIA")
        nb = sum(1 for a in sub if str(a.get("criticidade") or "").upper() not in ("CRITICA", "ALTA", "MEDIA"))
        tot_c += nc
        tot_a += na
        tot_m += nm
        tot_b += nb
        rc = t_res.add_row().cells
        vals = [elbl, str(nc), str(na), str(nm), str(nb), str(len(sub)), eimp]
        for ci, v in enumerate(vals):
            rc[ci].text = v
            for run in rc[ci].paragraphs[0].runs:
                run.font.size = Pt(8.2)

    rc_tot = t_res.add_row().cells
    tot_vals = ["TOTAL GERAL DE APONTAMENTOS", str(tot_c), str(tot_a), str(tot_m), str(tot_b), str(len(achados)), "Saneamento obrigatório prévio à divulgação no PNCP"]
    for ci, v in enumerate(tot_vals):
        _set_cell_bg(rc_tot[ci], "F1F5F9")
        rc_tot[ci].text = v
        for run in rc_tot[ci].paragraphs[0].runs:
            run.bold = True
            run.font.size = Pt(8.2)

    # SEÇÕES 1, 2 e 3: TABELAS DETALHADAS COM AS 3 COLUNAS ESSENCIAIS PEDIDAS PELO USUÁRIO:
    # (1) O Que Está Errado no Documento | (2) Norma que Ampara + Citação Direta do Extrato | (3) Como Deve Ser o Texto Após a Correção
    for sec_num, (ek, elbl, _) in enumerate(eixos_cfg, start=1):
        sub = [a for a in achados if str(a.get("eixo") or "NORMATIVA").upper() == ek]
        if not sub:
            continue
        p_s = doc.add_paragraph()
        p_s.paragraph_format.space_before = Pt(10)
        rs = p_s.add_run(f"{elbl.upper()}")
        rs.bold = True
        rs.font.size = Pt(10)
        rs.font.color.rgb = RGBColor(30, 45, 61)

        tbl = doc.add_table(rows=1, cols=6)
        tbl.style = "Table Grid"
        headers = [
            "ID / Prioridade",
            "Documento / Item",
            "O Que Está Errado (Inconsistência Identificada)",
            "Norma de Amparo & Citação Direta (Extrato da Norma)",
            "Como Deve Ser o Texto Após a Correção & Providência",
            "Setor Responsável",
        ]
        for i, h in enumerate(headers):
            c = tbl.rows[0].cells[i]
            _set_cell_bg(c, "1E2D3D")
            c.text = h
            for run in c.paragraphs[0].runs:
                run.bold = True
                run.font.size = Pt(8)
                run.font.color.rgb = RGBColor(255, 255, 255)

        for idx, a in enumerate(sub, start=1):
            row = tbl.add_row().cells
            crit = str(a.get("criticidade") or "ALTA").upper()
            row[0].text = f"☐ {a.get('id') or f'A-{idx:02d}'}\n[{crit}]"
            row[1].text = str(a.get("localizacao") or "Autos")
            row[2].text = f"{a.get('titulo') or ''}\n\n{a.get('descricao_erro') or ''}".strip()

            base_l = str(a.get("base_legal") or "Lei nº 14.133/2021").strip()
            ext_l = str(a.get("extrato_norma") or "").strip()
            row[3].text = f"Norma: {base_l}\n\nCitação / Extrato:\n\"{ext_l.strip('\"')}\"" if ext_l else f"Norma: {base_l}"

            txt_corr = str(a.get("texto_apos_correcao") or a.get("trecho_sugerido") or "").strip()
            acao_corr = str(a.get("acao_corretiva") or "").strip()
            col4_parts = []
            if txt_corr:
                col4_parts.append(f"Redação Corrigida:\n\"{txt_corr}\"")
            if acao_corr and acao_corr != txt_corr:
                col4_parts.append(f"Providência:\n{acao_corr}")
            row[4].text = "\n\n".join(col4_parts) or "Promover o saneamento conforme fundamentação."
            row[5].text = str(a.get("setor_responsavel") or "Licitações")

            for c in row:
                for p_cell in c.paragraphs:
                    for run in p_cell.runs:
                        run.font.size = Pt(7.8)

        # Se for Seção 1 e houver Tabela 1.1 de Divergências de CATMAT
        if sec_num == 1 and div_catmat:
            p_cat = doc.add_paragraph()
            p_cat.paragraph_format.space_before = Pt(6)
            rcat = p_cat.add_run("Tabela 1.1 — Detalhamento das Divergências de Códigos CATMAT/CATSER e Unidades de Medida (DFD × ETP × TR)")
            rcat.bold = True
            rcat.font.size = Pt(9)
            rcat.font.color.rgb = RGBColor(153, 27, 27)

            t_cat = doc.add_table(rows=1, cols=6)
            t_cat.style = "Table Grid"
            for i, h in enumerate(["Item", "Descrição no ETP / TR", "CATMAT Atual", "Descrição Puxada no DFD (Catálogo)", "Unidade (DFD vs TR)", "Correção Necessária"]):
                c = t_cat.rows[0].cells[i]
                _set_cell_bg(c, "334960")
                c.text = h
                for run in c.paragraphs[0].runs:
                    run.bold = True
                    run.font.size = Pt(7.8)
                    run.font.color.rgb = RGBColor(255, 255, 255)
            for dc in div_catmat:
                if not isinstance(dc, dict):
                    continue
                rc = t_cat.add_row().cells
                rc[0].text = str(dc.get("item") or "")
                rc[1].text = str(dc.get("descricao_etp_tr") or "")
                rc[2].text = str(dc.get("catmat_atual") or "")
                rc[3].text = str(dc.get("descricao_dfd_catalogo") or "")
                rc[4].text = str(dc.get("unidade_dfd_vs_tr") or "")
                rc[5].text = str(dc.get("correcao_necessaria") or "")
                for c in rc:
                    for run in c.paragraphs[0].runs:
                        run.font.size = Pt(7.8)

    # 4. ROTEIRO SEQUENCIAL RECOMENDADO PARA SANEAMENTO
    if roteiro:
        p_r = doc.add_paragraph()
        p_r.paragraph_format.space_before = Pt(10)
        rr = p_r.add_run("4. ROTEIRO SEQUENCIAL RECOMENDADO PARA SANEAMENTO CRONOLÓGICO")
        rr.bold = True
        rr.font.size = Pt(10)
        rr.font.color.rgb = RGBColor(30, 45, 61)

        t_rot = doc.add_table(rows=1, cols=4)
        t_rot.style = "Table Grid"
        for i, h in enumerate(["Etapa", "Atividade de Saneamento", "Itens Atendidos", "Setor Executor"]):
            c = t_rot.rows[0].cells[i]
            _set_cell_bg(c, "1E2D3D")
            c.text = h
            for run in c.paragraphs[0].runs:
                run.bold = True
                run.font.size = Pt(8)
                run.font.color.rgb = RGBColor(255, 255, 255)
        for idx_r, step in enumerate(roteiro, start=1):
            rc = t_rot.add_row().cells
            if isinstance(step, dict):
                rc[0].text = f"☐ {step.get('etapa') or f'{idx_r}ª Etapa'}"
                rc[1].text = str(step.get("atividade") or "")
                rc[2].text = str(step.get("itens_atendidos") or "—")
                rc[3].text = str(step.get("setor_executor") or "Licitações")
            else:
                rc[0].text = f"☐ {idx_r}ª Etapa"
                rc[1].text = str(step)
                rc[2].text = "—"
                rc[3].text = "Setor Competente"
            for c in rc:
                for run in c.paragraphs[0].runs:
                    run.font.size = Pt(8)

    # 5. TERMO DE CONFERÊNCIA E ATESTE DE SANEAMENTO
    p_t = doc.add_paragraph()
    p_t.paragraph_format.space_before = Pt(10)
    rt = p_t.add_run("5. TERMO DE CONFERÊNCIA E ATESTE DE SANEAMENTO PRÉVIO À PUBLICAÇÃO NO PNCP")
    rt.bold = True
    rt.font.size = Pt(10)
    rt.font.color.rgb = RGBColor(30, 45, 61)

    t_sig = doc.add_table(rows=2, cols=3)
    t_sig.style = "Table Grid"
    sig_hdrs = [
        "SETOR REQUISITANTE\n(Responsável pela Demanda / ETP)",
        "SEÇÃO DE INTENDÊNCIA / LICITAÇÕES\n(Agente de Contratação)",
        "AUTORIDADE COMPETENTE\n(Ordenador de Despesas)",
    ]
    for i, sh in enumerate(sig_hdrs):
        c = t_sig.rows[0].cells[i]
        _set_cell_bg(c, "F1F5F9")
        c.text = sh
        c.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        for run in c.paragraphs[0].runs:
            run.bold = True
            run.font.size = Pt(8)
        c2 = t_sig.rows[1].cells[i]
        c2.text = "\n________________________________________\nNome / Posto-Graduação / Cargo\nData: ____ / ____ / 2026"
        c2.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        for run in c2.paragraphs[0].runs:
            run.font.size = Pt(8)

    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()


def build_despacho_saneamento_pdf_bytes(payload: Dict[str, Any]) -> bytes:
    """
    Gera o CHECKLIST DETALHADO DE CONFORMIDADE NORMATIVA E SANEAMENTO PROCESSUAL em PDF (Paisagem A4)
    para máxima legibilidade das colunas:
    [ID/Risco | Documento/Item | O Que Está Errado | Norma + Citação Direta | Como Deve Ser o Texto Após Correção | Responsável]
    """
    om_name = (payload.get("om_name") or "MARINHA DO BRASIL").strip()
    processo_ref = (payload.get("processo_ref") or payload.get("filename") or "Processo Administrativo").strip()
    meta = payload.get("metadados_processo") or {}
    parecer = payload.get("parecer_preventivo") or {}
    achados = payload.get("achados") or []
    div_catmat = payload.get("divergencias_catmat") or []
    roteiro = payload.get("roteiro_saneamento") or []
    now_str = datetime.now().strftime("%d/%m/%Y às %H:%M")

    buf = io.BytesIO()
    doc = SimpleDocTemplate(
        buf,
        pagesize=landscape(A4),
        leftMargin=1.2 * cm,
        rightMargin=1.2 * cm,
        topMargin=1.2 * cm,
        bottomMargin=1.2 * cm,
    )
    styles = getSampleStyleSheet()
    st_hdr = ParagraphStyle("H", parent=styles["Normal"], fontName="Helvetica-Bold", fontSize=11, alignment=TA_CENTER, textColor=colors.HexColor("#1e2d3d"))
    st_sub = ParagraphStyle("S", parent=styles["Normal"], fontName="Helvetica-Bold", fontSize=8.5, alignment=TA_CENTER, textColor=colors.HexColor("#9a7632"), spaceAfter=6)
    st_sec = ParagraphStyle("Sec", parent=styles["Normal"], fontName="Helvetica-Bold", fontSize=9.5, textColor=colors.HexColor("#1e2d3d"), spaceBefore=8, spaceAfter=4)
    st_body = ParagraphStyle("B", parent=styles["Normal"], fontName="Helvetica", fontSize=8, leading=11, alignment=TA_JUSTIFY, spaceAfter=4)
    st_cell = ParagraphStyle("C", parent=styles["Normal"], fontName="Helvetica", fontSize=7.2, leading=9.4)
    st_hdr_cell = ParagraphStyle("HC", parent=styles["Normal"], fontName="Helvetica-Bold", fontSize=7.5, leading=9.5, textColor=colors.white)

    veredito_raw = str(parecer.get("veredito") or "SOBRESTAR_PARA_SANEAMENTO").replace("_", " ")
    story: List[Any] = [
        Paragraph(om_name.upper(), st_hdr),
        Paragraph("CHECKLIST DETALHADO DE CONFORMIDADE NORMATIVA E SANEAMENTO PROCESSUAL — LEI Nº 14.133/2021", st_sub),
        Paragraph(
            f"<b>Processo:</b> {meta.get('numero_processo') or processo_ref} &nbsp;|&nbsp; "
            f"<b>Modalidade:</b> {meta.get('modalidade_fundamento') or 'Lei nº 14.133/2021'} &nbsp;|&nbsp; "
            f"<b>Valor Estimado:</b> {meta.get('valor_total_estimado') or 'Conforme autos'} &nbsp;|&nbsp; "
            f"<b>Data da Auditoria:</b> {now_str}<br/>"
            f"<b>PARECER PREVENTIVO ({veredito_raw}):</b> {parecer.get('resumo_executivo') or ''}",
            st_body,
        ),
    ]

    eixos_cfg = [
        ("NORMATIVA", "1. CHECKLIST DE INCONSISTÊNCIAS NORMATIVAS (LEI Nº 14.133/2021 E INS)"),
        ("PROCESSUAL", "2. CHECKLIST DE ERROS PROCESSUAIS E INSTRUÇÃO (ART. 72 DA LEI Nº 14.133/2021)"),
        ("CONTRADICAO_INTERNA", "3. CHECKLIST DE CONTRADIÇÕES INTERNAS E REVISÃO REDACIONAL"),
    ]

    for sec_idx, (ek, etitle) in enumerate(eixos_cfg, start=1):
        sub = [a for a in achados if str(a.get("eixo") or "NORMATIVA").upper() == ek]
        if not sub:
            continue
        story.append(Paragraph(f"{etitle} ({len(sub)} apontamento(s))", st_sec))
        tdata = [[
            Paragraph("ID / Risco", st_hdr_cell),
            Paragraph("Documento / Item", st_hdr_cell),
            Paragraph("O Que Está Errado no Documento", st_hdr_cell),
            Paragraph("Norma & Citação Direta (Extrato)", st_hdr_cell),
            Paragraph("Como Deve Ficar o Texto Após Correção", st_hdr_cell),
            Paragraph("Responsável", st_hdr_cell),
        ]]
        for idx, a in enumerate(sub, start=1):
            base_l = str(a.get("base_legal") or "Lei nº 14.133/2021").strip()
            ext_l = str(a.get("extrato_norma") or "").strip()
            norma_html = f"<b>{base_l}</b>" + (f"<br/><br/><i>\"{ext_l.strip('\"')}\"</i>" if ext_l else "")

            txt_corr = str(a.get("texto_apos_correcao") or a.get("trecho_sugerido") or "").strip()
            acao_corr = str(a.get("acao_corretiva") or "").strip()
            corr_html = ""
            if txt_corr:
                corr_html += f"<b>Redação Corrigida:</b><br/>\"{txt_corr}\""
            if acao_corr and acao_corr != txt_corr:
                corr_html += ("<br/><br/>" if corr_html else "") + f"<b>Providência:</b> {acao_corr}"

            tdata.append([
                Paragraph(f"☐ <b>{a.get('id') or f'A-{idx:02d}'}</b><br/>[{a.get('criticidade') or 'ALTA'}]", st_cell),
                Paragraph(str(a.get("localizacao") or "Autos"), st_cell),
                Paragraph(f"<b>{a.get('titulo') or ''}</b><br/>{a.get('descricao_erro') or ''}", st_cell),
                Paragraph(norma_html, st_cell),
                Paragraph(corr_html or "Sanear conforme norma.", st_cell),
                Paragraph(str(a.get("setor_responsavel") or "Licitações"), st_cell),
            ])
        t = Table(tdata, colWidths=[2.1 * cm, 3.2 * cm, 6.8 * cm, 6.8 * cm, 6.0 * cm, 2.4 * cm], repeatRows=1)
        t.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1e2d3d")),
            ("GRID", (0, 0), (-1, -1), 0.35, colors.HexColor("#cbd5e1")),
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#f8fafc")]),
        ]))
        story.append(t)

        if sec_idx == 1 and div_catmat:
            story.append(Paragraph("Tabela 1.1 — Detalhamento das Divergências de CATMAT e Unidades de Medida (DFD × ETP × TR)", st_sec))
            cdata = [[
                Paragraph("Item", st_hdr_cell),
                Paragraph("Descrição no ETP / TR", st_hdr_cell),
                Paragraph("CATMAT Atual", st_hdr_cell),
                Paragraph("Descrição Puxada no DFD", st_hdr_cell),
                Paragraph("Unidade (DFD vs TR)", st_hdr_cell),
                Paragraph("Correção Necessária", st_hdr_cell),
            ]]
            for dc in div_catmat:
                if not isinstance(dc, dict):
                    continue
                cdata.append([
                    Paragraph(str(dc.get("item") or ""), st_cell),
                    Paragraph(str(dc.get("descricao_etp_tr") or ""), st_cell),
                    Paragraph(str(dc.get("catmat_atual") or ""), st_cell),
                    Paragraph(str(dc.get("descricao_dfd_catalogo") or ""), st_cell),
                    Paragraph(str(dc.get("unidade_dfd_vs_tr") or ""), st_cell),
                    Paragraph(str(dc.get("correcao_necessaria") or ""), st_cell),
                ])
            tc = Table(cdata, colWidths=[1.4 * cm, 5.8 * cm, 2.5 * cm, 7.6 * cm, 3.2 * cm, 6.8 * cm], repeatRows=1)
            tc.setStyle(TableStyle([
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#334960")),
                ("GRID", (0, 0), (-1, -1), 0.35, colors.HexColor("#cbd5e1")),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#fff1f2")]),
            ]))
            story.append(tc)

    if roteiro:
        story.append(Spacer(1, 6))
        story.append(Paragraph("4. ROTEIRO SEQUENCIAL RECOMENDADO PARA SANEAMENTO", st_sec))
        rdata = [[
            Paragraph("Etapa", st_hdr_cell),
            Paragraph("Atividade de Saneamento Cronológica", st_hdr_cell),
            Paragraph("Itens Atendidos", st_hdr_cell),
            Paragraph("Setor Executor", st_hdr_cell),
        ]]
        for idx_r, step in enumerate(roteiro, start=1):
            if isinstance(step, dict):
                rdata.append([
                    Paragraph(f"☐ <b>{step.get('etapa') or f'{idx_r}ª Etapa'}</b>", st_cell),
                    Paragraph(str(step.get("atividade") or ""), st_cell),
                    Paragraph(str(step.get("itens_atendidos") or "—"), st_cell),
                    Paragraph(str(step.get("setor_executor") or "Licitações"), st_cell),
                ])
            else:
                rdata.append([
                    Paragraph(f"☐ <b>{idx_r}ª Etapa</b>", st_cell),
                    Paragraph(str(step), st_cell),
                    Paragraph("—", st_cell),
                    Paragraph("Setor Competente", st_cell),
                ])
        tr = Table(rdata, colWidths=[2.5 * cm, 17.5 * cm, 3.8 * cm, 3.5 * cm], repeatRows=1)
        tr.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1e2d3d")),
            ("GRID", (0, 0), (-1, -1), 0.35, colors.HexColor("#cbd5e1")),
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ]))
        story.append(tr)

    doc.build(story)
    return buf.getvalue()
