# -*- coding: utf-8 -*-
"""
MÓDULO DE ENGENHARIA DE DOCUMENTOS AUDITADOS (LEI Nº 14.133/2021 & AGU)
======================================================================
Preserva fielmente o formato, layout, hierarquia de seções, cabeçalhos,
tabelas e diagramação de documentos originais (PDF, DOCX, TXT),
gerando revisões jurídicas oficiais com controle de alterações visual:
- AMARELO: Itens corrigidos, inseridos ou adequados às normas da AGU/NLLC.
- VERMELHO COM TACHADO: Dispositivos revogados, obsoletos ou extirpados.
- PRETO: Texto original mantido íntegro.

Fornece exportadores duais de alta fidelidade:
- Word (.DOCX): Com tabelas nativas, parágrafos formatados e runs com realce/tachado.
- PDF (.PDF): Via ReportLab com Brasão oficial, cabeçalho da Marinha, tabelas zebradas e numeração de páginas.
"""

from __future__ import annotations
import difflib
import html
import io
import os
import re
from typing import Any, Dict, List, Optional, Tuple

import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_COLOR_INDEX
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
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
from reportlab.pdfgen import canvas
import pdfplumber

BASE_DIR = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))


# ==============================================================================
# 1. FUNÇÕES AUXILIARES DE DIFF E ESCAPE
# ==============================================================================
def xml_escape(text: str) -> str:
    """Escapa caracteres especiais XML para renderização segura no ReportLab."""
    if not text:
        return ""
    return (
        text.replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace('"', "&quot;")
    )


def compute_word_diff(original_text: str, revised_text: str) -> List[Dict[str, str]]:
    """
    Compara o texto original com a versão revisada no nível de palavras.
    Retorna uma lista de segmentos estruturados:
    - {'type': 'equal', 'text': '...'}
    - {'type': 'removed', 'text': '...'} (Vermelho / Tachado)
    - {'type': 'added', 'text': '...'}   (Amarelo / Negrito)
    """
    if not original_text and not revised_text:
        return []
    if not original_text:
        return [{"type": "added", "text": revised_text}]
    if not revised_text:
        return [{"type": "removed", "text": original_text}]
    if original_text.strip() == revised_text.strip():
        return [{"type": "equal", "text": original_text}]

    w_orig = original_text.split()
    w_rev = revised_text.split()

    matcher = difflib.SequenceMatcher(None, w_orig, w_rev)
    segments: List[Dict[str, str]] = []

    for tag, i1, i2, j1, j2 in matcher.get_opcodes():
        if tag == "equal":
            txt = " ".join(w_orig[i1:i2])
            if txt:
                segments.append({"type": "equal", "text": txt})
        elif tag == "delete":
            txt = " ".join(w_orig[i1:i2])
            if txt:
                segments.append({"type": "removed", "text": txt})
        elif tag == "insert":
            txt = " ".join(w_rev[j1:j2])
            if txt:
                segments.append({"type": "added", "text": txt})
        elif tag == "replace":
            txt_del = " ".join(w_orig[i1:i2])
            txt_ins = " ".join(w_rev[j1:j2])
            if txt_del:
                segments.append({"type": "removed", "text": txt_del})
            if txt_ins:
                segments.append({"type": "added", "text": txt_ins})

    return segments


def set_docx_cell_background(cell, hex_color: str):
    """Aplica cor de fundo a uma célula de tabela Word via XML nativo."""
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_color.replace("#", "").upper())
    tcPr.append(shd)


def set_docx_cell_margins(cell, top=100, bottom=100, left=140, right=140):
    """Define margens internas de célula Word."""
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement("w:tcMar")
    for m, val in [("top", top), ("bottom", bottom), ("left", left), ("right", right)]:
        node = OxmlElement(f"w:{m}")
        node.set(qn("w:w"), str(val))
        node.set(qn("w:type"), "dxa")
        tcMar.append(node)
    tcPr.append(tcMar)


# ==============================================================================
# 2. PARSER DE DOCUMENTOS COM PRESERVAÇÃO DE DIAGRAMAÇÃO
# ==============================================================================
class DocumentLayoutParser:
    """
    Extrai o layout integral de documentos PDF, DOCX e TXT preservando
    seções, subcláusulas, parágrafos, tabelas, metadados e assinaturas.
    """

    @classmethod
    def parse_pdf(cls, file_bytes: bytes, filename: str = "") -> Dict[str, Any]:
        doc_structure: Dict[str, Any] = {
            "original_format": "pdf",
            "filename": filename,
            "doc_type": "TR",
            "title": "TERMO DE REFERÊNCIA (COMPRAS.GOV.BR)",
            "uasg": "782340",
            "orgao": "CAPITANIA FLUVIAL DE JUAZEIRO",
            "processo": "",
            "metadata_box": {},
            "elements": [],
            "signatures": [],
        }

        with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
            raw_page_items = []
            for pno, page in enumerate(pdf.pages):
                tables = page.find_tables()
                t_bboxes = [t.bbox for t in tables]

                # Extrai palavras fora das tabelas com coordenadas verticais
                words = page.extract_words()
                outside_words = [
                    w for w in words
                    if not any(
                        b[0] - 2 <= w["x0"] <= b[2] + 2 and b[1] - 2 <= w["top"] <= b[3] + 2
                        for b in t_bboxes
                    )
                ]

                lines_with_pos = []
                if outside_words:
                    outside_words.sort(key=lambda w: (w["top"], w["x0"]))
                    curr_line = [outside_words[0]]
                    for w in outside_words[1:]:
                        if abs(w["top"] - curr_line[-1]["top"]) < 3.5:
                            curr_line.append(w)
                        else:
                            text_str = " ".join(x["text"] for x in curr_line).strip()
                            if text_str:
                                lines_with_pos.append({"top": curr_line[0]["top"], "text": text_str})
                            curr_line = [w]
                    if curr_line:
                        text_str = " ".join(x["text"] for x in curr_line).strip()
                        if text_str:
                            lines_with_pos.append({"top": curr_line[0]["top"], "text": text_str})

                # Agrupa elementos da página ordenados por 'top'
                page_items = []
                for l in lines_with_pos:
                    page_items.append(("TEXT", l["top"], l["text"]))
                for t in tables:
                    t_data = t.extract()
                    clean_rows = [[(c or "").strip() for c in r] for r in t_data if any(r)]
                    if clean_rows:
                        page_items.append(("TABLE", t.bbox[1], clean_rows))

                page_items.sort(key=lambda x: x[1])
                for it in page_items:
                    raw_page_items.append((pno + 1, it[0], it[2]))

        # Interpretação estrutural do documento
        elements: List[Dict[str, Any]] = []
        current_clause: Optional[Dict[str, Any]] = None
        current_section_num = "0"
        current_section_title = "PREÂMBULO / INFORMAÇÕES GERAIS"
        detected_signatures: List[Dict[str, str]] = []

        for pno, it_type, val in raw_page_items:
            if it_type == "TABLE":
                if current_clause:
                    elements.append(current_clause)
                    current_clause = None
                elements.append({
                    "type": "TABLE",
                    "page": pno,
                    "section_num": current_section_num,
                    "headers": val[0] if len(val) > 0 else [],
                    "rows": val[1:] if len(val) > 1 else [],
                })
                continue

            line = str(val).strip()
            if not line:
                continue

            # Rodapés de paginação tipo "1 de 6" ou identificadores de UASG isolados
            if re.search(r"^\d+\s+de\s+\d+$", line):
                continue
            if re.search(r"^UASG\s*(\d+)$", line):
                doc_structure["uasg"] = re.search(r"^UASG\s*(\d+)$", line).group(1)
                continue
            if line in ["CAPITANIA FLUVIAL DE JUAZEIRO", "MARINHA DO BRASIL"]:
                doc_structure["orgao"] = line
                continue

            # Título do Documento
            if re.match(r"^(Estudo Técnico Preliminar|Termo de Referência|Documento de Formalização da Demanda)\s*([\d/]+)?", line, re.I):
                doc_structure["title"] = line
                if "Estudo Técnico" in line:
                    doc_structure["doc_type"] = "ETP"
                elif "Termo de Referência" in line:
                    doc_structure["doc_type"] = "TR"
                elif "Formalização" in line:
                    doc_structure["doc_type"] = "DFD"
                continue

            if "Número do processo:" in line:
                doc_structure["processo"] = line.split(":", 1)[1].strip()

            # Detecção de Assinaturas (Nomes em caixa alta seguidos de cargo)
            if re.match(r"^(Membro da comissão|Responsável|Ordenador|Fiscal|Chefe|Encarregado)", line, re.I):
                prev_text = ""
                if current_clause and current_clause["text"]:
                    words = current_clause["text"].split()
                    if words:
                        prev_text = words[-1]
                detected_signatures.append({"role": line, "name": prev_text or "RESPONSÁVEL PELA MINUTA"})
                continue

            # Seção Principal (Heading 1): Ex: "1. Informações Básicas" ou "2. Descrição da necessidade"
            sec_m = re.match(r"^(\d{1,2})\.\s+([A-ZÁÉÍÓÚÂÊÔÃÕÇ\w\s\-/()]{3,})$", line)
            if sec_m:
                if current_clause:
                    elements.append(current_clause)
                    current_clause = None
                current_section_num = sec_m.group(1)
                current_section_title = sec_m.group(2).strip()
                elements.append({
                    "type": "HEADING_1",
                    "number": current_section_num,
                    "title": current_section_title,
                    "page": pno,
                })
                continue

            # Subcláusula numerada: Ex: "2.1. Contratação de serviços..." ou "7.2.1."
            cl_m = re.match(r"^(\d{1,2}(?:\.\d{1,3})+)\.?\s*(.*)$", line)
            if cl_m:
                if current_clause:
                    elements.append(current_clause)
                current_clause = {
                    "type": "CLAUSE",
                    "clause_number": cl_m.group(1),
                    "section_num": current_section_num,
                    "section_title": current_section_title,
                    "text": cl_m.group(2).strip(),
                    "page": pno,
                    "modified": False,
                    "diff_segments": [],
                }
                continue

            # Linha de continuação de subcláusula ou parágrafo livre
            if current_clause:
                current_clause["text"] = (current_clause["text"] + " " + line).strip()
            else:
                # Subtítulos textuais sem numeração
                if len(line) < 70 and not line.endswith(".") and line.isupper():
                    elements.append({
                        "type": "HEADING_2",
                        "title": line,
                        "section_num": current_section_num,
                        "page": pno,
                    })
                else:
                    elements.append({
                        "type": "PARAGRAPH",
                        "text": line,
                        "section_num": current_section_num,
                        "page": pno,
                        "modified": False,
                        "diff_segments": [],
                    })

        if current_clause:
            elements.append(current_clause)

        doc_structure["elements"] = elements
        doc_structure["signatures"] = detected_signatures
        return doc_structure

    @classmethod
    def parse_docx(cls, file_bytes: bytes, filename: str = "") -> Dict[str, Any]:
        """Extrai parágrafos, seções e tabelas nativas de arquivo .DOCX."""
        doc = docx.Document(io.BytesIO(file_bytes))
        doc_structure: Dict[str, Any] = {
            "original_format": "docx",
            "filename": filename,
            "doc_type": "TR",
            "title": "TERMO DE REFERÊNCIA (WORD)",
            "uasg": "782340",
            "orgao": "CAPITANIA FLUVIAL DE JUAZEIRO",
            "processo": "",
            "metadata_box": {},
            "elements": [],
            "signatures": [],
        }

        elements: List[Dict[str, Any]] = []
        current_clause: Optional[Dict[str, Any]] = None
        current_section_num = "0"
        current_section_title = "PREÂMBULO"

        for block in doc.element.body:
            if block.tag.endswith("p"):
                p = docx.text.paragraph.Paragraph(block, doc)
                line = p.text.strip()
                if not line:
                    continue

                if re.match(r"^(Estudo Técnico|Termo de Referência|Documento de Formalização)", line, re.I):
                    doc_structure["title"] = line
                    continue

                sec_m = re.match(r"^(\d{1,2})\.\s+([A-ZÁÉÍÓÚÂÊÔÃÕÇ\w\s\-/()]{3,})$", line)
                if sec_m:
                    if current_clause:
                        elements.append(current_clause)
                        current_clause = None
                    current_section_num = sec_m.group(1)
                    current_section_title = sec_m.group(2).strip()
                    elements.append({
                        "type": "HEADING_1",
                        "number": current_section_num,
                        "title": current_section_title,
                        "page": 1,
                    })
                    continue

                cl_m = re.match(r"^(\d{1,2}(?:\.\d{1,3})+)\.?\s*(.*)$", line)
                if cl_m:
                    if current_clause:
                        elements.append(current_clause)
                    current_clause = {
                        "type": "CLAUSE",
                        "clause_number": cl_m.group(1),
                        "section_num": current_section_num,
                        "section_title": current_section_title,
                        "text": cl_m.group(2).strip(),
                        "page": 1,
                        "modified": False,
                        "diff_segments": [],
                    }
                    continue

                if current_clause:
                    current_clause["text"] = (current_clause["text"] + " " + line).strip()
                else:
                    elements.append({
                        "type": "PARAGRAPH",
                        "text": line,
                        "section_num": current_section_num,
                        "page": 1,
                        "modified": False,
                        "diff_segments": [],
                    })

            elif block.tag.endswith("tbl"):
                tbl = docx.table.Table(block, doc)
                rows_data = []
                for row in tbl.rows:
                    r_text = [cell.text.strip() for cell in row.cells]
                    if any(r_text):
                        rows_data.append(r_text)
                if rows_data:
                    if current_clause:
                        elements.append(current_clause)
                        current_clause = None
                    elements.append({
                        "type": "TABLE",
                        "page": 1,
                        "section_num": current_section_num,
                        "headers": rows_data[0],
                        "rows": rows_data[1:] if len(rows_data) > 1 else [],
                    })

        if current_clause:
            elements.append(current_clause)

        doc_structure["elements"] = elements
        return doc_structure

    @classmethod
    def parse_text(cls, raw_text: str, filename: str = "minuta_texto.txt") -> Dict[str, Any]:
        """Extrai seções, subcláusulas e parágrafos de texto puro digitado ou colado."""
        doc_structure: Dict[str, Any] = {
            "original_format": "txt",
            "filename": filename,
            "doc_type": "TR",
            "title": "TERMO DE REFERÊNCIA / MINUTA JURÍDICA",
            "uasg": "782340",
            "orgao": "CAPITANIA FLUVIAL DE JUAZEIRO",
            "processo": "",
            "metadata_box": {},
            "elements": [],
            "signatures": [],
        }

        elements: List[Dict[str, Any]] = []
        current_clause: Optional[Dict[str, Any]] = None
        current_section_num = "0"
        current_section_title = "PREÂMBULO"

        lines = [l.strip() for l in raw_text.split("\n") if l.strip()]
        for line in lines:
            if re.match(r"^(Estudo Técnico|Termo de Referência|Documento de Formalização)", line, re.I):
                doc_structure["title"] = line
                continue

            sec_m = re.match(r"^(\d{1,2})\.\s+([A-ZÁÉÍÓÚÂÊÔÃÕÇ\w\s\-/()]{3,})$", line)
            if sec_m:
                if current_clause:
                    elements.append(current_clause)
                    current_clause = None
                current_section_num = sec_m.group(1)
                current_section_title = sec_m.group(2).strip()
                elements.append({
                    "type": "HEADING_1",
                    "number": current_section_num,
                    "title": current_section_title,
                    "page": 1,
                })
                continue

            cl_m = re.match(r"^(\d{1,2}(?:\.\d{1,3})+)\.?\s*(.*)$", line)
            if cl_m:
                if current_clause:
                    elements.append(current_clause)
                current_clause = {
                    "type": "CLAUSE",
                    "clause_number": cl_m.group(1),
                    "section_num": current_section_num,
                    "section_title": current_section_title,
                    "text": cl_m.group(2).strip(),
                    "page": 1,
                    "modified": False,
                    "diff_segments": [],
                }
                continue

            if current_clause:
                current_clause["text"] = (current_clause["text"] + " " + line).strip()
            else:
                elements.append({
                    "type": "PARAGRAPH",
                    "text": line,
                    "section_num": current_section_num,
                    "page": 1,
                    "modified": False,
                    "diff_segments": [],
                })

        if current_clause:
            elements.append(current_clause)

        doc_structure["elements"] = elements
        return doc_structure



# ==============================================================================
# 3. MOTOR DE AUDITORIA JURÍDICA E TRACK CHANGES (AMARELO / VERMELHO TACHADO)
# ==============================================================================
class DocumentLegalAuditor:
    """
    Realiza a auditoria de conformidade com a Lei nº 14.133/2021 e Modelos da AGU:
    1. Identifica e extirpa referências à legislação revogada (Lei 8.666/93, 10.520/02, Decretos antigos).
    2. Substitui sanções inválidas (ex: suspensão de 2 anos -> impedimento de até 3 anos do Art. 156, III).
    3. Injeta cláusulas obrigatórias omissas (Gestão/Fiscalização Art. 117, Critérios de Sustentabilidade Art. 18, XI).
    4. Gera o diff com marcação exata em Amarelo (adicionado) e Vermelho com Tachado (removido).
    """

    RULES = [
        # Lei 8.666/1993 e 10.520/2002
        (
            r"\b(?:Lei\s+(?:Federal\s+)?n[ºo°]?\s*8\.?666(?:/1993|/93)?|Lei\s+8\.666)\b",
            "Lei nº 14.133/2021",
            "Lei nº 8.666/1993 totalmente revogada em 30/12/2023. Substituída pela Nova Lei de Licitações (Lei nº 14.133/2021).",
        ),
        (
            r"\b(?:Lei\s+(?:Federal\s+)?n[ºo°]?\s*10\.?520(?:/2002|/02)?|Lei\s+10\.520)\b",
            "Lei nº 14.133/2021",
            "Lei nº 10.520/2002 revogada. O pregão agora é disciplinado diretamente pelo Art. 28, I da Lei nº 14.133/2021.",
        ),
        (
            r"\b(?:Decreto\s+n[ºo°]?\s*10\.?024(?:/2019)?|Decreto\s+10\.024)\b",
            "Decreto nº 11.462/2023",
            "Decreto nº 10.024/2019 revogado no âmbito federal pelo Decreto nº 11.462/2023 (Regulamento de Licitação e SRP).",
        ),
        (
            r"\b(?:Decreto\s+n[ºo°]?\s*3\.?555(?:/2000)?|Decreto\s+n[ºo°]?\s*5\.?450(?:/2005)?)\b",
            "Decreto nº 11.462/2023",
            "Decretos federais revogados e substituídos pelo Decreto nº 11.462/2023.",
        ),
        (
            r"\bsuspens[ãa]o\s+tempor[áa]ria\s+de\s+participa[çc][ãa]o\s+em\s+licita[çc][ãa]o\b.*?(?:2|dois)\s+anos",
            "impedimento de licitar e contratar no âmbito da Administração Pública federal pelo prazo de até 3 (três) anos (Art. 156, inciso III da Lei nº 14.133/2021)",
            "A sanção de suspensão temporária por 2 anos da Lei 8.666 foi extinta. O Art. 156, III prevê impedimento por até 3 anos.",
        ),
        (
            r"\bdeclara[çc][ãa]o\s+de\s+inidoneidade\b.*?(?:2|dois)\s+anos",
            "declaração de inidoneidade para licitar ou contratar pelo prazo mínimo de 3 (três) anos e máximo de 6 (seis) anos (Art. 156, inciso IV da Lei nº 14.133/2021)",
            "Sob a Lei nº 14.133/2021, o prazo da declaração de inidoneidade é de 3 a 6 anos (Art. 156, § 5º).",
        ),
        (
            r"\bart(?:igo)?\.?\s*24,?\s*inciso\s*II\s+da\s+Lei\s*8\.?666\b",
            "art. 75, inciso II da Lei nº 14.133/2021",
            "Dispensa de licitação em razão do valor agora regida pelo Art. 75, inciso II da Lei nº 14.133/2021.",
        ),
        (
            r"\bart(?:igo)?\.?\s*24,?\s*inciso\s*I\s+da\s+Lei\s*8\.?666\b",
            "art. 75, inciso I da Lei nº 14.133/2021",
            "Dispensa para obras e serviços de engenharia agora regida pelo Art. 75, inciso I da Lei nº 14.133/2021.",
        ),
    ]

    @classmethod
    def audit_and_diff(cls, doc_structure: Dict[str, Any]) -> Dict[str, Any]:
        elements = doc_structure.get("elements", [])
        findings = []
        total_clauses = 0
        modified_clauses = 0

        for el in elements:
            if el.get("type") in ["CLAUSE", "PARAGRAPH"]:
                total_clauses += 1
                orig_text = el.get("text", "")
                revised_text = orig_text
                clause_findings = []

                for pattern, replacement, reason in cls.RULES:
                    if re.search(pattern, revised_text, flags=re.IGNORECASE):
                        revised_text = re.sub(pattern, replacement, revised_text, flags=re.IGNORECASE)
                        clause_findings.append({
                            "type": "VULNERABILIDADE_JURIDICA",
                            "severity": "CRITICA",
                            "reason": reason,
                            "original_clause": el.get("clause_number") or "Parágrafo",
                        })

                if revised_text != orig_text:
                    el["modified"] = True
                    el["revised_text"] = revised_text
                    el["diff_segments"] = compute_word_diff(orig_text, revised_text)
                    modified_clauses += 1
                    findings.extend(clause_findings)
                else:
                    el["diff_segments"] = [{"type": "equal", "text": orig_text}]

        # Verificação de Cláusulas Obrigatórias Ausentes
        all_text = " ".join(e.get("text", "") for e in elements)
        omissions = []

        # 1. Gestão e Fiscalização do Contrato (Art. 117 da Lei 14.133/2021)
        if not re.search(r"\b(?:fiscal\s+do\s+contrato|gestor\s+do\s+contrato|fiscaliza[çc][ãa]o\s+técnica)\b", all_text, re.I):
            omissions.append({
                "title": "MODELO DE GESTÃO E FISCALIZAÇÃO DO CONTRATO",
                "clause_num": "GESTÃO CONTRATUAL (ART. 117)",
                "text": "A execução do contrato deverá ser acompanhada e fiscalizada por 1 (um) ou mais fiscais do contrato, representantes da Administração especialmente designados conforme os arts. 117 da Lei nº 14.133/2021 e do Decreto nº 11.246/2022, permitida a contratação de terceiros para assisti-los e subsidiá-los com informações pertinentes a essa atribuição.",
                "reason": "Exigência expressa do Art. 40, § 1º, inciso VII e Art. 117 da Lei nº 14.133/2021.",
            })

        # 2. Critérios de Sustentabilidade Ambiental (Art. 18, XI da Lei 14.133/2021)
        if not re.search(r"\b(?:sustentabilidade|impactos?\s+ambientais?|rotulagem\s+ambiental|ibama)\b", all_text, re.I):
            omissions.append({
                "title": "CRITÉRIOS DE SUSTENTABILIDADE AMBIENTAL",
                "clause_num": "SUSTENTABILIDADE (ART. 18, XI)",
                "text": "O fornecedor deverá atender aos critérios e práticas de sustentabilidade ambiental na execução contratual e fabricação dos bens, observando a redução de resíduos, descarte adequado de embalagens e conformidade com as normas ambientais vigentes (Instrução Normativa IBAMA e Guias de Contratações Sustentáveis da AGU).",
                "reason": "Obrigatório segundo o Art. 18, inciso XI da Lei nº 14.133/2021 e Guia Nacional de Contratações Sustentáveis da AGU.",
            })

        for om in omissions:
            elements.append({
                "type": "HEADING_1",
                "number": "+",
                "title": om["title"],
                "page": 1,
            })
            elements.append({
                "type": "CLAUSE",
                "clause_number": om["clause_num"],
                "text": om["text"],
                "modified": True,
                "diff_segments": [{"type": "added", "text": om["text"]}],
                "page": 1,
            })
            findings.append({
                "type": "OMISSAO_OBRIGATORIA",
                "severity": "ALTA",
                "reason": om["reason"],
                "original_clause": om["clause_num"],
            })
            modified_clauses += 1

        compliance_score = max(0.0, 100.0 - (len(findings) * 12.0))

        doc_structure["audit_summary"] = {
            "total_clauses": total_clauses,
            "modified_clauses": modified_clauses,
            "findings_count": len(findings),
            "compliance_score": compliance_score,
            "findings": findings,
        }
        return doc_structure


# ==============================================================================
# 4. GERADOR OFICIAL WORD (.DOCX) COM DIAGRAMAÇÃO INTEGRAL E TRACK CHANGES
# ==============================================================================
def build_audited_document_docx_bytes(doc_structure: Dict[str, Any], org_name: str = "MARINHA DO BRASIL") -> bytes:
    doc = docx.Document()

    # Margens Padrão ABNT / Governo Federal (2 cm / 0.8 in)
    for section in doc.sections:
        section.top_margin = Inches(0.8)
        section.bottom_margin = Inches(0.8)
        section.left_margin = Inches(0.8)
        section.right_margin = Inches(0.8)

    # 1. Brasão Oficial da República
    brasao_path = os.path.join(BASE_DIR, "brasao_oficial.png")
    if os.path.exists(brasao_path):
        p_img = doc.add_paragraph()
        p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_img.paragraph_format.space_before = Pt(0)
        p_img.paragraph_format.space_after = Pt(2)
        run_img = p_img.add_run()
        run_img.add_picture(brasao_path, width=Inches(1.05))

    # 2. Cabeçalho Oficial
    p_head = doc.add_paragraph()
    p_head.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_head.paragraph_format.line_spacing = 1.05
    p_head.paragraph_format.space_after = Pt(6)
    r_head = p_head.add_run(
        f"REPÚBLICA FEDERATIVA DO BRASIL\nMINISTÉRIO DA DEFESA\nCOMANDO DA MARINHA\n{doc_structure.get('orgao') or org_name}\n"
    )
    r_head.font.name = "Calibri"
    r_head.font.size = Pt(10)
    r_head.font.bold = True
    r_head.font.color.rgb = RGBColor(30, 45, 61)

    # 3. Título do Documento
    p_tit = doc.add_paragraph()
    p_tit.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_tit.paragraph_format.space_after = Pt(8)
    r_tit = p_tit.add_run(doc_structure.get("title", "DOCUMENTO ADMINISTRATIVO AUDITADO") + "\n")
    r_tit.font.name = "Calibri"
    r_tit.font.size = Pt(12)
    r_tit.font.bold = True
    r_tit.font.color.rgb = RGBColor(30, 45, 61)

    if doc_structure.get("processo"):
        r_proc = p_tit.add_run(f"Processo Administrativo: {doc_structure.get('processo')}\n")
        r_proc.font.name = "Calibri"
        r_proc.font.size = Pt(9.5)
        r_proc.font.italic = True

    # 4. Legenda e Convenção de Auditoria (Caixa Destacada)
    p_leg = doc.add_paragraph()
    p_leg.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p_leg.paragraph_format.space_before = Pt(4)
    p_leg.paragraph_format.space_after = Pt(12)
    r_leg_title = p_leg.add_run("CONVENÇÃO DE AUDITORIA JURÍDICA (LEI Nº 14.133/2021 & AGU):\n")
    r_leg_title.font.name = "Calibri"
    r_leg_title.font.size = Pt(9)
    r_leg_title.font.bold = True
    r_leg_title.font.color.rgb = RGBColor(15, 23, 42)

    r_leg_yel = p_leg.add_run("[AMARELO COM NEGRITO]")
    r_leg_yel.font.name = "Calibri"
    r_leg_yel.font.size = Pt(9)
    r_leg_yel.font.bold = True
    r_leg_yel.font.highlight_color = WD_COLOR_INDEX.YELLOW

    r_leg_t1 = p_leg.add_run(" = Textos inseridos, saneados ou adequados às minutas padronizadas da AGU.   ")
    r_leg_t1.font.name = "Calibri"
    r_leg_t1.font.size = Pt(9)

    r_leg_red = p_leg.add_run("[VERMELHO TACHADO]")
    r_leg_red.font.name = "Calibri"
    r_leg_red.font.size = Pt(9)
    r_leg_red.font.strike = True
    r_leg_red.font.color.rgb = RGBColor(220, 38, 38)

    r_leg_t2 = p_leg.add_run(" = Dispositivos revogados, normas obsoletas (Lei 8.666/93) ou extirpadas da minuta.")
    r_leg_t2.font.name = "Calibri"
    r_leg_t2.font.size = Pt(9)

    # 5. Renderização dos Elementos na Ordem Original
    for el in doc_structure.get("elements", []):
        el_type = el.get("type")

        # Heading 1 (Seção Principal)
        if el_type == "HEADING_1":
            p = doc.add_paragraph()
            p.paragraph_format.space_before = Pt(12)
            p.paragraph_format.space_after = Pt(4)
            p.paragraph_format.keep_with_next = True
            sec_num = el.get("number", "")
            sec_title = el.get("title", "")
            r = p.add_run(f"{sec_num}. {sec_title}" if sec_num and sec_num != "+" else sec_title)
            r.font.name = "Calibri"
            r.font.size = Pt(11)
            r.font.bold = True
            r.font.color.rgb = RGBColor(30, 45, 61)

        # Heading 2 (Subseção Livre)
        elif el_type == "HEADING_2":
            p = doc.add_paragraph()
            p.paragraph_format.space_before = Pt(8)
            p.paragraph_format.space_after = Pt(3)
            p.paragraph_format.keep_with_next = True
            r = p.add_run(el.get("title", ""))
            r.font.name = "Calibri"
            r.font.size = Pt(10)
            r.font.bold = True
            r.font.color.rgb = RGBColor(51, 65, 85)

        # Cláusula Numerada
        elif el_type == "CLAUSE":
            p = doc.add_paragraph()
            p.paragraph_format.space_before = Pt(2)
            p.paragraph_format.space_after = Pt(5)
            p.paragraph_format.line_spacing = 1.15

            # Número da subcláusula
            cl_num = el.get("clause_number", "")
            if cl_num:
                r_num = p.add_run(f"{cl_num}. ")
                r_num.font.name = "Calibri"
                r_num.font.size = Pt(10)
                r_num.font.bold = True
                if el.get("modified"):
                    r_num.font.highlight_color = WD_COLOR_INDEX.YELLOW

            diffs = el.get("diff_segments") or [{"type": "equal", "text": el.get("text", "")}]
            for seg in diffs:
                st = seg.get("type")
                txt = seg.get("text", "") + " "
                r = p.add_run(txt)
                r.font.name = "Calibri"
                r.font.size = Pt(10)
                if st == "added":
                    r.font.bold = True
                    r.font.highlight_color = WD_COLOR_INDEX.YELLOW
                elif st == "removed":
                    r.font.strike = True
                    r.font.color.rgb = RGBColor(220, 38, 38)

        # Parágrafo Livre
        elif el_type == "PARAGRAPH":
            p = doc.add_paragraph()
            p.paragraph_format.space_before = Pt(2)
            p.paragraph_format.space_after = Pt(4)
            p.paragraph_format.line_spacing = 1.15
            diffs = el.get("diff_segments") or [{"type": "equal", "text": el.get("text", "")}]
            for seg in diffs:
                st = seg.get("type")
                txt = seg.get("text", "") + " "
                r = p.add_run(txt)
                r.font.name = "Calibri"
                r.font.size = Pt(10)
                if st == "added":
                    r.font.bold = True
                    r.font.highlight_color = WD_COLOR_INDEX.YELLOW
                elif st == "removed":
                    r.font.strike = True
                    r.font.color.rgb = RGBColor(220, 38, 38)

        # Tabela Nativa Preservada
        elif el_type == "TABLE":
            headers = el.get("headers", [])
            rows = el.get("rows", [])
            if not headers and not rows:
                continue

            num_cols = max(len(headers), max((len(r) for r in rows), default=0))
            if num_cols == 0:
                continue

            tbl = doc.add_table(rows=len(rows) + (1 if headers else 0), cols=num_cols)
            tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
            tbl.autofit = True

            curr_row = 0
            if headers:
                for col_idx, h in enumerate(headers[:num_cols]):
                    cell = tbl.cell(curr_row, col_idx)
                    cell.text = str(h)
                    set_docx_cell_background(cell, "1E2D3D")
                    set_docx_cell_margins(cell, 80, 80, 120, 120)
                    for r in cell.paragraphs[0].runs:
                        r.font.name = "Calibri"
                        r.font.size = Pt(9)
                        r.font.bold = True
                        r.font.color.rgb = RGBColor(255, 255, 255)
                curr_row += 1

            for row_idx, r_data in enumerate(rows):
                fill_color = "F8FAFC" if row_idx % 2 == 1 else "FFFFFF"
                for col_idx, val in enumerate(r_data[:num_cols]):
                    cell = tbl.cell(curr_row, col_idx)
                    cell.text = str(val)
                    if fill_color != "FFFFFF":
                        set_docx_cell_background(cell, fill_color)
                    set_docx_cell_margins(cell, 60, 60, 100, 100)
                    for r in cell.paragraphs[0].runs:
                        r.font.name = "Calibri"
                        r.font.size = Pt(8.5)
                        r.font.color.rgb = RGBColor(30, 41, 59)
                curr_row += 1

            doc.add_paragraph().paragraph_format.space_after = Pt(4)

    # 6. Assinaturas
    signatures = doc_structure.get("signatures", [])
    if signatures:
        p_sig_title = doc.add_paragraph()
        p_sig_title.paragraph_format.space_before = Pt(20)
        p_sig_title.paragraph_format.space_after = Pt(12)
        p_sig_title.paragraph_format.keep_with_next = True
        r_sig_t = p_sig_title.add_run("ASSINATURAS DOS RESPONSÁVEIS:")
        r_sig_t.font.name = "Calibri"
        r_sig_t.font.size = Pt(10)
        r_sig_t.font.bold = True

        for sig in signatures:
            p_s = doc.add_paragraph()
            p_s.paragraph_format.space_before = Pt(6)
            p_s.paragraph_format.space_after = Pt(2)
            p_s.paragraph_format.line_spacing = 1.05
            p_s.paragraph_format.keep_with_next = True
            r_line = p_s.add_run("_____________________________________________________\n")
            r_line.font.name = "Calibri"
            r_name = p_s.add_run(f"{sig.get('name', 'RESPONSÁVEL TÉCNICO')}\n")
            r_name.font.name = "Calibri"
            r_name.font.size = Pt(9.5)
            r_name.font.bold = True
            r_role = p_s.add_run(sig.get("role", "Membro da Equipe de Planejamento"))
            r_role.font.name = "Calibri"
            r_role.font.size = Pt(8.5)
            r_role.font.italic = True

    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()


# ==============================================================================
# 5. GERADOR OFICIAL PDF (.PDF) COM REPORTLAB E NUMERAÇÃO DE PÁGINAS
# ==============================================================================
class NumberedCanvas(canvas.Canvas):
    """Dois passos de renderização para imprimir 'Página X de Y' e cabeçalho formal."""

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self._saved_page_states = []

    def showPage(self):
        self._saved_page_states.append(dict(self.__dict__))
        self._startPage()

    def save(self):
        num_pages = len(self._saved_page_states)
        for state in self._saved_page_states:
            self.__dict__.update(state)
            self.draw_page_decorations(num_pages)
            super().showPage()
        super().save()

    def draw_page_decorations(self, total_pages: int):
        self.saveState()
        self.setFont("Helvetica", 8)
        self.setFillColor(colors.HexColor("#64748B"))

        # Cabeçalho a partir da página 2
        if self._pageNumber > 1:
            self.drawString(54, 800, "MINUTA ADMINISTRATIVA AUDITADA — LEI Nº 14.133/2021 & AGU")
            self.setStrokeColor(colors.HexColor("#CBD5E1"))
            self.setLineWidth(0.5)
            self.line(54, 794, 541, 794)

        # Rodapé em todas as páginas
        page_str = f"Página {self._pageNumber} de {total_pages}"
        self.drawRightString(541, 36, page_str)
        self.drawString(54, 36, "Documento Saneado e Auditado Eletronicamente — Padrão AGU / TCU")
        self.setStrokeColor(colors.HexColor("#CBD5E1"))
        self.setLineWidth(0.5)
        self.line(54, 46, 541, 46)
        self.restoreState()


def build_audited_document_pdf_bytes(doc_structure: Dict[str, Any], org_name: str = "MARINHA DO BRASIL") -> bytes:
    buf = io.BytesIO()
    doc = SimpleDocTemplate(
        buf,
        pagesize=A4,
        leftMargin=54,
        rightMargin=54,
        topMargin=54,
        bottomMargin=54,
    )

    styles = getSampleStyleSheet()

    style_head = ParagraphStyle(
        "GovHead",
        parent=styles["Normal"],
        fontName="Helvetica-Bold",
        fontSize=9,
        leading=11,
        alignment=1,
        textColor=colors.HexColor("#1E2D3D"),
        spaceAfter=4,
    )

    style_title = ParagraphStyle(
        "GovTitle",
        parent=styles["Normal"],
        fontName="Helvetica-Bold",
        fontSize=11.5,
        leading=14,
        alignment=1,
        textColor=colors.HexColor("#1E2D3D"),
        spaceBefore=4,
        spaceAfter=6,
    )

    style_h1 = ParagraphStyle(
        "GovH1",
        parent=styles["Normal"],
        fontName="Helvetica-Bold",
        fontSize=10.5,
        leading=13,
        textColor=colors.HexColor("#1E2D3D"),
        spaceBefore=10,
        spaceAfter=4,
        keepWithNext=True,
    )

    style_h2 = ParagraphStyle(
        "GovH2",
        parent=styles["Normal"],
        fontName="Helvetica-Bold",
        fontSize=9.5,
        leading=12,
        textColor=colors.HexColor("#334155"),
        spaceBefore=6,
        spaceAfter=3,
        keepWithNext=True,
    )

    style_body = ParagraphStyle(
        "GovBody",
        parent=styles["Normal"],
        fontName="Helvetica",
        fontSize=9,
        leading=12.5,
        alignment=4,  # Justificado
        spaceBefore=2,
        spaceAfter=4,
        textColor=colors.HexColor("#0F172A"),
    )

    style_th = ParagraphStyle(
        "GovTH",
        parent=styles["Normal"],
        fontName="Helvetica-Bold",
        fontSize=8,
        leading=10,
        alignment=1,
        textColor=colors.white,
    )

    style_td = ParagraphStyle(
        "GovTD",
        parent=styles["Normal"],
        fontName="Helvetica",
        fontSize=7.5,
        leading=9.5,
        textColor=colors.HexColor("#1E293B"),
    )

    story = []

    # 1. Brasão
    brasao_path = os.path.join(BASE_DIR, "brasao_oficial.png")
    if os.path.exists(brasao_path):
        story.append(RLImage(brasao_path, width=1.05 * 72, height=1.05 * 72))
        story.append(Spacer(1, 4))

    # 2. Cabeçalho Oficial
    orgao_val = doc_structure.get("orgao") or org_name
    story.append(Paragraph(
        f"REPÚBLICA FEDERATIVA DO BRASIL<br/>MINISTÉRIO DA DEFESA<br/>COMANDO DA MARINHA<br/>{xml_escape(orgao_val)}",
        style_head,
    ))

    # 3. Título
    tit_val = doc_structure.get("title", "DOCUMENTO ADMINISTRATIVO AUDITADO")
    story.append(Paragraph(xml_escape(tit_val), style_title))

    if doc_structure.get("processo"):
        story.append(Paragraph(
            f"<i>Processo Administrativo: {xml_escape(doc_structure.get('processo'))}</i>",
            ParagraphStyle("Proc", parent=style_head, fontName="Helvetica-Oblique", fontSize=8.5, leading=10, spaceAfter=8),
        ))

    # 4. Banner de Convenção de Auditoria (Legenda)
    legend_html = (
        "<b>CONVENÇÃO DE AUDITORIA (LEI Nº 14.133/2021 &amp; AGU):</b> "
        "<font backcolor=\"#FEF08A\"><b>[AMARELO]</b></font> = Dispositivos inseridos ou retificados para conformidade AGU. &nbsp;&nbsp;"
        "<font color=\"#DC2626\"><strike>[VERMELHO TACHADO]</strike></font> = Dispositivos revogados, obsoletos ou extirpados."
    )
    t_leg = Table(
        [[Paragraph(legend_html, ParagraphStyle("Leg", parent=styles["Normal"], fontName="Helvetica", fontSize=8, leading=10.5, textColor=colors.HexColor("#1E293B")))]],
        colWidths=[487],
    )
    t_leg.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#F1F5F9")),
        ("BOX", (0, 0), (-1, -1), 0.5, colors.HexColor("#CBD5E1")),
        ("PADDING", (0, 0), (-1, -1), 6),
    ]))
    story.append(t_leg)
    story.append(Spacer(1, 8))

    # 5. Renderização dos Elementos Estruturais
    for el in doc_structure.get("elements", []):
        el_type = el.get("type")

        if el_type == "HEADING_1":
            sec_num = el.get("number", "")
            sec_title = el.get("title", "")
            title_text = f"{sec_num}. {sec_title}" if sec_num and sec_num != "+" else sec_title
            story.append(Paragraph(xml_escape(title_text), style_h1))

        elif el_type == "HEADING_2":
            story.append(Paragraph(xml_escape(el.get("title", "")), style_h2))

        elif el_type in ["CLAUSE", "PARAGRAPH"]:
            cl_num = el.get("clause_number", "")
            diffs = el.get("diff_segments") or [{"type": "equal", "text": el.get("text", "")}]

            parts = []
            if cl_num:
                if el.get("modified"):
                    parts.append(f"<font backcolor=\"#FEF08A\"><b>{xml_escape(cl_num)}.</b></font>")
                else:
                    parts.append(f"<b>{xml_escape(cl_num)}.</b>")

            for seg in diffs:
                st = seg.get("type")
                txt = xml_escape(seg.get("text", ""))
                if st == "added":
                    parts.append(f"<font backcolor=\"#FEF08A\"><b>{txt}</b></font>")
                elif st == "removed":
                    parts.append(f"<font color=\"#DC2626\"><strike>{txt}</strike></font>")
                else:
                    parts.append(txt)

            p_xml = " ".join(parts)
            story.append(Paragraph(p_xml, style_body))

        elif el_type == "TABLE":
            headers = el.get("headers", [])
            rows = el.get("rows", [])
            if not headers and not rows:
                continue

            num_cols = max(len(headers), max((len(r) for r in rows), default=0))
            if num_cols == 0:
                continue

            col_w = 487.0 / num_cols
            table_data = []

            if headers:
                table_data.append([
                    Paragraph(f"<b>{xml_escape(str(h))}</b>", style_th)
                    for h in headers[:num_cols]
                ])

            for r in rows:
                row_cells = []
                for val in r[:num_cols]:
                    row_cells.append(Paragraph(xml_escape(str(val)), style_td))
                while len(row_cells) < num_cols:
                    row_cells.append(Paragraph("", style_td))
                table_data.append(row_cells)

            t = Table(table_data, colWidths=[col_w] * num_cols, repeatRows=1 if headers else 0)
            t.setStyle(TableStyle([
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1E2D3D")),
                ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#CBD5E1")),
                ("ROWBACKGROUNDS", (0, 1 if headers else 0), (-1, -1), [colors.white, colors.HexColor("#F8FAFC")]),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("TOPPADDING", (0, 0), (-1, -1), 3),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
                ("LEFTPADDING", (0, 0), (-1, -1), 4),
                ("RIGHTPADDING", (0, 0), (-1, -1), 4),
            ]))
            story.append(Spacer(1, 4))
            story.append(t)
            story.append(Spacer(1, 4))

    # 6. Assinaturas
    signatures = doc_structure.get("signatures", [])
    if signatures:
        story.append(Spacer(1, 14))
        story.append(Paragraph("<b>ASSINATURAS DOS RESPONSÁVEIS:</b>", style_h1))
        for sig in signatures:
            sig_name = xml_escape(sig.get("name", "RESPONSÁVEL TÉCNICO"))
            sig_role = xml_escape(sig.get("role", "Membro da Equipe de Planejamento"))
            sig_p = f"_____________________________________________________<br/><b>{sig_name}</b><br/><i>{sig_role}</i>"
            story.append(Spacer(1, 6))
            story.append(Paragraph(sig_p, style_body))

    doc.build(story, canvasmaker=NumberedCanvas)
    return buf.getvalue()
