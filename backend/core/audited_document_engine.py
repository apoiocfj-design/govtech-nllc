# -*- coding: utf-8 -*-
"""
MÓDULO MASTER DE ENGENHARIA DE DOCUMENTOS AUDITADOS (LEI Nº 14.133/2021 & AGU)
=============================================================================
Arquitetura Enterprise com Controle de Permissão Cláusula a Cláusula e
Preservação Fiel de Formato e Diagramação Original:

1. DocumentSessionStore:
   Mantém em disco temporário (TTL 60 min) o arquivo original (.DOCX ou .PDF)
   para edição cirúrgica fiel sem perda de formatação.

2. DocumentLayoutParser (Alta Fidelidade):
   - DOCX: Mapeia cada parágrafo e cada célula de tabela com índice exato no XML.
   - PDF: Extrai tamanho real da página, margens, fontes, tamanhos, negrito/itálico,
     alinhamentos reais (x0/x1), tabelas com larguras proporcionais de colunas e
     metadados/assinaturas do Compras.gov.br (TR, ETP, DFD, Editais).

3. DocumentLegalAuditor v3.0 (Unificado NLLC + AGU + RAG):
   - NÃO muta o documento sem permissão do usuário!
   - Analisa parágrafos, cláusulas E células de tabelas.
   - Produz propostas estruturadas (`findings`) com diff exato DE -> PARA,
     fundamentação legal e status `pending` aguardando aprovação expressa.

4. Exportador Cirúrgico Word (.DOCX) e PDF (.PDF):
   - Se original era .DOCX: abre o PRÓPRIO .DOCX original e edita in-place
     preservando 100% de cabeçalhos, rodapés, brasões, estilos, fontes, margens
     e tabelas originais!
   - Se original era .PDF: reconstrói o .DOCX preservando margens, fontes,
     tamanhos, alinhamentos e tabelas proporcionais do PDF original!
   - Aplica SOMENTE as alterações permitidas pelo usuário:
     * AMARELO (WD_COLOR_INDEX.YELLOW + negrito): texto alterado ou inserido.
     * VERMELHO TACHADO (#DC2626 + strike): texto retirado/revogado.
     * ORIGINAL INTACTO: texto não alterado ou recusado pelo usuário.
"""

from __future__ import annotations
import copy
import difflib
import io
import os
import re
import tempfile
import threading
import time
import uuid
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
)
from reportlab.pdfgen import canvas
import pdfplumber

BASE_DIR = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))


# ==============================================================================
# 0. SESSÃO DE ARQUIVOS ORIGINAIS EM DISCO (TTL 60 MIN)
# ==============================================================================
class DocumentSessionStore:
    """Armazena os bytes originais do arquivo enviado em disco temporário com TTL."""

    def __init__(self, ttl_seconds: int = 3600, max_sessions: int = 40):
        self._lock = threading.Lock()
        self._sessions: Dict[str, Dict[str, Any]] = {}
        self._ttl = ttl_seconds
        self._max_sessions = max_sessions

    def _cleanup_locked(self) -> None:
        now = time.time()
        expired = [
            sid for sid, meta in self._sessions.items()
            if now - meta.get("updated_at", now) > self._ttl
        ]
        if len(self._sessions) - len(expired) > self._max_sessions:
            ordered = sorted(
                ((meta.get("updated_at", 0), sid) for sid, meta in self._sessions.items() if sid not in expired)
            )
            excess = (len(self._sessions) - len(expired)) - self._max_sessions
            expired.extend(sid for _, sid in ordered[:max(0, excess)])

        for sid in expired:
            meta = self._sessions.pop(sid, None)
            if meta and meta.get("file_path"):
                try:
                    if os.path.exists(meta["file_path"]):
                        os.remove(meta["file_path"])
                except OSError:
                    pass

    def save(self, original_bytes: bytes, filename: str, original_format: str, doc_structure: Dict[str, Any]) -> str:
        sid = f"doc_{uuid.uuid4().hex[:16]}"
        ext = f".{original_format.lower().lstrip('.')}" if original_format else ".bin"
        fd, tmp_path = tempfile.mkstemp(prefix="govtech_doc_", suffix=ext)
        with os.fdopen(fd, "wb") as f:
            f.write(original_bytes)

        now = time.time()
        with self._lock:
            self._cleanup_locked()
            self._sessions[sid] = {
                "file_path": tmp_path,
                "filename": filename,
                "original_format": original_format.lower(),
                "doc_structure": copy.deepcopy(doc_structure),
                "created_at": now,
                "updated_at": now,
            }
        return sid

    def get(self, sid: str) -> Optional[Dict[str, Any]]:
        if not sid:
            return None
        with self._lock:
            meta = self._sessions.get(sid)
            if not meta:
                return None
            meta["updated_at"] = time.time()
            raw_bytes = None
            if meta.get("file_path") and os.path.exists(meta["file_path"]):
                with open(meta["file_path"], "rb") as f:
                    raw_bytes = f.read()
            return {
                "original_bytes": raw_bytes,
                "filename": meta["filename"],
                "original_format": meta["original_format"],
                "doc_structure": copy.deepcopy(meta["doc_structure"]),
            }


document_session_store = DocumentSessionStore()


# ==============================================================================
# 1. FUNÇÕES AUXILIARES DE DIFF, FONTES E ESCAPE
# ==============================================================================
def xml_escape(text: str) -> str:
    """Escapa caracteres especiais XML para renderização segura no ReportLab."""
    if not text:
        return ""
    return (
        str(text)
        .replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace('"', "&quot;")
    )


def compute_word_diff(original_text: str, revised_text: str) -> List[Dict[str, str]]:
    """
    Compara texto original vs revisado no nível de palavras preservando pontuação.
    Retorna segmentos:
    - {'type': 'equal', 'text': '...'}
    - {'type': 'removed', 'text': '...'}  (Vermelho + Tachado)
    - {'type': 'added', 'text': '...'}    (Amarelo + Negrito)
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

    matcher = difflib.SequenceMatcher(None, w_orig, w_rev, autojunk=False)
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


def _map_pdf_font_to_word(pdf_fontname: str) -> Tuple[str, bool, bool]:
    """Mapeia nomes de fontes embutidas de PDF para fontes Word equivalentes + bold/italic."""
    fn = str(pdf_fontname or "")
    # Remove prefixo subset (ex: 'HQXQFE+LiberationSerif-Bold' -> 'LiberationSerif-Bold')
    if "+" in fn:
        fn = fn.split("+", 1)[1]
    low = fn.lower()
    is_bold = any(k in low for k in ("bold", "black", "heavy", "demi", "-bd", "semibold"))
    is_italic = any(k in low for k in ("italic", "oblique", "-it", "slant"))

    if "serif" in low or "times" in low or "georgia" in low or "garamond" in low:
        family = "Times New Roman"
    elif "mono" in low or "courier" in low or "consolas" in low:
        family = "Courier New"
    elif "arial" in low or "helvetica" in low or "liberation" in low or "sans" in low:
        family = "Arial"
    else:
        family = "Calibri"
    return family, is_bold, is_italic


def set_docx_cell_background(cell, hex_color: str):
    """Aplica cor de fundo a uma célula de tabela Word via XML nativo."""
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_color.replace("#", "").upper())
    tcPr.append(shd)


def set_docx_cell_margins(cell, top=80, bottom=80, left=120, right=120):
    """Define margens internas de célula Word."""
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement("w:tcMar")
    for m, val in [("top", top), ("bottom", bottom), ("left", left), ("right", right)]:
        node = OxmlElement(f"w:{m}")
        node.set(qn("w:w"), str(val))
        node.set(qn("w:type"), "dxa")
        tcMar.append(node)
    tcPr.append(tcMar)


def set_docx_table_borders(table, color: str = "CBD5E1", sz: str = "4"):
    """Define bordas limpas em tabelas Word criadas a partir de PDF."""
    tblPr = table._tbl.tblPr
    borders = OxmlElement("w:tblBorders")
    for side in ("top", "left", "bottom", "right", "insideH", "insideV"):
        el = OxmlElement(f"w:{side}")
        el.set(qn("w:val"), "single")
        el.set(qn("w:sz"), sz)
        el.set(qn("w:space"), "0")
        el.set(qn("w:color"), color)
        borders.append(el)
    tblPr.append(borders)


# ==============================================================================
# 2. PARSER DE DOCUMENTOS COM PRESERVAÇÃO REAL DE DIAGRAMAÇÃO
# ==============================================================================
class DocumentLayoutParser:
    """
    Extrai o layout integral de documentos PDF, DOCX e TXT preservando:
    - DOCX: índice exato no XML (parágrafos e células de tabela) para cirurgia in-place.
    - PDF: dimensões de página, margens, tipografia (fonte/tamanho/negrito/itálico),
      alinhamento real, tabelas com larguras proporcionais de colunas e assinaturas.
    """

    @classmethod
    def parse_pdf(cls, file_bytes: bytes, filename: str = "") -> Dict[str, Any]:
        doc_structure: Dict[str, Any] = {
            "original_format": "pdf",
            "filename": filename,
            "doc_type": "TR",
            "title": "TERMO DE REFERÊNCIA (COMPRAS.GOV.BR)",
            "uasg": "",
            "orgao": "MARINHA DO BRASIL",
            "processo": "",
            "page_layout": {
                "width_pt": 595.3,
                "height_pt": 841.9,
                "margin_left_pt": 54.0,
                "margin_right_pt": 54.0,
                "margin_top_pt": 54.0,
                "margin_bottom_pt": 54.0,
                "default_font": "Arial",
                "default_size_pt": 10.0,
            },
            "metadata_box": {},
            "elements": [],
            "signatures": [],
        }

        raw_page_items = []
        min_x0 = 9999.0
        max_x1 = 0.0
        font_counts: Dict[str, int] = {}

        with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
            if pdf.pages:
                p0 = pdf.pages[0]
                doc_structure["page_layout"]["width_pt"] = float(p0.width)
                doc_structure["page_layout"]["height_pt"] = float(p0.height)

            for pno, page in enumerate(pdf.pages):
                pw = float(page.width)
                tables = page.find_tables()
                t_bboxes = [t.bbox for t in tables]

                words = page.extract_words(extra_attrs=["fontname", "size"]) or []
                outside_words = [
                    w for w in words
                    if not any(
                        b[0] - 2 <= w["x0"] <= b[2] + 2 and b[1] - 2 <= w["top"] <= b[3] + 2
                        for b in t_bboxes
                    )
                ]

                lines_with_pos = []
                if outside_words:
                    outside_words.sort(key=lambda w: (round(w["top"], 1), w["x0"]))
                    curr_line = [outside_words[0]]
                    for w in outside_words[1:]:
                        if abs(w["top"] - curr_line[-1]["top"]) < 3.8:
                            curr_line.append(w)
                        else:
                            lines_with_pos.append(cls._summarize_pdf_line(curr_line, pw))
                            curr_line = [w]
                    if curr_line:
                        lines_with_pos.append(cls._summarize_pdf_line(curr_line, pw))

                page_items = []
                for l in lines_with_pos:
                    if not l["text"]:
                        continue
                    min_x0 = min(min_x0, l["x0"])
                    max_x1 = max(max_x1, l["x1"])
                    font_counts[l["font_family"]] = font_counts.get(l["font_family"], 0) + 1
                    page_items.append(("TEXT", l["top"], l))

                for t in tables:
                    t_data = t.extract()
                    clean_rows = [[(c or "").replace("\n", " ").strip() for c in r] for r in (t_data or []) if any(r)]
                    if clean_rows:
                        # Extrai larguras proporcionais reais das colunas
                        col_ratios = cls._extract_table_col_ratios(t, len(clean_rows[0]))
                        page_items.append(("TABLE", t.bbox[1], {
                            "rows": clean_rows,
                            "col_ratios": col_ratios,
                            "bbox": list(t.bbox),
                        }))

                page_items.sort(key=lambda x: x[1])
                for it in page_items:
                    raw_page_items.append((pno + 1, it[0], it[2]))

        # Calcula margens reais inferidas
        pw_total = doc_structure["page_layout"]["width_pt"]
        if 20.0 <= min_x0 <= 120.0:
            doc_structure["page_layout"]["margin_left_pt"] = round(min_x0, 1)
        if max_x1 > 0 and 20.0 <= (pw_total - max_x1) <= 120.0:
            doc_structure["page_layout"]["margin_right_pt"] = round(pw_total - max_x1, 1)
        if font_counts:
            doc_structure["page_layout"]["default_font"] = max(font_counts.items(), key=lambda kv: kv[1])[0]

        # Interpretação estrutural preservando tipografia e alinhamento
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
                rows_all = val["rows"]
                elements.append({
                    "type": "TABLE",
                    "page": pno,
                    "section_num": current_section_num,
                    "headers": rows_all[0] if len(rows_all) > 0 else [],
                    "rows": rows_all[1:] if len(rows_all) > 1 else [],
                    "col_ratios": val.get("col_ratios") or [],
                })
                continue

            line = str(val["text"]).strip()
            if not line:
                continue

            style_meta = {
                "font_family": val["font_family"],
                "font_size": val["font_size"],
                "is_bold": val["is_bold"],
                "is_italic": val["is_italic"],
                "alignment": val["alignment"],
            }

            # Rodapés de paginação tipo "1 de 6" ou identificadores de UASG isolados
            if re.search(r"^\d+\s+de\s+\d+$", line):
                continue
            if re.search(r"^UASG\s*(\d+)$", line):
                doc_structure["uasg"] = re.search(r"^UASG\s*(\d+)$", line).group(1)
                continue
            if re.match(r"^(CAPITANIA FLUVIAL|MARINHA DO BRASIL|MINISTÉRIO DA DEFESA|COMANDO DA MARINHA)", line, re.I) and pno == 1 and len(elements) < 3:
                doc_structure["orgao"] = line
                continue

            # Título principal (ex: "Termo de Referência 37/2026" ou artefatos combinados no mesmo PDF)
            if re.match(r"^(Estudo Técnico Preliminar|Termo de Referência|Documento de Formalização da Demanda|Edital|Aviso de Contratação)\s*([\d/\-]+)?", line, re.I):
                if not doc_structure.get("title") or doc_structure["title"] == "TERMO DE REFERÊNCIA (COMPRAS.GOV.BR)":
                    doc_structure["title"] = line
                low_l = line.lower()
                if "estudo técnico" in low_l:
                    doc_structure["doc_type"] = "ETP"
                elif "termo de referência" in low_l:
                    doc_structure["doc_type"] = "TR"
                elif "formalização" in low_l:
                    doc_structure["doc_type"] = "DFD"
                elif "aviso" in low_l:
                    doc_structure["doc_type"] = "AVISO_DISPENSA"
                if current_clause:
                    elements.append(current_clause)
                    current_clause = None
                elements.append({
                    "type": "PARAGRAPH",
                    "text": line,
                    "original_text": line,
                    "section_num": current_section_num,
                    "section_title": line,
                    "page": pno,
                    "style": {**style_meta, "is_bold": True},
                    "modified": False,
                    "diff_segments": [],
                })
                continue

            if "Número do processo:" in line:
                doc_structure["processo"] = line.split(":", 1)[1].strip()

            # Assinaturas
            if re.match(r"^(Membro da comissão|Responsável|Ordenador|Fiscal|Chefe|Encarregado|Autoridade Competente)", line, re.I) and len(line) < 90:
                prev_text = ""
                if current_clause and current_clause["text"]:
                    words_prev = current_clause["text"].split()
                    if words_prev:
                        prev_text = " ".join(words_prev[-4:])
                detected_signatures.append({"role": line, "name": prev_text or "RESPONSÁVEL PELA MINUTA"})
                continue

            # Seção Principal (Heading 1): "1. Informações Básicas" ou fonte >= 12.5pt bold
            sec_m = re.match(r"^(\d{1,2})\.\s+([A-ZÁÉÍÓÚÂÊÔÃÕÇ\w\s\-/(),]{3,})$", line)
            if sec_m and (val["is_bold"] or val["font_size"] >= 11.5 or len(line) < 95):
                if current_clause:
                    elements.append(current_clause)
                    current_clause = None
                current_section_num = sec_m.group(1)
                current_section_title = sec_m.group(2).strip()
                elements.append({
                    "type": "HEADING_1",
                    "number": current_section_num,
                    "title": current_section_title,
                    "text": f"{current_section_num}. {current_section_title}",
                    "page": pno,
                    "style": style_meta,
                })
                continue

            # Subcláusula numerada: "2.1. Contratação..." ou "7.2.1."
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
                    "style": style_meta,
                    "modified": False,
                    "diff_segments": [],
                }
                continue

            # Subtítulos em negrito/fonte maior fora de cláusula
            if (val["font_size"] >= 12.0 and val["is_bold"] and len(line) < 90) or (
                not current_clause and len(line) < 75 and not line.endswith(".") and (val["is_bold"] or line.isupper())
            ):
                if current_clause:
                    elements.append(current_clause)
                    current_clause = None
                elements.append({
                    "type": "HEADING_2",
                    "title": line,
                    "text": line,
                    "section_num": current_section_num,
                    "page": pno,
                    "style": style_meta,
                })
                continue

            # Continuação de subcláusula (se mesma página ou continuação natural sem quebra de bloco)
            if current_clause and not (val["is_bold"] and len(line) < 60):
                current_clause["text"] = (current_clause["text"] + " " + line).strip()
            else:
                if current_clause:
                    elements.append(current_clause)
                    current_clause = None
                elements.append({
                    "type": "PARAGRAPH",
                    "text": line,
                    "section_num": current_section_num,
                    "section_title": current_section_title,
                    "page": pno,
                    "style": style_meta,
                    "modified": False,
                    "diff_segments": [],
                })

        if current_clause:
            elements.append(current_clause)

        doc_structure["elements"] = elements
        doc_structure["signatures"] = detected_signatures
        return doc_structure

    @staticmethod
    def _summarize_pdf_line(words: List[Dict[str, Any]], page_width: float) -> Dict[str, Any]:
        text_str = " ".join(w["text"] for w in words).strip()
        x0 = min(float(w["x0"]) for w in words)
        x1 = max(float(w["x1"]) for w in words)
        top = float(words[0]["top"])
        sizes = [float(w.get("size") or 10.0) for w in words]
        avg_size = round(sum(sizes) / max(1, len(sizes)), 1)

        # Fonte predominante da linha
        first_fn = words[0].get("fontname") or ""
        font_family, is_bold, is_italic = _map_pdf_font_to_word(first_fn)
        bold_words = sum(1 for w in words if _map_pdf_font_to_word(w.get("fontname", ""))[1])
        is_bold = bold_words >= max(1, len(words) * 0.6)

        # Inferência de alinhamento
        left_gap = x0
        right_gap = max(0.0, page_width - x1)
        line_w = x1 - x0
        if line_w < page_width * 0.65 and abs(left_gap - right_gap) < 28.0 and left_gap > 85.0:
            alignment = "CENTER"
        elif right_gap < 70.0 and left_gap > 200.0:
            alignment = "RIGHT"
        elif line_w > page_width * 0.72:
            alignment = "JUSTIFY"
        else:
            alignment = "LEFT"

        return {
            "top": top,
            "x0": x0,
            "x1": x1,
            "text": text_str,
            "font_family": font_family,
            "font_size": avg_size,
            "is_bold": is_bold,
            "is_italic": is_italic,
            "alignment": alignment,
        }

    @staticmethod
    def _extract_table_col_ratios(table_obj: Any, num_cols: int) -> List[float]:
        if num_cols <= 0:
            return []
        try:
            bbox = table_obj.bbox
            total_w = max(1.0, float(bbox[2]) - float(bbox[0]))
            # Extrai x0/x1 únicos das células da primeira linha
            row_cells = table_obj.rows[0].cells if table_obj.rows else []
            widths = []
            for c in row_cells[:num_cols]:
                if c is not None:
                    widths.append(max(10.0, float(c[2]) - float(c[0])))
                else:
                    widths.append(total_w / num_cols)
            s = sum(widths) or total_w
            return [round(w / s, 4) for w in widths]
        except Exception:
            return [round(1.0 / num_cols, 4)] * num_cols

    @classmethod
    def parse_docx(cls, file_bytes: bytes, filename: str = "") -> Dict[str, Any]:
        """
        Extrai parágrafos, seções e tabelas nativas de arquivo .DOCX preservando
        o índice exato (`docx_body_idx` / `docx_p_idx` / `docx_tbl_idx`) no XML original
        para permitir edição cirúrgica in-place 100% fiel ao layout!
        """
        doc = docx.Document(io.BytesIO(file_bytes))
        doc_structure: Dict[str, Any] = {
            "original_format": "docx",
            "filename": filename,
            "doc_type": "TR",
            "title": os.path.splitext(filename)[0] if filename else "TERMO DE REFERÊNCIA (WORD)",
            "uasg": "",
            "orgao": "MARINHA DO BRASIL",
            "processo": "",
            "metadata_box": {},
            "elements": [],
            "signatures": [],
        }

        elements: List[Dict[str, Any]] = []
        current_section_num = "0"
        current_section_title = "PREÂMBULO"
        p_idx = -1
        tbl_idx = -1

        for body_idx, block in enumerate(doc.element.body):
            tag = block.tag.split("}")[-1] if "}" in block.tag else block.tag

            if tag == "p":
                p_idx += 1
                p = docx.text.paragraph.Paragraph(block, doc)
                line = p.text.strip()
                if not line:
                    continue

                if re.match(r"^(Estudo Técnico|Termo de Referência|Documento de Formalização|Edital|Aviso de Contratação)", line, re.I):
                    doc_structure["title"] = line

                if "Número do processo:" in line:
                    doc_structure["processo"] = line.split(":", 1)[1].strip()

                p_style_name = (p.style.name if p.style else "Normal") or "Normal"
                is_heading_style = "heading" in p_style_name.lower() or "título" in p_style_name.lower()

                sec_m = re.match(r"^(\d{1,2})\.\s+([A-ZÁÉÍÓÚÂÊÔÃÕÇ\w\s\-/(),]{3,})$", line)
                if sec_m and (is_heading_style or len(line) < 95):
                    current_section_num = sec_m.group(1)
                    current_section_title = sec_m.group(2).strip()
                    elements.append({
                        "type": "HEADING_1",
                        "number": current_section_num,
                        "title": current_section_title,
                        "text": line,
                        "page": 1,
                        "docx_p_idx": p_idx,
                        "docx_body_idx": body_idx,
                    })
                    continue

                cl_m = re.match(r"^(\d{1,2}(?:\.\d{1,3})+)\.?\s*(.*)$", line)
                if cl_m:
                    elements.append({
                        "type": "CLAUSE",
                        "clause_number": cl_m.group(1),
                        "section_num": current_section_num,
                        "section_title": current_section_title,
                        "text": line,  # Mantém texto integral do parágrafo para cirurgia exata!
                        "page": 1,
                        "docx_p_idx": p_idx,
                        "docx_body_idx": body_idx,
                        "modified": False,
                        "diff_segments": [],
                    })
                    continue

                elements.append({
                    "type": "PARAGRAPH",
                    "text": line,
                    "section_num": current_section_num,
                    "section_title": current_section_title,
                    "page": 1,
                    "docx_p_idx": p_idx,
                    "docx_body_idx": body_idx,
                    "modified": False,
                    "diff_segments": [],
                })

            elif tag == "tbl":
                tbl_idx += 1
                tbl = docx.table.Table(block, doc)
                rows_data = []
                for row in tbl.rows:
                    r_text = [cell.text.strip() for cell in row.cells]
                    rows_data.append(r_text)
                if any(any(c for c in r) for r in rows_data):
                    elements.append({
                        "type": "TABLE",
                        "page": 1,
                        "section_num": current_section_num,
                        "docx_tbl_idx": tbl_idx,
                        "docx_body_idx": body_idx,
                        "headers": rows_data[0],
                        "rows": rows_data[1:] if len(rows_data) > 1 else [],
                    })

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
            "uasg": "",
            "orgao": "MARINHA DO BRASIL",
            "processo": "",
            "metadata_box": {},
            "elements": [],
            "signatures": [],
        }

        elements: List[Dict[str, Any]] = []
        current_clause: Optional[Dict[str, Any]] = None
        current_section_num = "0"
        current_section_title = "PREÂMBULO"

        lines = [l.strip() for l in (raw_text or "").split("\n") if l.strip()]
        for line in lines:
            if re.match(r"^(Estudo Técnico|Termo de Referência|Documento de Formalização)", line, re.I):
                doc_structure["title"] = line
                continue

            sec_m = re.match(r"^(\d{1,2})\.\s+([A-ZÁÉÍÓÚÂÊÔÃÕÇ\w\s\-/(),]{3,})$", line)
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
                    "text": line,
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
                    "section_title": current_section_title,
                    "page": 1,
                    "modified": False,
                    "diff_segments": [],
                })

        if current_clause:
            elements.append(current_clause)

        doc_structure["elements"] = elements
        return doc_structure


# ==============================================================================
# 3. MOTOR DE AUDITORIA JURÍDICA NLLC + AGU + RAG v3.0 (COM PERMISSÃO)
# ==============================================================================
class DocumentLegalAuditor:
    """
    Motor unificado de auditoria jurídica (Lei nº 14.133/2021 + AGU + TCU + IN 58/65/81):
    - Identifica vulnerabilidades em parágrafos, cláusulas E células de tabelas.
    - Detecta omissões obrigatórias (Art. 117, Art. 18 XI, Art. 75 § 1º, Art. 150, Art. 92).
    - NÃO altera o documento automaticamente sem permissão do usuário (`require_permission=True`).
    - Aplica exclusivamente as alterações aprovadas via `apply_approved_decisions(...)`.
    """

    # Tuplas: (regex, substituição, título curto, explicação técnica, base_legal, severidade)
    RULES = [
        (
            r"\b(?:Lei\s+(?:Federal\s+)?n[ºo°]?\s*8\.?666(?:/1993|/93)?|Lei\s+8\.666)\b",
            "Lei nº 14.133/2021",
            "Referência à Lei Revogada nº 8.666/1993",
            "A Lei nº 8.666/1993 foi integralmente revogada em 30/12/2023. Toda minuta deve fundamentar-se na Nova Lei de Licitações e Contratos (Lei nº 14.133/2021).",
            "Art. 193, II da Lei nº 14.133/2021",
            "CRITICA",
        ),
        (
            r"\b(?:Lei\s+(?:Federal\s+)?n[ºo°]?\s*10\.?520(?:/2002|/02)?|Lei\s+10\.520)\b",
            "Lei nº 14.133/2021",
            "Referência à Lei do Pregão Revogada (10.520/2002)",
            "A Lei nº 10.520/2002 foi revogada. A modalidade Pregão é agora disciplinada diretamente pelo Art. 28, I e Art. 29 da Lei nº 14.133/2021.",
            "Art. 28, I e Art. 193, II da Lei nº 14.133/2021",
            "CRITICA",
        ),
        (
            r"\b(?:Decreto\s+n[ºo°]?\s*10\.?024(?:/2019)?|Decreto\s+10\.024)\b",
            "Instrução Normativa SEGES/ME nº 73/2022 e Lei nº 14.133/2021",
            "Decreto nº 10.024/2019 (Regulamento Antigo do Pregão)",
            "O Decreto nº 10.024/2019 regulava o pregão sob a Lei 10.520/02. Na NLLC, o procedimento eletrônico segue a IN SEGES/ME nº 73/2022 e a Lei nº 14.133/2021.",
            "IN SEGES/ME nº 73/2022; Lei nº 14.133/2021",
            "ALTA",
        ),
        (
            r"\b(?:Decreto\s+n[ºo°]?\s*7\.?892(?:/2013)?|Decreto\s+7\.892)\b",
            "Decreto nº 11.462/2023",
            "Decreto de Registro de Preços Revogado (7.892/2013)",
            "O Sistema de Registro de Preços (SRP) no âmbito da Administração Pública Federal passou a ser regulamentado pelo Decreto nº 11.462/2023.",
            "Decreto Federal nº 11.462/2023; Art. 82 da Lei nº 14.133/2021",
            "CRITICA",
        ),
        (
            r"\b(?:Decreto\s+n[ºo°]?\s*3\.?555(?:/2000)?|Decreto\s+n[ºo°]?\s*5\.?450(?:/2005)?)\b",
            "Decreto nº 11.462/2023 e Lei nº 14.133/2021",
            "Decretos Federais de Licitação Revogados",
            "Os Decretos nº 3.555/2000 e nº 5.450/2005 foram revogados e substituídos pelo novo marco regulatório da Lei nº 14.133/2021.",
            "Lei nº 14.133/2021; Decreto nº 11.462/2023",
            "CRITICA",
        ),
        (
            r"\bInstru[çc][ãa]o\s+Normativa\s+(?:SEGES/ME\s+)?n[ºo°]?\s*73(?:/2020|,\s*de\s*2020)\b",
            "Instrução Normativa SEGES/ME nº 65/2021",
            "IN nº 73/2020 (Pesquisa de Preços Antiga)",
            "Para contratações sob a Lei nº 14.133/2021, a pesquisa de preços é regida pela Instrução Normativa SEGES/ME nº 65/2021 (e não pela IN 73/2020).",
            "Art. 23 da Lei nº 14.133/2021; IN SEGES/ME nº 65/2021",
            "ALTA",
        ),
        (
            r"\bsuspens[ãa]o\s+tempor[áa]ria\s+de\s+participa[çc][ãa]o\s+em\s+licita[çc][ãa]o\b.*?(?:2|dois)\s+anos",
            "impedimento de licitar e contratar no âmbito da Administração Pública direta e indireta federal pelo prazo de até 3 (três) anos (Art. 156, inciso III da Lei nº 14.133/2021)",
            "Sanção Administrativa Extinta (Suspensão de 2 Anos)",
            "A sanção de suspensão temporária por até 2 anos (Art. 87, III da Lei 8.666) foi extinta. O Art. 156, III da NLLC prevê impedimento de licitar e contratar por até 3 anos.",
            "Art. 156, inciso III e § 4º da Lei nº 14.133/2021",
            "CRITICA",
        ),
        (
            r"\bdeclara[çc][ãa]o\s+de\s+inidoneidade\b.*?(?:2|dois)\s+anos",
            "declaração de inidoneidade para licitar ou contratar pelo prazo mínimo de 3 (três) anos e máximo de 6 (seis) anos (Art. 156, inciso IV e § 5º da Lei nº 14.133/2021)",
            "Prazo Inválido de Declaração de Inidoneidade",
            "Sob a Lei nº 14.133/2021, o prazo da declaração de inidoneidade é fixado entre o mínimo de 3 anos e o máximo de 6 anos (Art. 156, § 5º).",
            "Art. 156, inciso IV e § 5º da Lei nº 14.133/2021",
            "CRITICA",
        ),
        (
            r"\bart(?:igo)?\.?\s*24,?\s*inciso\s*II\s+da\s+Lei\s*(?:n[ºo°]?\s*)?8\.?666(?:/93|/1993)?\b",
            "art. 75, inciso II da Lei nº 14.133/2021",
            "Dispensa de Licitação com Base Revogada (Art. 24, II)",
            "A dispensa em razão do valor para outros serviços e compras é agora disciplinada pelo Art. 75, inciso II da Lei nº 14.133/2021.",
            "Art. 75, inciso II da Lei nº 14.133/2021",
            "CRITICA",
        ),
        (
            r"\bart(?:igo)?\.?\s*24,?\s*inciso\s*I\s+da\s+Lei\s*(?:n[ºo°]?\s*)?8\.?666(?:/93|/1993)?\b",
            "art. 75, inciso I da Lei nº 14.133/2021",
            "Dispensa de Engenharia com Base Revogada (Art. 24, I)",
            "A dispensa em razão do valor para obras e serviços de engenharia é regida pelo Art. 75, inciso I da Lei nº 14.133/2021.",
            "Art. 75, inciso I da Lei nº 14.133/2021",
            "CRITICA",
        ),
        (
            r"\bart(?:igo)?\.?\s*25\s+da\s+Lei\s*(?:n[ºo°]?\s*)?8\.?666(?:/93|/1993)?\b",
            "art. 74 da Lei nº 14.133/2021",
            "Inexigibilidade com Base Revogada (Art. 25 Lei 8.666)",
            "As hipóteses de inexigibilidade de licitação passaram a constar do Art. 74 da Lei nº 14.133/2021.",
            "Art. 74 da Lei nº 14.133/2021",
            "CRITICA",
        ),
        (
            r"\b(?:EIRELI|Empresa\s+Individual\s+de\s+Responsabilidade\s+Limitada)\b",
            "Sociedade Limitada Unipessoal – SLU",
            "Menção à Figura Jurídica Extinta (EIRELI)",
            "A EIRELI foi extinta pelo Art. 41 da Lei nº 14.195/2021 e substituída automaticamente pela Sociedade Limitada Unipessoal (SLU). Padrão AGU atualizado.",
            "Art. 66 da Lei nº 14.133/2021 c/c Art. 41 da Lei nº 14.195/2021",
            "ALTA",
        ),
        (
            r"\b12\s*\(doze\)\s*meses\s*contados\s*da\s*emiss[ãa]o\s*da\s*Nota\s*de\s*Empenho\b",
            "60 (sessenta) dias contados da emissão e recebimento da Nota de Empenho, na forma do art. 105 da Lei nº 14.133/2021 (compatível com entrega imediata)",
            "Vigência Incompatível com Entrega Imediata (12 Meses)",
            "Segundo Nota Explicativa da AGU, compras com entrega imediata sem obrigações futuras devem ter vigência limitada ao ciclo de entrega, recebimento e pagamento (não 12 meses).",
            "Art. 6º, X, Art. 95, I e Art. 105 da Lei nº 14.133/2021; IN 81/2022",
            "ALTA",
        ),
        (
            r"\bvedada\s+a\s+participa[çc][ãa]o\s+de\s+(?:sociedades\s+)?cooperativas\b",
            "permitida a participação de sociedades cooperativas, desde que atendam integralmente aos requisitos do art. 16 da Lei nº 14.133/2021",
            "Vedação Genérica a Cooperativas (Restrição à Competitividade)",
            "O TCU (Acórdãos 2.054/2018 e 1.636/2021-Plenário) proíbe vedar sumariamente a participação de cooperativas sem justificativa técnica idônea.",
            "Art. 9º, I, 'a' e Art. 16 da Lei nº 14.133/2021; Acórdão TCU 2.054/2018",
            "ALTA",
        ),
        (
            r"\bInstru[çc][ãa]o\s+Normativa\s+(?:SLTI/MP|SEGES/MP)\s+n[ºo°]?\s*0?5(?:/2017|,\s*de\s*2017)\b",
            "Instrução Normativa SEGES/ME nº 98/2022 (aplicável à Lei nº 14.133/2021)",
            "IN nº 05/2017 sem Remissão à IN SEGES/ME nº 98/2022",
            "A aplicação da IN 05/2017 sob a Lei nº 14.133/2021 depende da autorização expressa da IN SEGES/ME nº 98/2022.",
            "IN SEGES/ME nº 98/2022; Lei nº 14.133/2021",
            "MEDIA",
        ),
        (
            r"\bLei\s+Complementar\s+n[ºo°]?\s*123/2006\b(?!.*art\.?\s*4[ºo°]\s+da\s+Lei\s+n[ºo°]?\s*14\.133)",
            "Lei Complementar nº 123/2006, observados os limites do art. 4º da Lei nº 14.133/2021",
            "Complementação Obrigatória do Art. 4º NLLC (ME/EPP)",
            "Os benefícios da LC 123/2006 na Lei nº 14.133/2021 submetem-se estritamente aos limites de receita bruta e contratação anual do Art. 4º, §§ 1º a 3º.",
            "Art. 4º, §§ 1º a 3º da Lei nº 14.133/2021",
            "MEDIA",
        ),
    ]

    MANDATORY_OMISSIONS = [
        {
            "key": "GESTAO_FISCALIZACAO",
            "regex": r"\b(?:fiscal\s+do\s+contrato|gestor\s+do\s+contrato|modelo\s+de\s+gest[ãa]o|fiscaliza[çc][ãa]o\s+t[ée]cnica|art\.?\s*117)\b",
            "title": "MODELO DE GESTÃO E FISCALIZAÇÃO DO CONTRATO",
            "clause_num": "GESTÃO CONTRATUAL (ART. 117)",
            "text": (
                "A execução do contrato deverá ser acompanhada e fiscalizada por 1 (um) ou mais fiscais do contrato, "
                "representantes da Administração especialmente designados conforme o art. 117 da Lei nº 14.133/2021 "
                "e o Decreto nº 11.246/2022, anotando em registro próprio todas as ocorrências relacionadas à execução."
            ),
            "explanation": "Ausência de cláusula expressa sobre o modelo de gestão e designação de fiscal do contrato.",
            "legal_foundation": "Art. 40, § 1º, VII e Art. 117 da Lei nº 14.133/2021; Decreto nº 11.246/2022",
            "severity": "ALTA",
        },
        {
            "key": "SUSTENTABILIDADE",
            "regex": r"\b(?:sustentabilidade|impactos?\s+ambientais?|rotulagem\s+ambiental|ibama|log[íi]stica\s+reversa)\b",
            "title": "CRITÉRIOS DE SUSTENTABILIDADE AMBIENTAL",
            "clause_num": "SUSTENTABILIDADE (ART. 18, XI)",
            "text": (
                "A contratada deverá atender aos critérios e práticas de sustentabilidade ambiental na execução do objeto, "
                "observando a redução de impactos ambientais, descarte adequado de resíduos e embalagens, e conformidade "
                "com o Guia Nacional de Contratações Sustentáveis da AGU (Art. 18, XI e Art. 144 da Lei nº 14.133/2021)."
            ),
            "explanation": "O documento não contempla critérios ou justificativa relativa à sustentabilidade socioambiental.",
            "legal_foundation": "Art. 5º, Art. 11, IV e Art. 18, XI da Lei nº 14.133/2021; Guia AGU",
            "severity": "ALTA",
        },
        {
            "key": "ADEQUACAO_ORCAMENTARIA",
            "regex": r"\b(?:adequa[çc][ãa]o\s+or[çc]ament[áa]ria|dota[çc][ãa]o\s+or[çc]ament[áa]ria|ptres|fonte\s+de\s+recursos|cr[ée]ditos?\s+or[çc]ament[áa]rios?)\b",
            "title": "ADEQUAÇÃO ORÇAMENTÁRIA",
            "clause_num": "ORÇAMENTO (ART. 150)",
            "text": (
                "As despesas decorrentes da presente contratação correrão à conta de recursos específicos consignados no "
                "Orçamento Geral da União, com indicação obrigatória da Gestão/Unidade, Fonte de Recursos, Programa de Trabalho "
                "(PTRES), Elemento de Despesa e Plano Interno (Art. 18, caput, Art. 72, IV e Art. 150 da Lei nº 14.133/2021)."
            ),
            "explanation": "Não localizada cláusula de adequação orçamentária com indicação dos créditos públicos.",
            "legal_foundation": "Art. 18, Art. 72, IV e Art. 150 da Lei nº 14.133/2021; IN SEGES 81/2022",
            "severity": "ALTA",
        },
        {
            "key": "CRITERIOS_PAGAMENTO",
            "regex": r"\b(?:liquida[çc][ãa]o|crit[ée]rios?\s+de\s+medi[çc][ãa]o|prazo\s+de\s+pagamento|ordem\s+cronol[óo]gica|nota\s+fiscal)\b",
            "title": "CRITÉRIOS DE MEDIÇÃO, LIQUIDAÇÃO E PAGAMENTO",
            "clause_num": "PAGAMENTO (ART. 92, V)",
            "text": (
                "O recebimento provisório e definitivo observará o art. 140 da Lei nº 14.133/2021. A liquidação e o pagamento "
                "serão efetuados em ordem cronológica mediante apresentação da Nota Fiscal/Fatura atestada pelo fiscal do contrato "
                "e verificação de regularidade fiscal no SICAF (Art. 92, V e VI da Lei nº 14.133/2021; IN SEGES/ME nº 77/2022)."
            ),
            "explanation": "Ausência de definição dos critérios de medição, recebimento, liquidação e pagamento.",
            "legal_foundation": "Art. 92, V-VI e Art. 140 da Lei nº 14.133/2021; IN SEGES/ME nº 77/2022",
            "severity": "MEDIA",
        },
    ]

    @classmethod
    def audit_and_diff(
        cls,
        doc_structure: Dict[str, Any],
        require_permission: bool = True,
    ) -> Dict[str, Any]:
        """
        Analisa integralmente o documento (parágrafos, cláusulas e células de tabela)
        e gera propostas estruturadas (`findings`).
        Quando `require_permission=True` (padrão seguro), mantém o documento original
        INTACTO até que o usuário conceda permissão expressa a cada alteração!
        """
        elements = doc_structure.get("elements", [])
        findings: List[Dict[str, Any]] = []
        total_clauses = 0
        finding_seq = 0

        # 1. Auditoria Bottom-Up em parágrafos e cláusulas
        for el_idx, el in enumerate(elements):
            el_type = el.get("type")
            if el_type in ("CLAUSE", "PARAGRAPH"):
                total_clauses += 1
                orig_text = el.get("text", "")
                if not orig_text:
                    continue

                # Inicializa sem mutação (original intacto)
                el["original_text"] = orig_text
                el["modified"] = False
                el["diff_segments"] = [{"type": "equal", "text": orig_text}]

                # Checa cada regra individualmente para dar controle granular ao usuário
                combined_revised = orig_text
                matched_rules = []
                for pattern, replacement, rule_title, reason, legal_basis, severity in cls.RULES:
                    if re.search(pattern, combined_revised, flags=re.IGNORECASE):
                        m = re.search(pattern, combined_revised, flags=re.IGNORECASE)
                        excerpt = m.group(0) if m else orig_text[:100]
                        combined_revised = re.sub(pattern, replacement, combined_revised, flags=re.IGNORECASE)
                        matched_rules.append((rule_title, reason, legal_basis, severity, excerpt))

                if matched_rules and combined_revised != orig_text:
                    finding_seq += 1
                    fid = f"FND_{finding_seq:03d}"
                    # Usa a maior severidade entre as regras disparadas no mesmo parágrafo
                    sev_rank = {"CRITICA": 3, "ALTA": 2, "MEDIA": 1}
                    top_rule = max(matched_rules, key=lambda r: sev_rank.get(r[3], 1))
                    explanations = " | ".join(dict.fromkeys(r[1] for r in matched_rules))
                    legals = "; ".join(dict.fromkeys(r[2] for r in matched_rules))
                    preview_diff = compute_word_diff(orig_text, combined_revised)

                    f_obj = {
                        "id": fid,
                        "type": "VULNERABILIDADE_JURIDICA",
                        "finding_type": "VULNERABILIDADE_JURIDICA",
                        "severity": top_rule[3],
                        "title": top_rule[0],
                        "section_title": el.get("section_title") or top_rule[0],
                        "clause_reference": el.get("clause_number") or f"Parágrafo {el_idx + 1}",
                        "original_clause": el.get("clause_number") or f"Parágrafo {el_idx + 1}",
                        "target_type": "element",
                        "element_idx": el_idx,
                        "docx_p_idx": el.get("docx_p_idx"),
                        "original_text": orig_text,
                        "original_excerpt": top_rule[4],
                        "suggested_text": combined_revised,
                        "agu_standard_clause": combined_revised,
                        "explanation": explanations,
                        "reason": explanations,
                        "legal_foundation": legals,
                        "preview_diff": preview_diff,
                        "status": "approved" if not require_permission else "pending",
                        "approved": not require_permission,
                    }
                    findings.append(f_obj)

                    if not require_permission:
                        el["modified"] = True
                        el["revised_text"] = combined_revised
                        el["diff_segments"] = preview_diff

            # 2. Auditoria Bottom-Up dentro de CÉLULAS DE TABELA
            elif el_type == "TABLE":
                all_rows = ([el.get("headers")] if el.get("headers") else []) + list(el.get("rows") or [])
                for r_i, row in enumerate(all_rows):
                    for c_i, cell_val in enumerate(row):
                        cell_str = str(cell_val or "").strip()
                        if not cell_str or len(cell_str) < 6:
                            continue
                        rev_cell = cell_str
                        matched_rules = []
                        for pattern, replacement, rule_title, reason, legal_basis, severity in cls.RULES:
                            if re.search(pattern, rev_cell, flags=re.IGNORECASE):
                                m = re.search(pattern, rev_cell, flags=re.IGNORECASE)
                                excerpt = m.group(0) if m else cell_str[:80]
                                rev_cell = re.sub(pattern, replacement, rev_cell, flags=re.IGNORECASE)
                                matched_rules.append((rule_title, reason, legal_basis, severity, excerpt))
                        if matched_rules and rev_cell != cell_str:
                            finding_seq += 1
                            fid = f"FND_{finding_seq:03d}"
                            top_rule = matched_rules[0]
                            preview_diff = compute_word_diff(cell_str, rev_cell)
                            findings.append({
                                "id": fid,
                                "type": "VULNERABILIDADE_JURIDICA",
                                "finding_type": "VULNERABILIDADE_JURIDICA",
                                "severity": top_rule[3],
                                "title": f"{top_rule[0]} (Tabela)",
                                "section_title": f"Tabela — Linha {r_i + 1}, Coluna {c_i + 1}",
                                "clause_reference": f"Tabela (L{r_i + 1}/C{c_i + 1})",
                                "original_clause": f"Tabela (L{r_i + 1}/C{c_i + 1})",
                                "target_type": "table_cell",
                                "element_idx": el_idx,
                                "docx_tbl_idx": el.get("docx_tbl_idx"),
                                "row_idx": r_i,
                                "col_idx": c_i,
                                "original_text": cell_str,
                                "original_excerpt": top_rule[4],
                                "suggested_text": rev_cell,
                                "agu_standard_clause": rev_cell,
                                "explanation": top_rule[1],
                                "reason": top_rule[1],
                                "legal_foundation": top_rule[2],
                                "preview_diff": preview_diff,
                                "status": "approved" if not require_permission else "pending",
                                "approved": not require_permission,
                            })

        # 3. Detecção Top-Down de Omissões Obrigatórias (AGU / Lei 14.133/2021)
        all_text_parts = []
        for e in elements:
            if e.get("text"):
                all_text_parts.append(str(e["text"]))
            if e.get("title"):
                all_text_parts.append(str(e["title"]))
            if e.get("type") == "TABLE":
                for row in ([e.get("headers") or []] + list(e.get("rows") or [])):
                    all_text_parts.extend(str(c) for c in row)
        all_text = " ".join(all_text_parts)

        for om in cls.MANDATORY_OMISSIONS:
            if not re.search(om["regex"], all_text, flags=re.IGNORECASE):
                finding_seq += 1
                fid = f"FND_{finding_seq:03d}"
                f_om = {
                    "id": fid,
                    "type": "OMISSAO_OBRIGATORIA",
                    "finding_type": "OMISSAO_OBRIGATORIA",
                    "severity": om["severity"],
                    "title": om["title"],
                    "section_title": om["title"],
                    "clause_reference": om["clause_num"],
                    "original_clause": om["clause_num"],
                    "target_type": "append_clause",
                    "element_idx": None,
                    "docx_p_idx": None,
                    "original_text": "",
                    "original_excerpt": "(Cláusula ausente na minuta original)",
                    "suggested_text": om["text"],
                    "agu_standard_clause": om["text"],
                    "explanation": om["explanation"],
                    "reason": om["explanation"],
                    "legal_foundation": om["legal_foundation"],
                    "preview_diff": [{"type": "added", "text": om["text"]}],
                    "status": "approved" if not require_permission else "pending",
                    "approved": not require_permission,
                }
                findings.append(f_om)

                if not require_permission:
                    elements.append({
                        "type": "HEADING_1",
                        "number": "+",
                        "title": om["title"],
                        "text": om["title"],
                        "page": 1,
                    })
                    elements.append({
                        "type": "CLAUSE",
                        "clause_number": om["clause_num"],
                        "section_title": om["title"],
                        "text": om["text"],
                        "modified": True,
                        "diff_segments": [{"type": "added", "text": om["text"]}],
                        "page": 1,
                    })

        crit_cnt = sum(1 for f in findings if f["severity"] == "CRITICA")
        high_cnt = sum(1 for f in findings if f["severity"] == "ALTA")
        med_cnt = sum(1 for f in findings if f["severity"] == "MEDIA")
        compliance_score = max(0.0, round(100.0 - (crit_cnt * 15.0 + high_cnt * 10.0 + med_cnt * 5.0), 1))

        doc_structure["findings"] = findings
        doc_structure["audit_summary"] = {
            "total_clauses": total_clauses,
            "modified_clauses": sum(1 for f in findings if f["approved"]),
            "findings_count": len(findings),
            "pending_count": sum(1 for f in findings if not f["approved"]),
            "compliance_score": compliance_score,
            "findings": findings,
        }
        return doc_structure

    @classmethod
    def apply_approved_decisions(
        cls,
        doc_structure: Dict[str, Any],
        approved_decisions: Optional[List[Dict[str, Any]]] = None,
    ) -> Tuple[Dict[str, Any], List[Dict[str, Any]]]:
        """
        Aplica EXCLUSIVAMENTE as alterações expressamente permitidas pelo usuário
        sobre uma cópia limpa de `doc_structure`.
        `approved_decisions`: lista de { 'id': 'FND_001', 'approved': True, 'custom_text': '...' (opcional) }
        Retorna (resolved_doc_structure, list_of_applied_actions).
        """
        resolved = copy.deepcopy(doc_structure)
        elements = resolved.get("elements", [])
        findings = resolved.get("findings") or (resolved.get("audit_summary") or {}).get("findings") or []

        # Limpa qualquer diff anterior nos elementos para aplicar estritamente o que foi aprovado
        for el in elements:
            if el.get("type") in ("CLAUSE", "PARAGRAPH"):
                base_txt = el.get("original_text") if el.get("original_text") is not None else el.get("text", "")
                el["text"] = base_txt
                el["modified"] = False
                el["diff_segments"] = [{"type": "equal", "text": base_txt}]
            elif el.get("type") == "TABLE":
                el["cell_diffs"] = {}

        # Mapa de decisões do usuário
        decision_map: Dict[str, Dict[str, Any]] = {}
        if approved_decisions is not None:
            for d in approved_decisions:
                fid = str(d.get("id") or "")
                if fid:
                    decision_map[fid] = d

        applied_actions: List[Dict[str, Any]] = []

        for f in findings:
            fid = str(f.get("id") or "")
            if approved_decisions is not None:
                dec = decision_map.get(fid)
                is_approved = bool(dec and dec.get("approved"))
                custom_txt = (dec.get("custom_text") if dec else None)
            else:
                is_approved = bool(f.get("approved"))
                custom_txt = f.get("custom_text")

            f["approved"] = is_approved
            f["status"] = "approved" if is_approved else "rejected"
            if not is_approved:
                continue

            new_text = str(custom_txt).strip() if (custom_txt and str(custom_txt).strip()) else str(f.get("suggested_text") or f.get("agu_standard_clause") or "")
            orig_text = str(f.get("original_text") or "")
            target_type = f.get("target_type") or "element"
            diff_segs = compute_word_diff(orig_text, new_text)

            action = {
                "id": fid,
                "target_type": target_type,
                "element_idx": f.get("element_idx"),
                "docx_p_idx": f.get("docx_p_idx"),
                "docx_tbl_idx": f.get("docx_tbl_idx"),
                "row_idx": f.get("row_idx"),
                "col_idx": f.get("col_idx"),
                "title": f.get("title") or f.get("section_title") or "Cláusula Inserida (AGU)",
                "clause_reference": f.get("clause_reference") or "",
                "original_text": orig_text,
                "new_text": new_text,
                "diff_segments": diff_segs,
            }
            applied_actions.append(action)

            if target_type == "element":
                e_idx = f.get("element_idx")
                if e_idx is not None and 0 <= e_idx < len(elements):
                    el = elements[e_idx]
                    el["modified"] = True
                    el["revised_text"] = new_text
                    el["diff_segments"] = diff_segs
            elif target_type == "table_cell":
                e_idx = f.get("element_idx")
                r_i = f.get("row_idx")
                c_i = f.get("col_idx")
                if e_idx is not None and 0 <= e_idx < len(elements) and r_i is not None and c_i is not None:
                    el = elements[e_idx]
                    el.setdefault("cell_diffs", {})[f"{r_i}_{c_i}"] = diff_segs
            elif target_type == "append_clause":
                elements.append({
                    "type": "HEADING_1",
                    "number": "+",
                    "title": action["title"],
                    "text": action["title"],
                    "page": 1,
                    "is_inserted": True,
                })
                elements.append({
                    "type": "CLAUSE",
                    "clause_number": action["clause_reference"],
                    "section_title": action["title"],
                    "text": new_text,
                    "modified": True,
                    "diff_segments": [{"type": "added", "text": new_text}],
                    "page": 1,
                    "is_inserted": True,
                })

        if "audit_summary" in resolved:
            resolved["audit_summary"]["modified_clauses"] = len(applied_actions)
        return resolved, applied_actions


# ==============================================================================
# 4. EXPORTADOR CIRÚRGICO WORD (.DOCX) — PRESERVAÇÃO 100% DO FORMATO ORIGINAL
# ==============================================================================
def _capture_paragraph_run_style(paragraph) -> Dict[str, Any]:
    """Captura fonte, tamanho, cor e itálico predominantes do parágrafo original."""
    style_info = {"name": None, "size": None, "bold": None, "italic": None, "color_rgb": None}
    for r in paragraph.runs:
        if r.text and r.text.strip():
            style_info["name"] = r.font.name
            style_info["size"] = r.font.size
            style_info["bold"] = r.font.bold
            style_info["italic"] = r.font.italic
            if r.font.color and r.font.color.rgb:
                style_info["color_rgb"] = r.font.color.rgb
            break
    return style_info


def _rewrite_paragraph_with_diff_preserving_style(paragraph, diff_segments: List[Dict[str, str]]) -> None:
    """
    Edita cirurgicamente os runs de um parágrafo nativo do python-docx mantendo
    INTACTO `<w:pPr>` (estilo, alinhamento, recuos, espaçamento, numeração e bordas!)
    e aplicando:
    - equal: exatamente a fonte/tamanho/cor original do parágrafo.
    - removed: cor vermelha (#DC2626) + tachado (strike = True).
    - added: destaque amarelo (WD_COLOR_INDEX.YELLOW) + negrito.
    """
    orig_style = _capture_paragraph_run_style(paragraph)

    # Remove apenas nós de texto/run antigos, preservando <w:pPr> intacto!
    p_el = paragraph._p
    for child in list(p_el):
        tag = child.tag.split("}")[-1] if "}" in child.tag else child.tag
        if tag in ("r", "hyperlink", "ins", "del"):
            p_el.remove(child)

    for idx, seg in enumerate(diff_segments):
        st = seg.get("type", "equal")
        txt = str(seg.get("text") or "")
        if not txt:
            continue
        suffix = " " if idx < len(diff_segments) - 1 else ""
        r = paragraph.add_run(txt + suffix)
        if orig_style["name"]:
            r.font.name = orig_style["name"]
        if orig_style["size"]:
            r.font.size = orig_style["size"]
        if orig_style["italic"] is not None:
            r.font.italic = orig_style["italic"]

        if st == "added":
            r.font.bold = True
            r.font.highlight_color = WD_COLOR_INDEX.YELLOW
            r.font.color.rgb = RGBColor(15, 23, 42)
        elif st == "removed":
            r.font.strike = True
            r.font.color.rgb = RGBColor(220, 38, 38)
        else:
            if orig_style["bold"] is not None:
                r.font.bold = orig_style["bold"]
            if orig_style["color_rgb"] is not None:
                r.font.color.rgb = orig_style["color_rgb"]


def apply_surgical_edits_to_original_docx(
    original_docx_bytes: bytes,
    applied_actions: List[Dict[str, Any]],
) -> bytes:
    """
    MODO CIRÚRGICO REAL PARA .DOCX ORIGINAL:
    Abre o arquivo .DOCX original intacto (mantendo cabeçalhos, rodapés, brasões,
    estilos, margens, colunas, bordas e tabelas originais 100% preservados) e aplica
    SOMENTE as alterações expressamente permitidas pelo usuário com marcação:
    - AMARELO: alterado / inserido
    - VERMELHO TACHADO: retirado
    """
    doc = docx.Document(io.BytesIO(original_docx_bytes))

    # Mapeia parágrafos e tabelas do body na mesma ordem de DocumentLayoutParser.parse_docx
    body_paragraphs = []
    body_tables = []
    for block in doc.element.body:
        tag = block.tag.split("}")[-1] if "}" in block.tag else block.tag
        if tag == "p":
            body_paragraphs.append(docx.text.paragraph.Paragraph(block, doc))
        elif tag == "tbl":
            body_tables.append(docx.table.Table(block, doc))

    # Detecta fonte predominante do documento original para novas cláusulas inseridas
    default_font_name = "Calibri"
    default_font_size = Pt(10.5)
    for bp in body_paragraphs:
        st = _capture_paragraph_run_style(bp)
        if st["name"]:
            default_font_name = st["name"]
            if st["size"]:
                default_font_size = st["size"]
            break

    # Aplica cada ação aprovada no local exato
    appended_clauses = []
    for act in applied_actions:
        t_type = act.get("target_type")
        diff_segs = act.get("diff_segments") or []

        if t_type == "element":
            p_idx = act.get("docx_p_idx")
            target_p = None
            if p_idx is not None and 0 <= p_idx < len(body_paragraphs):
                target_p = body_paragraphs[p_idx]
            else:
                # Fallback por busca de texto exato
                orig_t = (act.get("original_text") or "").strip()
                if orig_t:
                    for bp in body_paragraphs:
                        if bp.text.strip() == orig_t or orig_t[:60] in bp.text:
                            target_p = bp
                            break
            if target_p is not None and diff_segs:
                _rewrite_paragraph_with_diff_preserving_style(target_p, diff_segs)

        elif t_type == "table_cell":
            t_idx = act.get("docx_tbl_idx")
            r_i = act.get("row_idx")
            c_i = act.get("col_idx")
            if t_idx is not None and 0 <= t_idx < len(body_tables):
                tbl = body_tables[t_idx]
                if r_i is not None and c_i is not None and r_i < len(tbl.rows) and c_i < len(tbl.rows[r_i].cells):
                    cell = tbl.rows[r_i].cells[c_i]
                    if cell.paragraphs:
                        _rewrite_paragraph_with_diff_preserving_style(cell.paragraphs[0], diff_segs)

        elif t_type == "append_clause":
            appended_clauses.append(act)

    # Insere cláusulas obrigatórias aprovadas no final do documento mantendo a fonte original
    for act in appended_clauses:
        p_h = doc.add_paragraph()
        p_h.paragraph_format.space_before = Pt(10)
        p_h.paragraph_format.space_after = Pt(4)
        p_h.paragraph_format.keep_with_next = True
        r_h = p_h.add_run(f"{act.get('title', 'CLÁUSULA INCLUÍDA (LEI Nº 14.133/2021)')}")
        r_h.font.name = default_font_name
        r_h.font.size = default_font_size
        r_h.font.bold = True
        r_h.font.highlight_color = WD_COLOR_INDEX.YELLOW

        p_c = doc.add_paragraph()
        p_c.paragraph_format.space_before = Pt(2)
        p_c.paragraph_format.space_after = Pt(6)
        p_c.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        cl_ref = act.get("clause_reference")
        prefix = f"{cl_ref}: " if cl_ref else ""
        r_c = p_c.add_run(f"{prefix}{act.get('new_text', '')}")
        r_c.font.name = default_font_name
        r_c.font.size = default_font_size
        r_c.font.bold = True
        r_c.font.highlight_color = WD_COLOR_INDEX.YELLOW

    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()


def build_audited_document_docx_bytes(
    doc_structure: Dict[str, Any],
    org_name: str = "MARINHA DO BRASIL",
    original_bytes: Optional[bytes] = None,
    applied_actions: Optional[List[Dict[str, Any]]] = None,
) -> bytes:
    """
    Exporta o documento revisado em Microsoft Word (.DOCX):
    - Se `original_bytes` for um .DOCX original: executa cirurgia XML in-place (100% do layout original preservado!).
    - Caso contrário (PDF/TXT): reconstrói um .DOCX fiel preservando tamanho de página, margens,
      fontes/tamanhos reais, alinhamentos e larguras proporcionais de tabelas do PDF original!
    """
    if original_bytes and str(doc_structure.get("original_format", "")).lower() == "docx" and applied_actions is not None:
        return apply_surgical_edits_to_original_docx(original_bytes, applied_actions)

    doc = docx.Document()
    pl = doc_structure.get("page_layout") or {}
    default_font = pl.get("default_font") or "Arial"
    default_size = float(pl.get("default_size_pt") or 10.0)

    # Preserva dimensões e margens reais do PDF original
    for section in doc.sections:
        if pl.get("width_pt") and pl.get("height_pt"):
            section.page_width = Pt(float(pl["width_pt"]))
            section.page_height = Pt(float(pl["height_pt"]))
        section.top_margin = Pt(float(pl.get("margin_top_pt", 54.0)))
        section.bottom_margin = Pt(float(pl.get("margin_bottom_pt", 54.0)))
        section.left_margin = Pt(float(pl.get("margin_left_pt", 54.0)))
        section.right_margin = Pt(float(pl.get("margin_right_pt", 54.0)))

    usable_width_pt = float(pl.get("width_pt", 595.3)) - float(pl.get("margin_left_pt", 54.0)) - float(pl.get("margin_right_pt", 54.0))

    align_map = {
        "CENTER": WD_ALIGN_PARAGRAPH.CENTER,
        "RIGHT": WD_ALIGN_PARAGRAPH.RIGHT,
        "JUSTIFY": WD_ALIGN_PARAGRAPH.JUSTIFY,
        "LEFT": WD_ALIGN_PARAGRAPH.LEFT,
    }

    # 1. Brasão Oficial da República
    brasao_path = os.path.join(BASE_DIR, "brasao_oficial.png")
    if os.path.exists(brasao_path):
        p_img = doc.add_paragraph()
        p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_img.paragraph_format.space_before = Pt(0)
        p_img.paragraph_format.space_after = Pt(2)
        p_img.add_run().add_picture(brasao_path, width=Inches(0.95))

    # 2. Cabeçalho do Órgão + Título Original
    p_head = doc.add_paragraph()
    p_head.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_head.paragraph_format.line_spacing = 1.08
    p_head.paragraph_format.space_after = Pt(4)
    uasg_str = f"UASG {doc_structure['uasg']} — " if doc_structure.get("uasg") else ""
    r_head = p_head.add_run(f"{uasg_str}{doc_structure.get('orgao') or org_name}\n")
    r_head.font.name = default_font
    r_head.font.size = Pt(10)
    r_head.font.bold = True
    r_head.font.color.rgb = RGBColor(30, 45, 61)

    p_tit = doc.add_paragraph()
    p_tit.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_tit.paragraph_format.space_after = Pt(6)
    r_tit = p_tit.add_run(doc_structure.get("title", "DOCUMENTO ADMINISTRATIVO AUDITADO") + "\n")
    r_tit.font.name = default_font
    r_tit.font.size = Pt(13)
    r_tit.font.bold = True
    r_tit.font.color.rgb = RGBColor(15, 23, 42)

    if doc_structure.get("processo"):
        r_proc = p_tit.add_run(f"Processo Administrativo: {doc_structure.get('processo')}")
        r_proc.font.name = default_font
        r_proc.font.size = Pt(9.5)
        r_proc.font.italic = True

    # 3. Legenda discreta apenas se houver alterações aprovadas
    has_any_change = any(
        el.get("modified") or el.get("cell_diffs") or el.get("is_inserted")
        for el in doc_structure.get("elements", [])
    )
    if has_any_change:
        p_leg = doc.add_paragraph()
        p_leg.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_leg.paragraph_format.space_before = Pt(2)
        p_leg.paragraph_format.space_after = Pt(8)
        r0 = p_leg.add_run("Legenda de Revisão (Lei 14.133/2021): ")
        r0.font.name = default_font
        r0.font.size = Pt(8)
        r0.font.bold = True
        r1 = p_leg.add_run("Amarelo = Alterado/Inserido")
        r1.font.name = default_font
        r1.font.size = Pt(8)
        r1.font.bold = True
        r1.font.highlight_color = WD_COLOR_INDEX.YELLOW
        r_sep = p_leg.add_run("   |   ")
        r_sep.font.name = default_font
        r_sep.font.size = Pt(8)
        r2 = p_leg.add_run("Vermelho Tachado = Retirado")
        r2.font.name = default_font
        r2.font.size = Pt(8)
        r2.font.strike = True
        r2.font.color.rgb = RGBColor(220, 38, 38)

    # 4. Elementos na ordem e formatação fiel do documento original
    for el in doc_structure.get("elements", []):
        el_type = el.get("type")
        st_meta = el.get("style") or {}
        f_name = st_meta.get("font_family") or default_font
        f_size = float(st_meta.get("font_size") or default_size)
        p_align = align_map.get(st_meta.get("alignment", "LEFT"), WD_ALIGN_PARAGRAPH.JUSTIFY)

        if el_type == "HEADING_1":
            p = doc.add_paragraph()
            p.paragraph_format.space_before = Pt(11)
            p.paragraph_format.space_after = Pt(4)
            p.paragraph_format.keep_with_next = True
            sec_num = el.get("number", "")
            sec_title = el.get("title") or el.get("text") or ""
            txt = f"{sec_num}. {sec_title}" if sec_num and sec_num != "+" and not sec_title.startswith(f"{sec_num}.") else sec_title
            r = p.add_run(txt)
            r.font.name = f_name
            r.font.size = Pt(max(11.0, min(14.0, f_size)))
            r.font.bold = True
            r.font.color.rgb = RGBColor(30, 45, 61)
            if el.get("is_inserted"):
                r.font.highlight_color = WD_COLOR_INDEX.YELLOW

        elif el_type == "HEADING_2":
            p = doc.add_paragraph()
            p.paragraph_format.space_before = Pt(7)
            p.paragraph_format.space_after = Pt(2)
            p.paragraph_format.keep_with_next = True
            r = p.add_run(el.get("title") or el.get("text") or "")
            r.font.name = f_name
            r.font.size = Pt(max(10.0, min(12.0, f_size)))
            r.font.bold = True
            r.font.color.rgb = RGBColor(51, 65, 85)

        elif el_type in ("CLAUSE", "PARAGRAPH"):
            p = doc.add_paragraph()
            p.alignment = p_align if st_meta.get("alignment") else WD_ALIGN_PARAGRAPH.JUSTIFY
            p.paragraph_format.space_before = Pt(2)
            p.paragraph_format.space_after = Pt(4)
            p.paragraph_format.line_spacing = 1.15

            cl_num = el.get("clause_number", "")
            diffs = el.get("diff_segments") or [{"type": "equal", "text": el.get("text", "")}]
            first_seg_text = (diffs[0].get("text") or "") if diffs else ""
            if cl_num and not first_seg_text.startswith(f"{cl_num}"):
                r_num = p.add_run(f"{cl_num}. ")
                r_num.font.name = f_name
                r_num.font.size = Pt(f_size)
                r_num.font.bold = True
                if el.get("is_inserted"):
                    r_num.font.highlight_color = WD_COLOR_INDEX.YELLOW

            for seg_i, seg in enumerate(diffs):
                st = seg.get("type", "equal")
                txt = str(seg.get("text") or "")
                if not txt:
                    continue
                suffix = " " if seg_i < len(diffs) - 1 else ""
                r = p.add_run(txt + suffix)
                r.font.name = f_name
                r.font.size = Pt(f_size)
                if st == "added":
                    r.font.bold = True
                    r.font.highlight_color = WD_COLOR_INDEX.YELLOW
                    r.font.color.rgb = RGBColor(15, 23, 42)
                elif st == "removed":
                    r.font.strike = True
                    r.font.color.rgb = RGBColor(220, 38, 38)
                else:
                    if st_meta.get("is_bold"):
                        r.font.bold = True
                    if st_meta.get("is_italic"):
                        r.font.italic = True

        elif el_type == "TABLE":
            headers = el.get("headers") or []
            rows = el.get("rows") or []
            if not headers and not rows:
                continue
            num_cols = max(len(headers), max((len(r) for r in rows), default=0))
            if num_cols == 0:
                continue

            col_ratios = el.get("col_ratios") or [1.0 / num_cols] * num_cols
            if len(col_ratios) < num_cols:
                col_ratios = [1.0 / num_cols] * num_cols

            tbl = doc.add_table(rows=len(rows) + (1 if headers else 0), cols=num_cols)
            tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
            tbl.autofit = False
            set_docx_table_borders(tbl, color="CBD5E1", sz="4")

            cell_diffs = el.get("cell_diffs") or {}

            def _fill_table_cell(cell, row_abs_idx: int, col_idx: int, raw_val: str, is_header: bool):
                cell.width = Pt(max(36.0, usable_width_pt * col_ratios[col_idx]))
                set_docx_cell_margins(cell, 60, 60, 90, 90)
                if is_header:
                    set_docx_cell_background(cell, "1E2D3D")
                elif row_abs_idx % 2 == 0:
                    set_docx_cell_background(cell, "F8FAFC")

                p_cell = cell.paragraphs[0]
                p_cell.paragraph_format.space_before = Pt(1)
                p_cell.paragraph_format.space_after = Pt(1)
                c_diff = cell_diffs.get(f"{row_abs_idx}_{col_idx}")
                if c_diff:
                    for si, seg in enumerate(c_diff):
                        st = seg.get("type", "equal")
                        txt = str(seg.get("text") or "") + (" " if si < len(c_diff) - 1 else "")
                        r = p_cell.add_run(txt)
                        r.font.name = default_font
                        r.font.size = Pt(8.5)
                        if st == "added":
                            r.font.bold = True
                            r.font.highlight_color = WD_COLOR_INDEX.YELLOW
                            r.font.color.rgb = RGBColor(15, 23, 42)
                        elif st == "removed":
                            r.font.strike = True
                            r.font.color.rgb = RGBColor(220, 38, 38)
                        elif is_header:
                            r.font.bold = True
                            r.font.color.rgb = RGBColor(255, 255, 255)
                else:
                    r = p_cell.add_run(str(raw_val or ""))
                    r.font.name = default_font
                    r.font.size = Pt(8.5)
                    if is_header:
                        r.font.bold = True
                        r.font.color.rgb = RGBColor(255, 255, 255)
                    else:
                        r.font.color.rgb = RGBColor(30, 41, 59)

            curr_row = 0
            if headers:
                for col_idx in range(num_cols):
                    val = headers[col_idx] if col_idx < len(headers) else ""
                    _fill_table_cell(tbl.cell(curr_row, col_idx), 0, col_idx, val, is_header=True)
                curr_row += 1

            for row_idx, r_data in enumerate(rows):
                abs_r = row_idx + (1 if headers else 0)
                for col_idx in range(num_cols):
                    val = r_data[col_idx] if col_idx < len(r_data) else ""
                    _fill_table_cell(tbl.cell(curr_row, col_idx), abs_r, col_idx, val, is_header=False)
                curr_row += 1

            doc.add_paragraph().paragraph_format.space_after = Pt(4)

    # 5. Assinaturas originais
    signatures = doc_structure.get("signatures", [])
    if signatures:
        p_sig_title = doc.add_paragraph()
        p_sig_title.paragraph_format.space_before = Pt(16)
        p_sig_title.paragraph_format.space_after = Pt(8)
        p_sig_title.paragraph_format.keep_with_next = True
        r_sig_t = p_sig_title.add_run("ASSINATURAS DOS RESPONSÁVEIS:")
        r_sig_t.font.name = default_font
        r_sig_t.font.size = Pt(10)
        r_sig_t.font.bold = True

        for sig in signatures:
            p_s = doc.add_paragraph()
            p_s.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p_s.paragraph_format.space_before = Pt(6)
            p_s.paragraph_format.space_after = Pt(4)
            p_s.add_run("_____________________________________________________\n").font.name = default_font
            r_name = p_s.add_run(f"{sig.get('name', 'RESPONSÁVEL TÉCNICO')}\n")
            r_name.font.name = default_font
            r_name.font.size = Pt(9.5)
            r_name.font.bold = True
            r_role = p_s.add_run(sig.get("role", "Membro da Equipe de Planejamento"))
            r_role.font.name = default_font
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
        if self._pageNumber > 1:
            self.drawString(54, 800, "MINUTA ADMINISTRATIVA AUDITADA — LEI Nº 14.133/2021 & AGU")
            self.setStrokeColor(colors.HexColor("#CBD5E1"))
            self.setLineWidth(0.5)
            self.line(54, 794, 541, 794)
        self.drawRightString(541, 36, f"Página {self._pageNumber} de {total_pages}")
        self.drawString(54, 36, "Documento Auditado Eletronicamente com Controle de Permissão — Padrão AGU / TCU")
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

    style_head = ParagraphStyle("GovHead", parent=styles["Normal"], fontName="Helvetica-Bold", fontSize=9, leading=11, alignment=1, textColor=colors.HexColor("#1E2D3D"), spaceAfter=4)
    style_title = ParagraphStyle("GovTitle", parent=styles["Normal"], fontName="Helvetica-Bold", fontSize=11.5, leading=14, alignment=1, textColor=colors.HexColor("#1E2D3D"), spaceBefore=4, spaceAfter=6)
    style_h1 = ParagraphStyle("GovH1", parent=styles["Normal"], fontName="Helvetica-Bold", fontSize=10.5, leading=13, textColor=colors.HexColor("#1E2D3D"), spaceBefore=10, spaceAfter=4, keepWithNext=True)
    style_h2 = ParagraphStyle("GovH2", parent=styles["Normal"], fontName="Helvetica-Bold", fontSize=9.5, leading=12, textColor=colors.HexColor("#334155"), spaceBefore=6, spaceAfter=3, keepWithNext=True)
    style_body = ParagraphStyle("GovBody", parent=styles["Normal"], fontName="Helvetica", fontSize=9, leading=12.5, alignment=4, spaceBefore=2, spaceAfter=4, textColor=colors.HexColor("#0F172A"))
    style_th = ParagraphStyle("GovTH", parent=styles["Normal"], fontName="Helvetica-Bold", fontSize=8, leading=10, alignment=1, textColor=colors.white)
    style_td = ParagraphStyle("GovTD", parent=styles["Normal"], fontName="Helvetica", fontSize=7.5, leading=9.5, textColor=colors.HexColor("#1E293B"))

    story = []
    brasao_path = os.path.join(BASE_DIR, "brasao_oficial.png")
    if os.path.exists(brasao_path):
        story.append(RLImage(brasao_path, width=1.0 * 72, height=1.0 * 72))
        story.append(Spacer(1, 4))

    orgao_val = doc_structure.get("orgao") or org_name
    story.append(Paragraph(f"REPÚBLICA FEDERATIVA DO BRASIL<br/>{xml_escape(orgao_val)}", style_head))
    story.append(Paragraph(xml_escape(doc_structure.get("title", "DOCUMENTO ADMINISTRATIVO AUDITADO")), style_title))

    if doc_structure.get("processo"):
        story.append(Paragraph(f"<i>Processo Administrativo: {xml_escape(doc_structure.get('processo'))}</i>", ParagraphStyle("Proc", parent=style_head, fontName="Helvetica-Oblique", fontSize=8.5, leading=10, spaceAfter=8)))

    legend_html = (
        "<b>CONVENÇÃO DE REVISÃO (LEI Nº 14.133/2021 &amp; AGU):</b> "
        "<font backcolor=\"#FEF08A\"><b>[AMARELO]</b></font> = Alterado/Inserido com sua permissão. &nbsp;&nbsp;"
        "<font color=\"#DC2626\"><strike>[VERMELHO TACHADO]</strike></font> = Retirado/Revogado."
    )
    t_leg = Table([[Paragraph(legend_html, ParagraphStyle("Leg", parent=styles["Normal"], fontName="Helvetica", fontSize=8, leading=10.5, textColor=colors.HexColor("#1E293B")))]], colWidths=[487])
    t_leg.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#F1F5F9")), ("BOX", (0, 0), (-1, -1), 0.5, colors.HexColor("#CBD5E1")), ("PADDING", (0, 0), (-1, -1), 6)]))
    story.append(t_leg)
    story.append(Spacer(1, 8))

    def _diff_to_rl_html(diffs: List[Dict[str, str]]) -> str:
        parts = []
        for seg in diffs:
            st = seg.get("type", "equal")
            txt = xml_escape(seg.get("text", ""))
            if st == "added":
                parts.append(f"<font backcolor=\"#FEF08A\"><b>{txt}</b></font>")
            elif st == "removed":
                parts.append(f"<font color=\"#DC2626\"><strike>{txt}</strike></font>")
            else:
                parts.append(txt)
        return " ".join(parts)

    for el in doc_structure.get("elements", []):
        el_type = el.get("type")
        if el_type == "HEADING_1":
            sec_num = el.get("number", "")
            sec_title = el.get("title") or el.get("text") or ""
            title_text = f"{sec_num}. {sec_title}" if sec_num and sec_num != "+" and not sec_title.startswith(f"{sec_num}.") else sec_title
            if el.get("is_inserted"):
                story.append(Paragraph(f"<font backcolor=\"#FEF08A\"><b>{xml_escape(title_text)}</b></font>", style_h1))
            else:
                story.append(Paragraph(xml_escape(title_text), style_h1))
        elif el_type == "HEADING_2":
            story.append(Paragraph(xml_escape(el.get("title") or el.get("text") or ""), style_h2))
        elif el_type in ("CLAUSE", "PARAGRAPH"):
            cl_num = el.get("clause_number", "")
            diffs = el.get("diff_segments") or [{"type": "equal", "text": el.get("text", "")}]
            first_t = (diffs[0].get("text") or "") if diffs else ""
            prefix = ""
            if cl_num and not first_t.startswith(f"{cl_num}"):
                prefix = f"<font backcolor=\"#FEF08A\"><b>{xml_escape(cl_num)}.</b></font> " if el.get("is_inserted") else f"<b>{xml_escape(cl_num)}.</b> "
            story.append(Paragraph(prefix + _diff_to_rl_html(diffs), style_body))
        elif el_type == "TABLE":
            headers = el.get("headers") or []
            rows = el.get("rows") or []
            if not headers and not rows:
                continue
            num_cols = max(len(headers), max((len(r) for r in rows), default=0))
            if num_cols == 0:
                continue
            col_ratios = el.get("col_ratios") or [1.0 / num_cols] * num_cols
            if len(col_ratios) < num_cols:
                col_ratios = [1.0 / num_cols] * num_cols
            col_widths = [max(32.0, 487.0 * r) for r in col_ratios[:num_cols]]
            cell_diffs = el.get("cell_diffs") or {}

            table_data = []
            if headers:
                table_data.append([
                    Paragraph(_diff_to_rl_html(cell_diffs[f"0_{ci}"]) if f"0_{ci}" in cell_diffs else f"<b>{xml_escape(str(h))}</b>", style_th)
                    for ci, h in enumerate(headers[:num_cols])
                ])
            for ri, r in enumerate(rows):
                abs_r = ri + (1 if headers else 0)
                row_cells = []
                for ci in range(num_cols):
                    val = r[ci] if ci < len(r) else ""
                    if f"{abs_r}_{ci}" in cell_diffs:
                        row_cells.append(Paragraph(_diff_to_rl_html(cell_diffs[f"{abs_r}_{ci}"]), style_td))
                    else:
                        row_cells.append(Paragraph(xml_escape(str(val)), style_td))
                table_data.append(row_cells)

            t = Table(table_data, colWidths=col_widths, repeatRows=1 if headers else 0)
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

    signatures = doc_structure.get("signatures", [])
    if signatures:
        story.append(Spacer(1, 14))
        story.append(Paragraph("<b>ASSINATURAS DOS RESPONSÁVEIS:</b>", style_h1))
        for sig in signatures:
            sig_name = xml_escape(sig.get("name", "RESPONSÁVEL TÉCNICO"))
            sig_role = xml_escape(sig.get("role", "Membro da Equipe de Planejamento"))
            story.append(Spacer(1, 6))
            story.append(Paragraph(f"_____________________________________________________<br/><b>{sig_name}</b><br/><i>{sig_role}</i>", style_body))

    doc.build(story, canvasmaker=NumberedCanvas)
    return buf.getvalue()
