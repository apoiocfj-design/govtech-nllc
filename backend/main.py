import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_COLOR_INDEX
# -*- coding: utf-8 -*-
"""
SERVIDOR UNIFICADO DO APLICATIVO GOVTECH NLLC (PORTA 8080)
Centraliza TODAS as funções em um único link estilo App Web (http://127.0.0.1:8080):
1. Interface Web Completa Estilo App (GET /)
2. Pesquisa de Preços Ao Vivo na API do Governo + Cálculo IN 65/2021 (GET /api/v1/prices/in65/{code})
3. Exportação de Planilha Excel Oficial Mapa Comparativo IN 65/2021 (GET /api/v1/prices/export-xlsx/{code})
4. Upload e Auditoria Jurídica RAG de PDF ou DOCX (POST /api/v1/rag/audit-upload)
5. Geração e Download de Arquivo Word (.DOCX) com Alterações Marcadas em Amarelo (POST /api/v1/rag/export-docx)
6. Download Direto dos Documentos já Processados da Capitania (GET /api/v1/files/{filename})
"""

import base64
import gzip
import html
import io
import json
import os
import re
import socket
import sys
import zipfile
from http.server import HTTPServer, BaseHTTPRequestHandler
from urllib.parse import urlparse, parse_qs

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
import pdfplumber

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CORE_DIR = os.path.join(BASE_DIR, "core")
if CORE_DIR not in sys.path:
    sys.path.insert(0, CORE_DIR)

from gov_api_client import PriceResearchController
from rag_pipeline import AGUClauseChunker, RAGLegalAuditor, OpenXMLDocxReader
from catalog_parser import CNBSCatalogEngine, CONSULTA_DETALHADA_LINKS
from report_generators import (
    build_in65_docx_bytes,
    build_in65_pdf_bytes,
    build_in65_justificativa_docx_bytes,
    build_in65_justificativa_pdf_bytes,
    build_consolidated_in65_excel_bytes,
    build_consolidated_in65_pdf_bytes,
    build_consolidated_in65_docx_bytes,
    build_consolidated_suppliers_dossier_pdf_bytes,
    build_batch_zip_to_file,
    build_batch_zip_bytes,
    build_audit_certificate_pdf_bytes,
    generate_statistical_infographic,
    build_supplier_compliance_dossier_docx_bytes,
    build_supplier_compliance_dossier_pdf_bytes,
)
from agu_batch_reports import (
    build_agu_nota_tecnica_docx_bytes,
    build_agu_nota_tecnica_pdf_bytes,
    build_consolidated_suppliers_dossier_docx_bytes,
)
from job_manager import batch_job_manager
from audited_document_engine import (
    DocumentLayoutParser,
    DocumentLegalAuditor,
    build_audited_document_docx_bytes,
    build_audited_document_pdf_bytes,
    document_session_store,
)
from pncp_extractor import (
    resolve_pncp_compra,
    fetch_pncp_arquivos,
    download_pncp_file_bytes,
    generate_pncp_dossier_pdf,
    resolve_pncp_webpage_url,
    capture_pncp_page_pdf,
)
from irp_client import search_irps, get_irp_detail
from sanctions_checker import check_supplier_sanctions
from audit_trail import audit_trail_manager
from pncp_integration import build_pncp_procurement_payload, validate_pncp_payload

price_controller = PriceResearchController()
rag_auditor = RAGLegalAuditor(similarity_threshold=0.45)
rag_auditor.index_knowledge_base()
catalog_engine = CNBSCatalogEngine()

FRONTEND_HTML_PATH = os.path.abspath(os.path.join(BASE_DIR, "..", "frontend", "preview_app.html"))
PESQUISA_DIR = r"C:\Users\CFJ-SECOM\.gemini\antigravity\scratch\pesquisa-precos"
DISPENSA_DIR = os.path.join(PESQUISA_DIR, "documentos_dispensa")


def get_primary_lan_ip() -> str:
    """Identifica o IP local na interface primária de rede (Wi-Fi/Ethernet)."""
    s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
    try:
        s.connect(("8.8.8.8", 80))
        return s.getsockname()[0]
    except Exception:
        return "127.0.0.1"
    finally:
        s.close()

PREBUILT_FILES = {
    "Mapa_Comparativo_800un.xlsx": os.path.join(PESQUISA_DIR, "Mapa_Comparativo_de_Precos_800un_OFICIAL.xlsx"),
    "Relatorio_Ajustes_Andaina_SAR.docx": os.path.join(PESQUISA_DIR, "Relatorio_Ajustes_Lei14133_Andaina_SAR.docx"),
    "Pacote_4_Docs_Editados_Amarelo.zip": os.path.join(PESQUISA_DIR, "Dispensa_Andaina_SAR_Documentos_EDITADOS_AMARELO.zip"),
    "DFD_53_2026_EDITADO_AMARELO.docx": os.path.join(DISPENSA_DIR, "DFD_53_2026_EDITADO_AMARELO.docx"),
    "ETP_24_2026_EDITADO_AMARELO.docx": os.path.join(DISPENSA_DIR, "ETP_24_2026_EDITADO_AMARELO.docx"),
    "TR_42_2026_EDITADO_AMARELO.docx": os.path.join(DISPENSA_DIR, "TR_42_2026_EDITADO_AMARELO.docx"),
    "AC_43_2026_EDITADO_AMARELO.docx": os.path.join(DISPENSA_DIR, "AC_43_2026_EDITADO_AMARELO.docx"),
}


def build_in65_excel_bytes(report: dict, quantity: int, uasg_name: str = "MARINHA DO BRASIL") -> bytes:
    """Gera em memória a planilha Excel Oficial do Mapa Comparativo de Preços (IN 65/2021)."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Mapa Comparativo IN 65"

    borda = Side(style="thin")
    B = Border(left=borda, right=borda, top=borda, bottom=borda)
    cinza = PatternFill("solid", fgColor="D9D9D9")
    verde = PatternFill("solid", fgColor="E2EFDA")
    brl = "R$ #,##0.00"
    ctr = Alignment(horizontal="center", vertical="center", wrap_text=True)
    esq = Alignment(horizontal="left", vertical="center", wrap_text=True)

    def merge_line(row_idx, text, bold=True, align="center", height=22):
        ws.merge_cells(f"A{row_idx}:I{row_idx}")
        ws[f"A{row_idx}"] = text
        ws[f"A{row_idx}"].font = Font(bold=bold, size=11 if bold else 10)
        ws[f"A{row_idx}"].alignment = Alignment(horizontal=align, vertical="center", wrap_text=True)
        ws.row_dimensions[row_idx].height = height

    merge_line(1, uasg_name, bold=True)
    merge_line(2, "MAPA COMPARATIVO DE PREÇOS (CONFORME IN SEGES/ME Nº 65/2021)", bold=True)
    merge_line(3, "1. Série de preços coletados e saneados via API Oficial Compras.gov.br (OAS 3.1):", bold=True, align="left")

    headers = [
        "ITEM", "DESCRIÇÃO DO OBJETO", "CATMAT/CATSER", "UNID.", "QTDE",
        "MÉDIA SANEADA", "MEDIANA SANEADA", "PREÇO UNIT. ESTIMADO", "PREÇO TOTAL ESTIMADO"
    ]
    for col_idx, h in enumerate(headers, 1):
        c = ws.cell(4, col_idx, h)
        c.font, c.fill, c.alignment, c.border = Font(bold=True, size=10), cinza, ctr, B
    ws.row_dimensions[4].height = 28

    sm = report["sanitized_metrics"]
    unit_est = float(report["final_estimated_unit_price"])
    # Fórmulas estatísticas abertas para auditoria do TCU / Controle Interno
    adopted_param = str(report.get("adopted_parameter", "")).upper()
    adopted_formula = "=G5" if "MEDIANA" in adopted_param else "=F5"

    row_vals = [
        1,
        report["item_description"],
        str(report["catmat_catser_code"]),
        report["target_supply_unit"],
        quantity,
        float(sm["mean"] or unit_est),
        float(sm["median"] or unit_est),
        adopted_formula,   # Fórmula aberta: aponta para a mediana (=G5) ou média (=F5)
        "=E5*H5",          # Fórmula aberta: Quantidade x Preço Estimado
    ]
    for col_idx, val in enumerate(row_vals, 1):
        c = ws.cell(5, col_idx, val)
        c.border = B
        c.alignment = esq if col_idx == 2 else ctr
        if col_idx in (6, 7, 8, 9):
            c.number_format = brl
        if col_idx in (8, 9):
            c.fill = verde
            c.font = Font(bold=True)
    ws.row_dimensions[5].height = 65

    merge_line(7, f"2. Parâmetro Adotado: {report['adopted_parameter']} | Coeficiente de Variação Saneado (CV): {sm['cv_percent']}% (Bruto: {report['raw_metrics']['cv_percent']}%)", bold=True, align="left")
    merge_line(8, f"3. Fundamentação Legal e Metodológica (Art. 6º IN 65/2021): {report['legal_methodology_justification']}", bold=False, align="left", height=55)

    for col_letter, width in zip("ABCDEFGHI", [8, 48, 15, 9, 10, 16, 16, 18, 20]):
        ws.column_dimensions[col_letter].width = width

    # Aba 2: Trilha de Auditoria das Cotações Válidas e Expurgadas com Fórmulas Abertas
    ws2 = wb.create_sheet("Auditoria PNCP Compras.gov")
    cols2 = ["Status", "UASG", "Órgão Comprador", "UF", "Nº Compra", "CNPJ Vencedor", "Fornecedor", "Data", "Unid.", "Qtd.", "Preço Unit. (R$)", "Justificativa Estatística", "Link Oficial Auditável", "Rank / Menor Preço"]
    for i, h in enumerate(cols2, 1):
        c = ws2.cell(1, i, h)
        c.font, c.fill, c.alignment, c.border = Font(bold=True), cinza, ctr, B

    all_quotes = report.get("valid_quotes", []) + report.get("purged_quotes", [])
    max_q_row = len(all_quotes) + 1 if all_quotes else 2
    for r_idx, q in enumerate(all_quotes, 2):
        vals = [
            q["purge_status"], q["buyer_uasg"], q["buyer_org_name"], q["buyer_uf"],
            q["purchase_number_year"], q["supplier_cnpj"], q["supplier_name"],
            q["purchase_date"], q["supply_unit"], q["quantity_homologated"],
            float(q["unit_price"]), q.get("purge_justification") or "", q["pncp_or_compras_uri"],
            f"=RANK(K{r_idx}, $K$2:$K${max_q_row}, 1)"  # Fórmula aberta de ranqueamento
        ]
        for c_idx, v in enumerate(vals, 1):
            cel = ws2.cell(r_idx, c_idx, v)
            cel.border = B
            if c_idx == 11:
                cel.number_format = brl
            elif c_idx == 14:
                cel.alignment = ctr
                cel.font = Font(bold=True)
            if c_idx == 13 and v:
                cel.hyperlink = v
                cel.style = "Hyperlink"

    # Seção Estatística Aberta no Rodapé da Trilha de Auditoria
    if all_quotes:
        stat_start = max_q_row + 2
        ws2.merge_cells(f"A{stat_start}:D{stat_start}")
        ws2[f"A{stat_start}"] = "DEMONSTRATIVO ESTATÍSTICO DE FÓRMULAS ABERTAS (ART. 6º DA IN 65/2021)"
        ws2[f"A{stat_start}"].font = Font(bold=True, size=10, color="1E2D3D")

        stat_rows = [
            ("Média Amostral das Cotações Válidas", f'=AVERAGEIF(A2:A{max_q_row}, "VALIDA", K2:K{max_q_row})', brl),
            ("Menor Preço Encontrado no Governo", f"=MIN(K2:K{max_q_row})", brl),
            ("Maior Preço Encontrado no Governo", f"=MAX(K2:K{max_q_row})", brl),
            ("Desvio Padrão Amostral (s)", f"=STDEV.S(K2:K{max_q_row})", brl),
            ("Total de Amostras Homologadas", f"=COUNT(K2:K{max_q_row})", "#,##0"),
            ("Amostras Saneadas Consideradas (Mínimo Legal: 3)", f'=COUNTIF(A2:A{max_q_row}, "VALIDA")', "#,##0"),
        ]
        for offset, (lbl, formula, nfmt) in enumerate(stat_rows, 1):
            curr_r = stat_start + offset
            ws2.cell(curr_r, 2, lbl).font = Font(bold=True, size=9)
            val_c = ws2.cell(curr_r, 4, formula)
            val_c.font = Font(bold=True, size=9)
            val_c.fill = verde if "Média" in lbl or "Saneadas" in lbl else cinza
            val_c.border = B
            val_c.number_format = nfmt

    for col_letter, width in zip("ABCDEFGHIJKLMN", [24, 10, 36, 6, 14, 18, 36, 12, 8, 10, 15, 45, 55, 18]):
        ws2.column_dimensions[col_letter].width = width

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def build_highlighted_docx_bytes(title: str, clauses: list, org_name: str = "MARINHA DO BRASIL") -> bytes:
    """Gera em memória um arquivo Word (.DOCX) oficial com Brasão da República e cláusulas editadas em AMARELO."""
    doc = docx.Document()
    for section in doc.sections:
        section.top_margin = Inches(0.8)
        section.bottom_margin = Inches(0.8)
        section.left_margin = Inches(0.8)
        section.right_margin = Inches(0.8)

    brasao_p = os.path.join(BASE_DIR, "brasao_oficial.png")
    if os.path.exists(brasao_p):
        p_img = doc.add_paragraph()
        p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_img.paragraph_format.space_before = Pt(0)
        p_img.paragraph_format.space_after = Pt(4)
        run_img = p_img.add_run()
        run_img.add_picture(brasao_p, width=Inches(1.1))

    p_head = doc.add_paragraph()
    p_head.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_head.paragraph_format.line_spacing = 1.05
    p_head.paragraph_format.space_after = Pt(12)

    r0 = p_head.add_run(f"REPÚBLICA FEDERATIVA DO BRASIL\nMINISTÉRIO DA DEFESA\nCOMANDO DA MARINHA\n{org_name}\n")
    r0.font.name = "Calibri"
    r0.font.size = Pt(10.5)
    r0.font.bold = True

    p_tit = doc.add_paragraph()
    p_tit.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_tit.paragraph_format.space_after = Pt(10)
    r_tit = p_tit.add_run(title + "\n")
    r_tit.font.name = "Calibri"
    r_tit.font.size = Pt(12)
    r_tit.font.bold = True

    p_note = doc.add_paragraph()
    p_note.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p_note.paragraph_format.space_after = Pt(14)
    r_note = p_note.add_run("NOTA DE AUDITORIA: Todas as cláusulas corrigidas/inseridas pelo Motor RAG (Padrão AGU / Lei nº 14.133/2021) estão destacadas em AMARELO abaixo.")
    r_note.font.name = "Calibri"
    r_note.font.size = Pt(9.5)
    r_note.font.bold = True
    r_note.font.highlight_color = WD_COLOR_INDEX.YELLOW

    for c in clauses:
        num = c.get("num") or c.get("clause_number") or ""
        sec = c.get("title") or c.get("section_title") or ""
        txt = c.get("text") or c.get("content") or ""
        mod = bool(c.get("modified") or c.get("is_modified_by_agu"))

        # Onde tem o nome rascunho, ignora a menção isolada
        if "rascunho" in txt.lower() and len(txt.strip()) < 15:
            continue

        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(6)
        p.paragraph_format.line_spacing = 1.15

        if num or sec:
            r_sec = p.add_run(f"[{num} - {sec}]\n")
            r_sec.font.name = "Calibri"
            r_sec.font.size = Pt(10)
            r_sec.font.bold = True
            if mod:
                r_sec.font.highlight_color = WD_COLOR_INDEX.YELLOW

        r_txt = p.add_run(txt)
        r_txt.font.name = "Calibri"
        r_txt.font.size = Pt(10.5)
        if mod:
            r_txt.font.bold = True
            r_txt.font.highlight_color = WD_COLOR_INDEX.YELLOW

    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()


class GovTechAppHandler(BaseHTTPRequestHandler):
    def _send_json(self, status_code: int, payload: dict):
        body = json.dumps(payload, ensure_ascii=False, separators=(",", ":")).encode("utf-8")
        accept_enc = self.headers.get("Accept-Encoding", "") if self.headers else ""
        use_gzip = ("gzip" in accept_enc) and (len(body) > 4096)
        if use_gzip:
            body = gzip.compress(body, compresslevel=5)
        self.send_response(status_code)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Access-Control-Allow-Origin", "*")
        if use_gzip:
            self.send_header("Content-Encoding", "gzip")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def _send_binary(self, content: bytes, content_type: str, filename: str):
        self.send_response(200)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Disposition", f'attachment; filename="{filename}"')
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Expose-Headers", "Content-Disposition, Content-Length")
        self.send_header("Content-Length", str(len(content)))
        self.end_headers()
        self.wfile.write(content)

    def _send_file_stream(self, file_path: str, content_type: str, filename: str):
        """Envia arquivo grande em streaming (chunks de 64 KB) direto do disco sem estourar RAM."""
        file_size = os.path.getsize(file_path)
        self.send_response(200)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Disposition", f'attachment; filename="{filename}"')
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Expose-Headers", "Content-Disposition, Content-Length")
        self.send_header("Content-Length", str(file_size))
        self.end_headers()
        with open(file_path, "rb") as f:
            while True:
                chunk = f.read(65536)
                if not chunk:
                    break
                self.wfile.write(chunk)

    def do_HEAD(self):
        self.send_response(200)
        self.send_header("Content-Type", "text/html; charset=utf-8")
        self.send_header("Access-Control-Allow-Origin", "*")
        self.end_headers()

    def do_OPTIONS(self):
        self.send_response(200)
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Methods", "GET, POST, OPTIONS, HEAD")
        self.send_header("Access-Control-Allow-Headers", "Content-Type, Authorization, X-Requested-With")
        self.end_headers()

    def do_GET(self):
        parsed = urlparse(self.path)
        path = parsed.path
        qs = parse_qs(parsed.query)

        # 0. Status e Download Streaming de Jobs Assíncronos (/api/v1/jobs/<job_id>/status|download)
        if path.startswith("/api/v1/jobs/"):
            parts = [p for p in path.split("/") if p]
            # ['api', 'v1', 'jobs', '<job_id>', 'status'|'download']
            if len(parts) == 5 and parts[4] == "status":
                job_info = batch_job_manager.get_status(parts[3])
                if not job_info:
                    self._send_json(404, {"error": "Job não encontrado ou expirado."})
                    return
                self._send_json(200, job_info)
                return
            if len(parts) == 5 and parts[4] == "download":
                dl = batch_job_manager.get_download_info(parts[3])
                if not dl:
                    self._send_json(404, {"error": "Arquivo do job não está pronto ou já expirou."})
                    return
                try:
                    self._send_file_stream(dl["file_path"], dl["content_type"], dl["filename"])
                except BrokenPipeError:
                    pass
                return

        # 1. Interface Unificada do Aplicativo
        if path in ("/", "/index.html"):
            with open(FRONTEND_HTML_PATH, "rb") as f:
                raw_content = f.read()
            accept_enc = self.headers.get("Accept-Encoding", "")
            if "gzip" in accept_enc:
                gz_content = gzip.compress(raw_content, compresslevel=6)
                self.send_response(200)
                self.send_header("Content-Type", "text/html; charset=utf-8")
                self.send_header("Content-Encoding", "gzip")
                self.send_header("Content-Length", str(len(gz_content)))
                self.send_header("Access-Control-Allow-Origin", "*")
                self.end_headers()
                self.wfile.write(gz_content)
            else:
                self.send_response(200)
                self.send_header("Content-Type", "text/html; charset=utf-8")
                self.send_header("Content-Length", str(len(raw_content)))
                self.send_header("Access-Control-Allow-Origin", "*")
                self.end_headers()
                self.wfile.write(raw_content)
            return

        # 1.1 Ícone Oficial Transparente do Aplicativo
        if path == "/app_logo_transparent.png":
            logo_path = os.path.join(BASE_DIR, "..", "frontend", "app_logo_transparent.png")
            if os.path.exists(logo_path):
                with open(logo_path, "rb") as f:
                    data = f.read()
                self.send_response(200)
                self.send_header("Content-Type", "image/png")
                self.send_header("Cache-Control", "public, max-age=86400")
                self.send_header("Content-Length", str(len(data)))
                self.send_header("Access-Control-Allow-Origin", "*")
                self.end_headers()
                self.wfile.write(data)
                return
            self._send_json(404, {"error": "Logo não encontrado."})
            return

        # 1.2 Brasão Oficial da República
        if path == "/brasao_oficial.png":
            brasao_path = os.path.join(BASE_DIR, "brasao_oficial.png")
            if not os.path.exists(brasao_path):
                brasao_path = os.path.join(BASE_DIR, "..", "frontend", "brasao_oficial.png")
            if os.path.exists(brasao_path):
                with open(brasao_path, "rb") as f:
                    data = f.read()
                self.send_response(200)
                self.send_header("Content-Type", "image/png")
                self.send_header("Cache-Control", "public, max-age=86400")
                self.send_header("Content-Length", str(len(data)))
                self.send_header("Access-Control-Allow-Origin", "*")
                self.end_headers()
                self.wfile.write(data)
                return
            self._send_json(404, {"error": "Brasão não encontrado."})
            return

        # 2. Healthcheck
        if path == "/api/v1/health":
            self._send_json(200, {
                "status": "ONLINE",
                "app": "GovTech NLLC Unified Web App",
                "catalog": catalog_engine.get_stats(),
            })
            return

        # 2.1 Busca por Palavras-Chave no Catálogo Oficial CNBS (CATMAT + CATSER)
        if path == "/api/v1/catalog/search":
            q = qs.get("q", [""])[0]
            item_type = qs.get("type", ["ALL"])[0]
            limit = int(qs.get("limit", ["25"])[0])
            res = catalog_engine.search(query=q, item_type=item_type, limit=limit)
            self._send_json(200, res)
            return

        # 2.2 Estatísticas do Catálogo CNBS Importado
        if path == "/api/v1/catalog/stats":
            self._send_json(200, catalog_engine.get_stats())
            return

        # 2.3 Links Oficiais da Consulta Detalhada (gov.br/compras)
        if path == "/api/v1/consulta-detalhada/links":
            self._send_json(200, {"categorias": CONSULTA_DETALHADA_LINKS})
            return

        # 2.4 Informações de Rede e Compartilhamento para Outros Usuários
        if path == "/api/v1/network-info":
            lan_ip = get_primary_lan_ip()
            self._send_json(200, {
                "status": "ONLINE",
                "lan_ip": lan_ip,
                "port": 8080,
                "local_url": "http://localhost:8080",
                "lan_url": f"http://{lan_ip}:8080",
                "shareable_instructions": (
                    f"Qualquer pessoa conectada ao mesmo Wi-Fi ou rede corporativa pode acessar o sistema "
                    f"diretamente no navegador pelo endereço: http://{lan_ip}:8080"
                )
            })
            return

        # 2.5 Perfil do Fornecedor e Prevenção de Conluio (CNPJ, CNAE, QSA e Certidões Oficiais)
        if path == "/api/v1/supplier/cnpj-info":
            cnpj_raw = qs.get("cnpj", [""])[0]
            clean_digits = re.sub(r"\D", "", str(cnpj_raw))
            if len(clean_digits) != 14:
                self._send_json(400, {"error": "CNPJ deve conter 14 dígitos numéricos."})
                return

            info = {
                "cnpj": f"{clean_digits[:2]}.{clean_digits[2:5]}.{clean_digits[5:8]}/{clean_digits[8:12]}-{clean_digits[12:]}",
                "razao_social": "Razão Social Não Localizada",
                "nome_fantasia": "",
                "situacao_cadastral": "INDISPONÍVEL",
                "data_situacao_cadastral": "",
                "porte": "N/I",
                "cnae_codigo": "",
                "cnae_descricao": "",
                "cnaes_secundarios": [],
                "qsa": [],
                "municipio": "",
                "uf": "",
                "links_certidoes": {
                    "sicaf": "https://www3.comprasnet.gov.br/sicaf-web/",
                    "cndt": "https://cndt-certidao.tst.jus.br/",
                    "fgts": "https://consulta-crf.caixa.gov.br/consultacrf/pages/consultaEmpregador.jsf",
                    "receita_federal": "https://solucoes.receita.fazenda.gov.br/Servicos/certidaointernet/pj/consultar/",
                    "tcu_inidoneos": "https://contas.tcu.gov.br/certidao/pesquisa/",
                    "portal_transparencia": f"https://portaldatransparencia.gov.br/pessoa-juridica/{clean_digits}"
                }
            }

            import requests
            # Provedor 1: minhareceita.org
            try:
                r = requests.get(f"https://minhareceita.org/{clean_digits}", timeout=5)
                if r.status_code == 200:
                    d = r.json()
                    info["razao_social"] = d.get("razao_social") or ""
                    info["nome_fantasia"] = d.get("nome_fantasia") or ""
                    info["situacao_cadastral"] = d.get("descricao_situacao_cadastral") or "ATIVA"
                    info["data_situacao_cadastral"] = d.get("data_situacao_cadastral") or ""
                    info["porte"] = d.get("porte") or ""
                    info["cnae_codigo"] = str(d.get("cnae_fiscal") or "")
                    info["cnae_descricao"] = d.get("cnae_fiscal_descricao") or ""
                    info["municipio"] = d.get("municipio") or ""
                    info["uf"] = d.get("uf") or ""
                    info["cnaes_secundarios"] = d.get("cnaes_secundarios") or []
                    info["qsa"] = [
                        {"nome": s.get("nome_socio"), "qualificacao": s.get("qualificacao_socio")}
                        for s in d.get("qsa", []) if s.get("nome_socio")
                    ]
                    self._send_json(200, {"status": "SUCCESS", "supplier": info})
                    return
            except Exception:
                pass

            # Provedor 2 (Fallback): receitaws.com.br
            try:
                r = requests.get(f"https://receitaws.com.br/v1/cnpj/{clean_digits}", timeout=5)
                if r.status_code == 200:
                    d = r.json()
                    info["razao_social"] = d.get("nome") or ""
                    info["nome_fantasia"] = d.get("fantasia") or ""
                    info["situacao_cadastral"] = d.get("situacao") or "ATIVA"
                    info["porte"] = d.get("porte") or ""
                    info["cnae_codigo"] = d.get("atividade_principal", [{}])[0].get("code", "")
                    info["cnae_descricao"] = d.get("atividade_principal", [{}])[0].get("text", "")
                    info["municipio"] = d.get("municipio") or ""
                    info["uf"] = d.get("uf") or ""
                    info["qsa"] = [
                        {"nome": s.get("nome"), "qualificacao": s.get("qual")}
                        for s in d.get("qsa", []) if s.get("nome")
                    ]
                    info["sanctions"] = check_supplier_sanctions(clean_digits)
                    self._send_json(200, {"status": "SUCCESS", "supplier": info})
                    return
            except Exception:
                pass

            info["sanctions"] = check_supplier_sanctions(clean_digits)
            self._send_json(200, {"status": "PARTIAL", "supplier": info, "message": "Dados cadastrais resumidos."})
            return

        # 2.6 Inteligência de Risco e Bloqueio Cautelar (CEIS, CNEP, TCU - Art. 14, III e Art. 68 da Lei 14.133/2021)
        if path == "/api/v1/supplier/sanctions-check":
            cnpj_raw = qs.get("cnpj", [""])[0]
            clean_digits = re.sub(r"\D", "", str(cnpj_raw))
            if not clean_digits:
                self._send_json(400, {"error": "CNPJ é obrigatório."})
                return
            res = check_supplier_sanctions(clean_digits)
            self._send_json(200, res)
            return

        # 3. Download de Arquivos Prontos da Capitania
        if path.startswith("/api/v1/files/"):
            fname = path.split("/api/v1/files/")[-1]
            target = PREBUILT_FILES.get(fname)
            if target and os.path.exists(target):
                with open(target, "rb") as f:
                    data = f.read()
                ctype = "application/octet-stream"
                self._send_binary(data, ctype, fname)
                return
            self._send_json(404, {"error": "Arquivo não encontrado."})
            return

        # 4. Pesquisa de Preços Ao Vivo (JSON)
        if path.startswith("/api/v1/prices/in65/"):
            code_str = path.rstrip("/").split("/")[-1]
            if not code_str.isdigit():
                self._send_json(400, {"error": "Informe um código CATMAT/CATSER numérico."})
                return
            catalog_code = int(code_str)
            supply_unit = qs.get("supply_unit", [""])[0]
            item_type = qs.get("item_type", ["MATERIAL"])[0]
            months_back = int(qs.get("months_back", ["12"])[0])
            uf = qs.get("uf", [None])[0] or None

            try:
                res = price_controller.execute_in65_research(
                    catalog_code=catalog_code,
                    target_supply_unit=supply_unit,
                    item_type=item_type,
                    months_back=months_back,
                    uf=uf,
                )
                self._send_json(200, res)
            except Exception as exc:
                self._send_json(500, {"error": str(exc)})
            return

        # 5. Exportar Planilha Excel Oficial IN 65/2021 Ao Vivo
        if path.startswith("/api/v1/prices/export-xlsx/"):
            code_str = path.rstrip("/").split("/")[-1]
            catalog_code = int(code_str) if code_str.isdigit() else 475109
            supply_unit = qs.get("supply_unit", [""])[0]
            item_type = qs.get("item_type", ["MATERIAL"])[0]
            quantity = int(float(qs.get("qty", qs.get("quantity", ["1"]))[0]))
            months_back = int(qs.get("months_back", ["12"])[0])
            uf = qs.get("uf", [None])[0] or None

            try:
                res = price_controller.execute_in65_research(
                    catalog_code=catalog_code,
                    target_supply_unit=supply_unit,
                    item_type=item_type,
                    months_back=months_back,
                    uf=uf,
                )
                if res.get("status") != "SUCCESS":
                    self._send_json(404, res)
                    return
                xlsx_bytes = build_in65_excel_bytes(res["report"], quantity=quantity)
                prefix = "CATSER" if item_type.upper() == "SERVICO" else "CATMAT"
                self._send_binary(
                    xlsx_bytes,
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    f"Mapa_Comparativo_IN65_{prefix}_{catalog_code}.xlsx",
                )
            except Exception as exc:
                self._send_json(500, {"error": str(exc)})
            return

        # 6. Exportar Relatório Formal de Pesquisa de Preços em Microsoft Word (.DOCX)
        if path.startswith("/api/v1/prices/export-docx/"):
            code_str = path.rstrip("/").split("/")[-1]
            catalog_code = int(code_str) if code_str.isdigit() else 475109
            supply_unit = qs.get("supply_unit", [""])[0]
            item_type = qs.get("item_type", ["MATERIAL"])[0]
            quantity = int(float(qs.get("qty", qs.get("quantity", ["1"]))[0]))
            months_back = int(qs.get("months_back", ["12"])[0])
            uf = qs.get("uf", [None])[0] or None
            elaborador = qs.get("elaborador", ["ENCARREGADO DO SETOR DE OBTENÇÃO"])[0]
            ordenador = qs.get("ordenador", ["ORDENADOR DE DESPESAS"])[0]
            om_name = qs.get("om_name", ["MARINHA DO BRASIL"])[0]

            try:
                res = price_controller.execute_in65_research(
                    catalog_code=catalog_code,
                    target_supply_unit=supply_unit,
                    item_type=item_type,
                    months_back=months_back,
                    uf=uf,
                )
                if res.get("status") != "SUCCESS":
                    self._send_json(404, res)
                    return
                docx_bytes = build_in65_docx_bytes(
                    res["report"],
                    quantity=quantity,
                    om_name=om_name,
                    elaborador_nome=elaborador,
                    ordenador_nome=ordenador,
                )
                prefix = "CATSER" if item_type.upper() == "SERVICO" else "CATMAT"
                self._send_binary(
                    docx_bytes,
                    "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    f"Relatorio_Pesquisa_Precos_{prefix}_{catalog_code}.docx",
                )
            except Exception as exc:
                self._send_json(500, {"error": str(exc)})
            return

        # 7. Exportar Relatório Formal de Pesquisa de Preços em PDF A4 Oficial (.PDF)
        if path.startswith("/api/v1/prices/export-pdf/"):
            code_str = path.rstrip("/").split("/")[-1]
            catalog_code = int(code_str) if code_str.isdigit() else 475109
            supply_unit = qs.get("supply_unit", [""])[0]
            item_type = qs.get("item_type", ["MATERIAL"])[0]
            quantity = int(float(qs.get("qty", qs.get("quantity", ["1"]))[0]))
            months_back = int(qs.get("months_back", ["12"])[0])
            uf = qs.get("uf", [None])[0] or None
            elaborador = qs.get("elaborador", ["ENCARREGADO DO SETOR DE OBTENÇÃO"])[0]
            ordenador = qs.get("ordenador", ["ORDENADOR DE DESPESAS"])[0]
            om_name = qs.get("om_name", ["MARINHA DO BRASIL"])[0]

            try:
                res = price_controller.execute_in65_research(
                    catalog_code=catalog_code,
                    target_supply_unit=supply_unit,
                    item_type=item_type,
                    months_back=months_back,
                    uf=uf,
                )
                if res.get("status") != "SUCCESS":
                    self._send_json(404, res)
                    return
                pdf_bytes = build_in65_pdf_bytes(
                    res["report"],
                    quantity=quantity,
                    om_name=om_name,
                    elaborador_nome=elaborador,
                    ordenador_nome=ordenador,
                )
                prefix = "CATSER" if item_type.upper() == "SERVICO" else "CATMAT"
                self._send_binary(
                    pdf_bytes,
                    "application/pdf",
                    f"Relatorio_Pesquisa_Precos_{prefix}_{catalog_code}.pdf",
                )
            except Exception as exc:
                self._send_json(500, {"error": str(exc)})
            return

        # 8. Exportar Despacho de Justificativa Circunstanciada em Word (.DOCX) (Art. 6º, § 5º da IN 65/2021)
        if path.startswith("/api/v1/prices/export-justificativa-docx/"):
            code_str = path.rstrip("/").split("/")[-1]
            catalog_code = int(code_str) if code_str.isdigit() else 475109
            supply_unit = qs.get("supply_unit", [""])[0]
            item_type = qs.get("item_type", ["MATERIAL"])[0]
            quantity = int(float(qs.get("qty", qs.get("quantity", ["1"]))[0]))
            months_back = int(qs.get("months_back", ["12"])[0])
            uf = qs.get("uf", [None])[0] or None
            ordenador = qs.get("ordenador", ["ORDENADOR DE DESPESAS"])[0]
            om_name = qs.get("om_name", ["MARINHA DO BRASIL"])[0]

            try:
                res = price_controller.execute_in65_research(
                    catalog_code=catalog_code,
                    target_supply_unit=supply_unit,
                    item_type=item_type,
                    months_back=months_back,
                    uf=uf,
                )
                if res.get("status") != "SUCCESS":
                    self._send_json(404, res)
                    return
                docx_bytes = build_in65_justificativa_docx_bytes(
                    res["report"],
                    quantity=quantity,
                    om_name=om_name,
                    ordenador_nome=ordenador,
                )
                prefix = "CATSER" if item_type.upper() == "SERVICO" else "CATMAT"
                self._send_binary(
                    docx_bytes,
                    "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    f"Justificativa_Circunstanciada_Art6_IN65_{prefix}_{catalog_code}.docx",
                )
            except Exception as exc:
                self._send_json(500, {"error": str(exc)})
            return

        # 9. Exportar Despacho de Justificativa Circunstanciada em PDF A4 Oficial (.PDF)
        if path.startswith("/api/v1/prices/export-justificativa-pdf/"):
            code_str = path.rstrip("/").split("/")[-1]
            catalog_code = int(code_str) if code_str.isdigit() else 475109
            supply_unit = qs.get("supply_unit", [""])[0]
            item_type = qs.get("item_type", ["MATERIAL"])[0]
            quantity = int(float(qs.get("qty", qs.get("quantity", ["1"]))[0]))
            months_back = int(qs.get("months_back", ["12"])[0])
            uf = qs.get("uf", [None])[0] or None
            ordenador = qs.get("ordenador", ["ORDENADOR DE DESPESAS"])[0]
            om_name = qs.get("om_name", ["MARINHA DO BRASIL"])[0]

            try:
                res = price_controller.execute_in65_research(
                    catalog_code=catalog_code,
                    target_supply_unit=supply_unit,
                    item_type=item_type,
                    months_back=months_back,
                    uf=uf,
                )
                if res.get("status") != "SUCCESS":
                    self._send_json(404, res)
                    return
                pdf_bytes = build_in65_justificativa_pdf_bytes(
                    res["report"],
                    quantity=quantity,
                    om_name=om_name,
                    ordenador_nome=ordenador,
                )
                prefix = "CATSER" if item_type.upper() == "SERVICO" else "CATMAT"
                self._send_binary(
                    pdf_bytes,
                    "application/pdf",
                    f"Justificativa_Circunstanciada_Art6_IN65_{prefix}_{catalog_code}.pdf",
                )
            except Exception as exc:
                self._send_json(500, {"error": str(exc)})
            return

        # 9.1 Exportar Raio-X de Certidões das Empresas em Word (.DOCX)
        if path.startswith("/api/v1/prices/export-certidoes-docx/"):
            code_str = path.rstrip("/").split("/")[-1]
            catalog_code = int(code_str) if code_str.isdigit() else 475109
            supply_unit = qs.get("supply_unit", [""])[0]
            item_type = qs.get("item_type", ["MATERIAL"])[0]
            months_back = int(qs.get("months_back", ["12"])[0])
            uf = qs.get("uf", [None])[0] or None
            elaborador = qs.get("elaborador", ["ENCARREGADO DO SETOR DE OBTENÇÃO"])[0]
            ordenador = qs.get("ordenador", ["ORDENADOR DE DESPESAS"])[0]
            om_name = qs.get("om_name", ["MARINHA DO BRASIL"])[0]

            try:
                res = price_controller.execute_in65_research(
                    catalog_code=catalog_code,
                    target_supply_unit=supply_unit,
                    item_type=item_type,
                    months_back=months_back,
                    uf=uf,
                )
                if res.get("status") != "SUCCESS":
                    self._send_json(404, res)
                    return
                docx_bytes = build_supplier_compliance_dossier_docx_bytes(
                    res["report"],
                    om_name=om_name,
                    elaborador_nome=elaborador,
                    ordenador_nome=ordenador,
                )
                prefix = "CATSER" if item_type.upper() == "SERVICO" else "CATMAT"
                self._send_binary(
                    docx_bytes,
                    "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    f"RaioX_Certidoes_Empresas_{prefix}_{catalog_code}.docx",
                )
            except Exception as exc:
                self._send_json(500, {"error": str(exc)})
            return

        # 9.2 Exportar Raio-X de Certidões das Empresas em PDF A4 Oficial (.PDF)
        if path.startswith("/api/v1/prices/export-certidoes-pdf/"):
            code_str = path.rstrip("/").split("/")[-1]
            catalog_code = int(code_str) if code_str.isdigit() else 475109
            supply_unit = qs.get("supply_unit", [""])[0]
            item_type = qs.get("item_type", ["MATERIAL"])[0]
            months_back = int(qs.get("months_back", ["12"])[0])
            uf = qs.get("uf", [None])[0] or None
            elaborador = qs.get("elaborador", ["ENCARREGADO DO SETOR DE OBTENÇÃO"])[0]
            ordenador = qs.get("ordenador", ["ORDENADOR DE DESPESAS"])[0]
            om_name = qs.get("om_name", ["MARINHA DO BRASIL"])[0]

            try:
                res = price_controller.execute_in65_research(
                    catalog_code=catalog_code,
                    target_supply_unit=supply_unit,
                    item_type=item_type,
                    months_back=months_back,
                    uf=uf,
                )
                if res.get("status") != "SUCCESS":
                    self._send_json(404, res)
                    return
                pdf_bytes = build_supplier_compliance_dossier_pdf_bytes(
                    res["report"],
                    om_name=om_name,
                    elaborador_nome=elaborador,
                    ordenador_nome=ordenador,
                )
                prefix = "CATSER" if item_type.upper() == "SERVICO" else "CATMAT"
                self._send_binary(
                    pdf_bytes,
                    "application/pdf",
                    f"RaioX_Certidoes_Empresas_{prefix}_{catalog_code}.pdf",
                )
            except Exception as exc:
                self._send_json(500, {"error": str(exc)})
            return

        # 10. Resolução de Contratação no PNCP + Lista Oficial de Arquivos (Edital, TR, Termos)
        if path == "/api/v1/pncp/compra-info":
            chave = qs.get("chave", qs.get("idCompra", qs.get("codigo", [""])))[0]
            if not chave:
                self._send_json(400, {"error": "Informe o parâmetro chave ou idCompra."})
                return
            try:
                compra = resolve_pncp_compra(chave)
                arquivos = []
                pncp_webpage_url = resolve_pncp_webpage_url(chave)
                if compra:
                    cnpj = compra.get("orgaoEntidadeCnpj")
                    ano = compra.get("anoCompraPncp")
                    seq = compra.get("sequencialCompraPncp")
                    if cnpj and ano and seq:
                        arquivos = fetch_pncp_arquivos(cnpj, int(ano), int(seq))
                self._send_json(200, {
                    "status": "SUCCESS",
                    "chave": chave,
                    "compra": compra,
                    "arquivos": arquivos,
                    "pncp_webpage_url": pncp_webpage_url
                })
            except Exception as exc:
                self._send_json(500, {"error": str(exc)})
            return

        # 11. Captura Literal da Página do Edital no Portal PNCP em PDF A4 (Chrome Headless)
        if path == "/api/v1/pncp/page-pdf" or path == "/api/v1/pncp/export-report-pdf":
            chave = qs.get("chave", qs.get("idCompra", qs.get("codigo", [""])))[0]
            req_url = qs.get("url", [""])[0]
            pncp_url = req_url or resolve_pncp_webpage_url(chave)
            if pncp_url or chave:
                page_pdf = capture_pncp_page_pdf(pncp_url, chave_or_id=chave)
                if page_pdf:
                    clean_id = re.sub(r"\W", "_", pncp_url.split("editais/")[-1] if (pncp_url and "editais/" in pncp_url) else (chave or "edital"))
                    self._send_binary(page_pdf, "application/pdf", f"PNCP_Pagina_Oficial_{clean_id}.pdf")
                    return
            # Fallback se a captura headless não puder ser gerada imediatamente
            quote_info = {
                "supplier_name": qs.get("supplier_name", [""])[0],
                "supplier_cnpj": qs.get("supplier_cnpj", [""])[0],
                "unit_price": float(qs.get("unit_price", ["0"])[0] or 0),
                "supply_unit": qs.get("supply_unit", ["UN"])[0],
                "buyer_uasg": qs.get("buyer_uasg", [""])[0],
                "buyer_name": qs.get("buyer_name", [""])[0],
                "purchase_number_year": qs.get("purchase_number_year", [""])[0],
                "item_description": qs.get("item_description", [""])[0],
                "homologation_date": qs.get("homologation_date", [""])[0],
                "purge_status": qs.get("purge_status", ["VALIDA"])[0],
            }
            try:
                compra = resolve_pncp_compra(chave) if chave else None
                arquivos = []
                if compra:
                    cnpj = compra.get("orgaoEntidadeCnpj")
                    ano = compra.get("anoCompraPncp")
                    seq = compra.get("sequencialCompraPncp")
                    if cnpj and ano and seq:
                        arquivos = fetch_pncp_arquivos(cnpj, int(ano), int(seq))
                pdf_bytes = generate_pncp_dossier_pdf(compra, arquivos, quote_info)
                clean_name = re.sub(r"\D", "", chave) if chave else "Avulsa"
                self._send_binary(
                    pdf_bytes,
                    "application/pdf",
                    f"Relatorio_Contratacao_PNCP_{clean_name}.pdf",
                )
            except Exception as exc:
                self._send_json(500, {"error": str(exc)})
            return

        # 12. Download Direto de Arquivo Binário do PNCP (Termos, Edital, Pareceres)
        if path == "/api/v1/pncp/download-file":
            cnpj = qs.get("cnpj", [""])[0]
            ano = qs.get("ano", [""])[0]
            seq = qs.get("seq", [""])[0]
            doc_seq = qs.get("doc_seq", [""])[0]
            fname = qs.get("filename", ["documento_pncp.pdf"])[0]
            if not (cnpj and ano and seq and doc_seq):
                self._send_json(400, {"error": "Informe cnpj, ano, seq e doc_seq."})
                return
            try:
                content, ctype = download_pncp_file_bytes(cnpj, int(ano), int(seq), int(doc_seq))
                if content:
                    self._send_binary(content, ctype or "application/pdf", fname)
                else:
                    self._send_json(404, {"error": "Arquivo não encontrado no PNCP."})
            except Exception as exc:
                self._send_json(500, {"error": str(exc)})
            return

        # 13. Motor de Busca de IRP (Intenções de Registro de Preços) no PNCP (Lei 14.133/2021)
        if path == "/api/v1/irp/search":
            q = qs.get("q", [""])[0]
            status = qs.get("status", ["aberta"])[0]
            uf = qs.get("uf", [""])[0]
            esfera = qs.get("esfera", [""])[0]
            pagina = int(qs.get("pagina", ["1"])[0] or 1)
            tam_pagina = int(qs.get("tam_pagina", qs.get("tamanhoPagina", ["20"]))[0] or 20)
            try:
                res = search_irps(q=q, status=status, uf=uf, esfera=esfera, pagina=pagina, tam_pagina=tam_pagina)
                self._send_json(200, res)
            except Exception as exc:
                self._send_json(500, {"status": "ERROR", "error": str(exc), "items": []})
            return

        # 14. Detalhes Completos, Itens e Arquivos da IRP / Edital SRP / Ata SRP (PNCP)
        if path == "/api/v1/irp/detail":
            cnpj = qs.get("cnpj", [""])[0]
            ano = qs.get("ano", [""])[0]
            seq = qs.get("seq", qs.get("sequencial", [""]))[0]
            doc_type = qs.get("doc_type", qs.get("type", ["auto"]))[0]
            if not (cnpj and ano and seq):
                self._send_json(400, {"status": "ERROR", "error": "Parâmetros 'cnpj', 'ano' e 'seq' são obrigatórios."})
                return
            try:
                res = get_irp_detail(cnpj=cnpj, ano=ano, sequencial=seq, doc_type=doc_type)
                self._send_json(200, res)
            except Exception as exc:
                self._send_json(500, {"status": "ERROR", "error": str(exc), "itens": [], "arquivos": []})
            return

        # 15. Trilha de Auditoria Imutável (Consulta e Verificação de Integridade Criptográfica SHA-256)
        if path == "/api/v1/audit/trail":
            integrity = audit_trail_manager.verify_integrity()
            entries = audit_trail_manager.get_recent_entries(limit=100)
            self._send_json(200, {
                "status": "SUCCESS",
                "integrity": integrity,
                "total_recorded": len(entries),
                "entries": entries,
            })
            return

        self._send_json(404, {"error": "Rota não encontrada."})

    def do_POST(self):
        parsed = urlparse(self.path)
        length = int(self.headers.get("Content-Length", "0"))
        raw_body = self.rfile.read(length).decode("utf-8") if length > 0 else "{}"

        # 0.001 Assistente de IA e Busca Semântica em Linguagem Natural (NLP)
        if parsed.path == "/api/v1/catalog/ai-suggest":
            try:
                data = json.loads(raw_body)
                prompt = data.get("prompt", "")
                result = catalog_engine.semantic_ai_suggest(prompt)
                self._send_json(200, result)
            except Exception as exc:
                self._send_json(400, {"status": "ERROR", "error": str(exc), "top_suggestions": []})
            return

        # 0.001B Gerador de Ofício de Manifestação de Interesse (IRP) / Pedido de Adesão (Carona) - DOCX ou PDF
        if parsed.path in ("/api/v1/irp/oficio/docx", "/api/v1/irp/oficio/pdf"):
            try:
                data = json.loads(raw_body)
                try:
                    from core.irp_oficio_generator import build_irp_oficio_docx_bytes, build_irp_oficio_pdf_bytes
                except ImportError:
                    from irp_oficio_generator import build_irp_oficio_docx_bytes, build_irp_oficio_pdf_bytes

                seq_tag = str(data.get("sequencial") or "IRP").replace("/", "_")
                ano_tag = str(data.get("ano") or "2026")
                if parsed.path.endswith("/docx"):
                    content = build_irp_oficio_docx_bytes(data)
                    self._send_binary(
                        content,
                        "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        f"Oficio_Adesao_IRP_{seq_tag}_{ano_tag}.docx",
                    )
                else:
                    content = build_irp_oficio_pdf_bytes(data)
                    self._send_binary(
                        content,
                        "application/pdf",
                        f"Oficio_Adesao_IRP_{seq_tag}_{ano_tag}.pdf",
                    )
            except Exception as exc:
                self._send_json(500, {"status": "ERROR", "error": f"Falha ao gerar Ofício: {exc}"})
            return

        # 0.001C Termômetro Instantâneo de Vantajosidade da Ata (Preço da Ata vs Mediana Nacional CATMAT/CATSER)
        if parsed.path == "/api/v1/irp/vantajosidade":
            try:
                data = json.loads(raw_body)
                item_desc = str(data.get("descricao") or "").strip()
                item_type = str(data.get("tipo_codigo") or "MATERIAL").upper()
                ata_price = float(data.get("valor_unitario") or 0.0)
                cod_cat = str(data.get("codigo_item_catalogo") or "").strip()

                resolved_code = int(cod_cat) if cod_cat.isdigit() else 0
                if not resolved_code and item_desc:
                    words = [w for w in re.split(r"[,;\s\-/]+", item_desc) if len(w) > 3]
                    q_cat = words[0] if words else item_desc[:15]
                    cat_res = catalog_engine.search(q_cat, item_type=item_type, limit=1)
                    if cat_res:
                        resolved_code = int(cat_res[0].get("code") or 0)

                if not resolved_code:
                    self._send_json(200, {
                        "status": "NO_CODE",
                        "message": "Não foi possível resolver um código CATMAT/CATSER para este item.",
                    })
                    return

                rep = fetch_and_analyze_item(resolved_code, item_type=item_type, description=item_desc)
                mediana_nac = float(rep.get("median_price") or rep.get("recommended_price") or 0.0)
                media_nac = float(rep.get("mean_price") or 0.0)
                valid_count = int(rep.get("valid_quotes_count") or 0)

                economia_pct = None
                veredito = "SEM_REFERENCIA"
                if mediana_nac > 0 and ata_price > 0:
                    economia_pct = round(((mediana_nac - ata_price) / mediana_nac) * 100.0, 2)
                    if economia_pct >= 0:
                        veredito = "VANTAJOSO"
                    elif economia_pct >= -10:
                        veredito = "NA_MEDIA"
                    else:
                        veredito = "ACIMA_MERCADO"

                self._send_json(200, {
                    "status": "SUCCESS",
                    "resolved_code": resolved_code,
                    "item_type": item_type,
                    "ata_price": ata_price,
                    "mediana_nacional": mediana_nac,
                    "media_nacional": media_nac,
                    "valid_quotes_count": valid_count,
                    "economia_pct": economia_pct,
                    "veredito": veredito,
                })
            except Exception as exc:
                self._send_json(500, {"status": "ERROR", "error": str(exc)})
            return

        # 0.002 Registro de Evento na Trilha de Auditoria Imutável (SHA-256)
        if parsed.path == "/api/v1/audit/log":
            try:
                data = json.loads(raw_body)
                user = data.get("user") or "OPERADOR DO SISTEMA"
                action = data.get("action") or "ACAO_DESCONHECIDA"
                target_item = str(data.get("target_item") or "")
                details = data.get("details") or {}
                legal_basis = data.get("legal_basis") or "IN SEGES/ME nº 65/2021"
                ip_addr = self.client_address[0] if self.client_address else "127.0.0.1"
                entry = audit_trail_manager.record_event(
                    user=user,
                    action=action,
                    target_item=target_item,
                    details=details,
                    legal_basis=legal_basis,
                    ip_address=ip_addr,
                )
                self._send_json(200, {"status": "SUCCESS", "entry": entry})
            except Exception as exc:
                self._send_json(400, {"status": "ERROR", "error": str(exc)})
            return

        # 0.005 Exportação Isolada do Infográfico Estatístico e Curva Gaussiana (.PNG HD)
        if parsed.path == "/api/v1/prices/export-infographic-image":
            try:
                data = json.loads(raw_body)
                rep = data.get("report") or {}
                sm = rep.get("sanitized_metrics") or {}
                valid_quotes = rep.get("valid_quotes") or []
                purged_quotes = rep.get("purged_quotes") or []

                img_bytes = generate_statistical_infographic(
                    valid_prices=[float(q.get("unit_price", 0.0)) for q in valid_quotes],
                    purged_prices=[float(q.get("unit_price", 0.0)) for q in purged_quotes],
                    cv_percent=float(sm.get("cv_percent", 0.0)),
                    mean_val=float(sm.get("mean", 0.0)),
                    median_val=float(sm.get("median", 0.0)),
                    std_val=float(sm.get("std_deviation", 0.0)),
                    adopted_param=str(rep.get("adopted_parameter", "MEDIANA")),
                )
                self._send_binary(img_bytes, "image/png", f"Infografico_Estatistico_IN65_{int(time.time())}.png")
            except Exception as exc:
                self._send_json(400, {"status": "ERROR", "error": str(exc)})
            return

        # 0.003 Emissão de Certificado Oficial da Trilha de Auditoria (.PDF)
        if parsed.path == "/api/v1/audit/export-certificate":
            try:
                data = json.loads(raw_body)
                om_name = data.get("om_name") or "MARINHA DO BRASIL"
                user_name = data.get("user_name") or "OPERADOR DO SISTEMA"
                pdf_bytes = build_audit_certificate_pdf_bytes(om_name=om_name, user_name=user_name)
                self._send_binary(pdf_bytes, "application/pdf", f"Certificado_Trilha_Auditoria_{int(time.time())}.pdf")
            except Exception as exc:
                self._send_json(500, {"error": str(exc)})
            return

        # 0.004 Interoperabilidade Bidirecional: Gerar Payload de Integração Oficial PNCP (.JSON)
        if parsed.path == "/api/v1/pncp/export-procurement-payload":
            try:
                data = json.loads(raw_body)
                items_in = data.get("items", [])
                om_name = data.get("om_name") or "MARINHA DO BRASIL"
                uasg_code = data.get("uasg_code") or "785000"
                cnpj_orgao = data.get("cnpj_orgao") or "00394502000144"
                modalidade_id = int(data.get("modalidade_id") or 8)
                numero_compra = data.get("numero_compra") or "00001/2026"
                ano_compra = int(data.get("ano_compra") or 2026)
                processo_nup = data.get("processo_nup") or ""
                objeto_desc = data.get("objeto_descricao") or ""

                integrity = audit_trail_manager.verify_integrity()
                latest_hash = integrity.get("latest_hash") or "N/A"

                payload = build_pncp_procurement_payload(
                    items_list=items_in,
                    om_name=om_name,
                    uasg_code=uasg_code,
                    cnpj_orgao=cnpj_orgao,
                    numero_compra=numero_compra,
                    ano_compra=ano_compra,
                    modalidade_id=modalidade_id,
                    processo_nup=processo_nup,
                    objeto_descricao=objeto_desc,
                    hash_auditoria=latest_hash,
                )
                validation = validate_pncp_payload(payload)

                audit_trail_manager.record_event(
                    user=om_name,
                    action="EXPORTACAO_PAYLOAD_PNCP",
                    target_item=f"LOTE_{len(items_in)}_ITENS",
                    details={"numeroCompra": numero_compra, "valorTotal": payload.get("valorTotalEstimado")},
                    legal_basis="Lei nº 14.133/2021 (Integração PNCP)",
                )

                self._send_json(200, {
                    "status": "SUCCESS",
                    "validation": validation,
                    "pncp_payload": payload,
                })
            except Exception as exc:
                self._send_json(400, {"status": "ERROR", "error": str(exc)})
            return

        # 0. Parser de Lista de Demanda Livre (Extração NLP -> CATMAT/CATSER)
        if parsed.path == "/api/v1/catalog/parse-demand":
            try:
                data = json.loads(raw_body)
                raw_text = data.get("text", "")
                items = catalog_engine.parse_free_demand_text(raw_text)
                self._send_json(200, {"status": "SUCCESS", "parsed_items": items})
            except Exception as exc:
                self._send_json(400, {"error": str(exc)})
            return

        # 0.005 Verificação Rápida em Paralelo de Demanda (Últimos 12 Meses) no Compras.gov.br
        if parsed.path == "/api/v1/catalog/check-demand":
            try:
                data = json.loads(raw_body)
                items_in = data.get("items", [])
                months_back = int(data.get("months_back", 12))
                demand_map = price_controller.check_12m_demand_batch(items_in, months_back=months_back)
                self._send_json(200, {"status": "SUCCESS", "demand": demand_map})
            except Exception as exc:
                self._send_json(500, {"error": str(exc)})
            return

        # 0.008 Recálculo Interativo da Cesta de Preços (Art. 6º, § 2º da IN 65/2021)
        if parsed.path == "/api/v1/prices/recalculate-custom":
            try:
                data = json.loads(raw_body)
                res = price_controller.recalculate_custom(data)
                self._send_json(200, res)
            except Exception as exc:
                self._send_json(400, {"error": str(exc)})
            return

        # 0.01 Pesquisa Automatizada de Preços em Lote (Múltiplos CATMAT/CATSER)
        if parsed.path in ["/api/v1/prices/batch-research", "/api/v1/prices/execute-batch-research", "/api/v1/prices/batch"]:
            try:
                data = json.loads(raw_body)
                items_in = data.get("items", [])
                months_back = int(data.get("months_back", 12))
                uf = data.get("uf") or None

                import concurrent.futures

                def _research_one(idx_it):
                    idx, it = idx_it
                    code_raw = it.get("code")
                    catalog_code = int(code_raw) if str(code_raw).isdigit() else 0
                    unit = str(it.get("unit", "")).strip().upper()
                    item_type = str(it.get("type", "MATERIAL")).strip().upper()
                    try:
                        qty = float(it.get("qty", 1) or 1)
                    except (TypeError, ValueError):
                        qty = 1.0
                    desc_fallback = str(it.get("description") or "").strip()
                    try:
                        res = price_controller.execute_in65_research(
                            catalog_code=catalog_code,
                            target_supply_unit=unit,
                            item_type=item_type,
                            months_back=months_back,
                            uf=uf,
                        ) if catalog_code else {"status": "NO_DATA", "message": "Código inválido."}
                    except Exception as exc_one:
                        res = {"status": "API_ERROR", "message": f"Falha na API do Compras.gov.br: {exc_one}"}
                    rep = res.get("report")
                    has_data = bool(res.get("status") == "SUCCESS" and rep and rep.get("sanitized_metrics"))
                    if has_data:
                        unit_est = float(rep.get("final_estimated_unit_price", 0.0))
                        resolved_unit = rep.get("target_supply_unit") or unit or "UN"
                        resolved_type = rep.get("item_type") or item_type
                    else:
                        unit_est = 0.0
                        resolved_unit = unit or "UN"
                        resolved_type = item_type
                        rep = {
                            "catmat_catser_code": catalog_code,
                            "item_description": desc_fallback or f"Item {catalog_code}",
                            "target_supply_unit": resolved_unit,
                            "valid_quotes": [],
                            "purged_quotes": [],
                            "sanitized_metrics": None,
                            "raw_metrics": None,
                        }
                    return {
                        "item_index": idx,
                        "code": catalog_code,
                        "type": resolved_type,
                        "unit": resolved_unit,
                        "qty": qty,
                        "has_data": has_data,
                        "status": res.get("status", "SUCCESS" if has_data else "NO_DATA"),
                        "message": res.get("message", "Pesquisa homologada com sucesso" if has_data else "Sem compras homologadas nos últimos 12 e 24 meses"),
                        "search_window_months": int(rep.get("search_window_months") or 12) if isinstance(rep, dict) else 12,
                        "is_extended_24m": bool(rep.get("is_extended_24m")) if isinstance(rep, dict) else False,
                        "unit_price": unit_est,
                        "total_price": qty * unit_est,
                        "report": rep,
                    }

                with concurrent.futures.ThreadPoolExecutor(max_workers=6) as pool:
                    results = list(pool.map(_research_one, list(enumerate(items_in, start=1))))
                total_lote = sum(r["total_price"] for r in results)

                self._send_json(200, {
                    "status": "SUCCESS",
                    "total_items": len(results),
                    "total_lote": total_lote,
                    "results": results,
                })
            except Exception as exc:
                self._send_json(500, {"error": str(exc)})
            return

        # 0.012 Captura da página oficial do PNCP em PDF (POST com {chave, url})
        if parsed.path == "/api/v1/pncp/page-pdf-post":
            try:
                data = json.loads(raw_body or "{}")
                chave = str(data.get("chave") or "").strip()
                pncp_url = str(data.get("url") or "").strip() or resolve_pncp_webpage_url(chave)
                if not (pncp_url or chave):
                    self._send_json(400, {"error": "Informe a chave da compra ou a URL do PNCP."})
                    return
                page_pdf = capture_pncp_page_pdf(pncp_url, chave_or_id=chave)
                if page_pdf:
                    clean_id = re.sub(r"\W", "_", pncp_url.split("editais/")[-1] if (pncp_url and "editais/" in pncp_url) else (chave or "edital"))
                    self._send_binary(page_pdf, "application/pdf", f"PNCP_Pagina_Oficial_{clean_id}.pdf")
                    return
                compra = resolve_pncp_compra(chave) if chave else None
                arquivos = []
                if compra:
                    cnpj = compra.get("orgaoEntidadeCnpj")
                    ano = compra.get("anoCompraPncp")
                    seq = compra.get("sequencialCompraPncp")
                    if cnpj and ano and seq:
                        arquivos = fetch_pncp_arquivos(cnpj, int(ano), int(seq))
                pdf_bytes = generate_pncp_dossier_pdf(compra, arquivos, data.get("quote_info") or {})
                clean_name = re.sub(r"\D", "", chave) if chave else "Avulsa"
                self._send_binary(pdf_bytes, "application/pdf", f"Relatorio_Contratacao_PNCP_{clean_name}.pdf")
            except Exception as exc:
                self._send_json(500, {"error": str(exc)})
            return

        # 0.015 Exportar documento de UM item a partir do relatório já pesquisado/curado (sem nova consulta ao Gov)
        if parsed.path == "/api/v1/prices/export-report":
            try:
                data = json.loads(raw_body)
                rep = data.get("report") or {}
                if not rep.get("sanitized_metrics"):
                    self._send_json(400, {"error": "Relatório sem cotações válidas. Pesquise o item no Compras.gov.br primeiro."})
                    return
                fmt = str(data.get("format", "pdf")).lower()
                try:
                    quantity = float(data.get("qty", 1) or 1)
                except (TypeError, ValueError):
                    quantity = 1.0
                if quantity.is_integer():
                    quantity = int(quantity)
                item_type = str(data.get("item_type", "MATERIAL")).upper()
                om_name = str(data.get("om_name") or "MARINHA DO BRASIL")
                elaborador = str(data.get("elaborador") or "ENCARREGADO DO SETOR DE OBTENÇÃO")
                ordenador = str(data.get("ordenador") or "ORDENADOR DE DESPESAS")
                code = rep.get("catmat_catser_code", "item")
                prefix = "CATSER" if item_type == "SERVICO" else "CATMAT"
                if fmt == "xlsx":
                    content = build_in65_excel_bytes(rep, quantity=quantity)
                    ctype, fname = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", f"Mapa_Comparativo_IN65_{prefix}_{code}.xlsx"
                elif fmt == "docx":
                    content = build_in65_docx_bytes(rep, quantity=quantity, om_name=om_name, elaborador_nome=elaborador, ordenador_nome=ordenador)
                    ctype, fname = "application/vnd.openxmlformats-officedocument.wordprocessingml.document", f"Relatorio_Pesquisa_Precos_{prefix}_{code}.docx"
                elif fmt == "just-docx":
                    content = build_in65_justificativa_docx_bytes(rep, quantity=quantity, om_name=om_name, ordenador_nome=ordenador)
                    ctype, fname = "application/vnd.openxmlformats-officedocument.wordprocessingml.document", f"Justificativa_Art6_IN65_{prefix}_{code}.docx"
                elif fmt == "just-pdf":
                    content = build_in65_justificativa_pdf_bytes(rep, quantity=quantity, om_name=om_name, ordenador_nome=ordenador)
                    ctype, fname = "application/pdf", f"Justificativa_Art6_IN65_{prefix}_{code}.pdf"
                elif fmt in ("certidoes-docx", "certidoes_docx"):
                    content = build_supplier_compliance_dossier_docx_bytes(rep, om_name=om_name, elaborador_nome=elaborador, ordenador_nome=ordenador)
                    ctype, fname = "application/vnd.openxmlformats-officedocument.wordprocessingml.document", f"RaioX_Certidoes_Empresas_{prefix}_{code}.docx"
                elif fmt in ("certidoes-pdf", "certidoes_pdf", "certidoes"):
                    content = build_supplier_compliance_dossier_pdf_bytes(rep, om_name=om_name, elaborador_nome=elaborador, ordenador_nome=ordenador)
                    ctype, fname = "application/pdf", f"RaioX_Certidoes_Empresas_{prefix}_{code}.pdf"
                else:
                    content = build_in65_pdf_bytes(rep, quantity=quantity, om_name=om_name, elaborador_nome=elaborador, ordenador_nome=ordenador)
                    ctype, fname = "application/pdf", f"Relatorio_Pesquisa_Precos_{prefix}_{code}.pdf"
                self._send_binary(content, ctype, fname)
            except Exception as exc:
                self._send_json(500, {"error": str(exc)})
            return

        # 0.019 Iniciar Job Assíncrono de Pacote ZIP do Lote (Zero HTTP 504 para 50, 150, 300+ itens)
        if parsed.path == "/api/v1/jobs/batch-zip":
            try:
                data = json.loads(raw_body)
                items_in = data.get("items", [])
                om_name = str(data.get("om_name", "MARINHA DO BRASIL"))
                ordenador = str(data.get("ordenador", "ORDENADOR DE DESPESAS"))
                elaborador = str(data.get("elaborador", "ENCARREGADO DO SETOR DE OBTENÇÃO"))
                mode = str(data.get("mode", "auto")).strip().lower()

                items_results = []
                for idx, it in enumerate(items_in, start=1):
                    rep = it.get("report")
                    qty = float(it.get("qty", 1) or 1)
                    if not rep or not rep.get("sanitized_metrics"):
                        continue
                    items_results.append({
                        "item_index": idx,
                        "code": rep.get("catmat_catser_code", it.get("code")),
                        "type": it.get("type", "MATERIAL"),
                        "unit": it.get("unit", rep.get("target_supply_unit", "UN")),
                        "qty": qty,
                        "report": rep,
                    })

                if not items_results:
                    self._send_json(404, {"error": "Nenhum item do lote possui cotações válidas para exportação."})
                    return

                fname = f"Processo_Completo_IN65_{len(items_results)}_itens.zip"

                def _worker(out_path: str, progress_cb):
                    return build_batch_zip_to_file(
                        items_results=items_results,
                        output_path=out_path,
                        om_name=om_name,
                        ordenador_nome=ordenador,
                        elaborador_nome=elaborador,
                        extract_gov_docs=False,
                        mode=mode,
                        progress_callback=progress_cb,
                    )

                job_status = batch_job_manager.submit_file_job(
                    kind="batch_zip",
                    filename=fname,
                    content_type="application/zip",
                    total_items=len(items_results),
                    worker_fn=_worker,
                )
                self._send_json(202, job_status)
            except Exception as exc:
                self._send_json(500, {"error": str(exc)})
            return

        # 0.02 Exportar Pacote ZIP Completo do Lote (.ZIP com relatórios Word, PDF e Excel de todos os itens)
        if parsed.path == "/api/v1/prices/export-batch-zip":
            try:
                data = json.loads(raw_body)
                items_in = data.get("items", [])
                om_name = str(data.get("om_name", "MARINHA DO BRASIL"))
                ordenador = str(data.get("ordenador", "ORDENADOR DE DESPESAS"))
                elaborador = str(data.get("elaborador", "ENCARREGADO DO SETOR DE OBTENÇÃO"))
                months_back = int(data.get("months_back", 12))
                uf = data.get("uf") or None

                items_results = []
                for idx, it in enumerate(items_in, start=1):
                    rep = it.get("report")
                    qty = float(it.get("qty", 1) or 1)
                    if not rep or not rep.get("sanitized_metrics"):
                        if len(items_in) <= 20:
                            code_raw = it.get("code")
                            catalog_code = int(code_raw) if str(code_raw).isdigit() else 475109
                            unit = str(it.get("unit", "")).strip().upper()
                            item_type = str(it.get("type", "MATERIAL")).strip().upper()
                            res = price_controller.execute_in65_research(
                                catalog_code=catalog_code,
                                target_supply_unit=unit,
                                item_type=item_type,
                                months_back=months_back,
                                uf=uf,
                            )
                            rep = res.get("report") or {}

                    if not rep or not rep.get("sanitized_metrics"):
                        continue

                    items_results.append({
                        "item_index": idx,
                        "code": rep.get("catmat_catser_code", it.get("code")),
                        "type": it.get("type", "MATERIAL"),
                        "unit": it.get("unit", rep.get("target_supply_unit", "UN")),
                        "qty": qty,
                        "report": rep,
                    })

                if not items_results:
                    self._send_json(404, {"error": "Nenhum item do lote possui cotações válidas para exportação."})
                    return

                zip_bytes = build_batch_zip_bytes(
                    items_results,
                    om_name=om_name,
                    ordenador_nome=ordenador,
                    elaborador_nome=elaborador,
                    extract_gov_docs=False,
                )
                self._send_binary(zip_bytes, "application/zip", f"Processo_Completo_IN65_{len(items_results)}_itens.zip")
            except Exception as exc:
                self._send_json(500, {"error": str(exc)})
            return

        # 0.03 Exportar Planilha Consolidada do Lote (.XLSX com todas as abas)
        if parsed.path == "/api/v1/prices/export-batch-xlsx":
            try:
                data = json.loads(raw_body)
                items_in = data.get("items", [])
                om_name = str(data.get("om_name", "MARINHA DO BRASIL"))
                months_back = int(data.get("months_back", 12))
                uf = data.get("uf") or None

                items_results = []
                for idx, it in enumerate(items_in, start=1):
                    rep = it.get("report")
                    qty = float(it.get("qty", 1) or 1)
                    if not rep or not rep.get("sanitized_metrics"):
                        if len(items_in) <= 20:
                            code_raw = it.get("code")
                            catalog_code = int(code_raw) if str(code_raw).isdigit() else 475109
                            unit = str(it.get("unit", "")).strip().upper()
                            item_type = str(it.get("type", "MATERIAL")).strip().upper()
                            res = price_controller.execute_in65_research(
                                catalog_code=catalog_code,
                                target_supply_unit=unit,
                                item_type=item_type,
                                months_back=months_back,
                                uf=uf,
                            )
                            rep = res.get("report") or {}

                    if not rep or not rep.get("sanitized_metrics"):
                        continue

                    items_results.append({
                        "item_index": idx,
                        "code": rep.get("catmat_catser_code", it.get("code")),
                        "type": it.get("type", "MATERIAL"),
                        "unit": it.get("unit", rep.get("target_supply_unit", "UN")),
                        "qty": qty,
                        "report": rep,
                    })

                xlsx_bytes = build_consolidated_in65_excel_bytes(items_results, uasg_name=om_name)
                self._send_binary(
                    xlsx_bytes,
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    f"Mapa_Comparativo_Consolidado_Lote_{len(items_results)}_itens.xlsx",
                )
            except Exception as exc:
                self._send_json(500, {"error": str(exc)})
            return

        # 0.04 Exportar Caderno Consolidado de Preços em PDF (A4 Landscape oficial para juntada SEI)
        if parsed.path == "/api/v1/prices/export-batch-pdf":
            try:
                data = json.loads(raw_body)
                items_in = data.get("items", [])
                om_name = str(data.get("om_name", "MARINHA DO BRASIL"))
                ordenador = str(data.get("ordenador", "ORDENADOR DE DESPESAS"))
                elaborador = str(data.get("elaborador", "ENCARREGADO DO SETOR DE OBTENÇÃO"))

                items_results = []
                for idx, it in enumerate(items_in, start=1):
                    rep = it.get("report")
                    qty = float(it.get("qty", 1) or 1)
                    if rep and rep.get("sanitized_metrics"):
                        items_results.append({
                            "item_index": idx,
                            "code": rep.get("catmat_catser_code", it.get("code")),
                            "type": it.get("type", "MATERIAL"),
                            "unit": it.get("unit", rep.get("target_supply_unit", "UN")),
                            "qty": qty,
                            "report": rep,
                        })

                if not items_results:
                    self._send_json(404, {"error": "Nenhum item do lote possui cotações válidas para exportação."})
                    return

                pdf_bytes = build_consolidated_in65_pdf_bytes(
                    items_results,
                    om_name=om_name,
                    elaborador_nome=elaborador,
                    ordenador_nome=ordenador,
                )
                self._send_binary(
                    pdf_bytes,
                    "application/pdf",
                    f"Caderno_Consolidado_Precos_Lote_{len(items_results)}_itens.pdf",
                )
            except Exception as exc:
                self._send_json(500, {"error": str(exc)})
            return

        # 0.05 Exportar Caderno Consolidado de Preços em Word (.DOCX Landscape oficial)
        if parsed.path == "/api/v1/prices/export-batch-docx":
            try:
                data = json.loads(raw_body)
                items_in = data.get("items", [])
                om_name = str(data.get("om_name", "MARINHA DO BRASIL"))
                ordenador = str(data.get("ordenador", "ORDENADOR DE DESPESAS"))
                elaborador = str(data.get("elaborador", "ENCARREGADO DO SETOR DE OBTENÇÃO"))

                items_results = []
                for idx, it in enumerate(items_in, start=1):
                    rep = it.get("report")
                    qty = float(it.get("qty", 1) or 1)
                    if rep and rep.get("sanitized_metrics"):
                        items_results.append({
                            "item_index": idx,
                            "code": rep.get("catmat_catser_code", it.get("code")),
                            "type": it.get("type", "MATERIAL"),
                            "unit": it.get("unit", rep.get("target_supply_unit", "UN")),
                            "qty": qty,
                            "report": rep,
                        })

                if not items_results:
                    self._send_json(404, {"error": "Nenhum item do lote possui cotações válidas para exportação."})
                    return

                docx_bytes = build_consolidated_in65_docx_bytes(
                    items_results,
                    om_name=om_name,
                    elaborador_nome=elaborador,
                    ordenador_nome=ordenador,
                )
                self._send_binary(
                    docx_bytes,
                    "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    f"Caderno_Consolidado_Precos_Lote_{len(items_results)}_itens.docx",
                )
            except Exception as exc:
                self._send_json(500, {"error": str(exc)})
            return

        # 0.06 Exportar Nota Técnica Oficial AGU (.DOCX ou .PDF) — Modelo SEGES/AGU
        if parsed.path in ("/api/v1/prices/export-agu-nota-tecnica-docx", "/api/v1/prices/export-agu-nota-tecnica-pdf"):
            try:
                data = json.loads(raw_body)
                items_in = data.get("items", [])
                om_name = str(data.get("om_name", "MARINHA DO BRASIL"))
                ordenador = str(data.get("ordenador", "ORDENADOR DE DESPESAS"))
                elaborador = str(data.get("elaborador", "ENCARREGADO DO SETOR DE OBTENÇÃO"))
                matricula = str(data.get("matricula", "A INFORMAR"))
                cidade = str(data.get("cidade", "Rio de Janeiro - RJ"))
                nota_num = str(data.get("nota_num", ""))
                objeto_geral = str(data.get("objeto_geral", ""))

                items_results = []
                for idx, it in enumerate(items_in, start=1):
                    rep = it.get("report")
                    qty = float(it.get("qty", 1) or 1)
                    if rep and rep.get("sanitized_metrics"):
                        items_results.append({
                            "item_index": idx,
                            "code": rep.get("catmat_catser_code", it.get("code")),
                            "type": it.get("type", "MATERIAL"),
                            "unit": it.get("unit", rep.get("target_supply_unit", "UN")),
                            "qty": qty,
                            "report": rep,
                        })

                if not items_results:
                    self._send_json(404, {"error": "Nenhum item do lote possui cotações válidas para gerar a Nota Técnica AGU."})
                    return

                if parsed.path.endswith("-docx"):
                    out_bytes = build_agu_nota_tecnica_docx_bytes(
                        items_results,
                        om_name=om_name,
                        ordenador_nome=ordenador,
                        elaborador_nome=elaborador,
                        elaborador_matricula=matricula,
                        cidade=cidade,
                        nota_num=nota_num,
                        objeto_geral=objeto_geral,
                    )
                    self._send_binary(
                        out_bytes,
                        "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        f"Nota_Tecnica_AGU_IN65_{len(items_results)}_itens.docx",
                    )
                else:
                    out_bytes = build_agu_nota_tecnica_pdf_bytes(
                        items_results,
                        om_name=om_name,
                        ordenador_nome=ordenador,
                        elaborador_nome=elaborador,
                        elaborador_matricula=matricula,
                        cidade=cidade,
                        nota_num=nota_num,
                        objeto_geral=objeto_geral,
                    )
                    self._send_binary(
                        out_bytes,
                        "application/pdf",
                        f"Nota_Tecnica_AGU_IN65_{len(items_results)}_itens.pdf",
                    )
            except Exception as exc:
                self._send_json(500, {"error": str(exc)})
            return

        # 0.07 Exportar Raio-X Consolidado de Certidões das Empresas do Lote (.DOCX ou .PDF)
        if parsed.path in ("/api/v1/prices/export-batch-certidoes-docx", "/api/v1/prices/export-batch-certidoes-pdf"):
            try:
                data = json.loads(raw_body)
                items_in = data.get("items", [])
                om_name = str(data.get("om_name", "MARINHA DO BRASIL"))
                ordenador = str(data.get("ordenador", "ORDENADOR DE DESPESAS"))
                elaborador = str(data.get("elaborador", "ENCARREGADO DO SETOR DE OBTENÇÃO"))

                items_results = []
                for idx, it in enumerate(items_in, start=1):
                    rep = it.get("report")
                    qty = float(it.get("qty", 1) or 1)
                    if rep and rep.get("sanitized_metrics"):
                        items_results.append({
                            "item_index": idx,
                            "code": rep.get("catmat_catser_code", it.get("code")),
                            "type": it.get("type", "MATERIAL"),
                            "unit": it.get("unit", rep.get("target_supply_unit", "UN")),
                            "qty": qty,
                            "report": rep,
                        })

                if not items_results:
                    self._send_json(404, {"error": "Nenhum item do lote possui cotações válidas para gerar o Raio-X de Certidões."})
                    return

                if parsed.path.endswith("-docx"):
                    out_bytes = build_consolidated_suppliers_dossier_docx_bytes(
                        items_results, om_name=om_name, ordenador_nome=ordenador, elaborador_nome=elaborador
                    )
                    self._send_binary(
                        out_bytes,
                        "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        f"RaioX_Certidoes_Empresas_{len(items_results)}_itens.docx",
                    )
                else:
                    out_bytes = build_consolidated_suppliers_dossier_pdf_bytes(
                        items_results, om_name=om_name, ordenador_nome=ordenador, elaborador_nome=elaborador
                    )
                    self._send_binary(
                        out_bytes,
                        "application/pdf",
                        f"RaioX_Certidoes_Empresas_{len(items_results)}_itens.pdf",
                    )
            except Exception as exc:
                self._send_json(500, {"error": str(exc)})
            return

        # 0.1 Gerador Dinâmico de Minutas (DFD ou TR) Customizado para o Item Pesquisado
        if parsed.path == "/api/v1/artifacts/generate-draft":
            try:
                data = json.loads(raw_body)
                artifact_type = str(data.get("artifact_type", "TR")).strip().upper()
                desc = str(data.get("item_description", "Item")).strip()
                code = str(data.get("catalog_code", "000000")).strip()
                unit = str(data.get("unit", "UN")).strip().upper() or "UN"
                qty = float(data.get("qty", 1) or 1)
                unit_price = float(data.get("unit_price", 0.0) or 0.0)
                total_price = qty * unit_price
                org_name = str(data.get("org_name", "MARINHA DO BRASIL")).strip()

                clauses = []
                findings = []
                if artifact_type == "DFD":
                    title = f"DOCUMENTO DE FORMALIZAÇÃO DA DEMANDA (DFD) — CÓD. {code}"
                    clauses = [
                        {
                            "num": "1.1",
                            "title": "JUSTIFICATIVA DA NECESSIDADE DA CONTRATAÇÃO",
                            "text": f"1.1. A presente demanda justifica-se pela necessidade operacional de aquisição/contratação de {desc} (Código {code}), destinada a manter a prontidão e a continuidade das atividades da {org_name}.",
                            "modified": False
                        },
                        {
                            "num": "2.1",
                            "title": "QUANTITATIVO E ESTIMATIVA DE VALOR (IN 65/2021)",
                            "text": f"2.1. Estima-se o quantitativo de {int(qty) if qty.is_integer() else qty} {unit} de {desc}, com valor unitário referencial saneado de R$ {unit_price:,.2f} (apurado em conformidade com o Art. 6º da IN SEGES/ME nº 65/2021), totalizando R$ {total_price:,.2f}.",
                            "modified": True
                        },
                        {
                            "num": "3.1",
                            "title": "PREVISÃO NO PLANO DE CONTRATAÇÕES ANUAL (PCA)",
                            "text": "3.1. A aquisição está prevista no Plano de Contratações Anual (PCA) do exercício corrente, observando o alinhamento com o planejamento orçamentário da Unidade Gestora.",
                            "modified": False
                        },
                        {
                            "num": "4.1",
                            "title": "ENQUADRAMENTO LEGAL DA CONTRATAÇÃO DIRETA",
                            "text": "4.1. O enquadramento inicial pretendido é de Dispensa de Licitação em razão do valor (Art. 75, inciso II, da Lei nº 14.133/2021), respeitado o limite de não-fracionamento anual para a mesma linha de fornecimento.",
                            "modified": True
                        }
                    ]
                else:
                    title = f"TERMO DE REFERÊNCIA (TR) — CÓD. {code} (LEI 14.133/2021)"
                    clauses = [
                        {
                            "num": "1.1",
                            "title": "DEFINIÇÃO DO OBJETO",
                            "text": f"1.1. O objeto da presente contratação é a aquisição/prestação de {desc}, código de catálogo {code}, com entrega imediata em remessa única, atendendo às especificações técnicas e de qualidade estabelecidas.",
                            "modified": False
                        },
                        {
                            "num": "1.5",
                            "title": "VIGÊNCIA DA CONTRATAÇÃO (PADRÃO AGU)",
                            "text": "1.5. O prazo de vigência da contratação é de 60 (sessenta) dias contados da emissão e recebimento da Nota de Empenho, na forma do art. 105 da Lei nº 14.133/2021, compatível com a entrega imediata em até 10 dias úteis e liquidação da despesa.",
                            "modified": True
                        },
                        {
                            "num": "2.1",
                            "title": "ESTIMATIVA DE CUSTO E ADEQUAÇÃO ORÇAMENTÁRIA (IN 65/2021)",
                            "text": f"2.1. O custo estimado total é de R$ {total_price:,.2f} ({int(qty) if qty.is_integer() else qty} {unit} ao custo unitário referencial saneado de R$ {unit_price:,.2f}), obtido a partir de pesquisa de preços realizada via API de Dados Abertos do Compras.gov.br e PNCP.",
                            "modified": True
                        },
                        {
                            "num": "2.3",
                            "title": "DECLARAÇÃO DE NÃO-FRACIONAMENTO DE DESPESA (ART. 75, § 1º)",
                            "text": "2.3. Declara-se, para os fins do art. 75, § 1º da Lei nº 14.133/2021, que o valor desta contratação foi somado às demais despesas da mesma natureza no exercício financeiro, não ultrapassando o limite legal de R$ 65.492,11 (Decreto nº 11.871/2023).",
                            "modified": True
                        },
                        {
                            "num": "3.2.12",
                            "title": "PARTICIPAÇÃO DE COOPERATIVAS (SÚMULA TCU)",
                            "text": "3.2.12. Será permitida a participação de sociedades cooperativas, desde que apresentem demonstrativo de atuação em regime cooperado e atendam integralmente ao art. 16 da Lei nº 14.133/2021 e à jurisprudência do TCU.",
                            "modified": True
                        },
                        {
                            "num": "4.1",
                            "title": "REQUISITOS DE HABILITAÇÃO FISCAL, SOCIAL E TRABALHISTA",
                            "text": "4.1. A regularidade do fornecedor será comprovada por meio da Certidão Conjunta RFB/PGFN, CNDT (Trabalhista), CRF do FGTS e consulta à regularidade cadastral no SICAF e cadastros de impedimento (CEIS/CNEP/TCU).",
                            "modified": True
                        }
                    ]

                self._send_json(200, {
                    "status": "SUCCESS",
                    "title": title,
                    "artifact_type": artifact_type,
                    "clauses": clauses,
                    "findings": []
                })
            except Exception as exc:
                self._send_json(400, {"error": str(exc)})
            return

        # 1. Auditar Arquivo Enviado (PDF ou DOCX em Base64) ou Texto Colado — COM PERMISSÃO DO USUÁRIO
        if parsed.path == "/api/v1/rag/audit-upload":
            try:
                data = json.loads(raw_body)
                artifact_type = data.get("artifact_type", "TR")
                filename = data.get("filename", "")
                file_b64 = data.get("file_base64", "")
                raw_text = data.get("text", "")

                doc_structure = None
                file_bytes = b""
                if file_b64:
                    file_bytes = base64.b64decode(file_b64)
                    if filename.lower().endswith(".pdf"):
                        doc_structure = DocumentLayoutParser.parse_pdf(file_bytes, filename=filename)

                        # Fallback de Visão Computacional se PDF for escaneado (< 100 caracteres legíveis)
                        total_chars = sum(len(e.get("text", "")) for e in doc_structure.get("elements", []))
                        if total_chars < 100:
                            try:
                                from google.cloud import vision
                                client = vision.ImageAnnotatorClient()
                                input_config = vision.InputConfig(content=file_bytes, mime_type="application/pdf")
                                features = [vision.Feature(type_=vision.Feature.Type.DOCUMENT_TEXT_DETECTION)]
                                requests_v = [vision.AnnotateFileRequest(features=features, input_config=input_config)]
                                response = client.batch_annotate_files(requests=requests_v)
                                ocr_lines = []
                                for file_resp in response.responses:
                                    for page_resp in file_resp.responses:
                                        if page_resp.full_text_annotation and page_resp.full_text_annotation.text:
                                            ocr_lines.extend([l.strip() for l in page_resp.full_text_annotation.text.split("\n") if l.strip()])
                                if ocr_lines:
                                    doc_structure = DocumentLayoutParser.parse_text("\n".join(ocr_lines), filename=filename)
                            except Exception as ocr_err:
                                print(f"[Cloud Vision OCR Notice] {ocr_err}")

                    elif filename.lower().endswith(".docx"):
                        doc_structure = DocumentLayoutParser.parse_docx(file_bytes, filename=filename)
                    else:
                        doc_structure = DocumentLayoutParser.parse_text(file_bytes.decode("utf-8", errors="ignore"), filename=filename)
                else:
                    file_bytes = (raw_text or "").encode("utf-8")
                    doc_structure = DocumentLayoutParser.parse_text(raw_text, filename=filename or "minuta_texto.txt")

                if not doc_structure:
                    doc_structure = DocumentLayoutParser.parse_text(raw_text or "", filename=filename or "minuta.txt")

                # Auditoria Jurídica SEM mutação automática: aguarda permissão expressa do usuário!
                audited_doc = DocumentLegalAuditor.audit_and_diff(doc_structure, require_permission=True)

                # Auditoria Semântica Profunda (Multi-Artefato + Prompt de Engenharia Jurídica via Gemini API)
                context_note = str(data.get("context_note") or "").strip()
                gemini_api_key = str(data.get("gemini_api_key") or "").strip()
                gemini_model = str(data.get("gemini_model") or "gemini-2.5-flash").strip()
                extra_files = data.get("extra_files") or []

                combined_chunks = [f"=== DOCUMENTO PRINCIPAL: {filename or 'Minuta'} ==="]
                for idx_el, el in enumerate(audited_doc.get("elements", []), start=1):
                    t_el = (el.get("text") or "").strip()
                    if t_el:
                        combined_chunks.append(f"[El {idx_el}] {t_el}")

                for ef in extra_files:
                    if not isinstance(ef, dict):
                        continue
                    ef_name = str(ef.get("filename") or "Anexo").strip()
                    ef_b64 = ef.get("file_base64") or ""
                    if not ef_b64:
                        continue
                    try:
                        ef_bytes = base64.b64decode(ef_b64)
                        if ef_name.lower().endswith(".pdf"):
                            ef_struct = DocumentLayoutParser.parse_pdf(ef_bytes, filename=ef_name)
                        elif ef_name.lower().endswith(".docx"):
                            ef_struct = DocumentLayoutParser.parse_docx(ef_bytes, filename=ef_name)
                        else:
                            ef_struct = DocumentLayoutParser.parse_text(ef_bytes.decode("utf-8", errors="ignore"), filename=ef_name)
                        combined_chunks.append(f"\n=== ANEXO DOS AUTOS: {ef_name} ===")
                        for ef_el in ef_struct.get("elements", []):
                            ef_t = (ef_el.get("text") or "").strip()
                            if ef_t:
                                combined_chunks.append(ef_t)
                    except Exception:
                        pass

                deep_audit_result = {"status": "SKIPPED", "achados": [], "roteiro_saneamento": []}
                try:
                    try:
                        from core.gemini_legal_auditor import run_gemini_deep_audit
                    except ImportError:
                        from gemini_legal_auditor import run_gemini_deep_audit
                    deep_audit_result = run_gemini_deep_audit(
                        combined_text="\n".join(combined_chunks),
                        context_note=context_note,
                        api_key_override=gemini_api_key,
                        model_name=gemini_model,
                    )
                except Exception as deep_exc:
                    deep_audit_result = {"status": "ERROR", "reason": str(deep_exc), "achados": [], "roteiro_saneamento": []}

                # Salva arquivo original em sessão para cirurgia fiel no .DOCX original durante a exportação
                doc_session_id = document_session_store.save(
                    original_bytes=file_bytes,
                    filename=filename or "minuta",
                    original_format=audited_doc.get("original_format", "docx"),
                    doc_structure=audited_doc,
                )

                # Estruturação retrocompatível de cláusulas para o frontend
                clauses_list = []
                for el in audited_doc.get("elements", []):
                    if el.get("type") in ["CLAUSE", "PARAGRAPH"]:
                        clauses_list.append({
                            "num": el.get("clause_number") or el.get("number") or "",
                            "title": el.get("section_title") or el.get("title") or "Cláusula Contratual",
                            "text": el.get("text") or "",
                            "modified": bool(el.get("modified")),
                            "diff_segments": el.get("diff_segments", []),
                        })

                response_payload = {
                    "status": "SUCCESS",
                    "doc_session_id": doc_session_id,
                    "doc_structure": audited_doc,
                    "title": audited_doc.get("title", "DOCUMENTO AUDITADO (LEI 14.133/2021)"),
                    "original_format": audited_doc.get("original_format", "pdf"),
                    "elements": audited_doc.get("elements", []),
                    "clauses": clauses_list,
                    "findings": audited_doc.get("findings") or audited_doc.get("audit_summary", {}).get("findings", []),
                    "compliance_score": audited_doc.get("audit_summary", {}).get("compliance_score", 100.0),
                    "audit_summary": audited_doc.get("audit_summary", {}),
                    "signatures": audited_doc.get("signatures", []),
                    "deep_audit": deep_audit_result,
                }
                self._send_json(200, response_payload)
            except Exception as exc:
                self._send_json(400, {"error": str(exc)})
            return

        # 1B. Exportar Despacho de Saneamento Preventivo (Word .DOCX ou PDF) gerado pela Auditoria Profunda
        if parsed.path in ("/api/v1/rag/despacho-saneamento/docx", "/api/v1/rag/despacho-saneamento/pdf"):
            try:
                data = json.loads(raw_body)
                try:
                    from core.gemini_legal_auditor import (
                        build_despacho_saneamento_docx_bytes,
                        build_despacho_saneamento_pdf_bytes,
                    )
                except ImportError:
                    from gemini_legal_auditor import (
                        build_despacho_saneamento_docx_bytes,
                        build_despacho_saneamento_pdf_bytes,
                    )
                if parsed.path.endswith("/docx"):
                    content = build_despacho_saneamento_docx_bytes(data)
                    self._send_binary(
                        content,
                        "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        "Despacho_Saneamento_Auditoria_NLLC.docx",
                    )
                else:
                    content = build_despacho_saneamento_pdf_bytes(data)
                    self._send_binary(
                        content,
                        "application/pdf",
                        "Despacho_Saneamento_Auditoria_NLLC.pdf",
                    )
            except Exception as exc:
                self._send_json(500, {"error": f"Falha ao gerar Despacho de Saneamento: {exc}"})
            return

        # 2. Exportar Documento Word (.DOCX) — Preservação 100% da Diagramação Original + Amarelo/Vermelho Tachado
        if parsed.path == "/api/v1/rag/export-docx":
            try:
                data = json.loads(raw_body)
                doc_session_id = str(data.get("doc_session_id") or "").strip()
                approved_decisions = data.get("approved_decisions")
                session_data = document_session_store.get(doc_session_id) if doc_session_id else None

                if session_data:
                    base_structure = session_data["doc_structure"]
                    resolved_doc, applied_actions = DocumentLegalAuditor.apply_approved_decisions(
                        base_structure,
                        approved_decisions=approved_decisions if isinstance(approved_decisions, list) else [],
                    )
                    docx_bytes = build_audited_document_docx_bytes(
                        resolved_doc,
                        original_bytes=session_data.get("original_bytes"),
                        applied_actions=applied_actions,
                    )
                    doc_title = resolved_doc.get("title", "Minuta_Auditada")
                else:
                    doc_structure = data.get("doc_structure")
                    if doc_structure and isinstance(approved_decisions, list):
                        doc_structure, _ = DocumentLegalAuditor.apply_approved_decisions(
                            doc_structure,
                            approved_decisions=approved_decisions,
                        )
                    elif not doc_structure:
                        title = data.get("title", "DOCUMENTO AUDITADO (LEI 14.133/2021)")
                        clauses = data.get("clauses", [])
                        doc_structure = {
                            "title": title,
                            "original_format": "docx",
                            "elements": [
                                {
                                    "type": "CLAUSE",
                                    "clause_number": c.get("num") or c.get("clause_number") or "",
                                    "section_title": c.get("title") or "",
                                    "text": c.get("text") or "",
                                    "modified": bool(c.get("modified")),
                                    "diff_segments": c.get("diff_segments") or [
                                        {"type": "added" if c.get("modified") else "equal", "text": c.get("text") or ""}
                                    ],
                                }
                                for c in clauses
                            ],
                        }
                    docx_bytes = build_audited_document_docx_bytes(doc_structure)
                    doc_title = doc_structure.get("title", "Minuta_Auditada")

                safe_title = re.sub(r"[^a-zA-Z0-9_-]", "_", str(doc_title))[:80]
                self._send_binary(
                    docx_bytes,
                    "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    f"{safe_title}_Revisado_Amarelo_Vermelho.docx",
                )
            except Exception as exc:
                self._send_json(500, {"error": str(exc)})
            return

        # 3. Exportar Documento PDF (.PDF) com Layout Preservado + Alterações Aprovadas
        if parsed.path == "/api/v1/rag/export-pdf":
            try:
                data = json.loads(raw_body)
                doc_session_id = str(data.get("doc_session_id") or "").strip()
                approved_decisions = data.get("approved_decisions")
                session_data = document_session_store.get(doc_session_id) if doc_session_id else None

                if session_data:
                    doc_structure, _ = DocumentLegalAuditor.apply_approved_decisions(
                        session_data["doc_structure"],
                        approved_decisions=approved_decisions if isinstance(approved_decisions, list) else [],
                    )
                else:
                    doc_structure = data.get("doc_structure")
                    if doc_structure and isinstance(approved_decisions, list):
                        doc_structure, _ = DocumentLegalAuditor.apply_approved_decisions(
                            doc_structure,
                            approved_decisions=approved_decisions,
                        )
                    elif not doc_structure:
                        title = data.get("title", "DOCUMENTO AUDITADO (LEI 14.133/2021)")
                        clauses = data.get("clauses", [])
                        doc_structure = {
                            "title": title,
                            "original_format": "pdf",
                            "elements": [
                                {
                                    "type": "CLAUSE",
                                    "clause_number": c.get("num") or c.get("clause_number") or "",
                                    "section_title": c.get("title") or "",
                                    "text": c.get("text") or "",
                                    "modified": bool(c.get("modified")),
                                    "diff_segments": c.get("diff_segments") or [
                                        {"type": "added" if c.get("modified") else "equal", "text": c.get("text") or ""}
                                    ],
                                }
                                for c in clauses
                            ],
                        }
                pdf_bytes = build_audited_document_pdf_bytes(doc_structure)
                safe_title = re.sub(r"[^a-zA-Z0-9_-]", "_", str(doc_structure.get("title", "Minuta_Auditada")))[:80]
                self._send_binary(
                    pdf_bytes,
                    "application/pdf",
                    f"{safe_title}_Revisado_Amarelo_Vermelho.pdf",
                )
            except Exception as exc:
                self._send_json(500, {"error": str(exc)})
            return

        self._send_json(404, {"error": "Endpoint POST não encontrado."})


if __name__ == "__main__":
    port = int(os.environ.get("PORT", "8080"))
    lan_ip = get_primary_lan_ip()
    try:
        from http.server import ThreadingHTTPServer as _BaseSrv
    except ImportError:
        _BaseSrv = HTTPServer

    class ResilientGovTechServer(_BaseSrv):
        allow_reuse_address = True
        daemon_threads = True

    server = ResilientGovTechServer(("0.0.0.0", port), GovTechAppHandler)
    print("=" * 72)
    print("🚀 SISTEMA GOVTECH NLLC ONLINE (LEI 14.133/2021 & IN 65/2021)")
    print("=" * 72)
    print(f"👉 Acesso Local na sua Máquina:        http://localhost:{port}")
    print(f"👉 Acesso Compartilhado na Rede Wi-Fi: http://{lan_ip}:{port}")
    print("=" * 72)
    print("💡 Para compartilhar com qualquer pessoa na mesma rede Wi-Fi, envie o link:")
    print(f"   http://{lan_ip}:{port}")
    print("=" * 72)
    server.serve_forever()
