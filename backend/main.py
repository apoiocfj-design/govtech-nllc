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
import html
import io
import json
import os
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
from report_generators import build_in65_docx_bytes, build_in65_pdf_bytes

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


def build_in65_excel_bytes(report: dict, quantity: int, uasg_name: str = "MARINHA DO BRASIL - CAPITANIA FLUVIAL DE JUAZEIRO") -> bytes:
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
    row_vals = [
        1,
        report["item_description"],
        str(report["catmat_catser_code"]),
        report["target_supply_unit"],
        quantity,
        float(sm["mean"] or unit_est),
        float(sm["median"] or unit_est),
        unit_est,
        quantity * unit_est,
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

    # Aba 2: Trilha de Auditoria das Cotações Válidas e Expurgadas
    ws2 = wb.create_sheet("Auditoria PNCP Compras.gov")
    cols2 = ["Status", "UASG", "Órgão Comprador", "UF", "Nº Compra", "CNPJ Vencedor", "Fornecedor", "Data", "Unid.", "Qtd.", "Preço Unit. (R$)", "Justificativa Estatística", "Link Oficial Auditável"]
    for i, h in enumerate(cols2, 1):
        c = ws2.cell(1, i, h)
        c.font, c.fill, c.alignment, c.border = Font(bold=True), cinza, ctr, B

    all_quotes = report.get("valid_quotes", []) + report.get("purged_quotes", [])
    for r_idx, q in enumerate(all_quotes, 2):
        vals = [
            q["purge_status"], q["buyer_uasg"], q["buyer_org_name"], q["buyer_uf"],
            q["purchase_number_year"], q["supplier_cnpj"], q["supplier_name"],
            q["purchase_date"], q["supply_unit"], q["quantity_homologated"],
            float(q["unit_price"]), q.get("purge_justification") or "", q["pncp_or_compras_uri"]
        ]
        for c_idx, v in enumerate(vals, 1):
            cel = ws2.cell(r_idx, c_idx, v)
            cel.border = B
            if c_idx == 11:
                cel.number_format = brl
            if c_idx == 13 and v:
                cel.hyperlink = v
                cel.style = "Hyperlink"

    for col_letter, width in zip("ABCDEFGHIJKLM", [24, 10, 36, 6, 14, 18, 36, 12, 8, 10, 15, 45, 55]):
        ws2.column_dimensions[col_letter].width = width

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def build_highlighted_docx_bytes(title: str, clauses: list) -> bytes:
    """Gera em memória um arquivo Word (.DOCX) com todas as cláusulas editadas marcadas em AMARELO."""
    p_list = []

    def add_p_xml(text_str, bold=False, highlight=False, size=22, color=None):
        rpr = ['<w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/>', f'<w:sz w:val="{size}"/>']
        if bold:
            rpr.append("<w:b/>")
        if highlight:
            rpr.append('<w:highlight w:val="yellow"/>')
        if color:
            rpr.append(f'<w:color w:val="{color}"/>')
        rpr.append("</w:rPr>")
        return (
            '<w:p><w:pPr><w:spacing w:after="120" w:line="260" w:lineRule="auto"/></w:pPr>'
            f'<w:r>{"".join(rpr)}<w:t xml:space="preserve">{html.escape(text_str, quote=True)}</w:t></w:r></w:p>'
        )

    p_list.append(add_p_xml("MARINHA DO BRASIL - CAPITANIA FLUVIAL DE JUAZEIRO", bold=True, size=28, color="003366"))
    p_list.append(add_p_xml(title, bold=True, size=24, color="222222"))
    p_list.append(add_p_xml(
        "NOTA DE AUDITORIA: Todas as cláusulas corrigidas/inseridas pelo Motor RAG (Padrão AGU / Lei 14.133/2021) estão destacadas em AMARELO abaixo.",
        bold=True, highlight=True, size=20
    ))

    for c in clauses:
        num = c.get("num") or c.get("clause_number") or ""
        sec = c.get("title") or c.get("section_title") or ""
        txt = c.get("text") or c.get("content") or ""
        mod = bool(c.get("modified") or c.get("is_modified_by_agu"))
        p_list.append(add_p_xml(f"[{num} - {sec}]", bold=True, highlight=mod, size=20, color="555555"))
        p_list.append(add_p_xml(txt, bold=mod, highlight=mod, size=22))

    content_types = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
        '</Types>'
    )
    rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
        '</Relationships>'
    )
    doc_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        f'<w:body>{"".join(p_list)}</w:body></w:document>'
    )

    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", content_types)
        z.writestr("_rels/.rels", rels)
        z.writestr("word/document.xml", doc_xml)
    return buf.getvalue()


class GovTechAppHandler(BaseHTTPRequestHandler):
    def _send_json(self, status_code: int, payload: dict):
        body = json.dumps(payload, ensure_ascii=False, indent=2).encode("utf-8")
        self.send_response(status_code)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def _send_binary(self, content: bytes, content_type: str, filename: str):
        self.send_response(200)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Disposition", f'attachment; filename="{filename}"')
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Content-Length", str(len(content)))
        self.end_headers()
        self.wfile.write(content)

    def do_GET(self):
        parsed = urlparse(self.path)
        path = parsed.path
        qs = parse_qs(parsed.query)

        # 1. Interface Unificada do Aplicativo
        if path in ("/", "/index.html"):
            with open(FRONTEND_HTML_PATH, "rb") as f:
                content = f.read()
            self.send_response(200)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.send_header("Content-Length", str(len(content)))
            self.end_headers()
            self.wfile.write(content)
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
            supply_unit = qs.get("supply_unit", ["UN"])[0]
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
            supply_unit = qs.get("supply_unit", ["UN"])[0]
            item_type = qs.get("item_type", ["MATERIAL"])[0]
            quantity = int(qs.get("qty", ["1"])[0])
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
            supply_unit = qs.get("supply_unit", ["UN"])[0]
            item_type = qs.get("item_type", ["MATERIAL"])[0]
            quantity = int(qs.get("qty", ["1"])[0])
            months_back = int(qs.get("months_back", ["12"])[0])
            uf = qs.get("uf", [None])[0] or None
            elaborador = qs.get("elaborador", ["ENCARREGADO DO SETOR DE OBTENÇÃO"])[0]
            ordenador = qs.get("ordenador", ["ORDENADOR DE DESPESAS"])[0]
            om_name = qs.get("om_name", ["CAPITANIA FLUVIAL DE JUAZEIRO"])[0]

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
            supply_unit = qs.get("supply_unit", ["UN"])[0]
            item_type = qs.get("item_type", ["MATERIAL"])[0]
            quantity = int(qs.get("qty", ["1"])[0])
            months_back = int(qs.get("months_back", ["12"])[0])
            uf = qs.get("uf", [None])[0] or None
            elaborador = qs.get("elaborador", ["ENCARREGADO DO SETOR DE OBTENÇÃO"])[0]
            ordenador = qs.get("ordenador", ["ORDENADOR DE DESPESAS"])[0]
            om_name = qs.get("om_name", ["CAPITANIA FLUVIAL DE JUAZEIRO"])[0]

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

        self._send_json(404, {"error": "Rota não encontrada."})

    def do_POST(self):
        parsed = urlparse(self.path)
        length = int(self.headers.get("Content-Length", "0"))
        raw_body = self.rfile.read(length).decode("utf-8") if length > 0 else "{}"

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
                org_name = str(data.get("org_name", "MARINHA DO BRASIL - CAPITANIA FLUVIAL DE JUAZEIRO")).strip()

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

        # 1. Auditar Arquivo Enviado (PDF ou DOCX em Base64) ou Texto Colado
        if parsed.path == "/api/v1/rag/audit-upload":
            try:
                data = json.loads(raw_body)
                artifact_type = data.get("artifact_type", "TR")
                filename = data.get("filename", "")
                file_b64 = data.get("file_base64", "")
                raw_text = data.get("text", "")

                extracted_lines = []
                if file_b64:
                    file_bytes = base64.b64decode(file_b64)
                    if filename.lower().endswith(".pdf"):
                        with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
                            for pg in pdf.pages:
                                t = pg.extract_text() or ""
                                extracted_lines.extend([l.strip() for l in t.split("\n") if l.strip()])
                    elif filename.lower().endswith(".docx"):
                        tmp_path = os.path.join(BASE_DIR, "_tmp_upload.docx")
                        with open(tmp_path, "wb") as f:
                            f.write(file_bytes)
                        extracted_lines = OpenXMLDocxReader.extract_paragraphs_from_docx(tmp_path)
                    else:
                        extracted_lines = [l.strip() for l in file_bytes.decode("utf-8", errors="ignore").split("\n") if l.strip()]
                else:
                    extracted_lines = [l.strip() for l in raw_text.split("\n") if l.strip()]

                chunks = AGUClauseChunker.chunk_paragraphs(extracted_lines, artifact_type=artifact_type)
                audit_report = rag_auditor.audit_user_chunks(chunks, target_artifact_type=artifact_type)
                audit_report["clauses"] = [
                    {
                        "num": c.clause_number,
                        "title": c.section_title,
                        "text": c.text,
                        "modified": False,
                    }
                    for c in chunks[:60]
                ]
                self._send_json(200, audit_report)
            except Exception as exc:
                self._send_json(400, {"error": str(exc)})
            return

        # 2. Exportar Documento Word (.DOCX) com Destaque Amarelo
        if parsed.path == "/api/v1/rag/export-docx":
            try:
                data = json.loads(raw_body)
                title = data.get("title", "TERMO DE REFERÊNCIA AUDITADO (LEI 14.133/2021)")
                clauses = data.get("clauses", [])
                docx_bytes = build_highlighted_docx_bytes(title, clauses)
                self._send_binary(
                    docx_bytes,
                    "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    "Documento_Auditado_AGU_Marcado_Amarelo.docx",
                )
            except Exception as exc:
                self._send_json(500, {"error": str(exc)})
            return

        self._send_json(404, {"error": "Endpoint POST não encontrado."})


if __name__ == "__main__":
    port = int(os.environ.get("PORT", "8080"))
    lan_ip = get_primary_lan_ip()
    server = HTTPServer(("0.0.0.0", port), GovTechAppHandler)
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
