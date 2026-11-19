# -*- coding: utf-8 -*-
"""
Helper para resolução de compras públicas no PNCP (Portal Nacional de Contratações Públicas)
e extração/geração do Relatório Oficial em PDF com lista completa de materiais e termos.
"""
import io
import re
import requests
from datetime import datetime
from reportlab import rl_config as _rl_config
_rl_config.useA85 = 0
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import cm, mm
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, HRFlowable

SESSION = requests.Session()
SESSION.headers.update({
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) GovTech-AuditEngine/1.0",
    "Accept": "application/json"
})

import os
import subprocess
import tempfile
import threading
import time
from xml.sax.saxutils import escape as xml_escape
import qrcode
from reportlab.platypus import Image as RLImage, KeepTogether, PageBreak

CHROME_PATHS = [
    r"C:\Program Files\Google\Chrome\Application\chrome.exe",
    r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
    r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
    "/usr/bin/chromium",
    "/usr/bin/chromium-browser",
    "/usr/bin/google-chrome",
]

CACHE_DIR = os.path.join(tempfile.gettempdir(), "pncp_pdf_cache")
os.makedirs(CACHE_DIR, exist_ok=True)

_PNCP_CACHE_LOCK = threading.Lock()
_PNCP_COMPRA_CACHE: dict = {}
_PNCP_CONSULTA_CACHE: dict = {}


def build_comprasnet_direct_url(chave_or_id: str) -> str:
    """Retorna a URL real e direta de acompanhamento no Compras.gov.br (nunca link genérico de busca)."""
    digits = re.sub(r"\D", "", str(chave_or_id or ""))
    if len(digits) == 16:
        digits = "0" + digits
    if len(digits) >= 17:
        return f"https://cnetmobile.estaleiro.serpro.gov.br/comprasnet-web/public/compras/acompanhamento-compra?compra={digits[:17]}"
    return ""


def parse_pncp_control_id(ctrl_id: str):
    """Extrai (cnpj, ano, seq_int, seq_6) de um Id contratação PNCP (ex: 00394502000144-1-010123/2026)."""
    m = re.match(r"^(\d{14})-\d+-(\d+)/(\d{4})$", str(ctrl_id or "").strip())
    if not m:
        return None
    cnpj = m.group(1)
    seq_int = int(m.group(2))
    ano = m.group(3)
    return cnpj, ano, seq_int, str(seq_int).zfill(6)


def fetch_pncp_compra_consulta(cnpj: str, ano, sequencial, fast: bool = False) -> dict:
    """
    Consulta a API oficial pública de consulta do PNCP:
    GET https://pncp.gov.br/api/consulta/v1/orgaos/{cnpj}/compras/{ano}/{int(sequencial)}
    Retorna os dados completos da página do PNCP (valores estimado/homologado, linkSistemaOrigem,
    datas de proposta, amparo legal, situação, SRP, objeto, unidade compradora).
    """
    if not (cnpj and ano and sequencial):
        return {}
    try:
        seq_int = int(str(sequencial).lstrip("0") or "0")
    except ValueError:
        return {}
    if seq_int <= 0:
        return {}
    cache_key = f"{cnpj}/{ano}/{seq_int}"
    with _PNCP_CACHE_LOCK:
        if cache_key in _PNCP_CONSULTA_CACHE:
            return _PNCP_CONSULTA_CACHE[cache_key]

    url = f"https://pncp.gov.br/api/consulta/v1/orgaos/{cnpj}/compras/{ano}/{seq_int}"
    result = {}
    max_attempts = 1 if fast else 2
    req_timeout = 3.5 if fast else 8
    for attempt in range(max_attempts):
        try:
            resp = SESSION.get(url, timeout=req_timeout)
            if resp.status_code == 200:
                data = resp.json()
                if isinstance(data, dict):
                    result = data
                    break
            elif not fast and resp.status_code in (429, 500, 502, 503, 504):
                time.sleep(0.4 * (attempt + 1))
                continue
            else:
                break
        except Exception:
            if not fast:
                time.sleep(0.3)

    with _PNCP_CACHE_LOCK:
        if len(_PNCP_CONSULTA_CACHE) > 2000:
            _PNCP_CONSULTA_CACHE.clear()
        _PNCP_CONSULTA_CACHE[cache_key] = result
    return result


def resolve_pncp_webpage_url(chave_or_id: str) -> str:
    """Retorna a URL literal direta da página da contratação no Portal PNCP (ex: https://pncp.gov.br/app/editais/00394502000144/2026/10123)."""
    if not chave_or_id:
        return None
    parsed = parse_pncp_control_id(chave_or_id)
    if parsed:
        cnpj, ano, seq_int, _ = parsed
        return f"https://pncp.gov.br/app/editais/{cnpj}/{ano}/{seq_int}"

    compra = resolve_pncp_compra(chave_or_id)
    if compra:
        ctrl = compra.get("numeroControlePNCP") or compra.get("numero_controle_pncp")
        if ctrl:
            p2 = parse_pncp_control_id(ctrl)
            if p2:
                cnpj, ano, seq_int, _ = p2
                return f"https://pncp.gov.br/app/editais/{cnpj}/{ano}/{seq_int}"
        cnpj = compra.get("orgaoEntidadeCnpj") or compra.get("cnpjOrgao") or (compra.get("orgaoEntidade") or {}).get("cnpj")
        ano = compra.get("anoCompraPncp") or compra.get("anoCompra")
        seq_raw = compra.get("sequencialCompraPncp") or compra.get("sequencialCompra")
        if cnpj and ano and seq_raw:
            try:
                seq_int = int(str(seq_raw))
                if seq_int > 0:
                    return f"https://pncp.gov.br/app/editais/{cnpj}/{ano}/{seq_int}"
            except ValueError:
                pass
    return None


def fetch_pncp_itens(cnpj: str, ano: int, sequencial: int, max_items: int = 150) -> list:
    """Busca a lista oficial completa de materiais e itens da contratação no PNCP."""
    if not (cnpj and ano and sequencial):
        return []
    url = f"https://pncp.gov.br/api/pncp/v1/orgaos/{cnpj}/compras/{ano}/{sequencial}/itens?pagina=1&tamanhoPagina={max_items}"
    try:
        resp = SESSION.get(url, timeout=12)
        if resp.status_code == 200:
            data = resp.json()
            if isinstance(data, list):
                return data
            if isinstance(data, dict) and "data" in data:
                return data["data"]
    except Exception as e:
        print(f"[PNCP Itens Fetch Error] {e}")
    return []


def build_official_pncp_html(compra: dict, itens: list, arquivos: list, pncp_webpage_url: str = "") -> str:
    """Gera um documento HTML oficial e completo do PNCP com todos os materiais e termos para impressão direta em PDF via Chrome."""
    orgao = compra.get("orgaoEntidadeRazaoSocial") or "ÓRGÃO PÚBLICO FEDERAL"
    uasg = compra.get("unidadeOrgaoCodigoUnidade") or compra.get("codigoOrgao") or "—"
    uf = compra.get("unidadeOrgaoUfSigla") or "BR"
    cnpj = compra.get("orgaoEntidadeCnpj") or "—"
    num_compra = compra.get("numeroCompra") or "—"
    ano_compra = compra.get("anoCompraPncp") or datetime.now().year
    num_controle = compra.get("numeroControlePNCP") or f"{cnpj}-1-{str(compra.get('sequencialCompraPncp', 0)).zfill(6)}/{ano_compra}"
    processo = compra.get("processo") or "—"
    objeto = compra.get("objetoCompra") or "Aquisição de Bens / Prestação de Serviços"
    modalidade = compra.get("modalidadeNome") or "Pregão Eletrônico"
    modo_disputa = compra.get("modoDisputaNomePncp") or compra.get("modoDisputaNome") or "Aberto / Fechado"
    amparo = compra.get("amparoLegalNome") or "Lei 14.133/2021"
    amparo_desc = compra.get("amparoLegalDescricao") or ""
    
    val_estimado = float(compra.get("valorTotalEstimado") or 0)
    val_homologado = float(compra.get("valorTotalHomologado") or 0)
    
    dt_pub = compra.get("dataPublicacaoPncp") or compra.get("dataInclusaoPncp") or ""
    if dt_pub and "T" in dt_pub:
        dt_pub = dt_pub.split("T")[0]
        
    portal_url = pncp_webpage_url or f"https://pncp.gov.br/app/editais/{cnpj}/{ano_compra}/{str(compra.get('sequencialCompraPncp', 1)).zfill(6)}"

    # Gera linhas da tabela de materiais/itens
    items_rows_html = ""
    for it in (itens or []):
        num_it = it.get("numeroItem") or "—"
        desc = it.get("descricao") or "Item da contratação"
        unid = it.get("unidadeMedida") or "UN"
        qtd = float(it.get("quantidade") or 1)
        v_unit_est = float(it.get("valorUnitarioEstimado") or 0)
        v_unit_hom = float(it.get("valorUnitarioHomologado") or 0)
        v_tot = float(it.get("valorTotal") or (v_unit_hom or v_unit_est) * qtd)
        fornec = it.get("nomeRazaoSocialFornecedor") or "—"
        cnpj_fornec = it.get("niFornecedor") or ""
        fornec_info = f"{fornec}" + (f"<br><span style='font-size:9px;color:#666;'>CNPJ: {cnpj_fornec}</span>" if cnpj_fornec else "")
        
        items_rows_html += f"""
        <tr>
            <td style="text-align:center;font-weight:bold;">{num_it}</td>
            <td><strong>{desc}</strong></td>
            <td style="text-align:center;">{qtd:.0f} {unid}</td>
            <td style="text-align:right;white-space:nowrap;">R$ {v_unit_est:,.2f}</td>
            <td style="text-align:right;white-space:nowrap;color:#082849;font-weight:bold;">R$ {v_unit_hom:,.2f}</td>
            <td style="text-align:right;white-space:nowrap;font-weight:bold;">R$ {v_tot:,.2f}</td>
            <td style="font-size:10px;">{fornec_info}</td>
        </tr>
        """

    if not items_rows_html:
        items_rows_html = """
        <tr>
            <td colspan="7" style="text-align:center;padding:12px;color:#777;">
                Item único registrado na contratação ou itens disponíveis nos termos anexos ao processo.
            </td>
        </tr>
        """

    # Gera linhas da tabela de arquivos/termos
    arquivos_rows_html = ""
    for arq in (arquivos or []):
        tipo_doc = arq.get("tipoDocumentoNome") or "Termo Oficial"
        titulo = arq.get("titulo") or "Documento Anexo"
        seq_doc = arq.get("sequencialDocumento") or 1
        dl_url = f"https://pncp.gov.br/pncp-api/v1/orgaos/{cnpj}/compras/{ano_compra}/{compra.get('sequencialCompraPncp', 1)}/arquivos/{seq_doc}"
        arquivos_rows_html += f"""
        <tr>
            <td style="font-weight:bold;color:#082849;">{tipo_doc}</td>
            <td>{titulo}</td>
            <td style="text-align:center;"><a href="{dl_url}" style="color:#0c4a78;font-weight:bold;text-decoration:none;">📥 Baixar no PNCP</a></td>
        </tr>
        """

    if not arquivos_rows_html:
        arquivos_rows_html = """
        <tr>
            <td colspan="3" style="text-align:center;padding:10px;color:#777;">
                Arquivos e editais disponíveis diretamente no link oficial do PNCP informado no cabeçalho.
            </td>
        </tr>
        """

    html = f"""<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="utf-8">
<title>PNCP - Edital e Lista de Materiais</title>
<style>
    @page {{
        size: A4;
        margin: 12mm 15mm 15mm 15mm;
    }}
    body {{
        font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Arial, sans-serif;
        color: #1a202c;
        background: #ffffff;
        margin: 0;
        padding: 0;
        font-size: 11px;
        line-height: 1.4;
    }}
    .header-box {{
        border-bottom: 2px solid #082849;
        padding-bottom: 10px;
        margin-bottom: 12px;
        display: flex;
        justify-content: space-between;
        align-items: center;
    }}
    .pncp-brand {{
        font-size: 16px;
        font-weight: 900;
        color: #082849;
        letter-spacing: -0.5px;
    }}
    .pncp-sub {{
        font-size: 10px;
        color: #4a5568;
        font-weight: bold;
        text-transform: uppercase;
    }}
    .doc-pill {{
        background: #082849;
        color: #ffffff;
        font-size: 10px;
        font-weight: bold;
        padding: 4px 8px;
        border-radius: 4px;
        text-align: right;
    }}
    .grid-info {{
        display: grid;
        grid-template-columns: 1fr 1fr;
        gap: 8px 16px;
        background: #f8fafc;
        border: 1px solid #e2e8f0;
        border-radius: 6px;
        padding: 10px 12px;
        margin-bottom: 14px;
    }}
    .info-item {{
        font-size: 10.5px;
    }}
    .info-label {{
        font-weight: bold;
        color: #4a5568;
        text-transform: uppercase;
        font-size: 9.5px;
        display: block;
    }}
    .info-val {{
        color: #0f172a;
        font-weight: 600;
    }}
    .section-title {{
        font-size: 12px;
        font-weight: 800;
        color: #082849;
        text-transform: uppercase;
        margin-top: 14px;
        margin-bottom: 6px;
        border-bottom: 1px solid #cbd5e1;
        padding-bottom: 3px;
        display: flex;
        justify-content: space-between;
        align-items: center;
    }}
    table {{
        width: 100%;
        border-collapse: collapse;
        margin-bottom: 14px;
        font-size: 10px;
    }}
    th {{
        background: #082849;
        color: #ffffff;
        font-weight: bold;
        text-align: left;
        padding: 6px 8px;
        font-size: 9.5px;
        text-transform: uppercase;
    }}
    td {{
        padding: 6px 8px;
        border-bottom: 1px solid #e2e8f0;
        vertical-align: top;
    }}
    tr:nth-child(even) td {{
        background: #f8fafc;
    }}
    .footer-audit {{
        border-top: 1px dashed #cbd5e1;
        padding-top: 8px;
        font-size: 9px;
        color: #64748b;
        display: flex;
        justify-content: space-between;
        align-items: center;
        margin-top: 16px;
    }}
</style>
</head>
<body>

    <div class="header-box">
        <div>
            <div class="pncp-brand">PORTAL NACIONAL DE CONTRATAÇÕES PÚBLICAS (PNCP)</div>
            <div class="pncp-sub">REPÚBLICA FEDERATIVA DO BRASIL • LEI Nº 14.133/2021</div>
        </div>
        <div class="doc-pill">
            EXTRATO OFICIAL DA CONTRATAÇÃO<br>
            <span style="font-weight:normal;font-size:9px;">Trilha de Auditoria Digital</span>
        </div>
    </div>

    <div class="grid-info">
        <div class="info-item">
            <span class="info-label">Órgão Comprador</span>
            <span class="info-val">{orgao} (UASG: {uasg} - {uf})</span>
        </div>
        <div class="info-item">
            <span class="info-label">CNPJ do Órgão</span>
            <span class="info-val">{cnpj}</span>
        </div>
        <div class="info-item">
            <span class="info-label">Nº da Compra / Ano</span>
            <span class="info-val">{modalidade} Nº {num_compra}/{ano_compra}</span>
        </div>
        <div class="info-item">
            <span class="info-label">Nº de Controle PNCP</span>
            <span class="info-val" style="color:#0c4a78;font-family:monospace;">{num_controle}</span>
        </div>
        <div class="info-item" style="grid-column: span 2;">
            <span class="info-label">Objeto da Contratação</span>
            <span class="info-val">{objeto}</span>
        </div>
        <div class="info-item">
            <span class="info-label">Processo Administrativo</span>
            <span class="info-val">{processo}</span>
        </div>
        <div class="info-item">
            <span class="info-label">Modo de Disputa / Amparo Legal</span>
            <span class="info-val">{modo_disputa} ({amparo})</span>
        </div>
        <div class="info-item">
            <span class="info-label">Valor Total Estimado</span>
            <span class="info-val">R$ {val_estimado:,.2f}</span>
        </div>
        <div class="info-item">
            <span class="info-label">Valor Total Homologado</span>
            <span class="info-val" style="color:#15803d;font-weight:bold;">R$ {val_homologado:,.2f}</span>
        </div>
        <div class="info-item" style="grid-column: span 2;">
            <span class="info-label">Link de Acesso Oficial no PNCP</span>
            <span class="info-val"><a href="{portal_url}" style="color:#0c4a78;text-decoration:none;">{portal_url} ↗</a></span>
        </div>
    </div>

    <div class="section-title">
        <span>Itens e Materiais da Contratação ({len(itens)} itens registrados)</span>
        <span style="font-size:10px;font-weight:normal;color:#64748b;">Valores Homologados / Adjudicados</span>
    </div>

    <table>
        <thead>
            <tr>
                <th style="width:35px;text-align:center;">Item</th>
                <th>Descrição do Material / Serviço</th>
                <th style="width:75px;text-align:center;">Qtd/Unid</th>
                <th style="width:85px;text-align:right;">Unit. Est.</th>
                <th style="width:90px;text-align:right;">Unit. Homolog.</th>
                <th style="width:90px;text-align:right;">Valor Total</th>
                <th style="width:140px;">Fornecedor Homologado</th>
            </tr>
        </thead>
        <tbody>
            {items_rows_html}
        </tbody>
    </table>

    <div class="section-title">
        <span>Termos e Documentos Oficiais Anexados</span>
        <span style="font-size:10px;font-weight:normal;color:#64748b;">Publicados no PNCP</span>
    </div>

    <table>
        <thead>
            <tr>
                <th style="width:180px;">Tipo do Documento</th>
                <th>Título / Identificação do Arquivo</th>
                <th style="width:120px;text-align:center;">Acesso Oficial</th>
            </tr>
        </thead>
        <tbody>
            {arquivos_rows_html}
        </tbody>
    </table>

    <div class="footer-audit">
        <div>
            Documento emitido eletronicamente via Sistema GovTech NLLC integrado às APIs Oficiais do PNCP e Compras.gov.br.<br>
            Autenticidade passível de confirmação mediante consulta ao Controle PNCP nº {num_controle}.
        </div>
        <div style="text-align:right;">
            Data de Emissão: {datetime.now().strftime('%d/%m/%Y às %H:%M:%S')}
        </div>
    </div>

</body>
</html>
"""
    return html


def capture_pncp_page_pdf(pncp_url: str, chave_or_id: str = None, timeout_sec: int = 25) -> bytes:
    """
    Gera o PDF A4 Oficial e literal do edital no Portal PNCP com lista completa de materiais e termos.
    Utiliza renderização completa em alta resolução via Google Chrome headless sem atrasos de SPA.
    """
    compra = None
    if chave_or_id:
        compra = resolve_pncp_compra(chave_or_id)
    if not compra and pncp_url:
        m = re.search(r"editais/(\d+)/(\d+)/(\d+)", pncp_url)
        if m:
            compra = {
                "orgaoEntidadeCnpj": m.group(1),
                "anoCompraPncp": int(m.group(2)),
                "sequencialCompraPncp": int(m.group(3))
            }

    clean_id = re.sub(r"\W", "_", pncp_url.split("editais/")[-1] if (pncp_url and "editais/" in pncp_url) else str(chave_or_id or "pncp_doc"))
    cache_path = os.path.join(CACHE_DIR, f"pncp_full_{clean_id}.pdf")
    if os.path.exists(cache_path) and os.path.getsize(cache_path) > 2000:
        with open(cache_path, "rb") as f:
            return f.read()

    browser_exe = None
    for p in CHROME_PATHS:
        if os.path.exists(p):
            browser_exe = p
            break

    # Se encontramos a compra nas APIs oficiais do Governo, montamos o dossiê completo de materiais
    if compra:
        cnpj = compra.get("orgaoEntidadeCnpj")
        ano = compra.get("anoCompraPncp")
        seq = compra.get("sequencialCompraPncp")
        
        # Se os dados detalhados da compra ainda não foram buscados
        if not compra.get("orgaoEntidadeRazaoSocial") and chave_or_id:
            c_full = resolve_pncp_compra(chave_or_id)
            if c_full:
                compra = c_full
                
        itens = fetch_pncp_itens(cnpj, ano, seq)
        arquivos = fetch_pncp_arquivos(cnpj, ano, seq)
        
        html_content = build_official_pncp_html(compra, itens, arquivos, pncp_webpage_url=pncp_url or "")
        
        # Se temos o Google Chrome instalado, renderiza o HTML completo em PDF (Leva ~1.2s e é 100% perfeito)
        if browser_exe:
            temp_html = os.path.join(CACHE_DIR, f"tmp_{int(time.time()*1000)}.html")
            temp_pdf = os.path.join(CACHE_DIR, f"tmp_{int(time.time()*1000)}.pdf")
            try:
                with open(temp_html, "w", encoding="utf-8") as f:
                    f.write(html_content)
                cmd = [
                    browser_exe,
                    "--headless",
                    "--disable-gpu",
                    "--no-sandbox",
                    f"--print-to-pdf={temp_pdf}",
                    temp_html
                ]
                subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, timeout=timeout_sec)
                if os.path.exists(temp_pdf) and os.path.getsize(temp_pdf) > 1000:
                    with open(temp_pdf, "rb") as f:
                        data = f.read()
                    try:
                        os.replace(temp_pdf, cache_path)
                    except Exception:
                        pass
                    return data
            except Exception as e:
                print(f"[Chrome Local HTML Render Error] {e}")
            finally:
                if os.path.exists(temp_html):
                    try:
                        os.remove(temp_html)
                    except Exception:
                        pass
                if os.path.exists(temp_pdf):
                    try:
                        os.remove(temp_pdf)
                    except Exception:
                        pass
                        
        # Fallback se não tiver Chrome: gera via ReportLab em memória
        return generate_pncp_dossier_pdf(compra, arquivos, quote_info={"pncp_or_compras_uri": pncp_url})

    # Caso não tenhamos a chave estruturada, chama Chrome headless diretamente na URL com virtual-time-budget
    if browser_exe and pncp_url:
        temp_out = os.path.join(CACHE_DIR, f"tmp_{int(time.time()*1000)}.pdf")
        cmd = [
            browser_exe,
            "--headless=new",
            "--disable-gpu",
            "--no-sandbox",
            "--virtual-time-budget=8000",
            f"--print-to-pdf={temp_out}",
            pncp_url
        ]
        try:
            subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, timeout=timeout_sec)
            if os.path.exists(temp_out) and os.path.getsize(temp_out) > 500:
                with open(temp_out, "rb") as f:
                    data = f.read()
                try:
                    os.replace(temp_out, cache_path)
                except Exception:
                    pass
                return data
        except Exception as exc:
            print(f"[PNCP Direct Capture Error] {exc}")
        finally:
            if os.path.exists(temp_out):
                try:
                    os.remove(temp_out)
                except Exception:
                    pass

    return None


def resolve_pncp_compra(chave_or_id: str):
    """
    Resolve os dados oficiais da contratação no PNCP com cache em memória (thread-safe) e retry:
    - Aceita Id contratação PNCP (ex: 00394502000144-1-010123/2026) -> consulta direta /api/consulta/v1
    - Aceita idCompra (16 ou 17 dígitos, ex: 75121206000702026) -> Dados Abertos 1.1_consultarContratacoes_PNCP_14133_Id
    """
    raw = str(chave_or_id or "").strip()
    if not raw:
        return None

    # Caso 1: Já é um Id contratação PNCP (ex: 00394502000144-1-010123/2026)
    parsed_ctrl = parse_pncp_control_id(raw)
    if parsed_ctrl:
        cnpj, ano, seq_int, seq_6 = parsed_ctrl
        consulta = fetch_pncp_compra_consulta(cnpj, ano, seq_int)
        if consulta:
            org = consulta.get("orgaoEntidade") or {}
            unid = consulta.get("unidadeOrgao") or {}
            amparo = consulta.get("amparoLegal") or {}
            return {
                "numeroControlePNCP": consulta.get("numeroControlePNCP") or f"{cnpj}-1-{seq_6}/{ano}",
                "orgaoEntidadeCnpj": org.get("cnpj") or cnpj,
                "orgaoEntidadeRazaoSocial": org.get("razaoSocial") or "",
                "unidadeOrgaoCodigoUnidade": unid.get("codigoUnidade") or "",
                "unidadeOrgaoNomeUnidade": unid.get("nomeUnidade") or "",
                "unidadeOrgaoMunicipioNome": unid.get("municipioNome") or "",
                "unidadeOrgaoUfSigla": unid.get("ufSigla") or "",
                "anoCompraPncp": int(consulta.get("anoCompra") or ano),
                "sequencialCompraPncp": int(consulta.get("sequencialCompra") or seq_int),
                "numeroCompra": consulta.get("numeroCompra") or "",
                "processo": consulta.get("processo") or "",
                "objetoCompra": consulta.get("objetoCompra") or "",
                "modalidadeNome": consulta.get("modalidadeNome") or "",
                "tipoInstrumentoConvocatorioNome": consulta.get("tipoInstrumentoConvocatorioNome") or "",
                "modoDisputaNomePncp": consulta.get("modoDisputaNome") or "",
                "amparoLegalNome": amparo.get("nome") or "",
                "amparoLegalDescricao": amparo.get("descricao") or "",
                "situacaoCompraNome": consulta.get("situacaoCompraNome") or "Divulgada no PNCP",
                "srp": bool(consulta.get("srp")),
                "valorTotalEstimado": consulta.get("valorTotalEstimado"),
                "valorTotalHomologado": consulta.get("valorTotalHomologado"),
                "dataPublicacaoPncp": consulta.get("dataPublicacaoPncp") or "",
                "dataAtualizacao": consulta.get("dataAtualizacao") or "",
                "dataAberturaProposta": consulta.get("dataAberturaProposta") or "",
                "dataEncerramentoProposta": consulta.get("dataEncerramentoProposta") or "",
                "linkSistemaOrigem": consulta.get("linkSistemaOrigem") or "",
            }
        return None

    # Caso 2: idCompra / idCompraItem (16 a 22 dígitos)
    chave_clean = re.sub(r"\D", "", raw)
    if len(chave_clean) == 16:
        chave_clean = "0" + chave_clean
    chave_17 = chave_clean[:17] if len(chave_clean) >= 17 else chave_clean
    if len(chave_17) < 12:
        return None

    with _PNCP_CACHE_LOCK:
        if chave_17 in _PNCP_COMPRA_CACHE:
            return _PNCP_COMPRA_CACHE[chave_17]

    url_gov = "https://dadosabertos.compras.gov.br/modulo-contratacoes/1.1_consultarContratacoes_PNCP_14133_Id"
    found = None
    candidates = [chave_17]
    if chave_17.startswith("0"):
        candidates.append(chave_17.lstrip("0"))

    for cand in candidates:
        for attempt in range(3):
            try:
                resp = SESSION.get(url_gov, params={"tipo": "idCompra", "codigo": cand}, timeout=8)
                if resp.status_code == 200:
                    res = resp.json().get("resultado") or []
                    if res:
                        found = res[0]
                        break
                    break
                elif resp.status_code in (429, 500, 502, 503, 504):
                    time.sleep(0.35 * (attempt + 1))
                    continue
                else:
                    break
            except Exception:
                time.sleep(0.3)
        if found:
            break

    with _PNCP_CACHE_LOCK:
        if len(_PNCP_COMPRA_CACHE) > 4000:
            _PNCP_COMPRA_CACHE.clear()
        _PNCP_COMPRA_CACHE[chave_17] = found
    return found


def fetch_pncp_arquivos(cnpj: str, ano: int, sequencial: int):
    """Busca a lista oficial de arquivos e termos publicados no PNCP."""
    if not (cnpj and ano and sequencial):
        return []
    url = f"https://pncp.gov.br/api/pncp/v1/orgaos/{cnpj}/compras/{ano}/{sequencial}/arquivos"
    try:
        resp = SESSION.get(url, timeout=9)
        if resp.status_code == 200:
            return resp.json()
    except Exception:
        pass
    return []


def download_pncp_file_bytes(cnpj: str, ano: int, sequencial: int, seq_doc: int):
    """Faz o download em memória de um arquivo do PNCP."""
    url = f"https://pncp.gov.br/pncp-api/v1/orgaos/{cnpj}/compras/{ano}/{sequencial}/arquivos/{seq_doc}"
    resp = SESSION.get(url, timeout=20)
    if resp.status_code == 200:
        return resp.content, resp.headers.get("Content-Type", "application/pdf")
    return None, None


def _make_qr_rl_image(url: str, size_cm: float = 2.6):
    """Gera QR Code compacto apontando diretamente para a página da contratação no PNCP."""
    qr = qrcode.QRCode(
        version=None,
        error_correction=qrcode.constants.ERROR_CORRECT_L,
        box_size=4,
        border=1,
    )
    qr.add_data(url or "https://pncp.gov.br")
    qr.make(fit=True)
    img = qr.make_image(fill_color="black", back_color="white")
    buf = io.BytesIO()
    img.save(buf, format="PNG")
    buf.seek(0)
    return RLImage(buf, width=size_cm * cm, height=size_cm * cm)


def _fmt_brl_pncp(val) -> str:
    try:
        v = float(val or 0.0)
    except Exception:
        v = 0.0
    return f"R$ {v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def _fmt_date_br(raw_dt: str) -> str:
    s = str(raw_dt or "").strip()
    if not s:
        return "—"
    m = re.match(r"^(\d{4})-(\d{2})-(\d{2})(?:[T ](\d{2}:\d{2}))?", s)
    if m:
        base = f"{m.group(3)}/{m.group(2)}/{m.group(1)}"
        return f"{base} {m.group(4)}" if m.group(4) else base
    return s[:16]


def _build_pncp_comprovante_story_block(
    compra_data: dict,
    arquivos: list,
    quote_info: dict,
    linked_quotes: list = None,
) -> list:
    """
    Constrói os flowables de 1 página fiel à tela oficial do Portal Nacional de Contratações
    Públicas (PNCP — ex: https://pncp.gov.br/app/editais/00394502000144/2026/10123),
    com QR Code oficial no topo direito, Id contratação PNCP, cards de Valor Total Estimado
    e Homologado, e links clicáveis tanto para o PNCP quanto para o Compras.gov.br.
    """
    styles = getSampleStyleSheet()
    st_portal = ParagraphStyle("PNCPPortal", parent=styles["Normal"], fontName="Helvetica-Bold", fontSize=9.5, leading=12, textColor=colors.HexColor("#1e3a5f"))
    st_h1 = ParagraphStyle("PNCPH1", parent=styles["Normal"], fontName="Helvetica-Bold", fontSize=14, leading=17, textColor=colors.HexColor("#0f172a"))
    st_upd = ParagraphStyle("PNCPUpd", parent=styles["Normal"], fontName="Helvetica-Oblique", fontSize=8, leading=10, textColor=colors.HexColor("#64748b"))
    st_body = ParagraphStyle("PNCPBody", parent=styles["Normal"], fontName="Helvetica", fontSize=8.4, leading=11.5, textColor=colors.HexColor("#1e293b"))
    st_lbl = ParagraphStyle("PNCPLbl", parent=styles["Normal"], fontName="Helvetica-Bold", fontSize=8.4, leading=11.5, textColor=colors.HexColor("#0f172a"))
    st_card_lbl = ParagraphStyle("PNCPCardLbl", parent=styles["Normal"], fontName="Helvetica-Bold", fontSize=7.2, leading=9.5, textColor=colors.HexColor("#0f172a"))
    st_card_val = ParagraphStyle("PNCPCardVal", parent=styles["Normal"], fontName="Helvetica-Bold", fontSize=11.5, leading=14, textColor=colors.HexColor("#0f172a"))
    st_hdr = ParagraphStyle("PNCPHdr", parent=styles["Normal"], fontName="Helvetica-Bold", fontSize=7.4, leading=9.5, textColor=colors.white, alignment=1)
    st_td = ParagraphStyle("PNCPTd", parent=styles["Normal"], fontName="Helvetica", fontSize=7.2, leading=9.2, textColor=colors.HexColor("#1e293b"))
    st_td_r = ParagraphStyle("PNCPTdR", parent=styles["Normal"], fontName="Helvetica-Bold", fontSize=7.2, leading=9.2, textColor=colors.HexColor("#0f172a"), alignment=2)

    compra = compra_data if isinstance(compra_data, dict) else {}
    qinfo = quote_info if isinstance(quote_info, dict) else {}

    # Extração combinada (Dados Abertos + PNCP Consulta v1 + Quote)
    org_obj = compra.get("orgaoEntidade") if isinstance(compra.get("orgaoEntidade"), dict) else {}
    unid_obj = compra.get("unidadeOrgao") if isinstance(compra.get("unidadeOrgao"), dict) else {}
    amp_obj = compra.get("amparoLegal") if isinstance(compra.get("amparoLegal"), dict) else {}

    cnpj = str(compra.get("orgaoEntidadeCnpj") or org_obj.get("cnpj") or "").strip()
    orgao_nome = str(compra.get("orgaoEntidadeRazaoSocial") or org_obj.get("razaoSocial") or qinfo.get("buyer_org_name") or "ÓRGÃO PÚBLICO FEDERAL").strip()
    uasg = str(compra.get("unidadeOrgaoCodigoUnidade") or unid_obj.get("codigoUnidade") or qinfo.get("buyer_uasg") or "—").strip()
    unid_nome = str(compra.get("unidadeOrgaoNomeUnidade") or unid_obj.get("nomeUnidade") or qinfo.get("nome_uasg") or "").strip()
    mun = str(compra.get("unidadeOrgaoMunicipioNome") or unid_obj.get("municipioNome") or qinfo.get("municipio") or "").strip()
    uf = str(compra.get("unidadeOrgaoUfSigla") or unid_obj.get("ufSigla") or qinfo.get("buyer_uf") or "BR").strip()
    localidade = f"{mun}/{uf}" if mun else uf

    num_compra = str(compra.get("numeroCompra") or qinfo.get("purchase_number_year") or "—").strip()
    ano = str(compra.get("anoCompraPncp") or compra.get("anoCompra") or "").strip()
    seq_raw = compra.get("sequencialCompraPncp") or compra.get("sequencialCompra")
    num_controle = str(compra.get("numeroControlePNCP") or qinfo.get("pncp_control_number") or "").strip()

    if not num_controle and cnpj and ano and seq_raw:
        try:
            num_controle = f"{cnpj}-1-{str(int(seq_raw)).zfill(6)}/{ano}"
        except Exception:
            pass

    parsed_ctrl = parse_pncp_control_id(num_controle)
    if parsed_ctrl:
        c_cnpj, c_ano, c_seq_int, _ = parsed_ctrl
        pncp_url = f"https://pncp.gov.br/app/editais/{c_cnpj}/{c_ano}/{c_seq_int}"
        cnpj = cnpj or c_cnpj
        ano = ano or c_ano
    else:
        raw_uri = str(qinfo.get("pncp_or_compras_uri") or "").strip()
        pncp_url = raw_uri if ("/app/editais/" in raw_uri and "?q=" not in raw_uri) else ""

    chave_17 = str(compra.get("idCompra") or qinfo.get("purchase_key") or "").strip()
    compras_url = str(compra.get("linkSistemaOrigem") or qinfo.get("comprasnet_uri") or build_comprasnet_direct_url(chave_17) or "").strip()
    primary_qr_url = pncp_url or compras_url or "https://pncp.gov.br"

    tipo_inst = str(compra.get("tipoInstrumentoConvocatorioNome") or compra.get("modalidadeNome") or qinfo.get("forma_compra") or "Contratação Pública").strip()
    titulo_compra = f"{tipo_inst} nº {num_compra}" + (f"/{ano}" if ano and "/" not in num_compra else "")
    modalidade = str(compra.get("modalidadeNome") or qinfo.get("forma_compra") or "—").strip()
    amparo = str(compra.get("amparoLegalNome") or amp_obj.get("nome") or "Lei nº 14.133/2021").strip()
    modo_disp = str(compra.get("modoDisputaNomePncp") or compra.get("modoDisputaNome") or "—").strip()
    srp_str = "Sim" if compra.get("srp") else "Não"
    situacao = str(compra.get("situacaoCompraNome") or "Divulgada no PNCP").strip()
    dt_pub = _fmt_date_br(compra.get("dataPublicacaoPncp") or compra.get("dataInclusaoPncp") or qinfo.get("purchase_date") or "")
    dt_upd = _fmt_date_br(compra.get("dataAtualizacao") or dt_pub)
    dt_ini = _fmt_date_br(compra.get("dataAberturaProposta") or "")
    dt_fim = _fmt_date_br(compra.get("dataEncerramentoProposta") or "")
    objeto = str(compra.get("objetoCompra") or qinfo.get("objeto_compra") or qinfo.get("item_description") or "Contratação Pública Homologada no Compras.gov.br / PNCP.").strip()

    val_est = compra.get("valorTotalEstimado")
    val_hom = compra.get("valorTotalHomologado")

    story = []

    # Cabeçalho estilo Portal PNCP + QR Code no topo direito (espelho da Imagem 2!)
    qr_img = _make_qr_rl_image(primary_qr_url, size_cm=2.5)
    left_hdr = [
        Paragraph("Portal Nacional de Contratações Públicas (PNCP) • Comprovante Oficial de Contratação", st_portal),
        Spacer(1, 0.12 * cm),
        Paragraph(xml_escape(titulo_compra), st_h1),
        Paragraph(f"Última atualização: {xml_escape(dt_upd)} &nbsp;|&nbsp; Data de divulgação no PNCP: {xml_escape(dt_pub)}", st_upd),
    ]
    t_top = Table([[left_hdr, qr_img]], colWidths=[15.0 * cm, 3.0 * cm])
    t_top.setStyle(TableStyle([
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("ALIGN", (1, 0), (1, 0), "RIGHT"),
        ("LEFTPADDING", (0, 0), (-1, -1), 0),
        ("RIGHTPADDING", (0, 0), (-1, -1), 0),
    ]))
    story.append(t_top)
    story.append(HRFlowable(width="100%", thickness=1.2, color=colors.HexColor("#1e3a5f"), spaceBefore=4, spaceAfter=8))

    # Bloco de Metadados Oficiais PNCP
    cnpj_fmt = (
        f"{cnpj[:2]}.{cnpj[2:5]}.{cnpj[5:8]}/{cnpj[8:12]}-{cnpj[12:14]}"
        if len(cnpj) == 14
        else (cnpj or "—")
    )
    meta_rows = [
        [
            Paragraph(f"<b>Órgão/Entidade:</b> {xml_escape(cnpj_fmt)} - {xml_escape(orgao_nome)}", st_body),
            Paragraph(f"<b>Unidade compradora:</b> {xml_escape(uasg)}" + (f" - {xml_escape(unid_nome)}" if unid_nome else ""), st_body),
        ],
        [
            Paragraph(f"<b>Localidade da Unidade:</b> {xml_escape(localidade)}", st_body),
            Paragraph(f"<b>Fonte/Plataforma:</b> Compras.gov.br (ID Compra: {xml_escape(chave_17 or '—')})", st_body),
        ],
        [
            Paragraph(f"<b>Modalidade:</b> {xml_escape(modalidade)} &nbsp;|&nbsp; <b>Amparo legal:</b> {xml_escape(amparo)}", st_body),
            Paragraph(f"<b>Modo de disputa:</b> {xml_escape(modo_disp)} &nbsp;|&nbsp; <b>SRP:</b> {srp_str}", st_body),
        ],
        [
            Paragraph(f"<b>Id contratação PNCP:</b> <font color='#0f766e'><b>{xml_escape(num_controle or 'Não indexado (Compras.gov.br direto)')}</b></font>", st_body),
            Paragraph(f"<b>Situação:</b> {xml_escape(situacao)} &nbsp;|&nbsp; <b>Propostas:</b> {xml_escape(dt_ini)} a {xml_escape(dt_fim)}", st_body),
        ],
    ]
    t_meta = Table(meta_rows, colWidths=[9.5 * cm, 8.5 * cm])
    t_meta.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#f8fafc")),
        ("BOX", (0, 0), (-1, -1), 0.5, colors.HexColor("#cbd5e1")),
        ("INNERGRID", (0, 0), (-1, -1), 0.35, colors.HexColor("#e2e8f0")),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("TOPPADDING", (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
        ("LEFTPADDING", (0, 0), (-1, -1), 6),
        ("RIGHTPADDING", (0, 0), (-1, -1), 6),
    ]))
    story.append(t_meta)
    story.append(Spacer(1, 0.22 * cm))

    # Objeto
    story.append(Paragraph(f"<b>Objeto:</b> {xml_escape(objeto[:500])}", st_body))
    story.append(Spacer(1, 0.25 * cm))

    # Cards de Valor Total Estimado e Valor Total Homologado (espelho visual dos cards da Imagem 2!)
    val_est_str = _fmt_brl_pncp(val_est) if val_est not in (None, "") else "Consultar no PNCP"
    val_hom_str = _fmt_brl_pncp(val_hom) if val_hom not in (None, "") else "Consultar no PNCP"
    t_cards = Table(
        [[
            Paragraph(f"VALOR TOTAL ESTIMADO DA COMPRA<br/><font size='11'><b>{val_est_str}</b></font>", st_card_lbl),
            Paragraph(f"VALOR TOTAL HOMOLOGADO DA COMPRA<br/><font size='11'><b>{val_hom_str}</b></font>", st_card_lbl),
            Paragraph(
                "<b>ACESSO DIRETO VERIFICÁVEL:</b><br/>"
                + (f'• <link href="{xml_escape(pncp_url)}" color="#1d4ed8"><u>Abrir Página no PNCP</u></link><br/>' if pncp_url else "")
                + (f'• <link href="{xml_escape(compras_url)}" color="#0f766e"><u>Abrir Acompanhamento Compras.gov.br</u></link>' if compras_url else ""),
                st_body,
            ),
        ]],
        colWidths=[5.2 * cm, 5.2 * cm, 7.6 * cm],
    )
    t_cards.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (0, 0), colors.HexColor("#e0f2fe")),
        ("BACKGROUND", (1, 0), (1, 0), colors.HexColor("#ccfbf1")),
        ("BACKGROUND", (2, 0), (2, 0), colors.HexColor("#f8fafc")),
        ("BOX", (0, 0), (0, 0), 0.7, colors.HexColor("#38bdf8")),
        ("BOX", (1, 0), (1, 0), 0.7, colors.HexColor("#2dd4bf")),
        ("BOX", (2, 0), (2, 0), 0.5, colors.HexColor("#cbd5e1")),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING", (0, 0), (-1, -1), 6),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
        ("LEFTPADDING", (0, 0), (-1, -1), 7),
        ("RIGHTPADDING", (0, 0), (-1, -1), 7),
    ]))
    story.append(t_cards)
    story.append(Spacer(1, 0.3 * cm))

    # Cotações / Itens da Pesquisa vinculados a esta compra
    q_list = linked_quotes if linked_quotes else ([qinfo] if qinfo.get("supplier_name") else [])
    if q_list:
        story.append(Paragraph("<b>ITENS E FORNECEDORES VINCULADOS NA PESQUISA DE PREÇOS (IN 65/2021):</b>", st_lbl))
        story.append(Spacer(1, 0.1 * cm))
        q_rows = [[
            Paragraph("Item / Cód.", st_hdr),
            Paragraph("Fornecedor Vencedor", st_hdr),
            Paragraph("CNPJ", st_hdr),
            Paragraph("Unid. / Qtd", st_hdr),
            Paragraph("Preço Unitário", st_hdr),
        ]]
        for lq in q_list[:25]:
            it_ref = lq.get("item_ref") or lq.get("catmat_code") or "—"
            q_rows.append([
                Paragraph(xml_escape(str(it_ref)), st_td),
                Paragraph(xml_escape(str(lq.get("supplier_name") or "—")[:55]), st_td),
                Paragraph(xml_escape(str(lq.get("supplier_cnpj") or "—")), st_td),
                Paragraph(xml_escape(f"{lq.get('supply_unit') or 'UN'} / {float(lq.get('quantity_homologated') or 1):g}"), st_td),
                Paragraph(_fmt_brl_pncp(lq.get("unit_price")), st_td_r),
            ])
        t_q = Table(q_rows, colWidths=[2.6 * cm, 6.8 * cm, 3.4 * cm, 2.5 * cm, 2.7 * cm], repeatRows=1)
        t_q.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1e293b")),
            ("GRID", (0, 0), (-1, -1), 0.35, colors.HexColor("#cbd5e1")),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("TOPPADDING", (0, 0), (-1, -1), 3),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
        ]))
        story.append(t_q)
        story.append(Spacer(1, 0.25 * cm))

    # Documentos / Arquivos publicados no PNCP (se disponíveis)
    if arquivos:
        story.append(Paragraph(f"<b>DOCUMENTOS PUBLICADOS NO PNCP ({len(arquivos)}):</b>", st_lbl))
        story.append(Spacer(1, 0.08 * cm))
        arq_rows = [[Paragraph("Tipo", st_hdr), Paragraph("Título", st_hdr), Paragraph("Publicação", st_hdr)]]
        for a in arquivos[:15]:
            arq_rows.append([
                Paragraph(xml_escape(str(a.get("tipoDocumentoNome") or "Documento")), st_td),
                Paragraph(xml_escape(str(a.get("titulo") or "Anexo")[:80]), st_td),
                Paragraph(_fmt_date_br(a.get("dataInclusao") or ""), st_td),
            ])
        t_a = Table(arq_rows, colWidths=[5.0 * cm, 10.0 * cm, 3.0 * cm], repeatRows=1)
        t_a.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#334155")),
            ("GRID", (0, 0), (-1, -1), 0.35, colors.HexColor("#cbd5e1")),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("TOPPADDING", (0, 0), (-1, -1), 2.5),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 2.5),
        ]))
        story.append(t_a)
        story.append(Spacer(1, 0.2 * cm))

    # Rodapé com URLs literais completas
    story.append(HRFlowable(width="100%", thickness=0.5, color=colors.HexColor("#cbd5e1"), spaceAfter=4))
    url_footer = []
    if pncp_url:
        url_footer.append(f"<b>URL Oficial PNCP:</b> {xml_escape(pncp_url)}")
    if compras_url:
        url_footer.append(f"<b>URL Compras.gov.br:</b> {xml_escape(compras_url)}")
    story.append(Paragraph(
        " &nbsp;|&nbsp; ".join(url_footer)
        + f"<br/><i>Comprovante auditável emitido em {datetime.now().strftime('%d/%m/%Y às %H:%M')} (Lei nº 14.133/2021 e IN SEGES/ME nº 65/2021). Escaneie o QR Code acima para conferir a contratação ao vivo.</i>",
        st_upd,
    ))
    return story


def generate_pncp_dossier_pdf(compra_data: dict, arquivos: list, quote_info: dict = None) -> bytes:
    """Gera o PDF Comprovante Oficial da Contratação no PNCP com QR Code e cards de valores."""
    compra = dict(compra_data or {})
    qinfo = dict(quote_info or {})

    # Enriquece automaticamente com /api/consulta/v1 para trazer Valor Total Estimado e Homologado oficiais
    ctrl = compra.get("numeroControlePNCP") or qinfo.get("pncp_control_number")
    parsed = parse_pncp_control_id(ctrl) if ctrl else None
    if parsed and (compra.get("valorTotalEstimado") is None or not compra.get("linkSistemaOrigem")):
        consulta = fetch_pncp_compra_consulta(parsed[0], parsed[1], parsed[2])
        if consulta:
            for k, v in consulta.items():
                if compra.get(k) in (None, ""):
                    compra[k] = v

    buf = io.BytesIO()
    doc = SimpleDocTemplate(
        buf,
        pagesize=A4,
        leftMargin=1.5 * cm,
        rightMargin=1.5 * cm,
        topMargin=1.3 * cm,
        bottomMargin=1.3 * cm,
    )
    story = _build_pncp_comprovante_story_block(compra, arquivos or [], qinfo)
    doc.build(story)
    return buf.getvalue()


def build_pncp_comprovantes_caderno_pdf_bytes(
    items_results: list,
    om_name: str = "MARINHA DO BRASIL",
    max_purchases: int = 80,
) -> bytes:
    """
    Gera o Caderno Consolidado de Comprovantes Oficiais PNCP (PDF A4 com 1 página por
    contratação única referenciada nas cotações válidas do lote), contendo QR Code direto,
    Id contratação PNCP, valores estimado/homologado e links clicáveis PNCP + Compras.gov.br.
    """
    import concurrent.futures

    # Agrupa todas as cotações válidas por contratação única (pncp_control_number ou purchase_key)
    unique_purchases: dict = {}
    for idx, it in enumerate(items_results, start=1):
        rep = it.get("report") or {}
        code = str(rep.get("catmat_catser_code") or it.get("code") or "—")
        for q in (rep.get("valid_quotes") or []):
            ctrl = str(q.get("pncp_control_number") or "").strip()
            pkey = str(q.get("purchase_key") or "").strip()
            uri = str(q.get("pncp_or_compras_uri") or "").strip()
            if not ctrl and "/app/editais/" in uri:
                m_u = re.match(r"https://pncp\.gov\.br/app/editais/(\d{14})/(\d{4})/(\d+)", uri)
                if m_u:
                    ctrl = f"{m_u.group(1)}-1-{m_u.group(3).zfill(6)}/{m_u.group(2)}"
            uid = ctrl or pkey or uri
            if not uid:
                continue
            if uid not in unique_purchases:
                unique_purchases[uid] = {
                    "ctrl": ctrl,
                    "pkey": pkey,
                    "sample_quote": dict(q),
                    "linked_quotes": [],
                }
            lq = dict(q)
            lq["item_ref"] = f"Item {idx} ({code})"
            unique_purchases[uid]["linked_quotes"].append(lq)

    entries = list(unique_purchases.values())[:max_purchases]

    # Hidratação paralela segura (limita chamadas externas simultâneas para manter ZIP ultrarrápido no Render)
    def _hydrate(idx_entry):
        e_pos, entry = idx_entry
        ctrl = entry["ctrl"]
        pkey = entry["pkey"]
        compra = {}
        # Limita chamadas HTTP externas ao PNCP a no máximo 12 compras não cacheadas com timeout curto (~2s total)
        if e_pos < 12:
            if ctrl:
                parsed = parse_pncp_control_id(ctrl)
                if parsed:
                    compra = fetch_pncp_compra_consulta(parsed[0], parsed[1], parsed[2], fast=True) or {}
                    if compra and not compra.get("numeroControlePNCP"):
                        compra["numeroControlePNCP"] = ctrl
            elif pkey and len(re.sub(r"\D", "", pkey)) >= 16:
                c_da = resolve_pncp_compra(pkey)
                if c_da:
                    compra = dict(c_da)
                    c_ctrl = compra.get("numeroControlePNCP")
                    if c_ctrl:
                        entry["ctrl"] = c_ctrl
                        p2 = parse_pncp_control_id(c_ctrl)
                        if p2:
                            cons = fetch_pncp_compra_consulta(p2[0], p2[1], p2[2], fast=True)
                            if cons:
                                for k, v in cons.items():
                                    if compra.get(k) in (None, ""):
                                        compra[k] = v
        entry["compra"] = compra
        return entry

    if entries:
        with concurrent.futures.ThreadPoolExecutor(max_workers=8) as pool:
            entries = list(pool.map(_hydrate, list(enumerate(entries))))

    buf = io.BytesIO()
    doc = SimpleDocTemplate(
        buf,
        pagesize=A4,
        leftMargin=1.5 * cm,
        rightMargin=1.5 * cm,
        topMargin=1.3 * cm,
        bottomMargin=1.3 * cm,
    )
    styles = getSampleStyleSheet()
    story = []

    if not entries:
        story.append(Paragraph(f"<b>{xml_escape(om_name.upper())} — COMPROVANTES DE CONTRATAÇÕES PNCP</b>", styles["Heading2"]))
        story.append(Paragraph("Nenhuma contratação pública com chave PNCP/Compras.gov.br identificada no lote.", styles["Normal"]))
    else:
        for e_idx, entry in enumerate(entries):
            if e_idx > 0:
                story.append(PageBreak())
            q_sample = dict(entry["sample_quote"])
            if entry["ctrl"] and not q_sample.get("pncp_control_number"):
                q_sample["pncp_control_number"] = entry["ctrl"]
            story.extend(
                _build_pncp_comprovante_story_block(
                    entry.get("compra") or {},
                    [],
                    q_sample,
                    linked_quotes=entry["linked_quotes"],
                )
            )

    doc.build(story)
    return buf.getvalue()
