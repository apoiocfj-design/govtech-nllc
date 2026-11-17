"""
Cliente de Consulta e Motor de Busca para Intenções de Registro de Preços (IRP) e Oportunidades SRP
Integrado ao Portal Nacional de Contratações Públicas (PNCP) e Compras.gov.br
Regulado pela Lei nº 14.133/2021 (Art. 86) e Decreto nº 11.462/2023.
"""

import requests
import json
import time
from datetime import datetime
from typing import Dict, Any, List, Optional

HEADERS = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
    'Accept': 'application/json'
}

# Cache em memória para buscas e detalhes com TTL de 5 minutos
_IRP_SEARCH_CACHE: Dict[str, Any] = {}
_IRP_DETAIL_CACHE: Dict[str, Any] = {}
CACHE_TTL_SECONDS = 300


def _parse_datetime(dt_str: Optional[str]) -> Optional[datetime]:
    if not dt_str:
        return None
    try:
        clean = dt_str.replace("Z", "").split(".")[0]
        if "T" in clean:
            parts = clean.split("T")
            time_parts = parts[1].split(":")
            if len(time_parts) == 3:
                return datetime.strptime(clean, "%Y-%m-%dT%H:%M:%S")
            return datetime.strptime(clean, "%Y-%m-%dT%H:%M")
        return datetime.strptime(clean, "%Y-%m-%d")
    except Exception:
        return None


def _format_datetime_br(dt_str: Optional[str]) -> str:
    dt = _parse_datetime(dt_str)
    if not dt:
        return dt_str or "Não informado"
    return dt.strftime("%d/%m/%Y às %H:%M")


def _calculate_days_remaining(prazo_str: Optional[str]) -> Optional[int]:
    dt = _parse_datetime(prazo_str)
    if not dt:
        return None
    now = datetime.now()
    diff = dt - now
    if diff.total_seconds() < 0:
        return -1
    return diff.days


def search_irps(
    q: str = "",
    status: str = "aberta",
    uf: str = "",
    pagina: int = 1,
    tam_pagina: int = 20
) -> Dict[str, Any]:
    """
    Pesquisa Intenções de Registro de Preços (IRP) ou Editais SRP no PNCP.
    status:
      - 'aberta': IRPs Abertas para Manifestação
      - 'em_andamento': IRPs Em Andamento / Divulgadas
      - 'todas': Todas as IRPs no PNCP
      - 'srp_edital': Editais de Licitação SRP (Pregões / Concorrências) recebendo propostas
    """
    q_clean = (q or "").strip()
    status_clean = (status or "aberta").strip().lower()
    uf_clean = (uf or "").strip().upper()
    pagina_num = max(1, int(pagina))
    tam_pagina_num = min(50, max(10, int(tam_pagina)))

    is_srp_edital = status_clean in ["srp_edital", "edital_srp", "edital"]

    cache_key = f"{q_clean.lower()}|{status_clean}|{uf_clean}|{pagina_num}|{tam_pagina_num}"
    cached = _IRP_SEARCH_CACHE.get(cache_key)
    if cached and (time.time() - cached["timestamp"] < CACHE_TTL_SECONDS):
        return cached["data"]

    url = "https://pncp.gov.br/api/search/"
    params = {
        "pagina": pagina_num,
        "tam_pagina": tam_pagina_num
    }

    if is_srp_edital:
        params["tipos_documento"] = "edital"
        params["srp"] = "true"
        params["status"] = "recebendo_proposta"
        if q_clean:
            params["q"] = q_clean
        else:
            params["q"] = "registro de precos"
    else:
        params["tipos_documento"] = "irp"
        if q_clean:
            params["q"] = q_clean
            if status_clean not in ["todas", "todos", ""]:
                params["status"] = status_clean
        else:
            params["status"] = status_clean if status_clean not in ["todas", "todos", ""] else "aberta"

    if uf_clean:
        params["uf"] = uf_clean

    try:
        resp = requests.get(url, params=params, headers=HEADERS, timeout=12)
        if resp.status_code != 200:
            return {
                "status": "ERROR",
                "message": f"Erro na consulta ao PNCP: HTTP {resp.status_code} - {resp.text[:120]}",
                "total": 0,
                "items": []
            }

        raw_data = resp.json()
        total = raw_data.get("total", 0)
        items_raw = raw_data.get("items", [])

        # Se houver filtro por UF mas a API do PNCP não filtrou estritamente no backend dela
        if uf_clean:
            items_filtered = [it for it in items_raw if str(it.get("uf") or "").upper() == uf_clean]
            if len(items_filtered) < len(items_raw) and len(items_filtered) > 0:
                items_raw = items_filtered

        formatted_items = []
        for it in items_raw:
            cnpj = it.get("orgao_cnpj") or ""
            ano = it.get("ano") or ""
            seq = it.get("numero_sequencial") or it.get("numero") or ""

            item_url = it.get("item_url") or ""
            if (not cnpj or not ano or not seq) and item_url:
                parts = item_url.strip("/").split("/")
                if len(parts) >= 4:
                    cnpj = cnpj or parts[1]
                    ano = ano or parts[2]
                    seq = seq or parts[3]

            prazo_raw = (
                it.get("prazo_manifestacao_interesse")
                or it.get("data_fim_recebimento_propostas")
                or it.get("data_fim_vigencia")
            )
            dias_rest = _calculate_days_remaining(prazo_raw)
            prazo_br = _format_datetime_br(prazo_raw)

            # Montagem visual de status
            doc_type = it.get("document_type") or ("edital" if is_srp_edital else "irp")
            is_edital = doc_type == "edital"

            if dias_rest is not None:
                if dias_rest > 5:
                    status_badge = {
                        "label": f"{'Recebendo Propostas' if is_edital else 'Aberta'} ({dias_rest} dias)",
                        "color": "emerald",
                        "is_open": True,
                        "urgency": "normal"
                    }
                elif dias_rest >= 0:
                    status_badge = {
                        "label": f"Crítico (Encerra em {dias_rest} dias)",
                        "color": "amber",
                        "is_open": True,
                        "urgency": "critical"
                    }
                else:
                    status_badge = {
                        "label": "Prazo Encerrado",
                        "color": "slate",
                        "is_open": False,
                        "urgency": "closed"
                    }
            else:
                status_badge = {
                    "label": "Recebendo Propostas (SRP)" if is_edital else "Aberta para Manifestação",
                    "color": "emerald",
                    "is_open": True,
                    "urgency": "normal"
                }

            # Link oficial
            if is_edital:
                link_pncp = f"https://pncp.gov.br/app/editais/{cnpj}/{ano}/{seq}" if (cnpj and ano and seq) else "https://pncp.gov.br"
                title_display = it.get("title") or f"Edital SRP nº {seq}/{ano}"
            else:
                link_pncp = f"https://pncp.gov.br/app/irp/{cnpj}/{ano}/{seq}" if (cnpj and ano and seq) else "https://pncp.gov.br"
                title_display = it.get("title") or f"IRP {seq}/{ano}"

            formatted_items.append({
                "id": it.get("id"),
                "document_type": doc_type,
                "is_edital": is_edital,
                "title": title_display,
                "description": it.get("description") or "Sem descrição informada no PNCP.",
                "numero_irp": seq,
                "ano_irp": ano,
                "numero_controle_pncp": it.get("numero_controle_pncp") or "",
                "orgao_cnpj": cnpj,
                "orgao_nome": it.get("orgao_nome") or "Órgão Não Informado",
                "unidade_nome": it.get("unidade_nome") or "",
                "uf": it.get("uf") or "",
                "municipio_nome": it.get("municipio_nome") or "",
                "esfera_nome": it.get("esfera_nome") or "Municipal",
                "modalidade_nome": it.get("modalidade_licitacao_nome") or "Pregão Eletrônico (SRP)",
                "permite_adesao": it.get("permite_adesao", True),
                "tipo_item_irp": it.get("tipo_item_irp") or "M",
                "prazo_manifestacao": prazo_raw,
                "prazo_manifestacao_formatado": prazo_br,
                "dias_restantes": dias_rest,
                "status_badge": status_badge,
                "data_publicacao": _format_datetime_br(it.get("data_publicacao_pncp")),
                "link_pncp": link_pncp,
                "cnpj_ano_seq": f"{cnpj}_{ano}_{seq}"
            })

        out = {
            "status": "SUCCESS",
            "total": total,
            "page": pagina_num,
            "page_size": tam_pagina_num,
            "is_edital_mode": is_srp_edital,
            "items": formatted_items
        }
        _IRP_SEARCH_CACHE[cache_key] = {"timestamp": time.time(), "data": out}
        return out
    except Exception as exc:
        return {
            "status": "ERROR",
            "message": f"Falha na comunicação com o PNCP: {str(exc)}",
            "total": 0,
            "items": []
        }


def get_irp_detail(cnpj: str, ano: str, sequencial: str) -> Dict[str, Any]:
    """
    Recupera os detalhes completos da IRP: cabeçalho, lista de itens e documentos anexos.
    """
    cnpj_clean = str(cnpj or "").strip().replace(".", "").replace("/", "").replace("-", "")
    ano_clean = str(ano or "").strip()
    seq_clean = str(sequencial or "").strip()

    cache_key = f"{cnpj_clean}_{ano_clean}_{seq_clean}"
    cached = _IRP_DETAIL_CACHE.get(cache_key)
    if cached and (time.time() - cached["timestamp"] < CACHE_TTL_SECONDS):
        return cached["data"]

    base_url = f"https://pncp.gov.br/api/pncp/v1/orgaos/{cnpj_clean}/irp/{ano_clean}/{seq_clean}"
    try:
        # 1. Dados cadastrais da IRP
        r_header = requests.get(base_url, headers=HEADERS, timeout=10)
        header_data = r_header.json() if r_header.status_code == 200 else {}

        # 2. Itens da IRP
        r_items = requests.get(f"{base_url}/itens", headers=HEADERS, timeout=10)
        items_data = r_items.json() if r_items.status_code == 200 else []

        # 3. Documentos/Arquivos anexados
        r_files = requests.get(f"{base_url}/arquivos", headers=HEADERS, timeout=10)
        files_data = r_files.json() if r_files.status_code == 200 else []

        prazo = header_data.get("dataHoraFinalManifestacaoInteresse")
        dias_rest = _calculate_days_remaining(prazo)
        prazo_br = _format_datetime_br(prazo)

        cleaned_items = []
        for it in items_data:
            cat_nome = it.get("categoriaItemNome") or ("Serviço" if it.get("materialOuServico") == "S" else "Material")
            cleaned_items.append({
                "numero_item": it.get("numeroItem"),
                "categoria": cat_nome,
                "tipo_codigo": "SERVICO" if it.get("materialOuServico") == "S" else "MATERIAL",
                "descricao": it.get("descricao") or "Item sem descrição informada",
                "unidade_fornecimento": it.get("unidadeFornecimento") or "UN",
                "criterio_julgamento": it.get("criterioJulgamentoNome") or "Menor Preço",
                "valor_unitario_estimado": float(it.get("valorUnitarioEstimado") or 0.0),
                "indicador_sigiloso": bool(it.get("indicadorValorSigiloso", False)),
                "codigo_item_catalogo": str(it.get("codigoItemCatalogo") or "")
            })

        cleaned_files = []
        for f in files_data:
            cleaned_files.append({
                "sequencial": f.get("sequencialDocumento"),
                "titulo": f.get("titulo") or f"Documento_{f.get('sequencialDocumento')}",
                "tipo": f.get("tipoDocumentoNome") or "Anexo / Termo de Referência",
                "url": f.get("urlDownload") or f"https://pncp.gov.br/pncp-api/v1/orgaos/{cnpj_clean}/irp/{ano_clean}/{seq_clean}/arquivos/{f.get('sequencialDocumento')}"
            })

        out = {
            "status": "SUCCESS",
            "cnpj": cnpj_clean,
            "ano": ano_clean,
            "sequencial": seq_clean,
            "numero_controle_pncp": header_data.get("numeroControlePNCP") or "",
            "processo": header_data.get("processo") or "",
            "objeto": header_data.get("objetoIrp") or "",
            "orgao_nome": header_data.get("orgaoEntidade", {}).get("razaoSocial") or "Órgão Não Informado",
            "unidade_nome": header_data.get("unidadeOrgao", {}).get("nomeUnidade") or "",
            "uf": header_data.get("unidadeOrgao", {}).get("ufSigla") or "",
            "municipio": header_data.get("unidadeOrgao", {}).get("municipioNome") or "",
            "modalidade": header_data.get("modalidadeNome") or "Pregão Eletrônico (SRP)",
            "amparo_legal": header_data.get("amparoLegalNome") or "Lei nº 14.133/2021",
            "prazo_manifestacao": prazo,
            "prazo_manifestacao_formatado": prazo_br,
            "dias_restantes": dias_rest,
            "total_itens": len(cleaned_items),
            "itens": cleaned_items,
            "total_arquivos": len(cleaned_files),
            "arquivos": cleaned_files,
            "link_oficial_pncp": f"https://pncp.gov.br/app/irp/{cnpj_clean}/{ano_clean}/{seq_clean}",
            "link_plataforma_origem": header_data.get("urlPlataformaOrigem") or ""
        }
        _IRP_DETAIL_CACHE[cache_key] = {"timestamp": time.time(), "data": out}
        return out
    except Exception as exc:
        return {
            "status": "ERROR",
            "message": f"Erro ao recuperar detalhes da IRP: {str(exc)}",
            "itens": [],
            "arquivos": []
        }
