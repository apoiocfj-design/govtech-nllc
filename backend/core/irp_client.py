"""
Cliente de Consulta e Motor de Busca Híbrido para Intenções de Registro de Preços (IRP),
Editais SRP e Atas de Registro de Preços (ARP / Adesão Carona) por Unidade da Federação (UF).
Integrado ao Portal Nacional de Contratações Públicas (PNCP) e Compras.gov.br.
Regulado pela Lei nº 14.133/2021 (Art. 86) e Decreto nº 11.462/2023.
"""

from __future__ import annotations

import time
from datetime import datetime
from typing import Any, Dict, List, Optional

import requests
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36"
    ),
    "Accept": "application/json, text/plain, */*",
    "Referer": "https://pncp.gov.br/app/irp",
    "Origin": "https://pncp.gov.br",
}

# Cache em memória para buscas e detalhes com TTL de 5 minutos
_IRP_SEARCH_CACHE: Dict[str, Any] = {}
_IRP_DETAIL_CACHE: Dict[str, Any] = {}
CACHE_TTL_SECONDS = 300

VALID_UFS = {
    "AC", "AL", "AM", "AP", "BA", "CE", "DF", "ES", "GO",
    "MA", "MG", "MS", "MT", "PA", "PB", "PE", "PI", "PR",
    "RJ", "RN", "RO", "RR", "RS", "SC", "SE", "SP", "TO",
}


def _build_resilient_session() -> requests.Session:
    """Cria sessão HTTP com connection pooling e retry automático para resiliência contra WAF/502/503/504."""
    session = requests.Session()
    session.headers.update(HEADERS)
    retry = Retry(
        total=3,
        read=2,
        connect=2,
        backoff_factor=0.6,
        status_forcelist=(429, 500, 502, 503, 504),
        allowed_methods=frozenset(["GET"]),
        raise_on_status=False,
    )
    adapter = HTTPAdapter(max_retries=retry, pool_connections=10, pool_maxsize=10)
    session.mount("https://", adapter)
    session.mount("http://", adapter)
    return session


_HTTP_SESSION = _build_resilient_session()


def _parse_datetime(dt_str: Optional[str]) -> Optional[datetime]:
    if not dt_str:
        return None
    try:
        clean = str(dt_str).replace("Z", "").split(".")[0].strip()
        if "T" in clean:
            parts = clean.split("T")
            time_parts = parts[1].split(":")
            if len(time_parts) == 3:
                return datetime.strptime(clean, "%Y-%m-%dT%H:%M:%S")
            return datetime.strptime(clean, "%Y-%m-%dT%H:%M")
        return datetime.strptime(clean[:10], "%Y-%m-%d")
    except Exception:
        return None


def _format_datetime_br(dt_str: Optional[str]) -> str:
    dt = _parse_datetime(dt_str)
    if not dt:
        return str(dt_str or "Não informado")
    if dt.hour == 0 and dt.minute == 0 and "T" not in str(dt_str or ""):
        return dt.strftime("%d/%m/%Y")
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


def _build_status_badge(dias_rest: Optional[int], doc_type: str) -> Dict[str, Any]:
    is_edital = doc_type == "edital"
    is_ata = doc_type == "ata"
    open_prefix = "Ata Vigente" if is_ata else ("Recebendo Propostas" if is_edital else "IRP Aberta")

    if dias_rest is not None:
        if dias_rest > 5:
            return {
                "label": f"{open_prefix} ({dias_rest} dias)",
                "color": "emerald",
                "is_open": True,
                "urgency": "normal",
            }
        if dias_rest >= 0:
            return {
                "label": f"Crítico (Encerra em {dias_rest} dias)",
                "color": "amber",
                "is_open": True,
                "urgency": "critical",
            }
        return {
            "label": "Prazo Encerrado",
            "color": "slate",
            "is_open": False,
            "urgency": "closed",
        }

    return {
        "label": (
            "Ata SRP Vigente (Adesão)"
            if is_ata
            else ("Recebendo Propostas (SRP)" if is_edital else "Aberta para Manifestação")
        ),
        "color": "emerald",
        "is_open": True,
        "urgency": "normal",
    }


def _format_search_item(it: Dict[str, Any], default_doc_type: str = "irp") -> Dict[str, Any]:
    """Padroniza um item vindo do índice /api/search/ do PNCP (irp, edital ou ata)."""
    doc_type = str(it.get("document_type") or default_doc_type).lower()
    is_edital = doc_type == "edital"
    is_ata = doc_type == "ata"

    cnpj = str(it.get("orgao_cnpj") or "").strip()
    ano = str(it.get("ano") or "").strip()
    seq = str(it.get("numero_sequencial") or it.get("numero") or "").strip()
    seq_compra_ata = str(it.get("numero_sequencial_compra_ata") or "").strip()

    item_url = str(it.get("item_url") or "").strip()
    if item_url:
        parts = [p for p in item_url.strip("/").split("/") if p]
        # Exemplos de item_url no PNCP:
        # /irp/{cnpj}/{ano}/{seq}
        # /compras/{cnpj}/{ano}/{seq}
        # /atas/{cnpj}/{ano}/{seq_compra}/{seq_ata}
        if len(parts) >= 4:
            cnpj = cnpj or parts[1]
            ano = ano or parts[2]
            if is_ata and len(parts) >= 5:
                seq_compra_ata = seq_compra_ata or parts[3]
                seq = seq or parts[4]
            else:
                seq = seq or parts[3]

    # Para atas de registro de preços, o endpoint de itens no PNCP usa o sequencial da COMPRA mãe
    detail_seq = seq_compra_ata if (is_ata and seq_compra_ata) else seq

    prazo_raw = (
        it.get("prazo_manifestacao_interesse")
        or it.get("data_fim_recebimento_propostas")
        or it.get("data_fim_vigencia")
    )
    dias_rest = _calculate_days_remaining(prazo_raw)
    prazo_br = _format_datetime_br(prazo_raw)
    status_badge = _build_status_badge(dias_rest, doc_type)

    if is_ata:
        if cnpj and ano and seq_compra_ata and seq:
            link_pncp = f"https://pncp.gov.br/app/atas/{cnpj}/{ano}/{seq_compra_ata}/{seq}"
        elif cnpj and ano and detail_seq:
            link_pncp = f"https://pncp.gov.br/app/editais/{cnpj}/{ano}/{detail_seq}"
        else:
            link_pncp = "https://pncp.gov.br/app/atas"
        title_display = it.get("title") or f"Ata SRP nº {seq}/{ano}"
    elif is_edital:
        link_pncp = (
            f"https://pncp.gov.br/app/editais/{cnpj}/{ano}/{detail_seq}"
            if (cnpj and ano and detail_seq)
            else "https://pncp.gov.br/app/editais"
        )
        title_display = it.get("title") or f"Edital SRP nº {detail_seq}/{ano}"
    else:
        link_pncp = (
            f"https://pncp.gov.br/app/irp/{cnpj}/{ano}/{detail_seq}"
            if (cnpj and ano and detail_seq)
            else "https://pncp.gov.br/app/irp"
        )
        title_display = it.get("title") or f"IRP {detail_seq}/{ano}"

    permite_adesao = it.get("permite_adesao")
    if permite_adesao is None:
        permite_adesao = True

    return {
        "id": it.get("id") or f"{doc_type}_{cnpj}_{ano}_{detail_seq}_{seq}",
        "document_type": doc_type,
        "is_edital": is_edital,
        "is_ata": is_ata,
        "title": title_display,
        "description": it.get("description") or "Sem descrição informada no PNCP.",
        "numero_irp": detail_seq,
        "numero_exibicao": it.get("numero") or seq or detail_seq,
        "numero_sequencial_compra_ata": seq_compra_ata,
        "ano_irp": ano,
        "numero_controle_pncp": it.get("numero_controle_pncp") or "",
        "orgao_cnpj": cnpj,
        "orgao_nome": it.get("orgao_nome") or "Órgão Não Informado",
        "unidade_nome": it.get("unidade_nome") or "",
        "uf": str(it.get("uf") or "").upper(),
        "municipio_nome": it.get("municipio_nome") or "",
        "esfera_nome": it.get("esfera_nome") or "Federal / Estadual / Municipal",
        "modalidade_nome": it.get("modalidade_licitacao_nome") or (
            "Ata de Registro de Preços (SRP)" if is_ata else "Pregão Eletrônico (SRP)"
        ),
        "permite_adesao": bool(permite_adesao),
        "tipo_item_irp": it.get("tipo_item_irp") or "M",
        "prazo_manifestacao": prazo_raw,
        "prazo_manifestacao_formatado": prazo_br,
        "dias_restantes": dias_rest,
        "status_badge": status_badge,
        "data_publicacao": _format_datetime_br(it.get("data_publicacao_pncp")),
        "link_pncp": link_pncp,
        "cnpj_ano_seq": f"{cnpj}_{ano}_{detail_seq}",
    }


def _query_pncp_search_api(params: Dict[str, Any]) -> Dict[str, Any]:
    """Executa chamada resiliente ao motor de busca /api/search/ do PNCP."""
    url = "https://pncp.gov.br/api/search/"
    resp = _HTTP_SESSION.get(url, params=params, timeout=10)
    if resp.status_code != 200:
        raise RuntimeError(f"PNCP Search HTTP {resp.status_code}")
    data = resp.json()
    if not isinstance(data, dict):
        raise RuntimeError("Resposta inesperada da API Search do PNCP")
    return data


def _fallback_consulta_v1_by_uf(
    q_clean: str,
    uf_clean: str,
    pagina_num: int,
    tam_pagina_num: int,
) -> Dict[str, Any]:
    """
    Fallback oficial via API Consulta v1 do PNCP (/api/consulta/v1/contratacoes/proposta)
    quando /api/search/ sofre bloqueio ou indisponibilidade. Filtra nativamente por UF.
    """
    url = "https://pncp.gov.br/api/consulta/v1/contratacoes/proposta"
    params: Dict[str, Any] = {
        "dataFinal": "20271231",
        "codigoModalidadeContratacao": "6",  # Pregão Eletrônico
        "pagina": str(pagina_num),
        "tamanhoPagina": str(min(50, max(15, tam_pagina_num))),
    }
    if uf_clean and uf_clean in VALID_UFS:
        params["uf"] = uf_clean

    resp = _HTTP_SESSION.get(url, params=params, timeout=12)
    if resp.status_code != 200:
        raise RuntimeError(f"PNCP Consulta v1 HTTP {resp.status_code}")

    payload = resp.json()
    raw_list = payload.get("data") or []
    total_registros = int(payload.get("totalRegistros") or len(raw_list))

    q_lower = q_clean.lower()
    formatted: List[Dict[str, Any]] = []

    for row in raw_list:
        unidade = row.get("unidadeOrgao") or {}
        orgao = row.get("orgaoEntidade") or {}
        row_uf = str(unidade.get("ufSigla") or "").upper()
        if uf_clean and row_uf != uf_clean:
            continue

        obj = str(row.get("objetoCompra") or "Sem descrição informada.")
        if q_lower and q_lower not in obj.lower() and q_lower not in str(orgao.get("razaoSocial") or "").lower():
            continue

        cnpj = str(orgao.get("cnpj") or "").strip()
        ano = str(row.get("anoCompra") or "").strip()
        seq = str(row.get("sequencialCompra") or "").strip()
        num_compra = str(row.get("numeroCompra") or seq).strip()
        prazo_raw = row.get("dataEncerramentoProposta")
        dias_rest = _calculate_days_remaining(prazo_raw)
        prazo_br = _format_datetime_br(prazo_raw)
        is_srp = bool(row.get("srp", True))

        link_pncp = (
            f"https://pncp.gov.br/app/editais/{cnpj}/{ano}/{seq}"
            if (cnpj and ano and seq)
            else "https://pncp.gov.br/app/editais"
        )

        formatted.append({
            "id": row.get("numeroControlePNCP") or f"edital_{cnpj}_{ano}_{seq}",
            "document_type": "edital",
            "is_edital": True,
            "is_ata": False,
            "title": f"{row.get('tipoInstrumentoConvocatorioNome') or 'Edital'} nº {num_compra}/{ano}" + (" (SRP)" if is_srp else ""),
            "description": obj,
            "numero_irp": seq,
            "numero_exibicao": num_compra,
            "numero_sequencial_compra_ata": "",
            "ano_irp": ano,
            "numero_controle_pncp": row.get("numeroControlePNCP") or "",
            "orgao_cnpj": cnpj,
            "orgao_nome": orgao.get("razaoSocial") or "Órgão Não Informado",
            "unidade_nome": unidade.get("nomeUnidade") or "",
            "uf": row_uf,
            "municipio_nome": unidade.get("municipioNome") or "",
            "esfera_nome": "Federal" if str(orgao.get("esferaId")) == "F" else ("Estadual" if str(orgao.get("esferaId")) == "E" else "Municipal"),
            "modalidade_nome": (row.get("modalidadeNome") or "Pregão - Eletrônico") + (" (SRP)" if is_srp else ""),
            "permite_adesao": is_srp,
            "tipo_item_irp": "M",
            "prazo_manifestacao": prazo_raw,
            "prazo_manifestacao_formatado": prazo_br,
            "dias_restantes": dias_rest,
            "status_badge": _build_status_badge(dias_rest, "edital"),
            "data_publicacao": _format_datetime_br(row.get("dataPublicacaoPncp")),
            "link_pncp": link_pncp,
            "cnpj_ano_seq": f"{cnpj}_{ano}_{seq}",
        })

    return {
        "total": total_registros if not q_lower else len(formatted),
        "items": formatted[:tam_pagina_num],
    }


def search_irps(
    q: str = "",
    status: str = "aberta",
    uf: str = "",
    pagina: int = 1,
    tam_pagina: int = 20,
) -> Dict[str, Any]:
    """
    Pesquisa Intenções de Registro de Preços (IRP), Editais SRP e Atas de Registro de Preços (ARP)
    no PNCP com filtragem rigorosa por Unidade da Federação (UF).

    status suportados:
      - 'aberta': IRPs Abertas + Oportunidades SRP/ARP vigentes na UF selecionada
      - 'em_andamento': IRPs Em Andamento / Divulgadas
      - 'todas': Todas as IRPs e Oportunidades SRP na UF
      - 'srp_edital': Editais e Pregões SRP recebendo propostas
      - 'ata_vigente': Atas de Registro de Preços (ARP) vigentes para Adesão (Carona)
    """
    q_clean = (q or "").strip()
    status_clean = (status or "aberta").strip().lower()
    uf_clean = (uf or "").strip().upper()
    if uf_clean and uf_clean not in VALID_UFS:
        uf_clean = ""

    pagina_num = max(1, int(pagina))
    tam_pagina_num = min(50, max(10, int(tam_pagina)))

    cache_key = f"{q_clean.lower()}|{status_clean}|{uf_clean}|{pagina_num}|{tam_pagina_num}"
    cached = _IRP_SEARCH_CACHE.get(cache_key)
    if cached and (time.time() - cached["timestamp"] < CACHE_TTL_SECONDS):
        return cached["data"]

    is_srp_edital = status_clean in ["srp_edital", "edital_srp", "edital"]
    is_ata_vigente = status_clean in ["ata_vigente", "ata", "arp"]

    formatted_items: List[Dict[str, Any]] = []
    total_found = 0
    source_mode = status_clean

    try:
        # 1. MODO ATA DE REGISTRO DE PREÇOS VIGENTE (ADESÃO / CARONA)
        if is_ata_vigente:
            params: Dict[str, Any] = {
                "tipos_documento": "ata",
                "ordenacao": "-data",
                "pagina": pagina_num,
                "tam_pagina": tam_pagina_num,
                "status": "vigente",
            }
            if q_clean:
                params["q"] = q_clean
            if uf_clean:
                params["ufs"] = uf_clean

            raw_data = _query_pncp_search_api(params)
            items_raw = raw_data.get("items") or []
            if uf_clean:
                items_raw = [it for it in items_raw if str(it.get("uf") or "").upper() == uf_clean]
            total_found = int(raw_data.get("total") or len(items_raw))
            formatted_items = [_format_search_item(it, "ata") for it in items_raw]

        # 2. MODO EDITAIS E PREGÕES SRP (RECEBENDO PROPOSTAS)
        elif is_srp_edital:
            params = {
                "tipos_documento": "edital",
                "ordenacao": "-data",
                "pagina": pagina_num,
                "tam_pagina": tam_pagina_num,
                "status": "recebendo_proposta",
            }
            if q_clean:
                params["q"] = q_clean
            if uf_clean:
                params["ufs"] = uf_clean

            raw_data = _query_pncp_search_api(params)
            items_raw = raw_data.get("items") or []
            if uf_clean:
                items_raw = [it for it in items_raw if str(it.get("uf") or "").upper() == uf_clean]
            total_found = int(raw_data.get("total") or len(items_raw))
            formatted_items = [_format_search_item(it, "edital") for it in items_raw]

        # 3. MODO IRP (COM EXPANSÃO AUTOMÁTICA PARA EDITAIS SRP & ATAS DA UF QUANDO UF TEM 0 IRPS)
        else:
            irp_status_param = "todos" if status_clean in ["todas", "todos", ""] else "aberta"
            params = {
                "tipos_documento": "irp",
                "ordenacao": "-data",
                "pagina": pagina_num,
                "tam_pagina": tam_pagina_num,
                "status": irp_status_param,
            }
            if q_clean:
                params["q"] = q_clean
            if uf_clean:
                # IMPORTANTE: a API /api/search/ do PNCP exige 'ufs' (no plural) para filtrar por estado!
                params["ufs"] = uf_clean

            raw_data = _query_pncp_search_api(params)
            items_raw = raw_data.get("items") or []
            # Blindagem estrita: jamais permitir itens de outra UF quando uf_clean foi selecionado
            if uf_clean:
                items_raw = [it for it in items_raw if str(it.get("uf") or "").upper() == uf_clean]

            total_found = int(raw_data.get("total") or len(items_raw)) if items_raw else 0
            formatted_items = [_format_search_item(it, "irp") for it in items_raw]

            # Se o usuário filtrou por uma UF específica (ou termo) e o módulo IRP puro do PNCP
            # possui 0 registros ativos naquela UF, complementamos automaticamente com Editais SRP
            # abertos e Atas de Registro de Preços vigentes daquela mesma UF!
            if len(formatted_items) == 0:
                source_mode = "hybrid_srp_uf"
                edital_params: Dict[str, Any] = {
                    "tipos_documento": "edital",
                    "ordenacao": "-data",
                    "pagina": pagina_num,
                    "tam_pagina": tam_pagina_num,
                    "status": "recebendo_proposta",
                }
                if q_clean:
                    edital_params["q"] = q_clean
                if uf_clean:
                    edital_params["ufs"] = uf_clean

                edital_data = _query_pncp_search_api(edital_params)
                edital_items = edital_data.get("items") or []
                if uf_clean:
                    edital_items = [it for it in edital_items if str(it.get("uf") or "").upper() == uf_clean]

                formatted_editais = [_format_search_item(it, "edital") for it in edital_items]

                # Busca também Atas SRP vigentes na UF para adesão (Carona)
                ata_params: Dict[str, Any] = {
                    "tipos_documento": "ata",
                    "ordenacao": "-data",
                    "pagina": pagina_num,
                    "tam_pagina": max(6, tam_pagina_num // 2),
                    "status": "vigente",
                }
                if q_clean:
                    ata_params["q"] = q_clean
                if uf_clean:
                    ata_params["ufs"] = uf_clean

                try:
                    ata_data = _query_pncp_search_api(ata_params)
                    ata_items = ata_data.get("items") or []
                    if uf_clean:
                        ata_items = [it for it in ata_items if str(it.get("uf") or "").upper() == uf_clean]
                    formatted_atas = [_format_search_item(it, "ata") for it in ata_items]
                except Exception:
                    ata_data = {}
                    formatted_atas = []

                # Intercala Editais abertos e Atas vigentes da UF selecionada
                combined: List[Dict[str, Any]] = []
                max_len = max(len(formatted_editais), len(formatted_atas))
                for i in range(max_len):
                    if i < len(formatted_editais):
                        combined.append(formatted_editais[i])
                    if i < len(formatted_atas):
                        combined.append(formatted_atas[i])

                formatted_items = combined[:tam_pagina_num]
                total_found = int(edital_data.get("total") or 0) + int(ata_data.get("total") or 0)

        out = {
            "status": "SUCCESS",
            "total": total_found,
            "page": pagina_num,
            "page_size": tam_pagina_num,
            "uf_filter": uf_clean,
            "source_mode": source_mode,
            "is_edital_mode": is_srp_edital,
            "items": formatted_items,
        }
        _IRP_SEARCH_CACHE[cache_key] = {"timestamp": time.time(), "data": out}
        return out

    except Exception as primary_exc:
        # Fallback automático para API Consulta v1 (/api/consulta/v1/contratacoes/proposta?uf=...)
        try:
            fb = _fallback_consulta_v1_by_uf(q_clean, uf_clean, pagina_num, tam_pagina_num)
            out = {
                "status": "SUCCESS",
                "total": fb["total"],
                "page": pagina_num,
                "page_size": tam_pagina_num,
                "uf_filter": uf_clean,
                "source_mode": "fallback_consulta_v1",
                "is_edital_mode": True,
                "items": fb["items"],
            }
            _IRP_SEARCH_CACHE[cache_key] = {"timestamp": time.time(), "data": out}
            return out
        except Exception as fb_exc:
            return {
                "status": "ERROR",
                "message": f"Falha na comunicação com o PNCP ({primary_exc} / {fb_exc})",
                "total": 0,
                "items": [],
            }


def get_irp_detail(
    cnpj: str,
    ano: str,
    sequencial: str,
    doc_type: str = "auto",
) -> Dict[str, Any]:
    """
    Recupera os detalhes completos da IRP, Edital SRP ou Ata SRP no PNCP:
    cabeçalho, lista de itens homologados/estimados e documentos anexos.
    Suporta resolução automática entre /irp/{ano}/{seq} e /compras/{ano}/{seq}.
    """
    cnpj_clean = "".join(ch for ch in str(cnpj or "") if ch.isdigit())
    ano_clean = "".join(ch for ch in str(ano or "") if ch.isdigit())
    seq_raw = "".join(ch for ch in str(sequencial or "") if ch.isdigit())
    seq_clean = str(int(seq_raw)) if seq_raw else ""
    doc_type_clean = (doc_type or "auto").strip().lower()

    if not (cnpj_clean and ano_clean and seq_clean):
        return {
            "status": "ERROR",
            "message": "Parâmetros CNPJ, Ano e Sequencial inválidos.",
            "itens": [],
            "arquivos": [],
        }

    cache_key = f"{cnpj_clean}_{ano_clean}_{seq_clean}_{doc_type_clean}"
    cached = _IRP_DETAIL_CACHE.get(cache_key)
    if cached and (time.time() - cached["timestamp"] < CACHE_TTL_SECONDS):
        return cached["data"]

    try:
        header_data: Dict[str, Any] = {}
        items_data: List[Dict[str, Any]] = []
        files_data: List[Dict[str, Any]] = []
        resolved_kind = "irp"

        # Ordem de tentativa inteligente conforme tipo do documento
        try_irp_first = doc_type_clean in ["irp", "auto"]

        if try_irp_first:
            irp_api_url = f"https://pncp.gov.br/api/pncp/v1/orgaos/{cnpj_clean}/irp/{ano_clean}/{seq_clean}"
            r_hdr = _HTTP_SESSION.get(irp_api_url, timeout=8)
            if r_hdr.status_code == 200:
                try:
                    header_data = r_hdr.json() or {}
                except Exception:
                    header_data = {}

            if header_data:
                r_it = _HTTP_SESSION.get(f"{irp_api_url}/itens", timeout=8)
                if r_it.status_code == 200:
                    try:
                        items_data = r_it.json() or []
                    except Exception:
                        items_data = []
                r_fl = _HTTP_SESSION.get(f"{irp_api_url}/arquivos", timeout=8)
                if r_fl.status_code == 200:
                    try:
                        files_data = r_fl.json() or []
                    except Exception:
                        files_data = []

        # Se não era IRP (ou se é Edital / Ata SRP), consulta os endpoints de Compras do PNCP
        if not header_data:
            resolved_kind = "edital" if doc_type_clean != "ata" else "ata"
            compra_consulta_url = f"https://pncp.gov.br/api/consulta/v1/orgaos/{cnpj_clean}/compras/{ano_clean}/{seq_clean}"
            compra_api_url = f"https://pncp.gov.br/api/pncp/v1/orgaos/{cnpj_clean}/compras/{ano_clean}/{seq_clean}"

            r_hdr = _HTTP_SESSION.get(compra_consulta_url, timeout=10)
            if r_hdr.status_code == 200:
                try:
                    header_data = r_hdr.json() or {}
                except Exception:
                    header_data = {}

            r_it = _HTTP_SESSION.get(f"{compra_api_url}/itens?pagina=1&tamanhoPagina=100", timeout=10)
            if r_it.status_code == 200:
                try:
                    items_data = r_it.json() or []
                except Exception:
                    items_data = []

            r_fl = _HTTP_SESSION.get(f"{compra_api_url}/arquivos", timeout=10)
            if r_fl.status_code == 200:
                try:
                    files_data = r_fl.json() or []
                except Exception:
                    files_data = []

        prazo = (
            header_data.get("dataHoraFinalManifestacaoInteresse")
            or header_data.get("dataEncerramentoProposta")
            or header_data.get("dataAberturaProposta")
        )
        dias_rest = _calculate_days_remaining(prazo)
        prazo_br = _format_datetime_br(prazo)

        cleaned_items: List[Dict[str, Any]] = []
        for it in items_data:
            if not isinstance(it, dict):
                continue
            mat_ou_serv = str(it.get("materialOuServico") or "M").upper()
            cat_nome = (
                it.get("categoriaItemNome")
                or it.get("itemCategoriaNome")
                or it.get("materialOuServicoNome")
                or ("Serviço" if mat_ou_serv == "S" else "Material")
            )
            unid = (
                it.get("unidadeFornecimento")
                or it.get("unidadeMedida")
                or "UN"
            )
            cod_cat = (
                it.get("codigoItemCatalogo")
                or it.get("catalogoCodigoItem")
                or ""
            )
            cleaned_items.append({
                "numero_item": it.get("numeroItem"),
                "categoria": cat_nome,
                "tipo_codigo": "SERVICO" if mat_ou_serv == "S" else "MATERIAL",
                "descricao": it.get("descricao") or "Item sem descrição informada",
                "quantidade": float(it.get("quantidade") or 1.0),
                "unidade_fornecimento": unid,
                "criterio_julgamento": it.get("criterioJulgamentoNome") or "Menor Preço",
                "valor_unitario_estimado": float(it.get("valorUnitarioEstimado") or 0.0),
                "valor_total": float(it.get("valorTotal") or 0.0),
                "indicador_sigiloso": bool(it.get("indicadorValorSigiloso") or it.get("orcamentoSigiloso") or False),
                "codigo_item_catalogo": str(cod_cat),
            })

        cleaned_files: List[Dict[str, Any]] = []
        resource_path = "irp" if resolved_kind == "irp" else "compras"
        for f in files_data:
            if not isinstance(f, dict):
                continue
            seq_doc = f.get("sequencialDocumento") or 1
            cleaned_files.append({
                "sequencial": seq_doc,
                "titulo": f.get("titulo") or f"Documento_{seq_doc}",
                "tipo": f.get("tipoDocumentoNome") or "Anexo / Edital / Termo de Referência",
                "url": (
                    f.get("url")
                    or f.get("uri")
                    or f.get("urlDownload")
                    or f"https://pncp.gov.br/api/pncp/v1/orgaos/{cnpj_clean}/{resource_path}/{ano_clean}/{seq_clean}/arquivos/{seq_doc}"
                ),
            })

        orgao_ent = header_data.get("orgaoEntidade") or {}
        unid_org = header_data.get("unidadeOrgao") or {}
        amparo_obj = header_data.get("amparoLegal") or {}
        amparo_nome = (
            header_data.get("amparoLegalNome")
            or (amparo_obj.get("nome") if isinstance(amparo_obj, dict) else "")
            or "Lei nº 14.133/2021"
        )

        link_oficial = (
            f"https://pncp.gov.br/app/irp/{cnpj_clean}/{ano_clean}/{seq_clean}"
            if resolved_kind == "irp"
            else f"https://pncp.gov.br/app/editais/{cnpj_clean}/{ano_clean}/{seq_clean}"
        )

        out = {
            "status": "SUCCESS",
            "resolved_kind": resolved_kind,
            "cnpj": cnpj_clean,
            "ano": ano_clean,
            "sequencial": seq_clean,
            "numero_controle_pncp": header_data.get("numeroControlePNCP") or "",
            "processo": header_data.get("processo") or "",
            "objeto": header_data.get("objetoIrp") or header_data.get("objetoCompra") or "",
            "orgao_nome": orgao_ent.get("razaoSocial") or "Órgão Não Informado",
            "unidade_nome": unid_org.get("nomeUnidade") or "",
            "uf": unid_org.get("ufSigla") or "",
            "municipio": unid_org.get("municipioNome") or "",
            "modalidade": header_data.get("modalidadeNome") or "Pregão Eletrônico (SRP)",
            "amparo_legal": amparo_nome,
            "valor_total_estimado": float(header_data.get("valorTotalEstimado") or 0.0),
            "valor_total_homologado": float(header_data.get("valorTotalHomologado") or 0.0),
            "prazo_manifestacao": prazo,
            "prazo_manifestacao_formatado": prazo_br,
            "dias_restantes": dias_rest,
            "total_itens": len(cleaned_items),
            "itens": cleaned_items,
            "total_arquivos": len(cleaned_files),
            "arquivos": cleaned_files,
            "link_oficial_pncp": link_oficial,
            "link_plataforma_origem": header_data.get("urlPlataformaOrigem") or header_data.get("linkSistemaOrigem") or "",
        }
        _IRP_DETAIL_CACHE[cache_key] = {"timestamp": time.time(), "data": out}
        return out
    except Exception as exc:
        return {
            "status": "ERROR",
            "message": f"Erro ao recuperar detalhes no PNCP: {str(exc)}",
            "itens": [],
            "arquivos": [],
        }
