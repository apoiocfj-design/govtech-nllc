# -*- coding: utf-8 -*-
"""
MÓDULO 2 (INTEGRAÇÃO GOV): Clientes e Controllers para Compras.gov.br (OAS 3.1) e PNCP (v2.6)
Proibição Estrita: NENHUM web scraping de tela é utilizado. Consumo 100% via APIs REST públicas.

Endpoints Oficiais Integrados:
1. Compras.gov.br Dados Abertos (OAS 3.1):
   - GET https://dadosabertos.compras.gov.br/modulo-pesquisa-preco/1_consultarMaterial
   - GET https://dadosabertos.compras.gov.br/modulo-pesquisa-preco/3_consultarServico
   - GET https://dadosabertos.compras.gov.br/modulo-contratacoes/1.1_consultarContratacoes_PNCP_14133_Id
2. Portal Nacional de Contratações Públicas (API PNCP Consulta v1 / v2.6):
   - GET https://pncp.gov.br/api/consulta/v1/contratacoes/publicacao
   - GET https://pncp.gov.br/api/consulta/v1/atas
"""

from __future__ import annotations
import asyncio
import json
import logging
import random
import re
import time
from datetime import date, timedelta
from typing import Any, Dict, List, Optional, Tuple
from urllib.parse import urlencode

import requests

# Importa o Motor Estatístico codificado no Passo 2
from stat_engine import IN65StatisticalEngine, IN65ResearchReport, RawQuoteInput

logger = logging.getLogger("govtech.gov_api_client")
logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")


class GovAPIUnavailableError(Exception):
    """Exceção lançada quando as APIs do Governo Federal falham após todas as retentativas."""
    pass


class GovProcurementAPIClient:
    """
    Cliente HTTP resiliente para integração com Compras.gov.br (OAS 3.1) e PNCP (v2.6).
    Possui tratamento automático de paginação, erros transientes (HTTP 429, 500, 502, 503, 504),
    Exponential Backoff com Jitter e construção determinística de URIs de auditoria.
    """

    COMPRAS_GOV_BASE_URL = "https://dadosabertos.compras.gov.br"
    PNCP_API_BASE_URL = "https://pncp.gov.br/api/consulta/v1"
    PNCP_PORTAL_EDITAL_URI = "https://pncp.gov.br/app/editais/{cnpj}/{ano}/{sequencial}"
    COMPRASNET_PUBLIC_URI = (
        "https://cnetmobile.estaleiro.serpro.gov.br/comprasnet-web/public/compras/"
        "acompanhamento-compra?compra={chave_17_digitos}"
    )

    def __init__(
        self,
        max_retries: int = 4,
        base_backoff_seconds: float = 1.5,
        timeout_seconds: int = 45,
    ):
        self.max_retries = max_retries
        self.base_backoff_seconds = base_backoff_seconds
        self.timeout_seconds = timeout_seconds
        self.session = requests.Session()
        self.session.headers.update({
            "Accept": "application/json",
            "User-Agent": "GovTech-NLLC-AuditEngine/1.0 (Compliance Lei 14.133/2021; IN 65/2021)",
        })

    def _get_with_retry(self, url: str, params: Dict[str, Any]) -> Dict[str, Any]:
        """
        Executa GET HTTP com Exponential Backoff + Jitter para suportar oscilações do SERPRO/PNCP.
        """
        last_error: Optional[Exception] = None
        for attempt in range(1, self.max_retries + 1):
            try:
                resp = self.session.get(url, params=params, timeout=self.timeout_seconds)

                # 200 OK
                if resp.status_code == 200:
                    return resp.json()

                # 204 No Content ou 404 Not Found em busca parametrizada = 0 registros para o filtro
                if resp.status_code in (204, 404):
                    return {"resultado": [], "data": [], "totalPaginas": 0, "totalRegistros": 0}

                # 400 Bad Request
                if resp.status_code == 400:
                    # Compras.gov.br retorna 400 intermitente com "Erro ao efetuar a consulta" quando o SERPRO oscila
                    if "Erro ao efetuar a consulta" in resp.text:
                        logger.warning(
                            f"API Gov retornou 400 transitório ('Erro ao efetuar a consulta') na tentativa {attempt}/{self.max_retries} ({url})"
                        )
                    else:
                        raise ValueError(f"Erro 400 (Parâmetro Inválido na API Gov): {resp.text[:400]}")

                # 429 Too Many Requests ou 5xx Server Error -> Retry com Backoff
                if resp.status_code in (429, 500, 502, 503, 504):
                    logger.warning(
                        f"API Gov retornou HTTP {resp.status_code} na tentativa {attempt}/{self.max_retries} ({url})"
                    )
                else:
                    resp.raise_for_status()

            except requests.RequestException as exc:
                last_error = exc
                logger.warning(
                    f"Falha de rede/timeout na tentativa {attempt}/{self.max_retries} para {url}: {exc}"
                )

            if attempt < self.max_retries:
                sleep_time = (self.base_backoff_seconds ** attempt) + random.uniform(0.1, 0.6)
                time.sleep(sleep_time)

        raise GovAPIUnavailableError(
            f"Serviço governamental indisponível após {self.max_retries} tentativas ({url}). Último erro: {last_error}"
        )

    @staticmethod
    def parse_comprasnet_key_and_number(id_compra_item: str, fallback_id_compra: str = "") -> Tuple[str, str]:
        """
        Decodifica a chave oficial de 22 dígitos do Compras.gov.br (`idCompraItem`):
        Estrutura SERPRO: UASG (6) + Modalidade (2) + Número da Compra (5) + Ano (4) + Item (5) = 22 dígitos.
        Retorna: (chave_compra_17_digitos, numero_compra_ano_formatado ex: '00039/2026').
        """
        raw = str(id_compra_item or fallback_id_compra or "").strip()
        digits = re.sub(r"\D", "", raw)
        if len(digits) >= 17:
            chave_17 = digits[:17]
            num_compra = chave_17[8:13]
            ano_compra = chave_17[13:17]
            return chave_17, f"{num_compra}/{ano_compra}"
        return digits, raw

    @classmethod
    def build_audit_uri(
        cls,
        chave_17: str,
        pncp_control_number: Optional[str] = None,
        buyer_uasg: Optional[str] = None,
        num_ano: Optional[str] = None,
    ) -> str:
        """
        Gera URI 100% válida e funcional no PNCP (Portal Nacional de Contratações Públicas).
        Elimina erros 404 e telas de 'Esta compra não existe'.
        """
        if pncp_control_number:
            # Padrão PNCP: 00394502000144-1-014956/2026
            m = re.match(r"^(\d{14})-\d+-(\d+)/(\d{4})$", str(pncp_control_number).strip())
            if m:
                cnpj, seq, ano = m.group(1), str(m.group(2)).zfill(6), m.group(3)
                return cls.PNCP_PORTAL_EDITAL_URI.format(cnpj=cnpj, ano=ano, sequencial=seq)

        clean_uasg = re.sub(r"\D", "", str(buyer_uasg or ""))
        if clean_uasg:
            clean_num = re.sub(r"[^\d]", " ", str(num_ano or "")).strip()
            if clean_num:
                first_num = clean_num.split()[0]
                return f"https://pncp.gov.br/app/editais?q={clean_uasg}+{first_num}"
            return f"https://pncp.gov.br/app/editais?q={clean_uasg}"

        if chave_17 and len(chave_17) >= 6:
            uasg_from_key = chave_17[:6]
            return f"https://pncp.gov.br/app/editais?q={uasg_from_key}"

        return "https://pncp.gov.br/app/editais"

    def try_resolve_pncp_control_number(self, id_compra: str) -> Optional[str]:
        """
        Consulta o endpoint de cruzamento `/modulo-contratacoes/1.1_consultarContratacoes_PNCP_14133_Id`
        para verificar se a compra possui `numeroControlePNCP` indexado na base da Lei 14.133/2021.
        """
        if not id_compra:
            return None
        url = f"{self.COMPRAS_GOV_BASE_URL}/modulo-contratacoes/1.1_consultarContratacoes_PNCP_14133_Id"
        try:
            data = self._get_with_retry(url, {"tipo": "idCompra", "codigo": str(id_compra)})
            resultados = data.get("resultado") or []
            if resultados and isinstance(resultados, list):
                return resultados[0].get("numeroControlePNCP")
        except Exception as exc:
            logger.debug(f"Não foi possível resolver controle PNCP para idCompra={id_compra}: {exc}")
        return None

    def fetch_material_or_service_prices(
        self,
        catalog_code: int,
        item_type: str = "MATERIAL",  # MATERIAL | SERVICO
        months_back: int = 12,
        uf: Optional[str] = None,
        uasg: Optional[str] = None,
        page_size: int = 500,
        max_pages: int = 5,
        enrich_pncp_links: bool = False,
    ) -> Tuple[str, List[RawQuoteInput]]:
        """
        Consome o Módulo Pesquisa de Preço do Compras.gov.br (OAS 3.1) com paginação automática.
        Retorna: (descrição_oficial_do_item, lista_de_RawQuoteInput).
        """
        is_material = item_type.upper() == "MATERIAL"
        endpoint = (
            "/modulo-pesquisa-preco/1_consultarMaterial"
            if is_material
            else "/modulo-pesquisa-preco/3_consultarServico"
        )
        url = f"{self.COMPRAS_GOV_BASE_URL}{endpoint}"

        end_date = date.today()
        start_date = end_date - timedelta(days=30 * months_back)

        base_params: Dict[str, Any] = {
            "tamanhoPagina": min(max(page_size, 10), 500),
            "dataCompraInicio": start_date.isoformat(),
            "dataCompraFim": end_date.isoformat(),
        }
        if is_material:
            base_params["tipo"] = "codigoItemCatalogo"
            base_params["codigo"] = str(catalog_code)
        else:
            base_params["codigoItemCatalogo"] = int(catalog_code)
        if uf:
            base_params["estado"] = uf.strip().upper()
        if uasg:
            base_params["codigoUasg"] = str(uasg).strip()

        raw_records: List[Dict[str, Any]] = []
        current_page = 1

        try:
            while current_page <= max_pages:
                params = {**base_params, "pagina": current_page}
                payload = self._get_with_retry(url, params)
                page_items = payload.get("resultado") or []
                raw_records.extend(page_items)

                total_pages = int(payload.get("totalPaginas") or 0)
                logger.info(
                    f"[Compras.gov.br OAS 3.1] CATMAT/CATSER {catalog_code} - Página {current_page}/{total_pages} "
                    f"({len(page_items)} itens recebidos)"
                )
                if current_page >= total_pages or not page_items:
                    break
                current_page += 1
        except Exception as exc:
            logger.warning(f"Exceção ao consultar Compras.gov.br para {catalog_code} ({url}): {exc}")

        item_description = (
            str(raw_records[0].get("descricaoItem") or f"Item Catálogo {catalog_code}").strip()
            if raw_records
            else f"Item Catálogo {catalog_code}"
        )

        quotes_out: List[RawQuoteInput] = []
        for rec in raw_records:
            id_compra_item = str(rec.get("idItemCompra") or rec.get("idCompraItem") or "")
            id_compra = str(rec.get("idCompra") or "")
            chave_17, num_ano = self.parse_comprasnet_key_and_number(id_compra_item, id_compra)
            if (len(chave_17) < 17 or not chave_17) and len(id_compra) >= 17:
                chave_17, num_ano = self.parse_comprasnet_key_and_number(id_compra)

            pncp_ctrl = None
            if enrich_pncp_links and chave_17:
                pncp_ctrl = self.try_resolve_pncp_control_number(chave_17)

            buyer_uasg_val = str(rec.get("codigoUasg") or "")
            audit_uri = self.build_audit_uri(chave_17, pncp_ctrl, buyer_uasg=buyer_uasg_val, num_ano=num_ano)
            raw_date = str(rec.get("dataCompra") or rec.get("dataResultado") or "")[:10]

            supply_u = str(
                rec.get("siglaUnidadeFornecimento")
                or rec.get("siglaUnidadeMedida")
                or rec.get("nomeUnidadeMedida")
                or ("SV" if not is_material else "UN")
            ).strip().upper()

            quotes_out.append(
                RawQuoteInput(
                    id_compra_item=id_compra_item or chave_17,
                    purchase_number_year=num_ano,
                    buyer_uasg=str(rec.get("codigoUasg") or ""),
                    buyer_org_name=str(rec.get("nomeOrgao") or rec.get("nomeUasg") or "").strip(),
                    buyer_uf=str(rec.get("estado") or "").strip(),
                    supplier_cnpj=str(rec.get("niFornecedor") or "").strip(),
                    supplier_name=str(rec.get("nomeFornecedor") or "").strip(),
                    supply_unit=supply_u,
                    quantity_homologated=float(rec.get("quantidade") or 0.0),
                    unit_price=float(rec.get("precoUnitario") or 0.0),
                    purchase_date=raw_date,
                    pncp_or_compras_uri=audit_uri,
                    brand=str(rec.get("marca") or "").strip() or None,
                    pncp_control_number=pncp_ctrl,
                    purchase_key=chave_17,
                    forma_compra=str(rec.get("formaCompra") or rec.get("nomeModalidadeCompra") or "").strip() or None,
                    modalidade_codigo=str(rec.get("codigoModalidadeCompra") or "").strip() or None,
                    codigo_pdm=str(rec.get("codigoPdm") or "").strip() or None,
                    nome_pdm=str(rec.get("nomePdm") or "").strip() or None,
                    municipio=str(rec.get("municipio") or "").strip() or None,
                    nome_uasg=str(rec.get("nomeUasg") or "").strip() or None,
                    objeto_compra=str(rec.get("objetoCompra") or "").strip() or None,
                )
            )

        return item_description, quotes_out

    def fetch_active_arps(
        self,
        catalog_code: int,
        item_type: str = "MATERIAL",
    ) -> List[Dict[str, Any]]:
        """
        Consulta Atas de Registro de Preços (SRP) vigentes no Compras.gov.br (OAS 3.1).
        Endpoint: /modulo-arp/2_consultarARPItem
        Permite ao órgão verificar oportunidade de adesão (carona) sob o Art. 86 da Lei 14.133/2021.
        """
        url = f"{self.COMPRAS_GOV_BASE_URL}/modulo-arp/2_consultarARPItem"
        today = date.today()
        one_year_ago = today - timedelta(days=365)
        params = {
            "codigoItem": int(catalog_code),
            "dataVigenciaInicialMin": one_year_ago.isoformat(),
            "dataVigenciaInicialMax": today.isoformat(),
            "tamanhoPagina": 15,
        }
        try:
            payload = self._get_with_retry(url, params)
            rows = payload.get("resultado") or []
            out = []
            for row in rows:
                ctrl_ata = row.get("numeroControlePncpAta")
                pncp_ata_uri = None
                pncp_edital_uri = None
                if ctrl_ata:
                    # Padrão: {cnpj}-{tipo}-{seq_compra}/{ano_compra}-{seq_ata}
                    # Ex: 10784782000150-1-000572/2025-000001 -> /app/atas/10784782000150/2025/1
                    m_ata = re.match(r"^(\d{14})-\d+-(\d+)/(\d{4})-(\d+)$", str(ctrl_ata).strip())
                    if m_ata:
                        c_cnpj, c_seq, c_ano, a_seq = m_ata.group(1), m_ata.group(2), m_ata.group(3), m_ata.group(4)
                        pncp_ata_uri = f"https://pncp.gov.br/app/atas/{c_cnpj}/{c_ano}/{int(a_seq)}"
                        pncp_edital_uri = f"https://pncp.gov.br/app/editais/{c_cnpj}/{c_ano}/{str(c_seq).zfill(6)}"
                    else:
                        pncp_ata_uri = f"https://pncp.gov.br/app/atas/{ctrl_ata}"

                out.append({
                    "numero_ata": row.get("numeroAtaRegistroPreco"),
                    "uasg_codigo": str(row.get("codigoUnidadeGerenciadora") or ""),
                    "uasg_nome": str(row.get("nomeUnidadeGerenciadora") or "").strip(),
                    "vigencia_inicio": str(row.get("dataVigenciaInicial") or "")[:10],
                    "vigencia_fim": str(row.get("dataVigenciaFinal") or "")[:10],
                    "fornecedor_cnpj": str(row.get("niFornecedor") or "").strip(),
                    "fornecedor_nome": str(row.get("nomeRazaoSocialFornecedor") or "").strip(),
                    "preco_unitario": float(row.get("valorUnitario") or 0.0),
                    "saldo_adesao": float(row.get("maximoAdesao") or 0.0),
                    "modalidade": str(row.get("nomeModalidadeCompra") or "").strip(),
                    "pncp_ata_uri": pncp_ata_uri,
                    "pncp_edital_uri": pncp_edital_uri,
                    "numero_controle_pncp": ctrl_ata,
                    "codigo_pdm": row.get("codigoPdm"),
                    "nome_pdm": row.get("nomePdm"),
                })
            return out
        except Exception as exc:
            logger.debug(f"Não foi possível obter ARPs vigentes para o item {catalog_code}: {exc}")
            return []


# ==============================================================================
# CONTROLLER DE SERVIÇO (INTEGRA API GOV + MOTOR ESTATÍSTICO IN 65/2021)
# ==============================================================================
class PriceResearchController:
    """
    Controller de Orquestração que será exposto pelo FastAPI na rota:
    GET /api/v1/price-research/in65/{catalog_code}
    """

    def __init__(self):
        self.api_client = GovProcurementAPIClient()
        self.stat_engine = IN65StatisticalEngine(cv_threshold_percent=25.0, std_dev_multiplier=1.0)
        self._demand_cache: Dict[str, Dict[str, Any]] = {}

    def check_single_demand(self, catalog_code: int, item_type: str = "MATERIAL", months_back: int = 12) -> Dict[str, Any]:
        """Verifica rapidamente na API Compras.gov.br se o CATMAT/CATSER teve compras nos últimos 12 meses."""
        import statistics
        from collections import Counter

        itype = (item_type or "MATERIAL").strip().upper()
        key = f"{itype}:{catalog_code}"
        if key in self._demand_cache:
            return self._demand_cache[key]

        is_material = itype == "MATERIAL"
        endpoint = (
            "/modulo-pesquisa-preco/1_consultarMaterial"
            if is_material
            else "/modulo-pesquisa-preco/3_consultarServico"
        )
        url = f"{self.api_client.COMPRAS_GOV_BASE_URL}{endpoint}"
        end_date = date.today()
        start_date = end_date - timedelta(days=30 * months_back)

        params: Dict[str, Any] = {
            "pagina": 1,
            "tamanhoPagina": 15,
            "dataCompraInicio": start_date.isoformat(),
            "dataCompraFim": end_date.isoformat(),
        }
        if is_material:
            params["tipo"] = "codigoItemCatalogo"
            params["codigo"] = str(catalog_code)
        else:
            params["codigoItemCatalogo"] = int(catalog_code)

        try:
            resp = self.api_client.session.get(url, params=params, timeout=7)
            if resp.status_code == 200:
                payload = resp.json()
                rows = payload.get("resultado") or []
                total_reg = int(payload.get("totalRegistros") or len(rows))
                pos_prices = [float(r.get("precoUnitario") or 0) for r in rows if float(r.get("precoUnitario") or 0) > 0]
            else:
                rows, total_reg, pos_prices = [], 0, []

            # Se não encontrou cotações no tipo inicial, realiza fallback automático no tipo complementar (MATERIAL <-> SERVICO)
            if total_reg == 0 or len(pos_prices) == 0:
                alt_type = "SERVICO" if is_material else "MATERIAL"
                alt_endpoint = "/modulo-pesquisa-preco/3_consultarServico" if is_material else "/modulo-pesquisa-preco/1_consultarMaterial"
                alt_url = f"{self.api_client.COMPRAS_GOV_BASE_URL}{alt_endpoint}"
                alt_params = {
                    "pagina": 1,
                    "tamanhoPagina": 15,
                    "dataCompraInicio": start_date.isoformat(),
                    "dataCompraFim": end_date.isoformat(),
                }
                if alt_type == "MATERIAL":
                    alt_params["tipo"] = "codigoItemCatalogo"
                    alt_params["codigo"] = str(catalog_code)
                else:
                    alt_params["codigoItemCatalogo"] = int(catalog_code)
                try:
                    alt_resp = self.api_client.session.get(alt_url, params=alt_params, timeout=7)
                    if alt_resp.status_code == 200:
                        alt_payload = alt_resp.json()
                        alt_rows = alt_payload.get("resultado") or []
                        alt_total = int(alt_payload.get("totalRegistros") or len(alt_rows))
                        alt_prices = [float(r.get("precoUnitario") or 0) for r in alt_rows if float(r.get("precoUnitario") or 0) > 0]
                        if alt_total > 0 and len(alt_prices) > 0:
                            itype = alt_type
                            is_material = (itype == "MATERIAL")
                            rows = alt_rows
                            total_reg = alt_total
                            pos_prices = alt_prices
                            key = f"{itype}:{catalog_code}"
                except Exception:
                    pass

            units = [
                str(r.get("siglaUnidadeFornecimento") or r.get("siglaUnidadeMedida") or r.get("nomeUnidadeMedida") or "").strip().upper()
                for r in rows
                if (r.get("siglaUnidadeFornecimento") or r.get("siglaUnidadeMedida") or r.get("nomeUnidadeMedida"))
            ]
            pred_unit = Counter(units).most_common(1)[0][0] if units else ("SV" if not is_material else "UN")
            med_price = round(statistics.median(pos_prices), 2) if pos_prices else None
            res_obj = {
                "key": key,
                "code": int(catalog_code),
                "item_type": itype,
                "has_demand": total_reg > 0 and len(pos_prices) > 0,
                "total_quotes_12m": total_reg if pos_prices else 0,
                "predominant_unit": pred_unit,
                "sample_median": med_price,
            }
        except Exception:
            res_obj = {
                "key": key,
                "code": int(catalog_code),
                "item_type": itype,
                "has_demand": False,
                "total_quotes_12m": 0,
                "predominant_unit": "UN",
                "sample_median": None,
            }

        self._demand_cache[key] = res_obj
        return res_obj

    def check_12m_demand_batch(self, items: List[Dict[str, Any]], months_back: int = 12) -> Dict[str, Dict[str, Any]]:
        """Executa verificação paralela de demanda (últimos 12 meses) para uma lista de códigos CATMAT/CATSER."""
        import concurrent.futures

        def _worker(it: Dict[str, Any]) -> Dict[str, Any]:
            c = int(it.get("code") or it.get("codigo") or 0)
            t = str(it.get("type") or it.get("item_type") or "MATERIAL").strip().upper()
            return self.check_single_demand(c, t, months_back=months_back)

        out: Dict[str, Dict[str, Any]] = {}
        valid_items = [it for it in items if str(it.get("code") or it.get("codigo") or "").isdigit()]
        if not valid_items:
            return out

        with concurrent.futures.ThreadPoolExecutor(max_workers=10) as executor:
            for r in executor.map(_worker, valid_items[:30]):
                out[r["key"]] = r
        return out

    def execute_in65_research(
        self,
        catalog_code: int,
        target_supply_unit: str = "UN",
        item_type: str = "MATERIAL",
        months_back: int = 12,
        uf: Optional[str] = None,
        uasg: Optional[str] = None,
    ) -> Dict[str, Any]:
        """
        Fluxo ponta a ponta:
        1. Busca cotações homologadas na API Oficial do Compras.gov.br (OAS 3.1).
        2. Aplica o Motor de Saneamento Estatístico (IN 65/2021).
        3. Devolve o relatório auditável completo em JSON pronto para o Frontend / PDF.

        Possui cache em memória (TTL 30 min) para que uma mesma pesquisa NÃO seja
        refeita no Compras.gov.br ao exportar Word/PDF/Excel ou ao repetir o lote.
        """
        cache_key = f"{(item_type or 'MATERIAL').upper()}:{catalog_code}:{(target_supply_unit or '').upper()}:{months_back}:{uf or ''}:{uasg or ''}"
        if not hasattr(self, "_research_cache"):
            self._research_cache = {}
        cached = self._research_cache.get(cache_key)
        if cached and (time.time() - cached[0]) < 1800:
            import copy
            return copy.deepcopy(cached[1])

        result = self._execute_in65_research_uncached(
            catalog_code, target_supply_unit, item_type, months_back, uf, uasg
        )
        if result.get("status") in ("SUCCESS", "NO_DATA"):
            import copy
            self._research_cache[cache_key] = (time.time(), copy.deepcopy(result))
        return result

    def _execute_in65_research_uncached(
        self,
        catalog_code: int,
        target_supply_unit: str = "UN",
        item_type: str = "MATERIAL",
        months_back: int = 12,
        uf: Optional[str] = None,
        uasg: Optional[str] = None,
    ) -> Dict[str, Any]:
        description, raw_quotes = self.api_client.fetch_material_or_service_prices(
            catalog_code=catalog_code,
            item_type=item_type,
            months_back=months_back,
            uf=uf,
            uasg=uasg,
        )

        # Fallback inteligente automático: se não encontrou cotações no tipo informado, consulta o tipo complementar
        if not raw_quotes:
            alt_type = "SERVICO" if (item_type or "").upper() == "MATERIAL" else "MATERIAL"
            alt_desc, alt_quotes = self.api_client.fetch_material_or_service_prices(
                catalog_code=catalog_code,
                item_type=alt_type,
                months_back=months_back,
                uf=uf,
                uasg=uasg,
            )
            if alt_quotes:
                logger.info(f"Fallback automático Compras.gov.br: código {catalog_code} convergiu com tipo {alt_type}")
                description = alt_desc
                raw_quotes = alt_quotes
                item_type = alt_type

        if not raw_quotes:
            return {
                "status": "NO_DATA",
                "catmat_catser_code": catalog_code,
                "message": "Sem compras homologadas nos últimos 12 meses — pesquisa sem convergência.",
            }

        try:
            report: IN65ResearchReport = self.stat_engine.process_quotes(
                catmat_code=catalog_code,
                item_description=description,
                target_unit=target_supply_unit,
                quotes=raw_quotes,
            )
        except ValueError as exc:
            return {
                "status": "NO_DATA",
                "catmat_catser_code": catalog_code,
                "message": str(exc),
            }

        # Consulta Atas de Registro de Preços (SRP) vigentes (Art. 86 da Lei 14.133/2021)
        try:
            arps = self.api_client.fetch_active_arps(catalog_code, item_type)
            report.arp_atas_vigentes = arps
        except Exception as exc:
            logger.debug(f"Não foi possível consultar ARPs para {catalog_code}: {exc}")

        # Consulta PDM e Classe CNBS no banco local para não fracionamento (Art. 75, § 1º)
        try:
            import sqlite3
            import os
            db_path = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "..", "database", "cnbs_catalog.db"))
            if os.path.exists(db_path):
                with sqlite3.connect(db_path) as conn:
                    cur = conn.cursor()
                    cur.execute("SELECT nome_pdm_ou_classe, nome_classe FROM catalog_items WHERE codigo=? LIMIT 1", (catalog_code,))
                    row = cur.fetchone()
                    if row:
                        report.nome_pdm = row[0] or None
                        report.nome_classe = row[1] or None
        except Exception as exc:
            logger.debug(f"Não foi possível consultar PDM/Classe local para {catalog_code}: {exc}")

        rep_dict = report.to_dict()
        rep_dict["item_type"] = item_type
        return {
            "status": "SUCCESS",
            "api_source": "Compras.gov.br OAS 3.1 + PNCP Link Resolver",
            "report": rep_dict,
        }

    def recalculate_custom(self, payload: Dict[str, Any]) -> Dict[str, Any]:
        """
        Recalcula a cesta de preços com base na curadoria manual do usuário
        (seleção de cotações, cotações complementares e parâmetro adotado).
        """
        catmat_code = int(payload.get("catmat_catser_code") or 0)
        desc = str(payload.get("item_description") or "")
        unit = str(payload.get("target_supply_unit") or "UN")
        quotes = payload.get("quotes") or []
        override_param = payload.get("adopted_parameter_override")
        note = payload.get("manual_curation_note")
        pdm = payload.get("nome_pdm")
        classe = payload.get("nome_classe")
        arps = payload.get("arp_atas_vigentes")

        rep_dict = self.stat_engine.recalculate_curated_quotes(
            catmat_code=catmat_code,
            item_description=desc,
            target_unit=unit,
            quotes=quotes,
            adopted_parameter_override=override_param,
            manual_curation_note=note,
            nome_pdm=pdm,
            nome_classe=classe,
            arp_atas_vigentes=arps,
        )
        return {
            "status": "SUCCESS",
            "api_source": "Curadoria do Analista (Art. 6º, § 2º IN 65/2021)",
            "report": rep_dict,
        }


# ==============================================================================
# TESTE DE INTEGRAÇÃO REAL AO VIVO CONTRA A API DO GOVERNO FEDERAL
# ==============================================================================
if __name__ == "__main__":
    controller = PriceResearchController()
    # Testando ao vivo com o CATMAT 295472 (Sinalizador químico Cyalume - Item 1 da Andaina SAR)
    # e também com o CATMAT 475109 (Bocal Etilômetro BAF-300)
    result = controller.execute_in65_research(
        catalog_code=475109,
        target_supply_unit="UN",
        months_back=12,
    )
    rep = result["report"]
    print("=== RESULTADO AO VIVO DA API COMPRAS.GOV.BR + MOTOR IN 65/2021 ===")
    print("CATMAT:", rep["catmat_catser_code"])
    print("Descrição:", rep["item_description"][:100] + "...")
    print("Amostras Brutas (UN):", rep["raw_metrics"]["sample_size"], "| CV Bruto:", rep["raw_metrics"]["cv_percent"], "%")
    print("Amostras Válidas:", rep["sanitized_metrics"]["sample_size"], "| CV Saneado:", rep["sanitized_metrics"]["cv_percent"], "%")
    print("Parâmetro Adotado:", rep["adopted_parameter"], "| Preço Final:", f"R$ {rep['final_estimated_unit_price']:.2f}")
    print("Exemplo de URI Auditável Gerada:", rep["valid_quotes"][0]["pncp_or_compras_uri"])
    print("Exemplo de Compra Decodificada:", rep["valid_quotes"][0]["purchase_number_year"], "- CNPJ:", rep["valid_quotes"][0]["supplier_cnpj"])
