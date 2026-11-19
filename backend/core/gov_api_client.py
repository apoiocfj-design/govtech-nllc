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
        Executa GET HTTP com Exponential Backoff + Jitter para suportar oscilações do SERPRO/PNCP,
        registrando logs estruturados (Google Cloud Logging / Observabilidade) de latência e falhas.
        """
        last_error: Optional[Exception] = None
        for attempt in range(1, self.max_retries + 1):
            start_time = time.perf_counter()
            try:
                resp = self.session.get(url, params=params, timeout=self.timeout_seconds)
                latency_ms = int((time.perf_counter() - start_time) * 1000)

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
                        structured_log = {
                            "severity": "WARNING",
                            "service": "compras-gov-api",
                            "latency": latency_ms,
                            "endpoint": url,
                            "status_code": 400,
                            "attempt": attempt,
                            "message": "API Gov retornou 400 transitório ('Erro ao efetuar a consulta')",
                        }
                        logger.warning(json.dumps(structured_log))
                    else:
                        raise ValueError(f"Erro 400 (Parâmetro Inválido na API Gov): {resp.text[:400]}")

                # 429 Too Many Requests ou 5xx Server Error -> Retry com Backoff e Log Estruturado
                if resp.status_code in (429, 500, 502, 503, 504):
                    structured_log = {
                        "severity": "WARNING",
                        "service": "compras-gov-api",
                        "latency": latency_ms,
                        "endpoint": url,
                        "status_code": resp.status_code,
                        "attempt": attempt,
                        "message": f"Instabilidade em gateway governamental: HTTP {resp.status_code}",
                    }
                    logger.warning(json.dumps(structured_log))
                else:
                    resp.raise_for_status()

            except requests.RequestException as exc:
                latency_ms = int((time.perf_counter() - start_time) * 1000)
                last_error = exc
                structured_log = {
                    "severity": "WARNING",
                    "service": "compras-gov-api",
                    "latency": latency_ms,
                    "endpoint": url,
                    "attempt": attempt,
                    "error": str(exc),
                    "message": f"Falha de rede ou timeout ao consultar API governamental: {exc}",
                }
                logger.warning(json.dumps(structured_log))

            if attempt < self.max_retries:
                sleep_time = (self.base_backoff_seconds ** attempt) + random.uniform(0.1, 0.6)
                time.sleep(sleep_time)

        raise GovAPIUnavailableError(
            f"Serviço governamental indisponível após {self.max_retries} tentativas ({url}). Último erro: {last_error}"
        )

    @staticmethod
    def parse_comprasnet_key_and_number(id_compra_item: str, fallback_id_compra: str = "") -> Tuple[str, str]:
        """
        Decodifica a chave oficial de 22 dígitos do Compras.gov.br (`idCompraItem`) ou 17 dígitos (`idCompra`).
        Estrutura SERPRO: UASG (6) + Modalidade (2) + Número da Compra (5) + Ano (4) + Item (5) = 22 dígitos.
        Se idCompra vier com 16 dígitos (UASG iniciada em zero sem o dígito 0), efetua o padding com '0' à esquerda.
        Retorna: (chave_compra_17_digitos, numero_compra_ano_formatado ex: '00039/2026').
        """
        item_digits = re.sub(r"\D", "", str(id_compra_item or ""))
        compra_digits = re.sub(r"\D", "", str(fallback_id_compra or ""))

        # Se id_compra_item tiver >= 17 dígitos (chave completa do item com 22 ou 17 dígitos)
        if len(item_digits) >= 17:
            chave_17 = item_digits[:17]
            num_compra = chave_17[8:13]
            ano_compra = chave_17[13:17]
            return chave_17, f"{num_compra}/{ano_compra}"

        # Caso contrário, verifica fallback_id_compra (muito frequente em compras de serviço CATSER)
        if len(compra_digits) == 16:
            compra_digits = "0" + compra_digits
        if len(compra_digits) >= 17:
            chave_17 = compra_digits[:17]
            num_compra = chave_17[8:13]
            ano_compra = chave_17[13:17]
            return chave_17, f"{num_compra}/{ano_compra}"

        return item_digits or compra_digits, str(id_compra_item or fallback_id_compra or "")

    @staticmethod
    def build_comprasnet_direct_uri(chave_17: str) -> Optional[str]:
        """URL direta e real da contratação no Compras.gov.br (sem página genérica de busca)."""
        digits = re.sub(r"\D", "", str(chave_17 or ""))
        if len(digits) == 16:
            digits = "0" + digits
        if len(digits) >= 17:
            return f"https://cnetmobile.estaleiro.serpro.gov.br/comprasnet-web/public/compras/acompanhamento-compra?compra={digits[:17]}"
        return None

    @classmethod
    def build_audit_uri(
        cls,
        chave_17: str,
        pncp_control_number: Optional[str] = None,
        buyer_uasg: Optional[str] = None,
        num_ano: Optional[str] = None,
    ) -> str:
        """
        Gera URI direta e funcional para a página real da contratação (NUNCA links genéricos '?q=...'):
        1. Prioridade máxima: Página oficial do Edital/Aviso no PNCP (/app/editais/{cnpj}/{ano}/{seq}).
        2. Fallback direto: Página real de acompanhamento da compra no Compras.gov.br (?compra={chave_17}).
        """
        if pncp_control_number:
            # Padrão PNCP: 00394502000144-1-010123/2026 -> /app/editais/00394502000144/2026/10123
            m = re.match(r"^(\d{14})-\d+-(\d+)/(\d{4})$", str(pncp_control_number).strip())
            if m:
                cnpj, seq_int, ano = m.group(1), int(m.group(2)), m.group(3)
                return f"https://pncp.gov.br/app/editais/{cnpj}/{ano}/{seq_int}"

        direct_compras = cls.build_comprasnet_direct_uri(chave_17)
        if direct_compras:
            return direct_compras

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
                if len(raw_records) >= 120 or current_page >= total_pages or not page_items:
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
            if len(chave_17) < 17 and len(re.sub(r"\D", "", id_compra)) in (16, 17):
                padded_compra = "0" + id_compra if len(re.sub(r"\D", "", id_compra)) == 16 else id_compra
                chave_17, num_ano = self.parse_comprasnet_key_and_number(padded_compra)

            pncp_ctrl = None
            if enrich_pncp_links and chave_17:
                pncp_ctrl = self.try_resolve_pncp_control_number(chave_17)

            buyer_uasg_val = str(rec.get("codigoUasg") or "")
            audit_uri = self.build_audit_uri(chave_17, pncp_ctrl, buyer_uasg=buyer_uasg_val, num_ano=num_ano)
            compras_direct = self.build_comprasnet_direct_uri(chave_17)
            raw_date = str(rec.get("dataCompra") or rec.get("dataResultado") or "")[:10]

            supply_u = str(
                rec.get("siglaUnidadeFornecimento")
                or rec.get("siglaUnidadeMedida")
                or rec.get("nomeUnidadeMedida")
                or ("SV" if not is_material else "UN")
            ).strip().upper()

            cap_val = rec.get("capacidadeUnidadeFornecimento")
            sigla_med = str(rec.get("siglaUnidadeMedida") or rec.get("nomeUnidadeMedida") or "").strip()
            cap_str = None
            if cap_val not in (None, "", 0, "0"):
                cap_str = f"{cap_val} {sigla_med}".strip()

            esfera_raw = str(rec.get("esfera") or "").strip().upper()
            esfera_map = {"F": "Federal", "E": "Estadual", "M": "Municipal", "D": "Distrital"}
            esfera_label = esfera_map.get(esfera_raw, esfera_raw or None)

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
                    codigo_classe=str(rec.get("codigoClasse") or "").strip() or None,
                    nome_classe=str(rec.get("nomeClasse") or "").strip() or None,
                    municipio=str(rec.get("municipio") or "").strip() or None,
                    nome_uasg=str(rec.get("nomeUasg") or "").strip() or None,
                    objeto_compra=str(rec.get("objetoCompra") or "").strip() or None,
                    source_inciso="Inciso I - Compras.gov.br",
                    esfera=esfera_label,
                    capacidade_unidade=cap_str,
                    numero_item_compra=str(rec.get("numeroItemCompra") or "").strip() or None,
                    criterio_julgamento=str(rec.get("criterioJulgamento") or "").strip() or None,
                    is_extended_24m=(months_back > 12),
                    comprasnet_uri=compras_direct,
                )
            )

        return item_description, quotes_out

    def fetch_contract_item_prices(
        self,
        catalog_code: int,
        item_type: str = "MATERIAL",
        months_back: int = 12,
    ) -> Tuple[str, List[RawQuoteInput]]:
        """
        Consulta complementar no Módulo Contratos do Compras.gov.br Dados Abertos (OAS 3.1):
        Endpoint oficial: /modulo-contratos/2_consultarContratosItem
        Fundamento: Art. 5º, Inciso II da IN SEGES/ME nº 65/2021 (contratações similares feitas
        pela Administração Pública).
        """
        url = f"{self.COMPRAS_GOV_BASE_URL}/modulo-contratos/2_consultarContratosItem"
        end_date = date.today()
        start_date = end_date - timedelta(days=30 * months_back)
        params: Dict[str, Any] = {
            "codigoItem": int(catalog_code),
            "dataVigenciaInicialMin": start_date.isoformat(),
            "dataVigenciaInicialMax": end_date.isoformat(),
            "tamanhoPagina": 50,
            "pagina": 1,
        }
        try:
            resp = self.session.get(url, params=params, timeout=10)
            if resp.status_code != 200:
                return "", []
            payload = resp.json()
            rows = payload.get("resultado") or []
            if not rows:
                return "", []

            desc = str(rows[0].get("descricaoComplementarItem") or rows[0].get("descricaoItem") or f"Item Catálogo {catalog_code}").strip()
            out: List[RawQuoteInput] = []
            for rec in rows:
                u_price = float(rec.get("valorUnitarioItem") or rec.get("valorUnitario") or 0.0)
                if u_price <= 0:
                    continue
                uasg = str(rec.get("codigoUnidadeGestora") or rec.get("codigoUasg") or "").strip()
                num_contrato = str(rec.get("numeroContrato") or rec.get("id") or "").strip()
                raw_date = str(rec.get("dataVigenciaInicial") or rec.get("dataAssinatura") or "")[:10]
                supply_u = str(rec.get("unidadeMedida") or rec.get("siglaUnidadeMedida") or ("SV" if item_type.upper() != "MATERIAL" else "UN")).strip().upper()
                ctrl_pncp = str(rec.get("numeroControlePncpContrato") or "").strip() or None
                audit_uri = f"https://pncp.gov.br/app/contratos/{ctrl_pncp}" if ctrl_pncp else f"https://contratos.comprasnet.gov.br/transparencia/contratos?numero={num_contrato}"

                out.append(
                    RawQuoteInput(
                        id_compra_item=str(rec.get("id") or f"CTR-{uasg}-{num_contrato}"),
                        purchase_number_year=f"CTR {num_contrato}" if num_contrato else "Contrato Público",
                        buyer_uasg=uasg,
                        buyer_org_name=str(rec.get("nomeOrgao") or rec.get("nomeUnidadeGestora") or "Órgão Público Federal").strip(),
                        buyer_uf=str(rec.get("ufUnidadeGestora") or rec.get("estado") or "").strip(),
                        supplier_cnpj=str(rec.get("niFornecedor") or rec.get("cnpjFornecedor") or "").strip(),
                        supplier_name=str(rec.get("nomeRazaoSocialFornecedor") or rec.get("nomeFornecedor") or "Fornecedor Contratado").strip(),
                        supply_unit=supply_u,
                        quantity_homologated=float(rec.get("quantidadeItem") or rec.get("quantidade") or 1.0),
                        unit_price=u_price,
                        purchase_date=raw_date,
                        pncp_or_compras_uri=audit_uri,
                        brand=str(rec.get("marca") or "").strip() or None,
                        pncp_control_number=ctrl_pncp,
                        forma_compra="Contrato Administrativo (Inciso II)",
                        codigo_pdm=str(rec.get("codigoPdm") or "").strip() or None,
                        nome_pdm=str(rec.get("nomePdm") or "").strip() or None,
                        codigo_classe=str(rec.get("codigoClasse") or "").strip() or None,
                        nome_classe=str(rec.get("nomeClasse") or "").strip() or None,
                        nome_uasg=str(rec.get("nomeUnidadeGestora") or "").strip() or None,
                        objeto_compra=str(rec.get("objetoContrato") or "").strip() or None,
                        source_inciso="Inciso II - Contratos Públicos (Compras.gov.br)",
                        esfera="Federal",
                        is_extended_24m=(months_back > 12),
                    )
                )
            return desc, out
        except Exception as exc:
            logger.debug(f"Consulta complementar de Contratos (Inciso II) para {catalog_code} indisponível: {exc}")
            return "", []

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
            resp = self.api_client.session.get(url, params=params, timeout=12)
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
                    alt_resp = self.api_client.session.get(alt_url, params=alt_params, timeout=12)
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
        allow_24m_fallback: bool = True,
    ) -> Dict[str, Any]:
        """
        Fluxo ponta a ponta:
        1. Busca cotações homologadas na API Oficial do Compras.gov.br (OAS 3.1) em até 12 meses.
        2. Complementa com Módulo de Contratos (Inciso II) quando a amostra for reduzida (< 3).
        3. Se não houver cotações válidas em 12 meses e allow_24m_fallback=True, amplia automaticamente
           para janela excepcional de até 24 meses (separando em pasta distinta no ZIP).
        4. Aplica o Motor de Saneamento Estatístico (IN 65/2021).
        5. Devolve o relatório auditável completo em JSON pronto para o Frontend / PDF / ZIP.
        """
        cache_key = (
            f"{(item_type or 'MATERIAL').upper()}:{catalog_code}:"
            f"{(target_supply_unit or '').upper()}:{months_back}:{uf or ''}:{uasg or ''}:fb{int(bool(allow_24m_fallback))}"
        )
        if not hasattr(self, "_research_cache"):
            self._research_cache = {}
        cached = self._research_cache.get(cache_key)
        if cached and (time.time() - cached[0]) < 1800:
            import copy
            return copy.deepcopy(cached[1])

        result = self._execute_in65_research_uncached(
            catalog_code,
            target_supply_unit,
            item_type,
            months_back,
            uf,
            uasg,
            allow_24m_fallback=allow_24m_fallback,
        )
        if result.get("status") in ("SUCCESS", "NO_DATA"):
            import copy
            self._research_cache[cache_key] = (time.time(), copy.deepcopy(result))
        return result

    def _fetch_quotes_for_window(
        self,
        catalog_code: int,
        item_type: str,
        months_back: int,
        uf: Optional[str],
        uasg: Optional[str],
    ) -> Tuple[str, List[RawQuoteInput], str]:
        """Executa busca primária (Inciso I) + fallback tipo complementar + complemento Contratos (Inciso II)."""
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
                logger.info(f"Fallback automático Compras.gov.br: código {catalog_code} convergiu com tipo {alt_type} ({months_back}M)")
                description = alt_desc
                raw_quotes = alt_quotes
                item_type = alt_type

        # Complemento automático via Módulo de Contratos (Art. 5º, Inciso II IN 65/2021) se amostra primária < 3
        pos_count = sum(1 for q in raw_quotes if (q.unit_price or 0) > 0)
        if pos_count < 3:
            c_desc, c_quotes = self.api_client.fetch_contract_item_prices(
                catalog_code=catalog_code,
                item_type=item_type,
                months_back=months_back,
            )
            if c_quotes:
                if not raw_quotes and c_desc:
                    description = c_desc
                raw_quotes.extend(c_quotes)

        return description, raw_quotes, item_type

    def _execute_in65_research_uncached(
        self,
        catalog_code: int,
        target_supply_unit: str = "UN",
        item_type: str = "MATERIAL",
        months_back: int = 12,
        uf: Optional[str] = None,
        uasg: Optional[str] = None,
        allow_24m_fallback: bool = True,
    ) -> Dict[str, Any]:
        effective_months = int(months_back or 12)
        used_24m_fallback = effective_months > 12

        description, raw_quotes, item_type = self._fetch_quotes_for_window(
            catalog_code=catalog_code,
            item_type=item_type,
            months_back=effective_months,
            uf=uf,
            uasg=uasg,
        )

        # Fallback automático 12M -> 24M quando o item não possui compras homologadas nos últimos 12 meses
        has_positive = any((q.unit_price or 0) > 0 for q in raw_quotes)
        if not has_positive and allow_24m_fallback and effective_months <= 12:
            logger.info(f"[Fallback 12M -> 24M] Sem cotações em 12 meses para {catalog_code}; ampliando para 24 meses.")
            desc_24, quotes_24, type_24 = self._fetch_quotes_for_window(
                catalog_code=catalog_code,
                item_type=item_type,
                months_back=24,
                uf=uf,
                uasg=uasg,
            )
            if any((q.unit_price or 0) > 0 for q in quotes_24):
                description = desc_24
                raw_quotes = quotes_24
                item_type = type_24
                effective_months = 24
                used_24m_fallback = True

        if not raw_quotes:
            return {
                "status": "NO_DATA",
                "catmat_catser_code": catalog_code,
                "search_window_months": effective_months,
                "message": "Sem compras homologadas nos últimos 12 e 24 meses — pesquisa sem convergência.",
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
                "search_window_months": effective_months,
                "message": str(exc),
            }

        report.search_window_months = effective_months
        report.is_extended_24m = used_24m_fallback
        if used_24m_fallback:
            report.fallback_24m_note = (
                "JANELA EXCEPCIONAL (13 a 24 MESES): Não foram localizadas contratações homologadas nos últimos 12 meses "
                "(Art. 5º, I e II da IN SEGES/ME nº 65/2021). Aplicou-se pesquisa ampliada até 24 meses conforme diretrizes "
                "do Caderno de Logística de Pesquisa de Preços (SEGES/MGI), recomendando-se justificativa nos autos e "
                "eventual atualização monetária (IPCA/IBGE) até a data do orçamento estimado."
            )
            report.legal_methodology_justification = (
                f"{report.legal_methodology_justification} {report.fallback_24m_note}"
            )

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
                        if not report.nome_pdm and row[0]:
                            report.nome_pdm = row[0]
                        if not report.nome_classe and row[1]:
                            report.nome_classe = row[1]
        except Exception as exc:
            logger.debug(f"Não foi possível consultar PDM/Classe local para {catalog_code}: {exc}")

        # Resolução paralela e deduplicada de Id contratação PNCP + URL direta PNCP + URL direta Compras.gov.br
        try:
            import concurrent.futures
            try:
                from pncp_extractor import (
                    build_comprasnet_direct_url,
                    parse_pncp_control_id,
                    resolve_pncp_compra,
                )
            except ImportError:
                from core.pncp_extractor import (
                    build_comprasnet_direct_url,
                    parse_pncp_control_id,
                    resolve_pncp_compra,
                )

            quotes_to_enrich = report.valid_quotes + report.purged_quotes[:25]
            unique_keys = []
            seen_keys = set()
            for q in quotes_to_enrich:
                raw_k = q.purchase_key or (q.id_compra_item[:17] if len(q.id_compra_item) >= 17 else "")
                k_digits = re.sub(r"\D", "", str(raw_k or ""))
                if len(k_digits) == 16:
                    k_digits = "0" + k_digits
                if len(k_digits) >= 17:
                    k17 = k_digits[:17]
                    q.purchase_key = k17
                    if not q.comprasnet_uri:
                        q.comprasnet_uri = build_comprasnet_direct_url(k17)
                    if k17 not in seen_keys and not (q.pncp_control_number and "/app/editais/" in (q.pncp_or_compras_uri or "")):
                        seen_keys.add(k17)
                        unique_keys.append(k17)

            resolved_map = {}
            if unique_keys:
                with concurrent.futures.ThreadPoolExecutor(max_workers=8) as executor:
                    for k17, comp in zip(unique_keys, executor.map(resolve_pncp_compra, unique_keys)):
                        if comp:
                            resolved_map[k17] = comp

            for q in quotes_to_enrich:
                k17 = q.purchase_key or ""
                comp = resolved_map.get(k17)
                if comp:
                    ctrl = str(comp.get("numeroControlePNCP") or "").strip()
                    parsed = parse_pncp_control_id(ctrl)
                    if parsed:
                        cnpj, ano, seq_int, seq_6 = parsed
                        q.pncp_control_number = f"{cnpj}-1-{seq_6}/{ano}"
                        q.pncp_or_compras_uri = f"https://pncp.gov.br/app/editais/{cnpj}/{ano}/{seq_int}"
                    if not q.objeto_compra and comp.get("objetoCompra"):
                        q.objeto_compra = str(comp.get("objetoCompra")).strip()
                # Garantia absoluta: NUNCA deixar link genérico '?q='
                if "?q=" in (q.pncp_or_compras_uri or "") and q.comprasnet_uri:
                    q.pncp_or_compras_uri = q.comprasnet_uri
        except Exception as exc_enrich:
            logger.debug(f"Não foi possível enriquecer links literais do PNCP: {exc_enrich}")

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

        # Enriquecimento de links PNCP literais no recálculo personalizado
        try:
            import concurrent.futures
            try:
                from pncp_extractor import resolve_pncp_webpage_url
            except ImportError:
                from core.pncp_extractor import resolve_pncp_webpage_url

            def _enrich_dict_quote(q_dict):
                uri = q_dict.get("pncp_or_compras_uri") or ""
                if "/editais/" in uri and "?q=" not in uri:
                    return
                key = q_dict.get("purchase_key") or (str(q_dict.get("id_compra_item", ""))[:17] if len(str(q_dict.get("id_compra_item", ""))) >= 17 else "")
                if key and len(key) == 16:
                    key = "0" + key
                if key and len(key) >= 17:
                    literal_url = resolve_pncp_webpage_url(key)
                    if literal_url:
                        q_dict["pncp_or_compras_uri"] = literal_url
                        m = re.match(r"https://pncp\.gov\.br/app/editais/(\d+)/(\d+)/(\d+)", literal_url)
                        if m:
                            q_dict["pncp_control_number"] = f"{m.group(1)}-1-{m.group(3)}/{m.group(2)}"

            with concurrent.futures.ThreadPoolExecutor(max_workers=8) as executor:
                list(executor.map(_enrich_dict_quote, rep_dict.get("valid_quotes", []) + rep_dict.get("purged_quotes", [])[:15]))
        except Exception:
            pass

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
