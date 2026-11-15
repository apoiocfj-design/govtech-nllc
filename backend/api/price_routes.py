# -*- coding: utf-8 -*-
"""
ROTAS FASTAPI (CONTROLLER REST): Pesquisa de Preços IN SEGES/ME nº 65/2021
Expõe os endpoints consumidos pelo Frontend Next.js para busca e auditoria de preços.
"""

from typing import Optional
from core.gov_api_client import GovAPIUnavailableError, PriceResearchController

# Caso FastAPI esteja instalado no ambiente de deploy, registra o APIRouter padrão;
# mantém fallback limpo para execução standalone sem dependências externas pesadas.
try:
    from fastapi import APIRouter, HTTPException, Query

    router = APIRouter(prefix="/api/v1/prices", tags=["IN 65/2021 - Pesquisa de Preços"])
    controller = PriceResearchController()

    @router.get("/in65/{catalog_code}")
    async def get_in65_price_research(
        catalog_code: int,
        supply_unit: str = Query("UN", description="Unidade de fornecimento alvo (ex: UN, CX, PCT)"),
        item_type: str = Query("MATERIAL", pattern="^(MATERIAL|SERVICO)$"),
        months_back: int = Query(12, ge=1, le=24, description="Janela temporal em meses (Art. 5º, I da IN 65/2021)"),
        uf: Optional[str] = Query(None, min_length=2, max_length=2, description="Filtro opcional de UF"),
        uasg: Optional[str] = Query(None, description="Filtro opcional de UASG compradora"),
    ):
        """
        Consulta ao vivo as APIs oficiais do Governo Federal (Compras.gov.br OAS 3.1 e PNCP v2.6),
        executa o saneamento estatístico de outliers e devolve a memória de cálculo auditável.
        """
        try:
            result = controller.execute_in65_research(
                catalog_code=catalog_code,
                target_supply_unit=supply_unit,
                item_type=item_type,
                months_back=months_back,
                uf=uf,
                uasg=uasg,
            )
            if result.get("status") == "NO_DATA":
                raise HTTPException(status_code=404, detail=result.get("message"))
            return result
        except ValueError as ve:
            raise HTTPException(status_code=400, detail=str(ve))
        except GovAPIUnavailableError as ge:
            raise HTTPException(status_code=503, detail=str(ge))
except ImportError:
    router = None
