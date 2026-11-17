# -*- coding: utf-8 -*-
"""
MÓDULO DE INTEROPERABILIDADE BIDIRECIONAL COM O PNCP (PORTAL NACIONAL DE CONTRATAÇÕES PÚBLICAS)
Gera e valida o payload JSON exato exigido pelo PNCP (Lei nº 14.133/2021) para inclusão automática
de Avisos de Contratação Direta (Dispensa de Licitação - Art. 75) e Editais de Pregão Eletrônico.
"""

from __future__ import annotations

import json
from datetime import datetime, timezone
from typing import Any, Dict, List, Optional


def build_pncp_procurement_payload(
    items_list: List[Dict[str, Any]],
    om_name: str = "MARINHA DO BRASIL",
    uasg_code: str = "785000",
    cnpj_orgao: str = "00394502000144",
    uf: str = "BA",
    municipio: str = "Juazeiro",
    processo_nup: str = "",
    numero_compra: str = "00001/2026",
    ano_compra: int = 2026,
    modalidade_id: int = 8,  # 8 = Dispensa de Licitação, 6 = Pregão Eletrônico
    modo_disputa_id: int = 1,  # 1 = Aberto, 2 = Fechado
    tipo_instrumento_id: int = 1,  # 1 = Aviso de Contratação Direta, 2 = Edital
    objeto_descricao: str = "",
    hash_auditoria: str = "",
) -> Dict[str, Any]:
    """
    Constrói o payload JSON em estrita conformidade com a API REST do PNCP.
    """
    now_iso = datetime.now(timezone.utc).isoformat()
    total_homologavel = sum(float(it.get("total_price") or it.get("valor_total") or (float(it.get("unit_price") or it.get("valor_unitario") or it.get("report", {}).get("final_estimated_unit_price") or 0.0) * int(it.get("qty") or it.get("quantidade") or 1))) for it in items_list)

    if not objeto_descricao:
        if len(items_list) == 1:
            desc = items_list[0].get("report", {}).get("item_description") or "Aquisição de item"
            objeto_descricao = f"Contratação referente a: {desc}."
        else:
            objeto_descricao = f"Aquisição de lote contendo {len(items_list)} itens para suprimento operacional da Unidade Administrativa."

    # Mapeamento dos itens do lote
    pncp_itens: List[Dict[str, Any]] = []
    for idx, item in enumerate(items_list, start=1):
        rep = item.get("report") or {}
        sm = rep.get("sanitized_metrics") or {}
        code = item.get("code") or item.get("codigo") or rep.get("catmat_catser_code") or 0
        item_type = str(item.get("type") or item.get("tipo") or rep.get("item_type") or "MATERIAL").upper()
        tipo_pncp = "S" if "SERV" in item_type else "M"
        unit_price = float(item.get("unit_price") or item.get("valor_unitario") or rep.get("final_estimated_unit_price") or 0.0)
        qty = int(item.get("qty") or item.get("quantidade") or 1)
        tot_price = float(item.get("total_price") or item.get("valor_total") or (unit_price * qty))
        valid_quotes = rep.get("valid_quotes") or []
        cv_val = float(sm.get("cv_percent") or 0.0)

        param_adotado = rep.get("adopted_parameter") or ("MEDIANA" if cv_val > 25.0 else "MEDIA")

        pncp_itens.append({
            "numeroItem": idx,
            "tipoItem": tipo_pncp,
            "codigoItem": int(code),
            "descricao": item.get("descricao") or item.get("desc") or rep.get("item_description") or f"Item {code}",
            "quantidade": qty,
            "unidadeMedida": item.get("unidade") or item.get("unit") or rep.get("target_supply_unit") or "UN",
            "valorUnitarioEstimado": round(unit_price, 2),
            "valorTotal": round(tot_price, 2),
            "criterioJulgamentoId": 1,  # 1 = Menor Preço
            "situacaoCompraItem": 1,   # 1 = Em andamento / Propostas
            "metodologiaPrecoReferencia": {
                "baseLegal": "IN SEGES/ME nº 65/2021",
                "parametroAdotado": param_adotado,
                "coeficienteVariacao": round(cv_val, 2),
                "amostrasValidas": len(valid_quotes),
                "amostrasOutliers": len(rep.get("purged_quotes") or []),
            },
        })

    payload = {
        "codigoModalidadeContratacao": modalidade_id,
        "modoDisputaId": modo_disputa_id,
        "tipoInstrumentoConvocatorioCodigo": tipo_instrumento_id,
        "numeroCompra": numero_compra,
        "anoCompra": ano_compra,
        "processo": processo_nup or f"{uasg_code}.{numero_compra.replace('/', '')}/{ano_compra}",
        "tipoObjetoId": 1 if any(i["tipoItem"] == "M" for i in pncp_itens) else 2,
        "objetoCompra": objeto_descricao,
        "informacaoComplementar": (
            "Pesquisa de Preços realizada estritamente sob as regras da Lei nº 14.133/2021 "
            "e da Instrução Normativa SEGES/ME nº 65/2021, processada e auditada no Sistema GovTech NLLC."
        ),
        "srp": False,
        "valorTotalEstimado": round(total_homologavel, 2),
        "orgaoEntidade": {
            "cnpj": cnpj_orgao,
            "codigoUnidadeCompradora": uasg_code,
            "nomeUnidadeCompradora": om_name,
            "ufSigla": uf,
            "municipioNome": municipio,
        },
        "itens": pncp_itens,
        "metadadosAuditoria": {
            "sistemaOrigem": "Sistema GovTech NLLC (Lei 14.133/2021)",
            "versaoModulo": "2.4.0-Enterprise",
            "hashIntegridade": hash_auditoria or "N/A",
            "dataGeracaoUtc": now_iso,
        },
    }

    return payload


def validate_pncp_payload(payload: Dict[str, Any]) -> Dict[str, Any]:
    """
    Valida a integridade do payload antes do envio à API oficial do PNCP.
    """
    errors = []
    if not payload.get("numeroCompra"):
        errors.append("Campo 'numeroCompra' é obrigatório.")
    if not payload.get("objetoCompra"):
        errors.append("Campo 'objetoCompra' é obrigatório.")
    if not payload.get("itens"):
        errors.append("A contratação deve conter pelo menos 1 item.")
    for it in payload.get("itens", []):
        if float(it.get("valorUnitarioEstimado", 0)) <= 0:
            errors.append(f"Item #{it.get('numeroItem')} possui valor unitário zero ou negativo.")
        if int(it.get("quantidade", 0)) <= 0:
            errors.append(f"Item #{it.get('numeroItem')} possui quantidade zero ou negativa.")

    return {
        "valid": len(errors) == 0,
        "errors": errors,
        "total_itens": len(payload.get("itens", [])),
        "valor_total": payload.get("valorTotalEstimado", 0.0),
    }
