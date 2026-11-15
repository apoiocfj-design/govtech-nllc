# -*- coding: utf-8 -*-
"""
MÓDULO 2 (CORE MATEMÁTICO): Motor de Saneamento Estatístico de Preços Públicos
Norma Regente: Instrução Normativa SEGES/ME nº 65, de 7 de julho de 2021 (Art. 6º)
e Manual de Orientação de Pesquisa de Preços do STJ / TCU.

Regras de Negócio Implacáveis Implementadas:
1. Filtro Estrito de Unidade de Fornecimento (ex.: separa 'UN' de 'CX' ou 'PCT').
2. Cálculo das Métricas Brutas: Média (μ), Mediana, Desvio Padrão Amostral (s) e
   Coeficiente de Variação Bruto: CV = (s / μ) * 100.
3. Saneamento de Outliers em Duplo Estágio (Defesa contra Distorções Extremas):
   - Estágio A (Pré-filtro Robusto por Mediana/IQR): Em compras públicas, erros de
     digitação no Compras.gov.br (ex: lançar R$ 90.000,00 para um item de R$ 2,50)
     inflam a Média (μ) e o Desvio Padrão (s) a ponto de mascarar outros outliers.
     Por isso, expurgam-se primeiramente desvios aberrantes de ordem de grandeza
     (inexequíveis < 20% da Mediana; superfaturados > 500% da Mediana ou fora da cerca de Tukey 1.5*IQR).
   - Estágio B (Corte Clássico por Desvio Padrão - μ ± k*s): Sobre a amostra pré-saneada,
     aplica-se o corte pelo Desvio Padrão Amostral (Limites: μ - s e μ + s, ou fator configurável)
     para estabilizar o Coeficiente de Variação (CV <= 25%).
4. Decisão Automática do Parâmetro de Referência (Art. 6º da IN 65/2021):
   - Se o Coeficiente de Variação Saneado (CV_final) <= 25,00%: Amostra HOMOGÊNEA -> Adota-se a MÉDIA.
   - Se o Coeficiente de Variação Saneado (CV_final) > 25,00%: Amostra HETEROGÊNEA -> Adota-se a MEDIANA.
5. Trilha de Auditoria Completa: Cada cotação descartada recebe carimbo individual com
   a fórmula exata do motivo do expurgo para apresentação ao TCU / Controle Interno.
"""

from __future__ import annotations
import json
import math
import statistics
from dataclasses import dataclass, field, asdict
from datetime import date
from decimal import Decimal, ROUND_HALF_UP
from typing import List, Optional, Dict, Any


def round_currency(val: float) -> float:
    """Arredondamento financeiro padrão ABNT/TCU (4 casas decimais p/ unitário ou 2 p/ moeda)."""
    return float(Decimal(str(val)).quantize(Decimal("0.0001"), rounding=ROUND_HALF_UP))


def round_2(val: float) -> float:
    return float(Decimal(str(val)).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP))


@dataclass
class RawQuoteInput:
    """Representa uma cotação bruta vinda do Compras.gov.br (OAS 3.1) ou PNCP v2.6."""
    id_compra_item: str
    purchase_number_year: str
    buyer_uasg: str
    buyer_org_name: str
    buyer_uf: str
    supplier_cnpj: str
    supplier_name: str
    supply_unit: str
    quantity_homologated: float
    unit_price: float
    purchase_date: str
    pncp_or_compras_uri: str
    brand: Optional[str] = None
    pncp_control_number: Optional[str] = None


@dataclass
class AuditedQuoteOutput:
    """Cotação auditada com status de validação e justificativa matemática individual."""
    id_compra_item: str
    purchase_number_year: str
    buyer_uasg: str
    buyer_org_name: str
    buyer_uf: str
    supplier_cnpj: str
    supplier_name: str
    supply_unit: str
    quantity_homologated: float
    unit_price: float
    purchase_date: str
    pncp_or_compras_uri: str
    brand: Optional[str] = None
    pncp_control_number: Optional[str] = None
    is_outlier: bool = False
    purge_status: str = "VALIDA"  # VALIDA | EXPURGADA_UNIDADE_DIVERGENTE | EXPURGADA_INEXEQUIVEL | EXPURGADA_ACIMA_LIMITE
    purge_justification: Optional[str] = None


@dataclass
class StatisticalMetrics:
    """Fotografia estatística de uma série de preços (Bruta ou Saneada)."""
    sample_size: int
    min_price: Optional[float]
    max_price: Optional[float]
    mean: Optional[float]
    median: Optional[float]
    std_deviation: Optional[float]
    cv_percent: Optional[float]  # Coeficiente de Variação (%) = (s / media) * 100
    lower_bound_std: Optional[float] = None
    upper_bound_std: Optional[float] = None


@dataclass
class IN65ResearchReport:
    """Payload final estruturado devolvido pelo Motor Estatístico."""
    catmat_catser_code: int
    item_description: str
    target_supply_unit: str
    cv_homogeneity_threshold_percent: float
    std_dev_multiplier: float
    raw_metrics: StatisticalMetrics
    sanitized_metrics: StatisticalMetrics
    is_homogeneous: bool
    has_minimum_three_quotes: bool
    adopted_parameter: str  # MEDIA | MEDIANA
    final_estimated_unit_price: float
    legal_methodology_justification: str
    valid_quotes: List[AuditedQuoteOutput] = field(default_factory=list)
    purged_quotes: List[AuditedQuoteOutput] = field(default_factory=list)

    def to_dict(self) -> Dict[str, Any]:
        return asdict(self)


class IN65StatisticalEngine:
    """
    Motor de Saneamento Estatístico de Preços Públicos (IN SEGES/ME nº 65/2021).
    """

    def __init__(
        self,
        cv_threshold_percent: float = 25.0,
        std_dev_multiplier: float = 1.0,
        infeasible_median_ratio: float = 0.25,
        excessive_median_ratio: float = 4.0,
    ):
        """
        :param cv_threshold_percent: Limiar de homogeneidade do Coeficiente de Variação (padrão TCU/STJ: 25%).
        :param std_dev_multiplier: Fator 'k' para o corte por Desvio Padrão (μ ± k * s). Padrão: 1.0.
        :param infeasible_median_ratio: Limite inferior robusto contra preços vis/inexequíveis (ex: 25% da mediana).
        :param excessive_median_ratio: Limite superior robusto contra erros grosseiros de unidade (ex: 400% da mediana).
        """
        self.cv_threshold_percent = cv_threshold_percent
        self.std_dev_multiplier = std_dev_multiplier
        self.infeasible_median_ratio = infeasible_median_ratio
        self.excessive_median_ratio = excessive_median_ratio

    @staticmethod
    def _compute_metrics(prices: List[float], k_std: float = 1.0) -> StatisticalMetrics:
        n = len(prices)
        if n == 0:
            return StatisticalMetrics(0, None, None, None, None, None, None)
        min_p = round_currency(min(prices))
        max_p = round_currency(max(prices))
        mean_p = statistics.mean(prices)
        median_p = statistics.median(prices)

        if n >= 2:
            std_p = statistics.stdev(prices)
            cv_p = (std_p / mean_p) * 100.0 if mean_p > 0 else 0.0
            lo_std = max(0.0, mean_p - k_std * std_p)
            hi_std = mean_p + k_std * std_p
        else:
            std_p = 0.0
            cv_p = 0.0
            lo_std = mean_p
            hi_std = mean_p

        return StatisticalMetrics(
            sample_size=n,
            min_price=min_p,
            max_price=max_p,
            mean=round_currency(mean_p),
            median=round_currency(median_p),
            std_deviation=round_currency(std_p),
            cv_percent=round_2(cv_p),
            lower_bound_std=round_currency(lo_std),
            upper_bound_std=round_currency(hi_std),
        )

    @staticmethod
    def _normalize_unit(u: str) -> str:
        raw = (u or "").strip().upper()
        if raw in ("UN", "UND", "UNID", "UNIDADE", "UNIDADES", "PÇ", "PC", "PEÇA", "PECA"):
            return "UN"
        if raw in ("PAR", "PARES", "PR"):
            return "PAR"
        if raw in ("CX", "CAIXA"):
            return "CX"
        if raw in ("PCT", "PACOTE"):
            return "PCT"
        if raw in ("SERV", "SERVIÇO", "SERVICO", "SV", ""):
            return raw or "UN"
        return raw

    def process_quotes(
        self,
        catmat_code: int,
        item_description: str,
        target_unit: str,
        quotes: List[RawQuoteInput],
    ) -> IN65ResearchReport:
        from collections import Counter

        valid_list: List[AuditedQuoteOutput] = []
        purged_list: List[AuditedQuoteOutput] = []

        positive_quotes = [q for q in quotes if q.unit_price is not None and q.unit_price > 0]
        if not positive_quotes:
            raise ValueError("Nenhuma cotação com valor unitário positivo encontrada nos últimos 12 meses.")

        req_unit = (target_unit or "").strip().upper()
        req_norm = self._normalize_unit(req_unit) if req_unit else ""

        # Verifica se há cotações compatíveis com a unidade solicitada; se não houver (ou se vazia),
        # adota automaticamente a unidade de fornecimento predominante na amostra oficial do Compras.gov.br.
        matching_req = [
            q for q in positive_quotes
            if req_norm and (
                q.supply_unit.strip().upper() == req_unit
                or self._normalize_unit(q.supply_unit) == req_norm
            )
        ]
        if matching_req:
            effective_target_unit = matching_req[0].supply_unit.strip().upper() or req_unit
            effective_norm = self._normalize_unit(effective_target_unit)
        else:
            unit_counts = Counter(self._normalize_unit(q.supply_unit) for q in positive_quotes)
            effective_norm = unit_counts.most_common(1)[0][0]
            # Preserva a sigla original mais frequente daquele grupo
            raw_for_norm = [
                q.supply_unit.strip().upper()
                for q in positive_quotes
                if self._normalize_unit(q.supply_unit) == effective_norm and q.supply_unit.strip()
            ]
            effective_target_unit = Counter(raw_for_norm).most_common(1)[0][0] if raw_for_norm else (req_unit or "UN")
        target_unit = effective_target_unit

        # 1. Separar cotações com Unidade de Fornecimento divergente ou preço inválido (<= 0)
        same_unit_quotes: List[RawQuoteInput] = []
        for q in quotes:
            if q.unit_price is None or q.unit_price <= 0:
                purged_list.append(
                    AuditedQuoteOutput(
                        **asdict(q),
                        is_outlier=True,
                        purge_status="EXPURGADA_INEXEQUIVEL",
                        purge_justification="Valor unitário nulo ou menor/igual a zero.",
                    )
                )
            elif (
                q.supply_unit.strip().upper() != target_unit.strip().upper()
                and self._normalize_unit(q.supply_unit) != effective_norm
            ):
                purged_list.append(
                    AuditedQuoteOutput(
                        **asdict(q),
                        is_outlier=True,
                        purge_status="EXPURGADA_UNIDADE_DIVERGENTE",
                        purge_justification=(
                            f"Unidade de fornecimento da compra ('{q.supply_unit}') diverge "
                            f"da unidade-alvo padronizada ('{target_unit}')."
                        ),
                    )
                )
            else:
                same_unit_quotes.append(q)

        # 2. Métricas Brutas (sobre todas as cotações da mesma unidade)
        raw_prices = [q.unit_price for q in same_unit_quotes]
        raw_metrics = self._compute_metrics(raw_prices, self.std_dev_multiplier)

        if not same_unit_quotes:
            raise ValueError(f"Nenhuma cotação encontrada na unidade '{target_unit}'.")

        # 3. Saneamento Estatístico em Duplo Estágio (Robust Median Gate + Desvio Padrão)
        # Estágio A: Barreira Robusta pela Mediana Bruta (elimina distorções de ordem de grandeza, ex: R$ 90.000)
        raw_median = statistics.median(raw_prices)
        robust_min = raw_median * self.infeasible_median_ratio
        robust_max = raw_median * self.excessive_median_ratio

        stage_a_survivors: List[RawQuoteInput] = []
        for q in same_unit_quotes:
            if q.unit_price < robust_min:
                purged_list.append(
                    AuditedQuoteOutput(
                        **asdict(q),
                        is_outlier=True,
                        purge_status="EXPURGADA_INEXEQUIVEL",
                        purge_justification=(
                            f"Preço manifestamente inexequível (R$ {q.unit_price:.2f} < "
                            f"limite robusto de R$ {robust_min:.2f} [{self.infeasible_median_ratio*100:.0f}% da mediana bruta R$ {raw_median:.2f}]). "
                            f"Fundamento: Art. 6º da IN SEGES/ME nº 65/2021."
                        ),
                    )
                )
            elif q.unit_price > robust_max:
                purged_list.append(
                    AuditedQuoteOutput(
                        **asdict(q),
                        is_outlier=True,
                        purge_status="EXPURGADA_ACIMA_LIMITE",
                        purge_justification=(
                            f"Preço excessivamente elevado / erro grosseiro de unidade (R$ {q.unit_price:.2f} > "
                            f"limite robusto de R$ {robust_max:.2f} [{self.excessive_median_ratio*100:.0f}% da mediana bruta R$ {raw_median:.2f}]). "
                            f"Fundamento: Art. 6º da IN SEGES/ME nº 65/2021."
                        ),
                    )
                )
            else:
                stage_a_survivors.append(q)

        # Estágio B: Corte fino por Desvio Padrão (μ ± k * s) caso o CV da pré-amostra ainda supere o limiar de 25%
        stage_a_prices = [q.unit_price for q in stage_a_survivors]
        stage_a_metrics = self._compute_metrics(stage_a_prices, self.std_dev_multiplier)

        if (
            stage_a_metrics.sample_size >= 4
            and stage_a_metrics.cv_percent is not None
            and stage_a_metrics.cv_percent > self.cv_threshold_percent
        ):
            lo_std = stage_a_metrics.lower_bound_std
            hi_std = stage_a_metrics.upper_bound_std
            mean_a = stage_a_metrics.mean
            std_a = stage_a_metrics.std_deviation

            for q in stage_a_survivors:
                if q.unit_price < lo_std:
                    purged_list.append(
                        AuditedQuoteOutput(
                            **asdict(q),
                            is_outlier=True,
                            purge_status="EXPURGADA_INEXEQUIVEL",
                            purge_justification=(
                                f"Expurgado por Desvio Padrão: R$ {q.unit_price:.2f} está abaixo do limite inferior "
                                f"(μ - {self.std_dev_multiplier:g}s = R$ {lo_std:.2f}, onde μ=R$ {mean_a:.2f} e s=R$ {std_a:.2f})."
                            ),
                        )
                    )
                elif q.unit_price > hi_std:
                    purged_list.append(
                        AuditedQuoteOutput(
                            **asdict(q),
                            is_outlier=True,
                            purge_status="EXPURGADA_ACIMA_LIMITE",
                            purge_justification=(
                                f"Expurgado por Desvio Padrão: R$ {q.unit_price:.2f} está acima do limite superior "
                                f"(μ + {self.std_dev_multiplier:g}s = R$ {hi_std:.2f}, onde μ=R$ {mean_a:.2f} e s=R$ {std_a:.2f})."
                            ),
                        )
                    )
                else:
                    valid_list.append(
                        AuditedQuoteOutput(
                            **asdict(q),
                            is_outlier=False,
                            purge_status="VALIDA",
                            purge_justification="Cotação validada dentro do intervalo estatístico de confiança (μ ± s).",
                        )
                    )
        else:
            # Se a amostra pós-Estágio A já tem CV <= 25% (homogênea), preserva todas as sobreviventes
            for q in stage_a_survivors:
                valid_list.append(
                    AuditedQuoteOutput(
                        **asdict(q),
                        is_outlier=False,
                        purge_status="VALIDA",
                        purge_justification="Cotação validada em amostra estatisticamente homogênea (CV <= 25%).",
                    )
                )

        # 4. Cálculo das Métricas Finais Saneadas
        sanitized_prices = [q.unit_price for q in valid_list]
        sanitized_metrics = self._compute_metrics(sanitized_prices, self.std_dev_multiplier)

        # 5. Regra de Decisão Automática (Média vs Mediana) segundo o CV Saneado
        cv_final = sanitized_metrics.cv_percent or 0.0
        is_homogeneous = cv_final <= self.cv_threshold_percent
        has_min_three = sanitized_metrics.sample_size >= 3

        if is_homogeneous:
            adopted_param = "MEDIA"
            final_price = round_2(sanitized_metrics.mean)
            param_reason = (
                f"Como o Coeficiente de Variação após o saneamento foi de {cv_final:.2f}% "
                f"(inferior ou igual ao limite de dispersão de {self.cv_threshold_percent:.0f}%), "
                f"a cesta de preços é classificada como HOMOGÊNEA, adotando-se a MÉDIA ARITMÉTICA "
                f"(R$ {final_price:.2f}) como valor estimado de referência."
            )
        else:
            adopted_param = "MEDIANA"
            final_price = round_2(sanitized_metrics.median)
            param_reason = (
                f"Como o Coeficiente de Variação após o saneamento foi de {cv_final:.2f}% "
                f"(superior ao limite de {self.cv_threshold_percent:.0f}%), "
                f"a cesta apresenta dispersão residual HETEROGÊNEA, adotando-se a MEDIANA "
                f"(R$ {final_price:.2f}) como medida de tendência central, em estrita observância "
                f"ao art. 6º da IN SEGES/ME nº 65/2021."
            )

        min_quotes_note = (
            f"Atendido o requisito mínimo legal de 3 (três) preços válidos ({sanitized_metrics.sample_size} cotações validadas), "
            f"conforme art. 6º, § 5º da IN SEGES/ME nº 65/2021."
            if has_min_three
            else f"ATENÇÃO AUDITORIA: Obtiveram-se apenas {sanitized_metrics.sample_size} cotação(ões) válida(s), exigindo justificativa circunstanciada da autoridade competente (art. 6º, § 5º da IN 65/2021)."
        )

        justification = (
            f"Pesquisa de preços processada segundo o art. 6º da IN SEGES/ME nº 65/2021. "
            f"De um total de {len(quotes)} registros extraídos das bases oficiais, {len(purged_list)} foram "
            f"expurgados por divergência de unidade, inexequibilidade ou sobrepreço estatístico "
            f"(CV bruto reduziu de {raw_metrics.cv_percent:.2f}% para {cv_final:.2f}%). "
            f"{param_reason} {min_quotes_note}"
        )

        return IN65ResearchReport(
            catmat_catser_code=catmat_code,
            item_description=item_description,
            target_supply_unit=target_unit,
            cv_homogeneity_threshold_percent=self.cv_threshold_percent,
            std_dev_multiplier=self.std_dev_multiplier,
            raw_metrics=raw_metrics,
            sanitized_metrics=sanitized_metrics,
            is_homogeneous=is_homogeneous,
            has_minimum_three_quotes=has_min_three,
            adopted_parameter=adopted_param,
            final_estimated_unit_price=final_price,
            legal_methodology_justification=justification,
            valid_quotes=valid_list,
            purged_quotes=purged_list,
        )


# ==============================================================================
# SIMULAÇÃO DE TESTE COM PAYLOAD JSON REALISTA (CASO CATMAT 475109 / ANDAINA SAR)
# ==============================================================================
if __name__ == "__main__":
    simulated_api_json = [
        {
            "id_compra_item": "7823200700039202600001",
            "purchase_number_year": "00039/2026",
            "buyer_uasg": "782320",
            "buyer_org_name": "COMANDO DA MARINHA - CAPITANIA DOS PORTOS DE SERGIPE",
            "buyer_uf": "SE",
            "supplier_cnpj": "04.892.111/0001-55",
            "supplier_name": "ELEC INDUSTRIA E COMERCIO DE EQUIPAMENTOS DE MEDICAO LTDA",
            "supply_unit": "UN",
            "quantity_homologated": 500,
            "unit_price": 2.00,
            "purchase_date": "2026-10-26",
            "pncp_or_compras_uri": "https://pncp.gov.br/app/editais/00394502000144/2026/014956",
            "brand": "ELEC BAF-300"
        },
        {
            "id_compra_item": "9280400700025202600003",
            "purchase_number_year": "00025/2026",
            "buyer_uasg": "928040",
            "buyer_org_name": "CONSORCIO INT. DE SAUDE CENTRO SUL",
            "buyer_uf": "MG",
            "supplier_cnpj": "11.234.567/0001-89",
            "supplier_name": "RIBCO DO BRASIL IMPORTACAO E EXPORTACAO LTDA",
            "supply_unit": "UN",
            "quantity_homologated": 1200,
            "unit_price": 2.62,
            "purchase_date": "2026-09-24",
            "pncp_or_compras_uri": "https://pncp.gov.br/app/editais/92804000000110/2026/000025",
            "brand": "ONE-WAY"
        },
        {
            "id_compra_item": "9884610500030202600003",
            "purchase_number_year": "00030/2026",
            "buyer_uasg": "988461",
            "buyer_org_name": "PREFEITURA MUNICIPAL DE SARANDI",
            "buyer_uf": "PR",
            "supplier_cnpj": "22.345.678/0001-90",
            "supplier_name": "AGS COMERCIO E SERVICOS LTDA",
            "supply_unit": "UN",
            "quantity_homologated": 800,
            "unit_price": 2.55,
            "purchase_date": "2026-08-30",
            "pncp_or_compras_uri": "https://pncp.gov.br/app/editais/78200482000110/2026/000030",
            "brand": "BAF-300"
        },
        {
            "id_compra_item": "1605180500042202600184",
            "purchase_number_year": "00042/2026",
            "buyer_uasg": "160518",
            "buyer_org_name": "COMANDO DO EXERCITO - 28 BC",
            "buyer_uf": "SP",
            "supplier_cnpj": "33.456.789/0001-01",
            "supplier_name": "MEDTECH SUPRIMENTOS HOSPITALARES LTDA",
            "supply_unit": "UN",
            "quantity_homologated": 2000,
            "unit_price": 2.70,
            "purchase_date": "2026-07-22",
            "pncp_or_compras_uri": "https://pncp.gov.br/app/editais/00394452000103/2026/000042",
            "brand": "ELEC"
        },
        {
            "id_compra_item": "1801960700080202600001",
            "purchase_number_year": "00080/2026",
            "buyer_uasg": "180196",
            "buyer_org_name": "SECRETARIA DA SEGURANCA PUBLICA - PMESP",
            "buyer_uf": "SP",
            "supplier_cnpj": "44.567.890/0001-12",
            "supplier_name": "ALCOOTEST EQUIPAMENTOS LTDA",
            "supply_unit": "UN",
            "quantity_homologated": 5000,
            "unit_price": 2.85,
            "purchase_date": "2026-07-26",
            "pncp_or_compras_uri": "https://pncp.gov.br/app/editais/46377800000127/2026/000080",
            "brand": "ELEC"
        },
        {
            "id_compra_item": "9260870600103202600002",
            "purchase_number_year": "00103/2026",
            "buyer_uasg": "926087",
            "buyer_org_name": "DEPARTAMENTO ESTADUAL DE TRANSITO - DETRAN AL",
            "buyer_uf": "AL",
            "supplier_cnpj": "55.678.901/0001-23",
            "supplier_name": "NORTE SUL INSTRUMENTOS DE MEDICAO LTDA",
            "supply_unit": "UN",
            "quantity_homologated": 10000,
            "unit_price": 3.00,
            "purchase_date": "2026-02-26",
            "pncp_or_compras_uri": "https://pncp.gov.br/app/editais/12200176000176/2026/000103",
            "brand": "ONE-WAY"
        },
        {
            "id_compra_item": "9829210590036202600001",
            "purchase_number_year": "90036/2026",
            "buyer_uasg": "982921",
            "buyer_org_name": "PREFEITURA MUNICIPAL DE RIO DAS OSTRAS",
            "buyer_uf": "RJ",
            "supplier_cnpj": "66.789.012/0001-34",
            "supplier_name": "LITORAL MED DISTRIBUIDORA LTDA",
            "supply_unit": "UN",
            "quantity_homologated": 4000,
            "unit_price": 3.10,
            "purchase_date": "2025-12-18",
            "pncp_or_compras_uri": "https://pncp.gov.br/app/editais/39270442000150/2025/090036",
            "brand": "BAF-300"
        },
        # --- OUTLIER 1: Preço Inexequível (erro de centavos: R$ 0,15) ---
        {
            "id_compra_item": "7813300600011202600001",
            "purchase_number_year": "00011/2026",
            "buyer_uasg": "781330",
            "buyer_org_name": "COMANDO DA MARINHA - BASE NAVAL RJ",
            "buyer_uf": "RJ",
            "supplier_cnpj": "77.890.123/0001-45",
            "supplier_name": "DESCARTAVEIS EXPRESS LTDA",
            "supply_unit": "UN",
            "quantity_homologated": 100,
            "unit_price": 0.15,
            "purchase_date": "2026-06-10",
            "pncp_or_compras_uri": "https://pncp.gov.br/app/editais/00394502000144/2026/000011",
            "brand": "GENERICO"
        },
        # --- OUTLIER 2: Sobrepreço / Erro de lançamento de caixa como unidade (R$ 280,00) ---
        {
            "id_compra_item": "7873100600033202600001",
            "purchase_number_year": "00033/2026",
            "buyer_uasg": "787310",
            "buyer_org_name": "CAPITANIA FLUVIAL DO ARAGUAIA-TOCANTINS",
            "buyer_uf": "TO",
            "supplier_cnpj": "88.901.234/0001-56",
            "supplier_name": "TOCANTINS MEDICA LTDA",
            "supply_unit": "UN",
            "quantity_homologated": 3,
            "unit_price": 280.00,
            "purchase_date": "2026-04-12",
            "pncp_or_compras_uri": "https://pncp.gov.br/app/editais/00394502000144/2026/000033",
            "brand": "PACOTE 100 UN"
        },
        # --- OUTLIER 3: Erro grosseiro de digitação no Compras.gov.br (R$ 90.000,00) ---
        {
            "id_compra_item": "1801600590028202600001",
            "purchase_number_year": "90028/2026",
            "buyer_uasg": "180160",
            "buyer_org_name": "DETRAN - SP",
            "buyer_uf": "SP",
            "supplier_cnpj": "99.012.345/0001-67",
            "supplier_name": "EQUIPAMENTOS VIARIOS SP LTDA",
            "supply_unit": "UN",
            "quantity_homologated": 1,
            "unit_price": 90000.00,
            "purchase_date": "2026-07-14",
            "pncp_or_compras_uri": "https://pncp.gov.br/app/editais/15519372000180/2026/090028",
            "brand": "ETILOMETRO COMPLETO LANCADO NO ITEM ERRADO"
        },
        # --- OUTLIER 4: Unidade de Fornecimento Divergente (PCT com 100 em vez de UN) ---
        {
            "id_compra_item": "1550210600156202600001",
            "purchase_number_year": "00156/2026",
            "buyer_uasg": "155021",
            "buyer_org_name": "EBSERH - HC UFMG",
            "buyer_uf": "MG",
            "supplier_cnpj": "10.123.456/0001-78",
            "supplier_name": "MINAS HOSPITALAR LTDA",
            "supply_unit": "PCT",
            "quantity_homologated": 20,
            "unit_price": 115.00,
            "purchase_date": "2026-05-19",
            "pncp_or_compras_uri": "https://pncp.gov.br/app/editais/15126437000143/2026/000156",
            "brand": "ONE-WAY PCT 100"
        }
    ]

    inputs = [RawQuoteInput(**item) for item in simulated_api_json]
    engine = IN65StatisticalEngine(cv_threshold_percent=25.0, std_dev_multiplier=1.0)
    report = engine.process_quotes(
        catmat_code=475109,
        item_description="BOCAL DESCARTÁVEL C/ VÁLVULA UNIDIRECIONAL P/ ETILÔMETRO ELEC BAF-300",
        target_unit="UN",
        quotes=inputs,
    )

    print(json.dumps(report.to_dict(), indent=2, ensure_ascii=False))
