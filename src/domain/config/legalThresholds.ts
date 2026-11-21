/**
 * FONTE ÚNICA DA VERDADE — LIMITES LEGAIS DE DISPENSA DE LICITAÇÃO (LEI Nº 14.133/2021)
 * Art. 75, incisos I e II c/c Art. 182 (atualização anual pelo IPCA-E via Decreto Federal).
 *
 * Corrige definitivamente a inconsistência histórica onde R$ 62.725,59 era exibido
 * ao lado do Decreto nº 11.871/2023 (que vigia apenas em 2024 com valor de R$ 59.906,02).
 */

export interface FiscalYearThreshold {
  readonly fiscalYear: number;
  readonly decreeNumber: string;
  readonly decreeDate: string;
  readonly art75IncisoI: number;   // Obras, serviços de engenharia ou manutenção de veículos automotores
  readonly art75IncisoII: number;  // Outros serviços e compras
  readonly warningRatio: number;   // Alerta preventivo do controle interno (70% do teto)
}

export const LEGAL_THRESHOLDS_BY_YEAR: Readonly<Record<number, FiscalYearThreshold>> = {
  2023: {
    fiscalYear: 2023,
    decreeNumber: "Decreto Federal nº 11.317/2022",
    decreeDate: "29/12/2022",
    art75IncisoI: 114416.65,
    art75IncisoII: 57208.33,
    warningRatio: 0.7,
  },
  2024: {
    fiscalYear: 2024,
    decreeNumber: "Decreto Federal nº 11.871/2023",
    decreeDate: "29/12/2023",
    art75IncisoI: 119812.02,
    art75IncisoII: 59906.02,
    warningRatio: 0.7,
  },
  2025: {
    fiscalYear: 2025,
    decreeNumber: "Decreto Federal nº 12.343/2024",
    decreeDate: "30/12/2024",
    art75IncisoI: 125451.15,
    art75IncisoII: 62725.59,
    warningRatio: 0.7,
  },
  2026: {
    fiscalYear: 2026,
    decreeNumber: "Decreto Federal nº 12.343/2024 (c/c Art. 182 da Lei nº 14.133/2021)",
    decreeDate: "30/12/2024",
    art75IncisoI: 125451.15,
    art75IncisoII: 62725.59,
    warningRatio: 0.7,
  },
};

export const DEFAULT_ACTIVE_FISCAL_YEAR = 2026;

export function getActiveLegalThreshold(fiscalYear: number = DEFAULT_ACTIVE_FISCAL_YEAR): FiscalYearThreshold {
  return LEGAL_THRESHOLDS_BY_YEAR[fiscalYear] ?? LEGAL_THRESHOLDS_BY_YEAR[DEFAULT_ACTIVE_FISCAL_YEAR];
}

export function formatBrlCurrency(value: number): string {
  const safe = Number.isFinite(value) ? value : 0;
  return `R$ ${safe.toLocaleString("pt-BR", { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`;
}

export interface FracionamentoEvaluation {
  readonly isWarning: boolean;
  readonly isExceeded: boolean;
  readonly accumulatedBrl: number;
  readonly limitBrl: number;
  readonly warningLimitBrl: number;
  readonly decreeCitation: string;
  readonly fiscalYear: number;
}

export function evaluateFracionamentoRisk(
  accumulatedValue: number,
  inciso: "I" | "II" = "II",
  fiscalYear: number = DEFAULT_ACTIVE_FISCAL_YEAR
): FracionamentoEvaluation {
  const cfg = getActiveLegalThreshold(fiscalYear);
  const limitBrl = inciso === "I" ? cfg.art75IncisoI : cfg.art75IncisoII;
  const warningLimitBrl = Math.round(limitBrl * cfg.warningRatio * 100) / 100;
  const safeVal = Number.isFinite(accumulatedValue) && accumulatedValue > 0 ? accumulatedValue : 0;
  return {
    isWarning: safeVal > warningLimitBrl,
    isExceeded: safeVal > limitBrl,
    accumulatedBrl: Math.round(safeVal * 100) / 100,
    limitBrl,
    warningLimitBrl,
    decreeCitation: cfg.decreeNumber,
    fiscalYear: cfg.fiscalYear,
  };
}
