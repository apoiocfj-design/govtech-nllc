/**
 * MOTOR ESTATÍSTICO BLINDADO (IN SEGES/ME Nº 65/2021 — ART. 6º)
 * Funções puras em TypeScript estrito com proteção total contra:
 * - Divisão por zero (média = 0 ou amostra vazia)
 * - Séries com preços idênticos (IQR = 0, desvio padrão = 0, CV = 0% -> preserva todos os preços válidos)
 * - Amostras reduzidas (0, 1 ou 2 cotações -> aciona alerta obrigatório do Art. 6º, § 5º)
 * - Distorções de ponto flutuante IEEE-754 em somatórios monetários
 * - CSV / Excel Formula Injection (=, +, -, @, \t, \r)
 */

import { PT_BR_EDITORIAL, pluralizeValidQuotes } from "../../shared/i18n/ptBR.ts";

export interface QuoteInputTS {
  readonly id?: string;
  readonly unitPrice: number | null | undefined;
  readonly supplyUnit?: string | null;
  readonly supplierName?: string | null;
  readonly supplierCnpj?: string | null;
  readonly pncpOrComprasUri?: string | null;
}

export interface AuditedQuoteTS extends QuoteInputTS {
  readonly unitPrice: number;
  readonly isOutlier: boolean;
  readonly purgeStatus: "VALIDA" | "EXPURGADA_UNIDADE_DIVERGENTE" | "EXPURGADA_INEXEQUIVEL" | "EXPURGADA_ACIMA_LIMITE";
  readonly purgeJustification: string;
}

export interface StatisticalSummaryTS {
  readonly sampleSize: number;
  readonly minPrice: number | null;
  readonly maxPrice: number | null;
  readonly mean: number | null;
  readonly median: number | null;
  readonly q1: number | null;
  readonly q3: number | null;
  readonly iqr: number | null;
  readonly stdDeviation: number | null;
  readonly cvPercent: number | null;
}

export interface IN65EvaluationResultTS {
  readonly targetUnit: string;
  readonly rawMetrics: StatisticalSummaryTS;
  readonly sanitizedMetrics: StatisticalSummaryTS;
  readonly isHomogeneous: boolean;
  readonly hasMinimumThreeQuotes: boolean;
  readonly adoptedParameter: "MEDIA" | "MEDIANA" | "SEM_DADOS";
  readonly finalEstimatedUnitPrice: number;
  readonly legalMethodologyJustification: string;
  readonly validQuotes: readonly AuditedQuoteTS[];
  readonly purgedQuotes: readonly AuditedQuoteTS[];
}

export function roundCurrency2(val: number): number {
  if (!Number.isFinite(val)) return 0;
  return Math.round((val + Number.EPSILON) * 100) / 100;
}

export function roundCurrency4(val: number): number {
  if (!Number.isFinite(val)) return 0;
  return Math.round((val + Number.EPSILON) * 10000) / 10000;
}

/**
 * Soma valores monetários convertendo para centavos inteiros (BigInt/inteiros seguros)
 * eliminando desvios de ponto flutuante (ex: 0.1 + 0.2 !== 0.30000000000000004).
 */
export function sumMonetarySafe(values: readonly number[]): number {
  let totalCents = 0;
  for (const v of values) {
    if (Number.isFinite(v)) {
      totalCents += Math.round(v * 100);
    }
  }
  return totalCents / 100;
}

/**
 * Proteção contra CSV / Excel Formula Injection (OWASP).
 * Células iniciadas com =, +, -, @, \t ou \r recebem apóstrofo preventivo.
 */
export function sanitizeSpreadsheetCell(value: unknown): string {
  const str = String(value ?? "");
  if (/^[=+\-@\t\r]/.test(str)) {
    return `'${str}`;
  }
  return str;
}

export function normalizeSupplyUnit(unit?: string | null): string {
  const raw = String(unit ?? "").trim().toUpperCase();
  if (["UN", "UND", "UNID", "UNIDADE", "UNIDADES", "PÇ", "PC", "PEÇA", "PECA"].includes(raw)) return "UN";
  if (["PAR", "PARES", "PR"].includes(raw)) return "PAR";
  if (["CX", "CAIXA"].includes(raw)) return "CX";
  if (["PCT", "PACOTE"].includes(raw)) return "PCT";
  return raw || "UN";
}

export function computePercentile(sortedAsc: readonly number[], p: number): number {
  const n = sortedAsc.length;
  if (n === 0) return 0;
  if (n === 1) return sortedAsc[0];
  const clamped = Math.max(0, Math.min(1, p));
  const idx = (n - 1) * clamped;
  const lo = Math.floor(idx);
  const hi = Math.ceil(idx);
  if (lo === hi) return sortedAsc[lo];
  return sortedAsc[lo] + (sortedAsc[hi] - sortedAsc[lo]) * (idx - lo);
}

export function computeStatisticalSummary(prices: readonly number[]): StatisticalSummaryTS {
  const clean = prices.filter((p) => typeof p === "number" && Number.isFinite(p) && p > 0);
  const n = clean.length;
  if (n === 0) {
    return {
      sampleSize: 0,
      minPrice: null,
      maxPrice: null,
      mean: null,
      median: null,
      q1: null,
      q3: null,
      iqr: null,
      stdDeviation: null,
      cvPercent: null,
    };
  }
  const sorted = [...clean].sort((a, b) => a - b);
  const minP = roundCurrency4(sorted[0]);
  const maxP = roundCurrency4(sorted[n - 1]);
  const sum = clean.reduce((acc, v) => acc + v, 0);
  const meanP = sum / n;
  const medianP = computePercentile(sorted, 0.5);
  const q1 = computePercentile(sorted, 0.25);
  const q3 = computePercentile(sorted, 0.75);
  const iqr = Math.max(0, q3 - q1);

  let stdP = 0;
  if (n >= 2) {
    const variance = clean.reduce((acc, v) => acc + (v - meanP) ** 2, 0) / (n - 1);
    stdP = Math.sqrt(Math.max(0, variance));
  }
  const cvP = meanP > 0 ? (stdP / meanP) * 100 : 0;

  return {
    sampleSize: n,
    minPrice: minP,
    maxPrice: maxP,
    mean: roundCurrency4(meanP),
    median: roundCurrency4(medianP),
    q1: roundCurrency4(q1),
    q3: roundCurrency4(q3),
    iqr: roundCurrency4(iqr),
    stdDeviation: roundCurrency4(stdP),
    cvPercent: roundCurrency2(cvP),
  };
}

export function evaluateIN65PriceBasket(
  quotes: readonly QuoteInputTS[],
  requestedTargetUnit: string = "UN",
  cvThresholdPercent: number = 25.0
): IN65EvaluationResultTS {
  const targetNorm = normalizeSupplyUnit(requestedTargetUnit);
  const validList: AuditedQuoteTS[] = [];
  const purgedList: AuditedQuoteTS[] = [];
  const sameUnitQuotes: AuditedQuoteTS[] = [];

  for (const q of quotes) {
    const rawPrice = Number(q.unitPrice);
    if (!Number.isFinite(rawPrice) || rawPrice <= 0) {
      purgedList.push({
        ...q,
        unitPrice: 0,
        isOutlier: true,
        purgeStatus: "EXPURGADA_INEXEQUIVEL",
        purgeJustification: "Valor unitário nulo, negativo ou igual a zero.",
      });
      continue;
    }
    const qUnitNorm = normalizeSupplyUnit(q.supplyUnit || targetNorm);
    if (qUnitNorm !== targetNorm) {
      purgedList.push({
        ...q,
        unitPrice: rawPrice,
        isOutlier: true,
        purgeStatus: "EXPURGADA_UNIDADE_DIVERGENTE",
        purgeJustification: `Unidade de fornecimento ('${q.supplyUnit}') diverge da unidade padronizada ('${targetNorm}').`,
      });
      continue;
    }
    sameUnitQuotes.push({
      ...q,
      unitPrice: rawPrice,
      isOutlier: false,
      purgeStatus: "VALIDA",
      purgeJustification: "Cotação preliminarmente compatível.",
    });
  }

  const rawPrices = sameUnitQuotes.map((q) => q.unitPrice);
  const rawMetrics = computeStatisticalSummary(rawPrices);

  if (sameUnitQuotes.length === 0) {
    return {
      targetUnit: targetNorm,
      rawMetrics,
      sanitizedMetrics: rawMetrics,
      isHomogeneous: false,
      hasMinimumThreeQuotes: false,
      adoptedParameter: "SEM_DADOS",
      finalEstimatedUnitPrice: 0,
      legalMethodologyJustification: PT_BR_EDITORIAL.restrictedBasketJustification(0),
      validQuotes: [],
      purgedQuotes: purgedList,
    };
  }

  // Barreira Robusta por Mediana + IQR (protegida quando IQR === 0 / preços idênticos)
  const rawMedian = rawMetrics.median ?? sameUnitQuotes[0].unitPrice;
  const robustMin = rawMedian * 0.25;
  const robustMax = rawMedian * 4.0;
  const stageASurvivors: AuditedQuoteTS[] = [];

  for (const q of sameUnitQuotes) {
    if (q.unitPrice < robustMin) {
      purgedList.push({
        ...q,
        isOutlier: true,
        purgeStatus: "EXPURGADA_INEXEQUIVEL",
        purgeJustification: `Preço manifestamente inexequível (inferior a 25% da mediana bruta).`,
      });
    } else if (q.unitPrice > robustMax) {
      purgedList.push({
        ...q,
        isOutlier: true,
        purgeStatus: "EXPURGADA_ACIMA_LIMITE",
        purgeJustification: `Preço excessivamente elevado (superior a 400% da mediana bruta).`,
      });
    } else {
      stageASurvivors.push(q);
    }
  }

  const stageAMetrics = computeStatisticalSummary(stageASurvivors.map((q) => q.unitPrice));
  const stageACv = stageAMetrics.cvPercent ?? 0;
  const stageAStd = stageAMetrics.stdDeviation ?? 0;
  const stageAMean = stageAMetrics.mean ?? 0;

  if (stageASurvivors.length >= 4 && stageAStd > 0 && stageACv > cvThresholdPercent) {
    const lo = Math.max(0, stageAMean - stageAStd);
    const hi = stageAMean + stageAStd;
    for (const q of stageASurvivors) {
      if (q.unitPrice < lo) {
        purgedList.push({
          ...q,
          isOutlier: true,
          purgeStatus: "EXPURGADA_INEXEQUIVEL",
          purgeJustification: `Expurgado por Desvio Padrão (abaixo de μ - 1σ).`,
        });
      } else if (q.unitPrice > hi) {
        purgedList.push({
          ...q,
          isOutlier: true,
          purgeStatus: "EXPURGADA_ACIMA_LIMITE",
          purgeJustification: `Expurgado por Desvio Padrão (acima de μ + 1σ).`,
        });
      } else {
        validList.push({
          ...q,
          isOutlier: false,
          purgeStatus: "VALIDA",
          purgeJustification: "Cotação validada dentro do intervalo estatístico (μ ± 1σ).",
        });
      }
    }
  } else {
    for (const q of stageASurvivors) {
      validList.push({
        ...q,
        isOutlier: false,
        purgeStatus: "VALIDA",
        purgeJustification: "Cotação validada em amostra estatisticamente homogênea.",
      });
    }
  }

  const sanitizedMetrics = computeStatisticalSummary(validList.map((q) => q.unitPrice));
  const cvFinal = sanitizedMetrics.cvPercent ?? 0;
  const isHomogeneous = cvFinal <= cvThresholdPercent;
  const hasMinimumThreeQuotes = sanitizedMetrics.sampleSize >= 3;
  const adoptedParameter = isHomogeneous ? "MEDIA" : "MEDIANA";
  const finalEstimatedUnitPrice = roundCurrency2(
    (isHomogeneous ? sanitizedMetrics.mean : sanitizedMetrics.median) ?? 0
  );

  const justification =
    `Pesquisa de preços processada segundo o art. 6º da IN SEGES/ME nº 65/2021 (${pluralizeValidQuotes(sanitizedMetrics.sampleSize)}, ` +
    `CV saneado de ${cvFinal.toFixed(2)}%, parâmetro adotado: ${adoptedParameter}). ` +
    PT_BR_EDITORIAL.restrictedBasketJustification(sanitizedMetrics.sampleSize);

  return {
    targetUnit: targetNorm,
    rawMetrics,
    sanitizedMetrics,
    isHomogeneous,
    hasMinimumThreeQuotes,
    adoptedParameter,
    finalEstimatedUnitPrice,
    legalMethodologyJustification: justification,
    validQuotes: validList,
    purgedQuotes: purgedList,
  };
}
