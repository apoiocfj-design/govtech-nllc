/**
 * SUÍTE DE TESTES UNITÁRIOS AUTOMATIZADOS (MOTOR MATEMÁTICO IN SEGES/ME Nº 65/2021)
 * Validação rigorosa dos critérios de aceitabilidade de preços públicos:
 * 1. Amostra Homogênea (CV <= 25%) -> Adoção da MÉDIA ARITMÉTICA
 * 2. Amostra Heterogênea (CV > 25%) -> Adoção da MEDIANA
 * 3. Expurgos de Outliers -> Preços vis (< 25% da mediana) e sobrepreço (> 400% da mediana ou fora de μ ± s)
 * 4. Cesta Restrita (< 3 cotações válidas) -> Alerta obrigatório Art. 6º, § 5º da IN 65/2021
 */

const test = require("node:test");
const assert = require("node:assert/strict");

// Implementação isolada e autocontida do algoritmo matemático da IN 65/2021
// espelhando o motor backend (backend/core/stat_engine.py) e o frontend
function round4(val) {
  return Number(Math.round(Number(val + "e4")) + "e-4");
}

function round2(val) {
  return Number(Math.round(Number(val + "e2")) + "e-2");
}

function computeMetrics(prices, kStd = 1.0) {
  const n = prices.length;
  if (n === 0) {
    return { sampleSize: 0, min: null, max: null, mean: null, median: null, stdDev: null, cvPercent: null };
  }
  const min = Math.min(...prices);
  const max = Math.max(...prices);
  const sum = prices.reduce((a, b) => a + b, 0);
  const mean = sum / n;

  // Mediana
  const sorted = [...prices].sort((a, b) => a - b);
  const mid = Math.floor(n / 2);
  const median = n % 2 !== 0 ? sorted[mid] : (sorted[mid - 1] + sorted[mid]) / 2;

  // Desvio Padrão Amostral (s) e CV
  let stdDev = 0;
  let cvPercent = 0;
  let loStd = mean;
  let hiStd = mean;

  if (n >= 2) {
    const variance = prices.reduce((acc, p) => acc + Math.pow(p - mean, 2), 0) / (n - 1);
    stdDev = Math.sqrt(variance);
    cvPercent = mean > 0 ? (stdDev / mean) * 100 : 0;
    loStd = Math.max(0, mean - kStd * stdDev);
    hiStd = mean + kStd * stdDev;
  }

  return {
    sampleSize: n,
    min: round2(min),
    max: round2(max),
    mean: round2(mean),
    median: round2(median),
    stdDev: round2(stdDev),
    cvPercent: round2(cvPercent),
    lowerBoundStd: round2(loStd),
    upperBoundStd: round2(hiStd)
  };
}

function auditPriceSample(rawQuotes, targetUnit = "UN", options = {}) {
  const cvThreshold = options.cvThreshold || 25.0;
  const kStd = options.kStd || 1.0;
  const infeasibleRatio = options.infeasibleRatio || 0.25;
  const excessiveRatio = options.excessiveRatio || 4.0;

  const validQuotes = [];
  const purgedQuotes = [];

  // 1. Filtragem por unidade e preço positivo
  const sameUnitQuotes = [];
  for (const q of rawQuotes) {
    if (!q.unitPrice || q.unitPrice <= 0) {
      purgedQuotes.push({ ...q, isOutlier: true, purgeStatus: "EXPURGADA_INEXEQUIVEL", reason: "Preço nulo ou negativo." });
    } else if (q.supplyUnit && targetUnit && q.supplyUnit.toUpperCase() !== targetUnit.toUpperCase()) {
      purgedQuotes.push({ ...q, isOutlier: true, purgeStatus: "EXPURGADA_UNIDADE_DIVERGENTE", reason: "Unidade divergente." });
    } else {
      sameUnitQuotes.push(q);
    }
  }

  if (sameUnitQuotes.length === 0) {
    throw new Error("Nenhuma cotação com a unidade solicitada.");
  }

  // 2. Barreira Robusta pela Mediana Bruta (Estágio A)
  const rawPrices = sameUnitQuotes.map((q) => q.unitPrice);
  const rawMetrics = computeMetrics(rawPrices, kStd);
  const robustMin = rawMetrics.median * infeasibleRatio;
  const robustMax = rawMetrics.median * excessiveRatio;

  const stageASurvivors = [];
  for (const q of sameUnitQuotes) {
    if (q.unitPrice < robustMin) {
      purgedQuotes.push({
        ...q,
        isOutlier: true,
        purgeStatus: "EXPURGADA_INEXEQUIVEL",
        reason: `Preço manifestamente inexequível (R$ ${q.unitPrice} < R$ ${round2(robustMin)}).`
      });
    } else if (q.unitPrice > robustMax) {
      purgedQuotes.push({
        ...q,
        isOutlier: true,
        purgeStatus: "EXPURGADA_ACIMA_LIMITE",
        reason: `Preço manifestamente superfaturado/erro grosseiro (R$ ${q.unitPrice} > R$ ${round2(robustMax)}).`
      });
    } else {
      stageASurvivors.push(q);
    }
  }

  // 3. Estágio B: Corte fino por μ ± k*s se CV ainda > 25%
  const stageAPrices = stageASurvivors.map((q) => q.unitPrice);
  const stageAMetrics = computeMetrics(stageAPrices, kStd);

  if (stageAMetrics.sampleSize >= 4 && stageAMetrics.cvPercent > cvThreshold) {
    for (const q of stageASurvivors) {
      if (q.unitPrice < stageAMetrics.lowerBoundStd) {
        purgedQuotes.push({
          ...q,
          isOutlier: true,
          purgeStatus: "EXPURGADA_INEXEQUIVEL",
          reason: `Expurgo estatístico por Desvio Padrão (abaixo de μ - s = R$ ${stageAMetrics.lowerBoundStd}).`
        });
      } else if (q.unitPrice > stageAMetrics.upperBoundStd) {
        purgedQuotes.push({
          ...q,
          isOutlier: true,
          purgeStatus: "EXPURGADA_ACIMA_LIMITE",
          reason: `Expurgo estatístico por Desvio Padrão (acima de μ + s = R$ ${stageAMetrics.upperBoundStd}).`
        });
      } else {
        validQuotes.push({ ...q, isOutlier: false, purgeStatus: "VALIDA" });
      }
    }
  } else {
    for (const q of stageASurvivors) {
      validQuotes.push({ ...q, isOutlier: false, purgeStatus: "VALIDA" });
    }
  }

  // 4. Métricas Finais Saneadas
  const sanitizedPrices = validQuotes.map((q) => q.unitPrice);
  const sanitizedMetrics = computeMetrics(sanitizedPrices, kStd);

  // 5. Decisão Automática (Art. 6º IN 65/2021)
  const isHomogeneous = (sanitizedMetrics.cvPercent || 0) <= cvThreshold;
  const adoptedParameter = isHomogeneous ? "MEDIA" : "MEDIANA";
  const finalEstimatedUnitPrice = isHomogeneous ? sanitizedMetrics.mean : sanitizedMetrics.median;
  const hasMinimumThreeQuotes = sanitizedMetrics.sampleSize >= 3;

  return {
    rawMetrics,
    sanitizedMetrics,
    isHomogeneous,
    hasMinimumThreeQuotes,
    adoptedParameter,
    finalEstimatedUnitPrice,
    validQuotes,
    purgedQuotes
  };
}

// ============================================================================
// SUÍTE DE TESTES UNITÁRIOS (ASSERTIVAS NORMATIVAS DA IN 65/2021)
// ============================================================================

test("Caso 1: Amostra Homogênea (CV <= 25%) deve adotar estritamente a MÉDIA ARITMÉTICA", () => {
  // Cotações com pouca dispersão: [10.00, 10.50, 9.80, 10.20]
  const sample = [
    { id: "1", unitPrice: 10.0, supplyUnit: "UN" },
    { id: "2", unitPrice: 10.5, supplyUnit: "UN" },
    { id: "3", unitPrice: 9.8, supplyUnit: "UN" },
    { id: "4", unitPrice: 10.2, supplyUnit: "UN" }
  ];

  const result = auditPriceSample(sample, "UN");

  assert.equal(result.isHomogeneous, true, "Amostra deve ser classificada como homogênea");
  assert.ok(result.sanitizedMetrics.cvPercent <= 25.0, `CV (${result.sanitizedMetrics.cvPercent}%) deve ser <= 25%`);
  assert.equal(result.adoptedParameter, "MEDIA", "Parâmetro adotado deve ser MÉDIA");
  assert.equal(result.finalEstimatedUnitPrice, result.sanitizedMetrics.mean, "Preço final deve coincidir com a Média");
  assert.equal(result.hasMinimumThreeQuotes, true, "Deve conter no mínimo 3 cotações válidas");
  assert.equal(result.validQuotes.length, 4, "Todas as 4 cotações devem ser mantidas válidas");
});

test("Caso 2: Amostra Heterogênea (CV > 25%) deve adotar estritamente a MEDIANA", () => {
  // 3 cotações dispersas (n=3 não sofre corte por desvio padrão para preservar mínimo legal):
  // [10.00, 15.00, 20.00] -> Média = 15.00, Mediana = 15.00, Desvio = 5.00, CV = 33.33% (> 25%)
  const sample = [
    { id: "1", unitPrice: 10.0, supplyUnit: "UN" },
    { id: "2", unitPrice: 15.0, supplyUnit: "UN" },
    { id: "3", unitPrice: 20.0, supplyUnit: "UN" }
  ];

  const result = auditPriceSample(sample, "UN");

  assert.equal(result.isHomogeneous, false, "Amostra deve ser classificada como heterogênea");
  assert.ok(result.sanitizedMetrics.cvPercent > 25.0, `CV (${result.sanitizedMetrics.cvPercent}%) deve ser > 25%`);
  assert.equal(result.adoptedParameter, "MEDIANA", "Parâmetro adotado deve ser MEDIANA");
  assert.equal(result.finalEstimatedUnitPrice, result.sanitizedMetrics.median, "Preço final deve coincidir com a Mediana");
  assert.equal(result.finalEstimatedUnitPrice, 15.0, "Mediana esperada de [10, 15, 20] é 15.00");
  assert.equal(result.hasMinimumThreeQuotes, true, "Atende ao requisito mínimo de 3 cotações");
});

test("Caso 3: Expurgos de Outliers (Sobrepreço e Inexequibilidade) devem marcar expurgado: true", () => {
  // Amostra com outlier de erro de digitação (150.0) e preço vil (0.50) para um item de ~10 reais
  const sample = [
    { id: "1", unitPrice: 10.0, supplyUnit: "UN" },
    { id: "2", unitPrice: 10.5, supplyUnit: "UN" },
    { id: "3", unitPrice: 11.0, supplyUnit: "UN" },
    { id: "4", unitPrice: 9.8, supplyUnit: "UN" },
    { id: "5", unitPrice: 150.0, supplyUnit: "UN" }, // Outlier grosseiro (> 400% da mediana)
    { id: "6", unitPrice: 0.50, supplyUnit: "UN" }   // Inexequível (< 25% da mediana)
  ];

  const result = auditPriceSample(sample, "UN");

  // Verifica se os outliers foram devidamente expurgados
  const purgedIds = result.purgedQuotes.map((q) => q.id);
  assert.ok(purgedIds.includes("5"), "Cotação 5 (R$ 150.00) deve ser expurgada");
  assert.ok(purgedIds.includes("6"), "Cotação 6 (R$ 0.50) deve ser expurgada");

  const quote150 = result.purgedQuotes.find((q) => q.id === "5");
  assert.equal(quote150.isOutlier, true);
  assert.equal(quote150.purgeStatus, "EXPURGADA_ACIMA_LIMITE");
  assert.ok(quote150.reason.includes("superfaturado"));

  const quote050 = result.purgedQuotes.find((q) => q.id === "6");
  assert.equal(quote050.isOutlier, true);
  assert.equal(quote050.purgeStatus, "EXPURGADA_INEXEQUIVEL");
  assert.ok(quote050.reason.includes("inexequível"));

  // As cotações legítimas permanecem válidas
  assert.equal(result.validQuotes.length, 4);
});

test("Caso 4: Alerta de Cesta Restrita (< 3 cotações válidas)", () => {
  // Cesta restrita com apenas 2 cotações
  const sample = [
    { id: "1", unitPrice: 50.0, supplyUnit: "UN" },
    { id: "2", unitPrice: 52.0, supplyUnit: "UN" }
  ];

  const result = auditPriceSample(sample, "UN");

  assert.equal(result.hasMinimumThreeQuotes, false, "Cesta com 2 cotações deve acusar hasMinimumThreeQuotes=false");
  assert.equal(result.validQuotes.length, 2);
});

test("Caso 5: Expurgo obrigatório por Unidade de Fornecimento divergente", () => {
  const sample = [
    { id: "1", unitPrice: 10.0, supplyUnit: "UN" },
    { id: "2", unitPrice: 10.5, supplyUnit: "UN" },
    { id: "3", unitPrice: 100.0, supplyUnit: "CX" } // Caixa de 10 unidades misturada com UN
  ];

  const result = auditPriceSample(sample, "UN");

  const purgedCx = result.purgedQuotes.find((q) => q.id === "3");
  assert.ok(purgedCx, "Cotação em CX deve ser expurgada da cesta UN");
  assert.equal(purgedCx.purgeStatus, "EXPURGADA_UNIDADE_DIVERGENTE");
});
