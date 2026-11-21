import test from "node:test";
import assert from "node:assert/strict";
import {
  evaluateFracionamentoRisk,
  getActiveLegalThreshold,
} from "../../domain/config/legalThresholds.ts";
import {
  evaluateIN65PriceBasket,
  sanitizeSpreadsheetCell,
  sumMonetarySafe,
} from "../../domain/statistics/in65Statistics.ts";
import {
  PT_BR_EDITORIAL,
  pluralizeObtainedQuotes,
  pluralizeValidQuotes,
} from "../../shared/i18n/ptBR.ts";
import {
  computeExponentialBackoffMs,
  GovQuoteZodSchema,
} from "../../shared/validation/govApiZod.ts";

test("1. Parametrização Legal: Decreto 12.343/2024 (R$ 62.725,59) vs Decreto 11.871/2023 (R$ 59.906,02)", () => {
  const t2024 = getActiveLegalThreshold(2024);
  assert.equal(t2024.art75IncisoII, 59906.02);
  assert.match(t2024.decreeNumber, /11\.871\/2023/);

  const t2026 = getActiveLegalThreshold(2026);
  assert.equal(t2026.art75IncisoII, 62725.59);
  assert.match(t2026.decreeNumber, /12\.343\/2024/);

  const risk = evaluateFracionamentoRisk(50000, "II", 2026);
  assert.equal(risk.isWarning, true);
  assert.equal(risk.isExceeded, false);
  assert.match(risk.decreeCitation, /12\.343\/2024/);
});

test("2. Motor Estatístico: Série com todos os preços idênticos (IQR=0, s=0, CV=0%, sem NaN/Infinity)", () => {
  const res = evaluateIN65PriceBasket([
    { unitPrice: 180.0, supplyUnit: "UN" },
    { unitPrice: 180.0, supplyUnit: "UN" },
    { unitPrice: 180.0, supplyUnit: "UN" },
    { unitPrice: 180.0, supplyUnit: "UN" },
  ]);
  assert.equal(res.sanitizedMetrics.sampleSize, 4);
  assert.equal(res.sanitizedMetrics.iqr, 0);
  assert.equal(res.sanitizedMetrics.stdDeviation, 0);
  assert.equal(res.sanitizedMetrics.cvPercent, 0);
  assert.equal(res.isHomogeneous, true);
  assert.equal(res.adoptedParameter, "MEDIA");
  assert.equal(res.finalEstimatedUnitPrice, 180.0);
});

test("3. Motor Estatístico: Edge cases com 0, 1 e 2 cotações válidas (Art. 6º, § 5º IN 65/2021)", () => {
  const res0 = evaluateIN65PriceBasket([
    { unitPrice: 0, supplyUnit: "UN" },
    { unitPrice: -15, supplyUnit: "UN" },
    { unitPrice: null, supplyUnit: "UN" },
  ]);
  assert.equal(res0.hasMinimumThreeQuotes, false);
  assert.equal(res0.validQuotes.length, 0);
  assert.equal(res0.adoptedParameter, "SEM_DADOS");

  const res1 = evaluateIN65PriceBasket([{ unitPrice: 95.5, supplyUnit: "UN" }]);
  assert.equal(res1.hasMinimumThreeQuotes, false);
  assert.equal(res1.validQuotes.length, 1);
  assert.equal(res1.finalEstimatedUnitPrice, 95.5);
  assert.ok(res1.legalMethodologyJustification.includes("1 cotação válida"));

  const res2 = evaluateIN65PriceBasket([
    { unitPrice: 100.0, supplyUnit: "UN" },
    { unitPrice: 110.0, supplyUnit: "UN" },
  ]);
  assert.equal(res2.hasMinimumThreeQuotes, false);
  assert.equal(res2.validQuotes.length, 2);
  assert.ok(res2.legalMethodologyJustification.includes("2 cotações válidas"));
});

test("4. Precisão Monetária Segura e Proteção contra CSV/Excel Formula Injection", () => {
  assert.equal(sumMonetarySafe([0.1, 0.2]), 0.3);
  assert.equal(sumMonetarySafe([19.99, 20.01, 59.95]), 99.95);

  assert.equal(sanitizeSpreadsheetCell("=SUM(A1:A10)"), "'=SUM(A1:A10)");
  assert.equal(sanitizeSpreadsheetCell("+cmd|' /C calc'!A0"), "'+cmd|' /C calc'!A0");
  assert.equal(sanitizeSpreadsheetCell("-100+200"), "'-100+200");
  assert.equal(sanitizeSpreadsheetCell("@IMPORTXML()"), "'@IMPORTXML()");
  assert.equal(sanitizeSpreadsheetCell("FORNECEDOR REGULAR LTDA"), "FORNECEDOR REGULAR LTDA");
});

test("5. Revisão Editorial Formal e Validação Zod + Backoff", () => {
  assert.equal(pluralizeValidQuotes(0), "0 cotações válidas");
  assert.equal(pluralizeValidQuotes(1), "1 cotação válida");
  assert.equal(pluralizeValidQuotes(4), "4 cotações válidas");
  assert.equal(pluralizeObtainedQuotes(1), "1 cotação obtida");
  assert.ok(!PT_BR_EDITORIAL.autoJustificationHeader.includes("p/ Anexar"));
  assert.ok(PT_BR_EDITORIAL.autoJustificationHeader.includes("para Anexar"));
  assert.equal(
    PT_BR_EDITORIAL.nonFractionationHeading,
    "Classificação para Controle de Não Fracionamento (Art. 75, § 1º da Lei nº 14.133/2021)"
  );

  const parsed = GovQuoteZodSchema.safeParse({
    unit_price: 42.5,
    supplier_name: "EMPRESA TESTE",
    supply_unit: "un",
  });
  assert.equal(parsed.success, true);

  const invalid = GovQuoteZodSchema.safeParse({ unit_price: -5 });
  assert.equal(invalid.success, false);

  assert.equal(computeExponentialBackoffMs(0, 500), 500);
  assert.equal(computeExponentialBackoffMs(1, 500), 1000);
  assert.equal(computeExponentialBackoffMs(2, 500), 2000);
});
