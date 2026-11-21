import test from "node:test";
import assert from "node:assert/strict";
import fs from "node:fs";
import path from "node:path";
import {
  buildUnifiedDocumentHeader,
  FORBIDDEN_BRANDING_STRINGS,
  formatNumberedItemFilename,
  normalizeProcessFilename,
  resolveBatchLifecycleState,
  ZIP_PROCESS_FOLDERS,
} from "../../infrastructure/documents/documentHeaderTemplate.ts";
import { maskSensitiveData } from "../../infrastructure/observability/auditLogger.ts";

test("Regra Inviolável Nº 2: Nenhum gerador de documento (Word/PDF/Excel/ZIP/PNCP) contém 'GOVTECH NLLC'", () => {
  const rootDir = path.resolve(process.cwd());
  const generatorFiles = [
    "backend/core/report_generators.py",
    "backend/core/agu_batch_reports.py",
    "backend/core/pncp_extractor.py",
    "backend/core/pncp_integration.py",
    "backend/core/audited_document_engine.py",
  ];

  for (const relPath of generatorFiles) {
    const absPath = path.join(rootDir, relPath);
    const content = fs.readFileSync(absPath, "utf-8");
    const upper = content.toUpperCase();

    for (const forbidden of FORBIDDEN_BRANDING_STRINGS) {
      assert.ok(
        !upper.includes(forbidden),
        `String proibida '${forbidden}' encontrada em ${relPath}`
      );
    }
  }
});

test("Cabeçalho Documental Unificado e Fuso Horário America/Bahia", () => {
  const hdr = buildUnifiedDocumentHeader({
    organizationName: "Capitania Fluvial de Juazeiro",
    documentTitle: "Caderno Consolidado de Preços",
  });
  assert.equal(hdr.line1Republic, "REPÚBLICA FEDERATIVA DO BRASIL");
  assert.equal(hdr.line2Organization, "CAPITANIA FLUVIAL DE JUAZEIRO");
  assert.equal(hdr.documentTitle, "CADERNO CONSOLIDADO DE PREÇOS");
  for (const forbidden of FORBIDDEN_BRANDING_STRINGS) {
    assert.ok(!JSON.stringify(hdr).toUpperCase().includes(forbidden));
  }
});

test("Normalização de nomes de arquivos do ZIP (Windows Explorer & SEI/SIGAD) e Árvore do Processo", () => {
  const fname = formatNumberedItemFilename(1, 241827, "Malote Lona / Algodão: Visor*Azul?", "pdf");
  assert.equal(fname, "Item_001_241827_Malote_Lona_Algodao_Visor_Azul.pdf");
  assert.ok(!/[\\/:*?"<>|]/.test(fname));

  assert.equal(ZIP_PROCESS_FOLDERS.rootIndexTxt, "00_Indice_do_Processo.txt");
  assert.equal(ZIP_PROCESS_FOLDERS.folder01Consolidacao, "01_Consolidacao_Geral");
  assert.equal(ZIP_PROCESS_FOLDERS.folder02PesquisaItem, "02_Pesquisa_por_Item");
  assert.equal(ZIP_PROCESS_FOLDERS.folder03AnaliseJustificativas, "03_Analise_Estatistica_e_Justificativas");
  assert.equal(ZIP_PROCESS_FOLDERS.folder04PlanilhasAuditoria, "04_Planilhas_e_Auditoria");
});

test("Máquina de Estados do Lote e Mascaramento LGPD/AppSec", () => {
  assert.equal(
    resolveBatchLifecycleState({
      trayCount: 0,
      isResearching: false,
      hasResults: false,
      currentSignature: "",
      lastResearchedSignature: "",
    }),
    "LOTE_VAZIO"
  );
  assert.equal(
    resolveBatchLifecycleState({
      trayCount: 5,
      isResearching: false,
      hasResults: true,
      currentSignature: "A:10",
      lastResearchedSignature: "A:5",
    }),
    "DESATUALIZADO"
  );

  const masked = maskSensitiveData("CPF 123.456.789-01 Token Bearer eyJhbGciOiJIUzI1NiJ9");
  assert.ok(!masked.includes("456.789"));
  assert.ok(masked.includes("[REDACTED]"));
});
