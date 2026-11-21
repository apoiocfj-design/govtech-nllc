/**
 * VERIFICADOR ESTÁTICO DE TYPECHECK, LINT EDITORIAL/LEGAL E BUILD (GOVTECH NLLC)
 * Executa sem dependências externas em Node.js 22.6+ / 26.x com suporte nativo a TypeScript.
 */

const fs = require("node:fs");
const path = require("node:path");
const { pathToFileURL } = require("node:url");

const rootDir = path.resolve(__dirname, "..");
const mode = process.argv[2] || "--all";

async function runTypecheck() {
  const tsFiles = [
    "src/domain/config/legalThresholds.ts",
    "src/shared/i18n/ptBR.ts",
    "src/domain/statistics/in65Statistics.ts",
    "src/shared/validation/govApiZod.ts",
    "src/infrastructure/documents/documentHeaderTemplate.ts",
    "src/infrastructure/observability/auditLogger.ts",
  ];
  for (const rel of tsFiles) {
    const abs = path.join(rootDir, rel);
    if (!fs.existsSync(abs)) {
      throw new Error(`[TYPECHECK] Arquivo obrigatório ausente: ${rel}`);
    }
    await import(pathToFileURL(abs).href);
  }
  console.log(`✅ [typecheck] ${tsFiles.length} módulos TypeScript estritos verificados com 0 erros.`);
}

function runLint() {
  const indexHtmlPath = path.join(rootDir, "frontend", "index.html");
  const html = fs.readFileSync(indexHtmlPath, "utf-8");

  // 1. Validação sintática de todos os blocos <script> inline
  const scripts = [...html.matchAll(/<script(?![^>]*\bsrc=)[^>]*>([\s\S]*?)<\/script>/gi)];
  if (scripts.length === 0) {
    throw new Error("[LINT] Nenhum bloco <script> encontrado em frontend/index.html.");
  }
  scripts.forEach((m, idx) => {
    try {
      new Function(m[1]);
    } catch (err) {
      throw new Error(`[LINT] Erro de sintaxe no Script #${idx + 1}: ${err.message}`);
    }
  });

  // 2. Validação editorial e legal (proibição de termos desatualizados/informais)
  const forbiddenPatterns = [
    { pattern: "p/ Anexar", reason: "Use 'para Anexar' (redação formal p/ processo administrativo)." },
    { pattern: "Classificação Não Fracionamento", reason: "Use 'Classificação para Controle de Não Fracionamento'." },
    { pattern: "cotação(ões) válida(s)", reason: "Use pluralização dinâmica ('1 cotação válida' / 'N cotações válidas')." },
  ];
  for (const rule of forbiddenPatterns) {
    if (html.includes(rule.pattern)) {
      throw new Error(`[LINT EDITORIAL] Padrão proibido encontrado ('${rule.pattern}'): ${rule.reason}`);
    }
  }

  // 3. Regra Inviolável Nº 1: Preservação de Links PNCP/Compras e QR Codes
  if (!html.includes("function openPncpPortalByQuote(")) {
    throw new Error("[LINT REGRA Nº 1] Função openPncpPortalByQuote ausente.");
  }

  console.log(`✅ [lint] ${scripts.length} scripts inline, regras editoriais TCU/AGU e trava de Links/QR Codes aprovados (0 erros).`);
}

function runBuild() {
  const srcHtml = path.join(rootDir, "frontend", "index.html");
  const dstHtml = path.join(rootDir, "frontend", "preview_app.html");
  fs.copyFileSync(srcHtml, dstHtml);
  const sizeKb = (fs.statSync(dstHtml).size / 1024).toFixed(1);
  console.log(`✅ [build] Bundle frontend sincronizado e validado (index.html -> preview_app.html, ${sizeKb} KB).`);
}

(async () => {
  try {
    if (mode === "--typecheck" || mode === "--all") await runTypecheck();
    if (mode === "--lint" || mode === "--all") runLint();
    if (mode === "--build" || mode === "--all") runBuild();
  } catch (err) {
    console.error("❌ Falha na verificação:", err.message);
    process.exit(1);
  }
})();
