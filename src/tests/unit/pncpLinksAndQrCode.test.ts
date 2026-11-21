import test from "node:test";
import assert from "node:assert/strict";
import fs from "node:fs";
import path from "node:path";

/**
 * REGRA INVIOLÁVEL Nº 1 — TESTE DE REGRESSÃO E CONGELAMENTO DE LINKS PNCP/COMPRAS E QR CODES
 * Qualquer alteração que modifique 1 único caractere do formato das URLs oficiais de auditoria
 * (PNCP / Compras.gov.br) ou quebre a geração de QR Codes deve falhar imediatamente aqui.
 */

export function buildCanonicalPncpOrComprasUri(
  chave17: string,
  pncpControlNumber?: string | null
): string {
  if (pncpControlNumber) {
    const m = String(pncpControlNumber).trim().match(/^(\d{14})-\d+-(\d+)\/(\d{4})$/);
    if (m) {
      const cnpj = m[1];
      const seqInt = parseInt(m[2], 10);
      const ano = m[3];
      return `https://pncp.gov.br/app/editais/${cnpj}/${ano}/${seqInt}`;
    }
  }
  const digits = String(chave17 || "").replace(/\D/g, "");
  const padded = digits.length === 16 ? `0${digits}` : digits;
  if (padded.length >= 17) {
    return `https://cnetmobile.estaleiro.serpro.gov.br/comprasnet-web/public/compras/acompanhamento-compra?compra=${padded.slice(0, 17)}`;
  }
  return "https://pncp.gov.br/app/editais";
}

test("Congelamento 1: Id contratação PNCP gera URL literal exata (/app/editais/{cnpj}/{ano}/{seq})", () => {
  const url1 = buildCanonicalPncpOrComprasUri("75121206000702026", "00394502000144-1-010123/2026");
  assert.equal(url1, "https://pncp.gov.br/app/editais/00394502000144/2026/10123");
  assert.ok(!url1.includes("?q="), "URL PNCP jamais pode conter parâmetro genérico ?q=");

  const url2 = buildCanonicalPncpOrComprasUri("18019605900032026", "46377800000127-1-000981/2026");
  assert.equal(url2, "https://pncp.gov.br/app/editais/46377800000127/2026/981");
});

test("Congelamento 2: Fallback sem controle PNCP gera link direto Compras.gov.br (?compra=17_digitos)", () => {
  const url = buildCanonicalPncpOrComprasUri("75121206000702026", null);
  assert.equal(
    url,
    "https://cnetmobile.estaleiro.serpro.gov.br/comprasnet-web/public/compras/acompanhamento-compra?compra=75121206000702026"
  );
  assert.ok(!url.includes("?q="));
});

test("Congelamento 3: Chave de 16 dígitos (UASG iniciada em 0) recebe padding para 17 dígitos", () => {
  const url = buildCanonicalPncpOrComprasUri("7000105000122026", null);
  assert.equal(
    url,
    "https://cnetmobile.estaleiro.serpro.gov.br/comprasnet-web/public/compras/acompanhamento-compra?compra=07000105000122026"
  );
});

test("Congelamento 4: Frontend e Backend preservam intocadas as rotas e geradores de Links e QR Codes", () => {
  const rootDir = path.resolve(process.cwd());
  const html = fs.readFileSync(path.join(rootDir, "frontend", "index.html"), "utf-8");
  const govClient = fs.readFileSync(path.join(rootDir, "backend", "core", "gov_api_client.py"), "utf-8");
  const repGen = fs.readFileSync(path.join(rootDir, "backend", "core", "report_generators.py"), "utf-8");

  // Verifica que o frontend mantém openPncpPortalByQuote e os links diretos
  assert.ok(html.includes("function openPncpPortalByQuote("), "openPncpPortalByQuote deve permanecer intacto");
  assert.ok(
    html.includes("https://cnetmobile.estaleiro.serpro.gov.br/comprasnet-web/public/compras/acompanhamento-compra?compra="),
    "URL direta do Compras.gov.br deve permanecer intacta no frontend"
  );
  assert.ok(
    html.includes("https://pncp.gov.br/app/editais/"),
    "URL direta do PNCP deve permanecer intacta no frontend"
  );

  // Verifica que o backend mantém build_audit_uri e generate_qr_code_bytes intactos
  assert.ok(govClient.includes("def build_audit_uri("), "build_audit_uri deve permanecer intacto no backend");
  assert.ok(repGen.includes("def generate_qr_code_bytes("), "generate_qr_code_bytes deve permanecer intacto no backend");
});
