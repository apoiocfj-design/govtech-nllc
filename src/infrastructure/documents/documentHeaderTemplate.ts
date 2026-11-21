/**
 * FONTE ÚNICA DA VERDADE — IDENTIDADE DOCUMENTAL UNIFICADA (WORD, PDF, EXCEL, PRINT A4 E ZIP)
 * Regra Inviolável Nº 2: Proibição absoluta de menção a "SISTEMA DE GESTÃO DE PREÇOS E CONTRATAÇÕES
 * PÚBLICAS — GOVTECH NLLC" ou "GOVTECH NLLC" em qualquer arquivo exportado ou impresso.
 * Fuso Horário Oficial: America/Bahia (UTC-03:00).
 */

export const FORBIDDEN_BRANDING_STRINGS: readonly string[] = [
  "SISTEMA DE GESTÃO DE PREÇOS E CONTRATAÇÕES PÚBLICAS — GOVTECH NLLC",
  "SISTEMA DE GESTAO DE PRECOS E CONTRATACOES PUBLICAS — GOVTECH NLLC",
  "SISTEMA GOVTECH NLLC",
  "GOVTECH NLLC",
] as const;

export interface DocumentHeaderInput {
  readonly organizationName: string;
  readonly sectorName?: string;
  readonly documentTitle: string;
  readonly normativeSubtitle?: string;
  readonly referenceDate?: Date;
}

export interface UnifiedDocumentHeader {
  readonly line1Republic: string;
  readonly line2Organization: string;
  readonly line3Sector: string;
  readonly documentTitle: string;
  readonly normativeSubtitle: string;
  readonly issuedAtBahia: string;
}

export function formatDateTimeBahia(date: Date = new Date()): string {
  try {
    return new Intl.DateTimeFormat("pt-BR", {
      timeZone: "America/Bahia",
      day: "2-digit",
      month: "2-digit",
      year: "numeric",
      hour: "2-digit",
      minute: "2-digit",
      second: "2-digit",
    }).format(date);
  } catch {
    return date.toISOString();
  }
}

export function buildUnifiedDocumentHeader(input: DocumentHeaderInput): UnifiedDocumentHeader {
  const org = String(input.organizationName || "MARINHA DO BRASIL").trim().toUpperCase();
  const sector = String(input.sectorName || "SETOR DE OBTENÇÃO E CONTRATAÇÕES PÚBLICAS").trim().toUpperCase();
  const title = String(input.documentTitle || "RELATÓRIO DE PESQUISA DE PREÇOS").trim().toUpperCase();
  const sub = String(
    input.normativeSubtitle || "Instrução Normativa SEGES/ME nº 65/2021 • Art. 23 da Lei Federal nº 14.133/2021"
  ).trim();

  return {
    line1Republic: "REPÚBLICA FEDERATIVA DO BRASIL",
    line2Organization: org,
    line3Sector: sector,
    documentTitle: title,
    normativeSubtitle: sub,
    issuedAtBahia: formatDateTimeBahia(input.referenceDate ?? new Date()),
  };
}

/**
 * Sanitiza nomes de arquivos para Windows Explorer e SEI/SIGAD:
 * Remove acentos quebrados e caracteres proibidos (\/:*?"<>|), aplicando padding 001, 002.
 */
export function normalizeProcessFilename(rawName: string, maxLen: number = 40): string {
  const ascii = String(rawName || "Item")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[\\/:*?"<>|]+/g, "_")
    .replace(/[^A-Za-z0-9_-]+/g, "_")
    .replace(/_+/g, "_")
    .replace(/^_|_$/g, "");
  return (ascii.slice(0, maxLen).replace(/_$/g, "") || "Item");
}

export function formatNumberedItemFilename(
  index1Based: number,
  catalogCode: string | number,
  shortDescription: string,
  extension: "pdf" | "docx" = "pdf"
): string {
  const num = String(Math.max(1, Math.trunc(index1Based))).padStart(3, "0");
  const code = String(catalogCode || "000000").replace(/[^A-Za-z0-9]/g, "");
  const slug = normalizeProcessFilename(shortDescription, 32);
  return `Item_${num}_${code}_${slug}.${extension}`;
}

export const ZIP_PROCESS_FOLDERS = {
  rootIndexTxt: "00_Indice_do_Processo.txt",
  rootManualPdf: "00_Manual_Didatico_do_Processo.pdf",
  folder01Consolidacao: "01_Consolidacao_Geral",
  folder02PesquisaItem: "02_Pesquisa_por_Item",
  folder03AnaliseJustificativas: "03_Analise_Estatistica_e_Justificativas",
  folder04PlanilhasAuditoria: "04_Planilhas_e_Auditoria",
} as const;

/**
 * Máquina de Estados Determinística do Lote (invalida o ZIP anterior quando o lote é editado).
 */
export type BatchLifecycleState =
  | "LOTE_VAZIO"
  | "LOTE_MONTADO"
  | "PESQUISANDO"
  | "CONSOLIDADO"
  | "DESATUALIZADO";

export function resolveBatchLifecycleState(params: {
  readonly trayCount: number;
  readonly isResearching: boolean;
  readonly hasResults: boolean;
  readonly currentSignature: string;
  readonly lastResearchedSignature: string;
}): BatchLifecycleState {
  if (params.trayCount <= 0) return "LOTE_VAZIO";
  if (params.isResearching) return "PESQUISANDO";
  if (!params.hasResults) return "LOTE_MONTADO";
  if (params.currentSignature !== params.lastResearchedSignature) return "DESATUALIZADO";
  return "CONSOLIDADO";
}
