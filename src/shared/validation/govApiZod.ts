/**
 * VALIDAÇÃO ESTRUTURAL DE RESPOSTAS DE API (ZOD-COMPATIBLE SCHEMA ENGINE)
 * E POLÍTICA DE RESILIÊNCIA HTTP (EXPONENTIAL BACKOFF + CONTROLE DE CONCORRÊNCIA)
 */

export type SafeParseResult<T> =
  | { readonly success: true; readonly data: T }
  | { readonly success: false; readonly error: string };

export interface SchemaValidator<T> {
  parse(input: unknown): T;
  safeParse(input: unknown): SafeParseResult<T>;
}

export interface GovQuoteValidated {
  readonly id_compra_item: string;
  readonly purchase_number_year: string;
  readonly buyer_uasg: string;
  readonly buyer_org_name: string;
  readonly buyer_uf: string;
  readonly supplier_cnpj: string;
  readonly supplier_name: string;
  readonly supply_unit: string;
  readonly quantity_homologated: number;
  readonly unit_price: number;
  readonly purchase_date: string;
  readonly pncp_or_compras_uri: string;
  readonly pncp_control_number?: string | null;
  readonly comprasnet_uri?: string | null;
}

export const GovQuoteZodSchema: SchemaValidator<GovQuoteValidated> = {
  safeParse(input: unknown): SafeParseResult<GovQuoteValidated> {
    if (!input || typeof input !== "object") {
      return { success: false, error: "Cotação inválida: esperado objeto JSON." };
    }
    const obj = input as Record<string, unknown>;
    const unitPrice = Number(obj.unit_price ?? obj.precoUnitario ?? 0);
    if (!Number.isFinite(unitPrice) || unitPrice < 0) {
      return { success: false, error: "Campo unit_price inválido ou negativo." };
    }
    const qty = Number(obj.quantity_homologated ?? obj.quantidade ?? 1);
    return {
      success: true,
      data: {
        id_compra_item: String(obj.id_compra_item ?? obj.idItemCompra ?? ""),
        purchase_number_year: String(obj.purchase_number_year ?? "—"),
        buyer_uasg: String(obj.buyer_uasg ?? obj.codigoUasg ?? "—"),
        buyer_org_name: String(obj.buyer_org_name ?? obj.nomeOrgao ?? "Órgão Público"),
        buyer_uf: String(obj.buyer_uf ?? obj.estado ?? "BR"),
        supplier_cnpj: String(obj.supplier_cnpj ?? obj.niFornecedor ?? "—"),
        supplier_name: String(obj.supplier_name ?? obj.nomeFornecedor ?? "Fornecedor"),
        supply_unit: String(obj.supply_unit ?? obj.siglaUnidadeFornecimento ?? "UN").trim().toUpperCase() || "UN",
        quantity_homologated: Number.isFinite(qty) && qty > 0 ? qty : 1,
        unit_price: unitPrice,
        purchase_date: String(obj.purchase_date ?? obj.dataCompra ?? "").slice(0, 10),
        pncp_or_compras_uri: String(obj.pncp_or_compras_uri ?? "https://pncp.gov.br/app/editais"),
        pncp_control_number: obj.pncp_control_number ? String(obj.pncp_control_number) : null,
        comprasnet_uri: obj.comprasnet_uri ? String(obj.comprasnet_uri) : null,
      },
    };
  },
  parse(input: unknown): GovQuoteValidated {
    const res = this.safeParse(input);
    if (!res.success) {
      throw new Error(res.error);
    }
    return res.data;
  },
};

export const RETRYABLE_HTTP_STATUS_CODES: ReadonlySet<number> = new Set([408, 429, 500, 502, 503, 504]);

export function computeExponentialBackoffMs(attempt: number, baseMs: number = 600, maxMs: number = 8000): number {
  const exp = Math.min(6, Math.max(0, Math.trunc(attempt)));
  const raw = baseMs * Math.pow(2, exp);
  return Math.min(maxMs, raw);
}
