/**
 * OBSERVABILIDADE E TELEMETRIA ESTRUTURADA COM MASCARAMENTO DE DADOS SENSÍVEIS (LGPD / APPSEC)
 */

export interface StructuredAuditLog {
  readonly timestampUtc: string;
  readonly severity: "INFO" | "WARNING" | "ERROR";
  readonly service: string;
  readonly event: string;
  readonly latencyMs?: number;
  readonly details: Readonly<Record<string, unknown>>;
}

export function maskSensitiveData(input: string): string {
  return String(input || "")
    .replace(/\b(\d{3})\.\d{3}\.\d{3}-(\d{2})\b/g, "$1.***.***-$2") // CPF
    .replace(/(Bearer\s+)[A-Za-z0-9._-]+/gi, "$1[REDACTED]")        // JWT / Bearer
    .replace(/(api[_-]?key\s*[:=]\s*)[^\s&"']+/gi, "$1[REDACTED]"); // API Key
}

export function createStructuredLog(
  severity: "INFO" | "WARNING" | "ERROR",
  service: string,
  event: string,
  details: Record<string, unknown> = {},
  latencyMs?: number
): StructuredAuditLog {
  const sanitizedDetails: Record<string, unknown> = {};
  for (const [k, v] of Object.entries(details)) {
    sanitizedDetails[k] = typeof v === "string" ? maskSensitiveData(v) : v;
  }
  const base: StructuredAuditLog = {
    timestampUtc: new Date().toISOString(),
    severity,
    service,
    event: maskSensitiveData(event),
    details: sanitizedDetails,
    ...(latencyMs !== undefined ? { latencyMs } : {}),
  };
  return base;
}
