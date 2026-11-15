// ============================================================================
// CONTRATOS TYPESCRIPT DO FRONTEND (ESPELHO ESTRITO DO BACKEND FASTAPI PYTHON)
// ============================================================================

export type FindingCategory =
  | "VULNERABILIDADE_JURIDICA"
  | "OMISSAO_OBRIGATORIA"
  | "SUGESTAO_CORRECAO";

export type SeverityLevel = "CRITICA" | "ALTA" | "MEDIA" | "BAIXA";

export type FindingResolutionStatus =
  | "PENDENTE"
  | "ACEITO_APLICADO"
  | "JUSTIFICADO_MANTIDO";

export interface LegalAuditFinding {
  id: string;
  finding_type: FindingCategory;
  severity: SeverityLevel;
  clause_reference: string;
  section_title: string;
  user_document_excerpt: string | null;
  agu_standard_clause: string;
  legal_foundation: string;
  similarity_score: number;
  explanation: string;
  status: FindingResolutionStatus;
  reviewer_justification?: string;
}

export interface DocumentClauseBlock {
  clause_number: string;
  section_title: string;
  content: string;
  is_modified_by_agu?: boolean;
}

export interface RAGAuditResponse {
  process_nup: string;
  artifact_type: "DFD" | "ETP" | "TR" | "AVISO_DISPENSA" | "EDITAL";
  artifact_title: string;
  agu_template_version: string;
  total_user_clauses_parsed: number;
  agu_reference_clauses_checked: number;
  compliance_score: number;
  vulnerabilities_count: number;
  omissions_count: number;
  clauses: DocumentClauseBlock[];
  findings: LegalAuditFinding[];
}

// ============================================================================
// TIPOS DO MÓDULO 2: CRAWLER ANALÍTICO DE PREÇOS (IN SEGES/ME Nº 65/2021)
// ============================================================================

export type PurgeStatus =
  | "VALIDA"
  | "EXPURGADA_UNIDADE_DIVERGENTE"
  | "EXPURGADA_INEXEQUIVEL"
  | "EXPURGADA_ACIMA_LIMITE";

export interface AuditedQuote {
  id_compra_item: string;
  purchase_number_year: string;
  buyer_uasg: string;
  buyer_org_name: string;
  buyer_uf: string;
  supplier_cnpj: string;
  supplier_name: string;
  supply_unit: string;
  quantity_homologated: number;
  unit_price: number;
  purchase_date: string;
  pncp_or_compras_uri: string;
  brand?: string | null;
  pncp_control_number?: string | null;
  is_outlier: boolean;
  purge_status: PurgeStatus;
  purge_justification?: string | null;
}

export interface StatisticalMetrics {
  sample_size: number;
  min_price: number | null;
  max_price: number | null;
  mean: number | null;
  median: number | null;
  std_deviation: number | null;
  cv_percent: number | null;
  lower_bound_std: number | null;
  upper_bound_std: number | null;
}

export interface IN65ResearchReport {
  catmat_catser_code: number;
  item_description: string;
  target_supply_unit: string;
  demanded_quantity: number;
  cv_homogeneity_threshold_percent: number;
  std_dev_multiplier: number;
  raw_metrics: StatisticalMetrics;
  sanitized_metrics: StatisticalMetrics;
  is_homogeneous: boolean;
  has_minimum_three_quotes: boolean;
  adopted_parameter: "MEDIA" | "MEDIANA" | "MENOR_VALOR";
  final_estimated_unit_price: number;
  legal_methodology_justification: string;
  valid_quotes: AuditedQuote[];
  purged_quotes: AuditedQuote[];
}
