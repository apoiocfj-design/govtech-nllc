-- ============================================================================
-- SCHEMA OFICIAL POSTGRESQL 16+ - SISTEMA GOVTECH DE AUDITORIA E PREÇOS (NLLC)
-- ============================================================================
CREATE EXTENSION IF NOT EXISTS "uuid-ossp";
CREATE EXTENSION IF NOT EXISTS "pgcrypto";

CREATE TABLE organizations (
    id UUID PRIMARY KEY DEFAULT uuid_generate_v4(),
    uasg_code VARCHAR(6) UNIQUE NOT NULL,
    name VARCHAR(255) NOT NULL,
    cnpj VARCHAR(14) UNIQUE NOT NULL,
    uf CHAR(2) NOT NULL,
    esfera VARCHAR(20) DEFAULT 'FEDERAL',
    poder VARCHAR(20) DEFAULT 'EXECUTIVO',
    created_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

CREATE TABLE users (
    id UUID PRIMARY KEY DEFAULT uuid_generate_v4(),
    organization_id UUID NOT NULL REFERENCES organizations(id) ON DELETE RESTRICT,
    full_name VARCHAR(255) NOT NULL,
    email VARCHAR(255) UNIQUE NOT NULL,
    military_rank_or_role VARCHAR(100),
    system_role VARCHAR(30) NOT NULL CHECK (system_role IN ('PLANEJADOR', 'AGENTE_CONTRATACAO', 'AUDITOR_CONTROLE_INTERNO', 'AUTORIDADE_COMPETENTE')),
    created_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

CREATE TABLE procurement_processes (
    id UUID PRIMARY KEY DEFAULT uuid_generate_v4(),
    organization_id UUID NOT NULL REFERENCES organizations(id),
    created_by UUID NOT NULL REFERENCES users(id),
    nup_number VARCHAR(30) UNIQUE NOT NULL,
    notice_number VARCHAR(20),
    modality VARCHAR(40) NOT NULL CHECK (modality IN ('DISPENSA_ELETRONICA', 'PREGAO_ELETRONICO', 'INEXIGIBILIDADE')),
    legal_framing VARCHAR(40) NOT NULL CHECK (legal_framing IN ('ART_75_I', 'ART_75_II', 'ART_74', 'ART_28_I')),
    object_summary TEXT NOT NULL,
    total_estimated_value NUMERIC(15, 4) DEFAULT 0.0000,
    status VARCHAR(30) NOT NULL DEFAULT 'EM_ELABORACAO',
    created_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
    updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

CREATE TABLE artifact_documents (
    id UUID PRIMARY KEY DEFAULT uuid_generate_v4(),
    process_id UUID NOT NULL REFERENCES procurement_processes(id) ON DELETE CASCADE,
    doc_type VARCHAR(20) NOT NULL CHECK (doc_type IN ('DFD', 'ETP', 'TR', 'AVISO_DISPENSA', 'EDITAL')),
    version_number INTEGER NOT NULL DEFAULT 1,
    original_filename VARCHAR(255) NOT NULL,
    storage_path TEXT NOT NULL,
    file_hash_sha256 CHAR(64) NOT NULL,
    parsed_text_json JSONB,
    agu_template_used VARCHAR(120),
    compliance_score NUMERIC(5, 2),
    uploaded_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
    UNIQUE(process_id, doc_type, version_number)
);

CREATE TABLE legal_audit_findings (
    id UUID PRIMARY KEY DEFAULT uuid_generate_v4(),
    artifact_id UUID NOT NULL REFERENCES artifact_documents(id) ON DELETE CASCADE,
    clause_identifier VARCHAR(50),
    finding_category VARCHAR(25) NOT NULL CHECK (finding_category IN ('VULNERABILIDADE_JURIDICA', 'OMISSAO_OBRIGATORIA', 'ERRO_MATERIAL', 'ADEQUACAO_REDACIONAL')),
    severity VARCHAR(15) NOT NULL CHECK (severity IN ('CRITICA', 'ALTA', 'MEDIA', 'BAIXA')),
    user_original_text TEXT,
    agu_standard_suggestion TEXT NOT NULL,
    legal_foundation TEXT NOT NULL,
    rag_similarity_score NUMERIC(5, 4),
    source_chunk_id VARCHAR(100),
    status VARCHAR(20) DEFAULT 'PENDENTE',
    reviewer_justification TEXT,
    created_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

CREATE TABLE catalog_items (
    catmat_catser_code INTEGER PRIMARY KEY,
    item_type VARCHAR(15) NOT NULL CHECK (item_type IN ('MATERIAL', 'SERVICO')),
    group_code INTEGER,
    class_code INTEGER,
    pdm_code INTEGER,
    pdm_name VARCHAR(255),
    official_description TEXT NOT NULL,
    is_active BOOLEAN DEFAULT TRUE,
    last_synced_at TIMESTAMPTZ DEFAULT NOW()
);

CREATE TABLE demand_items (
    id UUID PRIMARY KEY DEFAULT uuid_generate_v4(),
    process_id UUID NOT NULL REFERENCES procurement_processes(id) ON DELETE CASCADE,
    item_number INTEGER NOT NULL,
    catmat_catser_code INTEGER NOT NULL REFERENCES catalog_items(catmat_catser_code),
    raw_demand_description TEXT,
    nlp_confidence_score NUMERIC(5, 4),
    supplemental_specs TEXT,
    supply_unit VARCHAR(20) NOT NULL,
    quantity NUMERIC(15, 4) NOT NULL CHECK (quantity > 0),
    estimated_unit_price NUMERIC(15, 4),
    estimated_total_price NUMERIC(15, 4) GENERATED ALWAYS AS (quantity * estimated_unit_price) STORED,
    UNIQUE(process_id, item_number)
);

CREATE TABLE price_research_sessions (
    id UUID PRIMARY KEY DEFAULT uuid_generate_v4(),
    demand_item_id UUID NOT NULL REFERENCES demand_items(id) ON DELETE CASCADE,
    executed_by UUID NOT NULL REFERENCES users(id),
    filter_uf CHAR(2),
    filter_supply_unit VARCHAR(20) NOT NULL,
    period_start DATE NOT NULL,
    period_end DATE NOT NULL,
    total_raw_quotes INTEGER NOT NULL,
    valid_quotes_count INTEGER NOT NULL,
    purged_outliers_count INTEGER NOT NULL,
    raw_mean NUMERIC(15, 4),
    raw_std_dev NUMERIC(15, 4),
    raw_cv_percent NUMERIC(8, 4),
    sanitized_min NUMERIC(15, 4),
    sanitized_mean NUMERIC(15, 4),
    sanitized_median NUMERIC(15, 4),
    sanitized_max NUMERIC(15, 4),
    sanitized_std_dev NUMERIC(15, 4),
    sanitized_cv_percent NUMERIC(8, 4),
    adopted_parameter VARCHAR(20) NOT NULL CHECK (adopted_parameter IN ('MEDIANA', 'MEDIA', 'MENOR_VALOR')),
    final_reference_price NUMERIC(15, 4) NOT NULL,
    outlier_method_used VARCHAR(50) NOT NULL DEFAULT 'DESVIO_PADRAO_E_IQR',
    methodology_justification TEXT NOT NULL,
    executed_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

CREATE TABLE price_quotes (
    id UUID PRIMARY KEY DEFAULT uuid_generate_v4(),
    session_id UUID NOT NULL REFERENCES price_research_sessions(id) ON DELETE CASCADE,
    api_source VARCHAR(30) NOT NULL CHECK (api_source IN ('COMPRAS_GOV_OAS31', 'PNCP_V26')),
    id_compra_item VARCHAR(30),
    purchase_number_year VARCHAR(20) NOT NULL,
    pncp_control_number VARCHAR(50),
    buyer_uasg VARCHAR(10) NOT NULL,
    buyer_org_name VARCHAR(255) NOT NULL,
    buyer_uf CHAR(2) NOT NULL,
    modality_code INTEGER,
    supplier_cnpj VARCHAR(20) NOT NULL,
    supplier_name VARCHAR(255) NOT NULL,
    brand VARCHAR(150),
    supply_unit VARCHAR(20) NOT NULL,
    quantity_homologated NUMERIC(15, 4) NOT NULL,
    unit_price NUMERIC(15, 4) NOT NULL,
    purchase_date DATE NOT NULL,
    pncp_or_compras_uri TEXT NOT NULL,
    is_outlier BOOLEAN NOT NULL DEFAULT FALSE,
    purge_status VARCHAR(30) NOT NULL DEFAULT 'VALIDA',
    purge_justification TEXT,
    raw_json_payload JSONB NOT NULL
);
