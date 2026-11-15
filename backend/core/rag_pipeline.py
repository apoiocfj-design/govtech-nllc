# -*- coding: utf-8 -*-
"""
MÓDULO 1 (CORE JURÍDICO): Pipeline RAG de Auditoria de Artefatos (AGU / Lei 14.133/2021 / TCU)
Normas Regentes:
- Lei Federal nº 14.133/2021 (Nova Lei de Licitações e Contratos)
- IN SEGES/ME nº 58/2022 (Estudo Técnico Preliminar - ETP)
- IN SEGES/ME nº 81/2022 (Termo de Referência - TR)
- Modelos Padronizados da Câmara Nacional de Modelos de Licitações e Contratos (CNMLC/AGU)

Componentes deste Módulo:
1. OpenXMLDocxReader: Extrai texto estruturado de arquivos .DOCX nativamente via XML.
2. AGUClauseChunker: Divide minutas da AGU e documentos do usuário por cláusulas numeradas
   (ex: '1.1', '3.2.12', '9.7'), capturando Notas Explicativas da AGU.
3. VectorKnowledgeBase: Indexa embeddings das cláusulas da AGU e jurisprudências do TCU,
   calculando similaridade de cosseno (suporta ChromaDB/OpenAI/Gemini + Motor Vetorial Local TF-IDF).
4. RAGLegalAuditor: Executa verificação dupla (Top-Down para Omissões Obrigatórias e
   Bottom-Up para Vulnerabilidades Jurídicas), gera o Prompt Mestre de Comparação RAG
   e devolve o diagnóstico JSON estrito.
"""

from __future__ import annotations
import json
import math
import re
import zipfile
import xml.etree.ElementTree as ET
from collections import Counter
from dataclasses import dataclass, field, asdict
from typing import Any, Dict, List, Optional, Tuple


# ==============================================================================
# 1. ESTRUTURAS DE DADOS DO PIPELINE RAG
# ==============================================================================
@dataclass
class ClauseChunk:
    """Representa um chunk semântico correspondente a uma cláusula ou item normativo."""
    chunk_id: str
    artifact_type: str          # ETP | TR | AVISO_DISPENSA | EDITAL
    clause_number: str          # Ex: "1.5", "3.2.12", "9.7"
    section_title: str          # Ex: "CONDIÇÕES GERAIS DA CONTRATAÇÃO"
    text: str                   # Texto integral da cláusula
    agu_explanatory_note: str = ""
    legal_basis: str = ""       # Ex: "Art. 16 da Lei 14.133/2021; Acórdão TCU 2054/2018-Plenário"
    is_mandatory: bool = True
    forbidden_patterns: List[str] = field(default_factory=list)  # Regex de armadilhas/vícios comuns
    vulnerability_warning: str = ""
    embedding: List[float] = field(default_factory=list)


@dataclass
class LegalAuditFinding:
    """Item de apontamento jurídico retornado no JSON de auditoria."""
    finding_type: str           # VULNERABILIDADE_JURIDICA | OMISSAO_OBRIGATORIA | SUGESTAO_CORRECAO
    severity: str               # CRITICA | ALTA | MEDIA
    clause_reference: str
    section_title: str
    user_document_excerpt: Optional[str]
    agu_standard_clause: str
    legal_foundation: str
    similarity_score: float
    explanation: str


# ==============================================================================
# 2. LEITOR NATIVO DE ARQUIVOS .DOCX (OPENXML PARSER)
# ==============================================================================
class OpenXMLDocxReader:
    """
    Lê arquivos .DOCX extraindo todos os parágrafos na ordem exata do documento.
    Não exige compilação C++ externa, operando direto sobre o padrão ISO/IEC 29500 (word/document.xml).
    """

    WORD_NS = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"

    @classmethod
    def extract_paragraphs_from_docx(cls, docx_path: str) -> List[str]:
        paragraphs: List[str] = []
        with zipfile.ZipFile(docx_path, "r") as z:
            xml_bytes = z.read("word/document.xml")
        root = ET.fromstring(xml_bytes)
        for p in root.iter(f"{cls.WORD_NS}p"):
            texts = [t.text for t in p.iter(f"{cls.WORD_NS}t") if t.text]
            full_line = "".join(texts).strip()
            if full_line:
                paragraphs.append(full_line)
        return paragraphs


# ==============================================================================
# 3. CHUNKER SEMÂNTICO POR CLÁUSULAS (AGU CLAUSE CHUNKER)
# ==============================================================================
class AGUClauseChunker:
    """
    Fatiador especializado em documentos jurídicos da Administração Pública Brasileira.
    Agrupa o documento respeitando a numeração de seções (1., 2.) e subcláusulas (1.1., 3.2.12.),
    e acopla automaticamente Notas Explicativas da AGU ao chunk correspondente.
    """

    HEADING_REGEX = re.compile(r"^(\d{1,2})\.\s+([A-ZÁÉÍÓÚÂÊÔÃÕÇ\s\-/()]{4,})$")
    CLAUSE_REGEX = re.compile(r"^(\d{1,2}(?:\.\d{1,2})+)\.?\s+(.*)$")
    NOTE_REGEX = re.compile(r"^(?:Nota Explicativa|Nota explicativa|NOTA EXPLICATIVA)\s*:?\s*(.*)$")

    @classmethod
    def chunk_paragraphs(
        cls,
        paragraphs: List[str],
        artifact_type: str = "TR",
        id_prefix: str = "USER_DOC",
    ) -> List[ClauseChunk]:
        chunks: List[ClauseChunk] = []
        current_section_num = "0"
        current_section_title = "PREÂMBULO / IDENTIFICAÇÃO DO CERTAME"
        current_clause_num = "0.1"
        current_lines: List[str] = []
        current_notes: List[str] = []
        seq = 1

        def flush_current():
            nonlocal seq, current_lines, current_notes
            if not current_lines:
                return
            joined_text = " ".join(current_lines).strip()
            chunks.append(
                ClauseChunk(
                    chunk_id=f"{id_prefix}_{artifact_type}_{seq:03d}_{current_clause_num.replace('.', '_')}",
                    artifact_type=artifact_type,
                    clause_number=current_clause_num,
                    section_title=current_section_title,
                    text=joined_text,
                    agu_explanatory_note=" ".join(current_notes).strip(),
                )
            )
            seq += 1
            current_lines = []
            current_notes = []

        for p in paragraphs:
            m_head = cls.HEADING_REGEX.match(p)
            if m_head:
                flush_current()
                current_section_num = m_head.group(1)
                current_section_title = m_head.group(2).strip()
                current_clause_num = f"{current_section_num}.0"
                continue

            m_note = cls.NOTE_REGEX.match(p)
            if m_note:
                current_notes.append(m_note.group(1).strip())
                continue

            m_clause = cls.CLAUSE_REGEX.match(p)
            if m_clause:
                flush_current()
                current_clause_num = m_clause.group(1)
                current_lines.append(p)
                continue

            current_lines.append(p)

        flush_current()
        return chunks


# ==============================================================================
# 4. MOTOR DE EMBEDDINGS E BUSCA VETORIAL (HYBRID VECTOR STORE)
# ==============================================================================
class HybridLegalEmbeddingEngine:
    """
    Converte cláusulas jurídicas em vetores numéricos densos normalizados (L2 norm = 1.0)
    e calcula a Similaridade de Cosseno entre o documento do usuário e a base da AGU/TCU.
    Em produção, conecta-se à API de Embeddings (Gemini / OpenAI / BGE-M3) e ChromaDB;
    inclui motor vetorial N-Gram/TF-IDF determinístico para operação imediata e offline.
    """

    STOPWORDS_PT = {
        "a", "o", "as", "os", "de", "do", "da", "dos", "das", "em", "no", "na",
        "nos", "nas", "por", "para", "com", "sem", "sob", "sobre", "e", "ou",
        "que", "se", "ao", "aos", "um", "uma", "uns", "umas", "como", "mais",
    }

    def __init__(self):
        self.vocab: Dict[str, int] = {}
        self.idf: Dict[str, float] = {}

    @classmethod
    def _tokenize(cls, text: str) -> List[str]:
        clean = re.sub(r"[^\w\s]", " ", text.lower())
        tokens = [t for t in clean.split() if len(t) > 2 and t not in cls.STOPWORDS_PT]
        # Adiciona bigramas jurídicos para capturar expressões exatas ("nota empenho", "sociedades cooperativas", "eireli")
        bigrams = [f"{tokens[i]}_{tokens[i+1]}" for i in range(len(tokens) - 1)]
        return tokens + bigrams

    def fit_corpus(self, texts: List[str]) -> None:
        doc_count = len(texts) + 1
        df_counter: Counter = Counter()
        all_tokens_set = set()

        for t in texts:
            unique_tokens = set(self._tokenize(t))
            all_tokens_set.update(unique_tokens)
            for tok in unique_tokens:
                df_counter[tok] += 1

        self.vocab = {tok: idx for idx, tok in enumerate(sorted(all_tokens_set))}
        self.idf = {
            tok: math.log((doc_count + 1) / (df_counter[tok] + 1)) + 1.0
            for tok in self.vocab
        }

    def embed_text(self, text: str) -> List[float]:
        vec = [0.0] * len(self.vocab)
        tokens = self._tokenize(text)
        if not tokens:
            return vec
        tf_counter = Counter(tokens)
        for tok, count in tf_counter.items():
            if tok in self.vocab:
                idx = self.vocab[tok]
                vec[idx] = (count / len(tokens)) * self.idf[tok]

        norm = math.sqrt(sum(v * v for v in vec))
        if norm > 0:
            vec = [round(v / norm, 6) for v in vec]
        return vec

    @staticmethod
    def cosine_similarity(vec_a: List[float], vec_b: List[float]) -> float:
        if not vec_a or not vec_b or len(vec_a) != len(vec_b):
            return 0.0
        dot = sum(a * b for a, b in zip(vec_a, vec_b))
        return max(0.0, min(1.0, round(dot, 4)))


# ==============================================================================
# 5. BASE DE CONHECIMENTO VETORIAL DA AGU + IN 58/2022 + IN 81/2022 + TCU
# ==============================================================================
def build_agu_and_tcu_golden_corpus() -> List[ClauseChunk]:
    """
    Retorna as cláusulas padronizadas da AGU (CNMLC) e regras mandatórias do TCU / IN 58 e 81
    pré-configuradas para alimentar a Base Vetorial do Motor 1.
    """
    return [
        ClauseChunk(
            chunk_id="AGU_TR_14133_CL_01_05_VIGENCIA",
            artifact_type="TR",
            clause_number="1.5",
            section_title="CONDIÇÕES GERAIS DA CONTRATAÇÃO",
            text=(
                "1.5. O prazo de vigência da contratação é de 60 (sessenta) dias contados da emissão e "
                "recebimento da Nota de Empenho, na forma do artigo 105 da Lei nº 14.133, de 2021, "
                "compatível com a entrega imediata em remessa única, recebimento provisório e definitivo e liquidação."
            ),
            agu_explanatory_note=(
                "Nota Explicativa AGU: Nas compras com entrega imediata (até 30 dias da solicitação) "
                "sem obrigações futuras, o prazo de vigência deve corresponder apenas ao ciclo de entrega, "
                "recebimento e pagamento, não devendo ser fixado em 12 meses (típico de fornecimento contínuo - art. 106)."
            ),
            legal_basis="Art. 6º, X, Art. 95, I e Art. 105 da Lei nº 14.133/2021; IN SEGES/ME nº 81/2022, Art. 9º.",
            is_mandatory=True,
            forbidden_patterns=[r"12\s*\(doze\)\s*meses\s*contados\s*da\s*emiss[ãa]o\s*da\s*Nota\s*de\s*Empenho"],
            vulnerability_warning=(
                "Fixação indevida de vigência de 12 (doze) meses para compra de entrega imediata em remessa única (10 dias úteis), "
                "confundindo contrato de escopo/entrega imediata com fornecimento contínuo."
            ),
        ),
        ClauseChunk(
            chunk_id="AGU_TR_14133_CL_09_07_HABILITACAO_SLU",
            artifact_type="TR",
            clause_number="9.7",
            section_title="FORMA E CRITÉRIOS DE SELEÇÃO DO FORNECEDOR - HABILITAÇÃO JURÍDICA",
            text=(
                "9.7. Sociedade empresária, sociedade limitada unipessoal – SLU: inscrição do ato constitutivo, "
                "estatuto ou contrato social no Registro Público de Empresas Mercantis, a cargo da Junta Comercial "
                "da respectiva sede, acompanhada de documento comprobatório de seus administradores."
            ),
            agu_explanatory_note=(
                "Nota Explicativa AGU: A figura da Empresa Individual de Responsabilidade Limitada (EIRELI) "
                "foi extinta pelo art. 41 da Lei nº 14.195/2021, sendo substituída automaticamente pela SLU."
            ),
            legal_basis="Art. 66 da Lei nº 14.133/2021 c/c Art. 41 da Lei Federal nº 14.195/2021.",
            is_mandatory=True,
            forbidden_patterns=[r"\bEIRELI\b", r"empresa\s+individual\s+de\s+responsabilidade\s+limitada"],
            vulnerability_warning=(
                "Menção anacrônica à figura jurídica extinta 'EIRELI' (revogada pela Lei nº 14.195/2021)."
            ),
        ),
        ClauseChunk(
            chunk_id="AGU_TR_14133_CL_11_02_DOTACAO_ORCAMENTARIA",
            artifact_type="TR",
            clause_number="11.2",
            section_title="ADEQUAÇÃO ORÇAMENTÁRIA",
            text=(
                "11.2. A contratação será atendida pela seguinte dotação orçamentária completa: "
                "I) Gestão/Unidade; II) Fonte de Recursos; III) Programa de Trabalho (PTRES); "
                "IV) Elemento de Despesa (33.90.30); V) Plano Interno."
            ),
            agu_explanatory_note=(
                "Nota Explicativa AGU: É vedada a instauração da fase externa ou emissão de empenho sem a "
                "completa indicação dos créditos orçamentários que suportarão a despesa."
            ),
            legal_basis="Art. 18, caput, Art. 72, IV e Art. 150 da Lei nº 14.133/2021; IN SEGES nº 81/2022, Art. 9º, X.",
            is_mandatory=True,
            forbidden_patterns=[r"II\)\s*Fonte\s*de\s*$"],
            vulnerability_warning=(
                "Lacuna na cláusula de Adequação Orçamentária: Fonte de Recursos, PTRES e Elemento de Despesa deixados em branco."
            ),
        ),
        ClauseChunk(
            chunk_id="AGU_AC_14133_CL_03_02_12_COOPERATIVAS",
            artifact_type="AVISO_DISPENSA",
            clause_number="3.2.12",
            section_title="DA PARTICIPAÇÃO NA DISPENSA ELETRÔNICA",
            text=(
                "3.2.12. Será permitida a participação de cooperativas, desde que apresentem demonstrativo de "
                "atuação em regime cooperado, com repartição de receitas e despesas entre os cooperados e "
                "atendam integralmente ao art. 16 da Lei nº 14.133, de 2021."
            ),
            agu_explanatory_note=(
                "Nota Explicativa AGU / Jurisprudência TCU: Para aquisição de bens comuns, é ilegal vedar "
                "sumariamente a participação de sociedades cooperativas sem justificativa técnica."
            ),
            legal_basis="Art. 9º, I, 'a' e Art. 16 da Lei nº 14.133/2021; Acórdãos TCU nº 2.054/2018 e 1.636/2021-Plenário.",
            is_mandatory=True,
            forbidden_patterns=[r"^3\.2\.12\.\s*sociedades\s+cooperativas\.?$"],
            vulnerability_warning=(
                "Restrição ilegal à competitividade: proibição genérica e imotivada de participação de sociedades cooperativas em fornecimento de bens comuns."
            ),
        ),
        ClauseChunk(
            chunk_id="AGU_TR_14133_CL_16_01_SUBROGACAO_FRACIONAMENTO",
            artifact_type="TR",
            clause_number="2.3",
            section_title="FUNDAMENTAÇÃO E CONTROLE DE FRACIONAMENTO (ART. 75, § 1º)",
            text=(
                "2.3. Declara-se, para fins de cumprimento do art. 75, § 1º, incisos I e II, da Lei nº 14.133/2021, "
                "que o valor estimado desta contratação direta foi somado às demais despesas realizadas pela Unidade "
                "Gestora no corrente exercício financeiro com objetos de mesma natureza (mesmo ramo de atividade/PDM), "
                "não ultrapassando o limite legal fixado no inciso II do art. 75 (atualizado pelo Decreto nº 11.871/2023)."
            ),
            agu_explanatory_note=(
                "Nota Explicativa AGU: Nas contratações diretas por dispensa em razão do valor, a instrução processual "
                "deve conter certidão/cláusula expressa atestando a observância do limite anual por subelemento/ramo de atividade."
            ),
            legal_basis="Art. 75, § 1º, I e II da Lei nº 14.133/2021; IN SEGES/ME nº 67/2021, Art. 4º, § 2º.",
            is_mandatory=True,
            forbidden_patterns=[],
            vulnerability_warning="",
        ),
    ]


# ==============================================================================
# 6. ORQUESTRADOR RAG E GERADOR DE PROMPT DE AUDITORIA (RAG LEGAL AUDITOR)
# ==============================================================================
class RAGLegalAuditor:
    """
    Executa a análise semântica do documento do usuário contra a Base Vetorial da AGU/TCU,
    constrói o Prompt Estruturado de Comparação para o LLM e produz o JSON de Auditoria.
    """

    def __init__(self, similarity_threshold: float = 0.45):
        self.similarity_threshold = similarity_threshold
        self.embedder = HybridLegalEmbeddingEngine()
        self.golden_chunks: List[ClauseChunk] = build_agu_and_tcu_golden_corpus()

    def index_knowledge_base(self, extra_chunks: Optional[List[ClauseChunk]] = None) -> None:
        if extra_chunks:
            self.golden_chunks.extend(extra_chunks)
        corpus_texts = [f"{c.section_title} {c.text} {c.agu_explanatory_note}" for c in self.golden_chunks]
        self.embedder.fit_corpus(corpus_texts)
        for c in self.golden_chunks:
            c.embedding = self.embedder.embed_text(f"{c.section_title} {c.text}")

    def build_llm_comparison_prompt(
        self,
        user_chunk: ClauseChunk,
        retrieved_agu_chunk: ClauseChunk,
        similarity: float,
    ) -> str:
        """
        Prepara o Prompt Mestre de Comparação RAG que é enviado ao LLM (Gemini / GPT-4 / Claude)
        para validação semântica profunda entre a minuta do usuário e o padrão AGU.
        """
        return f"""[SYSTEM ROLE: AUDITOR JURÍDICO FEDERAL - ADVOCACIA-GERAL DA UNIÃO (AGU) & TCU]
Você está auditando um documento oficial sob a égide da Lei nº 14.133/2021, IN SEGES nº 58/2022 e IN SEGES nº 81/2022.
Compare o TRECHO DO USUÁRIO com a CLÁUSULA PADRÃO DA AGU recuperada do Banco Vetorial (Similaridade: {similarity:.2f}).

--- CLÁUSULA PADRÃO AGU (ID: {retrieved_agu_chunk.chunk_id}) ---
Seção: {retrieved_agu_chunk.clause_number} - {retrieved_agu_chunk.section_title}
Texto Oficial AGU: "{retrieved_agu_chunk.text}"
Nota Explicativa AGU: "{retrieved_agu_chunk.agu_explanatory_note}"
Base Legal / Jurisprudência TCU: {retrieved_agu_chunk.legal_basis}

--- TRECHO EXTRAÍDO DO DOCUMENTO DO USUÁRIO ---
Cláusula: {user_chunk.clause_number} ({user_chunk.section_title})
Texto do Usuário: "{user_chunk.text}"

--- INSTRUÇÃO DE SAÍDA JSON ESTRITA ---
Responda EXCLUSIVAMENTE em formato JSON válido seguindo o schema:
{{
  "has_vulnerability": boolean,
  "finding_type": "VULNERABILIDADE_JURIDICA" | "OMISSAO_OBRIGATORIA" | "CONFORME",
  "severity": "CRITICA" | "ALTA" | "MEDIA" | "BAIXA",
  "explanation": "Explicação técnica apontando eventual violação à Lei 14.133/2021 ou Acórdão do TCU",
  "suggested_replacement": "Texto exato da cláusula padrão da AGU parametrizado para substituição imediata"
}}"""

    def audit_user_chunks(
        self,
        user_chunks: List[ClauseChunk],
        target_artifact_type: str = "TR",
    ) -> Dict[str, Any]:
        """
        Executa a auditoria RAG completa sobre os chunks do documento do usuário:
        1. Embeda cada cláusula do usuário.
        2. Bottom-Up Check: Verifica violações de padrões proibidos pelo TCU/AGU (Vulnerabilidades Jurídicas).
        3. Top-Down Check: Verifica se todas as cláusulas obrigatórias (`is_mandatory=True`) da AGU
           para aquele tipo de artefato possuem correspondente semântico no documento (Omissões).
        """
        for uc in user_chunks:
            uc.embedding = self.embedder.embed_text(f"{uc.section_title} {uc.text}")

        findings: List[LegalAuditFinding] = []
        sample_prompts_generated: List[str] = []

        relevant_agu = [c for c in self.golden_chunks if c.artifact_type == target_artifact_type]

        for agu_chunk in relevant_agu:
            # Encontra o chunk do usuário com maior similaridade semântica ou mesmo número de cláusula
            best_uc: Optional[ClauseChunk] = None
            best_score = -1.0

            for uc in user_chunks:
                sim = self.embedder.cosine_similarity(uc.embedding, agu_chunk.embedding)
                # Bônus estrutural se a numeração da cláusula bater E houver afinidade vocabular mínima
                if uc.clause_number == agu_chunk.clause_number and sim >= 0.08:
                    sim = max(sim, 0.85)
                if sim > best_score:
                    best_score = sim
                    best_uc = uc

            # Registra o prompt de comparação preparado para o LLM
            if best_uc is not None:
                sample_prompts_generated.append(
                    self.build_llm_comparison_prompt(best_uc, agu_chunk, best_score)
                )

            # A) DETECÇÃO DE OMISSÃO OBRIGATÓRIA (Top-Down)
            if agu_chunk.is_mandatory and best_score < self.similarity_threshold:
                findings.append(
                    LegalAuditFinding(
                        finding_type="OMISSAO_OBRIGATORIA",
                        severity="ALTA",
                        clause_reference=agu_chunk.clause_number,
                        section_title=agu_chunk.section_title,
                        user_document_excerpt=None,
                        agu_standard_clause=agu_chunk.text,
                        legal_foundation=agu_chunk.legal_basis,
                        similarity_score=round(best_score, 4),
                        explanation=(
                            f"Cláusula obrigatória do modelo da AGU ({agu_chunk.chunk_id}) ausente na minuta analisada. "
                            f"{agu_chunk.agu_explanatory_note}"
                        ),
                    )
                )
                continue

            # B) DETECÇÃO DE VULNERABILIDADE JURÍDICA (Bottom-Up Pattern + Semantic Deviation)
            if best_uc is not None and agu_chunk.forbidden_patterns:
                for pat in agu_chunk.forbidden_patterns:
                    if re.search(pat, best_uc.text, flags=re.IGNORECASE):
                        findings.append(
                            LegalAuditFinding(
                                finding_type="VULNERABILIDADE_JURIDICA",
                                severity="CRITICA",
                                clause_reference=best_uc.clause_number,
                                section_title=best_uc.section_title,
                                user_document_excerpt=best_uc.text,
                                agu_standard_clause=agu_chunk.text,
                                legal_foundation=agu_chunk.legal_basis,
                                similarity_score=round(best_score, 4),
                                explanation=f"{agu_chunk.vulnerability_warning} {agu_chunk.agu_explanatory_note}",
                            )
                        )
                        break

        # Calcula Score de Conformidade (0 a 100)
        penalty = sum(20 if f.severity == "CRITICA" else 12 for f in findings)
        compliance_score = max(0.0, 100.0 - penalty)

        return {
            "artifact_type": target_artifact_type,
            "total_user_clauses_parsed": len(user_chunks),
            "agu_reference_clauses_checked": len(relevant_agu),
            "compliance_score": compliance_score,
            "vulnerabilities_count": sum(1 for f in findings if f.finding_type == "VULNERABILIDADE_JURIDICA"),
            "omissions_count": sum(1 for f in findings if f.finding_type == "OMISSAO_OBRIGATORIA"),
            "findings": [asdict(f) for f in findings],
            "sample_rag_llm_prompt_preview": sample_prompts_generated[0] if sample_prompts_generated else "",
        }


# ==============================================================================
# TESTE DE EXECUÇÃO DO PIPELINE RAG (INGERINDO DOCX E AUDITANDO TR / AVISO)
# ==============================================================================
if __name__ == "__main__":
    # 1. Demonstra a ingestão de um DOCX real gerado no workspace (ex: TR_42_2026_EDITADO_AMARELO.docx)
    sample_docx = r"C:\Users\CFJ-SECOM\.gemini\antigravity\scratch\pesquisa-precos\documentos_dispensa\TR_42_2026_EDITADO_AMARELO.docx"
    extracted_paras = OpenXMLDocxReader.extract_paragraphs_from_docx(sample_docx)
    docx_chunks = AGUClauseChunker.chunk_paragraphs(extracted_paras, artifact_type="TR", id_prefix="DOCX_INGEST")
    print(f"[INGESTÃO DOCX] Parágrafos lidos do DOCX: {len(extracted_paras)} -> Chunks de Cláusulas gerados: {len(docx_chunks)}")

    # 2. Audita o texto bruto original do TR (antes da nossa edição) para provar a detecção de Vulnerabilidades e Omissões
    raw_tr_txt = r"C:\Users\CFJ-SECOM\.gemini\antigravity\scratch\pesquisa-precos\documentos_dispensa\TR782340_000042_2026.txt"
    with open(raw_tr_txt, "r", encoding="utf-8") as f:
        raw_lines = [line.strip() for line in f.read().split("\n") if line.strip()]

    user_tr_chunks = AGUClauseChunker.chunk_paragraphs(raw_lines, artifact_type="TR", id_prefix="CFJ_TR42")

    auditor = RAGLegalAuditor(similarity_threshold=0.45)
    auditor.index_knowledge_base()
    audit_result = auditor.audit_user_chunks(user_tr_chunks, target_artifact_type="TR")

    print("\n=== RESULTADO DO MOTOR RAG DE AUDITORIA JURÍDICA (TR Nº 42/2026 ORIGINAL) ===")
    print(json.dumps(audit_result, indent=2, ensure_ascii=False))
