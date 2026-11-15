"use client";

import React, { useState, useMemo } from "react";
import {
  RAGAuditResponse,
  LegalAuditFinding,
  DocumentClauseBlock,
} from "../types/govtech";

interface TRLegalReviewerProps {
  initialAuditData: RAGAuditResponse;
  onExportSanitizedDocx?: (updatedClauses: DocumentClauseBlock[]) => void;
}

export const TRLegalReviewer: React.FC<TRLegalReviewerProps> = ({
  initialAuditData,
  onExportSanitizedDocx,
}) => {
  const [clauses, setClauses] = useState<DocumentClauseBlock[]>(
    initialAuditData.clauses
  );
  const [findings, setFindings] = useState<LegalAuditFinding[]>(
    initialAuditData.findings
  );
  const [selectedFindingId, setSelectedFindingId] = useState<string | null>(
    initialAuditData.findings[0]?.id ?? null
  );
  const [filterType, setFilterType] = useState<"ALL" | "VULNERABILIDADE_JURIDICA" | "OMISSAO_OBRIGATORIA">("ALL");

  // Recalcula o Score de Conformidade dinamicamente conforme o usuário aplica as correções da AGU
  const dynamicScore = useMemo(() => {
    const pending = findings.filter((f) => f.status === "PENDENTE");
    const penalty = pending.reduce(
      (acc, f) => acc + (f.severity === "CRITICA" ? 20 : 12),
      0
    );
    return Math.max(0, 100 - penalty);
  }, [findings]);

  // Aplica com 1 clique a Cláusula Padrão da AGU no documento e marca em amarelo
  const handleApplyAGUClause = (finding: LegalAuditFinding) => {
    setClauses((prev) => {
      const exists = prev.some((c) => c.clause_number === finding.clause_reference);
      if (exists) {
        return prev.map((c) =>
          c.clause_number === finding.clause_reference
            ? {
                ...c,
                content: finding.agu_standard_clause,
                is_modified_by_agu: true,
              }
            : c
        );
      } else {
        // Caso seja uma OMISSÃO OBRIGATÓRIA, insere a cláusula faltante no documento
        return [
          ...prev,
          {
            clause_number: finding.clause_reference,
            section_title: finding.section_title,
            content: finding.agu_standard_clause,
            is_modified_by_agu: true,
          },
        ].sort((a, b) => a.clause_number.localeCompare(b.clause_number, undefined, { numeric: true }));
      }
    });

    setFindings((prev) =>
      prev.map((f) =>
        f.id === finding.id ? { ...f, status: "ACEITO_APLICADO" } : f
      )
    );
  };

  const filteredFindings = findings.filter((f) =>
    filterType === "ALL" ? true : f.finding_type === filterType
  );

  return (
    <div className="min-h-screen bg-slate-50 text-slate-900 flex flex-col">
      {/* BARRA SUPERIOR DE COMANDO E SCORE DE COMPLIANCE */}
      <header className="bg-slate-900 text-white px-6 py-4 shadow-md flex flex-wrap items-center justify-between gap-4 border-b border-slate-700">
        <div>
          <div className="flex items-center gap-2 text-xs uppercase tracking-wider text-amber-400 font-semibold">
            <span>Módulo 1 • Auditoria Jurídica RAG (CNMLC / AGU & TCU)</span>
            <span>•</span>
            <span>NUP: {initialAuditData.process_nup}</span>
          </div>
          <h1 className="text-xl font-bold mt-1">{initialAuditData.artifact_title}</h1>
          <p className="text-xs text-slate-400">
            Matriz Vetorial de Referência: {initialAuditData.agu_template_version} •{" "}
            {initialAuditData.total_user_clauses_parsed} cláusulas analisadas
          </p>
        </div>

        <div className="flex items-center gap-6">
          {/* Medidor de Conformidade */}
          <div className="bg-slate-800 px-4 py-2 rounded-lg border border-slate-700 flex items-center gap-3">
            <div>
              <span className="block text-[11px] text-slate-400 uppercase">
                Índice de Conformidade NLLC
              </span>
              <span
                className={`text-2xl font-extrabold ${
                  dynamicScore >= 90
                    ? "text-emerald-400"
                    : dynamicScore >= 70
                    ? "text-amber-400"
                    : "text-rose-400"
                }`}
              >
                {dynamicScore.toFixed(0)}%
              </span>
            </div>
          </div>

          <button
            onClick={() => onExportSanitizedDocx?.(clauses)}
            className="bg-amber-400 hover:bg-amber-300 text-slate-950 font-bold px-4 py-2.5 rounded-lg text-sm shadow transition flex items-center gap-2"
          >
            <span>Exportar .DOCX Marcado em Amarelo</span>
          </button>
        </div>
      </header>

      {/* CORPO SPLIT-SCREEN: DOCUMENTO (ESQUERDA) vs DIAGNÓSTICO RAG AGU/TCU (DIREITA) */}
      <div className="flex-1 grid grid-cols-1 lg:grid-cols-12 gap-6 p-6 max-w-[1600px] w-full mx-auto">
        {/* COLUNA ESQUERDA (7 COLUNAS): VISUALIZADOR ESTILO "REVISOR DE TEXTO" */}
        <section className="lg:col-span-7 bg-white rounded-xl shadow-sm border border-slate-200 flex flex-col overflow-hidden">
          <div className="bg-slate-100 px-5 py-3 border-b border-slate-200 flex items-center justify-between">
            <span className="font-semibold text-sm text-slate-700">
              Visualizador da Minuta (Grifos Automáticos por Cláusula)
            </span>
            <div className="flex items-center gap-3 text-xs">
              <span className="flex items-center gap-1">
                <span className="w-3 h-3 rounded-full bg-rose-500 inline-block"></span>
                Vulnerabilidade TCU
              </span>
              <span className="flex items-center gap-1">
                <span className="w-3 h-3 rounded-full bg-yellow-300 border border-amber-500 inline-block"></span>
                Editado / Saneado AGU
              </span>
            </div>
          </div>

          <div className="p-6 space-y-4 overflow-y-auto max-h-[75vh] font-serif leading-relaxed text-sm">
            {clauses.map((clause) => {
              const relatedFinding = findings.find(
                (f) => f.clause_reference === clause.clause_reference || f.clause_reference === clause.clause_number
              );
              const hasPendingError =
                relatedFinding && relatedFinding.status === "PENDENTE";

              return (
                <div
                  key={clause.clause_number}
                  onClick={() => relatedFinding && setSelectedFindingId(relatedFinding.id)}
                  className={`p-3.5 rounded-lg border transition cursor-pointer ${
                    clause.is_modified_by_agu
                      ? "bg-yellow-200/80 border-amber-400 text-slate-950 font-medium shadow-sm"
                      : hasPendingError
                      ? "bg-rose-50 border-rose-300 hover:bg-rose-100/70"
                      : "bg-white border-transparent hover:border-slate-200"
                  }`}
                >
                  <div className="flex items-center justify-between mb-1 font-sans">
                    <span className="text-xs font-bold uppercase tracking-wide text-slate-500">
                      Cláusula {clause.clause_number} — {clause.section_title}
                    </span>
                    {clause.is_modified_by_agu && (
                      <span className="text-[11px] bg-amber-500 text-slate-950 font-bold px-2 py-0.5 rounded">
                        CORRIGIDO PELA MINUTA AGU (DESTAQUE AMARELO)
                      </span>
                    )}
                    {hasPendingError && (
                      <span className="text-[11px] bg-rose-600 text-white font-bold px-2 py-0.5 rounded">
                        {relatedFinding.finding_type === "VULNERABILIDADE_JURIDICA"
                          ? "VIOLAÇÃO LEI 14.133 / TCU"
                          : "REQUER REVISÃO"}
                      </span>
                    )}
                  </div>
                  <p className="text-slate-800">{clause.content}</p>
                </div>
              );
            })}
          </div>
        </section>

        {/* COLUNA DIREITA (5 COLUNAS): PAINEL DE APONTAMENTOS RAG (VULNERABILIDADES + OMISSÕES) */}
        <aside className="lg:col-span-5 flex flex-col gap-4">
          {/* Filtros Rápidos */}
          <div className="bg-white p-3 rounded-xl shadow-sm border border-slate-200 flex items-center gap-2">
            <button
              onClick={() => setFilterType("ALL")}
              className={`px-3 py-1.5 rounded-lg text-xs font-bold transition ${
                filterType === "ALL"
                  ? "bg-slate-900 text-white"
                  : "bg-slate-100 text-slate-600 hover:bg-slate-200"
              }`}
            >
              Todos ({findings.length})
            </button>
            <button
              onClick={() => setFilterType("VULNERABILIDADE_JURIDICA")}
              className={`px-3 py-1.5 rounded-lg text-xs font-bold transition ${
                filterType === "VULNERABILIDADE_JURIDICA"
                  ? "bg-rose-600 text-white"
                  : "bg-rose-50 text-rose-700 hover:bg-rose-100"
              }`}
            >
              Vulnerabilidades ({findings.filter((f) => f.finding_type === "VULNERABILIDADE_JURIDICA").length})
            </button>
            <button
              onClick={() => setFilterType("OMISSAO_OBRIGATORIA")}
              className={`px-3 py-1.5 rounded-lg text-xs font-bold transition ${
                filterType === "OMISSAO_OBRIGATORIA"
                  ? "bg-amber-600 text-white"
                  : "bg-amber-50 text-amber-800 hover:bg-amber-100"
              }`}
            >
              Omissões AGU ({findings.filter((f) => f.finding_type === "OMISSAO_OBRIGATORIA").length})
            </button>
          </div>

          {/* Lista de Cartões de Auditoria */}
          <div className="space-y-4 overflow-y-auto max-h-[70vh] pr-1">
            {filteredFindings.map((finding) => {
              const isSelected = finding.id === selectedFindingId;
              const isResolved = finding.status === "ACEITO_APLICADO";

              return (
                <div
                  key={finding.id}
                  onClick={() => setSelectedFindingId(finding.id)}
                  className={`bg-white rounded-xl p-5 border transition shadow-sm ${
                    isSelected
                      ? "ring-2 ring-slate-900 border-slate-900"
                      : "border-slate-200"
                  }`}
                >
                  <div className="flex items-center justify-between gap-2 mb-2">
                    <span
                      className={`text-[11px] font-extrabold px-2.5 py-0.5 rounded uppercase ${
                        finding.finding_type === "VULNERABILIDADE_JURIDICA"
                          ? "bg-rose-100 text-rose-800 border border-rose-300"
                          : "bg-amber-100 text-amber-900 border border-amber-300"
                      }`}
                    >
                      {finding.finding_type === "VULNERABILIDADE_JURIDICA"
                        ? "Vulnerabilidade Jurídica"
                        : "Omissão de Cláusula Obrigatória"}
                    </span>
                    <span className="text-xs font-mono text-slate-500">
                      Item {finding.clause_reference} • Similaridade RAG:{" "}
                      {(finding.similarity_score * 100).toFixed(0)}%
                    </span>
                  </div>

                  <h3 className="font-bold text-sm text-slate-900 mb-1">
                    {finding.section_title}
                  </h3>
                  <p className="text-xs text-slate-600 mb-3 leading-relaxed">
                    {finding.explanation}
                  </p>

                  {/* Base Legal / Acórdão TCU */}
                  <div className="bg-slate-50 p-2.5 rounded border border-slate-200 mb-3 text-xs">
                    <span className="font-bold text-slate-700 block">
                      Fundamento Legal & Jurisprudência TCU:
                    </span>
                    <span className="text-slate-600">{finding.legal_foundation}</span>
                  </div>

                  {/* Sugestão Oficial AGU */}
                  <div className="bg-yellow-100/90 border border-amber-400 p-3 rounded-lg mb-3 text-xs">
                    <span className="font-bold text-amber-950 block mb-1 uppercase tracking-wider text-[10px]">
                      Redação Oficial Padronizada da AGU (Para Substituição):
                    </span>
                    <p className="text-slate-900 font-medium leading-relaxed">
                      “{finding.agu_standard_clause}”
                    </p>
                  </div>

                  <div className="flex items-center justify-end gap-2">
                    {isResolved ? (
                      <span className="text-xs font-bold text-emerald-700 bg-emerald-50 border border-emerald-200 px-3 py-1.5 rounded-lg">
                        ✓ Cláusula AGU Aplicada em Amarelo no Documento
                      </span>
                    ) : (
                      <button
                        onClick={(e) => {
                          e.stopPropagation();
                          handleApplyAGUClause(finding);
                        }}
                        className="bg-slate-900 hover:bg-slate-800 text-white text-xs font-bold px-4 py-2 rounded-lg transition shadow-sm"
                      >
                        Substituir pela Cláusula Padrão da AGU
                      </button>
                    )}
                  </div>
                </div>
              );
            })}
          </div>
        </aside>
      </div>
    </div>
  );
};
