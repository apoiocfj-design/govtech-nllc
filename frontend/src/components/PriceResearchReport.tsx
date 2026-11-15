"use client";

import React, { useState } from "react";
import { IN65ResearchReport } from "../types/govtech";

interface PriceResearchReportProps {
  report: IN65ResearchReport;
  onExportAuditPdf?: () => void;
}

export const PriceResearchReport: React.FC<PriceResearchReportProps> = ({
  report,
  onExportAuditPdf,
}) => {
  const [activeTab, setActiveTab] = useState<"VALIDAS" | "EXPURGADAS">("VALIDAS");

  const totalEstimatedGlobal =
    report.demanded_quantity * report.final_estimated_unit_price;

  return (
    <div className="min-h-screen bg-slate-50 text-slate-900 p-6">
      <div className="max-w-7xl mx-auto space-y-6">
        {/* CABEÇALHO DO RELATÓRIO E BOTÃO DE EXPORTAÇÃO PDF AUDITÁVEL */}
        <header className="bg-slate-900 text-white p-6 rounded-2xl shadow-md flex flex-wrap items-center justify-between gap-4">
          <div>
            <div className="flex items-center gap-2 text-xs font-bold uppercase tracking-wider text-emerald-400">
              <span>Módulo 2 • Crawler Analítico IN SEGES/ME nº 65/2021</span>
              <span>•</span>
              <span>API Compras.gov.br OAS 3.1 & PNCP v2.6</span>
            </div>
            <h1 className="text-xl font-bold mt-1">
              CATMAT {report.catmat_catser_code} — {report.item_description}
            </h1>
            <p className="text-xs text-slate-400 mt-1">
              Unidade Padronizada: <strong>{report.target_supply_unit}</strong> •
              Quantidade Demandada: <strong>{report.demanded_quantity}</strong>
            </p>
          </div>

          <button
            onClick={onExportAuditPdf}
            className="bg-emerald-500 hover:bg-emerald-400 text-slate-950 font-extrabold px-5 py-3 rounded-xl text-sm shadow transition"
          >
            Exportar Mapa Comparativo & Trilha PNCP (PDF)
          </button>
        </header>

        {/* CARTÕES DE MÉTRICAS ESTATÍSTICAS (SANEAMENTO IN 65/2021) */}
        <section className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-5 gap-4">
          <div className="bg-white p-5 rounded-xl border border-slate-200 shadow-sm">
            <span className="text-xs font-semibold uppercase text-slate-500">
              Coeficiente de Variação (CV)
            </span>
            <div className="mt-2 flex items-baseline gap-2">
              <span
                className={`text-2xl font-extrabold ${
                  report.is_homogeneous ? "text-emerald-600" : "text-amber-600"
                }`}
              >
                {report.sanitized_metrics.cv_percent?.toFixed(2)}%
              </span>
              <span className="text-xs text-slate-400 line-through">
                Bruto: {report.raw_metrics.cv_percent?.toFixed(1)}%
              </span>
            </div>
            <span className="mt-2 inline-block text-[11px] font-bold px-2 py-0.5 rounded bg-emerald-50 text-emerald-800 border border-emerald-200">
              {report.is_homogeneous
                ? "Série Homogênea (≤ 25%)"
                : "Série Heterogênea (> 25%)"}
            </span>
          </div>

          <div className="bg-white p-5 rounded-xl border border-slate-200 shadow-sm">
            <span className="text-xs font-semibold uppercase text-slate-500">
              Média Saneada (μ)
            </span>
            <div className="mt-2 text-2xl font-bold text-slate-900">
              R$ {report.sanitized_metrics.mean?.toFixed(2)}
            </div>
            <span className="text-xs text-slate-500 mt-1 block">
              Desvio Padrão (s): R$ {report.sanitized_metrics.std_deviation?.toFixed(2)}
            </span>
          </div>

          <div className="bg-white p-5 rounded-xl border border-slate-200 shadow-sm">
            <span className="text-xs font-semibold uppercase text-slate-500">
              Mediana Saneada (Md)
            </span>
            <div className="mt-2 text-2xl font-bold text-slate-900">
              R$ {report.sanitized_metrics.median?.toFixed(2)}
            </div>
            <span className="text-xs text-slate-500 mt-1 block">
              Faixa: R$ {report.sanitized_metrics.min_price?.toFixed(2)} a R${" "}
              {report.sanitized_metrics.max_price?.toFixed(2)}
            </span>
          </div>

          <div className="bg-white p-5 rounded-xl border border-slate-200 shadow-sm">
            <span className="text-xs font-semibold uppercase text-slate-500">
              Amostras Auditadas
            </span>
            <div className="mt-2 text-2xl font-bold text-slate-900">
              {report.sanitized_metrics.sample_size} Válidas
            </div>
            <span className="text-xs text-rose-600 font-semibold mt-1 block">
              {report.purged_quotes.length} Outliers Expurgados
            </span>
          </div>

          <div className="bg-slate-900 text-white p-5 rounded-xl shadow-sm border border-slate-800">
            <span className="text-xs font-bold uppercase text-amber-400">
              Valor Estimado Final ({report.adopted_parameter})
            </span>
            <div className="mt-1 text-2xl font-extrabold text-white">
              R$ {report.final_estimated_unit_price.toFixed(2)}{" "}
              <span className="text-xs font-normal text-slate-400">/{report.target_supply_unit}</span>
            </div>
            <span className="text-xs text-emerald-400 font-bold mt-1 block">
              Total Item: R$ {totalEstimatedGlobal.toLocaleString("pt-BR", { minimumFractionDigits: 2 })}
            </span>
          </div>
        </section>

        {/* PARECER AUTOMÁTICO DE FUNDAMENTAÇÃO METODOLÓGICA (ART. 6º IN 65/2021) */}
        <section className="bg-amber-50/90 border border-amber-300 rounded-xl p-4 text-xs text-slate-800 leading-relaxed">
          <strong className="uppercase tracking-wider text-amber-950 block mb-1">
            Nota Técnica de Fundamentação Metodológica (Art. 6º da IN SEGES/ME nº 65/2021):
          </strong>
          {report.legal_methodology_justification}
        </section>

        {/* TABELA DE RASTREABILIDADE (COTAÇÕES VÁLIDAS vs EXPURGADAS COM LINK PNCP) */}
        <section className="bg-white rounded-2xl shadow-sm border border-slate-200 overflow-hidden">
          <div className="px-6 py-4 bg-slate-100 border-b border-slate-200 flex items-center justify-between">
            <div className="flex items-center gap-3">
              <button
                onClick={() => setActiveTab("VALIDAS")}
                className={`px-4 py-2 rounded-lg text-xs font-bold transition ${
                  activeTab === "VALIDAS"
                    ? "bg-slate-900 text-white shadow-sm"
                    : "bg-white text-slate-600 border border-slate-300"
                }`}
              >
                Cesta de Preços Válidos ({report.valid_quotes.length})
              </button>
              <button
                onClick={() => setActiveTab("EXPURGADAS")}
                className={`px-4 py-2 rounded-lg text-xs font-bold transition ${
                  activeTab === "EXPURGADAS"
                    ? "bg-rose-700 text-white shadow-sm"
                    : "bg-white text-rose-700 border border-rose-200"
                }`}
              >
                Memória de Expurgos / Outliers ({report.purged_quotes.length})
              </button>
            </div>
            <span className="text-xs text-slate-500">
              Clique em <strong>Ver no PNCP / Compras.gov</strong> para abrir a ata de homologação oficial
            </span>
          </div>

          <div className="overflow-x-auto">
            <table className="w-full text-left border-collapse text-xs">
              <thead>
                <tr className="bg-slate-50 border-b border-slate-200 text-slate-600 uppercase">
                  <th className="py-3.5 px-4 font-bold">UASG / Órgão Comprador</th>
                  <th className="py-3.5 px-4 font-bold">Nº da Compra</th>
                  <th className="py-3.5 px-4 font-bold">Fornecedor Vencedor & CNPJ</th>
                  <th className="py-3.5 px-4 font-bold">Data</th>
                  <th className="py-3.5 px-4 font-bold text-right">Qtd.</th>
                  <th className="py-3.5 px-4 font-bold text-right">Preço Unit.</th>
                  <th className="py-3.5 px-4 font-bold">Parecer Estatístico / Motivo</th>
                  <th className="py-3.5 px-4 font-bold text-center">Prova Auditável</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-200">
                {(activeTab === "VALIDAS" ? report.valid_quotes : report.purged_quotes).map(
                  (q) => (
                    <tr
                      key={q.id_compra_item}
                      className={
                        q.is_outlier ? "bg-rose-50/40 hover:bg-rose-50" : "hover:bg-slate-50"
                      }
                    >
                      <td className="py-3 px-4">
                        <span className="font-bold text-slate-900 block">
                          UASG {q.buyer_uasg} ({q.buyer_uf})
                        </span>
                        <span className="text-slate-600">{q.buyer_org_name}</span>
                      </td>
                      <td className="py-3 px-4 font-mono font-bold text-slate-800">
                        {q.purchase_number_year}
                        <span className="block text-[10px] font-normal text-slate-400">
                          ID: {q.id_compra_item.slice(0, 17)}
                        </span>
                      </td>
                      <td className="py-3 px-4">
                        <span className="font-semibold text-slate-800 block">
                          {q.supplier_name}
                        </span>
                        <span className="font-mono text-slate-500">CNPJ: {q.supplier_cnpj}</span>
                      </td>
                      <td className="py-3 px-4 text-slate-600">{q.purchase_date}</td>
                      <td className="py-3 px-4 text-right font-mono">
                        {q.quantity_homologated} {q.supply_unit}
                      </td>
                      <td className="py-3 px-4 text-right font-mono font-extrabold text-sm text-slate-900">
                        R$ {q.unit_price.toFixed(2)}
                      </td>
                      <td className="py-3 px-4 max-w-xs">
                        <span
                          className={`inline-block px-2 py-0.5 rounded text-[10px] font-bold mb-1 ${
                            q.is_outlier
                              ? "bg-rose-100 text-rose-800"
                              : "bg-emerald-100 text-emerald-800"
                          }`}
                        >
                          {q.purge_status}
                        </span>
                        <p className="text-[11px] text-slate-600">{q.purge_justification}</p>
                      </td>
                      <td className="py-3 px-4 text-center">
                        <a
                          href={q.pncp_or_compras_uri}
                          target="_blank"
                          rel="noopener noreferrer"
                          className="inline-block bg-slate-900 hover:bg-slate-700 text-amber-300 font-bold px-3 py-1.5 rounded-lg text-[11px] transition"
                        >
                          Abrir Compra Oficial ↗
                        </a>
                      </td>
                    </tr>
                  )
                )}
              </tbody>
            </table>
          </div>
        </section>
      </div>
    </div>
  );
};
