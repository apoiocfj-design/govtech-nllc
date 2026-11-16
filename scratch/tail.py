
# ---- appended: replace the 3i/3j/3l section by a full helper injection ----
HELPERS = r'''
    // =========================================================================
    // NAVEGAÇÃO ENTRE PASSOS (ITENS DO LOTE <-> RESULTADOS)
    // =========================================================================
    function switchBatchWorkflowView(view) {
      const tray = document.getElementById("batchTraySection");
      const results = document.getElementById("batchResultsSection");
      if (view === "results" && !(batchResearchResults || []).length) {
        showAppNotification("Execute primeiro o passo 2 (Pesquisar Preços).", "warn");
        return;
      }
      if (tray) tray.classList.toggle("hidden", view !== "tray");
      if (results) results.classList.toggle("hidden", view !== "results");
      updateWorkflowStepper();
      window.scrollTo({ top: 0, behavior: "smooth" });
    }

    // =========================================================================
    // CONTADOR REGRESSIVO COM TEMPO ESTIMADO (FEEDBACK IMEDIATO AO USUÁRIO)
    // =========================================================================
    let countdownInterval = null;
    function showCountdownProgress(options) {
      const modal = document.getElementById("progressCountdownModal");
      if (!modal) return;
      const icon = document.getElementById("countdownModalIcon");
      const title = document.getElementById("countdownModalTitle");
      const subtitle = document.getElementById("countdownModalSubtitle");
      const step = document.getElementById("countdownProgressStep");
      const pct = document.getElementById("countdownProgressPercent");
      const bar = document.getElementById("countdownProgressBar");
      const secText = document.getElementById("countdownSecondsText");
      const steps = options.steps || ["Consultando dados oficiais...", "Processando...", "Finalizando..."];
      const totalMs = (options.estimatedSeconds || 4) * 1000;
      let elapsedMs = 0;
      if (title) title.textContent = options.title || "Processando solicitação";
      if (subtitle) subtitle.textContent = options.subtitle || "";
      if (icon) icon.textContent = options.icon || "⏳";
      if (bar) bar.style.width = "0%";
      if (pct) pct.textContent = "0%";
      if (secText) secText.textContent = `~${options.estimatedSeconds || 4}s`;
      if (step) step.textContent = "Iniciando comunicação com servidores oficiais...";
      modal.classList.remove("hidden");
      clearInterval(countdownInterval);
      countdownInterval = setInterval(() => {
        elapsedMs += 100;
        const remaining = Math.ceil((totalMs - elapsedMs) / 1000);
        if (secText) secText.textContent = remaining > 0 ? `${remaining}s` : "finalizando...";
        const p = Math.min(95, Math.round((elapsedMs / totalMs) * 90));
        if (bar) bar.style.width = `${p}%`;
        if (pct) pct.textContent = `${p}%`;
        if (step && elapsedMs >= 600) {
          const k = Math.min(steps.length - 1, Math.floor((elapsedMs / totalMs) * steps.length));
          step.textContent = steps[k];
        }
      }, 100);
    }
    function hideCountdownProgress() {
      const modal = document.getElementById("progressCountdownModal");
      const bar = document.getElementById("countdownProgressBar");
      const pct = document.getElementById("countdownProgressPercent");
      const step = document.getElementById("countdownProgressStep");
      clearInterval(countdownInterval);
      if (bar) bar.style.width = "100%";
      if (pct) pct.textContent = "100%";
      if (step) step.textContent = "Concluído!";
      setTimeout(() => { if (modal) modal.classList.add("hidden"); }, 350);
    }

    // PDF do edital no PNCP (gerado a partir do extrato oficial) com contador
    function openPncpPagePdfWithCountdown(idx) {
      const q = (window.currentAllQuotes || [])[idx];
      if (!q) return;
      downloadFromPost("/api/v1/pncp/page-pdf-post", {
        chave: q.purchase_key || "", url: q.pncp_or_compras_uri || ""
      }, `PNCP_Edital_${q.purchase_key || "documento"}.pdf`, "Gerando extrato do Edital no PNCP (PDF)");
    }

    // Inclusão opcional de uma Ata SRP como cotação da cesta (Art. 5º, II da IN 65/2021)
    function addArpToQuotationBasket(arpIdx) {
      if (!activePriceReport || !(activePriceReport.arp_atas_vigentes || [])[arpIdx]) return;
      const arp = activePriceReport.arp_atas_vigentes[arpIdx];
      const newQuote = {
        id_compra_item: `ARP-${arp.numero_ata || arpIdx}`,
        buyer_uasg: arp.uasg_codigo || "",
        buyer_org_name: arp.uasg_nome || "Órgão Gerenciador",
        buyer_uf: "BR",
        purchase_number_year: `Ata SRP ${arp.numero_ata || ''}`,
        supplier_name: arp.fornecedor_nome || "Fornecedor Registrado",
        supplier_cnpj: arp.fornecedor_cnpj || "",
        supply_unit: activePriceReport.target_supply_unit || "UN",
        quantity_homologated: 1,
        purchase_date: arp.vigencia_inicio || "",
        unit_price: Number(arp.preco_unitario || 0),
        is_outlier: false,
        is_selected: true,
        purge_status: "VALIDA",
        purge_justification: `Preço registrado em Ata SRP vigente (${arp.uasg_nome}), incluído pelo analista (Art. 5º, II da IN 65/2021).`,
        pncp_or_compras_uri: arp.pncp_ata_uri || "https://pncp.gov.br/app/atas",
        purchase_key: `ARP-${arp.numero_ata || arpIdx}`
      };
      window.currentAllQuotes = window.currentAllQuotes || [];
      if (window.currentAllQuotes.some(q => q.purchase_key === newQuote.purchase_key)) {
        showAppNotification("Esta ata já está na cesta.", "warn");
        return;
      }
      window.currentAllQuotes.push(newQuote);
      recalculateCuratedQuotes();
    }

'''
anchor = "    function renderBatchTray() {"
assert html.count(anchor) == 1, html.count(anchor)
html = html.replace(anchor, HELPERS + anchor)
print("ok: helpers injected")

html = html.replace("processBatchResearch(true)", "processBatchResearch()")
open(P, "w", encoding="utf-8", newline="").write(html.replace("\n", "\r\n"))
print("DONE")
