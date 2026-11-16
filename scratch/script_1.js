
    // =========================================================================
    // MODAL DE PERFIL CADASTRAL DO FORNECEDOR, CNAE E QSA (PREVENÇÃO DE CONLUIO)
    // =========================================================================
    async function openSupplierCnpjModal(cnpj, supplierName) {
      const cleanDigits = (cnpj || "").replace(/\D/g, "");
      if (!cleanDigits) {
        showAppNotification("CNPJ não informado para consulta.", "warn");
        return;
      }

      const modal = document.getElementById("supplierCnpjModal");
      const loading = document.getElementById("supplierModalLoading");
      const content = document.getElementById("supplierModalContent");
      const titleEl = document.getElementById("modalSupplierCnpjTitle");

      if (titleEl) titleEl.textContent = `CNPJ: ${cnpj}`;
      if (modal) modal.classList.remove("hidden");
      if (loading) loading.classList.remove("hidden");
      if (content) content.classList.add("hidden");

      try {
        const resp = await fetch(`/api/v1/supplier/cnpj-info?cnpj=${cleanDigits}`);
        const data = await resp.json();
        if (loading) loading.classList.add("hidden");
        if (content) content.classList.remove("hidden");

        const sup = data.supplier || {};
        document.getElementById("supRazaoSocial").textContent = sup.razao_social || decodeURIComponent(supplierName || 'Fornecedor');
        document.getElementById("supNomeFantasia").textContent = sup.nome_fantasia ? `(${sup.nome_fantasia})` : "";
        document.getElementById("supSituacaoCadastral").textContent = sup.situacao_cadastral || "ATIVA";
        document.getElementById("supPorte").textContent = sup.porte || "Demais";
        document.getElementById("supMunicipioUf").textContent = `${sup.municipio || '—'} / ${sup.uf || 'BR'}`;
        document.getElementById("supDataSituacao").textContent = sup.data_situacao_cadastral || "—";
        document.getElementById("supCnaeCodigo").textContent = sup.cnae_codigo || "—";
        document.getElementById("supCnaeDescricao").textContent = sup.cnae_descricao || "Atividade econômica registrada";

        const qsa = sup.qsa || [];
        const qsaCountEl = document.getElementById("supQsaCount");
        const qsaListEl = document.getElementById("supQsaList");
        if (qsaCountEl) qsaCountEl.textContent = `${qsa.length} sócio(s)`;
        if (qsaListEl) {
          if (qsa.length > 0) {
            qsaListEl.innerHTML = qsa.map(s => `
              <li class="py-1 flex items-center justify-between">
                <strong>${s.nome || 'Sócio'}</strong>
                <span class="text-slate-500">${s.qualificacao || 'Sócio-Administrador'}</span>
              </li>
            `).join("");
          } else {
            qsaListEl.innerHTML = `<li class="py-1 text-slate-400 italic">Quadro societário individual ou não informado.</li>`;
          }
        }

        const lTransp = document.getElementById("supLinkTransp");
        if (lTransp) lTransp.href = `https://portaldatransparencia.gov.br/pessoa-juridica/${cleanDigits}`;
      } catch (e) {
        if (loading) loading.classList.add("hidden");
        if (content) content.classList.remove("hidden");
        showAppNotification("Erro ao consultar dados cadastrais: " + e.message, "error");
      }
    }

    function closeSupplierCnpjModal() {
      const modal = document.getElementById("supplierCnpjModal");
      if (modal) modal.classList.add("hidden");
    }

    // =========================================================================
    // MODAL DE COTAÇÃO COMPLEMENTAR (ART. 5º DA IN 65/2021)
    // =========================================================================
    function openAddComplementaryQuoteModal() {
      const modal = document.getElementById("complementaryQuoteModal");
      if (modal) modal.classList.remove("hidden");
    }

    function closeAddComplementaryQuoteModal() {
      const modal = document.getElementById("complementaryQuoteModal");
      if (modal) modal.classList.add("hidden");
    }

    function saveComplementaryQuote() {
      const priceVal = parseFloat(document.getElementById("compQuotePrice").value);
      if (isNaN(priceVal) || priceVal <= 0) {
        showAppNotification("Informe um preço unitário válido maior que zero.", "warn");
        return;
      }
      const supplier = document.getElementById("compQuoteSupplier").value.trim() || "Cotação Complementar";
      const dateVal = document.getElementById("compQuoteDate").value || new Date().toISOString().slice(0, 10);
      const inciso = document.getElementById("compQuoteInciso").value;
      const cnpjVal = document.getElementById("compQuoteCnpj").value.trim() || "";
      const brandVal = document.getElementById("compQuoteBrand").value.trim() || "Complementar";
      const justVal = document.getElementById("compQuoteJust").value.trim() || `Cotação complementar inserida conforme ${inciso}.`;

      if (!window.currentAllQuotes) window.currentAllQuotes = [];

      window.currentAllQuotes.push({
        id_compra_item: `COMP_${Date.now()}`,
        purchase_number_year: "Cotação Complementar",
        buyer_uasg: "Órgão Licitante",
        buyer_org_name: inciso,
        buyer_uf: "BR",
        supplier_cnpj: cnpjVal || "N/A",
        supplier_name: supplier,
        supply_unit: (activePriceReport && activePriceReport.target_supply_unit) || "UN",
        quantity_homologated: 1,
        unit_price: priceVal,
        purchase_date: dateVal,
        pncp_or_compras_uri: brandVal.startsWith("http") ? brandVal : "https://pncp.gov.br",
        brand: brandVal,
        is_outlier: false,
        is_selected: true,
        purge_status: "VALIDA",
        purge_justification: justVal,
        source_inciso: inciso
      });

      closeAddComplementaryQuoteModal();
      recalculateCuratedQuotes();
      showAppNotification("Cotação complementar inserida na cesta!", "success");
    }
  