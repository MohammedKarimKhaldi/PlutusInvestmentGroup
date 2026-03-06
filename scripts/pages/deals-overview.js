    const AppCore = window.AppCore;
    const DEALS_STORAGE_KEY = (AppCore && AppCore.STORAGE_KEYS && AppCore.STORAGE_KEYS.deals) || "deals_data_v1";
    let dealsData = [];

    const STAGE_LABELS = {
      "prospect": "Prospect",
      "signing": "Signing",
      "onboarding": "Onboarding",
      "contacting investors": "Contacting investors"
    };

    function loadDealsData() {
      dealsData = AppCore ? AppCore.loadDealsData() : (Array.isArray(DEALS) ? JSON.parse(JSON.stringify(DEALS)) : []);
    }

    function saveDealsData() {
      if (AppCore) {
        AppCore.saveDealsData(dealsData);
        return;
      }
      try {
        localStorage.setItem(DEALS_STORAGE_KEY, JSON.stringify(dealsData));
      } catch (e) {
        console.warn("Failed to save deals to storage", e);
      }
    }

    function normalizeValue(value) {
      if (AppCore) return AppCore.normalizeValue(value);
      return String(value || "").trim().toLowerCase();
    }

    function toDealId(value) {
      return String(value || "")
        .trim()
        .toLowerCase()
        .replace(/[^a-z0-9]+/g, "-")
        .replace(/^-+|-+$/g, "")
        .slice(0, 80);
    }

    function parseNumericAmount(value) {
      const raw = String(value == null ? "" : value).trim();
      if (!raw) return null;
      const cleaned = raw.replace(/,/g, "");
      const parsed = Number(cleaned);
      if (!Number.isNaN(parsed) && Number.isFinite(parsed)) return parsed;
      return raw;
    }

    function buildUniqueDealId(baseId) {
      const safeBase = toDealId(baseId) || `deal-${Date.now()}`;
      const existing = new Set((Array.isArray(dealsData) ? dealsData : []).map((deal) => normalizeValue(deal.id)));
      if (!existing.has(safeBase)) return safeBase;
      let i = 2;
      while (existing.has(`${safeBase}-${i}`)) i += 1;
      return `${safeBase}-${i}`;
    }

    function setupAddDealForm() {
      const form = document.getElementById("deal-form");
      if (!form) return;

      const nameInput = document.getElementById("deal-name-input");
      const companyInput = document.getElementById("deal-company-input");
      const seniorInput = document.getElementById("deal-senior-input");
      const juniorInput = document.getElementById("deal-junior-input");
      const targetInput = document.getElementById("deal-target-input");
      const raisedInput = document.getElementById("deal-raised-input");
      const currencyInput = document.getElementById("deal-currency-input");
      const dashboardInput = document.getElementById("deal-dashboard-input");
      const cashInput = document.getElementById("deal-cash-input");
      const equityInput = document.getElementById("deal-equity-input");
      const retainerInput = document.getElementById("deal-retainer-input");
      const summaryInput = document.getElementById("deal-summary-input");
      const panel = document.getElementById("add-deal-panel");

      form.addEventListener("submit", (event) => {
        event.preventDefault();

        const name = String(nameInput.value || "").trim();
        const company = String(companyInput.value || "").trim();
        const seniorOwner = String(seniorInput.value || "").trim();
        const juniorOwner = String(juniorInput.value || "").trim();
        if (!name || !company || !seniorOwner || !juniorOwner) return;

        const newDeal = {
          id: buildUniqueDealId(`${company}-${name}`),
          name,
          company,
          stage: "prospect",
          seniorOwner,
          juniorOwner,
          owner: seniorOwner,
          targetAmount: parseNumericAmount(targetInput.value),
          raisedAmount: parseNumericAmount(raisedInput.value),
          currency: String(currencyInput.value || "USD").trim() || "USD",
          fundraisingDashboardId: String(dashboardInput.value || "").trim(),
          CashCommission: String(cashInput.value || "").trim(),
          EquityCommission: String(equityInput.value || "").trim(),
          Retainer: String(retainerInput.value || "").trim(),
          summary: String(summaryInput.value || "").trim(),
        };

        dealsData.push(newDeal);
        saveDealsData();
        renderDeals();
        form.reset();
        currencyInput.value = "USD";
        if (panel) panel.open = false;
      });
    }

    function formatAmountCell(amount, currency) {
      if (amount == null || amount === "") {
        return `<span class="amount-unknown">??</span>`;
      }
      if (typeof amount === "string") {
        return amount;
      }
      if (typeof amount === "number" && !isNaN(amount)) {
        try {
          const formatted = new Intl.NumberFormat("en-US", {
            style: "currency",
            currency: currency || "USD",
            maximumFractionDigits: 0
          }).format(amount);
          return formatted;
        } catch {
          return `${currency || "USD"} ${Number(amount).toLocaleString()}`;
        }
      }
      return `<span class="amount-unknown">??</span>`;
    }

    function stageClass(stage) {
      const s = String(stage || "").toLowerCase();
      if (s === "prospect") return "stage-prospect";
      if (s === "signing") return "stage-signing";
      if (s === "onboarding") return "stage-onboarding";
      if (s === "contacting investors") return "stage-contacting";
      return "stage-prospect";
    }

    function getSeniorOwner(deal) {
      return (deal && (deal.seniorOwner || deal.owner)) || "–";
    }

    function getJuniorOwner(deal) {
      return (deal && deal.juniorOwner) || "–";
    }

    function renderDeals() {
      const body = document.getElementById("deals-body");
      const metaRow = document.getElementById("meta-row");
      const footerMeta = document.getElementById("footer-meta");
      loadDealsData();
      if (!Array.isArray(dealsData)) return;

      body.innerHTML = "";

      let counts = {
        total: dealsData.length,
        prospect: 0,
        signing: 0,
        onboarding: 0,
        contacting: 0
      };

      const formatTextOrUnknown = (value) => {
        if (value == null || String(value).trim() === "") {
          return `<span class="amount-unknown">??</span>`;
        }
        return String(value);
      };

      dealsData.forEach(deal => {
        const tr = document.createElement("tr");
        const stage = String(deal.stage || "").toLowerCase();
        if (stage === "prospect") counts.prospect++;
        else if (stage === "signing") counts.signing++;
        else if (stage === "onboarding") counts.onboarding++;
        else if (stage === "contacting investors") counts.contacting++;

        tr.innerHTML = `
          <td class="name-cell">
            <a class="action-link" href="deal-details.html?id=${encodeURIComponent(deal.id)}">${deal.name || "–"}</a>
          </td>
          <td>${deal.company || "–"}</td>
          <td>
            <span class="stage-badge">
              <span class="stage-dot ${stageClass(stage)}"></span>
              ${STAGE_LABELS[stage] || "Prospect"}
            </span>
          </td>
          <td>${getSeniorOwner(deal)} / ${getJuniorOwner(deal)}</td>
          <td class="amount">${formatAmountCell(deal.targetAmount, deal.currency)}</td>
          <td class="amount">${formatAmountCell(deal.raisedAmount, deal.currency)}</td>
          <td>${formatTextOrUnknown(deal.CashCommission)}</td>
          <td>${formatTextOrUnknown(deal.EquityCommission)}</td>
          <td>${formatTextOrUnknown(deal.Retainer)}</td>
          <td>
            <a class="action-link" href="deal-details.html?id=${encodeURIComponent(deal.id)}">View deal</a>
          </td>
        `;

        body.appendChild(tr);
      });

      metaRow.innerHTML = "";
      const chips = [
        { label: "Total deals", value: counts.total },
        { label: "Prospect", value: counts.prospect },
        { label: "Signing", value: counts.signing },
        { label: "Onboarding", value: counts.onboarding },
        { label: "Contacting investors", value: counts.contacting },
      ];

      chips.forEach(ch => {
        const div = document.createElement("div");
        div.className = "chip";
        div.innerHTML = `<strong>${ch.value}</strong> ${ch.label}`;
        metaRow.appendChild(div);
      });

      footerMeta.textContent = `${counts.total} active deal${counts.total === 1 ? "" : "s"} in pipeline.`;
    }

    document.addEventListener("DOMContentLoaded", () => {
      loadDealsData();
      setupAddDealForm();
      renderDeals();
    });
