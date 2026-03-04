    const AppCore = window.AppCore;
    let dealsData = [];

    const STAGE_LABELS = {
      "prospect": "Prospect",
      "onboarding": "Onboarding",
      "contacting investors": "Contacting investors"
    };

    function loadDealsData() {
      dealsData = AppCore ? AppCore.loadDealsData() : (Array.isArray(DEALS) ? JSON.parse(JSON.stringify(DEALS)) : []);
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
        else if (stage === "onboarding") counts.onboarding++;
        else if (stage === "contacting investors") counts.contacting++;

        tr.innerHTML = `
          <td class="name-cell">
            <a class="action-link" href="deal.html?id=${encodeURIComponent(deal.id)}">${deal.name || "–"}</a>
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
            <a class="action-link" href="deal.html?id=${encodeURIComponent(deal.id)}">View deal</a>
          </td>
        `;

        body.appendChild(tr);
      });

      metaRow.innerHTML = "";
      const chips = [
        { label: "Total deals", value: counts.total },
        { label: "Prospect", value: counts.prospect },
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

    document.addEventListener("DOMContentLoaded", renderDeals);
