    const AppCore = window.AppCore;
    const DEALS_STORAGE_KEY = (AppCore && AppCore.STORAGE_KEYS && AppCore.STORAGE_KEYS.deals) || "deals_data_v1";
    let dealsData = [];
    let accountingAccessState = { restricted: false, allowed: true };
    let dealsFilterState = {
      keyword: "",
      stage: "all",
      owner: "all",
      setup: "all",
    };

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
        return AppCore.saveDealsData(dealsData);
      }
      try {
        localStorage.setItem(DEALS_STORAGE_KEY, JSON.stringify(dealsData));
      } catch (e) {
        console.warn("Failed to save deals to storage", e);
      }
      return Promise.resolve();
    }

    function normalizeValue(value) {
      if (AppCore) return AppCore.normalizeValue(value);
      return String(value || "").trim().toLowerCase();
    }

    function buildPageUrl(pageId, params) {
      if (AppCore && typeof AppCore.getPageUrl === "function") {
        return AppCore.getPageUrl(pageId, params);
      }
      if (window.PlutusAppConfig && typeof window.PlutusAppConfig.buildPageHref === "function") {
        return window.PlutusAppConfig.buildPageHref(pageId, params);
      }
      const query = new URLSearchParams(params || {});
      const queryString = query.toString();
      return queryString ? `${pageId}.html?${queryString}` : `${pageId}.html`;
    }

    function hasDashboardLinked(deal) {
      return Boolean(String(deal && deal.fundraisingDashboardId || "").trim());
    }

    function hasDeckLinked(deal) {
      return Boolean(String(deal && deal.deckUrl || "").trim());
    }

    function normalizeDealLegalLinksForFilters(deal) {
      const source =
        deal && Array.isArray(deal.legalLinks)
          ? deal.legalLinks
          : deal && Array.isArray(deal.legalAspects)
            ? deal.legalAspects
            : [];

      return source
        .map((entry) => {
          if (typeof entry === "string") return String(entry || "").trim();
          if (!entry || typeof entry !== "object") return "";
          return String(entry.url || entry.href || entry.link || "").trim();
        })
        .filter(Boolean);
    }

    function hasLegalLinked(deal) {
      return normalizeDealLegalLinksForFilters(deal).length > 0;
    }

    function getSubOwners(deal) {
      const source = deal && deal.subOwners;
      const values = Array.isArray(source)
        ? source
        : typeof source === "string"
          ? source.split(/[\n,;]+/)
          : [];
      return Array.from(
        new Set(
          values
            .map((entry) => String(entry || "").trim())
            .filter(Boolean)
        )
      );
    }

    function getDealOwners(deal) {
      return [
        String(deal && (deal.seniorOwner || deal.owner) || "").trim(),
        String(deal && deal.juniorOwner || "").trim(),
        ...getSubOwners(deal),
      ].filter(Boolean);
    }

    function parseSubOwnersInput(value) {
      return Array.from(
        new Set(
          String(value || "")
            .split(/[\n,;]+/)
            .map((entry) => String(entry || "").trim())
            .filter(Boolean)
        )
      );
    }

    function getPeopleInChargeText(deal) {
      const primaryOwners = [getSeniorOwner(deal), getJuniorOwner(deal)]
        .map((entry) => String(entry || "").trim())
        .filter((entry) => entry && entry !== "–");
      const subOwners = getSubOwners(deal);
      if (!subOwners.length) {
        return `${getSeniorOwner(deal)} / ${getJuniorOwner(deal)}`;
      }
      return `${primaryOwners.join(" / ") || "–"} · +${subOwners.length} sub`;
    }

    function syncDealFilterStateFromUi() {
      const searchInput = document.getElementById("deal-search-input");
      const stageFilter = document.getElementById("deal-stage-filter");
      const ownerFilter = document.getElementById("deal-owner-filter");
      const setupFilter = document.getElementById("deal-setup-filter");

      dealsFilterState = {
        keyword: String(searchInput && searchInput.value || "").trim().toLowerCase(),
        stage: String(stageFilter && stageFilter.value || "all").trim().toLowerCase() || "all",
        owner: String(ownerFilter && ownerFilter.value || "all").trim().toLowerCase() || "all",
        setup: String(setupFilter && setupFilter.value || "all").trim().toLowerCase() || "all",
      };
    }

    function applyDealFilterStateToUi() {
      const searchInput = document.getElementById("deal-search-input");
      const stageFilter = document.getElementById("deal-stage-filter");
      const ownerFilter = document.getElementById("deal-owner-filter");
      const setupFilter = document.getElementById("deal-setup-filter");

      if (searchInput) searchInput.value = dealsFilterState.keyword || "";
      if (stageFilter) stageFilter.value = dealsFilterState.stage || "all";
      if (ownerFilter) ownerFilter.value = dealsFilterState.owner || "all";
      if (setupFilter) setupFilter.value = dealsFilterState.setup || "all";
    }

    function populateDealOwnerFilter() {
      const ownerFilter = document.getElementById("deal-owner-filter");
      if (!ownerFilter) return;

      const currentValue = String(dealsFilterState.owner || ownerFilter.value || "all").trim().toLowerCase() || "all";
      const owners = Array.from(
        new Set(
          (Array.isArray(dealsData) ? dealsData : [])
            .flatMap((deal) => getDealOwners(deal))
            .map((owner) => owner.trim())
            .filter(Boolean)
        )
      ).sort((a, b) => a.localeCompare(b));

      ownerFilter.innerHTML = ['<option value="all">All owners</option>']
        .concat(owners.map((owner) => `<option value="${owner.toLowerCase()}">${owner}</option>`))
        .join("");

      ownerFilter.value = owners.some((owner) => owner.toLowerCase() === currentValue) ? currentValue : "all";
      dealsFilterState.owner = ownerFilter.value;
    }

    function matchesDealFilters(deal) {
      if (!deal) return false;

      const keyword = dealsFilterState.keyword;
      if (keyword) {
        const haystack = [
          deal.id,
          deal.name,
          deal.company,
          deal.stage,
          getSeniorOwner(deal),
          getJuniorOwner(deal),
          getSubOwners(deal).join(" "),
          deal.currency,
          deal.summary,
        ]
          .join(" ")
          .toLowerCase();
        if (!haystack.includes(keyword)) return false;
      }

      const stageValue = normalizeValue(deal.stage);
      if (dealsFilterState.stage !== "all" && stageValue !== dealsFilterState.stage) {
        return false;
      }

      if (dealsFilterState.owner !== "all") {
        const owners = getDealOwners(deal).map((owner) => owner.toLowerCase());
        if (!owners.includes(dealsFilterState.owner)) return false;
      }

      const hasDashboard = hasDashboardLinked(deal);
      const hasDeck = hasDeckLinked(deal);
      const hasLegal = hasLegalLinked(deal);
      if (dealsFilterState.setup === "missing-dashboard" && hasDashboard) return false;
      if (dealsFilterState.setup === "missing-deck" && hasDeck) return false;
      if (dealsFilterState.setup === "missing-legal" && hasLegal) return false;
      if (dealsFilterState.setup === "missing-any" && hasDashboard && hasDeck && hasLegal) return false;
      if (dealsFilterState.setup === "ready" && !(hasDashboard && hasDeck && hasLegal)) return false;

      return true;
    }

    function getFilteredDeals() {
      return (Array.isArray(dealsData) ? dealsData : []).filter((deal) => matchesDealFilters(deal));
    }

    function buildDealStageCounts(source) {
      const counts = {
        total: Array.isArray(source) ? source.length : 0,
        prospect: 0,
        signing: 0,
        onboarding: 0,
        contacting: 0,
      };

      (Array.isArray(source) ? source : []).forEach((deal) => {
        const stage = String(deal && deal.stage || "").toLowerCase();
        if (stage === "prospect") counts.prospect += 1;
        else if (stage === "signing") counts.signing += 1;
        else if (stage === "onboarding") counts.onboarding += 1;
        else if (stage === "contacting investors") counts.contacting += 1;
      });

      return counts;
    }

    function renderStageFilterChips(metaRow, counts) {
      if (!metaRow) return;
      const chips = [
        { value: "all", label: "All deals", count: counts.total },
        { value: "prospect", label: "Prospect", count: counts.prospect },
        { value: "signing", label: "Signing", count: counts.signing },
        { value: "onboarding", label: "Onboarding", count: counts.onboarding },
        { value: "contacting investors", label: "Contacting investors", count: counts.contacting },
      ];

      metaRow.innerHTML = chips.map((chip) => `
        <button class="chip${dealsFilterState.stage === chip.value ? " is-active" : ""}" type="button" data-deal-stage-chip="${chip.value}">
          <strong>${chip.count}</strong> ${chip.label}
        </button>
      `).join("");
    }

    function updateDealFilterSummary(filteredDeals) {
      const summaryEl = document.getElementById("deal-filter-summary");
      if (!summaryEl) return;
      const ownerFilter = document.getElementById("deal-owner-filter");
      const setupFilter = document.getElementById("deal-setup-filter");

      const total = Array.isArray(dealsData) ? dealsData.length : 0;
      const shown = Array.isArray(filteredDeals) ? filteredDeals.length : 0;
      const summaryParts = [`Showing ${shown} of ${total} deal${total === 1 ? "" : "s"}`];

      if (dealsFilterState.stage !== "all") {
        summaryParts.push(`stage: ${STAGE_LABELS[dealsFilterState.stage] || dealsFilterState.stage}`);
      }
      if (dealsFilterState.owner !== "all") {
        const ownerLabel = ownerFilter && ownerFilter.selectedOptions && ownerFilter.selectedOptions[0]
          ? ownerFilter.selectedOptions[0].text
          : dealsFilterState.owner;
        summaryParts.push(`owner: ${ownerLabel}`);
      }
      if (dealsFilterState.setup !== "all") {
        const setupLabel = setupFilter && setupFilter.selectedOptions && setupFilter.selectedOptions[0]
          ? setupFilter.selectedOptions[0].text
          : dealsFilterState.setup.replace(/-/g, " ");
        summaryParts.push(`setup: ${setupLabel}`);
      }
      if (dealsFilterState.keyword) {
        summaryParts.push(`search: "${dealsFilterState.keyword}"`);
      }

      summaryEl.textContent = `${summaryParts.join(" / ")}.`;
    }

    function setupDealFilters() {
      const searchInput = document.getElementById("deal-search-input");
      const stageFilter = document.getElementById("deal-stage-filter");
      const ownerFilter = document.getElementById("deal-owner-filter");
      const setupFilter = document.getElementById("deal-setup-filter");
      const resetBtn = document.getElementById("btn-reset-deal-filters");
      const metaRow = document.getElementById("meta-row");

      [searchInput, stageFilter, ownerFilter, setupFilter].forEach((control) => {
        if (!control) return;
        const eventName = control.tagName === "INPUT" ? "input" : "change";
        control.addEventListener(eventName, () => {
          syncDealFilterStateFromUi();
          renderDeals();
        });
      });

      if (resetBtn) {
        resetBtn.addEventListener("click", () => {
          dealsFilterState = {
            keyword: "",
            stage: "all",
            owner: "all",
            setup: "all",
          };
          applyDealFilterStateToUi();
          renderDeals();
        });
      }

      if (metaRow) {
        metaRow.addEventListener("click", (event) => {
          const chip = event.target.closest("[data-deal-stage-chip]");
          if (!chip) return;
          const nextStage = String(chip.getAttribute("data-deal-stage-chip") || "all").trim().toLowerCase() || "all";
          dealsFilterState.stage = dealsFilterState.stage === nextStage && nextStage !== "all" ? "all" : nextStage;
          applyDealFilterStateToUi();
          renderDeals();
        });
      }
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

    function setDealFormStatus(message, isError) {
      const status = document.getElementById("deal-form-status");
      if (!status) return;
      status.textContent = message || "";
      status.style.color = isError ? "#fecdd3" : "var(--text-dim)";
    }

    function setupAddDealForm() {
      const form = document.getElementById("deal-form");
      if (!form) return;

      const nameInput = document.getElementById("deal-name-input");
      const companyInput = document.getElementById("deal-company-input");
      const seniorInput = document.getElementById("deal-senior-input");
      const juniorInput = document.getElementById("deal-junior-input");
      const subOwnersInput = document.getElementById("deal-sub-owners-input");
      const targetInput = document.getElementById("deal-target-input");
      const raisedInput = document.getElementById("deal-raised-input");
      const currencyInput = document.getElementById("deal-currency-input");
      const dashboardInput = document.getElementById("deal-dashboard-input");
      const cashInput = document.getElementById("deal-cash-input");
      const equityInput = document.getElementById("deal-equity-input");
      const retainerInput = document.getElementById("deal-retainer-input");
      const summaryInput = document.getElementById("deal-summary-input");
      const dashboardUrlInput = document.getElementById("deal-dashboard-url-input");
      const dashboardDescInput = document.getElementById("deal-dashboard-description-input");
      const dashboardFundsInput = document.getElementById("deal-dashboard-sheet-funds-input");
      const dashboardFoInput = document.getElementById("deal-dashboard-sheet-fo-input");
      const dashboardFiguresInput = document.getElementById("deal-dashboard-sheet-figures-input");
      const panel = document.getElementById("add-deal-panel");

      form.addEventListener("submit", async (event) => {
        event.preventDefault();

        const name = String(nameInput.value || "").trim();
        const company = String(companyInput.value || "").trim();
        const seniorOwner = String(seniorInput.value || "").trim();
        const juniorOwner = String(juniorInput.value || "").trim();
        const subOwners = parseSubOwnersInput(subOwnersInput && subOwnersInput.value);
        if (!name || !company || !seniorOwner || !juniorOwner) return;

        const dashboardUrl = String(dashboardUrlInput && dashboardUrlInput.value || "").trim();
        let dashboardId = String(dashboardInput.value || "").trim();
        let createdDealId = "";
        if (!dashboardId && dashboardUrl) {
          dashboardId = toDealId(company || name);
        }
        if (dashboardInput) dashboardInput.value = dashboardId;

        setDealFormStatus(dashboardUrl ? "Creating deal and linked dashboard..." : "Creating deal...", false);

        try {
          if (dashboardUrl) {
            if (!dashboardId) {
              throw new Error("Dashboard ID is required when you provide a dashboard link.");
            }
            if (!AppCore || typeof AppCore.upsertDashboardConfigEntry !== "function") {
              throw new Error("Dashboard sync is unavailable in this mode.");
            }

            await AppCore.upsertDashboardConfigEntry({
              id: dashboardId,
              name,
              description: String(dashboardDescInput && dashboardDescInput.value || "").trim(),
              excelUrl: dashboardUrl,
              sheets: {
                funds: String(dashboardFundsInput && dashboardFundsInput.value || "").trim() || "funds",
                familyOffices: String(dashboardFoInput && dashboardFoInput.value || "").trim() || "f.o.",
                figures: String(dashboardFiguresInput && dashboardFiguresInput.value || "").trim() || "figure",
              },
            });
          }

          createdDealId = buildUniqueDealId(`${company}-${name}`);
          const newDeal = {
            id: createdDealId,
            name,
            company,
            stage: "prospect",
            seniorOwner,
            juniorOwner,
            subOwners,
            owner: seniorOwner,
            targetAmount: parseNumericAmount(targetInput.value),
            raisedAmount: parseNumericAmount(raisedInput.value),
            currency: String(currencyInput.value || "USD").trim() || "USD",
            fundraisingDashboardId: dashboardId,
            CashCommission: String(cashInput.value || "").trim(),
            EquityCommission: String(equityInput.value || "").trim(),
            Retainer: String(retainerInput.value || "").trim(),
            summary: String(summaryInput.value || "").trim(),
          };

          dealsData.push(newDeal);
          await saveDealsData();
          renderDeals();
          form.reset();
          currencyInput.value = "USD";
          setDealFormStatus(dashboardUrl ? "Deal and dashboard created." : "Deal created.", false);
          if (panel) panel.open = false;
        } catch (error) {
          if (createdDealId) {
            dealsData = dealsData.filter((deal) => !(deal && normalizeValue(deal.id) === normalizeValue(createdDealId)));
          }
          setDealFormStatus(
            error instanceof Error ? error.message : "Failed to create deal.",
            true,
          );
        }
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

    function getDashboardConfig() {
      if (AppCore && typeof AppCore.getDashboardConfig === "function") {
        return AppCore.getDashboardConfig();
      }
      return window.DASHBOARD_CONFIG || { dashboards: [], settings: {} };
    }

    function getDealDashboard(deal) {
      const config = getDashboardConfig();
      if (AppCore && typeof AppCore.getDashboardForDeal === "function") {
        return AppCore.getDashboardForDeal(deal, config);
      }

      const dashboardId = normalizeValue(deal && deal.fundraisingDashboardId);
      if (!dashboardId || !Array.isArray(config.dashboards)) return null;
      return config.dashboards.find((dashboard) => normalizeValue(dashboard.id) === dashboardId) || null;
    }

    async function refreshAccountingAccessState() {
      if (!AppCore || typeof AppCore.getPageAccessStatus !== "function") {
        accountingAccessState = { restricted: false, allowed: true };
        return accountingAccessState;
      }

      try {
        accountingAccessState = await AppCore.getPageAccessStatus("accounting");
      } catch {
        accountingAccessState = { restricted: false, allowed: true };
      }
      return accountingAccessState;
    }

    function renderDealActions(deal) {
      const detailHref = buildPageUrl("deal-details", { id: deal.id });
      const accountingHref = buildPageUrl("accounting", { id: deal.id });
      const dashboard = getDealDashboard(deal);
      const links = [
        `<a class="action-link" href="${detailHref}">View deal</a>`,
      ];

      if (!accountingAccessState.restricted || accountingAccessState.allowed) {
        links.push(`<a class="action-link" href="${accountingHref}">Accounting</a>`);
      }

      if (dashboard && dashboard.id) {
        const dashboardHref = buildPageUrl("investor-dashboard", { dashboard: dashboard.id });
        links.push(`<a class="action-link" href="${dashboardHref}">Open dashboard</a>`);
      }

      return links.join(' <span style="color: var(--text-dim);">·</span> ');
    }

    function renderDeals() {
      const body = document.getElementById("deals-body");
      const metaRow = document.getElementById("meta-row");
      const footerMeta = document.getElementById("footer-meta");
      loadDealsData();
      if (!Array.isArray(dealsData)) return;
      populateDealOwnerFilter();
      applyDealFilterStateToUi();

      const filteredDeals = getFilteredDeals();
      const allCounts = buildDealStageCounts(dealsData);

      body.innerHTML = "";

      const formatTextOrUnknown = (value) => {
        if (value == null || String(value).trim() === "") {
          return `<span class="amount-unknown">??</span>`;
        }
        return String(value);
      };

      filteredDeals.forEach(deal => {
        const tr = document.createElement("tr");
        const stage = String(deal.stage || "").toLowerCase();

        tr.innerHTML = `
          <td class="name-cell">
            <a class="action-link" href="${buildPageUrl("deal-details", { id: deal.id })}">${deal.name || "–"}</a>
          </td>
          <td>${deal.company || "–"}</td>
          <td>
            <span class="stage-badge">
              <span class="stage-dot ${stageClass(stage)}"></span>
              ${STAGE_LABELS[stage] || "Prospect"}
            </span>
          </td>
          <td>${getPeopleInChargeText(deal)}</td>
          <td class="amount">${formatAmountCell(deal.targetAmount, deal.currency)}</td>
          <td class="amount">${formatAmountCell(deal.raisedAmount, deal.currency)}</td>
          <td>${formatTextOrUnknown(deal.CashCommission)}</td>
          <td>${formatTextOrUnknown(deal.EquityCommission)}</td>
          <td>${formatTextOrUnknown(deal.Retainer)}</td>
          <td>${renderDealActions(deal)}</td>
        `;

        body.appendChild(tr);
      });

      if (!filteredDeals.length) {
        body.innerHTML = '<tr><td colspan="10" class="amount-unknown">No deals match the current filters.</td></tr>';
      }

      renderStageFilterChips(metaRow, allCounts);
      updateDealFilterSummary(filteredDeals);
      footerMeta.textContent = filteredDeals.length === dealsData.length
        ? `${allCounts.total} active deal${allCounts.total === 1 ? "" : "s"} in pipeline.`
        : `${filteredDeals.length} shown of ${allCounts.total} active deal${allCounts.total === 1 ? "" : "s"}.`;
    }

    document.addEventListener("DOMContentLoaded", () => {
      const initialize = async () => {
        if (AppCore && typeof AppCore.refreshDashboardConfigFromShareDrive === "function") {
          await AppCore.refreshDashboardConfigFromShareDrive();
        }
        await refreshAccountingAccessState();
        loadDealsData();
        setupDealFilters();
        setupAddDealForm();
        renderDeals();
      };
      initialize();
      if (AppCore) {
        window.addEventListener("appcore:deals-updated", renderDeals);
        window.addEventListener("appcore:dashboard-config-updated", renderDeals);
        window.addEventListener("appcore:graph-session-updated", async () => {
          await refreshAccountingAccessState();
          renderDeals();
        });
      }
    });
