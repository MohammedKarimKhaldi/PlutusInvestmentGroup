    const AppCore = window.AppCore;
    const DealIntegrity = window.DealIntegrity || null;
    const DEALS_STORAGE_KEY = (AppCore && AppCore.STORAGE_KEYS && AppCore.STORAGE_KEYS.deals) || "deals_data_v1";
    let dealsData = [];
    let accountingAccessState = { restricted: false, allowed: true };
    let dealIntegrityCache = new Map();
    let ownershipSnapshotState = {
      loading: false,
      loaded: false,
      error: "",
      configured: false,
      entriesByDealId: new Map(),
      unlinkedEntries: [],
      linkedDeals: 0,
      rows: 0,
    };
    let dealsFilterState = {
      portfolio: "active",
      keyword: "",
      stage: "all",
      owner: "all",
      retainer: "all",
      setup: "all",
    };

    const STAGE_LABELS = {
      "prospect": "Prospect",
      "signing": "Signing",
      "onboarding": "Onboarding",
      "contacting investors": "Contacting investors"
    };
    const DEAL_LIFECYCLE_LABELS = {
      active: "Active",
      finished: "Finished",
      closed: "Closed - not concluded",
    };

    function normalizeDealLifecycleStatus(value) {
      const normalized = normalizeValue(value);
      if (normalized === "finished") return "finished";
      if (normalized === "closed") return "closed";
      return "active";
    }

    function getDealLifecycleStatus(deal) {
      return normalizeDealLifecycleStatus(deal && (deal.lifecycleStatus || deal.dealStatus));
    }

    function isDealClosedLifecycle(status) {
      return status === "finished" || status === "closed";
    }

    function renderLifecycleBadge(status) {
      const normalized = normalizeDealLifecycleStatus(status);
      return `<span class="deal-lifecycle-badge is-${normalized}">${DEAL_LIFECYCLE_LABELS[normalized] || "Active"}</span>`;
    }

    function matchesPortfolioFilter(deal) {
      const lifecycleStatus = getDealLifecycleStatus(deal);
      if (dealsFilterState.portfolio === "active") {
        return !isDealClosedLifecycle(lifecycleStatus);
      }
      if (dealsFilterState.portfolio === "closed") {
        return isDealClosedLifecycle(lifecycleStatus);
      }
      return true;
    }

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

    function escapeHtml(value) {
      return String(value == null ? "" : value)
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;")
        .replace(/'/g, "&#39;");
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

    function getRequestedDealReference() {
      try {
        const params = new URLSearchParams(window.location.search || "");
        return String(params.get("deal") || params.get("id") || "").trim();
      } catch {
        return "";
      }
    }

    function applyRouteDealFilterFromUrl() {
      const reference = getRequestedDealReference();
      if (!reference) return;
      dealsFilterState = {
        ...dealsFilterState,
        portfolio: "all",
        stage: "all",
        owner: "all",
        keyword: reference.toLowerCase(),
      };
    }

    function hasDashboardLinked(deal) {
      return Boolean(String(deal && deal.fundraisingDashboardId || "").trim());
    }

    function hasDeckLinked(deal) {
      if (DealIntegrity && typeof DealIntegrity.hasDeckLinked === "function") {
        return DealIntegrity.hasDeckLinked(deal);
      }
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

    function getDealRetainerState(deal) {
      if (AppCore && typeof AppCore.getDealRetainerState === "function") {
        return AppCore.getDealRetainerState(deal);
      }
      const rawValue = String(deal && (deal.retainerMonthly != null ? deal.retainerMonthly : deal && deal.Retainer) || "").trim();
      const amount = Number(rawValue.replace(/,/g, "").replace(/[^0-9.-]/g, ""));
      const hasRetainer = Number.isFinite(amount) && amount > 0;
      return {
        rawValue,
        amount: hasRetainer ? amount : 0,
        hasRetainer,
        bucket: hasRetainer ? "with-retainer" : "no-retainer",
        label: hasRetainer ? "With retainer" : "0 / no retainer",
      };
    }

    function matchesDealRetainerFilter(deal, filterValue) {
      if (AppCore && typeof AppCore.matchesDealRetainerFilter === "function") {
        return AppCore.matchesDealRetainerFilter(deal, filterValue);
      }
      const normalized = normalizeValue(filterValue);
      if (!normalized || normalized === "all") return true;
      return getDealRetainerState(deal).bucket === normalized;
    }

    function sortDealsByRetainerState(source, fallbackComparator) {
      if (AppCore && typeof AppCore.sortDealsByRetainerState === "function") {
        return AppCore.sortDealsByRetainerState(source, fallbackComparator);
      }
      return (Array.isArray(source) ? source.slice() : []).sort((left, right) => {
        const leftState = getDealRetainerState(left);
        const rightState = getDealRetainerState(right);
        if (leftState.hasRetainer !== rightState.hasRetainer) {
          return leftState.hasRetainer ? -1 : 1;
        }
        return typeof fallbackComparator === "function" ? fallbackComparator(left, right) : 0;
      });
    }

    function clearDealIntegrityCache() {
      dealIntegrityCache = new Map();
    }

    function getOwnershipSnapshotEntry(deal) {
      if (!deal || !deal.id) return null;
      return ownershipSnapshotState.entriesByDealId.get(normalizeValue(deal.id)) || null;
    }

    function getDealIntegrityReport(deal) {
      const key = normalizeValue(deal && deal.id);
      if (!key) {
        return DealIntegrity && typeof DealIntegrity.buildDealIntegrityReport === "function"
          ? DealIntegrity.buildDealIntegrityReport(deal)
          : { counts: { total: 0, ownership: 0, accounting: 0, legal: 0, setup: 0 }, topIssues: [], fullyLinked: false };
      }
      if (dealIntegrityCache.has(key)) return dealIntegrityCache.get(key);

      const report = DealIntegrity && typeof DealIntegrity.buildDealIntegrityReport === "function"
        ? DealIntegrity.buildDealIntegrityReport(deal, {
          ownershipAvailable: ownershipSnapshotState.loaded,
          ownershipEntry: getOwnershipSnapshotEntry(deal),
        })
        : { counts: { total: 0, ownership: 0, accounting: 0, legal: 0, setup: 0 }, topIssues: [], fullyLinked: false };

      dealIntegrityCache.set(key, report);
      return report;
    }

    function getDealKeywords(deal) {
      const source = deal && deal.keywords;
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

    function normalizeDealSectors(deal) {
      const source = deal && (
        Array.isArray(deal.sectors)
          ? deal.sectors
          : typeof deal.sectors === "string"
            ? deal.sectors.split(/[\n,;]+/)
            : typeof deal.sector === "string"
              ? deal.sector.split(/[\n,;]+/)
              : []
      );
      const normalized = Array.from(
        new Set(
          (Array.isArray(source) ? source : [])
            .map((entry) => String(entry || "").trim())
            .filter(Boolean),
        ),
      );
      if (deal && typeof deal === "object") {
        deal.sectors = normalized;
        deal.sector = normalized.join(", ");
      }
      return normalized;
    }

    function parseDealSectorsInput(value) {
      return Array.from(
        new Set(
          String(value || "")
            .split(/[\n,;]+/)
            .map((entry) => String(entry || "").trim())
            .filter(Boolean)
        )
      );
    }

    function parseDealKeywordsInput(value) {
      return Array.from(
        new Set(
          String(value || "")
            .split(/[\n,;]+/)
            .map((entry) => String(entry || "").trim())
            .filter(Boolean)
        )
      );
    }

    function getDealProfileSummaryParts(deal) {
      const sectorLabel = normalizeDealSectors(deal).join(", ");
      return [
        sectorLabel,
        String(deal && deal.location || "").trim(),
        String(deal && deal.fundingStage || "").trim(),
        String(deal && deal.revenue || "").trim(),
      ].filter(Boolean);
    }

    function buildDealCompanyCell(deal) {
      const company = String(deal && deal.company || "").trim() || "–";
      const meta = getDealProfileSummaryParts(deal);
      if (!meta.length) return escapeHtml(company);
      return `
        <div class="deal-company-stack">
          <span class="deal-company-name">${escapeHtml(company)}</span>
          <span class="deal-company-meta">${escapeHtml(meta.join(" · "))}</span>
        </div>
      `;
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
      const portfolioFilter = document.getElementById("deal-portfolio-filter");
      const stageFilter = document.getElementById("deal-stage-filter");
      const ownerFilter = document.getElementById("deal-owner-filter");
      const retainerFilter = document.getElementById("deal-retainer-filter");
      const setupFilter = document.getElementById("deal-setup-filter");

      dealsFilterState = {
        portfolio: String(portfolioFilter && portfolioFilter.value || "active").trim().toLowerCase() || "active",
        keyword: String(searchInput && searchInput.value || "").trim().toLowerCase(),
        stage: String(stageFilter && stageFilter.value || "all").trim().toLowerCase() || "all",
        owner: String(ownerFilter && ownerFilter.value || "all").trim().toLowerCase() || "all",
        retainer: String(retainerFilter && retainerFilter.value || "all").trim().toLowerCase() || "all",
        setup: String(setupFilter && setupFilter.value || "all").trim().toLowerCase() || "all",
      };
    }

    function applyDealFilterStateToUi() {
      const searchInput = document.getElementById("deal-search-input");
      const portfolioFilter = document.getElementById("deal-portfolio-filter");
      const stageFilter = document.getElementById("deal-stage-filter");
      const ownerFilter = document.getElementById("deal-owner-filter");
      const retainerFilter = document.getElementById("deal-retainer-filter");
      const setupFilter = document.getElementById("deal-setup-filter");

      if (searchInput) searchInput.value = dealsFilterState.keyword || "";
      if (portfolioFilter) portfolioFilter.value = dealsFilterState.portfolio || "active";
      if (stageFilter) stageFilter.value = dealsFilterState.stage || "all";
      if (ownerFilter) ownerFilter.value = dealsFilterState.owner || "all";
      if (retainerFilter) retainerFilter.value = dealsFilterState.retainer || "all";
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
      if (!matchesPortfolioFilter(deal)) return false;
      if (!matchesDealRetainerFilter(deal, dealsFilterState.retainer)) return false;

      const keyword = dealsFilterState.keyword;
      if (keyword) {
        const haystack = [
          deal.id,
          deal.name,
          deal.company,
          getDealKeywords(deal).join(" "),
          normalizeDealSectors(deal).join(" "),
          deal.location,
          deal.fundingStage,
          deal.revenue,
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
      const integrityReport = getDealIntegrityReport(deal);
      if (dealsFilterState.setup === "missing-dashboard" && hasDashboard) return false;
      if (dealsFilterState.setup === "missing-deck" && hasDeck) return false;
      if (dealsFilterState.setup === "missing-legal" && hasLegal) return false;
      if (dealsFilterState.setup === "missing-any" && hasDashboard && hasDeck && hasLegal) return false;
      if (dealsFilterState.setup === "ready" && !(hasDashboard && hasDeck && hasLegal)) return false;
      if (dealsFilterState.setup === "irregularities" && !integrityReport.counts.total) return false;
      if (dealsFilterState.setup === "ownership-attention" && !integrityReport.counts.ownership) return false;
      if (dealsFilterState.setup === "accounting-attention" && !integrityReport.counts.accounting) return false;
      if (dealsFilterState.setup === "legal-attention" && !integrityReport.counts.legal) return false;
      if (dealsFilterState.setup === "fully-linked" && !integrityReport.fullyLinked) return false;

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
      const allLabel = dealsFilterState.portfolio === "closed"
        ? "Finished / closed"
        : dealsFilterState.portfolio === "all"
          ? "All deals"
          : "Active deals";
      const chips = [
        { value: "all", label: allLabel, count: counts.total },
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
      const portfolioFilter = document.getElementById("deal-portfolio-filter");
      const ownerFilter = document.getElementById("deal-owner-filter");
      const retainerFilter = document.getElementById("deal-retainer-filter");
      const setupFilter = document.getElementById("deal-setup-filter");

      const total = (Array.isArray(dealsData) ? dealsData : []).filter((deal) => matchesPortfolioFilter(deal)).length;
      const shown = Array.isArray(filteredDeals) ? filteredDeals.length : 0;
      const summaryParts = [`Showing ${shown} of ${total} deal${total === 1 ? "" : "s"}`];

      if (dealsFilterState.portfolio !== "active") {
        const portfolioLabel = portfolioFilter && portfolioFilter.selectedOptions && portfolioFilter.selectedOptions[0]
          ? portfolioFilter.selectedOptions[0].text
          : dealsFilterState.portfolio;
        summaryParts.push(`portfolio: ${portfolioLabel}`);
      }
      if (dealsFilterState.stage !== "all") {
        summaryParts.push(`stage: ${STAGE_LABELS[dealsFilterState.stage] || dealsFilterState.stage}`);
      }
      if (dealsFilterState.owner !== "all") {
        const ownerLabel = ownerFilter && ownerFilter.selectedOptions && ownerFilter.selectedOptions[0]
          ? ownerFilter.selectedOptions[0].text
          : dealsFilterState.owner;
        summaryParts.push(`owner: ${ownerLabel}`);
      }
      if (dealsFilterState.retainer !== "all") {
        const retainerLabel = retainerFilter && retainerFilter.selectedOptions && retainerFilter.selectedOptions[0]
          ? retainerFilter.selectedOptions[0].text
          : dealsFilterState.retainer.replace(/-/g, " ");
        summaryParts.push(`retainer: ${retainerLabel}`);
      }
      if (dealsFilterState.setup !== "all") {
        const setupLabel = setupFilter && setupFilter.selectedOptions && setupFilter.selectedOptions[0]
          ? setupFilter.selectedOptions[0].text
          : dealsFilterState.setup.replace(/-/g, " ");
        summaryParts.push(`connections: ${setupLabel}`);
      }
      if (dealsFilterState.keyword) {
        summaryParts.push(`search: "${dealsFilterState.keyword}"`);
      }

      summaryEl.textContent = `${summaryParts.join(" / ")}.`;
    }

    function setupDealFilters() {
      const searchInput = document.getElementById("deal-search-input");
      const portfolioFilter = document.getElementById("deal-portfolio-filter");
      const stageFilter = document.getElementById("deal-stage-filter");
      const ownerFilter = document.getElementById("deal-owner-filter");
      const retainerFilter = document.getElementById("deal-retainer-filter");
      const setupFilter = document.getElementById("deal-setup-filter");
      const resetBtn = document.getElementById("btn-reset-deal-filters");
      const metaRow = document.getElementById("meta-row");
      const integrityRow = document.getElementById("deal-integrity-row");

      [searchInput, portfolioFilter, stageFilter, ownerFilter, retainerFilter, setupFilter].forEach((control) => {
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
            portfolio: "active",
            keyword: "",
            stage: "all",
            owner: "all",
            retainer: "all",
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

      if (integrityRow) {
        integrityRow.addEventListener("click", (event) => {
          const chip = event.target.closest("[data-deal-setup-chip]");
          if (!chip) return;
          const nextSetup = String(chip.getAttribute("data-deal-setup-chip") || "all").trim().toLowerCase() || "all";
          dealsFilterState.setup = dealsFilterState.setup === nextSetup && nextSetup !== "all" ? "all" : nextSetup;
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
      const sectorInput = document.getElementById("deal-sector-input");
      const locationInput = document.getElementById("deal-location-input");
      const fundingStageInput = document.getElementById("deal-funding-stage-input");
      const revenueInput = document.getElementById("deal-revenue-input");
      const keywordsInput = document.getElementById("deal-keywords-input");
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
        const sectors = parseDealSectorsInput(sectorInput && sectorInput.value);
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
            lifecycleStatus: "active",
            seniorOwner,
            juniorOwner,
            subOwners,
            sectors,
            sector: sectors.join(", "),
            location: String(locationInput && locationInput.value || "").trim(),
            fundingStage: String(fundingStageInput && fundingStageInput.value || "").trim(),
            revenue: String(revenueInput && revenueInput.value || "").trim(),
            keywords: parseDealKeywordsInput(keywordsInput && keywordsInput.value),
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

    function getOwnershipDashboardConfig() {
      const config = getDashboardConfig();
      if (AppCore && typeof AppCore.getDashboardById === "function") {
        return AppCore.getDashboardById(config, "deal-ownership");
      }
      return Array.isArray(config && config.dashboards)
        ? config.dashboards.find((entry) => normalizeValue(entry && entry.id) === "deal-ownership") || null
        : null;
    }

    function sanitizeWorkbookRows(rows) {
      if (!Array.isArray(rows)) return [];
      return rows.map((row) => {
        const cleaned = {};
        Object.entries(row || {}).forEach(([key, value]) => {
          const header = String(key || "").trim();
          if (!header || /^__EMPTY/i.test(header)) return;
          cleaned[header] = value;
        });
        return cleaned;
      });
    }

    function getRowValueByHeaderPattern(row, patterns) {
      const headers = Object.keys(row || {});
      const matchHeader = headers.find((header) => patterns.some((pattern) => pattern.test(header)));
      return matchHeader ? String(row[matchHeader] || "").trim() : "";
    }

    function getPrimaryDealLabel(row, headers) {
      const prioritizedHeader = headers.find((header) => /deal/i.test(header))
        || headers.find((header) => /company/i.test(header))
        || headers.find((header) => /project/i.test(header));
      return prioritizedHeader ? String(row[prioritizedHeader] || "").trim() : "";
    }

    function getStaffLabel(row) {
      return getRowValueByHeaderPattern(row, [/^staff$/i, /^name$/i, /^lead$/i, /staff/i, /owner/i]) || "Unassigned";
    }

    function getRoleLabel(row) {
      return getRowValueByHeaderPattern(row, [/^role$/i, /title/i, /position/i, /function/i, /team/i]);
    }

    function getTopRoleSummary(summary, limit = 2) {
      if (!summary || !summary.roles || !summary.roles.size) return "Roles not tagged";
      return Array.from(summary.roles.entries())
        .sort((a, b) => b[1] - a[1] || a[0].localeCompare(b[0]))
        .slice(0, limit)
        .map(([role, count]) => count > 1 ? `${role} x${count}` : role)
        .join(", ");
    }

    function findDealForOwnershipRow(row, headers) {
      const references = headers
        .filter((header) => /deal|company|project/i.test(header))
        .map((header) => row[header])
        .filter(Boolean);

      if (AppCore && typeof AppCore.findDealByReference === "function") {
        return AppCore.findDealByReference(dealsData, references);
      }
      return (Array.isArray(dealsData) ? dealsData : []).find((deal) => references.some((entry) => normalizeValue(entry) === normalizeValue(deal && deal.id)));
    }

    async function resolveOwnershipWorkbookUrl(url) {
      const normalized = String(url || "").toLowerCase();
      const isSharePointLink = normalized.includes("sharepoint.com") || normalized.includes("1drv.ms");
      if (isSharePointLink && AppCore && typeof AppCore.resolveShareDriveDownloadUrl === "function") {
        try {
          const resolved = await AppCore.resolveShareDriveDownloadUrl(url);
          if (resolved) return resolved;
        } catch (error) {
          console.warn("[DealsOverview] Ownership URL resolution failed, using original URL", error);
        }
      }
      return url;
    }

    function buildDownloadHintUrl(url) {
      const raw = String(url || "").trim();
      if (!raw) return raw;
      if (/([?&])download=1(?:&|$)/i.test(raw)) return raw;
      return raw.includes("?") ? `${raw}&download=1` : `${raw}?download=1`;
    }

    function looksLikeHtmlBuffer(buffer) {
      try {
        const bytes = new Uint8Array(buffer || new ArrayBuffer(0));
        const sample = bytes.slice(0, 512);
        const text = new TextDecoder("utf-8").decode(sample).trim().toLowerCase();
        return text.startsWith("<!doctype html") || text.startsWith("<html") || text.includes("<head") || text.includes("<body");
      } catch {
        return false;
      }
    }

    async function downloadWorkbookBuffer(url) {
      return AppCore && typeof AppCore.downloadBinary === "function"
        ? AppCore.downloadBinary(url, { cache: "no-store" })
        : fetch(url).then((response) => {
          if (!response.ok) throw new Error("Failed to download ownership workbook");
          return response.arrayBuffer();
        });
    }

    async function fetchWorkbook(url) {
      const attempts = [String(url || "").trim()].filter(Boolean);
      const hintedUrl = buildDownloadHintUrl(url);
      if (hintedUrl && hintedUrl !== attempts[0]) attempts.push(hintedUrl);

      let lastError = null;
      for (let index = 0; index < attempts.length; index += 1) {
        const attemptUrl = attempts[index];
        try {
          const buffer = await downloadWorkbookBuffer(attemptUrl);
          if (looksLikeHtmlBuffer(buffer)) {
            throw new Error("Downloaded content is HTML instead of an Excel file.");
          }
          return XLSX.read(buffer, { type: "array" });
        } catch (error) {
          lastError = error;
        }
      }

      throw lastError || new Error("Failed to download ownership workbook.");
    }

    function buildOwnershipSnapshot(rows) {
      const headers = rows.length ? Object.keys(rows[0]) : [];
      const summaries = new Map();

      rows.forEach((row) => {
        const match = findDealForOwnershipRow(row, headers);
        const primaryDealLabel = getPrimaryDealLabel(row, headers) || "Unlabelled staffing line";
        const summaryKey = match && match.id
          ? `deal:${String(match.id)}`
          : `unlinked:${String(primaryDealLabel || "").trim().toLowerCase()}`;

        if (!summaries.has(summaryKey)) {
          summaries.set(summaryKey, {
            key: summaryKey,
            assignments: 0,
            staff: new Set(),
            roles: new Map(),
            label: primaryDealLabel,
            match,
          });
        }

        const summary = summaries.get(summaryKey);
        summary.assignments += 1;
        const staffLabel = getStaffLabel(row);
        const roleLabel = getRoleLabel(row);
        if (staffLabel) summary.staff.add(staffLabel);
        if (roleLabel) {
          summary.roles.set(roleLabel, (summary.roles.get(roleLabel) || 0) + 1);
        }
      });

      const entries = Array.from(summaries.values()).map((summary) => {
        const staffNames = Array.from(summary.staff || []).sort((left, right) => left.localeCompare(right));
        return {
          ...summary,
          staffNames,
          staffCount: staffNames.length,
          roleSummary: getTopRoleSummary(summary, 3),
          linkStatus: summary.match
            ? (staffNames.length ? "linked" : "missing-staffing")
            : "unlinked",
        };
      });

      return {
        entriesByDealId: new Map(
          entries
            .filter((entry) => entry.match && entry.match.id)
            .map((entry) => [normalizeValue(entry.match.id), entry])
        ),
        unlinkedEntries: entries.filter((entry) => entry.linkStatus === "unlinked"),
        linkedDeals: entries.filter((entry) => entry.match && entry.linkStatus !== "unlinked").length,
        rows: rows.length,
      };
    }

    function renderOwnershipConnectionBanner() {
      const banner = document.getElementById("deal-integrity-banner");
      if (!banner) return;

      if (ownershipSnapshotState.loading) {
        banner.className = "integrity-banner is-visible";
        banner.innerHTML = `
          <div class="integrity-banner-title">Checking ownership coverage</div>
          <div class="integrity-banner-copy">Pulling the staffing workbook so deal overview can flag missing ownership coverage and orphan staffing clusters.</div>
        `;
        return;
      }

      if (ownershipSnapshotState.error) {
        banner.className = "integrity-banner is-visible is-warning";
        banner.innerHTML = `
          <div class="integrity-banner-title">Ownership connection needs attention</div>
          <div class="integrity-banner-copy">${escapeHtml(ownershipSnapshotState.error)} Deal overview is still showing accounting and legal irregularities, but staffing linkage checks are temporarily limited.</div>
        `;
        return;
      }

      if (!ownershipSnapshotState.configured) {
        banner.className = "integrity-banner is-visible";
        banner.innerHTML = `
          <div class="integrity-banner-title">Ownership dashboard not configured yet</div>
          <div class="integrity-banner-copy">Add a deal-ownership dashboard config entry to raise staffing irregularities directly from deal overview.</div>
        `;
        return;
      }

      const unlinkedCount = ownershipSnapshotState.unlinkedEntries.length;
      banner.className = "integrity-banner is-visible";
      banner.innerHTML = `
        <div class="integrity-banner-title">Ownership sync is live</div>
        <div class="integrity-banner-copy">
          ${ownershipSnapshotState.linkedDeals} deal${ownershipSnapshotState.linkedDeals === 1 ? "" : "s"} matched from ${ownershipSnapshotState.rows} staffing row${ownershipSnapshotState.rows === 1 ? "" : "s"}.
          ${unlinkedCount ? `${unlinkedCount} staffing cluster${unlinkedCount === 1 ? " is" : "s are"} still unmatched and should be reviewed in Deal Ownership.` : "No orphan staffing clusters are currently detected."}
        </div>
      `;
    }

    async function refreshOwnershipSnapshot() {
      clearDealIntegrityCache();
      ownershipSnapshotState = {
        ...ownershipSnapshotState,
        loading: true,
        loaded: false,
        error: "",
      };
      renderOwnershipConnectionBanner();

      const dashboard = getOwnershipDashboardConfig();
      if (!dashboard || !dashboard.excelUrl) {
        ownershipSnapshotState = {
          loading: false,
          loaded: false,
          error: "",
          configured: false,
          entriesByDealId: new Map(),
          unlinkedEntries: [],
          linkedDeals: 0,
          rows: 0,
        };
        renderDeals();
        return;
      }

      if (typeof XLSX === "undefined") {
        ownershipSnapshotState = {
          loading: false,
          loaded: false,
          error: "SheetJS is not available, so the ownership workbook could not be parsed.",
          configured: true,
          entriesByDealId: new Map(),
          unlinkedEntries: [],
          linkedDeals: 0,
          rows: 0,
        };
        renderDeals();
        return;
      }

      try {
        const downloadUrl = await resolveOwnershipWorkbookUrl(dashboard.excelUrl);
        const workbook = await fetchWorkbook(downloadUrl);
        const sheetName = (dashboard.sheets && dashboard.sheets.funds) || workbook.SheetNames[0];
        const sheet = workbook.Sheets[sheetName] || workbook.Sheets[workbook.SheetNames[0]];
        if (!sheet) throw new Error("Ownership staffing sheet not found.");

        const rows = sanitizeWorkbookRows(XLSX.utils.sheet_to_json(sheet, { defval: "" }));
        const snapshot = buildOwnershipSnapshot(rows);
        ownershipSnapshotState = {
          loading: false,
          loaded: true,
          error: "",
          configured: true,
          ...snapshot,
        };
      } catch (error) {
        ownershipSnapshotState = {
          loading: false,
          loaded: false,
          error: error instanceof Error ? error.message : "Failed to load ownership workbook.",
          configured: true,
          entriesByDealId: new Map(),
          unlinkedEntries: [],
          linkedDeals: 0,
          rows: 0,
        };
      }

      clearDealIntegrityCache();
      renderDeals();
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

    function renderIntegrityPill(label, state) {
      const className = state === "good"
        ? "deal-integrity-pill is-good"
        : state === "attention"
          ? "deal-integrity-pill is-warning"
          : "deal-integrity-pill is-muted";
      return `<span class="${className}">${escapeHtml(label)}</span>`;
    }

    function renderDealIntegrityCell(deal) {
      const report = getDealIntegrityReport(deal);
      const summary = report.topIssues && report.topIssues.length
        ? report.topIssues.map((issue) => issue.title).join(" · ")
        : "Ownership, accounting, legal, and setup are aligned.";

      const ownershipState = ownershipSnapshotState.loaded
        ? report.statuses.ownership
        : (report.counts.ownership ? "attention" : "unknown");

      return `
        <div class="deal-integrity-cell">
          <div class="deal-integrity-pill-row">
            ${renderIntegrityPill("Ownership", ownershipState)}
            ${renderIntegrityPill("Accounting", report.statuses.accounting)}
            ${renderIntegrityPill("Legal", report.statuses.legal)}
          </div>
          <div class="deal-integrity-summary">
            <strong>${report.counts.total ? `${report.counts.total} irregularit${report.counts.total === 1 ? "y" : "ies"}` : "Fully connected"}</strong>
            <span> · ${escapeHtml(summary)}</span>
          </div>
        </div>
      `;
    }

    function renderDealActions(deal) {
      const detailHref = buildPageUrl("deal-details", { id: deal.id });
      const ownershipHref = buildPageUrl("deal-ownership", { deal: deal.id });
      const accountingHref = buildPageUrl("accounting", { id: deal.id });
      const legalHref = buildPageUrl("legal-management", { id: deal.id });
      const dashboard = getDealDashboard(deal);
      const links = [
        `<a class="action-link" href="${detailHref}">Deal record</a>`,
        `<a class="action-link" href="${ownershipHref}">Ownership</a>`,
        `<a class="action-link" href="${legalHref}">Legal</a>`,
      ];

      if (!accountingAccessState.restricted || accountingAccessState.allowed) {
        links.push(`<a class="action-link" href="${accountingHref}">Accounting</a>`);
      }

      if (dashboard && dashboard.id) {
        const dashboardHref = buildPageUrl("investor-dashboard", { dashboard: dashboard.id });
        links.push(`<a class="action-link" href="${dashboardHref}">Open dashboard</a>`);
      }

      return `<div class="deal-action-cluster">${links.join('<span style="color: var(--text-dim);">·</span>')}</div>`;
    }

    function renderIntegrityFilterChips(source) {
      const row = document.getElementById("deal-integrity-row");
      if (!row) return;

      const deals = Array.isArray(source) ? source : [];
      const counts = {
        all: deals.length,
        irregularities: 0,
        "ownership-attention": 0,
        "accounting-attention": 0,
        "legal-attention": 0,
        "fully-linked": 0,
      };

      deals.forEach((deal) => {
        const report = getDealIntegrityReport(deal);
        if (report.counts.total) counts.irregularities += 1;
        if (report.counts.ownership) counts["ownership-attention"] += 1;
        if (report.counts.accounting) counts["accounting-attention"] += 1;
        if (report.counts.legal) counts["legal-attention"] += 1;
        if (report.fullyLinked) counts["fully-linked"] += 1;
      });

      const chips = [
        { value: "all", label: "All", count: counts.all },
        { value: "irregularities", label: "Any irregularity", count: counts.irregularities },
        { value: "ownership-attention", label: "Ownership", count: counts["ownership-attention"] },
        { value: "accounting-attention", label: "Accounting", count: counts["accounting-attention"] },
        { value: "legal-attention", label: "Legal", count: counts["legal-attention"] },
        { value: "fully-linked", label: "Fully connected", count: counts["fully-linked"] },
      ];

      row.innerHTML = chips.map((chip) => `
        <button class="chip deal-integrity-chip${dealsFilterState.setup === chip.value ? " is-active" : ""}" type="button" data-deal-setup-chip="${chip.value}">
          <strong>${chip.count}</strong> ${chip.label}
        </button>
      `).join("");
    }

    function renderDeals() {
      const body = document.getElementById("deals-body");
      const metaRow = document.getElementById("meta-row");
      const footerMeta = document.getElementById("footer-meta");
      loadDealsData();
      if (!Array.isArray(dealsData)) return;
      clearDealIntegrityCache();
      populateDealOwnerFilter();
      applyDealFilterStateToUi();

      const filteredDeals = sortDealsByRetainerState(
        getFilteredDeals(),
        (left, right) => normalizeValue(left && (left.company || left.name || left.id)).localeCompare(
          normalizeValue(right && (right.company || right.name || right.id)),
        ),
      );
      const portfolioDeals = (Array.isArray(dealsData) ? dealsData : []).filter((deal) => matchesPortfolioFilter(deal));
      const allCounts = buildDealStageCounts(portfolioDeals);

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
        const lifecycleStatus = getDealLifecycleStatus(deal);
        const retainerState = getDealRetainerState(deal);
        const stageCell = `
          <div class="deal-stage-cell">
            <span class="stage-badge">
              <span class="stage-dot ${stageClass(stage)}"></span>
              ${STAGE_LABELS[stage] || "Prospect"}
            </span>
            ${lifecycleStatus !== "active" ? renderLifecycleBadge(lifecycleStatus) : ""}
          </div>
        `;

        tr.innerHTML = `
          <td class="name-cell">
            <a class="action-link" href="${buildPageUrl("deal-details", { id: deal.id })}">${deal.name || "–"}</a>
          </td>
          <td>${buildDealCompanyCell(deal)}</td>
          <td>${stageCell}</td>
          <td>${getPeopleInChargeText(deal)}</td>
          <td class="amount">${formatAmountCell(deal.targetAmount, deal.currency)}</td>
          <td class="amount">${formatAmountCell(deal.raisedAmount, deal.currency)}</td>
          <td>${formatTextOrUnknown(deal.CashCommission)}</td>
          <td>${formatTextOrUnknown(deal.EquityCommission)}</td>
          <td>${retainerState.hasRetainer ? formatTextOrUnknown(retainerState.rawValue) : `<span class="chip">${escapeHtml(retainerState.label)}</span>`}</td>
          <td>${renderDealIntegrityCell(deal)}</td>
          <td>${renderDealActions(deal)}</td>
        `;

        body.appendChild(tr);
      });

      if (!filteredDeals.length) {
        body.innerHTML = '<tr><td colspan="11" class="amount-unknown">No deals match the current filters.</td></tr>';
      }

      renderStageFilterChips(metaRow, allCounts);
      renderIntegrityFilterChips(portfolioDeals);
      renderOwnershipConnectionBanner();
      updateDealFilterSummary(filteredDeals);
      const portfolioLabel = dealsFilterState.portfolio === "closed"
        ? "finished / closed deal"
        : dealsFilterState.portfolio === "all"
          ? "deal"
          : "active deal";
      const flaggedCount = portfolioDeals.filter((deal) => getDealIntegrityReport(deal).counts.total > 0).length;
      const footerBase = filteredDeals.length === portfolioDeals.length
        ? `${allCounts.total} ${portfolioLabel}${allCounts.total === 1 ? "" : "s"} shown.`
        : `${filteredDeals.length} shown of ${allCounts.total} ${portfolioLabel}${allCounts.total === 1 ? "" : "s"}.`;
      footerMeta.textContent = `${footerBase} ${flaggedCount} ${flaggedCount === 1 ? "deal has" : "deals have"} active irregularities.`;
    }

    document.addEventListener("DOMContentLoaded", () => {
      const initialize = async () => {
        if (AppCore && typeof AppCore.refreshDashboardConfigFromShareDrive === "function") {
          await AppCore.refreshDashboardConfigFromShareDrive();
        }
        await refreshAccountingAccessState();
        loadDealsData();
        applyRouteDealFilterFromUrl();
        setupDealFilters();
        setupAddDealForm();
        renderDeals();
        refreshOwnershipSnapshot();
      };
      initialize();
      if (AppCore) {
        window.addEventListener("appcore:deals-updated", () => {
          clearDealIntegrityCache();
          renderDeals();
        });
        window.addEventListener("appcore:dashboard-config-updated", () => {
          refreshOwnershipSnapshot();
        });
        window.addEventListener("appcore:graph-session-updated", async () => {
          await refreshAccountingAccessState();
          renderDeals();
          refreshOwnershipSnapshot();
        });
      }
    });
