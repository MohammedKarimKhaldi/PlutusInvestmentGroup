(function initDealStaffingPage() {
  const STORAGE_KEY =
    (window.AppCore && window.AppCore.STORAGE_KEYS && window.AppCore.STORAGE_KEYS.staffing) ||
    "deal_staffing_v1";
  const PROXIES = window.DASHBOARD_PROXIES || [];
  const STAFFING_SETTINGS =
    (window.DASHBOARD_CONFIG && window.DASHBOARD_CONFIG.settings && window.DASHBOARD_CONFIG.settings.staffing) || {};
  const DEALS_STORAGE_KEY =
    (window.AppCore && window.AppCore.STORAGE_KEYS && window.AppCore.STORAGE_KEYS.deals) || "deals_data_v1";

  const REFRESH_MS = 5 * 60 * 1000;
  let workbookData = { sheets: [], activeSheet: "", rows: [], columns: [] };
  let linkedDeals = [];
  let activeDealIdFilter = "";
  let activeDealKeyword = "";

  const refs = {
    status: document.getElementById("sync-status"),
    refreshBtn: document.getElementById("btn-refresh"),
    search: document.getElementById("search"),
    sheetSelect: document.getElementById("sheet-select"),
    head: document.getElementById("staffing-head"),
    body: document.getElementById("staffing-body"),
    kTotalRows: document.getElementById("k-total-rows"),
    kUniqueDeals: document.getElementById("k-unique-deals"),
    kOpenRoles: document.getElementById("k-open-roles"),
    kLastSync: document.getElementById("k-last-sync"),
  };

  function normalizeKey(value) {
    return String(value || "").replace(/\s+/g, " ").trim();
  }

  function asLower(value) {
    return String(value || "").trim().toLowerCase();
  }

  function normalizeForMatch(value) {
    return asLower(value).replace(/[^a-z0-9]+/g, " ").replace(/\s+/g, " ").trim();
  }

  function isHiddenColumnKey(key) {
    const normalized = asLower(key);
    return normalized.startsWith("__empty") || normalized.startsWith("__");
  }

  function loadLinkedDeals() {
    if (window.AppCore && typeof window.AppCore.loadDealsData === "function") {
      const loaded = window.AppCore.loadDealsData();
      return Array.isArray(loaded) ? loaded : [];
    }
    return Array.isArray(window.DEALS) ? window.DEALS : [];
  }

  function findLinkedDealId(row) {
    const rawDeal = normalizeForMatch(row && row.Deal);
    if (!rawDeal || !linkedDeals.length) return "";

    for (const deal of linkedDeals) {
      const id = normalizeForMatch(deal.id);
      const name = normalizeForMatch(deal.name);
      const company = normalizeForMatch(deal.company);
      if (!id && !name && !company) continue;

      const exact = rawDeal === id || rawDeal === name || rawDeal === company;
      const close =
        (name && (rawDeal.includes(name) || name.includes(rawDeal))) ||
        (company && (rawDeal.includes(company) || company.includes(rawDeal)));

      if (exact || close) return String(deal.id || "").trim();
    }
    return "";
  }

  function readUrlFilters() {
    const params = new URLSearchParams(window.location.search);
    activeDealIdFilter = asLower(params.get("dealId"));

    const dealName = normalizeKey(params.get("deal"));
    const companyName = normalizeKey(params.get("company"));
    const keyword = dealName || companyName;
    activeDealKeyword = keyword;
    if (keyword) refs.search.value = keyword;
  }

  function applyDealFilterStatus() {
    if (!activeDealIdFilter) return;
    refs.status.textContent = activeDealKeyword
      ? `Showing staffing for "${activeDealKeyword}"`
      : "Showing staffing for selected deal";
  }

  function detectHeaderRow(sheet, maxScanRows = 30) {
    if (!sheet) return 0;
    const range = XLSX.utils.decode_range(sheet["!ref"] || "A1:A1");
    const upperBound = Math.min(range.e.r, maxScanRows - 1);

    let bestRow = 0;
    let bestScore = -1;

    for (let r = range.s.r; r <= upperBound; r++) {
      let nonEmpty = 0;
      let score = 0;
      for (let c = range.s.c; c <= range.e.c; c++) {
        const cellRef = XLSX.utils.encode_cell({ r, c });
        const cell = sheet[cellRef];
        const value = normalizeKey(cell && cell.v);
        if (!value) continue;
        nonEmpty++;
        const low = value.toLowerCase();
        if (
          low.includes("deal") ||
          low.includes("company") ||
          low.includes("role") ||
          low.includes("staff") ||
          low.includes("owner") ||
          low.includes("status")
        ) {
          score += 2;
        }
      }
      score += nonEmpty;
      if (score > bestScore) {
        bestScore = score;
        bestRow = r;
      }
    }

    return bestRow;
  }

  function normalizeRows(rows) {
    return rows
      .map((row) => {
        const next = {};
        Object.keys(row || {}).forEach((key) => {
          const normalized = normalizeKey(key);
          if (!normalized) return;
          if (isHiddenColumnKey(normalized)) return;
          next[normalized] = row[key];
        });
        const linkedDealId = findLinkedDealId(next);
        if (linkedDealId) next.__linkedDealId = linkedDealId;
        return next;
      })
      .filter((row) => Object.keys(row).length > 0);
  }

  function collectColumns(rows) {
    const ordered = [];
    const seen = new Set();
    (rows || []).forEach((row) => {
      Object.keys(row || {}).forEach((key) => {
        if (isHiddenColumnKey(key)) return;
        if (seen.has(key)) return;
        seen.add(key);
        ordered.push(key);
      });
    });
    return ordered;
  }

  function getDealValue(row) {
    const keys = Object.keys(row || {});
    const preferred =
      keys.find((k) => asLower(k) === "deal") ||
      keys.find((k) => asLower(k).includes("deal")) ||
      keys.find((k) => asLower(k).includes("company")) ||
      "";
    return preferred ? normalizeKey(row[preferred]) : "";
  }

  function isOpenRole(row) {
    const keys = Object.keys(row || {});
    const statusKey = keys.find((k) => asLower(k).includes("status"));
    const status = asLower(statusKey ? row[statusKey] : "");
    if (!status) return false;
    return status.includes("open") || status.includes("vacant") || status.includes("pending");
  }

  function saveCache(payload) {
    try {
      localStorage.setItem(STORAGE_KEY, JSON.stringify(payload));
    } catch (error) {
      console.warn("[Staffing] Failed to write cache", error);
    }
  }

  function loadCache() {
    try {
      const raw = localStorage.getItem(STORAGE_KEY);
      if (!raw) return null;
      const parsed = JSON.parse(raw);
      if (!parsed || !Array.isArray(parsed.sheets)) return null;
      return parsed;
    } catch (error) {
      console.warn("[Staffing] Failed to read cache", error);
      return null;
    }
  }

  async function fetchWorkbook(url) {
    for (let i = 0; i < PROXIES.length; i++) {
      const fetchUrl = PROXIES[i](url);
      try {
        const response = await fetch(fetchUrl);
        if (!response.ok) throw new Error(`HTTP ${response.status}`);

        let buffer;
        if (fetchUrl.includes("allorigins")) {
          const json = await response.json();
          const b64 = (json && json.contents && json.contents.split(",")[1]) || (json && json.contents) || "";
          const binary = atob(b64);
          buffer = new ArrayBuffer(binary.length);
          const view = new Uint8Array(buffer);
          for (let n = 0; n < binary.length; n++) view[n] = binary.charCodeAt(n);
        } else {
          buffer = await response.arrayBuffer();
        }

        return XLSX.read(buffer, { type: "array" });
      } catch (error) {
        console.warn(`[Staffing] Proxy ${i} failed`, error);
      }
    }
    throw new Error("All staffing proxies failed.");
  }

  function parseWorkbook(wb) {
    const sheetNames = wb.SheetNames || [];
    const sheets = sheetNames.map((name) => {
      const sheet = wb.Sheets[name];
      const headerRow = detectHeaderRow(sheet);
      const rows = normalizeRows(XLSX.utils.sheet_to_json(sheet, { range: headerRow, defval: "" }));
      const columns = collectColumns(rows);
      return { name, rows, columns };
    });
    return sheets;
  }

  function updateSheetSelect() {
    refs.sheetSelect.innerHTML = "";
    workbookData.sheets.forEach((sheet) => {
      const option = document.createElement("option");
      option.value = sheet.name;
      option.textContent = sheet.name;
      option.selected = sheet.name === workbookData.activeSheet;
      refs.sheetSelect.appendChild(option);
    });
  }

  function updateKPIs(rows, fetchedAt) {
    refs.kTotalRows.textContent = String(rows.length);
    const deals = new Set();
    let openRoles = 0;
    rows.forEach((row) => {
      const deal = getDealValue(row);
      if (deal) deals.add(deal);
      if (isOpenRole(row)) openRoles += 1;
    });
    refs.kUniqueDeals.textContent = String(deals.size);
    refs.kOpenRoles.textContent = String(openRoles);
    refs.kLastSync.textContent = fetchedAt ? new Date(fetchedAt).toLocaleString() : "—";
  }

  function renderTable(rows, columns) {
    linkedDeals = loadLinkedDeals();
    refs.head.innerHTML = "";
    refs.body.innerHTML = "";
    if (!rows.length) return;

    const tableColumns = Array.isArray(columns) && columns.length ? columns : collectColumns(rows);

    refs.head.innerHTML = `<tr>${tableColumns.map((col) => `<th>${col}</th>`).join("")}</tr>`;

    const keyword = asLower(refs.search.value);
    rows
      .filter((row) => {
        if (!activeDealIdFilter) return true;
        return asLower(row.__linkedDealId) === activeDealIdFilter;
      })
      .filter((row) => {
        if (!keyword) return true;
        return JSON.stringify(row).toLowerCase().includes(keyword);
      })
      .forEach((row) => {
        const cells = tableColumns
          .map((col) => {
            const value = row[col] || "—";
            if (col === "Deal" && row.__linkedDealId) {
              return `<td><a class="action-link" href="deal-details.html?id=${encodeURIComponent(
                row.__linkedDealId,
              )}">${value}</a></td>`;
            }
            return `<td>${value}</td>`;
          })
          .join("");
        refs.body.insertAdjacentHTML("beforeend", `<tr>${cells}</tr>`);
      });
  }

  function setActiveSheet(name, fetchedAt) {
    const sheet =
      workbookData.sheets.find((item) => item.name === name) ||
      workbookData.sheets[0] ||
      { rows: [], columns: [] };
    workbookData.activeSheet = sheet.name || "";
    workbookData.rows = sheet.rows || [];
    workbookData.columns = sheet.columns || [];
    updateSheetSelect();
    updateKPIs(workbookData.rows, fetchedAt);
    renderTable(workbookData.rows, workbookData.columns);
    applyDealFilterStatus();
  }

  async function syncStaffing() {
    linkedDeals = loadLinkedDeals();
    const staffingUrl = normalizeKey(STAFFING_SETTINGS.excelUrl);
    if (!staffingUrl) {
      refs.status.textContent = "Missing staffing URL in config";
      return;
    }

    refs.status.textContent = "Syncing staffing...";
    try {
      const wb = await fetchWorkbook(staffingUrl);
      const sheets = parseWorkbook(wb);
      const preferred = normalizeKey(STAFFING_SETTINGS.sheetName);
      const firstName = (sheets[0] && sheets[0].name) || "";
      const activeSheet =
        (preferred && sheets.find((s) => asLower(s.name) === asLower(preferred)) && preferred) || firstName;

      const payload = {
        sourceUrl: staffingUrl,
        fetchedAt: new Date().toISOString(),
        activeSheet,
        sheets,
      };
      saveCache(payload);

      workbookData.sheets = sheets;
      setActiveSheet(activeSheet, payload.fetchedAt);
      refs.status.textContent = "Live staffing sync active";
      applyDealFilterStatus();
    } catch (error) {
      console.error("[Staffing] Sync failed", error);
      const cached = loadCache();
      if (cached) {
        workbookData.sheets = cached.sheets;
        setActiveSheet(cached.activeSheet || (cached.sheets[0] && cached.sheets[0].name) || "", cached.fetchedAt);
        refs.status.textContent = "Using cached staffing data";
        applyDealFilterStatus();
      } else {
        refs.status.textContent = "Staffing sync failed";
        refs.body.innerHTML = "<tr><td>Unable to load staffing data.</td></tr>";
      }
    }
  }

  refs.refreshBtn.addEventListener("click", syncStaffing);
  refs.search.addEventListener("input", () => renderTable(workbookData.rows, workbookData.columns));
  refs.sheetSelect.addEventListener("change", (event) => setActiveSheet(event.target.value, new Date().toISOString()));

  window.addEventListener("storage", (event) => {
    if (event.key === STORAGE_KEY) {
      const cached = loadCache();
      if (!cached) return;
      workbookData.sheets = cached.sheets;
      setActiveSheet(cached.activeSheet || (cached.sheets[0] && cached.sheets[0].name) || "", cached.fetchedAt);
      refs.status.textContent = "Updated from shared cache";
      return;
    }

    if (event.key === DEALS_STORAGE_KEY) {
      renderTable(workbookData.rows, workbookData.columns);
      refs.status.textContent = "Deal links refreshed";
    }
  });

  readUrlFilters();
  syncStaffing();
  window.setInterval(syncStaffing, REFRESH_MS);
})();
