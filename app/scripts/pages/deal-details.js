const AppCore = window.AppCore;
const DEALS_STORAGE_KEY = (AppCore && AppCore.STORAGE_KEYS && AppCore.STORAGE_KEYS.deals) || "deals_data_v1";
const TASKS_STORAGE_KEY = (AppCore && AppCore.STORAGE_KEYS && AppCore.STORAGE_KEYS.tasks) || "owner_tasks_v1";

let allDeals = [];
let allTasks = [];
let currentDeal = null;
let groupMode = "owner";
const AUTO_CONTACT_TASK_PREFIX = (AppCore && AppCore.AUTO_CONTACT_TASK_PREFIX) || "auto-contact-status";

const STAGE_LABELS = {
  prospect: "Prospect",
  signing: "Signing",
  onboarding: "Onboarding",
  "contacting investors": "Contacting investors",
};
const STAGE_ORDER = ["prospect", "signing", "onboarding", "contacting investors"];
const SHAREDRIVE_URL_STORAGE_KEY = "sharedrive_url_v2";
const deckPickerState = {
  shareUrl: "",
  root: null,
  stack: [],
  items: [],
};

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

function loadDealsData() {
  allDeals = AppCore ? AppCore.loadDealsData() : (Array.isArray(DEALS) ? JSON.parse(JSON.stringify(DEALS)) : []);
}

function saveDealsData() {
  if (AppCore) {
    return AppCore.saveDealsData(allDeals);
  }
  try {
    localStorage.setItem(DEALS_STORAGE_KEY, JSON.stringify(allDeals));
  } catch (e) {
    console.warn("Failed to save deals to storage", e);
  }
  return Promise.resolve();
}

function loadTasksForDeal() {
  allTasks = AppCore ? AppCore.loadTasksData() : (Array.isArray(TASKS) ? JSON.parse(JSON.stringify(TASKS)) : []);
}

function saveTasksData() {
  if (AppCore) {
    AppCore.saveTasksData(allTasks);
    return;
  }
  try {
    localStorage.setItem(TASKS_STORAGE_KEY, JSON.stringify(allTasks));
  } catch (e) {
    console.warn("Failed to save tasks to storage", e);
  }
}

function stageClass(stage) {
  const s = normalizeValue(stage);
  if (s === "prospect") return "stage-prospect";
  if (s === "signing") return "stage-signing";
  if (s === "onboarding") return "stage-onboarding";
  if (s === "contacting investors") return "stage-contacting";
  return "stage-prospect";
}

function relatedTasksForCurrentDeal() {
  if (!currentDeal) return [];
  return allTasks.filter((task) => normalizeValue(task.dealId) === normalizeValue(currentDeal.id));
}

function isSigningTask(task) {
  const hay = `${task && task.title || ""} ${task && task.type || ""} ${task && task.notes || ""}`;
  const normalized = normalizeValue(hay);
  return (
    normalized.includes("sign") ||
    normalized.includes("signature") ||
    normalized.includes("contract") ||
    normalized.includes("legal")
  );
}

function getSigningTaskProgress() {
  const signingTasks = relatedTasksForCurrentDeal().filter(isSigningTask);
  const doneCount = signingTasks.filter((task) => normalizeValue(task.status) === "done").length;
  return {
    total: signingTasks.length,
    done: doneCount,
    allDone: signingTasks.length === 0 || doneCount === signingTasks.length,
  };
}

function advanceDealStage(nextStage) {
  if (!currentDeal) return;
  currentDeal.stage = nextStage;
  saveDealsData();
  renderDealHeader();
  populateDealForm();
  renderRelatedTasks();
}

function refreshStageButton() {
  const button = document.getElementById("btn-complete-stage");
  if (!button || !currentDeal) return;

  const stage = normalizeValue(currentDeal.stage);
  const signing = getSigningTaskProgress();

  if (stage === "prospect") {
    button.disabled = false;
    button.textContent = "Move to signing";
    return;
  }
  if (stage === "signing") {
    button.disabled = false;
    button.textContent = "Move to onboarding";
    return;
  }
  if (stage === "onboarding") {
    button.disabled = !signing.allDone;
    button.textContent = signing.total === 0
      ? "Move to contacting investors"
      : (signing.allDone
        ? "Move to contacting investors"
        : `Complete signing tasks (${signing.done}/${signing.total})`);
    return;
  }

  button.disabled = true;
  button.textContent = "Final stage reached";
}

function setupStageProgressionButton() {
  const button = document.getElementById("btn-complete-stage");
  if (!button) return;

  button.addEventListener("click", () => {
    if (!currentDeal) return;
    const stage = normalizeValue(currentDeal.stage);
    const signing = getSigningTaskProgress();

    if (stage === "prospect") {
      advanceDealStage("signing");
      return;
    }
    if (stage === "signing") {
      advanceDealStage("onboarding");
      return;
    }
    if (stage === "onboarding") {
      if (!signing.allDone) {
        window.alert(`Complete signing tasks first (${signing.done}/${signing.total} done).`);
        return;
      }
      advanceDealStage("contacting investors");
      return;
    }
  });
}

function autoPromoteFromOnboardingIfReady() {
  if (!currentDeal) return false;
  if (normalizeValue(currentDeal.stage) !== "onboarding") return false;

  const signing = getSigningTaskProgress();
  if (!(signing.total > 0 && signing.allDone)) return false;

  currentDeal.stage = "contacting investors";
  saveDealsData();
  return true;
}

function getSeniorOwner(deal) {
  return (deal && (deal.seniorOwner || deal.owner)) || "";
}

function getJuniorOwner(deal) {
  return (deal && deal.juniorOwner) || "";
}

function formatAmount(amount, currency) {
  if (amount == null || amount === "") return "–";
  if (typeof amount === "string") return amount;
  if (typeof amount === "number" && !isNaN(amount)) {
    try {
      return new Intl.NumberFormat("en-US", {
        style: "currency",
        currency: currency || "USD",
        maximumFractionDigits: 0,
      }).format(amount);
    } catch {
      return `${currency || "USD"} ${Number(amount).toLocaleString()}`;
    }
  }
  return "–";
}

function formatDate(value) {
  if (!value) return "No due date";
  const dt = new Date(value);
  if (isNaN(dt.getTime())) return value;
  return dt.toLocaleDateString("en-GB", { day: "2-digit", month: "short" });
}

function escapeHtml(value) {
  return String(value == null ? "" : value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

function getDashboardConfig() {
  if (AppCore && typeof AppCore.getDashboardConfig === "function") {
    return AppCore.getDashboardConfig();
  }
  return window.DASHBOARD_CONFIG || { dashboards: [], settings: {} };
}

function getDashboardForCurrentDeal() {
  const dashboardsConfig = getDashboardConfig();
  let dashboard =
    (window.AppCore && typeof window.AppCore.getDashboardForDeal === "function")
      ? window.AppCore.getDashboardForDeal(currentDeal, dashboardsConfig)
      : null;
  if (!dashboard && currentDeal) {
    const dashboardId = normalizeValue(currentDeal.fundraisingDashboardId);
    const dashboards = Array.isArray(dashboardsConfig.dashboards) ? dashboardsConfig.dashboards : [];
    dashboard = dashboards.find((entry) => normalizeValue(entry.id) === dashboardId) || null;
  }
  return dashboard;
}

function getDefaultDeckShareUrl() {
  try {
    const remembered = String(localStorage.getItem(SHAREDRIVE_URL_STORAGE_KEY) || "").trim();
    if (remembered) return remembered;
  } catch {
    // ignore storage failures
  }

  const sharedriveTasksUrl =
    window.SHAREDRIVE_TASKS &&
    window.SHAREDRIVE_TASKS.tasks &&
    typeof window.SHAREDRIVE_TASKS.tasks.shareUrl === "string"
      ? window.SHAREDRIVE_TASKS.tasks.shareUrl.trim()
      : "";
  if (sharedriveTasksUrl) return sharedriveTasksUrl;

  const config = getDashboardConfig();
  return (
    config &&
    config.settings &&
    config.settings.sharedDeals &&
    typeof config.settings.sharedDeals.shareUrl === "string"
      ? config.settings.sharedDeals.shareUrl.trim()
      : ""
  );
}

function getDeckLabel(deal) {
  if (!deal) return "No deck linked";
  return String(deal.deckName || deal.deckUrl || "").trim() || "No deck linked";
}

function refreshDeckEditorSummary() {
  const nameInput = document.getElementById("deal-input-deck-name");
  const urlInput = document.getElementById("deal-input-deck-url");
  const pathInput = document.getElementById("deal-input-deck-parent-path");
  const nameEl = document.getElementById("deal-deck-editor-name");
  const metaEl = document.getElementById("deal-deck-editor-meta");

  if (!nameInput || !urlInput || !pathInput || !nameEl || !metaEl) return;

  const deckName = String(nameInput.value || "").trim();
  const deckUrl = String(urlInput.value || "").trim();
  const parentPath = String(pathInput.value || "").trim();

  if (!deckUrl) {
    nameEl.textContent = "No deck selected";
    metaEl.textContent = "Browse Sharedrive and select a PDF deck for this deal.";
    return;
  }

  nameEl.textContent = deckName || deckUrl;
  metaEl.textContent = parentPath || deckUrl;
}

function applySelectedDeck(item) {
  const nameInput = document.getElementById("deal-input-deck-name");
  const urlInput = document.getElementById("deal-input-deck-url");
  const pathInput = document.getElementById("deal-input-deck-parent-path");
  const statusEl = document.getElementById("deal-save-status");
  if (!nameInput || !urlInput || !pathInput) return;

  nameInput.value = String((item && item.name) || "").trim();
  urlInput.value = String((item && item.webUrl) || "").trim();
  pathInput.value = String((item && item.parentPath) || "").trim();
  refreshDeckEditorSummary();

  if (statusEl) {
    statusEl.textContent = "Deck selected in form. Save deal updates to keep the link.";
  }
}

function closeDeckPicker() {
  const modal = document.getElementById("deck-picker-modal");
  if (modal) modal.hidden = true;
}

function setDeckPickerStatus(message) {
  const statusEl = document.getElementById("deck-picker-status");
  if (statusEl) statusEl.textContent = message || "";
}

function renderDeckPickerPath() {
  const pathEl = document.getElementById("deck-picker-path");
  if (!pathEl) return;

  if (!deckPickerState.root) {
    pathEl.textContent = "Not connected to a Sharedrive folder yet.";
    return;
  }

  const names = deckPickerState.stack.map((entry) => entry.name).filter(Boolean);
  pathEl.textContent = names.join(" / ") || deckPickerState.root.name || "Sharedrive root";
}

function isPdfItem(item) {
  const name = String((item && item.name) || "").trim();
  const mimeType = String((item && item.mimeType) || "").trim().toLowerCase();
  return /\.pdf$/i.test(name) || mimeType === "application/pdf";
}

function renderDeckPickerItems(items) {
  const listEl = document.getElementById("deck-picker-list");
  if (!listEl) return;

  const sorted = (Array.isArray(items) ? items.slice() : []).sort((a, b) => {
    if (Boolean(a.isFolder) !== Boolean(b.isFolder)) return a.isFolder ? -1 : 1;
    return String(a.name || "").localeCompare(String(b.name || ""));
  });
  deckPickerState.items = sorted;

  if (!sorted.length) {
    listEl.innerHTML = '<div class="deal-task-item"><span class="task-notes">No folders or PDF files found here.</span></div>';
    return;
  }

  listEl.innerHTML = sorted
    .filter((item) => item.isFolder || isPdfItem(item))
    .map((item) => {
      const title = String(item.name || "(Unnamed item)");
      const meta = item.isFolder
        ? `${item.childCount != null ? `${item.childCount} item${item.childCount === 1 ? "" : "s"}` : "Folder"}`
        : `${item.mimeType || "PDF file"}${item.parentPath ? ` · ${item.parentPath}` : ""}`;
      const actionLabel = item.isFolder ? "Open folder" : "Use this PDF";
      return `
        <div class="deck-picker-item">
          <div class="deck-picker-item-copy">
            <div class="deck-picker-item-title">${escapeHtml(title)}</div>
            <div class="deck-picker-item-meta">${escapeHtml(meta)}</div>
          </div>
          <div class="deck-picker-item-actions">
            <button class="btn" type="button" data-deck-item-id="${escapeHtml(item.id)}" data-deck-item-kind="${item.isFolder ? "folder" : "pdf"}">${actionLabel}</button>
          </div>
        </div>
      `;
    })
    .join("");

  if (!listEl.children.length) {
    listEl.innerHTML = '<div class="deal-task-item"><span class="task-notes">This folder has files, but none of them are PDFs.</span></div>';
  }
}

async function requestDeckPickerChildren(parentItemId) {
  const shareUrl = String(deckPickerState.shareUrl || "").trim();
  if (!shareUrl) {
    throw new Error("Provide a Sharedrive folder URL first.");
  }

  if (window.PlutusDesktop && typeof window.PlutusDesktop.listShareDriveChildren === "function") {
    const result = await window.PlutusDesktop.listShareDriveChildren({
      shareUrl,
      parentItemId: parentItemId || "",
    });
    if (!result || !result.ok) {
      throw new Error((result && result.error) || "Failed to load Sharedrive items.");
    }
    return result.data || {};
  }

  if (AppCore && typeof AppCore.listShareDriveChildren === "function") {
    return AppCore.listShareDriveChildren({ shareUrl, parentItemId: parentItemId || "" });
  }

  throw new Error("Sharedrive browsing is unavailable on this platform.");
}

async function loadDeckPickerFolder(targetEntry, options = {}) {
  const { resetStack = false } = options;
  const folderId = targetEntry && targetEntry.id ? targetEntry.id : "";

  setDeckPickerStatus("Loading Sharedrive folder...");
  const data = await requestDeckPickerChildren(folderId);
  deckPickerState.root = data.root || deckPickerState.root;

  if (resetStack) {
    const root = data.root || {};
    deckPickerState.stack = [{ id: data.parentItemId || root.id || "", name: root.name || "Sharedrive root" }];
  } else if (targetEntry && targetEntry.id) {
    deckPickerState.stack.push({ id: targetEntry.id, name: targetEntry.name || "Folder" });
  }

  renderDeckPickerPath();
  renderDeckPickerItems(Array.isArray(data.items) ? data.items : []);
  setDeckPickerStatus(`Showing ${Array.isArray(data.items) ? data.items.length : 0} item(s)`);
}

function openDeckPicker() {
  const modal = document.getElementById("deck-picker-modal");
  const shareUrlInput = document.getElementById("deck-share-url-input");
  if (!modal || !shareUrlInput) return;

  modal.hidden = false;
  deckPickerState.shareUrl = deckPickerState.shareUrl || getDefaultDeckShareUrl();
  shareUrlInput.value = deckPickerState.shareUrl;
  renderDeckPickerPath();

  if (!deckPickerState.shareUrl) {
    setDeckPickerStatus("Paste a Sharedrive folder URL to start browsing.");
    return;
  }

  loadDeckPickerFolder(null, { resetStack: true }).catch((error) => {
    setDeckPickerStatus(error instanceof Error ? error.message : "Failed to load Sharedrive items.");
  });
}

function refreshDashboardLinkEditor() {
  const select = document.getElementById("deal-input-dashboard-select");
  const input = document.getElementById("deal-input-dashboard");
  const help = document.getElementById("deal-dashboard-link-help");
  if (!select || !input || !help) return;

  const dashboardsConfig = getDashboardConfig();
  const dashboards = Array.isArray(dashboardsConfig.dashboards) ? dashboardsConfig.dashboards.slice() : [];
  const sortedDashboards = dashboards.sort((a, b) =>
    String((a && (a.name || a.id)) || "").localeCompare(String((b && (b.name || b.id)) || ""))
  );

  const currentValue = String(input.value || "").trim();
  const currentKey = normalizeValue(currentValue);
  const matchedDashboard = sortedDashboards.find((entry) => normalizeValue(entry && entry.id) === currentKey) || null;

  select.innerHTML = ['<option value="">No linked dashboard</option>']
    .concat(sortedDashboards.map((entry) => {
      const id = String(entry && entry.id ? entry.id : "").trim();
      const name = String(entry && (entry.name || entry.id) ? (entry.name || entry.id) : id).trim();
      return `<option value="${id}">${name} (${id})</option>`;
    }))
    .concat('<option value="__custom__">Custom dashboard ID</option>')
    .join("");

  if (!currentValue) {
    select.value = "";
    help.textContent = "This deal is currently not linked to any fundraising dashboard.";
  } else if (matchedDashboard) {
    select.value = matchedDashboard.id;
    help.textContent = `Linked to ${matchedDashboard.name || matchedDashboard.id}. You can switch to another dashboard or clear the link.`;
  } else {
    select.value = "__custom__";
    help.textContent = `Custom dashboard ID: ${currentValue}. Keep this if you want to link the deal before creating the dashboard profile.`;
  }
}

function parseNumericAmount(value) {
  if (value == null) return null;
  const raw = String(value).trim();
  if (!raw) return null;

  const cleaned = raw.replace(/,/g, "");
  const parsed = Number(cleaned);
  if (!Number.isNaN(parsed) && Number.isFinite(parsed)) return parsed;
  return raw;
}

function detectHeaderRow(sheet, maxScanRows = 20) {
  if (!sheet) return 3;
  const range = XLSX.utils.decode_range(sheet["!ref"] || "A1:A1");
  const upperBound = Math.min(range.e.r, maxScanRows - 1);

  for (let r = range.s.r; r <= upperBound; r++) {
    const rowValues = [];
    for (let c = range.s.c; c <= range.e.c; c++) {
      const cellAddress = XLSX.utils.encode_cell({ r, c });
      const cell = sheet[cellAddress];
      if (cell && cell.v !== undefined && cell.v !== null && String(cell.v).trim() !== "") {
        rowValues.push(String(cell.v).trim().toLowerCase());
      }
    }
    if (rowValues.includes("investor")) return r;
  }
  return 3;
}

function getContactStatusKey(row) {
  const forward = normalizeValue(row["Moving Forward"]);
  const meeting = normalizeValue(row["Meeting with Company"]);
  const call = normalizeValue(row["Call/Meeting"]);

  if (forward === "yes") return "moving_forward";
  if (forward === "waiting" || meeting === "waiting" || call === "waiting") return "waiting";
  if (forward === "no") return "passed";
  if (meeting === "yes") return "meeting_done";
  if (call === "yes") return "contact_started";
  return "target";
}

function toIdFragment(value) {
  return String(value || "")
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, "-")
    .replace(/^-+|-+$/g, "")
    .slice(0, 80);
}

function buildAutoInvestorTasks(rows) {
  const createdAt = new Date().toISOString().slice(0, 10);

  return rows
    .map((row) => {
      const investor = String(row["Investor"] || "").trim();
      if (!investor) return null;

      const statusKey = getContactStatusKey(row);

      let category = null;
      let taskStatus = "in progress";
      let label = "";

      if (statusKey === "meeting_done") {
        category = "meeting";
        taskStatus = "in progress";
        label = "Meeting";
      } else if (statusKey === "target") {
        category = "target";
        taskStatus = "waiting";
        label = "Target";
      } else if (statusKey === "contact_started" || statusKey === "waiting" || statusKey === "moving_forward") {
        category = "ongoing";
        taskStatus = statusKey === "waiting" ? "waiting" : "in progress";
        label = "Ongoing";
      }

      if (!category) return null;

      const investorId = toIdFragment(investor);
      return {
        id: `${AUTO_CONTACT_TASK_PREFIX}-${currentDeal.id}-${category}-${investorId}`,
        owner: getSeniorOwner(currentDeal) || "System",
        dealId: currentDeal.id,
        title: `[Auto] ${label}: ${investor}`,
        type: `Contact ${label}`,
        status: taskStatus,
        dueDate: "",
        notes: `Synced from dashboard on ${createdAt}.`,
        metaSource: "dashboard-contact-status",
        metaInvestor: investor,
        metaCategory: category,
      };
    })
    .filter(Boolean);
}

async function fetchDashboardWorkbook(excelUrl, proxies) {
  if (!excelUrl || !Array.isArray(proxies) || !proxies.length) return null;

  const looksLikeSharePointLink = (url) => {
    if (!url) return false;
    const normalized = String(url).toLowerCase();
    return (
      normalized.includes("sharepoint.com/:") ||
      normalized.includes("sharepoint.com/_layouts/15/doc.aspx") ||
      normalized.includes("1drv.ms/")
    );
  };

  const resolveShareLinkDownloadUrl = async (url) => {
    if (!looksLikeSharePointLink(url)) return null;
    if (window.AppCore && typeof window.AppCore.resolveShareDriveDownloadUrl === "function") {
      return window.AppCore.resolveShareDriveDownloadUrl(url);
    }
    return null;
  };

  const fetchWorkbookFromUrl = async (fetchUrl, expectsJsonWrapper) => {
    let buffer;
    if (expectsJsonWrapper) {
      const response = await fetch(fetchUrl);
      if (!response.ok) throw new Error(`Download failed`);
      const json = await response.json();
      if (!json.contents) throw new Error("No contents in allorigins response");
      const b64 = json.contents.split(",")[1] || json.contents;
      const binaryString = atob(b64);
      buffer = new ArrayBuffer(binaryString.length);
      const view = new Uint8Array(buffer);
      for (let j = 0; j < binaryString.length; j++) view[j] = binaryString.charCodeAt(j);
    } else {
      buffer = AppCore && typeof AppCore.downloadBinary === "function"
        ? await AppCore.downloadBinary(fetchUrl, { cache: "no-store" })
        : await fetch(fetchUrl).then((response) => {
            if (!response.ok) throw new Error("Download failed");
            return response.arrayBuffer();
          });
    }

    return XLSX.read(buffer, { type: "array" });
  };

  try {
    const resolved = await resolveShareLinkDownloadUrl(excelUrl);
    if (resolved) {
      return await fetchWorkbookFromUrl(resolved, false);
    }
  } catch (err) {
    console.warn("[Deal] SharePoint link resolution failed", err);
    return null;
  }

  for (let i = 0; i < proxies.length; i++) {
    try {
      const fetchUrl = proxies[i](excelUrl);
      return await fetchWorkbookFromUrl(fetchUrl, fetchUrl.includes("allorigins"));
    } catch (err) {
      console.warn(`[Deal] Dashboard sync proxy ${i} failed`, err);
    }
  }

  return null;
}

async function syncContactStatusTasksFromDashboard() {
  if (!currentDeal || !window.XLSX) return;
  const dashboardsConfig =
    (window.AppCore && typeof window.AppCore.getDashboardConfig === "function")
      ? window.AppCore.getDashboardConfig()
      : window.DASHBOARD_CONFIG;
  if (!dashboardsConfig || !Array.isArray(dashboardsConfig.dashboards)) return;

  const dashboardId = normalizeValue(currentDeal.fundraisingDashboardId);
  if (!dashboardId) return;

  const dashboard = dashboardsConfig.dashboards.find(
    (d) => normalizeValue(d.id) === dashboardId,
  );
  if (!dashboard || !dashboard.excelUrl) return;

  const workbook = await fetchDashboardWorkbook(
    dashboard.excelUrl,
    window.DASHBOARD_PROXIES || [],
  );
  if (!workbook) return;

  const sheetNames = dashboard.sheets || { funds: "funds", familyOffices: "f.o." };
  const getSheet = (name) => {
    if (!name) return null;
    const matchName = workbook.SheetNames.find((n) => normalizeValue(n).includes(normalizeValue(name)));
    return matchName ? workbook.Sheets[matchName] : null;
  };

  const fundsSheet = getSheet(sheetNames.funds);
  const foSheet = getSheet(sheetNames.familyOffices);
  if (!fundsSheet || !foSheet) return;

  const headerRowIndex = detectHeaderRow(fundsSheet);
  const vcRows = XLSX.utils.sheet_to_json(fundsSheet, { range: headerRowIndex, defval: "" }).filter((r) => r["Investor"]);
  const foRows = XLSX.utils.sheet_to_json(foSheet, { range: headerRowIndex, defval: "" }).filter((r) => r["Investor"]);
  const allRows = [...vcRows, ...foRows];

  const autoTasks = buildAutoInvestorTasks(allRows);

  allTasks = allTasks.filter(
    (task) => !(
      normalizeValue(task.dealId) === normalizeValue(currentDeal.id)
      && (normalizeValue(task.metaSource) === "dashboard-contact-status"
        || normalizeValue(task.id).startsWith(`${AUTO_CONTACT_TASK_PREFIX}-${normalizeValue(currentDeal.id)}-`))
    ),
  );

  allTasks.push(...autoTasks);
  saveTasksData();
  renderRelatedTasks();
}

function renderDealHeader() {
  if (!currentDeal) return;

  document.getElementById("crumb-deal-name").textContent = currentDeal.name || "Deal";
  document.getElementById("deal-name").textContent = currentDeal.name || "Deal";
  document.getElementById("deal-subtitle").textContent =
    `${currentDeal.company || "Company"} • ${STAGE_LABELS[normalizeValue(currentDeal.stage)] || "Prospect"}`;

  document.getElementById("deal-company").textContent = currentDeal.company || "–";
  document.getElementById("deal-senior").textContent = getSeniorOwner(currentDeal) || "–";
  document.getElementById("deal-junior").textContent = getJuniorOwner(currentDeal) || "–";
  document.getElementById("deal-target").textContent = formatAmount(currentDeal.targetAmount, currentDeal.currency);
  document.getElementById("deal-raised").textContent = formatAmount(currentDeal.raisedAmount, currentDeal.currency);

  const textOrDash = (value) => (value == null || String(value).trim() === "" ? "–" : String(value));
  document.getElementById("deal-cash-comm").textContent = textOrDash(currentDeal.CashCommission);
  document.getElementById("deal-equity-comm").textContent = textOrDash(currentDeal.EquityCommission);
  document.getElementById("deal-retainer").textContent = textOrDash(currentDeal.Retainer);

  const progressPercent =
    currentDeal.targetAmount && !isNaN(currentDeal.targetAmount) && currentDeal.raisedAmount && !isNaN(currentDeal.raisedAmount)
      ? Math.min(100, Math.round((currentDeal.raisedAmount / currentDeal.targetAmount) * 100))
      : 0;

  document.getElementById("deal-progress").style.width = `${progressPercent}%`;

  const stage = normalizeValue(currentDeal.stage);
  const stageLabel = STAGE_LABELS[stage] || "Prospect";
  document.getElementById("deal-stage-label").textContent = stageLabel;
  document.getElementById("deal-stage-dot").className = `stage-dot ${stageClass(stage)}`;
  renderStageTracker(stage);

  document.getElementById("deal-summary").textContent =
    currentDeal.summary ||
    "This deal does not yet have a custom description. Use this space to track the high-level context, goals, and key investors targeted for this round.";

  const btnDeck = document.getElementById("btn-open-deck");
  const deckNameEl = document.getElementById("deal-deck-name");
  const deckMetaEl = document.getElementById("deal-deck-meta");
  const deckUrl = String(currentDeal.deckUrl || "").trim();
  const deckName = getDeckLabel(currentDeal);

  if (deckNameEl) deckNameEl.textContent = deckName;
  if (deckMetaEl) {
    deckMetaEl.textContent = deckUrl
      ? (String(currentDeal.deckParentPath || "").trim() || deckUrl)
      : "Choose a PDF from Sharedrive and save the deal to keep the link.";
  }
  if (btnDeck) {
    if (deckUrl) {
      btnDeck.disabled = false;
      btnDeck.textContent = "Open linked deck";
      btnDeck.onclick = () => {
        window.open(deckUrl, "_blank", "noopener,noreferrer");
      };
    } else {
      btnDeck.disabled = true;
      btnDeck.textContent = "No deck linked";
      btnDeck.onclick = null;
    }
  }

  const btnDashboard = document.getElementById("btn-open-dashboard");
  const btnAccounting = document.getElementById("btn-open-accounting");
  const dashboard = getDashboardForCurrentDeal();

  if (btnAccounting) {
    if (currentDeal && currentDeal.id) {
      btnAccounting.disabled = false;
      btnAccounting.onclick = () => {
        window.location.href = buildPageUrl("accounting", { id: currentDeal.id });
      };
    } else {
      btnAccounting.disabled = true;
      btnAccounting.onclick = null;
    }
  }

  if (dashboard && dashboard.id) {
    btnDashboard.disabled = false;
    btnDashboard.textContent = "Open fundraising dashboard";
    btnDashboard.onclick = () => {
      window.location.href = buildPageUrl("investor-dashboard", { dashboard: dashboard.id });
    };
  } else {
    btnDashboard.disabled = true;
    btnDashboard.textContent = "No dashboard linked";
    btnDashboard.onclick = null;
  }

  refreshStageButton();
}

function renderStageTracker(stage) {
  const trackerValue = document.getElementById("deal-stage-tracker-value");
  const trackerFill = document.getElementById("deal-stage-track-fill");
  const stepNodes = Array.from(document.querySelectorAll(".deal-stage-step"));
  if (!trackerValue || !trackerFill || !stepNodes.length) return;

  const stageIndex = Math.max(0, STAGE_ORDER.indexOf(stage));
  const ratio = ((stageIndex + 1) / STAGE_ORDER.length) * 100;
  const label = STAGE_LABELS[stage] || "Prospect";

  trackerValue.textContent = `${label} · ${stageIndex + 1}/${STAGE_ORDER.length}`;
  trackerFill.style.width = `${ratio}%`;

  stepNodes.forEach((node, index) => {
    node.classList.toggle("active", index <= stageIndex);
  });
}

function populateDealForm() {
  if (!currentDeal) return;

  document.getElementById("deal-input-name").value = currentDeal.name || "";
  document.getElementById("deal-input-company").value = currentDeal.company || "";
  document.getElementById("deal-input-senior").value = getSeniorOwner(currentDeal);
  document.getElementById("deal-input-junior").value = getJuniorOwner(currentDeal);
  document.getElementById("deal-input-stage").value = normalizeValue(currentDeal.stage) || "prospect";
  document.getElementById("deal-input-target").value = currentDeal.targetAmount ?? "";
  document.getElementById("deal-input-raised").value = currentDeal.raisedAmount ?? "";
  document.getElementById("deal-input-currency").value = currentDeal.currency || "USD";
  document.getElementById("deal-input-dashboard").value = currentDeal.fundraisingDashboardId || "";
  refreshDashboardLinkEditor();
  document.getElementById("deal-input-cash").value = currentDeal.CashCommission || "";
  document.getElementById("deal-input-equity").value = currentDeal.EquityCommission || "";
  document.getElementById("deal-input-retainer").value = currentDeal.Retainer || "";
  document.getElementById("deal-input-summary").value = currentDeal.summary || "";
  document.getElementById("deal-input-deck-name").value = currentDeal.deckName || "";
  document.getElementById("deal-input-deck-url").value = currentDeal.deckUrl || "";
  document.getElementById("deal-input-deck-parent-path").value = currentDeal.deckParentPath || "";
  refreshDeckEditorSummary();
}

function renderRelatedTasks() {
  if (!currentDeal) return;

  const related = allTasks.filter((task) => normalizeValue(task.dealId) === normalizeValue(currentDeal.id));
  const listEl = document.getElementById("deal-task-list");
  const summaryEl = document.getElementById("deal-task-summary");

  if (!related.length) {
    listEl.innerHTML = '<div class="deal-task-item"><span class="task-notes">No tasks are linked to this deal yet.</span></div>';
    summaryEl.textContent = "No tasks currently linked to this deal.";
    refreshStageButton();
    return;
  }

  const today = new Date();
  today.setHours(0, 0, 0, 0);

  let overdueCount = 0;
  const groups = new Map();

  related.forEach((task) => {
    const status = normalizeValue(task.status);
    if (task.dueDate && status !== "done") {
      const d = new Date(task.dueDate);
      if (!isNaN(d.getTime())) {
        d.setHours(0, 0, 0, 0);
        if (d < today) overdueCount += 1;
      }
    }

    const key = groupMode === "type"
      ? (task.type && String(task.type).trim() ? String(task.type).trim() : "Uncategorized")
      : (task.owner && String(task.owner).trim() ? String(task.owner).trim() : "Unassigned");

    if (!groups.has(key)) groups.set(key, []);
    groups.get(key).push(task);
  });

  const renderTaskItem = (task) => {
    const status = normalizeValue(task.status);
    const statusLabel = status ? status.charAt(0).toUpperCase() + status.slice(1) : "In progress";

    let overdue = false;
    if (task.dueDate && status !== "done") {
      const d = new Date(task.dueDate);
      if (!isNaN(d.getTime())) {
        d.setHours(0, 0, 0, 0);
        overdue = d < today;
      }
    }

    const dueLabel = formatDate(task.dueDate) + (overdue ? ' · <span class="overdue-label">Overdue</span>' : "");
    const statusClassName =
      status === "waiting" ? "badge-status-waiting" :
        status === "done" ? "badge-status-done" : "badge-status-in-progress";

    const owner = task.owner || "Unassigned";
    const ownerLink = buildPageUrl("owner-tasks", { owner });

    return `
      <div class="deal-task-item">
        <div class="deal-task-header">
          <span class="deal-task-title">${task.title || "Task"}</span>
          <span class="${statusClassName}">${statusLabel}</span>
        </div>
        <div class="deal-task-meta">
          <span>${dueLabel}</span>
          <span>Owner: <a class="task-link" href="${ownerLink}">${owner}</a></span>
        </div>
        ${task.type ? `<div class="task-notes">Type: ${task.type}</div>` : ""}
        ${task.notes ? `<div class="task-notes">${task.notes}</div>` : ""}
      </div>
    `;
  };

  const groupsHtml = [...groups.entries()]
    .sort(([a], [b]) => a.localeCompare(b))
    .map(([groupName, items]) => {
      const doneCount = items.filter((task) => normalizeValue(task.status) === "done").length;
      const progress = items.length ? Math.round((doneCount / items.length) * 100) : 0;

      return `
        <section class="deal-task-group">
          <div class="deal-task-group-header">
            <div class="deal-task-group-title">${groupName}</div>
            <div class="deal-task-progress">
              <div class="deal-task-progress-row">
                <span>${doneCount}/${items.length} done</span>
                <span>${progress}%</span>
              </div>
              <div class="deal-task-progress-bar"><div class="deal-task-progress-fill" style="width:${progress}%;"></div></div>
            </div>
          </div>
          <div class="deal-task-group-list">
            ${items.map(renderTaskItem).join("")}
          </div>
        </section>
      `;
    })
    .join("");

  listEl.innerHTML = groupsHtml;
  summaryEl.textContent =
    `${related.length} task${related.length === 1 ? "" : "s"} linked to this deal · ${overdueCount} overdue · grouped by ${groupMode}.`;

  if (autoPromoteFromOnboardingIfReady()) {
    renderDealHeader();
    populateDealForm();
  } else {
    refreshStageButton();
  }
}

function setupGroupButtons() {
  const buttons = Array.from(document.querySelectorAll(".task-group-btn"));
  buttons.forEach((button) => {
    button.addEventListener("click", () => {
      buttons.forEach((btn) => btn.classList.remove("active"));
      button.classList.add("active");
      groupMode = button.getAttribute("data-group") || "owner";
      renderRelatedTasks();
    });
  });
}

function setupDealForm() {
  const form = document.getElementById("deal-form");
  const statusEl = document.getElementById("deal-save-status");
  const dashboardSelect = document.getElementById("deal-input-dashboard-select");
  const dashboardInput = document.getElementById("deal-input-dashboard");
  const clearLinkBtn = document.getElementById("btn-clear-dashboard-link");
  const browseDeckBtn = document.getElementById("btn-browse-deck");
  const clearDeckBtn = document.getElementById("btn-clear-deck-link");

  if (dashboardSelect && dashboardInput) {
    dashboardSelect.addEventListener("change", () => {
      if (dashboardSelect.value === "__custom__") {
        dashboardInput.focus();
      } else {
        dashboardInput.value = dashboardSelect.value;
      }
      refreshDashboardLinkEditor();
    });

    dashboardInput.addEventListener("input", () => {
      refreshDashboardLinkEditor();
    });
  }

  if (clearLinkBtn && dashboardInput) {
    clearLinkBtn.addEventListener("click", () => {
      dashboardInput.value = "";
      refreshDashboardLinkEditor();
      if (statusEl) statusEl.textContent = "Dashboard link cleared in form";
      window.setTimeout(() => {
        if (statusEl && statusEl.textContent === "Dashboard link cleared in form") {
          statusEl.textContent = "";
        }
      }, 1200);
    });
  }

  if (browseDeckBtn) {
    browseDeckBtn.addEventListener("click", () => {
      openDeckPicker();
    });
  }

  if (clearDeckBtn) {
    clearDeckBtn.addEventListener("click", () => {
      document.getElementById("deal-input-deck-name").value = "";
      document.getElementById("deal-input-deck-url").value = "";
      document.getElementById("deal-input-deck-parent-path").value = "";
      refreshDeckEditorSummary();
      if (statusEl) statusEl.textContent = "Deck link cleared in form";
      window.setTimeout(() => {
        if (statusEl && statusEl.textContent === "Deck link cleared in form") {
          statusEl.textContent = "";
        }
      }, 1200);
    });
  }

  form.addEventListener("submit", async (event) => {
    event.preventDefault();
    if (!currentDeal) return;

    currentDeal.name = document.getElementById("deal-input-name").value.trim();
    currentDeal.company = document.getElementById("deal-input-company").value.trim();
    currentDeal.seniorOwner = document.getElementById("deal-input-senior").value.trim();
    currentDeal.juniorOwner = document.getElementById("deal-input-junior").value.trim();
    currentDeal.owner = currentDeal.seniorOwner; // legacy compatibility
    currentDeal.stage = document.getElementById("deal-input-stage").value;
    currentDeal.targetAmount = parseNumericAmount(document.getElementById("deal-input-target").value);
    currentDeal.raisedAmount = parseNumericAmount(document.getElementById("deal-input-raised").value);
    currentDeal.currency = document.getElementById("deal-input-currency").value.trim() || "USD";
    currentDeal.fundraisingDashboardId = document.getElementById("deal-input-dashboard").value.trim();
    currentDeal.CashCommission = document.getElementById("deal-input-cash").value.trim();
    currentDeal.EquityCommission = document.getElementById("deal-input-equity").value.trim();
    currentDeal.Retainer = document.getElementById("deal-input-retainer").value.trim();
    currentDeal.summary = document.getElementById("deal-input-summary").value.trim();
    currentDeal.deckName = document.getElementById("deal-input-deck-name").value.trim();
    currentDeal.deckUrl = document.getElementById("deal-input-deck-url").value.trim();
    currentDeal.deckParentPath = document.getElementById("deal-input-deck-parent-path").value.trim();

    statusEl.textContent = "Saving...";
    try {
      await saveDealsData();
      renderDealHeader();
      populateDealForm();
      statusEl.textContent = currentDeal.fundraisingDashboardId || currentDeal.deckUrl ? "Saved and linked" : "Saved";
      window.setTimeout(() => {
        statusEl.textContent = "";
      }, 1500);
    } catch (error) {
      statusEl.textContent = error instanceof Error ? error.message : "Save failed";
    }
  });
}

function setupDealTaskForm() {
  const form = document.getElementById("deal-task-form");
  const statusEl = document.getElementById("task-save-status");

  if (currentDeal) {
    document.getElementById("task-input-owner").value = getSeniorOwner(currentDeal) || "";
  }

  form.addEventListener("submit", (event) => {
    event.preventDefault();
    if (!currentDeal) return;

    const titleInput = document.getElementById("task-input-title");
    const ownerInput = document.getElementById("task-input-owner");
    const typeInput = document.getElementById("task-input-type");
    const statusInput = document.getElementById("task-input-status");
    const dueInput = document.getElementById("task-input-due");
    const notesInput = document.getElementById("task-input-notes");

    const title = titleInput.value.trim();
    const owner = ownerInput.value.trim();
    if (!title || !owner) return;

    const newTask = {
      id: `t-${Date.now()}-${Math.floor(Math.random() * 1000)}`,
      owner,
      dealId: currentDeal.id,
      title,
      type: typeInput.value.trim() || "",
      status: statusInput.value || "in progress",
      dueDate: dueInput.value || "",
      notes: notesInput.value.trim() || "",
    };

    allTasks.push(newTask);
    saveTasksData();
    renderRelatedTasks();

    titleInput.value = "";
    typeInput.value = "";
    dueInput.value = "";
    notesInput.value = "";

    statusEl.textContent = "Task added";
    window.setTimeout(() => {
      statusEl.textContent = "";
    }, 1500);
  });
}

function showDealEditor() {
  const form = document.getElementById("deal-form");
  const toggle = document.getElementById("btn-toggle-deal-form");
  if (!form || !toggle || !form.hidden) return;
  form.hidden = false;
  toggle.textContent = "Hide editor";
}

function setupDeckPicker() {
  const modal = document.getElementById("deck-picker-modal");
  const shareUrlInput = document.getElementById("deck-share-url-input");
  const loadRootBtn = document.getElementById("btn-load-deck-root");
  const closeBtn = document.getElementById("btn-close-deck-picker");
  const backBtn = document.getElementById("btn-deck-go-back");
  const listEl = document.getElementById("deck-picker-list");
  const linkDeckBtn = document.getElementById("btn-link-deck");

  if (linkDeckBtn) {
    linkDeckBtn.addEventListener("click", () => {
      showDealEditor();
      openDeckPicker();
    });
  }

  if (closeBtn) {
    closeBtn.addEventListener("click", closeDeckPicker);
  }

  if (modal) {
    modal.addEventListener("click", (event) => {
      if (event.target === modal) closeDeckPicker();
    });
  }

  document.addEventListener("keydown", (event) => {
    if (event.key === "Escape" && modal && !modal.hidden) {
      closeDeckPicker();
    }
  });

  if (loadRootBtn && shareUrlInput) {
    loadRootBtn.addEventListener("click", async () => {
      deckPickerState.shareUrl = shareUrlInput.value.trim();
      if (!deckPickerState.shareUrl) {
        setDeckPickerStatus("Paste a Sharedrive folder URL first.");
        return;
      }
      try {
        localStorage.setItem(SHAREDRIVE_URL_STORAGE_KEY, deckPickerState.shareUrl);
      } catch {
        // ignore storage failures
      }
      try {
        await loadDeckPickerFolder(null, { resetStack: true });
      } catch (error) {
        setDeckPickerStatus(error instanceof Error ? error.message : "Failed to load Sharedrive items.");
      }
    });
  }

  if (backBtn) {
    backBtn.addEventListener("click", async () => {
      if (deckPickerState.stack.length <= 1) {
        setDeckPickerStatus("Already at the root folder.");
        return;
      }
      deckPickerState.stack.pop();
      const current = deckPickerState.stack[deckPickerState.stack.length - 1];
      try {
        setDeckPickerStatus("Loading previous folder...");
        const data = await requestDeckPickerChildren(current.id);
        renderDeckPickerPath();
        renderDeckPickerItems(Array.isArray(data.items) ? data.items : []);
        setDeckPickerStatus(`Showing ${Array.isArray(data.items) ? data.items.length : 0} item(s)`);
      } catch (error) {
        setDeckPickerStatus(error instanceof Error ? error.message : "Failed to load Sharedrive items.");
      }
    });
  }

  if (listEl) {
    listEl.addEventListener("click", async (event) => {
      const button = event.target.closest("button[data-deck-item-id]");
      if (!button) return;

      const itemId = String(button.getAttribute("data-deck-item-id") || "").trim();
      const kind = String(button.getAttribute("data-deck-item-kind") || "").trim();
      if (!itemId) return;
      const selectedItem = (deckPickerState.items || []).find((entry) => String(entry.id || "").trim() === itemId);
      if (!selectedItem) return;

      if (kind === "folder") {
        try {
          await loadDeckPickerFolder({ id: itemId, name: selectedItem.name || "Folder" });
        } catch (error) {
          setDeckPickerStatus(error instanceof Error ? error.message : "Failed to load Sharedrive items.");
        }
        return;
      }

      applySelectedDeck(selectedItem);
      closeDeckPicker();
    });
  }
}

function setupFormToggles() {
  const dealForm = document.getElementById("deal-form");
  const dealToggle = document.getElementById("btn-toggle-deal-form");
  const taskForm = document.getElementById("deal-task-form");
  const taskToggle = document.getElementById("btn-toggle-task-form");

  if (dealToggle && dealForm) {
    dealToggle.addEventListener("click", () => {
      const willShow = dealForm.hidden;
      dealForm.hidden = !willShow;
      dealToggle.textContent = willShow ? "Hide editor" : "Edit deal";
    });
  }

  if (taskToggle && taskForm) {
    taskToggle.addEventListener("click", () => {
      const willShow = taskForm.hidden;
      taskForm.hidden = !willShow;
      taskToggle.textContent = willShow ? "Hide form" : "Add task";
    });
  }
}

async function loadDealPage() {
  const params = new URLSearchParams(window.location.search);
  const id = params.get("id");

  const emptyState = document.getElementById("empty-state");
  const shell = document.getElementById("shell");

  if (AppCore) {
    if (typeof AppCore.refreshDashboardConfigFromShareDrive === "function") {
      await AppCore.refreshDashboardConfigFromShareDrive();
    }
    if (typeof AppCore.refreshDealsFromShareDrive === "function") {
      await AppCore.refreshDealsFromShareDrive("deal-details");
    }
    window.addEventListener("appcore:tasks-updated", (event) => {
      if (event && event.detail && Array.isArray(event.detail.tasks)) {
        allTasks = event.detail.tasks;
        if (currentDeal) renderRelatedTasks();
      }
    });
    window.addEventListener("appcore:deals-updated", (event) => {
      if (!event || !event.detail || !Array.isArray(event.detail.deals)) return;
      allDeals = event.detail.deals;
      currentDeal = allDeals.find((deal) => normalizeValue(deal.id) === normalizeValue(id)) || currentDeal;
      if (currentDeal) {
        document.getElementById("empty-state").style.display = "none";
        document.getElementById("shell").style.display = "block";
        renderDealHeader();
        populateDealForm();
        renderRelatedTasks();
      }
    });
    window.addEventListener("appcore:dashboard-config-updated", () => {
      if (currentDeal) {
        renderDealHeader();
        populateDealForm();
      }
    });
  }

  loadDealsData();
  loadTasksForDeal();

  if (!id || !Array.isArray(allDeals)) {
    emptyState.style.display = "block";
    return;
  }

  currentDeal = allDeals.find((deal) => normalizeValue(deal.id) === normalizeValue(id));
  if (!currentDeal) {
    emptyState.style.display = "block";
    return;
  }

  shell.style.display = "block";

  renderDealHeader();
  populateDealForm();
  renderRelatedTasks();
  setupDealForm();
  setupDeckPicker();
  setupDealTaskForm();
  setupFormToggles();
  setupGroupButtons();
  setupStageProgressionButton();
  syncContactStatusTasksFromDashboard();
}

document.addEventListener("DOMContentLoaded", loadDealPage);
