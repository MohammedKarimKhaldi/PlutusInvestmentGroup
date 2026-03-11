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

function normalizeValue(value) {
  if (AppCore) return AppCore.normalizeValue(value);
  return String(value || "").trim().toLowerCase();
}

function loadDealsData() {
  allDeals = AppCore ? AppCore.loadDealsData() : (Array.isArray(DEALS) ? JSON.parse(JSON.stringify(DEALS)) : []);
}

function saveDealsData() {
  if (AppCore) {
    AppCore.saveDealsData(allDeals);
    return;
  }
  try {
    localStorage.setItem(DEALS_STORAGE_KEY, JSON.stringify(allDeals));
  } catch (e) {
    console.warn("Failed to save deals to storage", e);
  }
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
    if (!window.PlutusDesktop || typeof window.PlutusDesktop.getShareDriveDownloadUrl !== "function") {
      return null;
    }
    const result = await window.PlutusDesktop.getShareDriveDownloadUrl({ shareUrl: url });
    if (!result || !result.ok || !result.data || !result.data.downloadUrl) {
      throw new Error((result && result.error) || "Failed to resolve SharePoint link.");
    }
    return result.data.downloadUrl;
  };

  const fetchWorkbookFromUrl = async (fetchUrl, expectsJsonWrapper) => {
    const response = await fetch(fetchUrl);
    if (!response.ok) throw new Error(`Download failed`);

    let buffer;
    if (expectsJsonWrapper) {
      const json = await response.json();
      if (!json.contents) throw new Error("No contents in allorigins response");
      const b64 = json.contents.split(",")[1] || json.contents;
      const binaryString = atob(b64);
      buffer = new ArrayBuffer(binaryString.length);
      const view = new Uint8Array(buffer);
      for (let j = 0; j < binaryString.length; j++) view[j] = binaryString.charCodeAt(j);
    } else {
      buffer = await response.arrayBuffer();
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
  if (!window.DASHBOARD_CONFIG || !Array.isArray(window.DASHBOARD_CONFIG.dashboards)) return;

  const dashboardId = normalizeValue(currentDeal.fundraisingDashboardId);
  if (!dashboardId) return;

  const dashboard = window.DASHBOARD_CONFIG.dashboards.find(
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

  const btnDashboard = document.getElementById("btn-open-dashboard");
  btnDashboard.onclick = () => {
    const dashId = currentDeal.fundraisingDashboardId || "biolux";
    window.location.href = `investor-dashboard.html?dashboard=${encodeURIComponent(dashId)}`;
  };

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
  document.getElementById("deal-input-cash").value = currentDeal.CashCommission || "";
  document.getElementById("deal-input-equity").value = currentDeal.EquityCommission || "";
  document.getElementById("deal-input-retainer").value = currentDeal.Retainer || "";
  document.getElementById("deal-input-summary").value = currentDeal.summary || "";
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
    const ownerLink = `owner-tasks.html?owner=${encodeURIComponent(owner)}`;

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

  form.addEventListener("submit", (event) => {
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

    saveDealsData();
    renderDealHeader();

    statusEl.textContent = "Saved";
    window.setTimeout(() => {
      statusEl.textContent = "";
    }, 1500);
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

function loadDealPage() {
  const params = new URLSearchParams(window.location.search);
  const id = params.get("id");

  const emptyState = document.getElementById("empty-state");
  const shell = document.getElementById("shell");

  loadDealsData();
  loadTasksForDeal();
  if (AppCore) {
    window.addEventListener("appcore:tasks-updated", (event) => {
      if (event && event.detail && Array.isArray(event.detail.tasks)) {
        allTasks = event.detail.tasks;
        if (currentDeal) renderRelatedTasks();
      }
    });
  }

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
  setupDealTaskForm();
  setupFormToggles();
  setupGroupButtons();
  setupStageProgressionButton();
  syncContactStatusTasksFromDashboard();
}

document.addEventListener("DOMContentLoaded", loadDealPage);
