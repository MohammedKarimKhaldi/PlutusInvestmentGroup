const AppCore = window.AppCore;
const DEALS_STORAGE_KEY = (AppCore && AppCore.STORAGE_KEYS && AppCore.STORAGE_KEYS.deals) || "deals_data_v1";
const TASKS_STORAGE_KEY = (AppCore && AppCore.STORAGE_KEYS && AppCore.STORAGE_KEYS.tasks) || "owner_tasks_v1";

let allDeals = [];
let allTasks = [];
let currentDeal = null;
let groupMode = "owner";
let accountingAccessState = { restricted: false, allowed: true };
const AUTO_CONTACT_TASK_PREFIX = (AppCore && AppCore.AUTO_CONTACT_TASK_PREFIX) || "auto-contact-status";

const STAGE_LABELS = {
  prospect: "Prospect",
  signing: "Signing",
  onboarding: "Onboarding",
  "contacting investors": "Contacting investors",
};
const STAGE_ORDER = ["prospect", "signing", "onboarding", "contacting investors"];
const DEAL_LIFECYCLE_LABELS = {
  active: "Active",
  finished: "Finished",
  closed: "Closed - not concluded",
};
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

function getDealRetainerState(deal) {
  if (AppCore && typeof AppCore.getDealRetainerState === "function") {
    return AppCore.getDealRetainerState(deal);
  }
  const rawValue = String(deal && (deal.retainerMonthly != null ? deal.retainerMonthly : deal && deal.Retainer) || "").trim();
  const cleaned = rawValue.replace(/,/g, "").replace(/[^0-9.-]/g, "");
  const amount = Number(cleaned);
  const hasRetainer = Number.isFinite(amount) && amount > 0;
  return {
    rawValue,
    amount: hasRetainer ? amount : 0,
    hasRetainer,
    bucket: hasRetainer ? "with-retainer" : "no-retainer",
    label: hasRetainer ? "With retainer" : "0 / no retainer",
  };
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

function stageClass(stage) {
  const s = normalizeValue(stage);
  if (s === "prospect") return "stage-prospect";
  if (s === "signing") return "stage-signing";
  if (s === "onboarding") return "stage-onboarding";
  if (s === "contacting investors") return "stage-contacting";
  return "stage-prospect";
}

function normalizeDealLifecycleStatus(value) {
  const normalized = normalizeValue(value);
  if (normalized === "finished") return "finished";
  if (normalized === "closed") return "closed";
  return "active";
}

function getCurrentDealLifecycleStatus() {
  return normalizeDealLifecycleStatus(currentDeal && (currentDeal.lifecycleStatus || currentDeal.dealStatus));
}

function isCurrentDealClosedLifecycle() {
  return getCurrentDealLifecycleStatus() !== "active";
}

async function updateDealLifecycleStatus(nextStatus) {
  if (!currentDeal) return;
  currentDeal.lifecycleStatus = normalizeDealLifecycleStatus(nextStatus);
  currentDeal.dealStatus = currentDeal.lifecycleStatus;
  await saveDealsData();
  renderDealHeader();
  populateDealForm();
  renderRelatedTasks();
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
  if (isCurrentDealClosedLifecycle()) return;
  currentDeal.stage = nextStage;
  saveDealsData();
  renderDealHeader();
  populateDealForm();
  renderRelatedTasks();
}

function refreshStageButton() {
  const button = document.getElementById("btn-complete-stage");
  if (!button || !currentDeal) return;

  if (isCurrentDealClosedLifecycle()) {
    button.disabled = true;
    button.textContent = getCurrentDealLifecycleStatus() === "finished" ? "Deal finished" : "Deal closed";
    return;
  }

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

function refreshDealLifecycleControls() {
  const finishedBtn = document.getElementById("btn-mark-deal-finished");
  const closedBtn = document.getElementById("btn-mark-deal-closed");
  const reopenBtn = document.getElementById("btn-reopen-deal");
  if (!finishedBtn || !closedBtn || !reopenBtn || !currentDeal) return;

  const lifecycleStatus = getCurrentDealLifecycleStatus();
  finishedBtn.hidden = lifecycleStatus === "finished";
  closedBtn.hidden = lifecycleStatus === "closed";
  reopenBtn.hidden = lifecycleStatus === "active";

  finishedBtn.disabled = lifecycleStatus === "finished";
  closedBtn.disabled = lifecycleStatus === "closed";
  reopenBtn.disabled = lifecycleStatus === "active";
}

function setupDealLifecycleButtons() {
  const finishedBtn = document.getElementById("btn-mark-deal-finished");
  const closedBtn = document.getElementById("btn-mark-deal-closed");
  const reopenBtn = document.getElementById("btn-reopen-deal");

  if (finishedBtn) {
    finishedBtn.addEventListener("click", async () => {
      if (!currentDeal) return;
      const confirmed = window.confirm(`Move "${currentDeal.name || "this deal"}" to finished deals?`);
      if (!confirmed) return;
      try {
        await updateDealLifecycleStatus("finished");
      } catch (error) {
        window.alert(error instanceof Error ? error.message : "Failed to update deal status.");
      }
    });
  }

  if (closedBtn) {
    closedBtn.addEventListener("click", async () => {
      if (!currentDeal) return;
      const confirmed = window.confirm(`Close "${currentDeal.name || "this deal"}" as not concluded?`);
      if (!confirmed) return;
      try {
        await updateDealLifecycleStatus("closed");
      } catch (error) {
        window.alert(error instanceof Error ? error.message : "Failed to update deal status.");
      }
    });
  }

  if (reopenBtn) {
    reopenBtn.addEventListener("click", async () => {
      if (!currentDeal) return;
      const confirmed = window.confirm(`Reopen "${currentDeal.name || "this deal"}" and move it back into active deals?`);
      if (!confirmed) return;
      try {
        await updateDealLifecycleStatus("active");
      } catch (error) {
        window.alert(error instanceof Error ? error.message : "Failed to update deal status.");
      }
    });
  }
}

function setupStageProgressionButton() {
  const button = document.getElementById("btn-complete-stage");
  if (!button) return;

  button.addEventListener("click", () => {
    if (!currentDeal) return;
    if (isCurrentDealClosedLifecycle()) return;
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
  if (isCurrentDealClosedLifecycle()) return false;
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

function normalizeDealSubOwners(deal) {
  const source = deal && deal.subOwners;
  const values = Array.isArray(source)
    ? source
    : typeof source === "string"
      ? source.split(/[\n,;]+/)
      : [];
  const normalized = values
    .map((value) => String(value || "").trim())
    .filter(Boolean);
  const deduped = Array.from(new Set(normalized));
  if (deal && typeof deal === "object") {
    deal.subOwners = deduped;
  }
  return deduped;
}

function parseDealSubOwnersInput(value) {
  return Array.from(
    new Set(
      String(value || "")
        .split(/[\n,;]+/)
        .map((entry) => String(entry || "").trim())
        .filter(Boolean),
    ),
  );
}

function normalizeDealKeywords(deal) {
  const source = deal && deal.keywords;
  const values = Array.isArray(source)
    ? source
    : typeof source === "string"
      ? source.split(/[\n,;]+/)
      : [];
  const normalized = values
    .map((value) => String(value || "").trim())
    .filter(Boolean);
  const deduped = Array.from(new Set(normalized));
  if (deal && typeof deal === "object") {
    deal.keywords = deduped;
  }
  return deduped;
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
        .map((value) => String(value || "").trim())
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
        .filter(Boolean),
    ),
  );
}

function parseDealKeywordsInput(value) {
  return Array.from(
    new Set(
      String(value || "")
        .split(/[\n,;]+/)
        .map((entry) => String(entry || "").trim())
        .filter(Boolean),
    ),
  );
}

function buildDealKeywordsHtml(keywords) {
  const values = Array.isArray(keywords) ? keywords : [];
  if (!values.length) {
    return '<span class="deal-keyword-chip is-empty">No thematics, sector tags, or keywords added yet.</span>';
  }
  return values
    .map((keyword) => `<span class="deal-keyword-chip">${escapeHtml(keyword)}</span>`)
    .join("");
}

function getAssignableOwnersForDeal(deal) {
  const seen = new Set();
  return [
    getSeniorOwner(deal),
    getJuniorOwner(deal),
    ...normalizeDealSubOwners(deal),
  ].filter((entry) => {
    const value = String(entry || "").trim();
    const key = normalizeValue(value);
    if (!key || seen.has(key)) return false;
    seen.add(key);
    return true;
  });
}

function renderOwnerOptionsIntoDatalist(listEl, values) {
  if (!listEl) return;
  listEl.innerHTML = "";
  values.forEach((value) => {
    const option = document.createElement("option");
    option.value = value;
    listEl.appendChild(option);
  });
}

function refreshDealTaskOwnerSuggestions() {
  const ownerInput = document.getElementById("task-input-owner");
  const datalist = document.getElementById("deal-task-owner-options");
  const helpEl = document.getElementById("task-owner-help");
  if (!ownerInput || !datalist) return;

  const owners = getAssignableOwnersForDeal(currentDeal);
  renderOwnerOptionsIntoDatalist(datalist, owners);

  if (helpEl) {
    helpEl.textContent = owners.length
      ? `Suggested assignees for this deal: ${owners.join(", ")}. These names also flow through to the task views.`
      : "Add senior, junior, or sub owners in the deal editor to suggest them here.";
  }

  if (!String(ownerInput.value || "").trim() && owners.length) {
    ownerInput.value = owners[0];
  }
}

function buildOwnerTaskLinksHtml(owners) {
  return owners
    .map((owner) => {
      const label = escapeHtml(owner);
      const href = buildPageUrl("owner-tasks", { owner });
      return `<a class="meta-owner-link" href="${href}">${label}</a>`;
    })
    .join("");
}

function normalizeDealContacts(deal, options = {}) {
  const { keepEmpty = false } = options;
  const source = deal && Array.isArray(deal.contacts) ? deal.contacts : [];
  const normalized = source
    .map((entry) => {
      if (!entry || typeof entry !== "object") return null;
      const name = String(entry.name || "").trim();
      const title = String(entry.title || "").trim();
      const email = String(entry.email || "").trim();
      const isPrimary = Boolean(entry.isPrimary);
      if (!keepEmpty && !name && !title && !email) return null;
      return { name, title, email, isPrimary };
    })
    .filter(Boolean);

  if (!normalized.length) {
    const fallbackEmail = String((deal && (deal.mainContactEmail || deal.email)) || "").trim();
    const fallbackName = String((deal && deal.mainContactName) || "").trim();
    const fallbackTitle = String((deal && deal.mainContactTitle) || "").trim();
    if (fallbackEmail || fallbackName || fallbackTitle) {
      normalized.push({
        name: fallbackName,
        title: fallbackTitle,
        email: fallbackEmail,
        isPrimary: true,
      });
    }
  }

  if (normalized.length && !normalized.some((entry) => entry.isPrimary)) {
    normalized[0].isPrimary = true;
  }

  if (deal && typeof deal === "object") {
    deal.contacts = normalized;
  }
  return normalized;
}

function getPrimaryDealContact(deal) {
  const contacts = normalizeDealContacts(deal);
  return contacts.find((entry) => entry.isPrimary) || contacts[0] || null;
}

function syncLegacyPrimaryContactFields(deal) {
  if (!deal || typeof deal !== "object") return;
  const primary = getPrimaryDealContact(deal);
  deal.mainContactName = primary ? primary.name : "";
  deal.mainContactTitle = primary ? primary.title : "";
  deal.mainContactEmail = primary ? primary.email : "";
  deal.email = primary ? primary.email : "";
}

function normalizeDealLegalLinks(deal, options = {}) {
  const { keepEmpty = false } = options;
  const source =
    deal && Array.isArray(deal.legalLinks)
      ? deal.legalLinks
      : deal && Array.isArray(deal.legalAspects)
        ? deal.legalAspects
        : [];

  const normalized = source
    .map((entry, index) => {
      if (typeof entry === "string") {
        const url = String(entry || "").trim();
        if (!keepEmpty && !url) return null;
        return {
          title: url ? `Legal link ${index + 1}` : "",
          url,
        };
      }
      if (!entry || typeof entry !== "object") return null;
      const title = String(entry.title || entry.label || entry.name || "").trim();
      const url = String(entry.url || entry.href || entry.link || "").trim();
      if (!keepEmpty && !title && !url) return null;
      return {
        title: title || (url ? `Legal link ${index + 1}` : ""),
        url,
      };
    })
    .filter(Boolean);

  if (deal && typeof deal === "object") {
    deal.legalLinks = normalized;
  }
  return normalized;
}

function toSafeExternalUrl(value) {
  const raw = String(value || "").trim();
  if (!raw) return "";
  try {
    const parsed = new URL(raw);
    const protocol = String(parsed.protocol || "").toLowerCase();
    if (protocol !== "http:" && protocol !== "https:") return "";
    return parsed.toString();
  } catch {
    return "";
  }
}

function buildLegalLinkEditorRow(link, index) {
  return `
    <div class="legal-link-editor-row" data-legal-link-row="${index}">
      <div class="contact-editor-grid">
        <label>Title<input type="text" data-legal-link-field="title" value="${escapeHtml(link.title || "")}" placeholder="Term sheet, NDA, counsel notes..." /></label>
        <label>URL<input type="url" data-legal-link-field="url" value="${escapeHtml(link.url || "")}" placeholder="https://..." /></label>
      </div>
      <div class="legal-link-editor-actions">
        <div class="linked-asset-meta">Use a full SharePoint, Google Drive, DocuSign, or other legal document link.</div>
        <button class="btn" type="button" data-remove-legal-link-row="${index}">Remove</button>
      </div>
    </div>
  `;
}

function renderDealLegalLinkEditor() {
  const listEl = document.getElementById("deal-legal-editor-list");
  if (!listEl || !currentDeal) return;
  const links = normalizeDealLegalLinks(currentDeal, { keepEmpty: true });
  listEl.innerHTML = links.length
    ? links.map((link, index) => buildLegalLinkEditorRow(link, index)).join("")
    : '<div class="legal-link-empty">No legal links yet. Add one to keep deal documents one click away.</div>';
}

function collectDealLegalLinksFromEditor() {
  const listEl = document.getElementById("deal-legal-editor-list");
  if (!listEl) return [];
  const rows = Array.from(listEl.querySelectorAll("[data-legal-link-row]"));
  return rows
    .map((row, index) => {
      const title = String((row.querySelector('[data-legal-link-field="title"]') || {}).value || "").trim();
      const url = String((row.querySelector('[data-legal-link-field="url"]') || {}).value || "").trim();
      if (!title && !url) return null;
      return {
        title: title || `Legal link ${index + 1}`,
        url,
      };
    })
    .filter(Boolean);
}

function renderDealLegalLinks() {
  const listEl = document.getElementById("deal-legal-links-list");
  const summaryEl = document.getElementById("deal-legal-links-summary");
  if (!listEl || !summaryEl || !currentDeal) return;

  const legalLinks = normalizeDealLegalLinks(currentDeal);
  if (!legalLinks.length) {
    summaryEl.textContent = "Add legal document and deal counsel links in the editor below.";
    listEl.innerHTML = '<div class="legal-link-empty">No legal links saved for this deal yet.</div>';
    return;
  }

  summaryEl.textContent = `${legalLinks.length} legal link${legalLinks.length === 1 ? "" : "s"} linked to this deal.`;
  listEl.innerHTML = legalLinks
    .map((link) => {
      const safeUrl = toSafeExternalUrl(link.url);
      if (!safeUrl) {
        return `
          <div class="legal-link-card is-invalid">
            <div class="legal-link-title">${escapeHtml(link.title || "Legal link")}</div>
            <div class="legal-link-meta">This entry does not have a valid http(s) URL yet.</div>
          </div>
        `;
      }

      return `
        <div class="legal-link-card">
          <div class="legal-link-title">${escapeHtml(link.title || "Legal link")}</div>
          <a class="btn legal-link-button" href="${escapeHtml(safeUrl)}" target="_blank" rel="noopener noreferrer">
            Open ${escapeHtml(link.title || "legal document")}
          </a>
        </div>
      `;
    })
    .join("");
}

function renderDealContactsSummary() {
  const listEl = document.getElementById("deal-contacts-list");
  const summaryEl = document.getElementById("deal-contacts-summary");
  if (!listEl || !summaryEl || !currentDeal) return;

  const contacts = normalizeDealContacts(currentDeal);
  if (!contacts.length) {
    summaryEl.textContent = "No company contacts saved for this deal yet.";
    listEl.innerHTML = '<div class="contact-card"><div class="contact-title">Add contacts in the deal editor to keep emails and titles on the deal.</div></div>';
    return;
  }

  const primary = getPrimaryDealContact(currentDeal);
  summaryEl.textContent = primary && primary.email
    ? `Main contact: ${primary.name || primary.email}${primary.title ? ` · ${primary.title}` : ""}`
    : `${contacts.length} contact${contacts.length === 1 ? "" : "s"} saved for this deal.`;

  listEl.innerHTML = contacts.map((contact) => `
    <div class="contact-card">
      <div class="contact-card-top">
        <div>
          <div class="contact-name">${escapeHtml(contact.name || contact.email || "Unnamed contact")}</div>
          <div class="contact-title">${escapeHtml(contact.title || "No title")}</div>
        </div>
        ${contact.isPrimary ? '<span class="contact-primary-badge">Main contact</span>' : ""}
      </div>
      ${contact.email ? `<a class="contact-email" href="mailto:${escapeHtml(contact.email)}">${escapeHtml(contact.email)}</a>` : '<div class="contact-title">No email</div>'}
    </div>
  `).join("");
}

function buildContactEditorRow(contact, index) {
  return `
    <div class="contact-editor-row" data-contact-row="${index}">
      <div class="contact-editor-grid">
        <label>Name<input type="text" data-contact-field="name" value="${escapeHtml(contact.name || "")}" /></label>
        <label>Title<input type="text" data-contact-field="title" value="${escapeHtml(contact.title || "")}" /></label>
        <label>Email<input type="email" data-contact-field="email" value="${escapeHtml(contact.email || "")}" /></label>
      </div>
      <div class="contact-editor-actions">
        <label class="contact-primary-toggle">
          <input type="radio" name="deal-contact-primary" data-contact-field="isPrimary" ${contact.isPrimary ? "checked" : ""} />
          <span>Main contact</span>
        </label>
        <button class="btn" type="button" data-remove-contact-row="${index}">Remove</button>
      </div>
    </div>
  `;
}

function renderDealContactEditor() {
  const listEl = document.getElementById("deal-contact-editor-list");
  if (!listEl || !currentDeal) return;
  const contacts = normalizeDealContacts(currentDeal, { keepEmpty: true });
  listEl.innerHTML = contacts.length
    ? contacts.map((contact, index) => buildContactEditorRow(contact, index)).join("")
    : '<div class="contact-editor-row"><div class="contact-title">No contacts yet. Add one to start tracking the main email contact.</div></div>';
}

function collectDealContactsFromEditor() {
  const listEl = document.getElementById("deal-contact-editor-list");
  if (!listEl) return [];
  const rows = Array.from(listEl.querySelectorAll("[data-contact-row]"));
  const contacts = rows.map((row) => ({
    name: String((row.querySelector('[data-contact-field="name"]') || {}).value || "").trim(),
    title: String((row.querySelector('[data-contact-field="title"]') || {}).value || "").trim(),
    email: String((row.querySelector('[data-contact-field="email"]') || {}).value || "").trim(),
    isPrimary: Boolean((row.querySelector('[data-contact-field="isPrimary"]') || {}).checked),
  })).filter((entry) => entry.name || entry.title || entry.email);

  if (contacts.length && !contacts.some((entry) => entry.isPrimary)) {
    contacts[0].isPrimary = true;
  }
  return contacts;
}

function normalizeInvestorContactStatus(value) {
  const normalized = normalizeValue(value);
  if (normalized === "queued") return "queued";
  if (normalized === "contacted") return "contacted";
  if (normalized === "follow-up") return "follow-up";
  if (normalized === "replied") return "replied";
  if (normalized === "invested") return "invested";
  if (normalized === "passed") return "passed";
  return "not-contacted";
}

function getInvestorContactStatusLabel(value) {
  const normalized = normalizeInvestorContactStatus(value);
  if (normalized === "queued") return "Queued";
  if (normalized === "contacted") return "Contacted";
  if (normalized === "follow-up") return "Follow up";
  if (normalized === "replied") return "Replied";
  if (normalized === "invested") return "Invested";
  if (normalized === "passed") return "Passed";
  return "Not contacted";
}

function normalizeDealInvestorContacts(deal) {
  const source = deal && Array.isArray(deal.investorContacts) ? deal.investorContacts : [];
  const normalized = source
    .map((entry, index) => {
      if (!entry || typeof entry !== "object") return null;
      const email = String(entry.email || "").trim();
      if (!email) return null;
      return {
        id: String(entry.id || `investor-${index}`).trim(),
        name: String(entry.name || "").trim(),
        firstName: String(entry.firstName || "").trim(),
        lastName: String(entry.lastName || "").trim(),
        email,
        contactStatus: normalizeInvestorContactStatus(entry.contactStatus),
        sourceMessageSubject: String(entry.sourceMessageSubject || "").trim(),
        sourceMessageWebLink: String(entry.sourceMessageWebLink || "").trim(),
        sourceReceivedAt: String(entry.sourceReceivedAt || "").trim(),
        sourceFromName: String(entry.sourceFromName || "").trim(),
        sourceFromEmail: String(entry.sourceFromEmail || "").trim(),
        notes: String(entry.notes || "").trim(),
        addedAt: String(entry.addedAt || "").trim(),
        updatedAt: String(entry.updatedAt || "").trim(),
      };
    })
    .filter(Boolean)
    .sort((left, right) => {
      const leftValue = Date.parse(left.updatedAt || left.addedAt || "") || 0;
      const rightValue = Date.parse(right.updatedAt || right.addedAt || "") || 0;
      if (leftValue !== rightValue) return rightValue - leftValue;
      return String(left.name || left.email).localeCompare(String(right.name || right.email));
    });

  if (deal && typeof deal === "object") {
    deal.investorContacts = normalized;
  }
  return normalized;
}

function renderDealInvestorContacts() {
  const listEl = document.getElementById("deal-investor-contacts-list");
  const summaryEl = document.getElementById("deal-investor-contacts-summary");
  if (!listEl || !summaryEl || !currentDeal) return;

  const investors = normalizeDealInvestorContacts(currentDeal);
  if (!investors.length) {
    summaryEl.textContent = "Track investors added from Outlook or manual entries for this deal.";
    listEl.innerHTML = '<div class="contact-card"><div class="contact-title">No investor contacts have been saved yet. Use the Outlook investors page to pull recipients onto this deal.</div></div>';
    return;
  }

  summaryEl.textContent = `${investors.length} investor contact${investors.length === 1 ? "" : "s"} tracked for this deal.`;
  listEl.innerHTML = investors.map((investor) => `
    <div class="contact-card investor-contact-card">
      <div class="contact-card-top">
        <div>
          <div class="contact-name">${escapeHtml(investor.name || investor.email)}</div>
          <div class="contact-title">
            ${escapeHtml(getInvestorContactStatusLabel(investor.contactStatus))}
            ${investor.sourceReceivedAt ? ` · ${escapeHtml(formatDate(investor.sourceReceivedAt))}` : ""}
          </div>
        </div>
        <span class="contact-primary-badge investor-status-badge investor-status-${escapeHtml(investor.contactStatus)}">
          ${escapeHtml(getInvestorContactStatusLabel(investor.contactStatus))}
        </span>
      </div>
      <a class="contact-email" href="mailto:${escapeHtml(investor.email)}">${escapeHtml(investor.email)}</a>
      ${investor.sourceMessageSubject
        ? `<div class="contact-title">Source email: ${escapeHtml(investor.sourceMessageSubject)}</div>`
        : '<div class="contact-title">Source email: Manual entry</div>'}
      ${investor.sourceMessageWebLink
        ? `<a class="task-link" href="${escapeHtml(investor.sourceMessageWebLink)}" target="_blank" rel="noopener noreferrer">Open source email</a>`
        : ""}
      ${investor.notes ? `<div class="task-notes">${escapeHtml(investor.notes)}</div>` : ""}
    </div>
  `).join("");
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

function getOwnershipDashboardConfig() {
  const dashboardsConfig = getDashboardConfig();
  if (!dashboardsConfig || !Array.isArray(dashboardsConfig.dashboards)) return null;
  if (AppCore && typeof AppCore.getDashboardById === "function") {
    return AppCore.getDashboardById(dashboardsConfig, "deal-ownership");
  }
  return dashboardsConfig.dashboards.find((entry) => normalizeValue(entry && entry.id) === "deal-ownership") || null;
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

function arrayBufferToBase64(payload) {
  let bytes = null;
  if (payload instanceof ArrayBuffer) {
    bytes = new Uint8Array(payload);
  } else if (ArrayBuffer.isView(payload)) {
    bytes = new Uint8Array(payload.buffer, payload.byteOffset, payload.byteLength);
  } else {
    return "";
  }

  let binary = "";
  const chunkSize = 0x8000;
  for (let index = 0; index < bytes.length; index += chunkSize) {
    const slice = bytes.subarray(index, Math.min(index + chunkSize, bytes.length));
    binary += String.fromCharCode.apply(null, slice);
  }
  return btoa(binary);
}

function looksLikeHtmlBuffer(buffer) {
  try {
    const bytes = new Uint8Array(buffer || new ArrayBuffer(0)).slice(0, 512);
    const text = new TextDecoder("utf-8").decode(bytes).trim().toLowerCase();
    return text.startsWith("<!doctype html") || text.startsWith("<html") || text.includes("<head") || text.includes("<body");
  } catch {
    return false;
  }
}

function getOwnershipWorkbookSheet(workbook, sheetsConfig) {
  if (!workbook || !Array.isArray(workbook.SheetNames) || !workbook.SheetNames.length) {
    return { sheetName: "", sheet: null };
  }

  const preferred = String(sheetsConfig && sheetsConfig.funds || "").trim();
  let sheetName = "";
  if (preferred) {
    sheetName = workbook.SheetNames.find((entry) => {
      const normalizedEntry = normalizeValue(entry);
      const normalizedPreferred = normalizeValue(preferred);
      return normalizedEntry === normalizedPreferred || normalizedEntry.includes(normalizedPreferred);
    }) || "";
  }
  if (!sheetName) {
    sheetName = workbook.SheetNames[0] || "";
  }
  return {
    sheetName,
    sheet: sheetName ? workbook.Sheets[sheetName] : null,
  };
}

function detectOwnershipHeaderRowIndex(rows, maxScanRows = 12) {
  const upperBound = Math.min(Array.isArray(rows) ? rows.length : 0, maxScanRows);
  for (let index = 0; index < upperBound; index += 1) {
    const headers = Array.isArray(rows[index]) ? rows[index].map((cell) => String(cell || "").trim()) : [];
    const hasDealHeader = headers.some((header) => /deal|company|project/i.test(header));
    const hasStaffHeader = headers.some((header) => /^staff$/i.test(header) || /^name$/i.test(header) || /^lead$/i.test(header) || /staff|owner/i.test(header));
    if (hasDealHeader && hasStaffHeader) {
      return index;
    }
  }
  return 0;
}

function findOwnershipHeaderIndex(headers, patterns) {
  return headers.findIndex((header) => patterns.some((pattern) => pattern.test(String(header || "").trim())));
}

function ensureRowLength(row, length) {
  const nextRow = Array.isArray(row) ? row.slice() : [];
  while (nextRow.length < length) nextRow.push("");
  return nextRow;
}

function buildCurrentDealReferenceSet(extraReferences = []) {
  if (!currentDeal) return new Set();
  return new Set(
    [
      currentDeal.id,
      currentDeal.name,
      currentDeal.company,
      currentDeal.fundraisingDashboardId,
      currentDeal.company && currentDeal.name ? `${currentDeal.company} ${currentDeal.name}` : "",
      currentDeal.name && currentDeal.company ? `${currentDeal.name} ${currentDeal.company}` : "",
      ...extraReferences,
    ]
      .map((value) => normalizeValue(value))
      .filter(Boolean),
  );
}

function rowMatchesCurrentDealInOwnershipSheet(row, dealColumnIndexes, dealReferenceSet) {
  const references = dealColumnIndexes
    .map((index) => row[index])
    .map((value) => String(value || "").trim())
    .filter(Boolean);
  if (!references.length) return false;

  if (AppCore && typeof AppCore.findDealByReference === "function") {
    const matchedDeal = AppCore.findDealByReference(allDeals, references);
    if (matchedDeal && normalizeValue(matchedDeal.id) === normalizeValue(currentDeal && currentDeal.id)) {
      return true;
    }
  }

  return references.some((value) => dealReferenceSet.has(normalizeValue(value)));
}

function isOwnershipSubOwnerRow(row, roleColumnIndex) {
  if (!Array.isArray(row) || roleColumnIndex < 0) return false;
  return normalizeValue(row[roleColumnIndex]).startsWith("sub owner");
}

async function syncDealSubOwnersToOwnershipWorkbook(options = {}) {
  if (!currentDeal) return { attempted: false, skipped: true };
  if (!window.XLSX) {
    throw new Error("Workbook sync is unavailable because the Excel parser did not load.");
  }
  if (
    !AppCore ||
    typeof AppCore.resolveShareDriveFile !== "function" ||
    typeof AppCore.downloadBinary !== "function" ||
    typeof AppCore.uploadShareDriveFile !== "function"
  ) {
    throw new Error("Workbook sync is unavailable in this mode.");
  }

  const ownershipDashboard = getOwnershipDashboardConfig();
  if (!ownershipDashboard || !ownershipDashboard.excelUrl) {
    return { attempted: false, skipped: true };
  }

  const resolvedFile = await AppCore.resolveShareDriveFile({
    shareUrl: ownershipDashboard.excelUrl,
  });
  if (!resolvedFile || !resolvedFile.downloadUrl) {
    throw new Error("Could not resolve the deal ownership workbook download URL.");
  }
  if (!resolvedFile.name || !resolvedFile.parentItemId) {
    throw new Error("Could not resolve the deal ownership workbook location for upload.");
  }

  const buffer = await AppCore.downloadBinary(resolvedFile.downloadUrl, { cache: "no-store" });
  if (looksLikeHtmlBuffer(buffer)) {
    throw new Error("Deal ownership workbook download returned an HTML page. Reconnect Sharedrive and try again.");
  }

  const workbook = XLSX.read(buffer, { type: "array" });
  const { sheetName, sheet } = getOwnershipWorkbookSheet(workbook, ownershipDashboard.sheets || {});
  if (!sheet || !sheetName) {
    throw new Error("Deal ownership sheet not found in the linked workbook.");
  }

  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });
  if (!rows.length) {
    throw new Error("Deal ownership sheet is empty.");
  }

  const headerRowIndex = detectOwnershipHeaderRowIndex(rows);
  const headers = ensureRowLength(rows[headerRowIndex], rows[headerRowIndex] ? rows[headerRowIndex].length : 0)
    .map((header) => String(header || "").trim());
  const dealColumnIndexes = headers
    .map((header, index) => (/deal|company|project/i.test(header) ? index : -1))
    .filter((index) => index >= 0);
  const primaryDealColumnIndex = findOwnershipHeaderIndex(headers, [/deal/i, /company/i, /project/i]);
  const dealNameColumnIndex = findOwnershipHeaderIndex(headers, [/deal/i, /project/i]);
  const companyColumnIndex = findOwnershipHeaderIndex(headers, [/company/i]);
  const staffColumnIndex = findOwnershipHeaderIndex(headers, [/^staff$/i, /^name$/i, /^lead$/i, /staff/i, /owner/i]);
  let roleColumnIndex = findOwnershipHeaderIndex(headers, [/^role$/i, /title/i, /position/i, /function/i, /team/i]);

  if (primaryDealColumnIndex < 0 || staffColumnIndex < 0) {
    throw new Error("Deal ownership sheet needs both a deal column and a staff/owner column.");
  }

  if (roleColumnIndex < 0) {
    roleColumnIndex = headers.length;
    headers.push("Role");
  }

  const rowLength = headers.length;
  const prefixRows = rows
    .slice(0, headerRowIndex + 1)
    .map((row, index) => (index === headerRowIndex ? headers.slice() : ensureRowLength(row, rowLength)));
  const dataRows = rows.slice(headerRowIndex + 1).map((row) => ensureRowLength(row, rowLength));
  const dealReferenceSet = buildCurrentDealReferenceSet([
    options.previousName,
    options.previousCompany,
    options.previousName && currentDeal && currentDeal.company ? `${options.previousName} ${currentDeal.company}` : "",
    options.previousCompany && currentDeal && currentDeal.name ? `${currentDeal.name} ${options.previousCompany}` : "",
    options.previousCompany && options.previousName ? `${options.previousCompany} ${options.previousName}` : "",
    options.previousName && options.previousCompany ? `${options.previousName} ${options.previousCompany}` : "",
  ]);
  const filteredDataRows = dataRows.filter((row) => !(
    rowMatchesCurrentDealInOwnershipSheet(row, dealColumnIndexes, dealReferenceSet) &&
    isOwnershipSubOwnerRow(row, roleColumnIndex)
  ));
  const removedCount = dataRows.length - filteredDataRows.length;

  const subOwners = normalizeDealSubOwners(currentDeal);
  const newRows = subOwners.map((subOwner) => {
    const row = new Array(rowLength).fill("");
    const dealLabel = currentDeal.name || currentDeal.company || currentDeal.id || "";
    if (dealNameColumnIndex >= 0) row[dealNameColumnIndex] = dealLabel;
    if (companyColumnIndex >= 0) row[companyColumnIndex] = currentDeal.company || dealLabel;
    if (dealNameColumnIndex < 0 && primaryDealColumnIndex >= 0) row[primaryDealColumnIndex] = dealLabel;
    row[staffColumnIndex] = subOwner;
    row[roleColumnIndex] = "Sub owner";
    return row;
  });

  if (!removedCount && !newRows.length) {
    return { attempted: false, skipped: true };
  }

  const nextRows = prefixRows.concat(filteredDataRows, newRows);
  const nextSheet = XLSX.utils.aoa_to_sheet(nextRows);
  if (sheet["!cols"]) nextSheet["!cols"] = sheet["!cols"];
  if (sheet["!merges"]) nextSheet["!merges"] = sheet["!merges"];
  workbook.Sheets[sheetName] = nextSheet;

  const updatedWorkbook = XLSX.write(workbook, { type: "array", bookType: "xlsx" });
  await AppCore.uploadShareDriveFile({
    shareUrl: ownershipDashboard.excelUrl,
    parentItemId: resolvedFile.parentItemId,
    fileName: resolvedFile.name,
    contentBase64: arrayBufferToBase64(updatedWorkbook),
    conflictBehavior: "replace",
  });

  return {
    attempted: true,
    sheetName,
    syncedCount: newRows.length,
    removedCount,
  };
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

  const lifecycleStatus = getCurrentDealLifecycleStatus();
  const lifecycleLabel = DEAL_LIFECYCLE_LABELS[lifecycleStatus] || "Active";
  const stageText = STAGE_LABELS[normalizeValue(currentDeal.stage)] || "Prospect";

  document.getElementById("crumb-deal-name").textContent = currentDeal.name || "Deal";
  document.getElementById("deal-name").textContent = currentDeal.name || "Deal";
  document.getElementById("deal-subtitle").textContent =
    `${currentDeal.company || "Company"} • ${stageText}${lifecycleStatus !== "active" ? ` • ${lifecycleLabel}` : ""}`;

  document.getElementById("deal-company").textContent = currentDeal.company || "–";
  document.getElementById("deal-senior").textContent = getSeniorOwner(currentDeal) || "–";
  document.getElementById("deal-junior").textContent = getJuniorOwner(currentDeal) || "–";
  const subOwners = normalizeDealSubOwners(currentDeal);
  const sectorLabel = normalizeDealSectors(currentDeal).join(", ");
  document.getElementById("deal-sub-owners").innerHTML = subOwners.length
    ? `<div class="meta-owner-links">${buildOwnerTaskLinksHtml(subOwners)}</div>`
    : "–";
  document.getElementById("deal-target").textContent = formatAmount(currentDeal.targetAmount, currentDeal.currency);
  document.getElementById("deal-raised").textContent = formatAmount(currentDeal.raisedAmount, currentDeal.currency);

  const textOrDash = (value) => (value == null || String(value).trim() === "" ? "–" : String(value));
  document.getElementById("deal-sector").textContent = textOrDash(sectorLabel);
  document.getElementById("deal-location").textContent = textOrDash(currentDeal.location);
  document.getElementById("deal-funding-stage").textContent = textOrDash(currentDeal.fundingStage);
  document.getElementById("deal-revenue").textContent = textOrDash(currentDeal.revenue);
  const keywordsEl = document.getElementById("deal-keywords");
  if (keywordsEl) {
    keywordsEl.innerHTML = buildDealKeywordsHtml(normalizeDealKeywords(currentDeal));
  }
  document.getElementById("deal-cash-comm").textContent = textOrDash(currentDeal.CashCommission);
  document.getElementById("deal-equity-comm").textContent = textOrDash(currentDeal.EquityCommission);
  const retainerState = getDealRetainerState(currentDeal);
  document.getElementById("deal-retainer").textContent = retainerState.rawValue || "0";
  const retainerStateEl = document.getElementById("deal-retainer-state");
  if (retainerStateEl) {
    retainerStateEl.textContent = retainerState.label;
  }

  const progressPercent =
    currentDeal.targetAmount && !isNaN(currentDeal.targetAmount) && currentDeal.raisedAmount && !isNaN(currentDeal.raisedAmount)
      ? Math.min(100, Math.round((currentDeal.raisedAmount / currentDeal.targetAmount) * 100))
      : 0;

  document.getElementById("deal-progress").style.width = `${progressPercent}%`;

  const stage = normalizeValue(currentDeal.stage);
  const stageLabel = STAGE_LABELS[stage] || "Prospect";
  document.getElementById("deal-stage-label").textContent = stageLabel;
  document.getElementById("deal-stage-dot").className = `stage-dot ${stageClass(stage)}`;
  const lifecycleBadge = document.getElementById("deal-lifecycle-badge");
  if (lifecycleBadge) {
    lifecycleBadge.textContent = lifecycleLabel;
    lifecycleBadge.className = `deal-lifecycle-badge is-${lifecycleStatus}`;
  }
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
  const btnOutlookInvestors = document.getElementById("btn-open-outlook-investors");
  const btnOutlookInvestorsSecondary = document.getElementById("btn-open-outlook-investors-secondary");
  const btnEditDashboardConfig = document.getElementById("btn-edit-dashboard-config");
  const btnAccounting = document.getElementById("btn-open-accounting");
  const btnLegalWorkspace = document.getElementById("btn-open-legal-workspace");
  const dashboard = getDashboardForCurrentDeal();

  [btnOutlookInvestors, btnOutlookInvestorsSecondary].forEach((button) => {
    if (!button) return;
    if (currentDeal && currentDeal.id) {
      button.disabled = false;
      button.onclick = () => {
        window.location.href = buildPageUrl("outlook-investor-sync", { id: currentDeal.id });
      };
    } else {
      button.disabled = true;
      button.onclick = null;
    }
  });

  if (btnAccounting) {
    if (accountingAccessState.restricted && !accountingAccessState.allowed) {
      btnAccounting.disabled = true;
      btnAccounting.textContent = "Accounting restricted";
      btnAccounting.onclick = null;
    } else if (currentDeal && currentDeal.id) {
      btnAccounting.disabled = false;
      btnAccounting.textContent = "Open accounting";
      btnAccounting.onclick = () => {
        window.location.href = buildPageUrl("accounting", { id: currentDeal.id });
      };
    } else {
      btnAccounting.disabled = true;
      btnAccounting.onclick = null;
    }
  }

  if (btnLegalWorkspace) {
    if (currentDeal && currentDeal.id) {
      btnLegalWorkspace.disabled = false;
      btnLegalWorkspace.textContent = "Open legal workspace";
      btnLegalWorkspace.onclick = () => {
        window.location.href = buildPageUrl("legal-management", { id: currentDeal.id });
      };
    } else {
      btnLegalWorkspace.disabled = true;
      btnLegalWorkspace.onclick = null;
    }
  }

  if (dashboard && dashboard.id) {
    btnDashboard.disabled = false;
    btnDashboard.textContent = "Open fundraising dashboard";
    btnDashboard.onclick = () => {
      window.location.href = buildPageUrl("investor-dashboard", { dashboard: dashboard.id });
    };
    if (btnEditDashboardConfig) {
      btnEditDashboardConfig.disabled = false;
      btnEditDashboardConfig.textContent = "Edit dashboard config";
      btnEditDashboardConfig.onclick = () => {
        window.location.href = buildPageUrl("investor-dashboard", { dashboard: dashboard.id, edit: "1" });
      };
    }
  } else {
    btnDashboard.disabled = true;
    btnDashboard.textContent = "No dashboard linked";
    btnDashboard.onclick = null;
    if (btnEditDashboardConfig) {
      btnEditDashboardConfig.disabled = true;
      btnEditDashboardConfig.textContent = "No dashboard to edit";
      btnEditDashboardConfig.onclick = null;
    }
  }

  syncLegacyPrimaryContactFields(currentDeal);
  renderDealLegalLinks();
  renderDealContactsSummary();
  renderDealInvestorContacts();
  refreshDealTaskOwnerSuggestions();
  refreshStageButton();
  refreshDealLifecycleControls();
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
  document.getElementById("deal-input-sub-owners").value = normalizeDealSubOwners(currentDeal).join("\n");
  document.getElementById("deal-input-stage").value = normalizeValue(currentDeal.stage) || "prospect";
  document.getElementById("deal-input-lifecycle-status").value = getCurrentDealLifecycleStatus();
  document.getElementById("deal-input-sector").value = normalizeDealSectors(currentDeal).join("\n");
  document.getElementById("deal-input-location").value = currentDeal.location || "";
  document.getElementById("deal-input-funding-stage").value = currentDeal.fundingStage || "";
  document.getElementById("deal-input-revenue").value = currentDeal.revenue || "";
  document.getElementById("deal-input-target").value = currentDeal.targetAmount ?? "";
  document.getElementById("deal-input-raised").value = currentDeal.raisedAmount ?? "";
  document.getElementById("deal-input-currency").value = currentDeal.currency || "USD";
  document.getElementById("deal-input-dashboard").value = currentDeal.fundraisingDashboardId || "";
  refreshDashboardLinkEditor();
  document.getElementById("deal-input-cash").value = currentDeal.CashCommission || "";
  document.getElementById("deal-input-equity").value = currentDeal.EquityCommission || "";
  document.getElementById("deal-input-retainer").value = currentDeal.Retainer || "";
  document.getElementById("deal-input-keywords").value = normalizeDealKeywords(currentDeal).join("\n");
  document.getElementById("deal-input-summary").value = currentDeal.summary || "";
  normalizeDealLegalLinks(currentDeal);
  document.getElementById("deal-input-deck-name").value = currentDeal.deckName || "";
  document.getElementById("deal-input-deck-url").value = currentDeal.deckUrl || "";
  document.getElementById("deal-input-deck-parent-path").value = currentDeal.deckParentPath || "";
  renderDealLegalLinkEditor();
  refreshDeckEditorSummary();
  renderDealContactEditor();
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
  const addLegalLinkBtn = document.getElementById("btn-add-legal-link-row");
  const legalLinkEditorList = document.getElementById("deal-legal-editor-list");
  const addContactBtn = document.getElementById("btn-add-contact-row");
  const contactEditorList = document.getElementById("deal-contact-editor-list");

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

  if (addLegalLinkBtn) {
    addLegalLinkBtn.addEventListener("click", () => {
      if (!currentDeal) return;
      const legalLinks = normalizeDealLegalLinks(currentDeal, { keepEmpty: true });
      legalLinks.push({
        title: "",
        url: "",
      });
      currentDeal.legalLinks = legalLinks;
      renderDealLegalLinkEditor();
      if (statusEl) statusEl.textContent = "Legal link row added in form";
    });
  }

  if (legalLinkEditorList) {
    legalLinkEditorList.addEventListener("click", (event) => {
      const removeBtn = event.target.closest("button[data-remove-legal-link-row]");
      if (!removeBtn || !currentDeal) return;
      const index = Number(removeBtn.getAttribute("data-remove-legal-link-row"));
      const legalLinks = normalizeDealLegalLinks(currentDeal, { keepEmpty: true });
      if (!Number.isFinite(index) || !legalLinks[index]) return;
      legalLinks.splice(index, 1);
      currentDeal.legalLinks = legalLinks;
      renderDealLegalLinkEditor();
      if (statusEl) statusEl.textContent = "Legal link removed in form";
    });
  }

  if (addContactBtn) {
    addContactBtn.addEventListener("click", () => {
      if (!currentDeal) return;
      const contacts = normalizeDealContacts(currentDeal);
      contacts.push({
        name: "",
        title: "",
        email: "",
        isPrimary: !contacts.length,
      });
      currentDeal.contacts = contacts;
      renderDealContactEditor();
      if (statusEl) statusEl.textContent = "Contact row added in form";
    });
  }

  if (contactEditorList) {
    contactEditorList.addEventListener("click", (event) => {
      const removeBtn = event.target.closest("button[data-remove-contact-row]");
      if (!removeBtn || !currentDeal) return;
      const index = Number(removeBtn.getAttribute("data-remove-contact-row"));
      const contacts = normalizeDealContacts(currentDeal);
      if (!Number.isFinite(index) || !contacts[index]) return;
      contacts.splice(index, 1);
      if (contacts.length && !contacts.some((entry) => entry.isPrimary)) {
        contacts[0].isPrimary = true;
      }
      currentDeal.contacts = contacts;
      renderDealContactEditor();
      if (statusEl) statusEl.textContent = "Contact removed in form";
    });
  }

  form.addEventListener("submit", async (event) => {
    event.preventDefault();
    if (!currentDeal) return;

    const previousName = String(currentDeal.name || "").trim();
    const previousCompany = String(currentDeal.company || "").trim();
    const previousSubOwners = normalizeDealSubOwners(currentDeal).slice();

    currentDeal.name = document.getElementById("deal-input-name").value.trim();
    currentDeal.company = document.getElementById("deal-input-company").value.trim();
    currentDeal.seniorOwner = document.getElementById("deal-input-senior").value.trim();
    currentDeal.juniorOwner = document.getElementById("deal-input-junior").value.trim();
    currentDeal.subOwners = parseDealSubOwnersInput(document.getElementById("deal-input-sub-owners").value);
    currentDeal.owner = currentDeal.seniorOwner; // legacy compatibility
    currentDeal.stage = document.getElementById("deal-input-stage").value;
    currentDeal.lifecycleStatus = normalizeDealLifecycleStatus(document.getElementById("deal-input-lifecycle-status").value);
    currentDeal.dealStatus = currentDeal.lifecycleStatus;
    currentDeal.sectors = parseDealSectorsInput(document.getElementById("deal-input-sector").value);
    currentDeal.sector = currentDeal.sectors.join(", ");
    currentDeal.location = document.getElementById("deal-input-location").value.trim();
    currentDeal.fundingStage = document.getElementById("deal-input-funding-stage").value.trim();
    currentDeal.revenue = document.getElementById("deal-input-revenue").value.trim();
    currentDeal.targetAmount = parseNumericAmount(document.getElementById("deal-input-target").value);
    currentDeal.raisedAmount = parseNumericAmount(document.getElementById("deal-input-raised").value);
    currentDeal.currency = document.getElementById("deal-input-currency").value.trim() || "USD";
    currentDeal.fundraisingDashboardId = document.getElementById("deal-input-dashboard").value.trim();
    currentDeal.CashCommission = document.getElementById("deal-input-cash").value.trim();
    currentDeal.EquityCommission = document.getElementById("deal-input-equity").value.trim();
    currentDeal.Retainer = document.getElementById("deal-input-retainer").value.trim();
    currentDeal.keywords = parseDealKeywordsInput(document.getElementById("deal-input-keywords").value);
    currentDeal.summary = document.getElementById("deal-input-summary").value.trim();
    const legalLinks = collectDealLegalLinksFromEditor();
    const invalidLegalLink = legalLinks.find((entry) => !toSafeExternalUrl(entry.url));
    if (invalidLegalLink) {
      statusEl.textContent = `Legal link "${invalidLegalLink.title || "Untitled"}" needs a valid http(s) URL.`;
      return;
    }
    currentDeal.legalLinks = legalLinks.map((entry, index) => ({
      title: entry.title || `Legal link ${index + 1}`,
      url: toSafeExternalUrl(entry.url),
    }));
    currentDeal.deckName = document.getElementById("deal-input-deck-name").value.trim();
    currentDeal.deckUrl = document.getElementById("deal-input-deck-url").value.trim();
    currentDeal.deckParentPath = document.getElementById("deal-input-deck-parent-path").value.trim();
    currentDeal.contacts = collectDealContactsFromEditor();
    syncLegacyPrimaryContactFields(currentDeal);

    const shouldSyncSubOwners =
      previousName !== currentDeal.name ||
      previousCompany !== currentDeal.company ||
      previousSubOwners.join("\n") !== currentDeal.subOwners.join("\n");

    statusEl.textContent = "Saving...";
    try {
      await saveDealsData();
      let subOwnerSyncResult = null;
      let subOwnerSyncError = null;
      if (shouldSyncSubOwners) {
        try {
          subOwnerSyncResult = await syncDealSubOwnersToOwnershipWorkbook({
            previousName,
            previousCompany,
          });
        } catch (error) {
          subOwnerSyncError = error;
        }
      }
      renderDealHeader();
      populateDealForm();
      if (subOwnerSyncError) {
        statusEl.textContent = `Saved locally, but sub owner workbook sync failed: ${subOwnerSyncError.message || "Unknown error"}`;
      } else if (subOwnerSyncResult && subOwnerSyncResult.attempted) {
        statusEl.textContent = subOwnerSyncResult.syncedCount
          ? `Saved and synced ${subOwnerSyncResult.syncedCount} sub owner${subOwnerSyncResult.syncedCount === 1 ? "" : "s"}`
          : "Saved and cleared sub owner sync rows";
      } else {
        statusEl.textContent = currentDeal.fundraisingDashboardId || currentDeal.deckUrl ? "Saved and linked" : "Saved";
      }
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

  refreshDealTaskOwnerSuggestions();

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
    await refreshAccountingAccessState();
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
    window.addEventListener("appcore:graph-session-updated", async () => {
      await refreshAccountingAccessState();
      if (currentDeal) renderDealHeader();
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
  setupDealLifecycleButtons();
  setupStageProgressionButton();
  syncContactStatusTasksFromDashboard();
}

document.addEventListener("DOMContentLoaded", loadDealPage);
