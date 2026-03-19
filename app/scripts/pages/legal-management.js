const AppCore = window.AppCore;
const DEALS_STORAGE_KEY = (AppCore && AppCore.STORAGE_KEYS && AppCore.STORAGE_KEYS.deals) || "deals_data_v1";
const LEGAL_SHAREDRIVE_URL_STORAGE_KEY = "sharedrive_url_v2";

let dealsData = [];
let currentSearch = "";
let showOnlyMissingLegal = false;
let selectedDealReference = "";
let selectedDealId = "";
const dirtyDealIds = new Set();
const legalPickerState = {
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
  const query = new URLSearchParams();
  Object.entries(params || {}).forEach(([key, value]) => {
    if (value == null || value === "") return;
    query.set(key, String(value));
  });
  const queryString = query.toString();
  return queryString ? `${pageId}.html?${queryString}` : `${pageId}.html`;
}

function loadDealsData() {
  dealsData = AppCore ? AppCore.loadDealsData() : (Array.isArray(window.DEALS) ? JSON.parse(JSON.stringify(window.DEALS)) : []);
}

function saveDealsData() {
  if (AppCore) return AppCore.saveDealsData(dealsData);
  try {
    localStorage.setItem(DEALS_STORAGE_KEY, JSON.stringify(dealsData));
  } catch (error) {
    console.warn("Failed to save deals to storage", error);
  }
  return Promise.resolve();
}

function getDashboardConfig() {
  if (AppCore && typeof AppCore.getDashboardConfig === "function") {
    return AppCore.getDashboardConfig();
  }
  return window.DASHBOARD_CONFIG || { dashboards: [], settings: {} };
}

function getDefaultLegalShareUrl() {
  try {
    const remembered = String(localStorage.getItem(LEGAL_SHAREDRIVE_URL_STORAGE_KEY) || "").trim();
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

function escapeHtml(value) {
  return String(value == null ? "" : value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

function getSelectedDealReference() {
  try {
    const params = new URLSearchParams(window.location.search || "");
    return String(params.get("id") || params.get("deal") || "").trim();
  } catch {
    return "";
  }
}

function findDealByReference(reference) {
  if (!Array.isArray(dealsData) || !dealsData.length) return null;
  if (AppCore && typeof AppCore.findDealByReference === "function") {
    return AppCore.findDealByReference(dealsData, reference);
  }

  const key = normalizeValue(reference);
  if (!key) return null;
  return dealsData.find((deal) => {
    const refs = [
      deal && deal.id,
      deal && deal.name,
      deal && deal.company,
    ].map((value) => normalizeValue(value)).filter(Boolean);
    return refs.includes(key);
  }) || null;
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

function deriveLegalTitleFromFileName(value) {
  const raw = String(value || "").trim();
  if (!raw) return "Legal document";
  const withoutExtension = raw.replace(/\.[a-z0-9]+$/i, "");
  return withoutExtension || raw;
}

function isLegalDocumentItem(item) {
  const name = String((item && item.name) || "").trim().toLowerCase();
  const mimeType = String((item && item.mimeType) || "").trim().toLowerCase();
  if (/\.(pdf|doc|docx|xls|xlsx|ppt|pptx)$/i.test(name)) return true;
  return [
    "application/pdf",
    "application/msword",
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    "application/vnd.ms-excel",
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    "application/vnd.ms-powerpoint",
    "application/vnd.openxmlformats-officedocument.presentationml.presentation",
  ].includes(mimeType);
}

function getLegalLinkStats(deal) {
  const links = normalizeDealLegalLinks(deal);
  const validCount = links.filter((link) => Boolean(toSafeExternalUrl(link.url))).length;
  return {
    total: links.length,
    valid: validCount,
    invalid: Math.max(links.length - validCount, 0),
  };
}

function markDirty(dealId, isDirty) {
  const key = normalizeValue(dealId);
  if (!key) return;
  if (isDirty) dirtyDealIds.add(key);
  else dirtyDealIds.delete(key);
}

function setStatus(message, isError) {
  const line = document.getElementById("legal-status-line");
  const pill = document.getElementById("legal-status-pill");
  if (line) {
    line.textContent = message || "";
    line.style.color = isError ? "#ef4444" : "var(--text-soft)";
  }
  if (pill) {
    pill.textContent = isError ? "Attention needed" : (dirtyDealIds.size ? "Unsaved changes" : "Ready");
  }
}

function getVisibleDeals() {
  const query = normalizeValue(currentSearch);
  const source = Array.isArray(dealsData) ? dealsData.slice() : [];
  const filtered = source.filter((deal) => {
    const haystack = [
      deal && deal.id,
      deal && deal.name,
      deal && deal.company,
    ].map((value) => normalizeValue(value)).join(" ");
    if (query && !haystack.includes(query)) return false;
    if (showOnlyMissingLegal && getLegalLinkStats(deal).total > 0) return false;
    return true;
  });

  filtered.sort((left, right) => normalizeValue(left.company || left.name).localeCompare(normalizeValue(right.company || right.name)));
  return filtered;
}

function getSelectedDeal() {
  if (!selectedDealId) return null;
  return dealsData.find((deal) => normalizeValue(deal.id) === normalizeValue(selectedDealId)) || null;
}

function ensureSelectedDeal(visibleDeals) {
  const visible = Array.isArray(visibleDeals) ? visibleDeals : [];
  const currentSelected = getSelectedDeal();
  if (currentSelected && visible.some((deal) => normalizeValue(deal.id) === normalizeValue(currentSelected.id))) {
    return currentSelected;
  }

  if (selectedDealReference) {
    const matchedByReference = visible.find((deal) => {
      const refs = [deal.id, deal.name, deal.company].map((value) => normalizeValue(value));
      return refs.includes(normalizeValue(selectedDealReference));
    });
    if (matchedByReference) {
      selectedDealId = matchedByReference.id;
      selectedDealReference = "";
      return matchedByReference;
    }
  }

  const first = visible[0] || null;
  selectedDealId = first ? first.id : "";
  return first;
}

function renderMetaRow(deals) {
  const row = document.getElementById("legal-meta-row");
  if (!row) return;

  const allDeals = Array.isArray(deals) ? deals : [];
  const withLegalLinks = allDeals.filter((deal) => getLegalLinkStats(deal).total > 0).length;
  const missingLegalLinks = allDeals.filter((deal) => getLegalLinkStats(deal).total === 0).length;
  const totalLegalLinks = allDeals.reduce((sum, deal) => sum + getLegalLinkStats(deal).total, 0);

  row.innerHTML = [
    `<div class="chip"><strong>${allDeals.length}</strong> deals shown</div>`,
    `<div class="chip"><strong>${withLegalLinks}</strong> with legal links</div>`,
    `<div class="chip"><strong>${missingLegalLinks}</strong> missing legal</div>`,
    `<div class="chip"><strong>${totalLegalLinks}</strong> legal links total</div>`,
    `<div class="chip"><strong>${dirtyDealIds.size}</strong> unsaved changes</div>`,
  ].join("");
}

function buildLegalStatusBadge(stats) {
  if (!stats.total) {
    return '<span class="legal-status-badge is-missing">Missing</span>';
  }
  if (stats.invalid) {
    return `<span class="legal-status-badge is-warning">${stats.invalid} invalid</span>`;
  }
  return '<span class="legal-status-badge is-ready">Ready</span>';
}

function renderDealsTable() {
  const body = document.getElementById("legal-deals-body");
  const toggleBtn = document.getElementById("btn-toggle-missing-legal");
  if (!body) return;

  const visibleDeals = getVisibleDeals();
  const selectedDeal = ensureSelectedDeal(visibleDeals);
  renderMetaRow(visibleDeals);

  if (toggleBtn) {
    toggleBtn.textContent = showOnlyMissingLegal ? "Show all deals" : "Show only missing legal";
    toggleBtn.classList.toggle("is-active", showOnlyMissingLegal);
  }

  if (!visibleDeals.length) {
    body.innerHTML = '<tr><td colspan="5">No deals match this filter.</td></tr>';
    renderEditorPanel();
    return;
  }

  body.innerHTML = visibleDeals.map((deal) => {
    const stats = getLegalLinkStats(deal);
    const isSelected = selectedDeal && normalizeValue(selectedDeal.id) === normalizeValue(deal.id);
    const openDealHref = buildPageUrl("deal-details", { id: deal.id });
    return `
      <tr class="legal-row${isSelected ? " is-selected" : ""}" data-select-deal-id="${escapeHtml(String(deal.id || ""))}">
        <td><strong>${escapeHtml(String(deal.name || "Untitled deal"))}</strong></td>
        <td>${escapeHtml(String(deal.company || "—"))}</td>
        <td><span class="legal-count"><strong>${stats.total}</strong> linked</span></td>
        <td>${buildLegalStatusBadge(stats)}</td>
        <td><a class="action-link" href="${openDealHref}">Open deal</a></td>
      </tr>
    `;
  }).join("");

  renderEditorPanel();
}

function renderDealsTableBodyOnly() {
  const body = document.getElementById("legal-deals-body");
  const toggleBtn = document.getElementById("btn-toggle-missing-legal");
  if (!body) return;

  const visibleDeals = getVisibleDeals();
  const selectedDeal = ensureSelectedDeal(visibleDeals);
  renderMetaRow(visibleDeals);

  if (toggleBtn) {
    toggleBtn.textContent = showOnlyMissingLegal ? "Show all deals" : "Show only missing legal";
    toggleBtn.classList.toggle("is-active", showOnlyMissingLegal);
  }

  if (!visibleDeals.length) {
    body.innerHTML = '<tr><td colspan="5">No deals match this filter.</td></tr>';
    return;
  }

  body.innerHTML = visibleDeals.map((deal) => {
    const stats = getLegalLinkStats(deal);
    const isSelected = selectedDeal && normalizeValue(selectedDeal.id) === normalizeValue(deal.id);
    const openDealHref = buildPageUrl("deal-details", { id: deal.id });
    return `
      <tr class="legal-row${isSelected ? " is-selected" : ""}" data-select-deal-id="${escapeHtml(String(deal.id || ""))}">
        <td><strong>${escapeHtml(String(deal.name || "Untitled deal"))}</strong></td>
        <td>${escapeHtml(String(deal.company || "—"))}</td>
        <td><span class="legal-count"><strong>${stats.total}</strong> linked</span></td>
        <td>${buildLegalStatusBadge(stats)}</td>
        <td><a class="action-link" href="${openDealHref}">Open deal</a></td>
      </tr>
    `;
  }).join("");
}

function buildEditorRow(link, index, isDirty) {
  return `
    <div class="legal-editor-row${isDirty ? " is-dirty" : ""}" data-legal-row="${index}">
      <div class="legal-editor-grid">
        <label>
          Title
          <input type="text" data-legal-field="title" value="${escapeHtml(link.title || "")}" placeholder="Term sheet, NDA, counsel notes..." />
        </label>
        <label>
          URL
          <input type="url" data-legal-field="url" value="${escapeHtml(link.url || "")}" placeholder="https://..." />
        </label>
      </div>
      <div class="legal-editor-actions">
        <div class="hint">Use a full SharePoint, Google Drive, DocuSign, or other legal document link.</div>
        <button class="btn" type="button" data-remove-legal-index="${index}">Remove</button>
      </div>
    </div>
  `;
}

function collectLegalLinksFromEditor() {
  const list = document.getElementById("legal-editor-list");
  if (!list) return [];
  return Array.from(list.querySelectorAll("[data-legal-row]"))
    .map((row, index) => {
      const title = String((row.querySelector('[data-legal-field="title"]') || {}).value || "").trim();
      const url = String((row.querySelector('[data-legal-field="url"]') || {}).value || "").trim();
      if (!title && !url) return null;
      return {
        title: title || `Legal link ${index + 1}`,
        url,
      };
    })
    .filter(Boolean);
}

function persistSelectedDealLinks() {
  const selectedDeal = getSelectedDeal();
  if (!selectedDeal) return;
  selectedDeal.legalLinks = collectLegalLinksFromEditor();
  markDirty(selectedDeal.id, true);
}

function renderPreviewGrid() {
  const grid = document.getElementById("legal-preview-grid");
  const selectedDeal = getSelectedDeal();
  if (!grid) return;

  if (!selectedDeal) {
    grid.innerHTML = "";
    return;
  }

  const links = normalizeDealLegalLinks(selectedDeal);
  grid.innerHTML = links.length
    ? links.map((link) => {
      const safeUrl = toSafeExternalUrl(link.url);
      if (!safeUrl) {
        return `
          <div class="legal-preview-card is-invalid">
            <div class="legal-preview-title">${escapeHtml(link.title || "Legal link")}</div>
            <div class="legal-preview-meta">This entry needs a valid http(s) URL.</div>
          </div>
        `;
      }

      return `
        <div class="legal-preview-card">
          <div class="legal-preview-title">${escapeHtml(link.title || "Legal link")}</div>
          <div class="legal-preview-meta">${escapeHtml(safeUrl)}</div>
          <div class="btn-row" style="margin-top:10px;">
            <a class="btn" href="${escapeHtml(safeUrl)}" target="_blank" rel="noopener noreferrer">Open</a>
          </div>
        </div>
      `;
    }).join("")
    : '<div class="legal-empty">No legal links saved for this deal yet.</div>';
}

function renderEditorPanel() {
  const panel = document.getElementById("legal-editor-panel");
  const title = document.getElementById("legal-editor-title");
  const subtitle = document.getElementById("legal-editor-subtitle");
  const list = document.getElementById("legal-editor-list");
  const openDealBtn = document.getElementById("btn-open-selected-deal");
  const selectedDeal = getSelectedDeal();
  if (!panel || !title || !subtitle || !list || !openDealBtn) return;

  if (!selectedDeal) {
    panel.hidden = true;
    return;
  }

  panel.hidden = false;
  const stats = getLegalLinkStats(selectedDeal);
  title.textContent = `${String(selectedDeal.name || "Deal")} · Legal links`;
  subtitle.textContent = `${String(selectedDeal.company || "Company")} · ${stats.total} legal link${stats.total === 1 ? "" : "s"} tracked`;
  openDealBtn.onclick = () => {
    window.location.href = buildPageUrl("deal-details", { id: selectedDeal.id });
  };

  const links = normalizeDealLegalLinks(selectedDeal, { keepEmpty: true });
  const isDirty = dirtyDealIds.has(normalizeValue(selectedDeal.id));
  list.innerHTML = links.length
    ? links.map((link, index) => buildEditorRow(link, index, isDirty)).join("")
    : '<div class="legal-empty">No legal links yet. Add one to start managing the legal aspects for this deal.</div>';

  renderPreviewGrid();
}

function closeLegalPicker() {
  const modal = document.getElementById("legal-picker-modal");
  if (modal) modal.hidden = true;
}

function setLegalPickerStatus(message) {
  const statusEl = document.getElementById("legal-picker-status");
  if (statusEl) statusEl.textContent = message || "";
}

function renderLegalPickerPath() {
  const pathEl = document.getElementById("legal-picker-path");
  if (!pathEl) return;

  if (!legalPickerState.root) {
    pathEl.textContent = "Not connected to a Sharedrive folder yet.";
    return;
  }

  const names = legalPickerState.stack.map((entry) => entry.name).filter(Boolean);
  pathEl.textContent = names.join(" / ") || legalPickerState.root.name || "Sharedrive root";
}

function renderLegalPickerItems(items) {
  const listEl = document.getElementById("legal-picker-list");
  if (!listEl) return;

  const sorted = (Array.isArray(items) ? items.slice() : []).sort((a, b) => {
    if (Boolean(a.isFolder) !== Boolean(b.isFolder)) return a.isFolder ? -1 : 1;
    return String(a.name || "").localeCompare(String(b.name || ""));
  });
  legalPickerState.items = sorted;

  const filtered = sorted.filter((item) => item.isFolder || isLegalDocumentItem(item));
  if (!filtered.length) {
    listEl.innerHTML = '<div class="legal-empty">No folders or supported legal documents found here.</div>';
    return;
  }

  listEl.innerHTML = filtered.map((item) => {
    const title = String(item.name || "(Unnamed item)");
    const meta = item.isFolder
      ? `${item.childCount != null ? `${item.childCount} item${item.childCount === 1 ? "" : "s"}` : "Folder"}`
      : `${item.mimeType || "Document"}${item.parentPath ? ` · ${item.parentPath}` : ""}`;
    const actionLabel = item.isFolder ? "Open folder" : "Use this document";
    return `
      <div class="legal-picker-item">
        <div class="legal-picker-item-copy">
          <div class="legal-picker-item-title">${escapeHtml(title)}</div>
          <div class="legal-picker-item-meta">${escapeHtml(meta)}</div>
        </div>
        <div class="legal-picker-item-actions">
          <button class="btn" type="button" data-legal-item-id="${escapeHtml(String(item.id || ""))}" data-legal-item-kind="${item.isFolder ? "folder" : "document"}">${actionLabel}</button>
        </div>
      </div>
    `;
  }).join("");
}

async function requestLegalPickerChildren(parentItemId) {
  const shareUrl = String(legalPickerState.shareUrl || "").trim();
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

async function loadLegalPickerFolder(targetEntry, options = {}) {
  const { resetStack = false } = options;
  const folderId = targetEntry && targetEntry.id ? targetEntry.id : "";

  setLegalPickerStatus("Loading Sharedrive folder...");
  const data = await requestLegalPickerChildren(folderId);
  legalPickerState.root = data.root || legalPickerState.root;

  if (resetStack) {
    const root = data.root || {};
    legalPickerState.stack = [{ id: data.parentItemId || root.id || "", name: root.name || "Sharedrive root" }];
  } else if (targetEntry && targetEntry.id) {
    legalPickerState.stack.push({ id: targetEntry.id, name: targetEntry.name || "Folder" });
  }

  renderLegalPickerPath();
  renderLegalPickerItems(Array.isArray(data.items) ? data.items : []);
  setLegalPickerStatus(`Showing ${Array.isArray(data.items) ? data.items.length : 0} item(s)`);
}

function openLegalPicker() {
  const modal = document.getElementById("legal-picker-modal");
  const shareUrlInput = document.getElementById("legal-share-url-input");
  const selectedDeal = getSelectedDeal();
  if (!modal || !shareUrlInput || !selectedDeal) return;

  modal.hidden = false;
  legalPickerState.shareUrl = legalPickerState.shareUrl || getDefaultLegalShareUrl();
  shareUrlInput.value = legalPickerState.shareUrl;
  renderLegalPickerPath();

  if (!legalPickerState.shareUrl) {
    setLegalPickerStatus("Paste a Sharedrive folder URL to start browsing.");
    return;
  }

  loadLegalPickerFolder(null, { resetStack: true }).catch((error) => {
    setLegalPickerStatus(error instanceof Error ? error.message : "Failed to load Sharedrive items.");
  });
}

function attachSelectedLegalDocument(item) {
  const selectedDeal = getSelectedDeal();
  if (!selectedDeal || !item) return;

  const url = String((item && item.webUrl) || "").trim();
  if (!toSafeExternalUrl(url)) {
    setStatus("Selected file does not have a valid document URL.", true);
    return;
  }

  const links = normalizeDealLegalLinks(selectedDeal, { keepEmpty: true });
  links.push({
    title: deriveLegalTitleFromFileName(item.name || "Legal document"),
    url,
  });
  selectedDeal.legalLinks = links;
  markDirty(selectedDeal.id, true);
  renderDealsTable();
  closeLegalPicker();
  setStatus("Legal document attached from Sharedrive. Save all changes to persist.", false);
}

function updateSelectedDealFromEditor() {
  const selectedDeal = getSelectedDeal();
  if (!selectedDeal) return;
  selectedDeal.legalLinks = collectLegalLinksFromEditor();
  markDirty(selectedDeal.id, true);
  renderPreviewGrid();
  renderMetaRow(getVisibleDeals());
  setStatus("Legal links updated. Save all changes to persist.", false);
}

async function saveAllChanges() {
  if (!dirtyDealIds.size) {
    setStatus("No changes to save.", false);
    return;
  }

  setStatus("Saving legal updates...", false);
  try {
    await saveDealsData();
    dirtyDealIds.clear();
    renderDealsTable();
    setStatus("Legal updates saved to the shared online deals data.", false);
  } catch (error) {
    setStatus(error instanceof Error ? error.message : "Failed to save legal updates.", true);
  }
}

function setupInteractions() {
  const body = document.getElementById("legal-deals-body");
  const editorList = document.getElementById("legal-editor-list");
  const addBtn = document.getElementById("btn-add-legal-row");
  const browseBtn = document.getElementById("btn-browse-legal-docs");
  const saveBtn = document.getElementById("save-legal-btn");
  const searchInput = document.getElementById("legal-search");
  const toggleMissingBtn = document.getElementById("btn-toggle-missing-legal");
  const pickerModal = document.getElementById("legal-picker-modal");
  const pickerCloseBtn = document.getElementById("btn-close-legal-picker");
  const pickerLoadBtn = document.getElementById("btn-load-legal-root");
  const pickerBackBtn = document.getElementById("btn-legal-go-back");
  const pickerShareUrlInput = document.getElementById("legal-share-url-input");
  const pickerList = document.getElementById("legal-picker-list");

  if (body) {
    body.addEventListener("click", (event) => {
      const row = event.target.closest("[data-select-deal-id]");
      if (!row) return;
      selectedDealId = String(row.getAttribute("data-select-deal-id") || "").trim();
      renderDealsTable();
    });
  }

  if (editorList) {
    editorList.addEventListener("input", () => {
      const previouslySelectedDeal = getSelectedDeal();
      const previousSelectedId = normalizeValue(previouslySelectedDeal && previouslySelectedDeal.id);
      updateSelectedDealFromEditor();
      const visibleDeals = getVisibleDeals();
      const selectedStillVisible = previousSelectedId && visibleDeals.some((deal) => normalizeValue(deal.id) === previousSelectedId);
      if (selectedStillVisible) {
        renderDealsTableBodyOnly();
      } else {
        renderDealsTable();
      }
    });

    editorList.addEventListener("click", (event) => {
      const removeBtn = event.target.closest("[data-remove-legal-index]");
      if (!removeBtn) return;
      const index = Number(removeBtn.getAttribute("data-remove-legal-index"));
      const selectedDeal = getSelectedDeal();
      if (!selectedDeal) return;
      const links = normalizeDealLegalLinks(selectedDeal, { keepEmpty: true });
      if (!Number.isFinite(index) || !links[index]) return;
      links.splice(index, 1);
      selectedDeal.legalLinks = links;
      markDirty(selectedDeal.id, true);
      renderDealsTable();
      setStatus("Legal link removed. Save all changes to persist.", false);
    });
  }

  if (addBtn) {
    addBtn.addEventListener("click", () => {
      const selectedDeal = getSelectedDeal();
      if (!selectedDeal) return;
      const links = normalizeDealLegalLinks(selectedDeal, { keepEmpty: true });
      links.push({
        title: "",
        url: "",
      });
      selectedDeal.legalLinks = links;
      markDirty(selectedDeal.id, true);
      renderEditorPanel();
      setStatus("Legal link row added. Save all changes to persist.", false);
    });
  }

  if (browseBtn) {
    browseBtn.addEventListener("click", () => {
      openLegalPicker();
    });
  }

  if (saveBtn) {
    saveBtn.addEventListener("click", saveAllChanges);
  }

  if (searchInput) {
    searchInput.addEventListener("input", (event) => {
      currentSearch = String(event.target && event.target.value || "");
      renderDealsTable();
    });
  }

  if (toggleMissingBtn) {
    toggleMissingBtn.addEventListener("click", () => {
      showOnlyMissingLegal = !showOnlyMissingLegal;
      renderDealsTable();
    });
  }

  if (pickerCloseBtn) {
    pickerCloseBtn.addEventListener("click", closeLegalPicker);
  }

  if (pickerModal) {
    pickerModal.addEventListener("click", (event) => {
      if (event.target === pickerModal) closeLegalPicker();
    });
  }

  document.addEventListener("keydown", (event) => {
    if (event.key === "Escape" && pickerModal && !pickerModal.hidden) {
      closeLegalPicker();
    }
  });

  if (pickerLoadBtn && pickerShareUrlInput) {
    pickerLoadBtn.addEventListener("click", async () => {
      legalPickerState.shareUrl = String(pickerShareUrlInput.value || "").trim();
      if (!legalPickerState.shareUrl) {
        setLegalPickerStatus("Paste a Sharedrive folder URL first.");
        return;
      }

      try {
        localStorage.setItem(LEGAL_SHAREDRIVE_URL_STORAGE_KEY, legalPickerState.shareUrl);
      } catch {
        // ignore storage failures
      }

      try {
        await loadLegalPickerFolder(null, { resetStack: true });
      } catch (error) {
        setLegalPickerStatus(error instanceof Error ? error.message : "Failed to load Sharedrive items.");
      }
    });
  }

  if (pickerBackBtn) {
    pickerBackBtn.addEventListener("click", async () => {
      if (legalPickerState.stack.length <= 1) {
        setLegalPickerStatus("Already at the root folder.");
        return;
      }

      legalPickerState.stack.pop();
      const current = legalPickerState.stack[legalPickerState.stack.length - 1];
      try {
        setLegalPickerStatus("Loading previous folder...");
        const data = await requestLegalPickerChildren(current.id);
        renderLegalPickerPath();
        renderLegalPickerItems(Array.isArray(data.items) ? data.items : []);
        setLegalPickerStatus(`Showing ${Array.isArray(data.items) ? data.items.length : 0} item(s)`);
      } catch (error) {
        setLegalPickerStatus(error instanceof Error ? error.message : "Failed to load Sharedrive items.");
      }
    });
  }

  if (pickerList) {
    pickerList.addEventListener("click", async (event) => {
      const button = event.target.closest("button[data-legal-item-id]");
      if (!button) return;

      const itemId = String(button.getAttribute("data-legal-item-id") || "").trim();
      const kind = String(button.getAttribute("data-legal-item-kind") || "").trim();
      if (!itemId) return;

      const selectedItem = (legalPickerState.items || []).find((entry) => String(entry.id || "").trim() === itemId);
      if (!selectedItem) return;

      if (kind === "folder") {
        try {
          await loadLegalPickerFolder({ id: itemId, name: selectedItem.name || "Folder" });
        } catch (error) {
          setLegalPickerStatus(error instanceof Error ? error.message : "Failed to load Sharedrive items.");
        }
        return;
      }

      attachSelectedLegalDocument(selectedItem);
    });
  }
}

function initializeLegalManagementPage() {
  selectedDealReference = getSelectedDealReference();
  loadDealsData();
  setupInteractions();
  renderDealsTable();
  setStatus("No unsaved changes.", false);

  if (AppCore) {
    window.addEventListener("appcore:deals-updated", (event) => {
      if (!event || !event.detail || !Array.isArray(event.detail.deals)) return;
      dealsData = event.detail.deals.slice();
      renderDealsTable();
    });
  }
}

document.addEventListener("DOMContentLoaded", initializeLegalManagementPage);
