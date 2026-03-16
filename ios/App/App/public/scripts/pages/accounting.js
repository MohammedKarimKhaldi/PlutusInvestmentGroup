const AppCore = window.AppCore;
const DEALS_STORAGE_KEY = (AppCore && AppCore.STORAGE_KEYS && AppCore.STORAGE_KEYS.deals) || "deals_data_v1";

let dealsData = [];
const dirtyDealIds = new Set();
let currentSearch = "";
let selectedDealReference = "";
let singleDealMode = false;
const INVOICE_SHAREDRIVE_URL_STORAGE_KEY = "invoice_sharedrive_url_v1";
const invoicePickerState = {
  shareUrl: "",
  root: null,
  stack: [],
  items: [],
};
const PLUTUS_INVOICE_ISSUER = {
  company: "Plutus Investment Group LLP",
  lines: ["83 Baker Street", "London", "W1U 6AG"],
  email: "mj@plutus-investment.com",
  phone: "+447398639471",
  bankHolder: "PLUTUS INVESTMENT GROUP LLP",
  accountNumber: "02297518",
  sortCode: "04-05-11",
  iban: "GB66CLRB04051102297518",
  swift: "CLRBGB22",
};
const INVOICE_TEMPLATE_FILE = "invoice-template.docx";

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

function getSelectedDealReference() {
  try {
    const params = new URLSearchParams(window.location.search || "");
    return String(params.get("id") || params.get("deal") || "").trim();
  } catch {
    return "";
  }
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

function getRetainerMonthly(deal) {
  if (!deal || typeof deal !== "object") return "";
  if (deal.retainerMonthly != null && String(deal.retainerMonthly).trim()) return String(deal.retainerMonthly).trim();
  if (deal.Retainer != null && String(deal.Retainer).trim()) return String(deal.Retainer).trim();
  return "";
}

function getPaymentDay(deal) {
  if (!deal || typeof deal !== "object") return "";
  const raw = deal.retainerPaymentDay != null ? String(deal.retainerPaymentDay).trim() : "";
  if (!raw) return "";
  const parsed = Number(raw);
  if (!Number.isFinite(parsed)) return "";
  return String(Math.min(31, Math.max(1, Math.round(parsed))));
}

function formatDayOrdinal(value) {
  const day = Number(value);
  if (!Number.isFinite(day) || day < 1 || day > 31) return "";
  const mod100 = day % 100;
  if (mod100 >= 11 && mod100 <= 13) return `${day}th`;
  const mod10 = day % 10;
  if (mod10 === 1) return `${day}st`;
  if (mod10 === 2) return `${day}nd`;
  if (mod10 === 3) return `${day}rd`;
  return `${day}th`;
}

function computeNextExpectedDate(dayValue) {
  const day = Number(dayValue);
  if (!Number.isFinite(day) || day < 1 || day > 31) return "Not set";

  const now = new Date();
  const year = now.getFullYear();
  const month = now.getMonth();
  const today = now.getDate();

  const monthDays = new Date(year, month + 1, 0).getDate();
  const targetDayThisMonth = Math.min(day, monthDays);
  if (targetDayThisMonth >= today) {
    return new Date(year, month, targetDayThisMonth).toLocaleDateString("en-GB", { day: "2-digit", month: "short", year: "numeric" });
  }

  const nextMonthDays = new Date(year, month + 2, 0).getDate();
  const targetDayNextMonth = Math.min(day, nextMonthDays);
  return new Date(year, month + 1, targetDayNextMonth).toLocaleDateString("en-GB", { day: "2-digit", month: "short", year: "numeric" });
}

function markDirty(dealId, isDirty) {
  const key = normalizeValue(dealId);
  if (!key) return;
  if (isDirty) dirtyDealIds.add(key);
  else dirtyDealIds.delete(key);
}

function setStatus(message, isError) {
  const line = document.getElementById("accounting-status-line");
  const pill = document.getElementById("accounting-status-text");
  if (line) {
    line.textContent = message || "";
    line.style.color = isError ? "#ef4444" : "var(--text-soft)";
  }
  if (pill) {
    pill.textContent = isError ? "Attention needed" : (dirtyDealIds.size ? "Unsaved changes" : "Ready");
  }
}

function updateMetaRow(filteredDeals) {
  const row = document.getElementById("accounting-meta-row");
  if (!row) return;

  const allDeals = Array.isArray(filteredDeals) ? filteredDeals : [];
  if (singleDealMode && allDeals.length) {
    const deal = allDeals[0];
    const textOrDash = (value) => {
      const raw = String(value == null ? "" : value).trim();
      return raw || "–";
    };
    row.innerHTML = [
      `<div class="chip"><strong>ID</strong> ${textOrDash(deal.id)}</div>`,
      `<div class="chip"><strong>Stage</strong> ${textOrDash(deal.stage)}</div>`,
      `<div class="chip"><strong>Senior</strong> ${textOrDash(deal.seniorOwner || deal.owner)}</div>`,
      `<div class="chip"><strong>Junior</strong> ${textOrDash(deal.juniorOwner)}</div>`,
      `<div class="chip"><strong>Currency</strong> ${textOrDash(deal.currency)}</div>`,
      `<div class="chip"><strong>Cash %</strong> ${textOrDash(deal.CashCommission)}</div>`,
      `<div class="chip"><strong>Equity %</strong> ${textOrDash(deal.EquityCommission)}</div>`,
      `<div class="chip"><strong>Unsaved</strong> ${dirtyDealIds.size}</div>`,
    ].join("");
    return;
  }

  const withRetainer = allDeals.filter((deal) => getRetainerMonthly(deal)).length;
  const withDay = allDeals.filter((deal) => getPaymentDay(deal)).length;

  row.innerHTML = [
    `<div class="chip"><strong>${allDeals.length}</strong> deals shown</div>`,
    `<div class="chip"><strong>${withRetainer}</strong> with retainer</div>`,
    `<div class="chip"><strong>${withDay}</strong> with payment day</div>`,
    `<div class="chip"><strong>${dirtyDealIds.size}</strong> unsaved changes</div>`,
  ].join("");
}

function getVisibleDeals() {
  if (singleDealMode) {
    const selected = findDealByReference(selectedDealReference);
    return selected ? [selected] : [];
  }

  const query = normalizeValue(currentSearch);
  const source = Array.isArray(dealsData) ? dealsData.slice() : [];
  source.sort((a, b) => normalizeValue(a.company || a.name).localeCompare(normalizeValue(b.company || b.name)));
  if (!query) return source;
  return source.filter((deal) => {
    const haystack = [
      deal && deal.name,
      deal && deal.company,
      deal && deal.id,
      deal && deal.seniorOwner,
      deal && deal.juniorOwner,
    ].map((value) => normalizeValue(value)).join(" ");
    return haystack.includes(query);
  });
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
      deal && deal.fundraisingDashboardId,
    ].map((value) => normalizeValue(value)).filter(Boolean);
    return refs.includes(key);
  }) || null;
}

function updatePageHeaderForMode() {
  const title = document.querySelector(".title-block h1");
  const subtitle = document.querySelector(".title-block p");
  const searchInput = document.getElementById("accounting-search");
  if (!singleDealMode) {
    if (title) title.textContent = "Retainers & Payment Days";
    if (subtitle) subtitle.textContent = "Manage monthly retainers and expected payment day for each deal.";
    if (searchInput) searchInput.disabled = false;
    return;
  }

  const deal = findDealByReference(selectedDealReference);
  if (deal) {
    if (title) title.textContent = `Accounting · ${String(deal.name || "Deal")}`;
    if (subtitle) subtitle.textContent = `${String(deal.company || "Company")} · Parameters loaded from deals.json`;
  } else {
    if (title) title.textContent = "Accounting · Deal not found";
    if (subtitle) subtitle.textContent = "The requested deal was not found in deals data.";
  }
  if (searchInput) {
    searchInput.value = "";
    searchInput.disabled = true;
  }
}

function escapeHtml(value) {
  return String(value == null ? "" : value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

function getSelectedDeal() {
  if (!singleDealMode) return null;
  return findDealByReference(selectedDealReference);
}

function isPdfReference(value) {
  return /\.pdf(?:$|[?#])/i.test(String(value || "").trim());
}

function normalizeDealInvoices(deal) {
  const source = deal && Array.isArray(deal.invoices) ? deal.invoices : [];
  const normalized = source
    .map((entry) => {
      if (typeof entry === "string") {
        const raw = entry.trim();
        if (!raw) return null;
        return {
          name: raw.split("/").pop() || "Invoice.pdf",
          url: raw,
          parentPath: "",
          addedAt: "",
        };
      }

      if (!entry || typeof entry !== "object") return null;
      const url = String(entry.url || entry.webUrl || "").trim();
      const name = String(entry.name || entry.fileName || "").trim() || (url.split("/").pop() || "Invoice.pdf");
      return {
        name,
        url,
        parentPath: String(entry.parentPath || "").trim(),
        addedAt: String(entry.addedAt || "").trim(),
      };
    })
    .filter((entry) => entry && entry.url && (isPdfReference(entry.url) || isPdfReference(entry.name)));

  return normalized;
}

function normalizeInvoiceDraft(deal) {
  const source = deal && deal.invoiceDraft && typeof deal.invoiceDraft === "object" ? deal.invoiceDraft : {};
  const invoiceDate = String(source.invoiceDate || "").trim() || new Date().toISOString().slice(0, 10);
  const dueDate = String(source.dueDate || "").trim() || invoiceDate;
  const currentMonthLabel = formatMonthYear(invoiceDate);
  return {
    clientName: String(source.clientName || deal.company || deal.name || "").trim(),
    addressLine1: String(source.addressLine1 || "").trim(),
    addressLine2: String(source.addressLine2 || "").trim(),
    country: String(source.country || "").trim(),
    invoiceNumber: String(source.invoiceNumber || buildDefaultInvoiceNumber(invoiceDate)).trim(),
    invoiceDate,
    dueDate,
    description: String(source.description || `Monthly Retainer Plutus - ${currentMonthLabel}`).trim(),
    amount: String(source.amount || formatAmountInput(getRetainerMonthly(deal))).trim(),
  };
}

function ensureInvoiceDraft(deal) {
  if (!deal || typeof deal !== "object") return null;
  const normalized = normalizeInvoiceDraft(deal);
  deal.invoiceDraft = normalized;
  return normalized;
}

function ensureDealInvoices(deal) {
  if (!deal || typeof deal !== "object") return [];
  const normalized = normalizeDealInvoices(deal);
  deal.invoices = normalized;
  return normalized;
}

function getDashboardConfig() {
  if (AppCore && typeof AppCore.getDashboardConfig === "function") {
    return AppCore.getDashboardConfig();
  }
  return window.DASHBOARD_CONFIG || { dashboards: [], settings: {} };
}

function getDefaultShareUrl() {
  const config = getDashboardConfig();
  const invoiceShareUrl =
    config &&
    config.settings &&
    typeof config.settings.invoiceShareUrl === "string"
      ? config.settings.invoiceShareUrl.trim()
      : "";
  if (invoiceShareUrl) return invoiceShareUrl;

  try {
    const remembered = String(localStorage.getItem(INVOICE_SHAREDRIVE_URL_STORAGE_KEY) || "").trim();
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

  return (
    config &&
    config.settings &&
    config.settings.sharedDeals &&
    typeof config.settings.sharedDeals.shareUrl === "string"
      ? config.settings.sharedDeals.shareUrl.trim()
      : ""
  );
}

function setInvoiceHistoryStatus(message, isError) {
  const node = document.getElementById("invoice-history-status");
  if (!node) return;
  node.textContent = message || "";
  node.style.color = isError ? "#ef4444" : "var(--text-soft)";
}

function setInvoiceBuilderStatus(message, isError) {
  const node = document.getElementById("invoice-builder-status");
  if (!node) return;
  node.textContent = message || "";
  node.style.color = isError ? "#ef4444" : "var(--text-soft)";
}

function formatAddedAt(value) {
  const raw = String(value || "").trim();
  if (!raw) return "Date unknown";
  const dt = new Date(raw);
  if (Number.isNaN(dt.getTime())) return raw;
  return dt.toLocaleDateString("en-GB", { day: "2-digit", month: "short", year: "numeric" });
}

function formatMonthYear(value) {
  const dt = new Date(value);
  if (Number.isNaN(dt.getTime())) {
    const now = new Date();
    return now.toLocaleDateString("en-GB", { month: "long", year: "numeric" });
  }
  return dt.toLocaleDateString("en-GB", { month: "long", year: "numeric" });
}

function formatLongDate(value) {
  const dt = new Date(value);
  if (Number.isNaN(dt.getTime())) return String(value || "").trim();
  return dt.toLocaleDateString("en-GB", { day: "numeric", month: "long", year: "numeric" });
}

function buildDefaultInvoiceNumber(value) {
  const dt = new Date(value);
  const date = Number.isNaN(dt.getTime()) ? new Date() : dt;
  const day = String(date.getDate()).padStart(2, "0");
  const month = String(date.getMonth() + 1).padStart(2, "0");
  const year = String(date.getFullYear()).slice(-2);
  return `01${day}${month}${year}`;
}

function formatAmountInput(value) {
  const raw = String(value == null ? "" : value).trim();
  if (!raw) return "";
  const normalized = raw.replace(/[^\d.,-]/g, "").replace(/,/g, "");
  const parsed = Number(normalized);
  if (!Number.isFinite(parsed)) return raw;
  return parsed.toLocaleString("en-GB", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

function parseAmount(value) {
  const raw = String(value == null ? "" : value).trim();
  if (!raw) return 0;
  const parsed = Number(raw.replace(/[^\d.-]/g, ""));
  return Number.isFinite(parsed) ? parsed : 0;
}

function formatCurrencyAmount(value, withSymbol) {
  const amount = parseAmount(value);
  const formatted = amount.toLocaleString("en-GB", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  return withSymbol ? `GBP ${formatted}` : formatted;
}

function getInvoiceDraftFieldMap() {
  return {
    clientName: document.getElementById("invoice-client-name"),
    addressLine1: document.getElementById("invoice-address-line-1"),
    addressLine2: document.getElementById("invoice-address-line-2"),
    country: document.getElementById("invoice-country"),
    invoiceNumber: document.getElementById("invoice-number"),
    invoiceDate: document.getElementById("invoice-date"),
    dueDate: document.getElementById("invoice-due-date"),
    description: document.getElementById("invoice-description"),
    amount: document.getElementById("invoice-amount"),
  };
}

function renderInvoiceBuilder() {
  const panel = document.getElementById("invoice-builder-panel");
  const subtitle = document.getElementById("invoice-builder-subtitle");
  const fieldMap = getInvoiceDraftFieldMap();
  if (!panel || !subtitle) return;

  if (!singleDealMode) {
    panel.hidden = true;
    return;
  }

  const deal = getSelectedDeal();
  panel.hidden = false;
  if (!deal) {
    subtitle.textContent = "Deal not found for invoice generation.";
    Object.values(fieldMap).forEach((input) => {
      if (input) input.value = "";
    });
    return;
  }

  const draft = ensureInvoiceDraft(deal);
  subtitle.textContent = `Invoice output for ${String(deal.company || deal.name || "this deal")}.`;
  Object.entries(fieldMap).forEach(([key, input]) => {
    if (!input) return;
    input.value = draft && draft[key] != null ? String(draft[key]) : "";
    input.classList.toggle("is-dirty", dirtyDealIds.has(normalizeValue(deal.id)));
  });
}

function buildInvoiceViewModel(deal) {
  const draft = ensureInvoiceDraft(deal) || {};
  return {
    issuer: PLUTUS_INVOICE_ISSUER,
    clientName: draft.clientName || deal.company || deal.name || "",
    addressLine1: draft.addressLine1 || "",
    addressLine2: draft.addressLine2 || "",
    country: draft.country || "",
    invoiceNumber: draft.invoiceNumber || buildDefaultInvoiceNumber(draft.invoiceDate),
    invoiceDate: formatLongDate(draft.invoiceDate),
    dueDate: formatLongDate(draft.dueDate),
    description: draft.description || `Monthly Retainer Plutus - ${formatMonthYear(draft.invoiceDate)}`,
    amountText: formatCurrencyAmount(draft.amount, false),
    amountTotalText: `£${formatCurrencyAmount(draft.amount, false)}`,
  };
}

function buildInvoiceFilename(deal, extension) {
  const draft = ensureInvoiceDraft(deal) || {};
  const base = [
    String(draft.clientName || deal.company || deal.name || "invoice").trim(),
    String(draft.invoiceNumber || "").trim(),
  ]
    .filter(Boolean)
    .join("-");
  const safe = base
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, "-")
    .replace(/^-+|-+$/g, "") || "invoice";
  return `${safe}.${extension}`;
}

function buildInvoiceHtmlDocument(model) {
  const clientLines = [model.clientName, model.addressLine1, model.addressLine2, model.country]
    .filter(Boolean)
    .map((line) => `<div>${escapeHtml(line)}</div>`)
    .join("");

  return `<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8" />
  <title>Invoice ${escapeHtml(model.invoiceNumber)}</title>
  <style>
    body { font-family: Arial, sans-serif; margin: 40px; color: #111827; }
    .issuer { text-align: right; font-size: 11pt; line-height: 1.4; }
    .rule { border-top: 1px solid #d1d5db; margin: 18px 0; }
    .head { width: 100%; border-collapse: collapse; margin-top: 24px; }
    .head td { vertical-align: top; width: 50%; }
    .invoice-title { text-align: right; }
    .invoice-title strong { display: block; font-size: 12pt; margin-bottom: 8px; }
    .details-table, .totals-table { width: 100%; border-collapse: collapse; margin-top: 20px; }
    .details-table th, .details-table td, .totals-table td { padding: 10px 12px; }
    .details-table th { background: #000; color: #fff; text-align: left; }
    .details-table th:last-child, .details-table td:last-child, .totals-table td:last-child { text-align: right; }
    .details-table tbody td { background: #f8f9fa; }
    .totals-table { width: 50%; margin-left: auto; }
    .totals-table .grand td { font-weight: 700; border-top: 1px solid #111827; }
    .payment { margin-top: 28px; border-top: 1px solid #d1d5db; padding-top: 16px; }
    .payment p { margin: 4px 0; }
  </style>
</head>
<body>
  <div class="issuer">
    <div>${escapeHtml(model.issuer.company)}</div>
    <div>${escapeHtml(model.issuer.lines[0])}</div>
    <div>${escapeHtml(model.issuer.lines[1])}</div>
    <div>${escapeHtml(model.issuer.lines[2])}</div>
    <div>${escapeHtml(model.issuer.email)}</div>
    <div>${escapeHtml(model.issuer.phone)}</div>
  </div>
  <div class="rule"></div>
  <table class="head">
    <tr>
      <td>${clientLines}</td>
      <td class="invoice-title">
        <strong>INVOICE ${escapeHtml(model.invoiceNumber)}</strong>
        <div><strong>${escapeHtml(model.invoiceDate)}</strong></div>
        <div style="margin-top:8px; color:#6b7280;">Payment due by ${escapeHtml(model.dueDate)}</div>
      </td>
    </tr>
  </table>
  <table class="details-table">
    <thead>
      <tr><th>Details</th><th>Price (£)</th></tr>
    </thead>
    <tbody>
      <tr><td>${escapeHtml(model.description)}</td><td>${escapeHtml(model.amountText)}</td></tr>
    </tbody>
  </table>
  <table class="totals-table">
    <tr><td>Net Total</td><td>${escapeHtml(model.amountText)}</td></tr>
    <tr class="grand"><td>GBP Total</td><td>${escapeHtml(model.amountTotalText)}</td></tr>
  </table>
  <div class="payment">
    <strong>Payment Details</strong>
    <p><strong>Bank Account Holder Name:</strong> ${escapeHtml(model.issuer.bankHolder)}</p>
    <p><strong>Account number:</strong> ${escapeHtml(model.issuer.accountNumber)}</p>
    <p><strong>Sort code:</strong> ${escapeHtml(model.issuer.sortCode)}</p>
    <p><strong>IBAN:</strong> ${escapeHtml(model.issuer.iban)}</p>
    <p><strong>SWIFT BIC:</strong> ${escapeHtml(model.issuer.swift)}</p>
  </div>
</body>
</html>`;
}

function downloadBlob(blob, fileName) {
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = fileName;
  document.body.appendChild(anchor);
  anchor.click();
  anchor.remove();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

function getInvoiceTemplatePath() {
  if (window.PlutusAppConfig && typeof window.PlutusAppConfig.getDataPath === "function") {
    return window.PlutusAppConfig.getDataPath(INVOICE_TEMPLATE_FILE);
  }
  return `../data/${INVOICE_TEMPLATE_FILE}`;
}

async function loadInvoiceTemplateBuffer() {
  const path = getInvoiceTemplatePath();
  try {
    const response = await fetch(path, { cache: "no-store" });
    if (response.ok) {
      return response.arrayBuffer();
    }
  } catch {
    // Fall through to XHR for file:// and embedded runtimes.
  }

  return new Promise((resolve, reject) => {
    try {
      const request = new XMLHttpRequest();
      request.open("GET", path, true);
      request.responseType = "arraybuffer";
      request.onload = () => {
        if ((request.status >= 200 && request.status < 300) || request.status === 0) {
          resolve(request.response);
          return;
        }
        reject(new Error("Invoice .docx template could not be loaded."));
      };
      request.onerror = () => reject(new Error("Invoice .docx template could not be loaded."));
      request.send();
    } catch {
      reject(new Error("Invoice .docx template could not be loaded."));
    }
  });
}

function getWordTextNodes(xmlDocument) {
  return Array.from(
    xmlDocument.getElementsByTagNameNS(
      "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
      "t",
    ),
  );
}

function replaceTextNodeValue(textNodes, sampleText, nextText, occurrenceIndex = 0) {
  let seen = 0;
  for (const node of textNodes) {
    if (node.textContent !== sampleText) continue;
    if (seen === occurrenceIndex) {
      node.textContent = nextText;
      return true;
    }
    seen += 1;
  }
  return false;
}

function replaceTextSequence(textNodes, sampleValues, replacementValues) {
  for (let index = 0; index <= textNodes.length - sampleValues.length; index += 1) {
    const isMatch = sampleValues.every((value, offset) => textNodes[index + offset].textContent === value);
    if (!isMatch) continue;
    replacementValues.forEach((value, offset) => {
      textNodes[index + offset].textContent = value;
    });
    return true;
  }
  return false;
}

async function generateInvoiceWord() {
  const deal = getSelectedDeal();
  if (!deal) {
    setInvoiceBuilderStatus("No deal loaded for invoice generation.", true);
    return;
  }

  if (!window.JSZip) {
    setInvoiceBuilderStatus("DOCX library is unavailable on this device.", true);
    return;
  }

  try {
    const model = buildInvoiceViewModel(deal);
    const templateBuffer = await loadInvoiceTemplateBuffer();
    const zip = await window.JSZip.loadAsync(templateBuffer);
    const documentXmlFile = zip.file("word/document.xml");
    if (!documentXmlFile) {
      throw new Error("Invoice template is missing word/document.xml.");
    }

    const documentXml = await documentXmlFile.async("string");
    const xmlDocument = new DOMParser().parseFromString(documentXml, "application/xml");
    const textNodes = getWordTextNodes(xmlDocument);

    replaceTextNodeValue(textNodes, "IQ500", model.clientName || "");
    replaceTextNodeValue(textNodes, "2057 Green Bay Road #1008", model.addressLine1 || "");
    replaceTextNodeValue(textNodes, "Highland Park, IL 60035", model.addressLine2 || "");
    replaceTextNodeValue(textNodes, "United States", model.country || "");
    replaceTextNodeValue(textNodes, "0160226", model.invoiceNumber || "");

    replaceTextSequence(
      textNodes,
      ["5", " ", "February", " 2026"],
      [model.invoiceDate || "", "", "", ""],
    );
    replaceTextSequence(
      textNodes,
      ["5 February", " 2026"],
      [model.dueDate || "", ""],
    );
    replaceTextSequence(
      textNodes,
      ["Monthly Retainer Plutus -", " ", "February", " 2026"],
      [model.description || "", "", "", ""],
    );

    replaceTextNodeValue(textNodes, "3,000.00", model.amountText || "", 0);
    replaceTextNodeValue(textNodes, "3,000.00", model.amountText || "", 1);
    replaceTextNodeValue(textNodes, "£3,000.00", model.amountTotalText || "");

    const serializedXml = new XMLSerializer().serializeToString(xmlDocument);
    zip.file("word/document.xml", serializedXml);

    const output = await zip.generateAsync({ type: "blob" });
    downloadBlob(
      output,
      buildInvoiceFilename(deal, "docx"),
    );
    setInvoiceBuilderStatus("DOCX invoice generated.", false);
  } catch (error) {
    setInvoiceBuilderStatus(
      error instanceof Error ? error.message : "Failed to generate DOCX invoice.",
      true,
    );
  }
}

function generateInvoicePdf() {
  const deal = getSelectedDeal();
  if (!deal) {
    setInvoiceBuilderStatus("No deal loaded for invoice generation.", true);
    return;
  }

  if (!window.jspdf || !window.jspdf.jsPDF) {
    setInvoiceBuilderStatus("PDF library is unavailable on this device.", true);
    return;
  }

  const model = buildInvoiceViewModel(deal);
  const { jsPDF } = window.jspdf;
  const doc = new jsPDF({ unit: "pt", format: "a4" });
  const left = 48;
  const right = 547;
  let y = 42;

  doc.setFont("helvetica", "normal");
  doc.setFontSize(10);
  [model.issuer.company, ...model.issuer.lines, model.issuer.email, model.issuer.phone].forEach((line) => {
    doc.text(String(line), right, y, { align: "right" });
    y += 14;
  });

  y += 8;
  doc.setDrawColor(217, 217, 217);
  doc.line(left, y, right, y);
  y += 30;

  doc.setFont("helvetica", "bold");
  doc.setFontSize(12);
  [model.clientName, model.addressLine1, model.addressLine2, model.country].filter(Boolean).forEach((line) => {
    doc.text(String(line), left, y);
    y += 16;
  });

  let headY = y - 48;
  doc.text(`INVOICE ${model.invoiceNumber}`, right, headY, { align: "right" });
  headY += 18;
  doc.text(model.invoiceDate, right, headY, { align: "right" });
  headY += 18;
  doc.setFont("helvetica", "normal");
  doc.setTextColor(107, 114, 128);
  doc.setFontSize(10);
  doc.text(`Payment due by ${model.dueDate}`, right, headY, { align: "right" });
  doc.setTextColor(17, 24, 39);

  y += 18;
  const tableTop = y;
  doc.setFillColor(0, 0, 0);
  doc.rect(left, tableTop, right - left, 24, "F");
  doc.setTextColor(255, 255, 255);
  doc.setFont("helvetica", "bold");
  doc.text("Details", left + 12, tableTop + 16);
  doc.text("Price (£)", right - 12, tableTop + 16, { align: "right" });

  const rowTop = tableTop + 24;
  doc.setFillColor(248, 249, 250);
  doc.rect(left, rowTop, right - left, 28, "F");
  doc.setTextColor(17, 24, 39);
  doc.setFont("helvetica", "normal");
  doc.text(model.description, left + 12, rowTop + 18, { maxWidth: 330 });
  doc.text(model.amountText, right - 12, rowTop + 18, { align: "right" });

  y = rowTop + 60;
  doc.text("Net Total", 400, y, { align: "right" });
  doc.text(model.amountText, right - 12, y, { align: "right" });
  y += 22;
  doc.setFont("helvetica", "bold");
  doc.text("GBP Total", 400, y, { align: "right" });
  doc.line(430, y - 14, right, y - 14);
  doc.text(model.amountTotalText, right - 12, y, { align: "right" });

  y += 38;
  doc.setFont("helvetica", "normal");
  doc.setDrawColor(217, 217, 217);
  doc.line(left, y, right, y);
  y += 24;
  doc.setFont("helvetica", "bold");
  doc.text("Payment Details", left, y);
  y += 22;
  doc.setFont("helvetica", "normal");
  [
    `Bank Account Holder Name: ${model.issuer.bankHolder}`,
    `Account number: ${model.issuer.accountNumber}`,
    `Sort code: ${model.issuer.sortCode}`,
    `IBAN: ${model.issuer.iban}`,
    `SWIFT BIC: ${model.issuer.swift}`,
  ].forEach((line) => {
    doc.text(line, left, y);
    y += 16;
  });

  doc.save(buildInvoiceFilename(deal, "pdf"));
  setInvoiceBuilderStatus("PDF invoice generated.", false);
}

function renderInvoiceHistory() {
  const panel = document.getElementById("invoice-history-panel");
  const subtitle = document.getElementById("invoice-history-subtitle");
  const list = document.getElementById("invoice-history-list");
  const loadBtn = document.getElementById("btn-load-invoice-pdf");
  if (!panel || !subtitle || !list || !loadBtn) return;

  if (!singleDealMode) {
    panel.hidden = true;
    return;
  }

  panel.hidden = false;
  const deal = getSelectedDeal();
  if (!deal) {
    subtitle.textContent = "Deal not found for this accounting route.";
    loadBtn.disabled = true;
    list.innerHTML = '<div class="invoice-item"><div class="invoice-copy">No deal loaded.</div></div>';
    return;
  }

  loadBtn.disabled = false;
  const invoices = ensureDealInvoices(deal);
  subtitle.textContent = `${invoices.length} invoice${invoices.length === 1 ? "" : "s"} linked to this deal.`;

  if (!invoices.length) {
    list.innerHTML = '<div class="invoice-item"><div class="invoice-copy">No historical invoices linked yet.</div></div>';
    return;
  }

  list.innerHTML = invoices
    .map((invoice, index) => {
      const label = invoice.name || "Invoice PDF";
      const meta = `${formatAddedAt(invoice.addedAt)}${invoice.parentPath ? ` · ${invoice.parentPath}` : ""}`;
      return `
        <div class="invoice-item">
          <div class="invoice-copy">
            <div class="invoice-title">${escapeHtml(label)}</div>
            <div class="invoice-meta">${escapeHtml(meta)}</div>
          </div>
          <div class="btn-row">
            <button class="btn" type="button" data-invoice-open-index="${index}">Open</button>
            <button class="btn danger" type="button" data-invoice-delete-index="${index}">Delete</button>
          </div>
        </div>
      `;
    })
    .join("");
}

function setInvoicePickerStatus(message) {
  const status = document.getElementById("invoice-picker-status");
  if (status) status.textContent = message || "";
}

function renderInvoicePickerPath() {
  const pathNode = document.getElementById("invoice-picker-path");
  if (!pathNode) return;

  if (!invoicePickerState.root) {
    pathNode.textContent = "Not connected to a Sharedrive folder yet.";
    return;
  }

  const names = invoicePickerState.stack.map((entry) => entry.name).filter(Boolean);
  pathNode.textContent = names.join(" / ") || invoicePickerState.root.name || "Sharedrive root";
}

function isPdfItem(item) {
  const name = String((item && item.name) || "").trim();
  const mimeType = String((item && item.mimeType) || "").trim().toLowerCase();
  return /\.pdf$/i.test(name) || mimeType === "application/pdf";
}

function renderInvoicePickerItems(items) {
  const list = document.getElementById("invoice-picker-list");
  if (!list) return;

  const sorted = (Array.isArray(items) ? items.slice() : []).sort((a, b) => {
    if (Boolean(a.isFolder) !== Boolean(b.isFolder)) return a.isFolder ? -1 : 1;
    return String(a.name || "").localeCompare(String(b.name || ""));
  });

  const filtered = sorted.filter((item) => item.isFolder || isPdfItem(item));
  invoicePickerState.items = filtered;

  if (!filtered.length) {
    list.innerHTML = '<div class="invoice-item"><div class="invoice-copy">No folders or PDF files found here.</div></div>';
    return;
  }

  list.innerHTML = filtered
    .map((item) => {
      const title = String(item.name || "(Unnamed item)");
      const meta = item.isFolder
        ? `${item.childCount != null ? `${item.childCount} item${item.childCount === 1 ? "" : "s"}` : "Folder"}`
        : `${item.mimeType || "PDF file"}${item.parentPath ? ` · ${item.parentPath}` : ""}`;
      const actionLabel = item.isFolder ? "Open folder" : "Use PDF";
      return `
        <div class="invoice-item">
          <div class="invoice-copy">
            <div class="invoice-title">${escapeHtml(title)}</div>
            <div class="invoice-meta">${escapeHtml(meta)}</div>
          </div>
          <div class="btn-row">
            <button class="btn" type="button" data-invoice-item-id="${escapeHtml(item.id)}" data-invoice-item-kind="${item.isFolder ? "folder" : "pdf"}">${actionLabel}</button>
          </div>
        </div>
      `;
    })
    .join("");
}

async function requestInvoicePickerChildren(parentItemId) {
  const shareUrl = String(invoicePickerState.shareUrl || "").trim();
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

async function loadInvoicePickerFolder(targetEntry, options = {}) {
  const { resetStack = false } = options;
  const parentItemId = targetEntry && targetEntry.id ? targetEntry.id : "";
  const data = await requestInvoicePickerChildren(parentItemId);

  const root = data.root || invoicePickerState.root || null;
  invoicePickerState.root = root;

  if (resetStack) {
    invoicePickerState.stack = [{ id: data.parentItemId || (root && root.id) || "", name: (root && root.name) || "Sharedrive root" }];
  } else if (targetEntry && targetEntry.id) {
    invoicePickerState.stack.push({ id: targetEntry.id, name: targetEntry.name || "Folder" });
  }

  renderInvoicePickerPath();
  renderInvoicePickerItems(data.items || []);
}

function closeInvoicePicker() {
  const modal = document.getElementById("invoice-picker-modal");
  if (modal) modal.hidden = true;
}

function openInvoicePicker() {
  const modal = document.getElementById("invoice-picker-modal");
  const shareUrlInput = document.getElementById("invoice-share-url-input");
  if (!modal || !shareUrlInput) return;

  invoicePickerState.shareUrl = getDefaultShareUrl();
  invoicePickerState.root = null;
  invoicePickerState.stack = [];
  invoicePickerState.items = [];
  shareUrlInput.value = invoicePickerState.shareUrl;
  setInvoicePickerStatus("");
  modal.hidden = false;

  if (!invoicePickerState.shareUrl) {
    renderInvoicePickerPath();
    renderInvoicePickerItems([]);
    setInvoicePickerStatus("Enter a Sharedrive folder URL, then click Load location.");
  }
}

function addInvoiceToDealFromPicker(item) {
  const deal = getSelectedDeal();
  if (!deal || !item) return;

  const url = String(item.webUrl || "").trim();
  if (!url) {
    setInvoiceHistoryStatus("Selected item has no PDF URL.", true);
    return;
  }

  const invoices = ensureDealInvoices(deal);
  const exists = invoices.some((entry) => normalizeValue(entry.url) === normalizeValue(url));
  if (exists) {
    setInvoiceHistoryStatus("This invoice is already linked.", false);
    return;
  }

  invoices.unshift({
    name: String(item.name || "Invoice.pdf").trim(),
    url,
    parentPath: String(item.parentPath || "").trim(),
    addedAt: new Date().toISOString(),
  });

  deal.invoices = invoices;
  markDirty(deal.id, true);
  renderTable();
  renderInvoiceHistory();
  setStatus("Invoice linked. Save all changes to persist.", false);
  setInvoiceHistoryStatus("Invoice added to this deal. Save all changes to keep it.", false);
}

function setupInvoiceHistory() {
  const loadBtn = document.getElementById("btn-load-invoice-pdf");
  const list = document.getElementById("invoice-history-list");
  const modal = document.getElementById("invoice-picker-modal");
  const closeBtn = document.getElementById("btn-close-invoice-picker");
  const loadRootBtn = document.getElementById("btn-load-invoice-root");
  const shareUrlInput = document.getElementById("invoice-share-url-input");
  const backBtn = document.getElementById("btn-invoice-go-back");
  const pickerList = document.getElementById("invoice-picker-list");

  if (loadBtn) {
    loadBtn.addEventListener("click", () => {
      if (!singleDealMode) return;
      openInvoicePicker();
    });
  }

  if (list) {
    list.addEventListener("click", (event) => {
      const openButton = event.target.closest("button[data-invoice-open-index]");
      if (openButton) {
        const index = Number(openButton.getAttribute("data-invoice-open-index"));
        const deal = getSelectedDeal();
        const invoices = ensureDealInvoices(deal);
        const selected = invoices[index];
        if (selected && selected.url) {
          window.open(selected.url, "_blank", "noopener,noreferrer");
        }
        return;
      }

      const deleteButton = event.target.closest("button[data-invoice-delete-index]");
      if (!deleteButton) return;
      const index = Number(deleteButton.getAttribute("data-invoice-delete-index"));
      const deal = getSelectedDeal();
      if (!deal) return;
      const invoices = ensureDealInvoices(deal);
      const selected = invoices[index];
      if (!selected) return;

      const confirmed = window.confirm(`Delete "${selected.name || "invoice"}" from this deal's invoice history?`);
      if (!confirmed) return;

      invoices.splice(index, 1);
      deal.invoices = invoices;
      markDirty(deal.id, true);
      renderTable();
      renderInvoiceHistory();
      setStatus("Invoice removed. Save all changes to persist.", false);
      setInvoiceHistoryStatus("Invoice removed from this deal history.", false);
    });
  }

  if (closeBtn) {
    closeBtn.addEventListener("click", closeInvoicePicker);
  }

  if (modal) {
    modal.addEventListener("click", (event) => {
      if (event.target === modal) closeInvoicePicker();
    });
  }

  if (loadRootBtn && shareUrlInput) {
    loadRootBtn.addEventListener("click", async () => {
      invoicePickerState.shareUrl = String(shareUrlInput.value || "").trim();
      if (!invoicePickerState.shareUrl) {
        setInvoicePickerStatus("Provide a Sharedrive folder URL first.");
        return;
      }

      try {
        localStorage.setItem(INVOICE_SHAREDRIVE_URL_STORAGE_KEY, invoicePickerState.shareUrl);
      } catch {
        // ignore storage failures
      }

      setInvoicePickerStatus("Loading folder...");
      try {
        invoicePickerState.root = null;
        invoicePickerState.stack = [];
        await loadInvoicePickerFolder(null, { resetStack: true });
        setInvoicePickerStatus("Choose a PDF invoice to link.");
      } catch (error) {
        setInvoicePickerStatus(error instanceof Error ? error.message : "Failed to load Sharedrive folder.");
      }
    });
  }

  if (backBtn) {
    backBtn.addEventListener("click", async () => {
      if (invoicePickerState.stack.length <= 1) {
        setInvoicePickerStatus("Already at root folder.");
        return;
      }

      invoicePickerState.stack.pop();
      const current = invoicePickerState.stack[invoicePickerState.stack.length - 1];
      setInvoicePickerStatus("Loading folder...");
      try {
        const data = await requestInvoicePickerChildren((current && current.id) || "");
        renderInvoicePickerPath();
        renderInvoicePickerItems(data.items || []);
        setInvoicePickerStatus("");
      } catch (error) {
        setInvoicePickerStatus(error instanceof Error ? error.message : "Failed to load folder.");
      }
    });
  }

  if (pickerList) {
    pickerList.addEventListener("click", async (event) => {
      const button = event.target.closest("button[data-invoice-item-id]");
      if (!button) return;

      const itemId = String(button.getAttribute("data-invoice-item-id") || "").trim();
      const kind = String(button.getAttribute("data-invoice-item-kind") || "").trim();
      if (!itemId) return;

      const selectedItem = (invoicePickerState.items || []).find((entry) => String(entry.id || "").trim() === itemId);
      if (!selectedItem) return;

      if (kind === "folder") {
        setInvoicePickerStatus("Loading folder...");
        try {
          await loadInvoicePickerFolder({ id: selectedItem.id, name: selectedItem.name }, { resetStack: false });
          setInvoicePickerStatus("");
        } catch (error) {
          setInvoicePickerStatus(error instanceof Error ? error.message : "Failed to open folder.");
        }
        return;
      }

      if (kind === "pdf") {
        addInvoiceToDealFromPicker(selectedItem);
        closeInvoicePicker();
      }
    });
  }
}

function setupInvoiceBuilder() {
  const panel = document.getElementById("invoice-builder-panel");
  const fieldMap = getInvoiceDraftFieldMap();
  const wordBtn = document.getElementById("btn-generate-invoice-word");
  const pdfBtn = document.getElementById("btn-generate-invoice-pdf");

  if (panel) {
    panel.addEventListener("input", (event) => {
      const target = event.target;
      const deal = getSelectedDeal();
      if (!deal || !target || !target.id) return;
      const draft = ensureInvoiceDraft(deal);
      const mapping = {
        "invoice-client-name": "clientName",
        "invoice-address-line-1": "addressLine1",
        "invoice-address-line-2": "addressLine2",
        "invoice-country": "country",
        "invoice-number": "invoiceNumber",
        "invoice-date": "invoiceDate",
        "invoice-due-date": "dueDate",
        "invoice-description": "description",
        "invoice-amount": "amount",
      };
      const key = mapping[target.id];
      if (!key) return;
      draft[key] = String(target.value || "").trim();
      if (key === "invoiceDate" && !draft.description) {
        draft.description = `Monthly Retainer Plutus - ${formatMonthYear(draft.invoiceDate)}`;
      }
      deal.invoiceDraft = draft;
      markDirty(deal.id, true);
      if (target.classList) target.classList.add("is-dirty");
      setStatus("Invoice details updated. Save all changes to persist.", false);
      setInvoiceBuilderStatus("Invoice details updated for this deal.", false);
    });
  }

  Object.values(fieldMap).forEach((input) => {
    if (!input) return;
    input.addEventListener("change", () => renderInvoiceBuilder());
  });

  if (wordBtn) {
    wordBtn.addEventListener("click", generateInvoiceWord);
  }
  if (pdfBtn) {
    pdfBtn.addEventListener("click", generateInvoicePdf);
  }
}

function handleInputChange(event) {
  const input = event.target;
  if (!input || !input.dataset || !input.dataset.dealId) return;

  const dealId = input.dataset.dealId;
  const field = input.dataset.field;
  const deal = dealsData.find((entry) => normalizeValue(entry.id) === normalizeValue(dealId));
  if (!deal) return;

  if (field === "retainerMonthly") {
    const value = String(input.value || "").trim();
    if (value) {
      deal.retainerMonthly = value;
      deal.Retainer = value;
    } else {
      delete deal.retainerMonthly;
      delete deal.Retainer;
    }
  } else if (field === "retainerPaymentDay") {
    const raw = String(input.value || "").trim();
    if (!raw) {
      delete deal.retainerPaymentDay;
    } else {
      const parsed = Number(raw);
      if (Number.isFinite(parsed)) {
        deal.retainerPaymentDay = Math.min(31, Math.max(1, Math.round(parsed)));
      }
    }
    input.value = getPaymentDay(deal);
  }

  markDirty(dealId, true);
  input.classList.add("is-dirty");
  renderTable();
  setStatus("You have unsaved accounting changes.", false);
}

async function saveAllChanges() {
  if (!dirtyDealIds.size) {
    setStatus("No changes to save.", false);
    return;
  }
  setStatus("Saving accounting changes...", false);
  try {
    await saveDealsData();
    dirtyDealIds.clear();
    document.querySelectorAll(".value-input.is-dirty").forEach((entry) => entry.classList.remove("is-dirty"));
    renderTable();
    renderInvoiceBuilder();
    renderInvoiceHistory();
    setStatus("Accounting changes saved.", false);
  } catch (error) {
    setStatus(error instanceof Error ? error.message : "Failed to save accounting changes.", true);
  }
}

function renderTable() {
  const body = document.getElementById("accounting-body");
  if (!body) return;

  const visibleDeals = getVisibleDeals();
  updateMetaRow(visibleDeals);

  if (!visibleDeals.length) {
    if (singleDealMode) {
      const allHref = buildPageUrl("accounting");
      const dealsHref = buildPageUrl("deals-overview");
      body.innerHTML = `<tr><td colspan="5">Deal not found. <a class="action-link" href="${dealsHref}">Back to deals</a> or <a class="action-link" href="${allHref}">open all accounting rows</a>.</td></tr>`;
      renderInvoiceBuilder();
      renderInvoiceHistory();
      return;
    }
    body.innerHTML = `<tr><td colspan="5">No deals match this filter.</td></tr>`;
    renderInvoiceBuilder();
    renderInvoiceHistory();
    return;
  }

  body.innerHTML = visibleDeals.map((deal) => {
    const dealId = String(deal.id || "");
    const accountingHref = buildPageUrl("accounting", { id: deal.id });
    const isDirty = dirtyDealIds.has(normalizeValue(dealId));
    const retainerMonthly = getRetainerMonthly(deal);
    const paymentDay = getPaymentDay(deal);
    const nextDate = computeNextExpectedDate(paymentDay);
    const ordinal = formatDayOrdinal(paymentDay);

    return `
      <tr>
        <td class="deal-cell"><a class="deal-link" href="${accountingHref}">${escapeHtml(String(deal.name || "Untitled deal"))}</a></td>
        <td>${String(deal.company || "—")}</td>
        <td>
          <input
            class="value-input ${isDirty ? "is-dirty" : ""}"
            type="text"
            value="${retainerMonthly.replace(/"/g, "&quot;")}"
            data-deal-id="${dealId.replace(/"/g, "&quot;")}"
            data-field="retainerMonthly"
            placeholder="e.g. 3000 USD"
          />
        </td>
        <td>
          <input
            class="value-input day-input ${isDirty ? "is-dirty" : ""}"
            type="number"
            min="1"
            max="31"
            value="${paymentDay}"
            data-deal-id="${dealId.replace(/"/g, "&quot;")}"
            data-field="retainerPaymentDay"
            placeholder="1-31"
          />
          <span class="hint">${ordinal || ""}</span>
        </td>
        <td class="next-date">${nextDate}</td>
      </tr>
    `;
  }).join("");
  renderInvoiceBuilder();
  renderInvoiceHistory();
}

function initializeAccountingPage() {
  selectedDealReference = getSelectedDealReference();
  singleDealMode = Boolean(selectedDealReference);
  loadDealsData();
  updatePageHeaderForMode();
  renderTable();
  setStatus(singleDealMode ? "Loaded single-deal accounting from deals data." : "No unsaved changes.", false);

  const searchInput = document.getElementById("accounting-search");
  if (searchInput) {
    searchInput.addEventListener("input", (event) => {
      if (singleDealMode) return;
      currentSearch = String(event.target && event.target.value || "");
      renderTable();
    });
  }

  const body = document.getElementById("accounting-body");
  if (body) {
    body.addEventListener("input", handleInputChange);
    body.addEventListener("change", handleInputChange);
  }

  const saveAllBtn = document.getElementById("save-all-btn");
  if (saveAllBtn) {
    saveAllBtn.addEventListener("click", saveAllChanges);
  }

  setupInvoiceHistory();
  setupInvoiceBuilder();
  renderInvoiceBuilder();
  renderInvoiceHistory();

  if (AppCore) {
    window.addEventListener("appcore:deals-updated", (event) => {
      if (event && event.detail && Array.isArray(event.detail.deals)) {
        dealsData = event.detail.deals.slice();
        updatePageHeaderForMode();
        renderTable();
        renderInvoiceBuilder();
        renderInvoiceHistory();
      }
    });
  }
}

document.addEventListener("DOMContentLoaded", initializeAccountingPage);
