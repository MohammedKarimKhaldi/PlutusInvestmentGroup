const AppCore = window.AppCore;
const DEALS_STORAGE_KEY = (AppCore && AppCore.STORAGE_KEYS && AppCore.STORAGE_KEYS.deals) || "deals_data_v1";

let dealsData = [];
const dirtyDealIds = new Set();
let currentSearch = "";
let selectedDealReference = "";
let singleDealMode = false;
let hideNoRetainerCompanies = false;
let groupByIncomingPayments = false;
const DEFAULT_CURRENCY = "GBP";
const COMMON_CURRENCIES = ["GBP", "USD", "EUR", "AED", "CHF"];
const PAYMENT_INTERVAL_OPTIONS = [1, 2, 3, 6, 12];
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
const INVOICE_STATUS_OPTIONS = [
  { value: "draft", label: "Draft" },
  { value: "prepared", label: "Prepared" },
  { value: "sent", label: "Sent" },
  { value: "part_paid", label: "Part paid" },
  { value: "paid", label: "Paid" },
  { value: "cancelled", label: "Cancelled" },
];

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

function getConfiguredAccountingAllowedEmails() {
  const page =
    window.PlutusAppConfig &&
    typeof window.PlutusAppConfig.getPage === "function"
      ? window.PlutusAppConfig.getPage("accounting")
      : null;
  return Array.isArray(page && page.allowedEmails)
    ? page.allowedEmails.map((entry) => normalizeValue(entry)).filter(Boolean)
    : [];
}

function renderAccountingAccessDenied(accessStatus) {
  const mainColumn = document.querySelector(".main-column");
  if (!mainColumn) return;

  const allowedEmails = (accessStatus && Array.isArray(accessStatus.allowedEmails) && accessStatus.allowedEmails.length)
    ? accessStatus.allowedEmails
    : getConfiguredAccountingAllowedEmails();
  const connectedPerson = accessStatus && accessStatus.person ? accessStatus.person : null;
  const connectedLabel = connectedPerson && connectedPerson.email
    ? `${connectedPerson.alias || connectedPerson.email} (${connectedPerson.email})`
    : "No permitted Microsoft account connected";

  mainColumn.innerHTML = `
    <div class="header">
      <div class="title-block">
        <h1>Accounting Access Restricted</h1>
        <p>This page is only available to approved Plutus finance users.</p>
      </div>
    </div>
    <div class="toolbar">
      <div class="toolbar-row"><strong>Signed in as:</strong> <span>${escapeHtml(connectedLabel)}</span></div>
      <div class="toolbar-row"><strong>Allowed emails:</strong> <span>${escapeHtml(allowedEmails.join(", "))}</span></div>
      <div class="toolbar-row">
        <a class="btn" href="${buildPageUrl("deals-overview")}">Back to deals</a>
        <a class="btn btn-primary" href="${buildPageUrl("sharedrive-folders")}">Switch Microsoft account</a>
      </div>
    </div>
  `;
}

async function ensureAccountingAccess() {
  const allowedEmails = getConfiguredAccountingAllowedEmails();
  if (!allowedEmails.length) return true;

  let accessStatus = null;
  if (AppCore && typeof AppCore.getPageAccessStatus === "function") {
    try {
      accessStatus = await AppCore.getPageAccessStatus("accounting");
    } catch {
      accessStatus = null;
    }
  }

  if (!accessStatus && AppCore && typeof AppCore.getCurrentConnectedPerson === "function") {
    try {
      const person = await AppCore.getCurrentConnectedPerson();
      const email = normalizeValue(person && person.email);
      accessStatus = {
        restricted: true,
        allowed: Boolean(email && allowedEmails.includes(email)),
        allowedEmails,
        person,
      };
    } catch {
      accessStatus = null;
    }
  }

  if (accessStatus && accessStatus.allowed) {
    return true;
  }

  renderAccountingAccessDenied(accessStatus || {
    restricted: true,
    allowed: false,
    allowedEmails,
    person: null,
  });
  return false;
}

function normalizeCurrencyCode(value, fallback = DEFAULT_CURRENCY) {
  const raw = String(value || "").trim().toUpperCase();
  const cleaned = raw.replace(/[^A-Z]/g, "").slice(0, 3);
  return cleaned || fallback;
}

function getDealCurrency(deal) {
  if (!deal || typeof deal !== "object") return DEFAULT_CURRENCY;
  return normalizeCurrencyCode(deal.currency, DEFAULT_CURRENCY);
}

function buildCurrencyOptionsHtml(selectedCurrency) {
  const selected = normalizeCurrencyCode(selectedCurrency, DEFAULT_CURRENCY);
  const values = COMMON_CURRENCIES.includes(selected)
    ? COMMON_CURRENCIES.slice()
    : COMMON_CURRENCIES.concat(selected);
  return values
    .map((currency) => `<option value="${currency}"${currency === selected ? " selected" : ""}>${currency}</option>`)
    .join("");
}

function getRetainerMonthly(deal) {
  if (!deal || typeof deal !== "object") return "";
  if (deal.retainerMonthly != null && String(deal.retainerMonthly).trim()) return String(deal.retainerMonthly).trim();
  if (deal.Retainer != null && String(deal.Retainer).trim()) return String(deal.Retainer).trim();
  return "";
}

function normalizePaymentIntervalMonths(value) {
  const parsed = Number(value);
  if (!Number.isFinite(parsed)) return 1;
  return Math.min(12, Math.max(1, Math.round(parsed)));
}

function getPaymentIntervalMonths(deal) {
  if (!deal || typeof deal !== "object") return 1;
  return normalizePaymentIntervalMonths(deal.retainerIntervalMonths || deal.paymentIntervalMonths || 1);
}

function buildPaymentIntervalOptionsHtml(selectedValue) {
  const selected = normalizePaymentIntervalMonths(selectedValue);
  return PAYMENT_INTERVAL_OPTIONS
    .map((value) => {
      const label =
        value === 1 ? "Monthly" :
          value === 3 ? "Quarterly" :
            value === 6 ? "Every 6 months" :
              value === 12 ? "Yearly" :
                `Every ${value} months`;
      return `<option value="${value}"${value === selected ? " selected" : ""}>${label}</option>`;
    })
    .join("");
}

function getRetainerNextPaymentDate(deal) {
  if (!deal || typeof deal !== "object") return "";
  return normalizeDateInput(deal.retainerNextPaymentDate || deal.nextPaymentDate || "");
}

function hasPositiveRetainer(deal) {
  return parseAmount(getRetainerMonthly(deal)) > 0;
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
  const nextDate = computeNextExpectedDateObject(dayValue);
  if (!nextDate) return "Not set";
  return nextDate.toLocaleDateString("en-GB", { day: "2-digit", month: "short", year: "numeric" });
}

function computeNextExpectedDateObject(dayValue) {
  const day = Number(dayValue);
  if (!Number.isFinite(day) || day < 1 || day > 31) return null;

  const now = new Date();
  const year = now.getFullYear();
  const month = now.getMonth();
  const today = now.getDate();

  const monthDays = new Date(year, month + 1, 0).getDate();
  const targetDayThisMonth = Math.min(day, monthDays);
  if (targetDayThisMonth >= today) {
    return new Date(year, month, targetDayThisMonth);
  }

  const nextMonthDays = new Date(year, month + 2, 0).getDate();
  const targetDayNextMonth = Math.min(day, nextMonthDays);
  return new Date(year, month + 1, targetDayNextMonth);
}

function addMonths(date, monthCount) {
  if (!(date instanceof Date) || Number.isNaN(date.getTime())) return null;
  const year = date.getFullYear();
  const month = date.getMonth();
  const day = date.getDate();
  const target = new Date(year, month + monthCount, 1);
  const monthDays = new Date(target.getFullYear(), target.getMonth() + 1, 0).getDate();
  target.setDate(Math.min(day, monthDays));
  return target;
}

function getNextScheduledPaymentDateObject(deal) {
  const intervalMonths = getPaymentIntervalMonths(deal);
  const explicitNextDate = getRetainerNextPaymentDate(deal);

  if (explicitNextDate) {
    const explicitDate = new Date(explicitNextDate);
    if (!Number.isNaN(explicitDate.getTime())) {
      const today = new Date();
      today.setHours(0, 0, 0, 0);
      let nextDate = new Date(explicitDate.getFullYear(), explicitDate.getMonth(), explicitDate.getDate());
      while (nextDate < today) {
        const advanced = addMonths(nextDate, intervalMonths);
        if (!advanced) break;
        nextDate = advanced;
      }
      return nextDate;
    }
  }

  if (intervalMonths > 1) {
    return null;
  }

  return computeNextExpectedDateObject(getPaymentDay(deal));
}

function formatScheduledPaymentDate(deal) {
  const nextDate = getNextScheduledPaymentDateObject(deal);
  if (!nextDate) return "Set next date";
  return nextDate.toLocaleDateString("en-GB", { day: "2-digit", month: "short", year: "numeric" });
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
    const primaryContact = getPrimaryDealContact(deal);
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
      `<div class="chip"><strong>Main contact</strong> ${textOrDash(primaryContact && (primaryContact.name || primaryContact.email))}</div>`,
      `<div class="chip"><strong>Unsaved</strong> ${dirtyDealIds.size}</div>`,
    ].join("");
    return;
  }

  const withRetainer = allDeals.filter((deal) => getRetainerMonthly(deal)).length;
  const withSchedule = allDeals.filter((deal) => Boolean(getNextScheduledPaymentDateObject(deal))).length;
  const invoiceStats = getInvoiceSummaryStats(allDeals);

  row.innerHTML = [
    `<div class="chip"><strong>${allDeals.length}</strong> deals shown</div>`,
    `<div class="chip"><strong>${withRetainer}</strong> with retainer</div>`,
    `<div class="chip"><strong>${withSchedule}</strong> with schedule</div>`,
    `<div class="chip"><strong>${invoiceStats.totalInvoices}</strong> invoices linked</div>`,
    `<div class="chip"><strong>${invoiceStats.outstanding}</strong> awaiting payment</div>`,
    `<div class="chip"><strong>${invoiceStats.overdue}</strong> overdue</div>`,
    `<div class="chip"><strong>${invoiceStats.paid}</strong> paid</div>`,
    `<div class="chip"><strong>${dirtyDealIds.size}</strong> unsaved changes</div>`,
  ].join("");
}

function updateViewControls() {
  const noRetainerBtn = document.getElementById("btn-toggle-no-retainer");
  const groupBtn = document.getElementById("btn-toggle-group-payments");
  const hintEl = document.getElementById("accounting-view-hint");

  if (noRetainerBtn) {
    noRetainerBtn.textContent = hideNoRetainerCompanies ? "Show 0 / no retainer" : "Hide 0 / no retainer";
    noRetainerBtn.classList.toggle("is-active", hideNoRetainerCompanies);
  }

  if (groupBtn) {
    groupBtn.textContent = groupByIncomingPayments ? "Ungroup incoming payments" : "Group by incoming payments";
    groupBtn.classList.toggle("is-active", groupByIncomingPayments);
  }

  if (hintEl) {
    const parts = [];
    parts.push(hideNoRetainerCompanies ? "Zero/no-retainer companies hidden." : "Showing all companies.");
    parts.push(groupByIncomingPayments ? "Grouped by next incoming payment date." : "Standard company order.");
    hintEl.textContent = parts.join(" ");
  }
}

function addCurrencyAmount(totals, currency, amount) {
  const key = normalizeCurrencyCode(currency, DEFAULT_CURRENCY);
  const numericAmount = Number(amount);
  if (!Number.isFinite(numericAmount) || numericAmount <= 0) return totals;
  totals[key] = (totals[key] || 0) + numericAmount;
  return totals;
}

function multiplyCurrencyTotals(totals, factor) {
  return Object.entries(totals || {}).reduce((accumulator, [currency, amount]) => {
    addCurrencyAmount(accumulator, currency, Number(amount) * factor);
    return accumulator;
  }, {});
}

function formatCurrencyTotalsHtml(totals, emptyLabel) {
  const entries = Object.entries(totals || {})
    .filter(([, amount]) => Number(amount) > 0)
    .sort(([left], [right]) => left.localeCompare(right));

  if (!entries.length) {
    return `<div class="analytics-empty">${escapeHtml(emptyLabel || "No values available.")}</div>`;
  }

  return `
    <div class="analytics-lines">
      ${entries.map(([currency, amount]) => `
        <div class="analytics-line">
          <span>${escapeHtml(currency)}</span>
          <strong>${escapeHtml(formatCurrencyAmount(amount, currency, true))}</strong>
        </div>
      `).join("")}
    </div>
  `;
}

function hasCurrencyTotals(totals) {
  return Object.values(totals || {}).some((amount) => Number(amount) > 0);
}

function getRetainerAmount(deal) {
  return parseAmount(getRetainerMonthly(deal));
}

function getRetainerDeals(deals) {
  return (Array.isArray(deals) ? deals : []).filter((deal) => hasPositiveRetainer(deal));
}

function buildRetainerTotalsByCurrency(deals) {
  return getRetainerDeals(deals).reduce((totals, deal) => {
    addCurrencyAmount(
      totals,
      getDealCurrency(deal),
      getRetainerAmount(deal) / getPaymentIntervalMonths(deal),
    );
    return totals;
  }, {});
}

function buildProjectionMonth(deals, monthOffset) {
  const anchor = new Date();
  const monthDate = new Date(anchor.getFullYear(), anchor.getMonth() + monthOffset, 1);
  const year = monthDate.getFullYear();
  const month = monthDate.getMonth();
  const totals = {};
  let scheduledCount = 0;

  getRetainerDeals(deals).forEach((deal) => {
    const scheduledDate = getNextScheduledPaymentDateObject(deal);
    if (!scheduledDate || Number.isNaN(scheduledDate.getTime())) return;
    if (scheduledDate.getFullYear() !== year || scheduledDate.getMonth() !== month) return;
    scheduledCount += 1;
    addCurrencyAmount(totals, getDealCurrency(deal), getRetainerAmount(deal));
  });

  return {
    label: monthDate.toLocaleDateString("en-GB", { month: "long", year: "numeric" }),
    totals,
    scheduledCount,
  };
}

function buildPaymentTimingInsights(deals) {
  const buckets = [
    { label: "Days 1-7", min: 1, max: 7 },
    { label: "Days 8-14", min: 8, max: 14 },
    { label: "Days 15-21", min: 15, max: 21 },
    { label: "Days 22-31", min: 22, max: 31 },
  ].map((bucket) => Object.assign({}, bucket, { count: 0, totals: {} }));

  getRetainerDeals(deals).forEach((deal) => {
    const nextScheduledDate = getNextScheduledPaymentDateObject(deal);
    const paymentDay = nextScheduledDate ? nextScheduledDate.getDate() : Number.NaN;
    if (!Number.isFinite(paymentDay)) return;
    const bucket = buckets.find((entry) => paymentDay >= entry.min && paymentDay <= entry.max);
    if (!bucket) return;
    bucket.count += 1;
    addCurrencyAmount(bucket.totals, getDealCurrency(deal), getRetainerAmount(deal));
  });

  return buckets;
}

function buildUnscheduledRetainerInsights(deals) {
  return getRetainerDeals(deals)
    .filter((deal) => !getNextScheduledPaymentDateObject(deal))
    .sort((left, right) => normalizeValue(left.company || left.name).localeCompare(normalizeValue(right.company || right.name)))
    .slice(0, 5);
}

function renderAnalytics(deals) {
  const cardsEl = document.getElementById("accounting-analytics-cards");
  const projectionEl = document.getElementById("accounting-analytics-projection");
  const insightsEl = document.getElementById("accounting-analytics-insights");
  const subtitleEl = document.getElementById("accounting-analytics-subtitle");
  if (!cardsEl || !projectionEl || !insightsEl || !subtitleEl) return;

  const visibleDeals = Array.isArray(deals) ? deals : [];
  const retainerDeals = getRetainerDeals(visibleDeals);
  const scheduledRetainerDeals = retainerDeals.filter((deal) => Boolean(getNextScheduledPaymentDateObject(deal)));
  const unscheduledRetainerDeals = retainerDeals.filter((deal) => !getNextScheduledPaymentDateObject(deal));
  const monthlyRecurringTotals = buildRetainerTotalsByCurrency(visibleDeals);
  const currentMonthProjection = buildProjectionMonth(visibleDeals, 0);
  const nextMonthProjection = buildProjectionMonth(visibleDeals, 1);
  const annualizedTotals = multiplyCurrencyTotals(monthlyRecurringTotals, 12);
  const invoiceStats = getInvoiceSummaryStats(visibleDeals);
  const paymentTiming = buildPaymentTimingInsights(visibleDeals);
  const unscheduledList = buildUnscheduledRetainerInsights(visibleDeals);
  const projectionMonths = [0, 1, 2].map((offset) => buildProjectionMonth(visibleDeals, offset));
  const projectionCard = projectionEl.parentElement;
  const insightCard = insightsEl.parentElement;

  subtitleEl.textContent = singleDealMode
    ? "Recurring cash and invoice analytics for this deal."
    : "Recurring cash, payment cadence, and incoming payment visibility based on the current accounting view.";

  const analyticsCards = [
    `
      <div class="analytics-card">
        <div class="analytics-card-title">Monthly Recurring Retainers</div>
        <div class="analytics-value">${retainerDeals.length}</div>
        <div class="analytics-subvalue">active retainer${retainerDeals.length === 1 ? "" : "s"}</div>
        ${formatCurrencyTotalsHtml(monthlyRecurringTotals, "No active retainers in this view.")}
      </div>
    `,
  ];

  if (currentMonthProjection.scheduledCount > 0 || hasCurrencyTotals(currentMonthProjection.totals)) {
    analyticsCards.push(`
      <div class="analytics-card">
        <div class="analytics-card-title">Scheduled This Month</div>
        <div class="analytics-value">${currentMonthProjection.scheduledCount}</div>
        <div class="analytics-subvalue">incoming payment${currentMonthProjection.scheduledCount === 1 ? "" : "s"} expected in ${escapeHtml(currentMonthProjection.label)}</div>
        ${formatCurrencyTotalsHtml(currentMonthProjection.totals, "")}
      </div>
    `);
  }

  if (nextMonthProjection.scheduledCount > 0 || hasCurrencyTotals(nextMonthProjection.totals)) {
    analyticsCards.push(`
      <div class="analytics-card">
        <div class="analytics-card-title">Scheduled Next Month</div>
        <div class="analytics-value">${nextMonthProjection.scheduledCount}</div>
        <div class="analytics-subvalue">incoming payment${nextMonthProjection.scheduledCount === 1 ? "" : "s"} expected in ${escapeHtml(nextMonthProjection.label)}</div>
        ${formatCurrencyTotalsHtml(nextMonthProjection.totals, "")}
      </div>
    `);
  }

  analyticsCards.push(`
      <div class="analytics-card">
        <div class="analytics-card-title">Annualized Run Rate</div>
        <div class="analytics-value">${retainerDeals.length}</div>
        <div class="analytics-subvalue">${unscheduledRetainerDeals.length} retainer${unscheduledRetainerDeals.length === 1 ? "" : "s"} still missing a payment schedule</div>
        ${formatCurrencyTotalsHtml(annualizedTotals, "No annualized cash run rate available.")}
      </div>
  `);

  cardsEl.innerHTML = analyticsCards.join("");

  const visibleProjectionMonths = projectionMonths.filter(
    (month) => month.scheduledCount > 0 || hasCurrencyTotals(month.totals),
  );

  if (projectionCard) {
    projectionCard.hidden = !visibleProjectionMonths.length;
  }

  projectionEl.innerHTML = visibleProjectionMonths.length
    ? visibleProjectionMonths.map((month) => `
        <div class="analytics-projection-item">
          <div class="analytics-projection-copy">
            <span>${escapeHtml(month.label)}</span>
            <small>${month.scheduledCount} scheduled payment${month.scheduledCount === 1 ? "" : "s"}</small>
          </div>
          <div>${formatCurrencyTotalsHtml(month.totals, "")}</div>
        </div>
      `).join("")
    : "";

  if (insightCard) {
    insightCard.hidden = false;
  }

  const visiblePaymentTiming = paymentTiming.filter(
    (bucket) => bucket.count > 0 || hasCurrencyTotals(bucket.totals),
  );

  const paymentTimingHtml = visiblePaymentTiming.length
    ? visiblePaymentTiming.map((bucket) => `
        <div class="analytics-insight-item">
          <div class="analytics-insight-copy">
            <span>${escapeHtml(bucket.label)}</span>
            <small>${bucket.count} scheduled retainer${bucket.count === 1 ? "" : "s"}</small>
          </div>
          <div>${formatCurrencyTotalsHtml(bucket.totals, "")}</div>
        </div>
      `).join("")
    : '<div class="analytics-empty">Add a payment day or next scheduled date to see timing distribution.</div>';

  const unscheduledHtml = unscheduledList.length
    ? unscheduledList.map((deal) => `
        <div class="analytics-insight-item">
          <div class="analytics-insight-copy">
            <span>${escapeHtml(String(deal.company || deal.name || "Deal"))}</span>
            <small>No payment schedule assigned yet</small>
          </div>
          <strong>${escapeHtml(formatCurrencyAmount(getRetainerAmount(deal), getDealCurrency(deal), true))}</strong>
        </div>
      `).join("")
    : '<div class="analytics-empty">All active retainers in this view have a payment schedule.</div>';

  insightsEl.innerHTML = `
    <div class="analytics-card-title">Payment Timing</div>
    ${paymentTimingHtml}
    <div class="analytics-card-title" style="margin-top:12px;">Unscheduled Retainers</div>
    ${unscheduledHtml}
    <div class="analytics-card-title" style="margin-top:12px;">Invoice Collection</div>
    <div class="analytics-lines">
      <div class="analytics-line"><span>Outstanding</span><strong>${invoiceStats.outstanding}</strong></div>
      <div class="analytics-line"><span>Overdue</span><strong>${invoiceStats.overdue}</strong></div>
      <div class="analytics-line"><span>Paid</span><strong>${invoiceStats.paid}</strong></div>
      <div class="analytics-line"><span>Total invoices</span><strong>${invoiceStats.totalInvoices}</strong></div>
    </div>
  `;
}

function getVisibleDeals() {
  if (singleDealMode) {
    const selected = findDealByReference(selectedDealReference);
    return selected ? [selected] : [];
  }

  const query = normalizeValue(currentSearch);
  const source = Array.isArray(dealsData) ? dealsData.slice() : [];
  const filtered = source.filter((deal) => {
    const contacts = normalizeDealContacts(deal);
    const haystack = [
      deal && deal.name,
      deal && deal.company,
      deal && deal.id,
      deal && deal.seniorOwner,
      deal && deal.juniorOwner,
      ...contacts.map((entry) => `${entry.name} ${entry.title} ${entry.email}`),
    ].map((value) => normalizeValue(value)).join(" ");
    return !query || haystack.includes(query);
  });

  const retainersFiltered = hideNoRetainerCompanies
    ? filtered.filter((deal) => hasPositiveRetainer(deal))
    : filtered;

  retainersFiltered.sort((a, b) => normalizeValue(a.company || a.name).localeCompare(normalizeValue(b.company || b.name)));
  return retainersFiltered;
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
    if (title) title.textContent = "Retainers & Schedules";
    if (subtitle) subtitle.textContent = "Manage retainers, payment cadence, and expected payment schedules for each deal.";
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

function normalizeDateInput(value) {
  const raw = String(value || "").trim();
  if (!raw) return "";
  return /^\d{4}-\d{2}-\d{2}$/.test(raw) ? raw : "";
}

function normalizeInvoiceStatus(value) {
  const raw = String(value || "").trim().toLowerCase();
  return INVOICE_STATUS_OPTIONS.some((option) => option.value === raw) ? raw : "draft";
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
          status: "draft",
          sentDate: "",
          dueDate: "",
          paidDate: "",
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
        status: normalizeInvoiceStatus(entry.status),
        sentDate: normalizeDateInput(entry.sentDate || entry.invoiceDate || entry.sentAt),
        dueDate: normalizeDateInput(entry.dueDate),
        paidDate: normalizeDateInput(entry.paidDate || entry.payDate || entry.paidAt),
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
    description: String(source.description || `Retainer Plutus - ${currentMonthLabel}`).trim(),
    currency: normalizeCurrencyCode(source.currency || getDealCurrency(deal), getDealCurrency(deal)),
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

function setDealContactsStatus(message, isError) {
  const node = document.getElementById("deal-contacts-status");
  if (!node) return;
  node.textContent = message || "";
  node.style.color = isError ? "#ef4444" : "var(--text-soft)";
}

function buildAccountingContactEditorRow(contact, index) {
  return `
    <div class="invoice-item">
      <div class="contact-meta-card">
        <div class="contact-editor-grid">
          <label class="invoice-field">
            <span>Name</span>
            <input data-contact-index="${index}" data-contact-field="name" type="text" value="${escapeHtml(contact.name || "")}" />
          </label>
          <label class="invoice-field">
            <span>Title</span>
            <input data-contact-index="${index}" data-contact-field="title" type="text" value="${escapeHtml(contact.title || "")}" />
          </label>
          <label class="invoice-field">
            <span>Email</span>
            <input data-contact-index="${index}" data-contact-field="email" type="email" value="${escapeHtml(contact.email || "")}" />
          </label>
          <div class="contact-editor-actions">
            <label class="contact-primary-toggle">
              <input type="radio" name="accounting-contact-primary" data-contact-index="${index}" data-contact-field="isPrimary" ${contact.isPrimary ? "checked" : ""} />
              <span>Main contact</span>
            </label>
            <button class="btn" type="button" data-remove-contact-index="${index}">Remove</button>
          </div>
        </div>
      </div>
    </div>
  `;
}

function renderDealContactsPanel() {
  const panel = document.getElementById("deal-contacts-panel");
  const list = document.getElementById("deal-contacts-editor-list");
  const subtitle = document.getElementById("deal-contacts-subtitle");
  if (!panel || !list || !subtitle) return;

  if (!singleDealMode) {
    panel.hidden = true;
    return;
  }

  const deal = getSelectedDeal();
  panel.hidden = false;
  if (!deal) {
    subtitle.textContent = "Deal not found for contact tracking.";
    list.innerHTML = '<div class="invoice-item"><div class="contact-subline">No deal loaded.</div></div>';
    return;
  }

  const contacts = normalizeDealContacts(deal, { keepEmpty: true });
  const primary = getPrimaryDealContact(deal);
  subtitle.textContent = primary && primary.email
    ? `Main contact: ${primary.name || primary.email}${primary.title ? ` · ${primary.title}` : ""}`
    : `${contacts.length} contact${contacts.length === 1 ? "" : "s"} linked to this deal.`;

  list.innerHTML = contacts.length
    ? contacts.map((contact, index) => buildAccountingContactEditorRow(contact, index)).join("")
    : '<div class="invoice-item"><div class="contact-subline">No contacts saved for this deal yet.</div></div>';
}

function updateDealContactField(index, field, nextValue, isChecked) {
  const deal = getSelectedDeal();
  if (!deal) return;
  const contacts = normalizeDealContacts(deal, { keepEmpty: true });
  const contact = contacts[index];
  if (!contact) return;

  if (field === "isPrimary") {
    if (!isChecked) return;
    contacts.forEach((entry, entryIndex) => {
      entry.isPrimary = entryIndex === index;
    });
  } else if (["name", "title", "email"].includes(field)) {
    contact[field] = String(nextValue || "").trim();
  } else {
    return;
  }

  deal.contacts = contacts.filter((entry) => entry.name || entry.title || entry.email);
  if (deal.contacts.length && !deal.contacts.some((entry) => entry.isPrimary)) {
    deal.contacts[0].isPrimary = true;
  }
  syncLegacyPrimaryContactFields(deal);
  markDirty(deal.id, true);
  renderDealContactsPanel();
  renderTable();
  setStatus("Deal contacts updated. Save all changes to sync online.", false);
  setDealContactsStatus("Deal contacts updated for this deal.", false);
}

function formatAddedAt(value) {
  const raw = String(value || "").trim();
  if (!raw) return "Date unknown";
  const dt = new Date(raw);
  if (Number.isNaN(dt.getTime())) return raw;
  return dt.toLocaleDateString("en-GB", { day: "2-digit", month: "short", year: "numeric" });
}

function formatStatusLabel(value) {
  const option = INVOICE_STATUS_OPTIONS.find((entry) => entry.value === normalizeInvoiceStatus(value));
  return option ? option.label : "Draft";
}

function formatShortDate(value) {
  const raw = normalizeDateInput(value);
  if (!raw) return "";
  const dt = new Date(raw);
  if (Number.isNaN(dt.getTime())) return raw;
  return dt.toLocaleDateString("en-GB", { day: "2-digit", month: "short", year: "numeric" });
}

function getLatestInvoiceRecord(deal) {
  const invoices = ensureDealInvoices(deal);
  if (!invoices.length) return null;

  return invoices
    .slice()
    .sort((a, b) => {
      const aDate = Date.parse(a.paidDate || a.dueDate || a.sentDate || a.addedAt || "");
      const bDate = Date.parse(b.paidDate || b.dueDate || b.sentDate || b.addedAt || "");
      return (Number.isFinite(bDate) ? bDate : 0) - (Number.isFinite(aDate) ? aDate : 0);
    })[0] || invoices[0] || null;
}

function isInvoiceOverdue(invoice) {
  if (!invoice || normalizeInvoiceStatus(invoice.status) === "paid" || normalizeInvoiceStatus(invoice.status) === "cancelled") {
    return false;
  }
  if (invoice.paidDate) return false;
  const dueDate = normalizeDateInput(invoice.dueDate);
  if (!dueDate) return false;
  const today = new Date().toISOString().slice(0, 10);
  return dueDate < today;
}

function getInvoiceDisplayStatus(invoice) {
  if (isInvoiceOverdue(invoice)) return "overdue";
  return normalizeInvoiceStatus(invoice && invoice.status);
}

function getInvoiceSummaryStats(deals) {
  const summary = {
    totalInvoices: 0,
    outstanding: 0,
    paid: 0,
    sent: 0,
    overdue: 0,
  };

  (Array.isArray(deals) ? deals : []).forEach((deal) => {
    ensureDealInvoices(deal).forEach((invoice) => {
      summary.totalInvoices += 1;
      const status = normalizeInvoiceStatus(invoice.status);
      if (isInvoiceOverdue(invoice)) summary.overdue += 1;
      if (["prepared", "sent", "part_paid"].includes(status)) summary.outstanding += 1;
      if (status === "paid") summary.paid += 1;
      if (status === "sent" || status === "part_paid") summary.sent += 1;
    });
  });

  return summary;
}

function buildInvoiceBadge(status) {
  const normalized = status === "overdue" ? "overdue" : normalizeInvoiceStatus(status);
  const label = normalized === "overdue" ? "Overdue" : formatStatusLabel(normalized);
  return `<span class="invoice-badge invoice-badge-${normalized}">${escapeHtml(label)}</span>`;
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

function formatCurrencyAmount(value, currency, withSymbol) {
  const amount = parseAmount(value);
  const formatted = amount.toLocaleString("en-GB", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  const code = normalizeCurrencyCode(currency, DEFAULT_CURRENCY);
  if (!withSymbol) return formatted;
  try {
    return new Intl.NumberFormat("en-GB", {
      style: "currency",
      currency: code,
      minimumFractionDigits: 2,
      maximumFractionDigits: 2,
    }).format(amount);
  } catch {
    return `${code} ${formatted}`;
  }
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
    currency: document.getElementById("invoice-currency"),
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
  if (fieldMap.currency) {
    fieldMap.currency.innerHTML = buildCurrencyOptionsHtml(draft.currency || getDealCurrency(deal));
  }
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
    description: draft.description || `Retainer Plutus - ${formatMonthYear(draft.invoiceDate)}`,
    currency: normalizeCurrencyCode(draft.currency || getDealCurrency(deal), getDealCurrency(deal)),
    amountText: formatCurrencyAmount(draft.amount, draft.currency || getDealCurrency(deal), false),
    amountTotalText: formatCurrencyAmount(draft.amount, draft.currency || getDealCurrency(deal), true),
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
      <tr><th>Details</th><th>Price (${escapeHtml(model.currency)})</th></tr>
    </thead>
    <tbody>
      <tr><td>${escapeHtml(model.description)}</td><td>${escapeHtml(model.amountText)}</td></tr>
    </tbody>
  </table>
  <table class="totals-table">
    <tr><td>Net Total</td><td>${escapeHtml(model.amountText)}</td></tr>
    <tr class="grand"><td>${escapeHtml(model.currency)} Total</td><td>${escapeHtml(model.amountTotalText)}</td></tr>
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
  doc.text(`Price (${model.currency})`, right - 12, tableTop + 16, { align: "right" });

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
  doc.text(`${model.currency} Total`, 400, y, { align: "right" });
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
      const status = normalizeInvoiceStatus(invoice.status);
      const displayStatus = getInvoiceDisplayStatus(invoice);
      const detailParts = [
        `Added ${formatAddedAt(invoice.addedAt)}`,
        `Status ${displayStatus === "overdue" ? "Overdue" : formatStatusLabel(status)}`,
      ];
      if (invoice.sentDate) detailParts.push(`Sent ${formatShortDate(invoice.sentDate)}`);
      if (invoice.dueDate) detailParts.push(`Due ${formatShortDate(invoice.dueDate)}`);
      if (invoice.paidDate) detailParts.push(`Paid ${formatShortDate(invoice.paidDate)}`);
      if (invoice.parentPath) detailParts.push(invoice.parentPath);
      const meta = detailParts.join(" · ");
      const statusOptions = INVOICE_STATUS_OPTIONS
        .map((option) => `<option value="${option.value}"${option.value === status ? " selected" : ""}>${escapeHtml(option.label)}</option>`)
        .join("");
      return `
        <div class="invoice-item">
          <div class="invoice-copy">
            <div class="invoice-title">${escapeHtml(label)}</div>
            <div class="invoice-meta-row">
              ${buildInvoiceBadge(displayStatus)}
              <div class="invoice-meta">${escapeHtml(meta)}</div>
            </div>
            <div class="invoice-status-grid">
              <label class="invoice-field">
                <span>Status</span>
                <select data-invoice-index="${index}" data-invoice-field="status">
                  ${statusOptions}
                </select>
              </label>
              <label class="invoice-field">
                <span>Sent date</span>
                <input data-invoice-index="${index}" data-invoice-field="sentDate" type="date" value="${escapeHtml(invoice.sentDate || "")}" />
              </label>
              <label class="invoice-field">
                <span>Due date</span>
                <input data-invoice-index="${index}" data-invoice-field="dueDate" type="date" value="${escapeHtml(invoice.dueDate || "")}" />
              </label>
              <label class="invoice-field">
                <span>Paid date</span>
                <input data-invoice-index="${index}" data-invoice-field="paidDate" type="date" value="${escapeHtml(invoice.paidDate || "")}" />
              </label>
            </div>
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

function updateInvoiceTrackingField(index, field, nextValue) {
  const deal = getSelectedDeal();
  if (!deal) return;

  const invoices = ensureDealInvoices(deal);
  const invoice = invoices[index];
  if (!invoice) return;

  const today = new Date().toISOString().slice(0, 10);
  if (field === "status") {
    invoice.status = normalizeInvoiceStatus(nextValue);
    if (invoice.status === "sent" && !invoice.sentDate) {
      invoice.sentDate = today;
    }
    if (invoice.status === "paid") {
      if (!invoice.sentDate) invoice.sentDate = today;
      if (!invoice.paidDate) invoice.paidDate = today;
    }
  } else if (field === "sentDate") {
    invoice.sentDate = normalizeDateInput(nextValue);
    if (invoice.sentDate && invoice.status === "draft") {
      invoice.status = "sent";
    }
  } else if (field === "dueDate") {
    invoice.dueDate = normalizeDateInput(nextValue);
  } else if (field === "paidDate") {
    invoice.paidDate = normalizeDateInput(nextValue);
    if (invoice.paidDate) {
      invoice.status = "paid";
      if (!invoice.sentDate) invoice.sentDate = invoice.paidDate;
    }
  } else {
    return;
  }

  deal.invoices = invoices;
  markDirty(deal.id, true);
  renderInvoiceHistory();
  setStatus("Invoice tracking updated. Save all changes to sync online.", false);
  setInvoiceHistoryStatus("Invoice tracking updated for this deal.", false);
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
  const draft = ensureInvoiceDraft(deal) || {};
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
    status: "draft",
    sentDate: normalizeDateInput(draft.invoiceDate),
    dueDate: normalizeDateInput(draft.dueDate),
    paidDate: "",
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
    list.addEventListener("change", (event) => {
      const target = event.target;
      if (!target || !target.dataset) return;
      const index = Number(target.dataset.invoiceIndex);
      const field = String(target.dataset.invoiceField || "").trim();
      if (!Number.isFinite(index) || !field) return;
      updateInvoiceTrackingField(index, field, target.value);
    });

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
        "invoice-currency": "currency",
        "invoice-amount": "amount",
      };
      const key = mapping[target.id];
      if (!key) return;
      draft[key] = key === "currency"
        ? normalizeCurrencyCode(target.value, getDealCurrency(deal))
        : String(target.value || "").trim();
      if (key === "invoiceDate" && !draft.description) {
        draft.description = `Retainer Plutus - ${formatMonthYear(draft.invoiceDate)}`;
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

function setupDealContactsPanel() {
  const addBtn = document.getElementById("btn-add-accounting-contact");
  const list = document.getElementById("deal-contacts-editor-list");

  if (addBtn) {
    addBtn.addEventListener("click", () => {
      const deal = getSelectedDeal();
      if (!deal) return;
      const contacts = normalizeDealContacts(deal, { keepEmpty: true });
      contacts.push({
        name: "",
        title: "",
        email: "",
        isPrimary: !contacts.length,
      });
      deal.contacts = contacts;
      renderDealContactsPanel();
      setDealContactsStatus("Contact row added. Save all changes to persist.", false);
    });
  }

  if (list) {
    list.addEventListener("input", (event) => {
      const target = event.target;
      if (!target || !target.dataset) return;
      const index = Number(target.dataset.contactIndex);
      const field = String(target.dataset.contactField || "").trim();
      if (!Number.isFinite(index) || !field || field === "isPrimary") return;
      updateDealContactField(index, field, target.value, target.checked);
    });

    list.addEventListener("change", (event) => {
      const target = event.target;
      if (!target || !target.dataset) return;
      const index = Number(target.dataset.contactIndex);
      const field = String(target.dataset.contactField || "").trim();
      if (!Number.isFinite(index) || !field) return;
      updateDealContactField(index, field, target.value, target.checked);
    });

    list.addEventListener("click", (event) => {
      const removeButton = event.target.closest("button[data-remove-contact-index]");
      if (!removeButton) return;
      const index = Number(removeButton.getAttribute("data-remove-contact-index"));
      const deal = getSelectedDeal();
      if (!deal) return;
      const contacts = normalizeDealContacts(deal, { keepEmpty: true });
      if (!Number.isFinite(index) || !contacts[index]) return;
      contacts.splice(index, 1);
      if (contacts.length && !contacts.some((entry) => entry.isPrimary)) {
        contacts[0].isPrimary = true;
      }
      deal.contacts = contacts;
      syncLegacyPrimaryContactFields(deal);
      markDirty(deal.id, true);
      renderDealContactsPanel();
      renderTable();
      setStatus("Deal contact removed. Save all changes to sync online.", false);
      setDealContactsStatus("Deal contact removed.", false);
    });
  }
}

function handleInputChange(event) {
  const input = event.target;
  if (!input || !input.dataset || !input.dataset.dealId) return;

  const dealId = input.dataset.dealId;
  const field = input.dataset.field;
  const isCommittedChange = event.type === "change";
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
    if (isCommittedChange) {
      input.value = getPaymentDay(deal);
    }
  } else if (field === "retainerIntervalMonths") {
    const intervalMonths = normalizePaymentIntervalMonths(input.value);
    deal.retainerIntervalMonths = intervalMonths;
    deal.paymentIntervalMonths = intervalMonths;
    if (intervalMonths === 1 && !getRetainerNextPaymentDate(deal)) {
      delete deal.retainerNextPaymentDate;
      delete deal.nextPaymentDate;
    }
  } else if (field === "retainerNextPaymentDate") {
    const normalizedDate = normalizeDateInput(input.value);
    if (normalizedDate) {
      deal.retainerNextPaymentDate = normalizedDate;
      deal.nextPaymentDate = normalizedDate;
    } else {
      delete deal.retainerNextPaymentDate;
      delete deal.nextPaymentDate;
    }
  } else if (field === "currency") {
    deal.currency = normalizeCurrencyCode(input.value, getDealCurrency(deal));
    const draft = ensureInvoiceDraft(deal);
    if (draft) {
      draft.currency = deal.currency;
      deal.invoiceDraft = draft;
    }
  }

  markDirty(dealId, true);
  input.classList.add("is-dirty");
  const visibleDeals = getVisibleDeals();
  updateMetaRow(visibleDeals);
  renderAnalytics(visibleDeals);
  if (isCommittedChange && ["retainerPaymentDay", "retainerIntervalMonths", "retainerNextPaymentDate"].includes(field)) {
    renderTable();
    return;
  }
  if (field === "currency") {
    renderTable();
    renderInvoiceBuilder();
    setStatus("You have unsaved accounting changes.", false);
    return;
  }
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
    setStatus("Accounting changes saved to the shared online deals data.", false);
  } catch (error) {
    setStatus(error instanceof Error ? error.message : "Failed to save accounting changes.", true);
  }
}

function getIncomingPaymentGroupDetails(deal) {
  const nextDate = getNextScheduledPaymentDateObject(deal);
  if (!nextDate) {
    return {
      key: "no-payment-date",
      label: "No incoming payment scheduled",
      sortValue: Number.MAX_SAFE_INTEGER,
    };
  }

  return {
    key: nextDate.toISOString().slice(0, 10),
    label: `Incoming ${nextDate.toLocaleDateString("en-GB", { day: "2-digit", month: "short", year: "numeric" })}`,
    sortValue: nextDate.getTime(),
  };
}

function renderAccountingRow(deal) {
  const dealId = String(deal.id || "");
  const accountingHref = buildPageUrl("accounting", { id: deal.id });
  const isDirty = dirtyDealIds.has(normalizeValue(dealId));
  const retainerMonthly = getRetainerMonthly(deal);
  const currency = getDealCurrency(deal);
  const cadenceMonths = getPaymentIntervalMonths(deal);
  const paymentDay = getPaymentDay(deal);
  const nextDate = formatScheduledPaymentDate(deal);
  const nextPaymentDate = getRetainerNextPaymentDate(deal);
  const ordinal = formatDayOrdinal(paymentDay);
  const primaryContact = getPrimaryDealContact(deal);
  const invoices = ensureDealInvoices(deal);
  const latestInvoice = getLatestInvoiceRecord(deal);
  const latestStatus = latestInvoice ? buildInvoiceBadge(getInvoiceDisplayStatus(latestInvoice)) : '<span class="hint">No invoice</span>';
  const latestSentDate = latestInvoice && latestInvoice.sentDate ? formatShortDate(latestInvoice.sentDate) : "—";
  const latestDueDate = latestInvoice && latestInvoice.dueDate ? formatShortDate(latestInvoice.dueDate) : "—";
  const latestPaidDate = latestInvoice && latestInvoice.paidDate ? formatShortDate(latestInvoice.paidDate) : "—";

  return `
    <tr>
      <td class="deal-cell"><a class="deal-link" href="${accountingHref}">${escapeHtml(String(deal.name || "Untitled deal"))}</a></td>
      <td>${String(deal.company || "—")}</td>
      <td>
        ${primaryContact
          ? `<div class="contact-stack"><span class="contact-name">${escapeHtml(primaryContact.name || primaryContact.email || "Contact")}</span><span class="contact-subline">${escapeHtml([primaryContact.title, primaryContact.email].filter(Boolean).join(" · ") || "No email")}</span></div>`
          : '<span class="hint">No contact</span>'}
      </td>
      <td>
        <input
          class="value-input ${isDirty ? "is-dirty" : ""}"
          type="text"
          value="${retainerMonthly.replace(/"/g, "&quot;")}"
          data-deal-id="${dealId.replace(/"/g, "&quot;")}"
          data-field="retainerMonthly"
          placeholder="e.g. 5000"
        />
      </td>
      <td>
        <select
          class="value-input ${isDirty ? "is-dirty" : ""}"
          data-deal-id="${dealId.replace(/"/g, "&quot;")}"
          data-field="currency"
        >
          ${buildCurrencyOptionsHtml(currency)}
        </select>
      </td>
      <td>
        <select
          class="value-input ${isDirty ? "is-dirty" : ""}"
          data-deal-id="${dealId.replace(/"/g, "&quot;")}"
          data-field="retainerIntervalMonths"
        >
          ${buildPaymentIntervalOptionsHtml(cadenceMonths)}
        </select>
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
      <td>
        <input
          class="value-input ${isDirty ? "is-dirty" : ""}"
          type="date"
          value="${nextPaymentDate}"
          data-deal-id="${dealId.replace(/"/g, "&quot;")}"
          data-field="retainerNextPaymentDate"
        />
      </td>
      <td class="next-date">${nextDate}</td>
      <td>
        <div class="invoice-cell">
          <strong>${invoices.length}</strong>
          <span class="hint">${invoices.length === 1 ? "invoice" : "invoices"}</span>
        </div>
      </td>
      <td>${latestStatus}</td>
      <td class="next-date">${latestSentDate}</td>
      <td class="next-date">${latestDueDate}</td>
      <td class="next-date">${latestPaidDate}</td>
    </tr>
  `;
}

function renderGroupedAccountingRows(deals) {
  const sorted = (Array.isArray(deals) ? deals.slice() : []).sort((left, right) => {
    const leftGroup = getIncomingPaymentGroupDetails(left);
    const rightGroup = getIncomingPaymentGroupDetails(right);
    if (leftGroup.sortValue !== rightGroup.sortValue) {
      return leftGroup.sortValue - rightGroup.sortValue;
    }
    return normalizeValue(left.company || left.name).localeCompare(normalizeValue(right.company || right.name));
  });

  const groups = [];
  sorted.forEach((deal) => {
    const details = getIncomingPaymentGroupDetails(deal);
    const existing = groups[groups.length - 1];
    if (!existing || existing.key !== details.key) {
      groups.push({
        key: details.key,
        label: details.label,
        deals: [deal],
      });
      return;
    }
    existing.deals.push(deal);
  });

  return groups.map((group) => `
    <tr class="group-row">
      <td colspan="14">
        <div class="group-row-copy">
          <span>${escapeHtml(group.label)}</span>
          <span class="group-row-meta">${group.deals.length} compan${group.deals.length === 1 ? "y" : "ies"}</span>
        </div>
      </td>
    </tr>
    ${group.deals.map((deal) => renderAccountingRow(deal)).join("")}
  `).join("");
}

function renderTable() {
  const body = document.getElementById("accounting-body");
  if (!body) return;

  const visibleDeals = getVisibleDeals();
  updateMetaRow(visibleDeals);
  updateViewControls();
  renderAnalytics(visibleDeals);

  if (!visibleDeals.length) {
    if (singleDealMode) {
      const allHref = buildPageUrl("accounting");
      const dealsHref = buildPageUrl("deals-overview");
      body.innerHTML = `<tr><td colspan="14">Deal not found. <a class="action-link" href="${dealsHref}">Back to deals</a> or <a class="action-link" href="${allHref}">open all accounting rows</a>.</td></tr>`;
      renderInvoiceBuilder();
      renderDealContactsPanel();
      renderInvoiceHistory();
      return;
    }
    body.innerHTML = `<tr><td colspan="14">No deals match this filter.</td></tr>`;
    renderInvoiceBuilder();
    renderDealContactsPanel();
    renderInvoiceHistory();
    return;
  }

  body.innerHTML = groupByIncomingPayments
    ? renderGroupedAccountingRows(visibleDeals)
    : visibleDeals.map((deal) => renderAccountingRow(deal)).join("");
  renderInvoiceBuilder();
  renderDealContactsPanel();
  renderInvoiceHistory();
}

async function initializeAccountingPage() {
  const hasAccess = await ensureAccountingAccess();
  if (!hasAccess) return;

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

  const toggleNoRetainerBtn = document.getElementById("btn-toggle-no-retainer");
  if (toggleNoRetainerBtn) {
    toggleNoRetainerBtn.addEventListener("click", () => {
      hideNoRetainerCompanies = !hideNoRetainerCompanies;
      renderTable();
    });
  }

  const toggleGroupPaymentsBtn = document.getElementById("btn-toggle-group-payments");
  if (toggleGroupPaymentsBtn) {
    toggleGroupPaymentsBtn.addEventListener("click", () => {
      groupByIncomingPayments = !groupByIncomingPayments;
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
  setupDealContactsPanel();
  renderInvoiceBuilder();
  renderDealContactsPanel();
  renderInvoiceHistory();
  updateViewControls();

  if (AppCore) {
    window.addEventListener("appcore:graph-session-updated", () => {
      window.location.reload();
    });
    window.addEventListener("appcore:deals-updated", (event) => {
      if (event && event.detail && Array.isArray(event.detail.deals)) {
        dealsData = event.detail.deals.slice();
        updatePageHeaderForMode();
        renderTable();
        renderInvoiceBuilder();
        renderDealContactsPanel();
        renderInvoiceHistory();
      }
    });
  }
}

document.addEventListener("DOMContentLoaded", initializeAccountingPage);
