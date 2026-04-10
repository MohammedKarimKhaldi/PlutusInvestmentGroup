const AppCore = window.AppCore;
const DAYS_OFF_STORAGE_KEY = (AppCore && AppCore.STORAGE_KEYS && AppCore.STORAGE_KEYS.daysOff) || "days_off_v1";
const DAY_MS = 24 * 60 * 60 * 1000;
const DEFAULT_TYPE = "annual leave";
const WEEKDAY_OFFSET = 1;
const CALENDAR_WEEKS = 6;
const ENTRY_TYPE_LABELS = {
  "annual leave": "Annual leave",
  "public holiday": "Public holiday",
  "personal day": "Personal day",
  "sick leave": "Sick leave",
  travel: "Travel",
  other: "Other",
};

let daysOffEntries = [];
let dealsData = [];
let tasksData = [];
let editingEntryId = "";
let currentSearch = "";
let currentStatusFilter = "all";
let currentTypeFilter = "all";
let currentCalendarMonth = startOfMonth(new Date());
let selectedDateKey = toDateKey(new Date());

function cloneArray(value) {
  return Array.isArray(value) ? JSON.parse(JSON.stringify(value)) : [];
}

function normalizeValue(value) {
  if (AppCore && typeof AppCore.normalizeValue === "function") {
    return AppCore.normalizeValue(value);
  }
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

function titleCase(value) {
  const text = String(value || "").trim();
  if (!text) return "";
  return text.replace(/\b\w/g, (letter) => letter.toUpperCase());
}

function parseDateOnly(value) {
  const raw = String(value || "").trim();
  if (!/^\d{4}-\d{2}-\d{2}$/.test(raw)) return null;

  const [year, month, day] = raw.split("-").map(Number);
  const date = new Date(year, month - 1, day);
  if (
    Number.isNaN(date.getTime()) ||
    date.getFullYear() !== year ||
    date.getMonth() !== month - 1 ||
    date.getDate() !== day
  ) {
    return null;
  }

  date.setHours(0, 0, 0, 0);
  return date;
}

function toDateKey(value) {
  const date = value instanceof Date ? new Date(value.getTime()) : parseDateOnly(value);
  if (!date || Number.isNaN(date.getTime())) return "";
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, "0");
  const day = String(date.getDate()).padStart(2, "0");
  return `${year}-${month}-${day}`;
}

function startOfMonth(value) {
  const date = value instanceof Date ? new Date(value.getTime()) : new Date();
  date.setHours(0, 0, 0, 0);
  return new Date(date.getFullYear(), date.getMonth(), 1);
}

function addDays(date, amount) {
  const next = new Date(date.getTime());
  next.setDate(next.getDate() + amount);
  next.setHours(0, 0, 0, 0);
  return next;
}

function isWeekendDate(value) {
  const date = value instanceof Date ? value : parseDateOnly(value);
  if (!date || Number.isNaN(date.getTime())) return false;
  const day = date.getDay();
  return day === 0 || day === 6;
}

function shiftMonth(date, amount) {
  return new Date(date.getFullYear(), date.getMonth() + amount, 1);
}

function formatDisplayDate(value) {
  const date = parseDateOnly(value);
  if (!date) return String(value || "");
  return date.toLocaleDateString("en-GB", {
    day: "2-digit",
    month: "short",
    year: "numeric",
  });
}

function formatMonthYear(date) {
  return date.toLocaleDateString("en-GB", {
    month: "long",
    year: "numeric",
  });
}

function formatDateRange(entry) {
  const start = formatDisplayDate(entry.startDate);
  const end = formatDisplayDate(entry.endDate || entry.startDate);
  if (!start) return "No dates";
  if (!end || end === start) return start;
  return `${start} - ${end}`;
}

function calculateCalendarDaysCount(entry) {
  const start = parseDateOnly(entry.startDate);
  const end = parseDateOnly(entry.endDate || entry.startDate);
  if (!start || !end) return 0;
  return Math.floor((end.getTime() - start.getTime()) / DAY_MS) + 1;
}

function calculateDaysCount(entry) {
  const start = parseDateOnly(entry.startDate);
  const end = parseDateOnly(entry.endDate || entry.startDate);
  if (!start || !end) return 0;

  let count = 0;
  for (let cursor = new Date(start.getTime()); cursor <= end; cursor = addDays(cursor, 1)) {
    if (!isWeekendDate(cursor)) count += 1;
  }
  return count;
}

function getEntryState(entry, today = new Date()) {
  const currentDay = new Date(today.getTime());
  currentDay.setHours(0, 0, 0, 0);

  const start = parseDateOnly(entry.startDate);
  const end = parseDateOnly(entry.endDate || entry.startDate);
  if (!start || !end) {
    return { key: "upcoming", label: "Scheduled" };
  }

  if (end < currentDay) return { key: "past", label: "Past" };
  if (start > currentDay) return { key: "upcoming", label: "Upcoming" };
  return { key: "active", label: "Away today" };
}

function getEntryTypeKey(value) {
  const key = normalizeValue(value);
  return ENTRY_TYPE_LABELS[key] ? key : (key || DEFAULT_TYPE);
}

function getEntryTypeLabel(value) {
  const key = getEntryTypeKey(value);
  return ENTRY_TYPE_LABELS[key] || titleCase(key);
}

function getTypeBadgeClass(type) {
  return `days-off-type-${getEntryTypeKey(type).replace(/[^a-z0-9]+/g, "-")}`;
}

function buildEntryId(entry, index) {
  const fragments = [
    normalizeValue(entry && entry.owner) || "team",
    getEntryTypeKey(entry && entry.type),
    String((entry && entry.startDate) || "").trim() || "date",
    String((entry && entry.endDate) || (entry && entry.startDate) || "").trim() || "date",
    String(index),
  ];

  return `days-off-${fragments.join("-").replace(/[^a-z0-9-]+/g, "-")}`;
}

function normalizeEntry(entry, index) {
  const owner = String((entry && entry.owner) || "").trim();
  const startDate = String((entry && (entry.startDate || entry.date)) || "").trim();
  let endDate = String((entry && (entry.endDate || entry.startDate || entry.date)) || "").trim();
  const parsedStart = parseDateOnly(startDate);
  const parsedEnd = parseDateOnly(endDate);

  if (parsedStart && !parsedEnd) {
    endDate = startDate;
  } else if (parsedStart && parsedEnd && parsedEnd < parsedStart) {
    endDate = startDate;
  }

  return {
    id: String((entry && entry.id) || "").trim() || buildEntryId(entry, index),
    owner,
    type: getEntryTypeKey(entry && entry.type),
    startDate,
    endDate: endDate || startDate,
    notes: String((entry && entry.notes) || "").trim(),
    createdAt: String((entry && entry.createdAt) || "").trim(),
    updatedAt: String((entry && entry.updatedAt) || "").trim(),
  };
}

function normalizeEntries(values) {
  return cloneArray(values)
    .map((entry, index) => normalizeEntry(entry, index))
    .filter((entry) => entry.owner || entry.startDate || entry.notes);
}

function loadDaysOffData() {
  if (AppCore && typeof AppCore.loadDaysOffData === "function") {
    daysOffEntries = normalizeEntries(AppCore.loadDaysOffData());
    return;
  }

  try {
    const raw = localStorage.getItem(DAYS_OFF_STORAGE_KEY);
    daysOffEntries = normalizeEntries(raw ? JSON.parse(raw) : window.DAYS_OFF);
  } catch {
    daysOffEntries = normalizeEntries(window.DAYS_OFF);
  }
}

function saveDaysOffData() {
  if (AppCore && typeof AppCore.saveDaysOffData === "function") {
    AppCore.saveDaysOffData(daysOffEntries);
  } else {
    try {
      localStorage.setItem(DAYS_OFF_STORAGE_KEY, JSON.stringify(daysOffEntries));
    } catch (error) {
      console.warn("Failed to save days off log", error);
    }
  }

  window.DAYS_OFF = cloneArray(daysOffEntries);
}

function loadRelatedData() {
  dealsData = AppCore && typeof AppCore.loadDealsData === "function"
    ? AppCore.loadDealsData()
    : cloneArray(window.DEALS);
  tasksData = AppCore && typeof AppCore.loadTasksData === "function"
    ? AppCore.loadTasksData()
    : cloneArray(window.TASKS);
}

function buildUniqueOwnerList(values) {
  const seen = new Set();
  return (Array.isArray(values) ? values : []).filter((entry) => {
    const value = String(entry || "").trim();
    const key = normalizeValue(value);
    if (!value || !key || key === "system" || seen.has(key)) return false;
    seen.add(key);
    return true;
  });
}

function getSubOwners(deal) {
  const source = deal && deal.subOwners;
  const values = Array.isArray(source)
    ? source
    : typeof source === "string"
      ? source.split(/[\n,;]+/)
      : [];
  return buildUniqueOwnerList(values);
}

function getAssignableOwnersForDeal(deal) {
  return buildUniqueOwnerList([
    deal && (deal.seniorOwner || deal.owner),
    deal && deal.juniorOwner,
    ...getSubOwners(deal),
  ]);
}

function getKnownPeople() {
  const ownersFromDeals = (Array.isArray(dealsData) ? dealsData : []).flatMap((deal) => getAssignableOwnersForDeal(deal));
  const ownersFromTasks = (Array.isArray(tasksData) ? tasksData : []).map((task) => task && task.owner);
  const ownersFromLog = (Array.isArray(daysOffEntries) ? daysOffEntries : []).map((entry) => entry && entry.owner);

  return buildUniqueOwnerList(["All team"].concat(ownersFromDeals, ownersFromTasks, ownersFromLog))
    .sort((left, right) => left.localeCompare(right));
}

function renderOwnerSuggestions() {
  const list = document.getElementById("days-off-owner-options");
  if (!list) return;

  list.innerHTML = getKnownPeople()
    .map((value) => `<option value="${escapeHtml(value)}"></option>`)
    .join("");
}

function setStatus(message, isError) {
  const pill = document.getElementById("days-off-status-pill");
  const line = document.getElementById("days-off-form-status");

  if (pill) {
    pill.textContent = message || "Calendar ready";
  }

  if (line) {
    line.textContent = message || "";
    line.style.color = isError ? "#dc2626" : "var(--text-soft)";
  }
}

function compareEntries(left, right) {
  const leftState = getEntryState(left).key;
  const rightState = getEntryState(right).key;
  const stateRank = { active: 0, upcoming: 1, past: 2 };
  if (stateRank[leftState] !== stateRank[rightState]) {
    return stateRank[leftState] - stateRank[rightState];
  }

  const leftStart = parseDateOnly(left.startDate);
  const rightStart = parseDateOnly(right.startDate);
  const leftEnd = parseDateOnly(left.endDate || left.startDate);
  const rightEnd = parseDateOnly(right.endDate || right.startDate);

  if (leftState === "past") {
    if (leftEnd && rightEnd && leftEnd.getTime() !== rightEnd.getTime()) {
      return rightEnd.getTime() - leftEnd.getTime();
    }
  } else if (leftStart && rightStart && leftStart.getTime() !== rightStart.getTime()) {
    return leftStart.getTime() - rightStart.getTime();
  }

  return normalizeValue(left.owner).localeCompare(normalizeValue(right.owner));
}

function matchesFilters(entry) {
  const state = getEntryState(entry).key;
  if (currentStatusFilter !== "all" && state !== currentStatusFilter) return false;
  if (currentTypeFilter !== "all" && getEntryTypeKey(entry.type) !== currentTypeFilter) return false;

  if (!currentSearch) return true;

  const haystack = [
    entry.owner,
    entry.type,
    entry.notes,
    entry.startDate,
    entry.endDate,
  ].map(normalizeValue).join(" ");

  return haystack.includes(currentSearch);
}

function getFilteredEntries() {
  return daysOffEntries.filter((entry) => matchesFilters(entry)).sort(compareEntries);
}

function entryIncludesDate(entry, dateKey) {
  const start = parseDateOnly(entry.startDate);
  const end = parseDateOnly(entry.endDate || entry.startDate);
  const date = parseDateOnly(dateKey);
  if (!start || !end || !date) return false;
  return start <= date && date <= end;
}

function getEntriesForDate(dateKey, entries) {
  return (Array.isArray(entries) ? entries : daysOffEntries)
    .filter((entry) => entryIncludesDate(entry, dateKey))
    .sort(compareEntries);
}

function getUniqueOwnerCount(entries) {
  const owners = new Set(
    (Array.isArray(entries) ? entries : [])
      .map((entry) => normalizeValue(entry.owner))
      .filter(Boolean),
  );
  return owners.size;
}

function renderMetaRow() {
  const row = document.getElementById("days-off-meta-row");
  if (!row) return;

  const today = new Date();
  today.setHours(0, 0, 0, 0);
  const upcomingLimit = addDays(today, 30);
  const activeToday = daysOffEntries.filter((entry) => getEntryState(entry, today).key === "active").length;
  const upcoming = daysOffEntries.filter((entry) => {
    const start = parseDateOnly(entry.startDate);
    return Boolean(start && start > today && start <= upcomingLimit);
  }).length;
  const publicHolidays = daysOffEntries.filter((entry) => getEntryTypeKey(entry.type) === "public holiday").length;
  const totalDays = daysOffEntries.reduce((sum, entry) => sum + calculateDaysCount(entry), 0);

  row.innerHTML = [
    `<div class="chip"><strong>${daysOffEntries.length}</strong> entries logged</div>`,
    `<div class="chip"><strong>${activeToday}</strong> active today</div>`,
    `<div class="chip"><strong>${upcoming}</strong> upcoming in 30 days</div>`,
    `<div class="chip"><strong>${publicHolidays}</strong> public holidays</div>`,
    `<div class="chip"><strong>${totalDays}</strong> total workdays logged</div>`,
  ].join("");
}

function getCalendarGridStart(monthDate) {
  const firstDay = startOfMonth(monthDate);
  const offset = (firstDay.getDay() + 7 - WEEKDAY_OFFSET) % 7;
  return addDays(firstDay, -offset);
}

function diffDays(left, right) {
  const start = left instanceof Date ? left : parseDateOnly(left);
  const end = right instanceof Date ? right : parseDateOnly(right);
  if (!start || !end) return 0;
  return Math.round((end.getTime() - start.getTime()) / DAY_MS);
}

function entryIntersectsRange(entry, rangeStart, rangeEnd) {
  const start = parseDateOnly(entry.startDate);
  const end = parseDateOnly(entry.endDate || entry.startDate);
  if (!start || !end) return false;
  return start <= rangeEnd && end >= rangeStart;
}

function compareCalendarSegments(left, right) {
  if (left.startColumn !== right.startColumn) {
    return left.startColumn - right.startColumn;
  }
  if (left.endColumnExclusive !== right.endColumnExclusive) {
    return right.endColumnExclusive - left.endColumnExclusive;
  }
  return normalizeValue(left.entry.owner).localeCompare(normalizeValue(right.entry.owner));
}

function buildWeekBarLanes(weekStart, weekEnd, entries) {
  const segments = (Array.isArray(entries) ? entries : [])
    .filter((entry) => entryIntersectsRange(entry, weekStart, weekEnd))
    .map((entry) => {
      const entryStart = parseDateOnly(entry.startDate);
      const entryEnd = parseDateOnly(entry.endDate || entry.startDate);
      const visibleStart = entryStart < weekStart ? weekStart : entryStart;
      const visibleEnd = entryEnd > weekEnd ? weekEnd : entryEnd;

      return {
        entry,
        startColumn: diffDays(weekStart, visibleStart) + 1,
        endColumnExclusive: diffDays(weekStart, visibleEnd) + 2,
        continuedLeft: entryStart < weekStart,
        continuedRight: entryEnd > weekEnd,
      };
    })
    .sort(compareCalendarSegments);

  const lanes = [];
  segments.forEach((segment) => {
    let targetLane = lanes.find((lane) => lane.lastEndExclusive <= segment.startColumn);
    if (!targetLane) {
      targetLane = { lastEndExclusive: 0, segments: [] };
      lanes.push(targetLane);
    }

    targetLane.segments.push(segment);
    targetLane.lastEndExclusive = segment.endColumnExclusive;
  });

  return lanes.map((lane) => lane.segments);
}

function ensureSelectedDateVisible() {
  const selectedDate = parseDateOnly(selectedDateKey);
  if (!selectedDate) {
    selectedDateKey = toDateKey(currentCalendarMonth);
    return;
  }

  if (
    selectedDate.getFullYear() !== currentCalendarMonth.getFullYear() ||
    selectedDate.getMonth() !== currentCalendarMonth.getMonth()
  ) {
    currentCalendarMonth = startOfMonth(selectedDate);
  }
}

function getCalendarCellTooltip(entry) {
  const parts = [
    entry.owner || "All team",
    getEntryTypeLabel(entry.type),
    formatDateRange(entry),
  ];
  if (entry.notes) parts.push(entry.notes);
  return parts.join(" | ");
}

function buildDaySummary(entries, dateKey) {
  const visibleEntries = Array.isArray(entries) ? entries : [];
  if (!visibleEntries.length) {
    return isWeekendDate(dateKey) ? "Weekend" : "";
  }

  const uniqueOwners = getUniqueOwnerCount(visibleEntries);
  const holidayCount = visibleEntries.filter((entry) => getEntryTypeKey(entry.type) === "public holiday").length;
  const firstOwner = normalizeValue(visibleEntries[0] && visibleEntries[0].owner);

  if (visibleEntries.length === 1 && holidayCount === 1 && firstOwner === "all team") {
    return "Holiday";
  }

  if (holidayCount && holidayCount === visibleEntries.length) {
    return `${holidayCount} holiday${holidayCount === 1 ? "" : "s"}`;
  }

  return `${uniqueOwners} away`;
}

function renderCalendar() {
  const monthEl = document.getElementById("days-off-calendar-month");
  const grid = document.getElementById("days-off-calendar-grid");
  if (!monthEl || !grid) return;

  ensureSelectedDateVisible();
  monthEl.textContent = formatMonthYear(currentCalendarMonth);

  const filteredEntries = getFilteredEntries();
  const calendarStart = getCalendarGridStart(currentCalendarMonth);
  const todayKey = toDateKey(new Date());
  const monthKey = `${currentCalendarMonth.getFullYear()}-${currentCalendarMonth.getMonth()}`;

  let markup = "";
  for (let weekIndex = 0; weekIndex < CALENDAR_WEEKS; weekIndex += 1) {
    const weekStart = addDays(calendarStart, weekIndex * 7);
    const weekEnd = addDays(weekStart, 6);
    const weekLanes = buildWeekBarLanes(weekStart, weekEnd, filteredEntries);

    const dayCells = Array.from({ length: 7 }, (_, dayIndex) => {
      const date = addDays(weekStart, dayIndex);
      const dateKey = toDateKey(date);
      const dateEntries = getEntriesForDate(dateKey, filteredEntries);
      const isOutsideMonth = `${date.getFullYear()}-${date.getMonth()}` !== monthKey;
      const classes = ["days-off-calendar-day"];

      if (dateEntries.length) classes.push("has-entries");
      if (isWeekendDate(date)) classes.push("is-weekend");
      if (isOutsideMonth) classes.push("is-outside-month");
      if (dateKey === todayKey) classes.push("is-today");
      if (dateKey === selectedDateKey) classes.push("is-selected");

      const countMarkup = dateEntries.length
        ? `<span class="days-off-calendar-day-count" title="${dateEntries.length} entr${dateEntries.length === 1 ? "y" : "ies"}">${dateEntries.length}</span>`
        : "";
      const summary = buildDaySummary(dateEntries, dateKey);

      return `
        <button
          class="${classes.join(" ")}"
          type="button"
          data-action="select-date"
          data-date-key="${escapeHtml(dateKey)}"
          aria-pressed="${dateKey === selectedDateKey ? "true" : "false"}"
        >
          <div class="days-off-calendar-day-head">
            <span class="days-off-calendar-day-number">${date.getDate()}</span>
            ${countMarkup}
          </div>
          <div class="days-off-calendar-day-meta">${summary ? escapeHtml(summary) : "&nbsp;"}</div>
        </button>
      `;
    }).join("");

    const lanesMarkup = weekLanes.length
      ? `
        <div class="days-off-calendar-bars">
          ${weekLanes.map((lane) => `
            <div class="days-off-calendar-lane">
              ${lane.map((segment) => {
                const classes = ["days-off-calendar-bar", getTypeBadgeClass(segment.entry.type)];
                if (segment.continuedLeft) classes.push("is-continued-left");
                if (segment.continuedRight) classes.push("is-continued-right");
                if (entryIncludesDate(segment.entry, selectedDateKey)) classes.push("is-on-selected-date");
                if (entryIncludesDate(segment.entry, todayKey)) classes.push("is-on-today");

                const ownerLabel = String(segment.entry.owner || "All team").trim() || "All team";
                const label = normalizeValue(ownerLabel) === "all team"
                  ? `${ownerLabel} · ${getEntryTypeLabel(segment.entry.type)}`
                  : ownerLabel;

                return `
                  <button
                    class="${classes.join(" ")}"
                    type="button"
                    data-action="edit-entry"
                    data-entry-id="${escapeHtml(segment.entry.id)}"
                    style="grid-column: ${segment.startColumn} / ${segment.endColumnExclusive};"
                    title="${escapeHtml(getCalendarCellTooltip(segment.entry))}"
                  >
                    <span class="days-off-calendar-bar-label">${escapeHtml(label)}</span>
                  </button>
                `;
              }).join("")}
            </div>
          `).join("")}
        </div>
      `
      : "";

    markup += `
      <div class="days-off-calendar-week">
        <div class="days-off-calendar-days-row">${dayCells}</div>
        ${lanesMarkup}
      </div>
    `;
  }

  grid.innerHTML = markup;
}

function renderSelectedDatePanel() {
  const titleEl = document.getElementById("days-off-selected-title");
  const subtitleEl = document.getElementById("days-off-selected-subtitle");
  const metaEl = document.getElementById("days-off-selected-meta");
  const listEl = document.getElementById("days-off-selected-list");
  if (!titleEl || !subtitleEl || !metaEl || !listEl) return;

  const selectedDate = parseDateOnly(selectedDateKey) || new Date();
  const selectedEntries = getEntriesForDate(toDateKey(selectedDate), getFilteredEntries());
  const uniqueOwners = getUniqueOwnerCount(selectedEntries);
  const publicHolidayCount = selectedEntries.filter((entry) => getEntryTypeKey(entry.type) === "public holiday").length;
  const todayKey = toDateKey(new Date());
  const isWeekend = isWeekendDate(selectedDate);

  titleEl.textContent = formatDisplayDate(toDateKey(selectedDate));
  subtitleEl.textContent = isWeekend
    ? "Weekend view in the team calendar."
    : selectedDateKey === todayKey
      ? "Today’s team availability."
      : "Selected date in the team calendar.";

  const metaChips = [
    `<div class="chip"><strong>${selectedEntries.length}</strong> entries</div>`,
    `<div class="chip"><strong>${uniqueOwners}</strong> ${uniqueOwners === 1 ? "person" : "people"}</div>`,
    `<div class="chip"><strong>${publicHolidayCount}</strong> holidays</div>`,
  ];
  if (isWeekend) {
    metaChips.push(`<div class="chip"><strong>Weekend</strong> Saturday / Sunday</div>`);
  }
  metaEl.innerHTML = metaChips.join("");

  if (!selectedEntries.length) {
    listEl.innerHTML = `
      <li class="days-off-empty-item">
        No one is marked away on ${escapeHtml(formatDisplayDate(toDateKey(selectedDate)))} with the current filters.
      </li>
    `;
    return;
  }

  listEl.innerHTML = selectedEntries.map((entry) => {
    const state = getEntryState(entry);
    const workdaysCount = calculateDaysCount(entry);
    const calendarDaysCount = calculateCalendarDaysCount(entry);
    const dayLabel = workdaysCount === 1 ? "workday" : "workdays";
    const countLabel = workdaysCount === calendarDaysCount
      ? `${workdaysCount} ${dayLabel}`
      : `${calendarDaysCount} calendar days · ${workdaysCount} ${dayLabel}`;

    return `
      <li class="days-off-selected-item">
        <div class="days-off-selected-item-top">
          <div>
            <div class="days-off-selected-person">${escapeHtml(entry.owner || "All team")}</div>
            <div class="days-off-selected-submeta">${escapeHtml(formatDateRange(entry))} · ${escapeHtml(countLabel)}</div>
          </div>
          <span class="days-off-type-badge ${getTypeBadgeClass(entry.type)}">${escapeHtml(getEntryTypeLabel(entry.type))}</span>
        </div>
        <div class="days-off-selected-submeta">
          <span class="days-off-state-badge days-off-state-${state.key}">${escapeHtml(state.label)}</span>
        </div>
        ${entry.notes ? `<div class="days-off-selected-notes">${escapeHtml(entry.notes)}</div>` : ""}
        <div class="days-off-selected-actions">
          <button class="btn" type="button" data-action="edit-entry" data-entry-id="${escapeHtml(entry.id)}">Edit</button>
          <button class="btn days-off-delete-btn" type="button" data-action="delete-entry" data-entry-id="${escapeHtml(entry.id)}">Delete</button>
        </div>
      </li>
    `;
  }).join("");
}

function renderTable() {
  const body = document.getElementById("days-off-body");
  const empty = document.getElementById("days-off-empty");
  const count = document.getElementById("days-off-count");
  if (!body || !empty || !count) return;

  const visibleEntries = getFilteredEntries();
  count.textContent = `${visibleEntries.length} shown`;

  if (!visibleEntries.length) {
    body.innerHTML = "";
    empty.hidden = false;
    return;
  }

  empty.hidden = true;
  body.innerHTML = visibleEntries.map((entry) => {
    const state = getEntryState(entry);
    const daysCount = calculateDaysCount(entry);
    return `
      <tr>
        <td class="name-cell">${escapeHtml(entry.owner || "All team")}</td>
        <td><span class="days-off-type-badge ${getTypeBadgeClass(entry.type)}">${escapeHtml(getEntryTypeLabel(entry.type))}</span></td>
        <td>${escapeHtml(formatDateRange(entry))}</td>
        <td>${daysCount}</td>
        <td><span class="days-off-state-badge days-off-state-${state.key}">${escapeHtml(state.label)}</span></td>
        <td class="days-off-notes-cell">${escapeHtml(entry.notes || "—")}</td>
        <td>
          <div class="days-off-actions">
            <button class="btn" type="button" data-action="edit-entry" data-entry-id="${escapeHtml(entry.id)}">Edit</button>
            <button class="btn days-off-delete-btn" type="button" data-action="delete-entry" data-entry-id="${escapeHtml(entry.id)}">Delete</button>
          </div>
        </td>
      </tr>
    `;
  }).join("");
}

function renderAll() {
  renderMetaRow();
  renderCalendar();
  renderSelectedDatePanel();
  renderTable();
  renderOwnerSuggestions();
}

function resetForm() {
  const form = document.getElementById("days-off-form");
  const typeInput = document.getElementById("days-off-type");
  const title = document.getElementById("days-off-form-title");
  const submitButton = document.getElementById("days-off-submit-btn");
  const cancelButton = document.getElementById("days-off-cancel-edit");

  if (form) form.reset();
  if (typeInput) typeInput.value = DEFAULT_TYPE;
  if (title) title.textContent = "Log time off";
  if (submitButton) submitButton.textContent = "Save entry";
  if (cancelButton) cancelButton.hidden = true;

  editingEntryId = "";
  syncEndDateLimit();
  setStatus("Ready to log time off.", false);
}

function prefillFormForDate(dateKey) {
  const parsed = parseDateOnly(dateKey);
  if (!parsed) return;

  resetForm();
  const title = document.getElementById("days-off-form-title");
  const startInput = document.getElementById("days-off-start-date");
  const endInput = document.getElementById("days-off-end-date");
  const ownerInput = document.getElementById("days-off-owner");
  const formCard = document.getElementById("days-off-form-card");

  if (title) title.textContent = `Log time off for ${formatDisplayDate(dateKey)}`;
  if (startInput) startInput.value = dateKey;
  if (endInput) endInput.value = dateKey;
  syncEndDateLimit();

  setStatus(`Selected ${formatDisplayDate(dateKey)}. Add the person and save.`, false);
  if (formCard && typeof formCard.scrollIntoView === "function") {
    formCard.scrollIntoView({ behavior: "smooth", block: "start" });
  }
  if (ownerInput) {
    ownerInput.focus();
    ownerInput.select();
  }
}

function setSelectedDate(dateKey, options = {}) {
  const parsed = parseDateOnly(dateKey);
  if (!parsed) return;

  selectedDateKey = toDateKey(parsed);
  if (options.syncMonth !== false) {
    currentCalendarMonth = startOfMonth(parsed);
  }

  renderCalendar();
  renderSelectedDatePanel();
}

function startEditing(entryId) {
  const entry = daysOffEntries.find((item) => item.id === entryId);
  if (!entry) return;

  editingEntryId = entry.id;
  setSelectedDate(entry.startDate || selectedDateKey, { syncMonth: true });
  document.getElementById("days-off-owner").value = entry.owner || "";
  document.getElementById("days-off-type").value = getEntryTypeKey(entry.type);
  document.getElementById("days-off-start-date").value = entry.startDate || "";
  document.getElementById("days-off-end-date").value = entry.endDate || entry.startDate || "";
  document.getElementById("days-off-notes").value = entry.notes || "";
  document.getElementById("days-off-form-title").textContent = "Edit time off entry";
  document.getElementById("days-off-submit-btn").textContent = "Update entry";
  document.getElementById("days-off-cancel-edit").hidden = false;

  syncEndDateLimit();
  setStatus("Editing existing entry.", false);
  const ownerInput = document.getElementById("days-off-owner");
  const formCard = document.getElementById("days-off-form-card");
  if (formCard && typeof formCard.scrollIntoView === "function") {
    formCard.scrollIntoView({ behavior: "smooth", block: "start" });
  }
  if (ownerInput) {
    ownerInput.focus();
    ownerInput.select();
  }
}

function removeEntry(entryId) {
  const nextEntries = daysOffEntries.filter((entry) => entry.id !== entryId);
  if (nextEntries.length === daysOffEntries.length) return;

  const removedWhileEditing = editingEntryId === entryId;
  daysOffEntries = nextEntries;
  saveDaysOffData();
  if (removedWhileEditing) {
    resetForm();
  }
  setStatus("Entry deleted.", false);
  renderAll();
}

function syncEndDateLimit() {
  const startInput = document.getElementById("days-off-start-date");
  const endInput = document.getElementById("days-off-end-date");
  if (!startInput || !endInput) return;

  endInput.min = startInput.value || "";
  if (!endInput.value && startInput.value) {
    endInput.value = startInput.value;
  }
  if (endInput.value && startInput.value && endInput.value < startInput.value) {
    endInput.value = startInput.value;
  }
}

function handleFormSubmit(event) {
  event.preventDefault();

  const ownerInput = document.getElementById("days-off-owner");
  const typeInput = document.getElementById("days-off-type");
  const startInput = document.getElementById("days-off-start-date");
  const endInput = document.getElementById("days-off-end-date");
  const notesInput = document.getElementById("days-off-notes");
  if (!ownerInput || !typeInput || !startInput || !endInput || !notesInput) return;

  const type = getEntryTypeKey(typeInput.value);
  const owner = String(ownerInput.value || "").trim() || (type === "public holiday" ? "All team" : "");
  const startDate = String(startInput.value || "").trim();
  const endDate = String(endInput.value || startDate).trim() || startDate;
  const notes = String(notesInput.value || "").trim();

  if (!owner) {
    setStatus("Please add a person or use All team.", true);
    ownerInput.focus();
    return;
  }

  if (!startDate) {
    setStatus("Please choose a start date.", true);
    startInput.focus();
    return;
  }

  if (parseDateOnly(endDate) && parseDateOnly(startDate) && endDate < startDate) {
    setStatus("End date cannot be before the start date.", true);
    endInput.focus();
    return;
  }

  const timestamp = new Date().toISOString();
  const normalizedEntry = normalizeEntry({
    id: editingEntryId || `days-off-${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 8)}`,
    owner,
    type,
    startDate,
    endDate,
    notes,
    createdAt: editingEntryId
      ? (daysOffEntries.find((entry) => entry.id === editingEntryId)?.createdAt || timestamp)
      : timestamp,
    updatedAt: timestamp,
  }, daysOffEntries.length);

  const successMessage = editingEntryId ? "Entry updated." : "Entry saved.";
  if (editingEntryId) {
    daysOffEntries = daysOffEntries.map((entry) => entry.id === editingEntryId ? normalizedEntry : entry);
  } else {
    daysOffEntries = [normalizedEntry].concat(daysOffEntries);
  }

  selectedDateKey = normalizedEntry.startDate || selectedDateKey;
  currentCalendarMonth = startOfMonth(parseDateOnly(selectedDateKey) || new Date());
  saveDaysOffData();
  resetForm();
  setStatus(successMessage, false);
  renderAll();
}

function handleAction(action, element) {
  const entryId = String((element && element.getAttribute("data-entry-id")) || "").trim();
  const dateKey = String((element && element.getAttribute("data-date-key")) || "").trim();

  if (action === "edit-entry" && entryId) {
    startEditing(entryId);
    return;
  }

  if (action === "delete-entry" && entryId) {
    removeEntry(entryId);
    return;
  }

  if (action === "select-date" && dateKey) {
    setSelectedDate(dateKey, { syncMonth: true });
    return;
  }

  if (action === "calendar-prev") {
    currentCalendarMonth = shiftMonth(currentCalendarMonth, -1);
    selectedDateKey = toDateKey(currentCalendarMonth);
    renderCalendar();
    renderSelectedDatePanel();
    return;
  }

  if (action === "calendar-next") {
    currentCalendarMonth = shiftMonth(currentCalendarMonth, 1);
    selectedDateKey = toDateKey(currentCalendarMonth);
    renderCalendar();
    renderSelectedDatePanel();
    return;
  }

  if (action === "calendar-today") {
    const todayKey = toDateKey(new Date());
    setSelectedDate(todayKey, { syncMonth: true });
    return;
  }

  if (action === "calendar-log-selected") {
    prefillFormForDate(selectedDateKey);
  }
}

function bindEvents() {
  const form = document.getElementById("days-off-form");
  const startInput = document.getElementById("days-off-start-date");
  const cancelButton = document.getElementById("days-off-cancel-edit");
  const searchInput = document.getElementById("days-off-search");
  const statusFilter = document.getElementById("days-off-status-filter");
  const typeFilter = document.getElementById("days-off-type-filter");

  if (form) {
    form.addEventListener("submit", handleFormSubmit);
  }

  if (startInput) {
    startInput.addEventListener("change", syncEndDateLimit);
  }

  if (cancelButton) {
    cancelButton.addEventListener("click", resetForm);
  }

  if (searchInput) {
    searchInput.addEventListener("input", () => {
      currentSearch = normalizeValue(searchInput.value);
      renderAll();
    });
  }

  if (statusFilter) {
    statusFilter.addEventListener("change", () => {
      currentStatusFilter = normalizeValue(statusFilter.value) || "all";
      renderAll();
    });
  }

  if (typeFilter) {
    typeFilter.addEventListener("change", () => {
      currentTypeFilter = normalizeValue(typeFilter.value) || "all";
      renderAll();
    });
  }

  document.addEventListener("click", (event) => {
    const actionEl = event.target.closest("[data-action]");
    if (!actionEl) return;
    handleAction(String(actionEl.getAttribute("data-action") || "").trim(), actionEl);
  });

  window.addEventListener("appcore:days-off-updated", (event) => {
    if (!event || !event.detail || !Array.isArray(event.detail.entries)) return;
    daysOffEntries = normalizeEntries(event.detail.entries);
    renderAll();
  });
}

function initializeDaysOffPage() {
  loadRelatedData();
  loadDaysOffData();
  bindEvents();
  resetForm();
  syncEndDateLimit();
  renderAll();
}

if (document.readyState === "loading") {
  document.addEventListener("DOMContentLoaded", initializeDaysOffPage, { once: true });
} else {
  initializeDaysOffPage();
}
