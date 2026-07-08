const DCI_WORK_UNITS = [
  "Zone 1 East",
  "Zone 1 West",
  "Zone 2",
  "Zone 2 SLANT",
  "Zone 2 QCMEG",
  "Zone 2 BATF",
  "Zone 2 RIITF",
  "Zone 4",
  "Zone 4 CIFG",
  "Zone 4 WCITF Macomb",
  "Zone 4 WCITF Quincy",
  "Zone 4 PMEG",
  "Zone 5",
  "Zone 5 TF6",
  "Zone 5 VMEG",
  "Zone 5 ECITF",
  "Zone 6",
  "Zone 6 PSEG",
  "Zone 6 SCIDTF",
  "Zone 6 MEGSI",
  "Zone 7",
  "Zone 7 SIEG",
  "Zone 7 SIDTF",
  "Zone 8",
  "Zone 8 SEIDTF",
  "SIU",
  "STIC",
  "Digital Crime Unit",
  "JTTF",
  "SOCOM",
  "Air Operations",
  "Statewide Gaming",
  "DDO",
  "ISC",
  "DDO FOIA Unit"
];

const ALL_WORK_UNITS_OPTION = "ALL";

const state = {
  records: [],
  scopedRecords: [],
  filteredRecords: [],
  selectedId: null,
  selectedWorkUnit: ALL_WORK_UNITS_OPTION,
  searchQuery: ""
};

const elements = {
  kpiTotalOpen: document.getElementById("kpiTotalOpen"),
  kpiDue10: document.getElementById("kpiDue10"),
  kpiDue5: document.getElementById("kpiDue5"),
  kpiDueToday: document.getElementById("kpiDueToday"),
  kpiOverdue: document.getElementById("kpiOverdue"),
  kpiAvgDaysOpen: document.getElementById("kpiAvgDaysOpen"),
  kpiClosedThisMonth: document.getElementById("kpiClosedThisMonth"),
  kpiOpenWorkUnits: document.getElementById("kpiOpenWorkUnits"),
  workUnitSummaryBody: document.getElementById("workUnitSummaryBody"),
  upcomingDeadlines: document.getElementById("upcomingDeadlines"),
  commandAlerts: document.getElementById("commandAlerts"),
  foiaTableBody: document.getElementById("foiaTableBody"),
  detailContent: document.getElementById("detailContent"),
  searchInput: document.getElementById("searchInput"),
  workUnitFilter: document.getElementById("workUnitFilter"),
  darkModeToggle: document.getElementById("darkModeToggle"),
  presentationToggle: document.getElementById("presentationToggle")
};

const today = startOfDay(new Date());

initialize();

async function initialize() {
  attachEventListeners();
  await loadData();
  populateWorkUnitFilter();
  applyFilters();
}

function attachEventListeners() {
  elements.searchInput.addEventListener("input", (event) => {
    state.searchQuery = event.target.value;
    applyFilters();
  });

  elements.workUnitFilter.addEventListener("change", (event) => {
    state.selectedWorkUnit = event.target.value;
    applyFilters();
  });

  elements.darkModeToggle.addEventListener("click", () => {
    document.body.classList.toggle("dark");
    elements.darkModeToggle.setAttribute("aria-pressed", String(document.body.classList.contains("dark")));
    elements.darkModeToggle.textContent = document.body.classList.contains("dark") ? "Light Mode" : "Dark Mode";
  });

  elements.presentationToggle.addEventListener("click", () => {
    document.body.classList.toggle("presentation-mode");
    elements.presentationToggle.setAttribute("aria-pressed", String(document.body.classList.contains("presentation-mode")));
    elements.presentationToggle.textContent = document.body.classList.contains("presentation-mode")
      ? "Exit Presentation Mode"
      : "Presentation Mode";
  });
}

async function loadData() {
  try {
    const response = await fetch("foia-data.json", { cache: "no-store" });
    if (!response.ok) {
      throw new Error(`Data request failed with status ${response.status}`);
    }

    const rawRecords = await response.json();

    state.records = rawRecords
      .map((record) => normalizeRecord(record))
      .filter((record) => DCI_WORK_UNITS.includes(record.dci_work_unit));
  } catch (error) {
    elements.foiaTableBody.innerHTML = `<tr><td colspan="8">Failed to load FOIA data: ${escapeHtml(error.message)}</td></tr>`;
  }
}

function populateWorkUnitFilter() {
  elements.workUnitFilter.innerHTML = [
    `<option value="${ALL_WORK_UNITS_OPTION}">All DCI Work Units</option>`,
    ...DCI_WORK_UNITS.map((workUnit) => `<option value="${escapeHtml(workUnit)}">${escapeHtml(workUnit)}</option>`)
  ].join("");
}

function applyFilters() {
  const normalizedQuery = state.searchQuery.trim().toLowerCase();

  state.scopedRecords = state.records.filter((record) => {
    if (state.selectedWorkUnit === ALL_WORK_UNITS_OPTION) {
      return true;
    }
    return record.dci_work_unit === state.selectedWorkUnit;
  });

  state.filteredRecords = state.scopedRecords.filter((record) => {
    if (!normalizedQuery) {
      return true;
    }

    const searchable = [
      record.request_id,
      record.requester,
      record.subject,
      record.dci_work_unit,
      record.status,
      record.assigned_to,
      record.notes
    ].join(" ").toLowerCase();

    return searchable.includes(normalizedQuery);
  });

  if (!state.filteredRecords.some((item) => item.request_id === state.selectedId)) {
    state.selectedId = state.filteredRecords[0]?.request_id ?? null;
  }

  renderAll();
}

function renderAll() {
  renderKpis();
  renderWorkUnitSummary();
  renderUpcomingDeadlines();
  renderCommandAlerts();
  renderRequestTable();
  renderDetailPanel();
}

function renderKpis() {
  const openRecords = state.scopedRecords.filter((record) => isOpenStatus(record.status));
  const due10 = openRecords.filter((record) => daysUntil(record.due_date) <= 10 && daysUntil(record.due_date) >= 0);
  const due5 = openRecords.filter((record) => daysUntil(record.due_date) <= 5 && daysUntil(record.due_date) >= 0);
  const dueToday = openRecords.filter((record) => daysUntil(record.due_date) === 0);
  const overdue = openRecords.filter((record) => daysUntil(record.due_date) < 0);

  const totalOpenDays = openRecords.reduce((sum, record) => sum + daysBetween(record.received_date, today), 0);
  const avgDaysOpen = openRecords.length ? Math.round(totalOpenDays / openRecords.length) : 0;

  const now = new Date();
  const closedThisMonth = state.scopedRecords.filter((record) => {
    if (!record.closed_date) {
      return false;
    }
    return record.closed_date.getMonth() === now.getMonth() && record.closed_date.getFullYear() === now.getFullYear();
  });

  const openWorkUnits = new Set(openRecords.map((record) => record.dci_work_unit));

  elements.kpiTotalOpen.textContent = String(openRecords.length);
  elements.kpiDue10.textContent = String(due10.length);
  elements.kpiDue5.textContent = String(due5.length);
  elements.kpiDueToday.textContent = String(dueToday.length);
  elements.kpiOverdue.textContent = String(overdue.length);
  elements.kpiAvgDaysOpen.textContent = String(avgDaysOpen);
  elements.kpiClosedThisMonth.textContent = String(closedThisMonth.length);
  elements.kpiOpenWorkUnits.textContent = String(openWorkUnits.size);
}

function renderWorkUnitSummary() {
  const unitsToRender = state.selectedWorkUnit === ALL_WORK_UNITS_OPTION
    ? DCI_WORK_UNITS
    : [state.selectedWorkUnit];

  const rows = unitsToRender
    .map((workUnit) => {
      const unitRecords = state.scopedRecords.filter((record) => record.dci_work_unit === workUnit);
      const openRecords = unitRecords.filter((record) => isOpenStatus(record.status));
      const due10 = openRecords.filter((record) => {
        const delta = daysUntil(record.due_date);
        return delta >= 0 && delta <= 10;
      }).length;
      const due5 = openRecords.filter((record) => {
        const delta = daysUntil(record.due_date);
        return delta >= 0 && delta <= 5;
      }).length;
      const dueToday = openRecords.filter((record) => daysUntil(record.due_date) === 0).length;
      const overdue = openRecords.filter((record) => daysUntil(record.due_date) < 0).length;

      return {
        workUnit,
        open: openRecords.length,
        due10,
        due5,
        dueToday,
        overdue
      };
    })
    .sort((a, b) => b.open - a.open || a.workUnit.localeCompare(b.workUnit));

  if (!rows.length) {
    elements.workUnitSummaryBody.innerHTML = "<tr><td colspan=\"6\">No DCI work unit workload data available.</td></tr>";
    return;
  }

  elements.workUnitSummaryBody.innerHTML = rows.map((item) => `
    <tr>
      <td>${escapeHtml(item.workUnit)}</td>
      <td>${item.open}</td>
      <td>${item.due10}</td>
      <td>${item.due5}</td>
      <td>${item.dueToday}</td>
      <td>${item.overdue}</td>
    </tr>
  `).join("");
}

function renderCommandAlerts() {
  const openRecords = state.scopedRecords.filter((record) => isOpenStatus(record.status));
  const overdueRecords = openRecords
    .map((record) => ({ ...record, delta: daysUntil(record.due_date) }))
    .filter((record) => record.delta < 0)
    .sort((a, b) => a.delta - b.delta);

  const upcomingFiveDays = openRecords
    .map((record) => ({ ...record, delta: daysUntil(record.due_date) }))
    .filter((record) => record.delta >= 0 && record.delta <= 5)
    .sort((a, b) => a.delta - b.delta);

  const dueToday = openRecords.filter((record) => daysUntil(record.due_date) === 0);
  const nextCritical = overdueRecords[0] || upcomingFiveDays[0] || null;

  const overduePreview = overdueRecords
    .slice(0, 2)
    .map((record) => `${record.request_id} - ${record.dci_work_unit} (${Math.abs(record.delta)}d overdue)`)
    .join(" | ");

  const upcomingPreview = upcomingFiveDays
    .slice(0, 2)
    .map((record) => `${record.request_id} - ${record.dci_work_unit} (${record.delta}d)`)
    .join(" | ");

  elements.commandAlerts.innerHTML = `
    <li class="alert-item">
      <div class="alert-title">Overdue Requests</div>
      <div class="alert-value">${overdueRecords.length}</div>
      <div class="alert-note">${escapeHtml(overduePreview || "No overdue requests")}</div>
    </li>
    <li class="alert-item">
      <div class="alert-title">Due in 5 Days</div>
      <div class="alert-value">${upcomingFiveDays.length}</div>
      <div class="alert-note">${escapeHtml(upcomingPreview || "No deadlines in the next 5 days")}</div>
    </li>
    <li class="alert-item">
      <div class="alert-title">Critical Today</div>
      <div class="alert-value">${dueToday.length}</div>
      <div class="alert-note">${escapeHtml(nextCritical ? `${nextCritical.request_id} - ${formatDueIn(nextCritical.delta)}` : "No critical deadlines")}</div>
    </li>
  `;
}

function renderUpcomingDeadlines() {
  const upcoming = state.scopedRecords
    .filter((record) => isOpenStatus(record.status))
    .map((record) => ({
      ...record,
      dueInDays: daysUntil(record.due_date)
    }))
    .sort((a, b) => a.dueInDays - b.dueInDays);

  if (!upcoming.length) {
    elements.upcomingDeadlines.innerHTML = "<li class=\"placeholder\">No upcoming open deadlines.</li>";
    return;
  }

  elements.upcomingDeadlines.innerHTML = upcoming.map((record) => {
    const timing = formatDueIn(record.dueInDays);
    return `
      <li class="deadline-item">
        <strong>${escapeHtml(record.request_id)} - ${escapeHtml(record.dci_work_unit)}</strong>
        <span>${escapeHtml(record.subject)}</span>
        <div class="deadline-meta">
          <span>Due ${formatDate(record.due_date)}</span>
          <span>${timing}</span>
        </div>
      </li>
    `;
  }).join("");
}

function renderRequestTable() {
  if (!state.filteredRecords.length) {
    elements.foiaTableBody.innerHTML = "<tr><td colspan=\"8\">No FOIA records match your search.</td></tr>";
    return;
  }

  elements.foiaTableBody.innerHTML = state.filteredRecords.map((record) => {
    const selectedClass = record.request_id === state.selectedId ? "selected" : "";
    const statusClass = `status-${record.status.toLowerCase().replace(/\s+/g, "-")}`;
    return `
      <tr data-request-id="${escapeHtml(record.request_id)}" class="${selectedClass}" tabindex="0">
        <td>${escapeHtml(record.request_id)}</td>
        <td>${escapeHtml(record.subject)}</td>
        <td>${escapeHtml(record.requester)}</td>
        <td>${escapeHtml(record.dci_work_unit)}</td>
        <td><span class="status-chip ${statusClass}">${escapeHtml(record.status)}</span></td>
        <td>${daysBetween(record.received_date, today)}</td>
        <td>${formatDate(record.due_date)}</td>
        <td>${escapeHtml(formatDueInShort(daysUntil(record.due_date)))}</td>
      </tr>
    `;
  }).join("");

  Array.from(elements.foiaTableBody.querySelectorAll("tr[data-request-id]")).forEach((row) => {
    const select = () => {
      state.selectedId = row.getAttribute("data-request-id");
      renderRequestTable();
      renderDetailPanel();
    };

    row.addEventListener("click", select);
    row.addEventListener("keydown", (event) => {
      if (event.key === "Enter" || event.key === " ") {
        event.preventDefault();
        select();
      }
    });
  });
}

function renderDetailPanel() {
  const selected = state.scopedRecords.find((record) => record.request_id === state.selectedId);

  if (!selected) {
    elements.detailContent.innerHTML = "<p class=\"placeholder\">Select a FOIA request to view detailed information.</p>";
    return;
  }

  const dueInDays = daysUntil(selected.due_date);
  const dueLabel = formatDueIn(dueInDays);
  const closedDate = selected.closed_date ? formatDate(selected.closed_date) : "N/A";

  elements.detailContent.innerHTML = `
    <h3>${escapeHtml(selected.request_id)}</h3>
    <div class="detail-grid">
      <span class="detail-label">DCI Work Unit</span><span>${escapeHtml(selected.dci_work_unit)}</span>
      <span class="detail-label">Requester</span><span>${escapeHtml(selected.requester)}</span>
      <span class="detail-label">Subject</span><span>${escapeHtml(selected.subject)}</span>
      <span class="detail-label">Status</span><span>${escapeHtml(selected.status)}</span>
      <span class="detail-label">Days Open</span><span>${daysBetween(selected.received_date, today)}</span>
      <span class="detail-label">Received</span><span>${formatDate(selected.received_date)}</span>
      <span class="detail-label">Due Date</span><span>${formatDate(selected.due_date)} (${dueLabel})</span>
      <span class="detail-label">Days Remaining</span><span>${escapeHtml(formatDueInShort(dueInDays))}</span>
      <span class="detail-label">Assigned To</span><span>${escapeHtml(selected.assigned_to)}</span>
      <span class="detail-label">Last Update</span><span>${formatDate(selected.last_update)}</span>
      <span class="detail-label">Closed Date</span><span>${closedDate}</span>
      <span class="detail-label">Notes</span><span>${escapeHtml(selected.notes)}</span>
    </div>
  `;
}

function normalizeRecord(record) {
  const workUnit = normalizeWorkUnit(record.dci_work_unit || record.work_unit);
  return {
    ...record,
    dci_work_unit: workUnit,
    status: normalizeStatus(record.status),
    received_date: parseDate(record.received_date),
    due_date: parseDate(record.due_date),
    last_update: parseDate(record.last_update),
    closed_date: record.closed_date ? parseDate(record.closed_date) : null
  };
}

function normalizeWorkUnit(value) {
  const normalized = String(value || "").trim();
  if (DCI_WORK_UNITS.includes(normalized)) {
    return normalized;
  }
  return "Unknown";
}

function normalizeStatus(status) {
  const normalized = String(status || "Open").trim().toLowerCase();
  if (normalized === "closed") {
    return "Closed";
  }
  if (normalized === "in review") {
    return "In Review";
  }
  return "Open";
}

function isOpenStatus(status) {
  return status !== "Closed";
}

function parseDate(value) {
  const date = new Date(value);
  return startOfDay(date);
}

function startOfDay(date) {
  const copy = new Date(date);
  copy.setHours(0, 0, 0, 0);
  return copy;
}

function daysBetween(first, second) {
  const millisecondsPerDay = 86400000;
  return Math.round((startOfDay(second) - startOfDay(first)) / millisecondsPerDay);
}

function daysUntil(date) {
  return daysBetween(today, date);
}

function formatDate(date) {
  if (!(date instanceof Date) || Number.isNaN(date.getTime())) {
    return "N/A";
  }
  return date.toLocaleDateString("en-US", {
    month: "short",
    day: "2-digit",
    year: "numeric"
  });
}

function formatDueIn(days) {
  if (days < 0) {
    return `${Math.abs(days)} day${Math.abs(days) === 1 ? "" : "s"} overdue`;
  }
  if (days === 0) {
    return "Due today";
  }
  return `Due in ${days} day${days === 1 ? "" : "s"}`;
}

function formatDueInShort(days) {
  if (days < 0) {
    return `${Math.abs(days)} overdue`;
  }
  if (days === 0) {
    return "Due today";
  }
  return String(days);
}

function escapeHtml(value) {
  return String(value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/\"/g, "&quot;")
    .replace(/'/g, "&#39;");
}