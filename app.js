const DIVISIONS = ["DCI", "DJS", "DOP", "DII", "DFS"];

const state = {
  records: [],
  filteredRecords: [],
  selectedId: null
};

const elements = {
  kpiTotalOpen: document.getElementById("kpiTotalOpen"),
  kpiDue10: document.getElementById("kpiDue10"),
  kpiDue5: document.getElementById("kpiDue5"),
  kpiDueToday: document.getElementById("kpiDueToday"),
  kpiOverdue: document.getElementById("kpiOverdue"),
  kpiAvgDaysOpen: document.getElementById("kpiAvgDaysOpen"),
  kpiClosedThisMonth: document.getElementById("kpiClosedThisMonth"),
  kpiRequestsByDivision: document.getElementById("kpiRequestsByDivision"),
  divisionChartBars: document.getElementById("divisionChartBars"),
  upcomingDeadlines: document.getElementById("upcomingDeadlines"),
  commandAlerts: document.getElementById("commandAlerts"),
  foiaTableBody: document.getElementById("foiaTableBody"),
  detailContent: document.getElementById("detailContent"),
  searchInput: document.getElementById("searchInput"),
  darkModeToggle: document.getElementById("darkModeToggle"),
  presentationToggle: document.getElementById("presentationToggle")
};

const today = startOfDay(new Date());

initialize();

async function initialize() {
  attachEventListeners();
  await loadData();
  applySearch("");
}

function attachEventListeners() {
  elements.searchInput.addEventListener("input", (event) => {
    applySearch(event.target.value);
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
      .filter((record) => DIVISIONS.includes(record.division));
  } catch (error) {
    elements.foiaTableBody.innerHTML = `<tr><td colspan="8">Failed to load FOIA data: ${escapeHtml(error.message)}</td></tr>`;
  }
}

function applySearch(query) {
  const normalizedQuery = query.trim().toLowerCase();
  state.filteredRecords = state.records.filter((record) => {
    if (!normalizedQuery) {
      return true;
    }

    const searchable = [
      record.request_id,
      record.requester,
      record.subject,
      record.division,
      record.status,
      record.assigned_to
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
  renderDivisionWorkloadChart();
  renderUpcomingDeadlines();
  renderCommandAlerts();
  renderRequestTable();
  renderDetailPanel();
}

function renderKpis() {
  const openRecords = state.records.filter((record) => isOpenStatus(record.status));
  const due10 = openRecords.filter((record) => daysUntil(record.due_date) <= 10 && daysUntil(record.due_date) >= 0);
  const due5 = openRecords.filter((record) => daysUntil(record.due_date) <= 5 && daysUntil(record.due_date) >= 0);
  const dueToday = openRecords.filter((record) => daysUntil(record.due_date) === 0);
  const overdue = openRecords.filter((record) => daysUntil(record.due_date) < 0);

  const totalOpenDays = openRecords.reduce((sum, record) => sum + daysBetween(record.received_date, today), 0);
  const avgDaysOpen = openRecords.length ? Math.round(totalOpenDays / openRecords.length) : 0;

  const now = new Date();
  const closedThisMonth = state.records.filter((record) => {
    if (!record.closed_date) {
      return false;
    }
    return record.closed_date.getMonth() === now.getMonth() && record.closed_date.getFullYear() === now.getFullYear();
  });

  const divisionCounts = getDivisionCounts(openRecords);

  elements.kpiTotalOpen.textContent = String(openRecords.length);
  elements.kpiDue10.textContent = String(due10.length);
  elements.kpiDue5.textContent = String(due5.length);
  elements.kpiDueToday.textContent = String(dueToday.length);
  elements.kpiOverdue.textContent = String(overdue.length);
  elements.kpiAvgDaysOpen.textContent = String(avgDaysOpen);
  elements.kpiClosedThisMonth.textContent = String(closedThisMonth.length);

  elements.kpiRequestsByDivision.innerHTML = DIVISIONS
    .map((division) => `<span>${division}: <strong>${divisionCounts[division] ?? 0}</strong></span>`)
    .join("");
}

function renderDivisionWorkloadChart() {
  const openByDivision = DIVISIONS.map((division) => {
    const openCount = state.records.filter((record) => record.division === division && isOpenStatus(record.status)).length;
    return { division, openCount };
  });

  const maxCount = Math.max(...openByDivision.map((item) => item.openCount), 1);

  elements.divisionChartBars.innerHTML = openByDivision.map((item) => {
    const percent = Math.round((item.openCount / maxCount) * 100);
    return `
      <li class="chart-row">
        <span class="chart-label">${item.division}</span>
        <div class="chart-track" role="img" aria-label="${item.division} has ${item.openCount} open requests">
          <div class="chart-bar" style="width: ${percent}%"></div>
        </div>
        <span class="chart-value">${item.openCount}</span>
      </li>
    `;
  }).join("");
}

function renderCommandAlerts() {
  const openRecords = state.records.filter((record) => isOpenStatus(record.status));
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
    .map((record) => `${record.request_id} (${Math.abs(record.delta)}d overdue)`)
    .join(" | ");

  const upcomingPreview = upcomingFiveDays
    .slice(0, 2)
    .map((record) => `${record.request_id} (${record.delta}d)`)
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
  const upcoming = state.records
    .filter((record) => isOpenStatus(record.status))
    .map((record) => ({
      ...record,
      dueInDays: daysUntil(record.due_date)
    }))
    .sort((a, b) => a.dueInDays - b.dueInDays)
    .slice(0, 8);

  if (!upcoming.length) {
    elements.upcomingDeadlines.innerHTML = "<li class=\"placeholder\">No upcoming open deadlines.</li>";
    return;
  }

  elements.upcomingDeadlines.innerHTML = upcoming.map((record) => {
    const timing = formatDueIn(record.dueInDays);
    return `
      <li class="deadline-item">
        <strong>${escapeHtml(record.request_id)} - ${escapeHtml(record.division)}</strong>
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
        <td>${escapeHtml(record.division)}</td>
        <td>${escapeHtml(record.requester)}</td>
        <td>${escapeHtml(record.subject)}</td>
        <td>${formatDate(record.received_date)}</td>
        <td>${formatDate(record.due_date)}</td>
        <td><span class="status-chip ${statusClass}">${escapeHtml(record.status)}</span></td>
        <td>${escapeHtml(record.assigned_to)}</td>
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
  const selected = state.records.find((record) => record.request_id === state.selectedId);

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
      <span class="detail-label">Division</span><span>${escapeHtml(selected.division)}</span>
      <span class="detail-label">Requester</span><span>${escapeHtml(selected.requester)}</span>
      <span class="detail-label">Subject</span><span>${escapeHtml(selected.subject)}</span>
      <span class="detail-label">Status</span><span>${escapeHtml(selected.status)}</span>
      <span class="detail-label">Received</span><span>${formatDate(selected.received_date)}</span>
      <span class="detail-label">Due Date</span><span>${formatDate(selected.due_date)} (${dueLabel})</span>
      <span class="detail-label">Assigned To</span><span>${escapeHtml(selected.assigned_to)}</span>
      <span class="detail-label">Last Update</span><span>${formatDate(selected.last_update)}</span>
      <span class="detail-label">Closed Date</span><span>${closedDate}</span>
      <span class="detail-label">Notes</span><span>${escapeHtml(selected.notes)}</span>
    </div>
  `;
}

function normalizeRecord(record) {
  return {
    ...record,
    status: normalizeStatus(record.status),
    received_date: parseDate(record.received_date),
    due_date: parseDate(record.due_date),
    last_update: parseDate(record.last_update),
    closed_date: record.closed_date ? parseDate(record.closed_date) : null
  };
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

function getDivisionCounts(records) {
  return records.reduce((acc, record) => {
    acc[record.division] = (acc[record.division] || 0) + 1;
    return acc;
  }, {});
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

function escapeHtml(value) {
  return String(value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/\"/g, "&quot;")
    .replace(/'/g, "&#39;");
}