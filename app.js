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
const GRAPH_ROOT = "https://graph.microsoft.com/v1.0";
const DEFAULT_GRAPH_SCOPES = ["User.Read", "Sites.Read.All"];
const LIVE_REFRESH_INTERVAL_MS = 5 * 60 * 1000;
const SHAREPOINT_TARGET = {
  siteUrl: "https://ilgov.sharepoint.com/teams/ISP.DCI.Commanders",
  listPathContains: "/Lists/DCI%20FOIA/",
  expectedDisplayName: "DCI FOIA"
};

const state = {
  records: [],
  availableWorkUnits: [],
  scopedRecords: [],
  filteredRecords: [],
  selectedId: null,
  selectedWorkUnit: ALL_WORK_UNITS_OPTION,
  searchQuery: "",
  diagnostics: {
    authStatus: "Not signed in",
    signedInUser: "N/A",
    siteResolved: false,
    listResolved: false,
    liveItemsLoaded: 0,
    dataSource: "None",
    lastUpdated: null
  }
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
  presentationToggle: document.getElementById("presentationToggle"),
  signInButton: document.getElementById("signInButton"),
  refreshNowButton: document.getElementById("refreshNowButton"),
  connectionStatusMessage: document.getElementById("connectionStatusMessage"),
  diagAuthStatus: document.getElementById("diagAuthStatus"),
  diagSignedInUser: document.getElementById("diagSignedInUser"),
  diagSiteResolved: document.getElementById("diagSiteResolved"),
  diagListResolved: document.getElementById("diagListResolved"),
  diagLiveItemsLoaded: document.getElementById("diagLiveItemsLoaded"),
  diagDataSource: document.getElementById("diagDataSource"),
  diagLastUpdated: document.getElementById("diagLastUpdated"),
  lastUpdatedValue: document.getElementById("lastUpdatedValue")
};

const today = startOfDay(new Date());
let msalInstance = null;
let liveRefreshTimerId = null;
let liveRefreshInProgress = false;
let signInHandlerAttached = false;

initialize();

async function initialize() {
  attachCoreEventListeners();
  setAuthControlsEnabled(false);
  renderDiagnostics();
  showStatusMessage("Sign in with Microsoft to load live SharePoint data.", "");
  setTableMessage("Sign in with Microsoft to load live SharePoint data.");

  console.info("Startup auth diagnostics", {
    graphConfigDetected: Boolean(window.GRAPH_CONFIG),
    msalLibraryDetected: Boolean(window.msal?.PublicClientApplication)
  });

  try {
    await initializeAuthentication();
    ensureSignInHandlerAttached();
    setAuthControlsEnabled(true);
    await trySilentLiveLoad();
  } catch (error) {
    console.error("MSAL initialization failed.", error);
    state.diagnostics.authStatus = "Initialization failed";
    renderDiagnostics();
    setAuthControlsEnabled(false);
    const reason = `Authentication setup failed: ${extractErrorMessage(error)}`;
    showStatusMessage(reason, "error");
    await loadDemoData(reason);
  }
}

function attachCoreEventListeners() {
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

  elements.refreshNowButton.addEventListener("click", async () => {
    await refreshLiveData({
      reason: "manual",
      showBusyState: true,
      fallbackToDemoOnFailure: false,
      keepCurrentDataOnFailure: true
    });
  });
}

function ensureSignInHandlerAttached() {
  if (signInHandlerAttached) {
    return;
  }
  elements.signInButton.addEventListener("click", async () => {
    await handleSignInClick();
  });
  signInHandlerAttached = true;
}

async function initializeAuthentication() {
  if (!window.msal || !window.msal.PublicClientApplication) {
    throw new Error("MSAL Browser failed to load.");
  }

  const config = window.GRAPH_CONFIG;
  if (!config?.clientId || !config?.tenantId || !config?.authority || !config?.redirectUri) {
    throw new Error("Microsoft Entra configuration is missing.");
  }

  if (msalInstance) {
    return;
  }

  const msalConfig = {
    auth: {
      clientId: config.clientId,
      authority: config.authority,
      redirectUri: config.redirectUri,
      postLogoutRedirectUri: config.redirectUri,
      navigateToLoginRequestUrl: false
    },
    cache: {
      cacheLocation: "sessionStorage",
      storeAuthStateInCookie: false
    }
  };

  msalInstance = new window.msal.PublicClientApplication(msalConfig);

  if (typeof msalInstance.initialize === "function") {
    await msalInstance.initialize();
  }

  await msalInstance.handleRedirectPromise();

  const accounts = msalInstance.getAllAccounts();
  if (accounts.length > 0) {
    msalInstance.setActiveAccount(accounts[0]);
  }
}

function setAuthControlsEnabled(enabled) {
  elements.signInButton.disabled = !enabled;
  elements.refreshNowButton.disabled = !enabled;
}

async function trySilentLiveLoad() {
  const account = getPreferredAccount();
  if (!account) {
    state.diagnostics.authStatus = "Ready to sign in";
    renderDiagnostics();
    return;
  }

  msalInstance.setActiveAccount(account);
  updateAuthDiagnostics(account, "Signed in");

  await refreshLiveData({
    reason: "startup",
    showBusyState: false,
    fallbackToDemoOnFailure: true,
    keepCurrentDataOnFailure: false
  });
  startLiveRefreshSchedule();
}

async function handleSignInClick() {
  if (!msalInstance) {
    showStatusMessage("Microsoft sign-in is unavailable because authentication initialization did not complete.", "error");
    return;
  }

  elements.signInButton.disabled = true;
  elements.signInButton.textContent = "Signing in...";

  try {
    const loginRequest = {
      scopes: getGraphScopes(),
      prompt: "select_account"
    };

    const loginResponse = await msalInstance.loginPopup(loginRequest);
    msalInstance.setActiveAccount(loginResponse.account);
    updateAuthDiagnostics(loginResponse.account, "Signed in");

    await refreshLiveData({
      reason: "signin",
      showBusyState: false,
      fallbackToDemoOnFailure: true,
      keepCurrentDataOnFailure: false
    });
    startLiveRefreshSchedule();
  } catch (error) {
    await handleLiveLoadFailure(error, {
      fallbackToDemo: true,
      keepCurrentData: false,
      nonIntrusive: false,
      reason: "signin"
    });
  } finally {
    elements.signInButton.disabled = false;
    elements.signInButton.textContent = "Sign in with Microsoft";
  }
}

async function refreshLiveData(options = {}) {
  const {
    reason = "scheduled",
    showBusyState = false,
    fallbackToDemoOnFailure = false,
    keepCurrentDataOnFailure = true
  } = options;

  if (liveRefreshInProgress) {
    return;
  }

  liveRefreshInProgress = true;

  if (showBusyState) {
    elements.refreshNowButton.disabled = true;
    elements.refreshNowButton.textContent = "Refreshing...";
  }

  try {
    await loadLiveSharePointData();
  } catch (error) {
    await handleLiveLoadFailure(error, {
      fallbackToDemo: fallbackToDemoOnFailure,
      keepCurrentData: keepCurrentDataOnFailure,
      nonIntrusive: reason === "scheduled" || reason === "manual",
      reason
    });
  } finally {
    liveRefreshInProgress = false;
    if (showBusyState) {
      elements.refreshNowButton.disabled = false;
      elements.refreshNowButton.textContent = "Refresh Now";
    }
  }
}

async function loadLiveSharePointData() {
  const token = await acquireGraphAccessToken();
  const site = await resolveSiteByUrl(token, SHAREPOINT_TARGET.siteUrl);
  state.diagnostics.siteResolved = true;

  const list = await resolveListByWebUrl(token, site.id, SHAREPOINT_TARGET);
  state.diagnostics.listResolved = true;

  const items = await fetchAllListItems(token, site.id, list.id);
  const fieldInternalNames = collectFieldNames(items);

  console.log("SharePoint live-data diagnostics", {
    siteId: site.id,
    listId: list.id,
    listDisplayName: list.displayName,
    listWebUrl: list.webUrl,
    itemCount: items.length,
    availableFieldInternalNames: fieldInternalNames
  });

  state.diagnostics.liveItemsLoaded = items.length;
  state.diagnostics.dataSource = "SharePoint";
  state.diagnostics.lastUpdated = new Date();
  showStatusMessage("Live SharePoint data connected", "live");
  renderDiagnostics();

  const graphRecords = items.map((item) => mapGraphItemToRecord(item));
  setDashboardRecords(graphRecords);
}

async function handleLiveLoadFailure(error, options = {}) {
  const {
    fallbackToDemo = false,
    keepCurrentData = true,
    nonIntrusive = true,
    reason = "scheduled"
  } = options;

  console.error("SharePoint authentication or loading failed.", error);
  const message = extractErrorMessage(error);
  const statusTone = nonIntrusive ? "warning" : "error";
  showStatusMessage(`Connection failed: ${message}`, statusTone);

  if (!keepCurrentData) {
    state.diagnostics.siteResolved = false;
    state.diagnostics.listResolved = false;
    state.diagnostics.liveItemsLoaded = 0;
  }

  if (!fallbackToDemo && keepCurrentData) {
    const failureLabel = reason === "manual" ? "Manual refresh failed" : "Automatic refresh failed";
    showStatusMessage(`${failureLabel}: ${message}. Current data remains visible.`, "warning");
  }

  renderDiagnostics();

  if (fallbackToDemo) {
    await loadDemoData(message);
  }
}

async function acquireGraphAccessToken() {
  const account = getPreferredAccount();
  if (!account) {
    throw new Error("No signed-in account. Use Sign in with Microsoft.");
  }

  try {
    const tokenResult = await msalInstance.acquireTokenSilent({
      account,
      scopes: getGraphScopes()
    });
    return tokenResult.accessToken;
  } catch (error) {
    if (!isInteractionRequired(error)) {
      throw error;
    }

    const tokenResult = await msalInstance.acquireTokenPopup({
      account,
      scopes: getGraphScopes()
    });
    return tokenResult.accessToken;
  }
}

async function graphFetch(token, url) {
  const response = await fetch(url, {
    headers: {
      Authorization: `Bearer ${token}`,
      Accept: "application/json"
    }
  });

  if (!response.ok) {
    const message = await response.text();
    throw new Error(`Graph request failed (${response.status}): ${message}`);
  }

  return response.json();
}

async function resolveSiteByUrl(token, siteUrl) {
  const parsed = new URL(siteUrl);
  const sitePath = parsed.pathname;
  const endpoint = `${GRAPH_ROOT}/sites/${parsed.hostname}:${sitePath}`;
  const site = await graphFetch(token, endpoint);
  console.log("Resolved SharePoint site", { siteId: site.id, siteWebUrl: site.webUrl || siteUrl });
  return site;
}

async function resolveListByWebUrl(token, siteId, target) {
  const endpoint = `${GRAPH_ROOT}/sites/${encodeURIComponent(siteId)}/lists?$select=id,displayName,webUrl`;
  const payload = await graphFetch(token, endpoint);
  const lists = payload.value || [];

  const normalizedTargetPath = target.listPathContains.toLowerCase();
  const normalizedDecodedTargetPath = decodeURIComponent(target.listPathContains).toLowerCase();

  const byWebUrl = lists.find((list) => {
    const normalizedWebUrl = String(list.webUrl || "").toLowerCase();
    return normalizedWebUrl.includes(normalizedTargetPath) || normalizedWebUrl.includes(normalizedDecodedTargetPath);
  });

  const byName = lists.find((list) => String(list.displayName || "").toLowerCase() === target.expectedDisplayName.toLowerCase());
  const resolved = byWebUrl || byName;

  if (!resolved) {
    throw new Error(`Unable to resolve list with path ${target.listPathContains} or name ${target.expectedDisplayName}.`);
  }

  console.log("Resolved SharePoint list", { id: resolved.id, displayName: resolved.displayName, webUrl: resolved.webUrl });

  return resolved;
}

async function fetchAllListItems(token, siteId, listId) {
  const encodedSiteId = encodeURIComponent(siteId);
  const encodedListId = encodeURIComponent(listId);
  let nextUrl = `${GRAPH_ROOT}/sites/${encodedSiteId}/lists/${encodedListId}/items?$expand=fields&$top=200`;
  const allItems = [];

  while (nextUrl) {
    const page = await graphFetch(token, nextUrl);
    allItems.push(...(page.value || []));
    nextUrl = page["@odata.nextLink"] || null;
  }

  return allItems;
}

function mapGraphItemToRecord(item) {
  const fields = item.fields || {};
  return {
    request_id: pickFirstField(fields, ["request_id", "requestId", "RequestID", "FOIANumber", "Title"]) || `ITEM-${item.id}`,
    dci_work_unit: pickFirstField(fields, ["dci_work_unit", "DCIWorkUnit", "WorkUnit"]),
    requester: pickFirstField(fields, ["requester", "Requester", "RequesterName"]),
    subject: pickFirstField(fields, ["subject", "Subject", "Title"]),
    received_date: pickFirstField(fields, ["received_date", "ReceivedDate", "DateReceived", "Created"]),
    due_date: pickFirstField(fields, ["due_date", "DueDate", "DateDue"]),
    status: pickFirstField(fields, ["status", "Status"]),
    assigned_to: pickFirstField(fields, ["assigned_to", "AssignedTo", "Owner"]),
    last_update: pickFirstField(fields, ["last_update", "LastUpdate", "Modified"]),
    closed_date: pickFirstField(fields, ["closed_date", "ClosedDate", "DateClosed"]),
    notes: pickFirstField(fields, ["notes", "Notes", "Comments"])
  };
}

function pickFirstField(fields, candidates) {
  for (const key of candidates) {
    const normalized = normalizeFieldValue(fields[key]);
    if (normalized) {
      return normalized;
    }
  }
  return "";
}

function normalizeFieldValue(value) {
  if (value === undefined || value === null) {
    return "";
  }
  if (Array.isArray(value)) {
    return value.map((item) => normalizeFieldValue(item)).filter(Boolean).join(", ");
  }
  if (typeof value === "object") {
    return value.displayName || value.email || value.title || value.LookupValue || value.Label || "";
  }
  return String(value).trim();
}

function collectFieldNames(items) {
  const names = new Set();
  items.forEach((item) => {
    Object.keys(item.fields || {}).forEach((key) => names.add(key));
  });
  return Array.from(names).sort((a, b) => a.localeCompare(b));
}

function getPreferredAccount() {
  if (!msalInstance) {
    return null;
  }
  return msalInstance.getActiveAccount() || msalInstance.getAllAccounts()[0] || null;
}

function getGraphScopes() {
  return Array.isArray(window.GRAPH_CONFIG?.scopes) && window.GRAPH_CONFIG.scopes.length
    ? window.GRAPH_CONFIG.scopes
    : DEFAULT_GRAPH_SCOPES;
}

function updateAuthDiagnostics(account, statusText) {
  state.diagnostics.authStatus = statusText;
  state.diagnostics.signedInUser = account?.username || account?.name || "N/A";
  renderDiagnostics();
}

function showStatusMessage(message, typeClass) {
  elements.connectionStatusMessage.textContent = message;
  elements.connectionStatusMessage.classList.remove("error", "demo", "live", "warning");
  if (typeClass) {
    elements.connectionStatusMessage.classList.add(typeClass);
  }
}

function renderDiagnostics() {
  elements.diagAuthStatus.textContent = state.diagnostics.authStatus;
  elements.diagSignedInUser.textContent = state.diagnostics.signedInUser;
  elements.diagSiteResolved.textContent = state.diagnostics.siteResolved ? "Yes" : "No";
  elements.diagListResolved.textContent = state.diagnostics.listResolved ? "Yes" : "No";
  elements.diagLiveItemsLoaded.textContent = String(state.diagnostics.liveItemsLoaded);
  elements.diagDataSource.textContent = state.diagnostics.dataSource;
  const lastUpdatedText = formatDateTime(state.diagnostics.lastUpdated);
  elements.diagLastUpdated.textContent = lastUpdatedText;
  elements.lastUpdatedValue.textContent = lastUpdatedText;
}

function setTableMessage(message) {
  elements.foiaTableBody.innerHTML = `<tr><td colspan="8">${escapeHtml(message)}</td></tr>`;
}

function setDashboardRecords(rawRecords) {
  const previousWorkUnit = state.selectedWorkUnit;
  const previousSelectedId = state.selectedId;

  state.records = rawRecords.map((record) => normalizeRecord(record));
  state.availableWorkUnits = buildAvailableWorkUnits(state.records);

  if (previousWorkUnit === ALL_WORK_UNITS_OPTION || state.availableWorkUnits.includes(previousWorkUnit)) {
    state.selectedWorkUnit = previousWorkUnit;
  } else {
    state.selectedWorkUnit = ALL_WORK_UNITS_OPTION;
  }

  state.selectedId = previousSelectedId;
  populateWorkUnitFilter();
  elements.workUnitFilter.value = state.selectedWorkUnit;
  applyFilters();
}

async function loadDemoData(reason) {
  try {
    const response = await fetch("foia-data.json", { cache: "no-store" });
    if (!response.ok) {
      throw new Error(`Data request failed with status ${response.status}`);
    }

    const rawRecords = await response.json();
    setDashboardRecords(rawRecords);
    state.diagnostics.dataSource = "Demo";
    state.diagnostics.liveItemsLoaded = 0;
    renderDiagnostics();
    showStatusMessage(`Demo data currently displayed. ${reason}`, "demo");
  } catch (fallbackError) {
    state.records = [];
    state.availableWorkUnits = [];
    populateWorkUnitFilter();
    applyFilters();
    showStatusMessage(`Failed to load demo data: ${extractErrorMessage(fallbackError)}`, "error");
    console.error("Demo fallback loading failed.", fallbackError);
  }
}

function extractErrorMessage(error) {
  if (!error) {
    return "Unknown error";
  }
  if (typeof error === "string") {
    return error;
  }
  if (error.message) {
    return error.message;
  }
  return "Unknown error";
}

function isInteractionRequired(error) {
  if (!error) {
    return false;
  }
  if (window.msal?.InteractionRequiredAuthError && error instanceof window.msal.InteractionRequiredAuthError) {
    return true;
  }
  const code = String(error.errorCode || "").toLowerCase();
  return code.includes("interaction") || code.includes("consent") || code.includes("login");
}

function startLiveRefreshSchedule() {
  if (liveRefreshTimerId) {
    return;
  }

  liveRefreshTimerId = window.setInterval(async () => {
    await refreshLiveData({
      reason: "scheduled",
      showBusyState: false,
      fallbackToDemoOnFailure: false,
      keepCurrentDataOnFailure: true
    });
  }, LIVE_REFRESH_INTERVAL_MS);
}

function formatDateTime(value) {
  if (!(value instanceof Date) || Number.isNaN(value.getTime())) {
    return "Never";
  }
  return value.toLocaleString("en-US", {
    month: "short",
    day: "2-digit",
    year: "numeric",
    hour: "2-digit",
    minute: "2-digit",
    second: "2-digit"
  });
}

function buildAvailableWorkUnits(records) {
  const fromData = new Set(
    records
      .map((record) => String(record.dci_work_unit || "").trim())
      .filter((value) => value)
  );

  const orderedKnown = DCI_WORK_UNITS.filter((unit) => fromData.has(unit));
  const extras = Array.from(fromData)
    .filter((unit) => !DCI_WORK_UNITS.includes(unit))
    .sort((a, b) => a.localeCompare(b));

  return [...orderedKnown, ...extras];
}

function populateWorkUnitFilter() {
  elements.workUnitFilter.innerHTML = [
    `<option value="${ALL_WORK_UNITS_OPTION}">All DCI Work Units</option>`,
    ...state.availableWorkUnits.map((workUnit) => `<option value="${escapeHtml(workUnit)}">${escapeHtml(workUnit)}</option>`)
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
    ? state.availableWorkUnits
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
  return normalized || "Unknown";
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
  if (!value) {
    return null;
  }
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) {
    return null;
  }
  return startOfDay(date);
}

function startOfDay(date) {
  if (!(date instanceof Date) || Number.isNaN(date.getTime())) {
    return null;
  }
  const copy = new Date(date);
  copy.setHours(0, 0, 0, 0);
  return copy;
}

function daysBetween(first, second) {
  if (!(first instanceof Date) || Number.isNaN(first.getTime()) || !(second instanceof Date) || Number.isNaN(second.getTime())) {
    return 0;
  }
  const millisecondsPerDay = 86400000;
  return Math.round((startOfDay(second) - startOfDay(first)) / millisecondsPerDay);
}

function daysUntil(date) {
  if (!(date instanceof Date) || Number.isNaN(date.getTime())) {
    return Number.POSITIVE_INFINITY;
  }
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
  if (!Number.isFinite(days)) {
    return "Due date unavailable";
  }
  if (days < 0) {
    return `${Math.abs(days)} day${Math.abs(days) === 1 ? "" : "s"} overdue`;
  }
  if (days === 0) {
    return "Due today";
  }
  return `Due in ${days} day${days === 1 ? "" : "s"}`;
}

function formatDueInShort(days) {
  if (!Number.isFinite(days)) {
    return "N/A";
  }
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
