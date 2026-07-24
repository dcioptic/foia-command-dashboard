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
const ALL_STATUSES_OPTION = "ALL_STATUSES";
const ALL_TIME_OPTION = "ALL_TIME";
const STATUS_FILTER_ALL_LABEL = "All Statuses";
const FALLBACK_STATUS_CHOICES = [
  "1. NEW",
  "2. IN PROGRESS",
  "3. PENDING WITH LEGAL",
  "4. WORK UNIT RESPONDED",
  "5. DCI COMPLETED"
];
const KPI_STATUS_ORDER = [...FALLBACK_STATUS_CHOICES];
const OPEN_STAGE_STATUSES = new Set([
  "1. NEW",
  "2. IN PROGRESS",
  "3. PENDING WITH LEGAL",
  "4. WORK UNIT RESPONDED"
]);
const COMPLETED_STATUS = "5. DCI COMPLETED";
const GRAPH_ROOT = "https://graph.microsoft.com/v1.0";
const DEFAULT_GRAPH_SCOPES = ["User.Read", "Sites.Read.All"];
const LIVE_REFRESH_INTERVAL_MS = 5 * 60 * 1000;
const SHAREPOINT_HOSTNAME = "ilgov.sharepoint.com";
const SHAREPOINT_SITE_PATH = "/teams/ISP.DCI.Commanders";
const SHAREPOINT_FOIA_LIST_PATH = "/Lists/DCI%20FOIA/";
const SHAREPOINT_FOIA_LIST_NAME_FALLBACKS = ["DCI FOIA", "DCI FOIA TRACKER", "DCI FOIA Tracker"];

const state = {
  records: [],
  availableWorkUnits: [],
  baseFilteredRecords: [],
  scopedRecords: [],
  filteredRecords: [],
  selectedId: null,
  selectedWorkUnit: ALL_WORK_UNITS_OPTION,
  selectedStatus: ALL_STATUSES_OPTION,
  availableStatuses: [...FALLBACK_STATUS_CHOICES],
  selectedTimePeriod: ALL_TIME_OPTION,
  customStartDate: "",
  customEndDate: "",
  searchQuery: "",
  diagnostics: {
    authStatus: "Not signed in",
    signedInUser: "N/A",
    siteResolved: false,
    listResolved: false,
    liveItemsLoaded: 0,
    dataSource: "None",
    lastUpdated: null
  },
  liveConnection: {
    site: {
      id: "",
      displayName: "",
      webUrl: ""
    },
    list: {
      id: "",
      displayName: "",
      webUrl: ""
    },
    fieldMap: {}
  }
};

const elements = {
  kpiStatusNew: document.getElementById("kpiStatusNew"),
  kpiStatusInProgress: document.getElementById("kpiStatusInProgress"),
  kpiStatusPendingLegal: document.getElementById("kpiStatusPendingLegal"),
  kpiStatusUnitResponded: document.getElementById("kpiStatusUnitResponded"),
  kpiStatusCompleted: document.getElementById("kpiStatusCompleted"),
  kpiDue10: document.getElementById("kpiDue10"),
  kpiDue5: document.getElementById("kpiDue5"),
  kpiDueToday: document.getElementById("kpiDueToday"),
  kpiOverdue: document.getElementById("kpiOverdue"),
  kpiAvgDaysInProgress: document.getElementById("kpiAvgDaysInProgress"),
  kpiReceivedThisMonth: document.getElementById("kpiReceivedThisMonth"),
  kpiCompletedThisMonth: document.getElementById("kpiCompletedThisMonth"),
  kpiScopeSummary: document.getElementById("kpiScopeSummary"),
  workUnitSummaryBody: document.getElementById("workUnitSummaryBody"),
  trendTableBody: document.getElementById("trendTableBody"),
  upcomingDeadlines: document.getElementById("upcomingDeadlines"),
  commandAlerts: document.getElementById("commandAlerts"),
  foiaTableBody: document.getElementById("foiaTableBody"),
  detailContent: document.getElementById("detailContent"),
  searchInput: document.getElementById("searchInput"),
  workUnitFilter: document.getElementById("workUnitFilter"),
  statusFilter: document.getElementById("statusFilter"),
  timePeriodFilter: document.getElementById("timePeriodFilter"),
  customStartDate: document.getElementById("customStartDate"),
  customEndDate: document.getElementById("customEndDate"),
  clearFiltersButton: document.getElementById("clearFiltersButton"),
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
const devWarningFlags = new Set();

initialize();

async function initialize() {
  hydrateInitialFiltersFromUrl();
  attachCoreEventListeners();
  setAuthControlsEnabled(false);
  renderDiagnostics();
  showStatusMessage("Sign in with Microsoft to load live SharePoint data.", "warning");
  setTableMessage("Sign in with Microsoft to load live SharePoint data.");

  console.log("MSAL Browser loaded:", Boolean(window.msal?.PublicClientApplication));

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
    setDashboardRecords([]);
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

  elements.statusFilter.addEventListener("change", (event) => {
    state.selectedStatus = event.target.value;
    applyFilters();
  });

  elements.timePeriodFilter.addEventListener("change", (event) => {
    state.selectedTimePeriod = event.target.value;
    applyFilters();
  });

  elements.customStartDate.addEventListener("change", (event) => {
    state.customStartDate = event.target.value || "";
    applyFilters();
  });

  elements.customEndDate.addEventListener("change", (event) => {
    state.customEndDate = event.target.value || "";
    applyFilters();
  });

  elements.clearFiltersButton.addEventListener("click", () => {
    resetFiltersToDefault();
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

function hydrateInitialFiltersFromUrl() {
  const params = new URLSearchParams(window.location.search);
  const statusFromUrl = String(params.get("status") || "").trim();
  if (!statusFromUrl || statusFromUrl === STATUS_FILTER_ALL_LABEL) {
    state.selectedStatus = ALL_STATUSES_OPTION;
    return;
  }
  state.selectedStatus = statusFromUrl;
}

function resetFiltersToDefault() {
  state.selectedWorkUnit = ALL_WORK_UNITS_OPTION;
  state.selectedStatus = ALL_STATUSES_OPTION;
  state.selectedTimePeriod = ALL_TIME_OPTION;
  state.customStartDate = "";
  state.customEndDate = "";
  state.searchQuery = "";

  elements.workUnitFilter.value = state.selectedWorkUnit;
  elements.statusFilter.value = state.selectedStatus;
  elements.timePeriodFilter.value = state.selectedTimePeriod;
  elements.customStartDate.value = "";
  elements.customEndDate.value = "";
  elements.searchInput.value = "";
}

function syncStatusQueryString() {
  const params = new URLSearchParams(window.location.search);
  const statusValue = state.selectedStatus === ALL_STATUSES_OPTION
    ? STATUS_FILTER_ALL_LABEL
    : state.selectedStatus;
  params.set("status", statusValue);
  const nextSearch = params.toString();
  const nextUrl = `${window.location.pathname}${nextSearch ? `?${nextSearch}` : ""}${window.location.hash}`;
  window.history.replaceState({}, "", nextUrl);
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
    showStatusMessage("Sign in with Microsoft to load live SharePoint data.", "warning");
    setDashboardRecords([]);
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
      fallbackToDemoOnFailure: false,
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

  const site = await resolveSharePointSite(token);
  state.liveConnection.site.id = site.id || "";
  state.liveConnection.site.displayName = site.displayName || "";
  state.liveConnection.site.webUrl = site.webUrl || "";
  state.diagnostics.siteResolved = true;
  renderDiagnostics();

  const list = await resolveFoiaList(token, site.id);
  state.liveConnection.list.id = list.id || "";
  state.liveConnection.list.displayName = list.displayName || "";
  state.liveConnection.list.webUrl = list.webUrl || "";
  state.diagnostics.listResolved = true;
  renderDiagnostics();

  const columns = await fetchAllListColumns(token, site.id, list.id);
  const fieldMap = buildSharePointFieldMap(columns);
  state.liveConnection.fieldMap = fieldMap;
  state.availableStatuses = deriveStatusChoices(columns, fieldMap.status);
  renderFieldMappingDiagnostics(fieldMap);

  const items = await fetchAllListItems(token, site.id, list.id);
  console.log("Number of list items returned:", items.length);
  items.slice(0, 3).forEach((item) => {
    console.log("SharePoint field keys:", Object.keys(item?.fields || {}));
  });

  const graphRecords = items.map((item) => mapGraphItemToRecord(item, fieldMap));
  logFieldMappingValidationSummary(items, graphRecords);

  state.diagnostics.liveItemsLoaded = items.length;
  state.diagnostics.dataSource = "SharePoint";
  state.diagnostics.lastUpdated = new Date();
  showStatusMessage("Live SharePoint data connected", "live");
  renderDiagnostics();

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
  const message = getLiveLoadFailureMessage(error);
  const statusTone = nonIntrusive ? "warning" : "error";
  showStatusMessage(message, statusTone);

  if (!keepCurrentData) {
    if (error?.stage === "site") {
      state.diagnostics.siteResolved = false;
      state.diagnostics.listResolved = false;
      state.liveConnection.site.id = "";
      state.liveConnection.site.displayName = "";
      state.liveConnection.site.webUrl = "";
    } else if (error?.stage === "list") {
      state.diagnostics.siteResolved = true;
      state.diagnostics.listResolved = false;
    } else {
      state.diagnostics.siteResolved = false;
      state.diagnostics.listResolved = false;
      state.liveConnection.site.id = "";
      state.liveConnection.site.displayName = "";
      state.liveConnection.site.webUrl = "";
    }
    state.diagnostics.liveItemsLoaded = 0;
    state.liveConnection.list.id = "";
    state.liveConnection.list.displayName = "";
    state.liveConnection.list.webUrl = "";
  }

  if (!fallbackToDemo && keepCurrentData) {
    const failureLabel = reason === "manual" ? "Manual refresh failed" : "Automatic refresh failed";
    showStatusMessage(`${failureLabel}: ${message} Current data remains visible.`, "warning");
  }

  renderDiagnostics();

  if (fallbackToDemo) {
    showStatusMessage(`${message} Live data remains required; demo data is disabled.`, "error");
    if (!keepCurrentData) {
      setDashboardRecords([]);
    }
  }
}

async function acquireGraphAccessToken() {
  const account = getPreferredAccount();
  if (!account) {
    throw new Error("No signed-in account. Use Sign in with Microsoft.");
  }

  const scopes = getGraphScopes();
  console.log("Graph token scopes requested:", scopes);

  try {
    const tokenResult = await msalInstance.acquireTokenSilent({
      account,
      scopes
    });
    return tokenResult.accessToken;
  } catch (error) {
    if (!isInteractionRequired(error)) {
      throw error;
    }

    const tokenResult = await msalInstance.acquireTokenPopup({
      account,
      scopes
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

  const status = response.status;
  const rawText = await response.text();

  if (!response.ok) {
    let errorPayload = null;
    if (rawText) {
      try {
        errorPayload = JSON.parse(rawText);
      } catch {
        errorPayload = null;
      }
    }

    const graphError = errorPayload?.error || rawText || "Unknown Graph error";
    console.error("Graph request failed", {
      url,
      status,
      error: graphError
    });
    const error = new Error(`Graph request failed (${status})`);
    error.graphStatus = status;
    error.graphError = graphError;
    throw error;
  }

  let payload = {};
  if (rawText) {
    try {
      payload = JSON.parse(rawText);
    } catch {
      throw new Error(`Graph response JSON parsing failed (${status})`);
    }
  }

  return {
    status,
    data: payload
  };
}

async function resolveSharePointSite(token) {
  const siteEndpoint = `${GRAPH_ROOT}/sites/${SHAREPOINT_HOSTNAME}:${encodeURI(SHAREPOINT_SITE_PATH)}?$select=id,displayName,webUrl`;
  console.log("Graph site endpoint requested:", siteEndpoint);

  try {
    const response = await graphFetch(token, siteEndpoint);
    console.log("Site response HTTP status:", response.status);

    const site = response.data || {};
    console.log("Resolved site:", {
      id: site.id || "",
      displayName: site.displayName || "",
      webUrl: site.webUrl || ""
    });

    if (!site.id) {
      throw new Error("SharePoint site response was missing a site id.");
    }

    return site;
  } catch (error) {
    throw createStageError(
      "site",
      "Signed in, but the ISP.DCI.Commanders SharePoint site could not be resolved.",
      error
    );
  }
}

async function resolveFoiaList(token, siteId) {
  const encodedSiteId = encodeURIComponent(siteId);
  let nextUrl = `${GRAPH_ROOT}/sites/${encodedSiteId}/lists?$select=id,displayName,name,webUrl`;
  const allLists = [];

  try {
    while (nextUrl) {
      const response = await graphFetch(token, nextUrl);
      allLists.push(...(response.data?.value || []));
      nextUrl = response.data?.["@odata.nextLink"] || null;
    }

    console.log("Number of lists returned:", allLists.length);
    console.table(
      allLists.map((list) => ({
        id: list.id || "",
        displayName: list.displayName || "",
        name: list.name || "",
        webUrl: list.webUrl || ""
      }))
    );

    const normalizedPathTarget = normalizeListUrl(SHAREPOINT_FOIA_LIST_PATH);
    const resolvedByUrl = allLists.find((list) => normalizeListUrl(list.webUrl).includes(normalizedPathTarget));

    const fallbackNames = new Set(SHAREPOINT_FOIA_LIST_NAME_FALLBACKS.map((name) => name.toLowerCase()));
    const resolvedByName = allLists.find((list) => {
      const displayName = String(list.displayName || "").trim().toLowerCase();
      return fallbackNames.has(displayName);
    });

    const resolved = resolvedByUrl || resolvedByName;

    if (!resolved) {
      throw new Error("Unable to locate FOIA list by URL or fallback display names.");
    }

    console.log("Resolved list:", {
      id: resolved.id || "",
      displayName: resolved.displayName || "",
      webUrl: resolved.webUrl || ""
    });

    return resolved;
  } catch (error) {
    throw createStageError(
      "list",
      "SharePoint site connected, but the DCI FOIA list could not be located.",
      error
    );
  }
}

async function fetchAllListItems(token, siteId, listId) {
  const encodedSiteId = encodeURIComponent(siteId);
  const encodedListId = encodeURIComponent(listId);
  let nextUrl = `${GRAPH_ROOT}/sites/${encodedSiteId}/lists/${encodedListId}/items?$expand=fields`;
  const allItems = [];

  while (nextUrl) {
    const page = await graphFetch(token, nextUrl);
    allItems.push(...(page.data?.value || []));
    nextUrl = page.data?.["@odata.nextLink"] || null;
  }

  return allItems;
}

function normalizeListUrl(value) {
  if (!value) {
    return "";
  }

  let normalized = String(value).trim().toLowerCase();
  try {
    normalized = decodeURIComponent(normalized);
  } catch {
    // Keep original normalized text when URL decoding is not possible.
  }

  const queryIndex = normalized.indexOf("?");
  if (queryIndex >= 0) {
    normalized = normalized.slice(0, queryIndex);
  }

  normalized = normalized.replace(/\/+$/, "");
  return normalized;
}

function createStageError(stage, message, cause) {
  const error = new Error(message);
  error.stage = stage;
  error.cause = cause;
  return error;
}

function getLiveLoadFailureMessage(error) {
  if (error?.stage === "site") {
    return "Signed in, but the ISP.DCI.Commanders SharePoint site could not be resolved.";
  }
  if (error?.stage === "list") {
    return "SharePoint site connected, but the DCI FOIA list could not be located.";
  }
  return `Connection failed: ${extractErrorMessage(error)}`;
}

function mapGraphItemToRecord(item, fieldMap) {
  const fields = item?.fields || {};
  const foiaNumber = readMappedText(fields, fieldMap.foiaNumber) || `ITEM-${item.id}`;
  const statusFromField = readMappedChoice(fields, fieldMap.status);
  const status = statusFromField || "Unknown";

  if (!statusFromField) {
    warnOnce("status-mapping-missing-value", "Status field mapping returned no value for one or more records. Status mapping may need attention.");
  }

  const workUnit = readMappedChoice(fields, fieldMap.workUnit);
  const subjectValue = readMappedText(fields, fieldMap.subjectName);
  const subject = subjectValue || foiaNumber;

  const createdDate = readMappedDate(fields, fieldMap.created);
  const dateStamped = readMappedDate(fields, fieldMap.dateStamped);
  const receivedDate = readMappedDate(fields, fieldMap.dateReceived) || createdDate;
  const fiveDaysOut = readMappedDate(fields, fieldMap.fiveDaysOut);
  const tenDaysOut = readMappedDate(fields, fieldMap.tenDaysOut);
  const dueDate = determineDueDate({
    explicitDueDate: readMappedDate(fields, fieldMap.dueDate),
    fiveDaysOut,
    tenDaysOut
  });

  const assignedTo = readMappedPerson(fields, fieldMap.assignedTo, item.id);
  const response = readMappedChoice(fields, fieldMap.response);
  const requester = readMappedText(fields, fieldMap.division) || response;

  const normalizedRecord = {
    id: String(item.id || ""),
    foiaNumber,
    status,
    foiaType: readMappedChoice(fields, fieldMap.foiaType),
    receivedDate,
    dateStamped,
    createdDate,
    fiveDaysOut,
    tenDaysOut,
    response,
    workUnit,
    assignedTo,
    subject,
    crimepadCase: readMappedText(fields, fieldMap.crimepadCase),
    tracsCase: readMappedText(fields, fieldMap.tracsCase),
    otherCase: readMappedText(fields, fieldMap.otherCase),
    modifiedDate: readMappedDate(fields, fieldMap.modified),
    dueDate,
    closedDate: readMappedDate(fields, fieldMap.closedDate),
    notes: readMappedText(fields, fieldMap.notes)
  };

  return {
    request_id: normalizedRecord.foiaNumber,
    dci_work_unit: normalizedRecord.workUnit,
    requester,
    subject: normalizedRecord.subject,
    received_date: normalizedRecord.receivedDate,
    due_date: normalizedRecord.dueDate,
    status: normalizedRecord.status,
    assigned_to: normalizedRecord.assignedTo,
    last_update: normalizedRecord.modifiedDate,
    closed_date: normalizedRecord.closedDate,
    notes: normalizedRecord.notes,
    foia_type: normalizedRecord.foiaType,
    response: normalizedRecord.response,
    created_date: normalizedRecord.createdDate,
    date_stamped: normalizedRecord.dateStamped,
    five_days_out: normalizedRecord.fiveDaysOut,
    ten_days_out: normalizedRecord.tenDaysOut,
    crimepad_case: normalizedRecord.crimepadCase,
    tracs_case: normalizedRecord.tracsCase,
    other_case: normalizedRecord.otherCase
  };
}

function normalizeColumnLabel(value) {
  return String(value || "")
    .toLowerCase()
    .replace(/[^a-z0-9]/g, "");
}

function detectColumnType(column) {
  if (column?.choice) {
    return "choice";
  }
  if (column?.dateTime) {
    return "dateTime";
  }
  if (column?.personOrGroup) {
    return "personOrGroup";
  }
  if (column?.lookup) {
    return "lookup";
  }
  if (column?.boolean) {
    return "boolean";
  }
  if (column?.number) {
    return "number";
  }
  if (column?.text) {
    return "text";
  }
  return "unknown";
}

async function fetchAllListColumns(token, siteId, listId) {
  const encodedSiteId = encodeURIComponent(siteId);
  const encodedListId = encodeURIComponent(listId);
  let nextUrl = `${GRAPH_ROOT}/sites/${encodedSiteId}/lists/${encodedListId}/columns?$select=id,name,displayName,description,hidden,readOnly,required,text,choice,dateTime,personOrGroup,lookup,boolean,number`;
  const allColumns = [];

  while (nextUrl) {
    const page = await graphFetch(token, nextUrl);
    allColumns.push(...(page.data?.value || []));
    nextUrl = page.data?.["@odata.nextLink"] || null;
  }

  console.table(
    allColumns.map((column) => ({
      displayName: column.displayName || "",
      name: column.name || "",
      id: column.id || "",
      hidden: Boolean(column.hidden),
      readOnly: Boolean(column.readOnly),
      detectedType: detectColumnType(column)
    }))
  );

  return allColumns;
}

function deriveStatusChoices(columns, statusInternalName) {
  if (!statusInternalName) {
    return [...FALLBACK_STATUS_CHOICES];
  }

  const statusColumn = columns.find((column) => column.name === statusInternalName);
  const liveChoices = Array.isArray(statusColumn?.choice?.choices)
    ? statusColumn.choice.choices.map((choice) => String(choice || "").trim()).filter(Boolean)
    : [];

  if (!liveChoices.length) {
    return [...FALLBACK_STATUS_CHOICES];
  }

  return liveChoices;
}

function findInternalName(variants, internalNameByDisplayName, fallbackByInternalName) {
  for (const variant of variants) {
    const normalized = normalizeColumnLabel(variant);
    if (!normalized) {
      continue;
    }

    const byDisplayName = internalNameByDisplayName.get(normalized);
    if (byDisplayName) {
      return byDisplayName;
    }

    const byInternalName = fallbackByInternalName.get(normalized);
    if (byInternalName) {
      return byInternalName;
    }
  }

  console.warn("Column mapping not found for variants:", variants.join(" | "));
  return "";
}

function buildSharePointFieldMap(columns) {
  const internalNameByDisplayName = new Map(
    columns
      .filter((column) => column.displayName && column.name)
      .map((column) => [normalizeColumnLabel(column.displayName), column.name])
  );

  const fallbackByInternalName = new Map(
    columns
      .filter((column) => column.name)
      .map((column) => [normalizeColumnLabel(column.name), column.name])
  );

  const fieldMap = {
    foiaNumber: findInternalName(["Title"], internalNameByDisplayName, fallbackByInternalName),
    status: findInternalName(["STATUS", "Status"], internalNameByDisplayName, fallbackByInternalName),
    division: findInternalName(["DIVISION", "Division"], internalNameByDisplayName, fallbackByInternalName),
    foirType: findInternalName(["FOIR TYPE", "FOIR Type"], internalNameByDisplayName, fallbackByInternalName),
    foiaType: findInternalName(["FOIA TYPE", "FOIR TYPE", "FOIA Type"], internalNameByDisplayName, fallbackByInternalName),
    attachments: findInternalName(["Attachments"], internalNameByDisplayName, fallbackByInternalName),
    dateStamped: findInternalName(["DATE STAMPED", "Date Stamped"], internalNameByDisplayName, fallbackByInternalName),
    dateReceived: findInternalName(["DATE DCI RECEIVED", "Date DCI Received"], internalNameByDisplayName, fallbackByInternalName),
    created: findInternalName(["Created"], internalNameByDisplayName, fallbackByInternalName) || "Created",
    fiveDaysOut: findInternalName(["5 Days Out", "Five Days Out"], internalNameByDisplayName, fallbackByInternalName),
    tenDaysOut: findInternalName(["10 Days Out", "Ten Days Out"], internalNameByDisplayName, fallbackByInternalName),
    response: findInternalName(["RESPONSE", "Response"], internalNameByDisplayName, fallbackByInternalName),
    workUnit: findInternalName(["DCI WORK UNIT", "DCI Work Unit"], internalNameByDisplayName, fallbackByInternalName),
    assignedTo: findInternalName(["Assigned to", "Assigned To"], internalNameByDisplayName, fallbackByInternalName),
    subjectName: findInternalName(["SUBJECT NAME", "Subject Name"], internalNameByDisplayName, fallbackByInternalName),
    crimepadCase: findInternalName(["CRIMEPAD CASE #", "CRIMEPAD CASE"], internalNameByDisplayName, fallbackByInternalName),
    tracsCase: findInternalName(["TRACS CASE #", "TRACS CASE"], internalNameByDisplayName, fallbackByInternalName),
    otherCase: findInternalName(["OTHER CASE #", "OTHER CASE"], internalNameByDisplayName, fallbackByInternalName),
    modified: findInternalName(["Modified"], internalNameByDisplayName, fallbackByInternalName) || "Modified",
    dueDate: findInternalName(["DUE DATE", "DATE DUE", "DueDate", "due_date"], internalNameByDisplayName, fallbackByInternalName),
    closedDate: findInternalName(["CLOSED DATE", "DATE CLOSED", "Closed Date"], internalNameByDisplayName, fallbackByInternalName),
    notes: findInternalName(["NOTES", "Notes", "COMMENTS", "Comments"], internalNameByDisplayName, fallbackByInternalName)
  };

  const mappingTable = [
    ["Title", fieldMap.foiaNumber],
    ["STATUS", fieldMap.status],
    ["DIVISION", fieldMap.division],
    ["FOIR TYPE", fieldMap.foirType],
    ["FOIA TYPE", fieldMap.foiaType],
    ["Attachments", fieldMap.attachments],
    ["DATE STAMPED", fieldMap.dateStamped],
    ["DATE DCI RECEIVED", fieldMap.dateReceived],
    ["Created", fieldMap.created],
    ["5 Days Out", fieldMap.fiveDaysOut],
    ["10 Days Out", fieldMap.tenDaysOut],
    ["RESPONSE", fieldMap.response],
    ["DCI WORK UNIT", fieldMap.workUnit],
    ["Assigned to", fieldMap.assignedTo],
    ["SUBJECT NAME", fieldMap.subjectName],
    ["CRIMEPAD CASE #", fieldMap.crimepadCase],
    ["TRACS CASE #", fieldMap.tracsCase],
    ["OTHER CASE #", fieldMap.otherCase],
    ["Modified", fieldMap.modified]
  ];

  console.table(
    mappingTable.map(([visibleName, internalName]) => ({
      visibleName,
      internalName: internalName || "(unmapped)"
    }))
  );

  return fieldMap;
}

function renderFieldMappingDiagnostics(fieldMap) {
  const diagnosticsSection = document.querySelector(".connection-diagnostics");
  if (!diagnosticsSection) {
    return;
  }

  let mappingBlock = document.getElementById("diagFieldMappings");
  if (!mappingBlock) {
    mappingBlock = document.createElement("div");
    mappingBlock.id = "diagFieldMappings";
    mappingBlock.className = "diag-field-mappings";
    diagnosticsSection.appendChild(mappingBlock);
  }

  const lines = [
    ["Title", fieldMap.foiaNumber],
    ["STATUS", fieldMap.status],
    ["FOIA TYPE", fieldMap.foiaType],
    ["DATE DCI RECEIVED", fieldMap.dateReceived],
    ["5 Days Out", fieldMap.fiveDaysOut],
    ["10 Days Out", fieldMap.tenDaysOut],
    ["RESPONSE", fieldMap.response],
    ["DCI WORK UNIT", fieldMap.workUnit],
    ["Assigned to", fieldMap.assignedTo],
    ["SUBJECT NAME", fieldMap.subjectName],
    ["CRIMEPAD CASE #", fieldMap.crimepadCase],
    ["TRACS CASE #", fieldMap.tracsCase],
    ["OTHER CASE #", fieldMap.otherCase],
    ["Modified", fieldMap.modified]
  ]
    .map(([visibleName, internalName]) => `${escapeHtml(visibleName)} -> ${escapeHtml(internalName || "(unmapped)")}`)
    .join("<br>");

  mappingBlock.innerHTML = `<h4>Temporary Field Mapping Diagnostics</h4><div>${lines}</div>`;
}

function readMappedValue(fields, internalName) {
  if (!internalName) {
    return undefined;
  }
  return fields?.[internalName];
}

function normalizeGraphFieldText(value) {
  if (value === undefined || value === null) {
    return "";
  }
  if (Array.isArray(value)) {
    return value.map((entry) => normalizeGraphFieldText(entry)).filter(Boolean).join(", ");
  }
  if (typeof value === "object") {
    return String(
      value.displayName ||
      value.LookupValue ||
      value.email ||
      value.title ||
      value.Label ||
      value.value ||
      ""
    ).trim();
  }
  return String(value).trim();
}

function readMappedText(fields, internalName) {
  return normalizeGraphFieldText(readMappedValue(fields, internalName));
}

function readMappedChoice(fields, internalName) {
  const value = readMappedValue(fields, internalName);
  if (Array.isArray(value)) {
    return value.map((entry) => normalizeGraphFieldText(entry)).filter(Boolean).join(", ");
  }
  return normalizeGraphFieldText(value);
}

function readMappedDate(fields, internalName) {
  const value = readMappedValue(fields, internalName);
  const normalizedText = normalizeGraphFieldText(value);
  if (!normalizedText) {
    return "";
  }
  const parsed = new Date(normalizedText);
  if (Number.isNaN(parsed.getTime())) {
    return "";
  }
  return parsed.toISOString();
}

function normalizePersonEntry(value) {
  if (value === undefined || value === null) {
    return "";
  }
  if (typeof value === "string") {
    return value.trim();
  }
  if (typeof value === "number") {
    return String(value);
  }
  if (typeof value === "object") {
    const text = value.displayName || value.email || value.LookupValue || value.title || "";
    if (text) {
      return String(text).trim();
    }

    const lookupId = value.id || value.lookupId || value.LookupId || value.Id;
    if (lookupId !== undefined && lookupId !== null) {
      return String(lookupId);
    }
  }
  return "";
}

function readMappedPerson(fields, internalName, itemId) {
  const value = readMappedValue(fields, internalName);
  if (value === undefined || value === null) {
    return "";
  }

  if (Array.isArray(value)) {
    const entries = value.map((entry) => normalizePersonEntry(entry)).filter(Boolean);
    if (entries.length) {
      return entries.join(", ");
    }
    warnOnce("assigned-to-array-requires-expansion", "Assigned to field returned array values that require person expansion.", {
      itemId,
      field: internalName
    });
    return "";
  }

  if (typeof value === "object") {
    const normalized = normalizePersonEntry(value);
    if (normalized) {
      if (/^\d+$/.test(normalized)) {
        warnOnce("assigned-to-lookup-id", "Assigned to value appears to be a lookup ID. Person expansion is required.", {
          itemId,
          field: internalName,
          lookupId: normalized
        });
      }
      return normalized;
    }

    warnOnce("assigned-to-object-requires-expansion", "Assigned to field returned object data that requires person expansion.", {
      itemId,
      field: internalName
    });
    return "";
  }

  if (typeof value === "number") {
    warnOnce("assigned-to-numeric-lookup-id", "Assigned to field returned a numeric lookup ID. Person expansion is required.", {
      itemId,
      field: internalName,
      lookupId: value
    });
    return String(value);
  }

  return String(value).trim();
}

function determineDueDate(values) {
  if (values.explicitDueDate) {
    return values.explicitDueDate;
  }

  if (values.fiveDaysOut && values.tenDaysOut) {
    const five = new Date(values.fiveDaysOut);
    const ten = new Date(values.tenDaysOut);
    if (!Number.isNaN(five.getTime()) && !Number.isNaN(ten.getTime()) && ten > five) {
      warnOnce(
        "due-date-inferred-from-thresholds",
        "Due date is being inferred from 10 Days Out and 5 Days Out values. Confirm this behavior matches your SharePoint deadline semantics."
      );
      return ten.toISOString();
    }
  }

  return "";
}

function warnOnce(key, message, details) {
  if (devWarningFlags.has(key)) {
    return;
  }
  devWarningFlags.add(key);
  if (details) {
    console.warn(message, details);
    return;
  }
  console.warn(message);
}

function logFieldMappingValidationSummary(items, records) {
  const summary = {
    totalItemsRetrieved: items.length,
    recordsWithMappedWorkUnits: records.filter((record) => String(record.dci_work_unit || "").trim() && record.dci_work_unit !== "Unknown").length,
    recordsWithMappedStatuses: records.filter((record) => String(record.status || "").trim() && record.status !== "Unknown").length,
    recordsWithReceivedDates: records.filter((record) => Boolean(record.received_date)).length,
    recordsWithFiveDaysOutValues: records.filter((record) => Boolean(record.five_days_out)).length,
    recordsWithTenDaysOutValues: records.filter((record) => Boolean(record.ten_days_out)).length,
    recordsWithSubjectNames: records.filter((record) => String(record.subject || "").trim()).length,
    recordsWithAssignedUsers: records.filter((record) => String(record.assigned_to || "").trim()).length
  };

  console.log("SharePoint mapping validation summary", summary);
}

function getPreferredAccount() {
  if (!msalInstance) {
    return null;
  }
  return msalInstance.getActiveAccount() || msalInstance.getAllAccounts()[0] || null;
}

function getGraphScopes() {
  const requestedScopes = Array.isArray(window.GRAPH_CONFIG?.scopes) && window.GRAPH_CONFIG.scopes.length
    ? window.GRAPH_CONFIG.scopes
    : DEFAULT_GRAPH_SCOPES;

  const scopeSet = new Set(requestedScopes);
  scopeSet.add("User.Read");
  scopeSet.add("Sites.Read.All");
  return Array.from(scopeSet);
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
  elements.foiaTableBody.innerHTML = `<tr><td colspan="9">${escapeHtml(message)}</td></tr>`;
}

function setDashboardRecords(rawRecords) {
  const previousWorkUnit = state.selectedWorkUnit;
  const previousStatus = state.selectedStatus;
  const previousSelectedId = state.selectedId;

  state.records = rawRecords.map((record) => normalizeRecord(record));
  state.availableWorkUnits = buildAvailableWorkUnits(state.records);
  state.availableStatuses = buildAvailableStatuses(state.records, state.availableStatuses);

  if (previousWorkUnit === ALL_WORK_UNITS_OPTION || state.availableWorkUnits.includes(previousWorkUnit)) {
    state.selectedWorkUnit = previousWorkUnit;
  } else {
    state.selectedWorkUnit = ALL_WORK_UNITS_OPTION;
  }

  if (previousStatus === ALL_STATUSES_OPTION || state.availableStatuses.some((status) => statusEquals(status, previousStatus))) {
    state.selectedStatus = resolveStatusOptionValue(previousStatus, state.availableStatuses);
  } else {
    state.selectedStatus = ALL_STATUSES_OPTION;
  }

  state.selectedId = previousSelectedId;
  populateWorkUnitFilter();
  populateStatusFilter();
  elements.workUnitFilter.value = state.selectedWorkUnit;
  elements.statusFilter.value = state.selectedStatus;
  elements.timePeriodFilter.value = state.selectedTimePeriod;
  elements.customStartDate.value = state.customStartDate;
  elements.customEndDate.value = state.customEndDate;
  elements.searchInput.value = state.searchQuery;
  applyFilters();
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

function buildAvailableStatuses(records, baselineChoices) {
  const preferredOrder = Array.isArray(baselineChoices) && baselineChoices.length
    ? baselineChoices
    : FALLBACK_STATUS_CHOICES;

  const foundStatuses = new Set(
    records
      .map((record) => String(record.status || "").trim())
      .filter(Boolean)
  );

  const orderedKnown = [...preferredOrder];
  const extras = Array.from(foundStatuses)
    .filter((status) => !preferredOrder.includes(status))
    .sort((a, b) => a.localeCompare(b));

  return [...orderedKnown, ...extras];
}

function resolveStatusOptionValue(value, availableStatuses) {
  if (value === ALL_STATUSES_OPTION) {
    return ALL_STATUSES_OPTION;
  }
  const matched = availableStatuses.find((status) => statusEquals(status, value));
  return matched || ALL_STATUSES_OPTION;
}

function populateWorkUnitFilter() {
  elements.workUnitFilter.innerHTML = [
    `<option value="${ALL_WORK_UNITS_OPTION}">All DCI Work Units</option>`,
    ...state.availableWorkUnits.map((workUnit) => `<option value="${escapeHtml(workUnit)}">${escapeHtml(workUnit)}</option>`)
  ].join("");
}

function populateStatusFilter() {
  elements.statusFilter.innerHTML = [
    `<option value="${ALL_STATUSES_OPTION}">${STATUS_FILTER_ALL_LABEL}</option>`,
    ...state.availableStatuses.map((status) => `<option value="${escapeHtml(status)}">${escapeHtml(status)}</option>`)
  ].join("");
}

function applyFilters() {
  const normalizedQuery = state.searchQuery.trim().toLowerCase();

  const periodRange = getTimePeriodRange(state.selectedTimePeriod);
  const customRange = getCustomDateRange(state.customStartDate, state.customEndDate);

  state.baseFilteredRecords = state.records.filter((record) => {
    if (state.selectedWorkUnit !== ALL_WORK_UNITS_OPTION && record.dci_work_unit !== state.selectedWorkUnit) {
      return false;
    }

    if (state.selectedStatus !== ALL_STATUSES_OPTION && !statusEquals(record.status, state.selectedStatus)) {
      return false;
    }

    if (!recordMatchesDateRange(record, periodRange.start, periodRange.end)) {
      return false;
    }

    if (!recordMatchesDateRange(record, customRange.start, customRange.end)) {
      return false;
    }

    return true;
  });

  state.scopedRecords = [...state.baseFilteredRecords];

  state.filteredRecords = state.baseFilteredRecords.filter((record) => {
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

  syncStatusQueryString();

  renderAll();
}

function renderAll() {
  renderKpis();
  renderWorkUnitSummary();
  renderTrendPanel();
  renderUpcomingDeadlines();
  renderCommandAlerts();
  renderRequestTable();
  renderDetailPanel();
}

function renderKpis() {
  const records = state.scopedRecords;
  const now = new Date();

  const statusCounts = KPI_STATUS_ORDER.map((statusLabel) => countByStatus(records, statusLabel));
  const activeRecords = records.filter((record) => isOpenStageStatus(record.status));

  const dueToday = activeRecords.filter((record) => hasValidDueDate(record) && daysUntil(record.due_date) === 0);
  const due5 = activeRecords.filter((record) => {
    if (!hasValidDueDate(record)) {
      return false;
    }
    const delta = daysUntil(record.due_date);
    return delta >= 1 && delta <= 5;
  });
  const due10 = activeRecords.filter((record) => {
    if (!hasValidDueDate(record)) {
      return false;
    }
    const delta = daysUntil(record.due_date);
    return delta >= 6 && delta <= 10;
  });
  const overdue = activeRecords.filter((record) => hasValidDueDate(record) && daysUntil(record.due_date) < 0);

  const inProgressDurations = activeRecords
    .map((record) => daysInProgress(record))
    .filter((days) => Number.isFinite(days));
  const avgDaysInProgress = inProgressDurations.length
    ? Math.round(inProgressDurations.reduce((sum, days) => sum + days, 0) / inProgressDurations.length)
    : 0;

  const monthScopeRecords = getCurrentMonthMetricScopeRecords();
  const receivedThisMonth = monthScopeRecords.filter((record) => isWithinMonth(getIntakeDate(record), now)).length;
  const completedThisMonth = monthScopeRecords.filter((record) => isCompletedStatus(record.status) && isWithinMonth(record.closed_date, now)).length;

  elements.kpiStatusNew.textContent = String(statusCounts[0]);
  elements.kpiStatusInProgress.textContent = String(statusCounts[1]);
  elements.kpiStatusPendingLegal.textContent = String(statusCounts[2]);
  elements.kpiStatusUnitResponded.textContent = String(statusCounts[3]);
  elements.kpiStatusCompleted.textContent = String(statusCounts[4]);
  elements.kpiDue10.textContent = String(state.selectedStatus === COMPLETED_STATUS ? 0 : due10.length);
  elements.kpiDue5.textContent = String(state.selectedStatus === COMPLETED_STATUS ? 0 : due5.length);
  elements.kpiDueToday.textContent = String(state.selectedStatus === COMPLETED_STATUS ? 0 : dueToday.length);
  elements.kpiOverdue.textContent = String(state.selectedStatus === COMPLETED_STATUS ? 0 : overdue.length);
  elements.kpiAvgDaysInProgress.textContent = String(avgDaysInProgress);
  elements.kpiReceivedThisMonth.textContent = String(receivedThisMonth);
  elements.kpiCompletedThisMonth.textContent = String(completedThisMonth);

  elements.kpiScopeSummary.textContent = buildFilterSummaryText();
}

function renderWorkUnitSummary() {
  const statusColumns = [...KPI_STATUS_ORDER];
  const records = state.scopedRecords;
  const monthReference = new Date();
  const monthScopeRecords = getCurrentMonthMetricScopeRecords();
  const unitsInScope = state.selectedWorkUnit === ALL_WORK_UNITS_OPTION
    ? Array.from(new Set(records.map((record) => record.dci_work_unit))).sort((a, b) => a.localeCompare(b))
    : [state.selectedWorkUnit];

  const rows = unitsInScope
    .map((workUnit) => {
      const unitRecords = records.filter((record) => record.dci_work_unit === workUnit);
      const statusCounts = statusColumns.map((status) => unitRecords.filter((record) => statusEquals(record.status, status)).length);
      const activeRecords = unitRecords.filter((record) => isOpenStageStatus(record.status));
      const due10 = activeRecords.filter((record) => {
        if (!hasValidDueDate(record)) {
          return false;
        }
        const delta = daysUntil(record.due_date);
        return delta >= 6 && delta <= 10;
      }).length;
      const due5 = activeRecords.filter((record) => {
        if (!hasValidDueDate(record)) {
          return false;
        }
        const delta = daysUntil(record.due_date);
        return delta >= 1 && delta <= 5;
      }).length;
      const dueToday = activeRecords.filter((record) => hasValidDueDate(record) && daysUntil(record.due_date) === 0).length;
      const overdue = activeRecords.filter((record) => hasValidDueDate(record) && daysUntil(record.due_date) < 0).length;
      const durations = activeRecords.map((record) => daysInProgress(record)).filter((days) => Number.isFinite(days));
      const avgDays = durations.length ? Math.round(durations.reduce((sum, days) => sum + days, 0) / durations.length) : 0;

      const monthUnitScope = monthScopeRecords.filter((record) => record.dci_work_unit === workUnit);
      const receivedThisMonth = monthUnitScope.filter((record) => isWithinMonth(getIntakeDate(record), monthReference)).length;
      const completedThisMonth = monthUnitScope.filter((record) => isCompletedStatus(record.status) && isWithinMonth(record.closed_date, monthReference)).length;

      const total = statusCounts.reduce((sum, count) => sum + count, 0);
      return {
        workUnit,
        statusCounts,
        due10,
        due5,
        dueToday,
        overdue,
        avgDays,
        receivedThisMonth,
        completedThisMonth,
        total
      };
    })
    .filter((row) => row.total > 0)
    .sort((a, b) => b.total - a.total || a.workUnit.localeCompare(b.workUnit));

  if (!rows.length) {
    elements.workUnitSummaryBody.innerHTML = "<tr><td colspan=\"14\">No DCI work unit workload data available.</td></tr>";
    return;
  }

  elements.workUnitSummaryBody.innerHTML = rows.map((item) => `
    <tr>
      <td>${escapeHtml(item.workUnit)}</td>
      <td title="1. NEW">${item.statusCounts[0]}</td>
      <td title="2. IN PROGRESS">${item.statusCounts[1]}</td>
      <td title="3. PENDING WITH LEGAL">${item.statusCounts[2]}</td>
      <td title="4. WORK UNIT RESPONDED">${item.statusCounts[3]}</td>
      <td title="5. DCI COMPLETED">${item.statusCounts[4]}</td>
      <td>${item.due10}</td>
      <td>${item.due5}</td>
      <td>${item.dueToday}</td>
      <td>${item.overdue}</td>
      <td>${item.avgDays}</td>
      <td>${item.receivedThisMonth}</td>
      <td>${item.completedThisMonth}</td>
      <td>${item.total}</td>
    </tr>
  `).join("");
}

function renderTrendPanel() {
  const trendRows = buildTrendRows();
  if (!trendRows.length) {
    elements.trendTableBody.innerHTML = "<tr><td colspan=\"3\">No trend data available for current filters.</td></tr>";
    return;
  }

  elements.trendTableBody.innerHTML = trendRows.map((row) => `
    <tr>
      <td>${escapeHtml(row.label)}</td>
      <td>${row.received}</td>
      <td>${row.completed}</td>
    </tr>
  `).join("");
}

function renderCommandAlerts() {
  const records = state.scopedRecords;
  const activeRecords = records.filter((record) => isOpenStageStatus(record.status));
  const monthScopeRecords = getCurrentMonthMetricScopeRecords();

  const due10 = activeRecords.filter((record) => hasValidDueDate(record) && daysUntil(record.due_date) >= 6 && daysUntil(record.due_date) <= 10).length;
  const due5 = activeRecords.filter((record) => hasValidDueDate(record) && daysUntil(record.due_date) >= 1 && daysUntil(record.due_date) <= 5).length;
  const dueToday = activeRecords.filter((record) => hasValidDueDate(record) && daysUntil(record.due_date) === 0).length;
  const overdue = activeRecords.filter((record) => hasValidDueDate(record) && daysUntil(record.due_date) < 0).length;
  const respondedWorkUnits = new Set(
    records
      .filter((record) => statusEquals(record.status, "4. WORK UNIT RESPONDED"))
      .map((record) => record.dci_work_unit)
      .filter(Boolean)
  );
  const completedThisMonth = monthScopeRecords.filter((record) => isCompletedStatus(record.status) && isWithinMonth(record.closed_date, new Date())).length;

  elements.commandAlerts.innerHTML = `
    <li class="alert-item"><div class="alert-title">New Requests Awaiting Action</div><div class="alert-value">${countByStatus(records, "1. NEW")}</div></li>
    <li class="alert-item"><div class="alert-title">Requests Currently In Progress</div><div class="alert-value">${countByStatus(records, "2. IN PROGRESS")}</div></li>
    <li class="alert-item"><div class="alert-title">Requests Pending With Legal</div><div class="alert-value">${countByStatus(records, "3. PENDING WITH LEGAL")}</div></li>
    <li class="alert-item"><div class="alert-title">Work Units That Have Responded</div><div class="alert-value">${respondedWorkUnits.size}</div></li>
    <li class="alert-item"><div class="alert-title">Due in 10 Days</div><div class="alert-value">${state.selectedStatus === COMPLETED_STATUS ? 0 : due10}</div></li>
    <li class="alert-item"><div class="alert-title">Due in 5 Days</div><div class="alert-value">${state.selectedStatus === COMPLETED_STATUS ? 0 : due5}</div></li>
    <li class="alert-item"><div class="alert-title">Due Today</div><div class="alert-value">${state.selectedStatus === COMPLETED_STATUS ? 0 : dueToday}</div></li>
    <li class="alert-item"><div class="alert-title">Overdue</div><div class="alert-value">${state.selectedStatus === COMPLETED_STATUS ? 0 : overdue}</div></li>
    <li class="alert-item"><div class="alert-title">DCI Completed This Month</div><div class="alert-value">${completedThisMonth}</div></li>
  `;
}

function renderUpcomingDeadlines() {
  const upcoming = state.scopedRecords
    .filter((record) => isOpenStageStatus(record.status) && hasValidDueDate(record))
    .map((record) => ({
      ...record,
      dueInDays: daysUntil(record.due_date)
    }))
    .sort((a, b) => a.dueInDays - b.dueInDays);

  if (!upcoming.length) {
    elements.upcomingDeadlines.innerHTML = "<li class=\"placeholder\">No active deadlines for current filters.</li>";
    return;
  }

  elements.upcomingDeadlines.innerHTML = upcoming.map((record) => {
    const timing = record.dueInDays < 0
      ? `${Math.abs(record.dueInDays)} day${Math.abs(record.dueInDays) === 1 ? "" : "s"} overdue`
      : `${record.dueInDays} day${record.dueInDays === 1 ? "" : "s"} remaining`;
    return `
      <li class="deadline-item">
        <strong>${escapeHtml(record.request_id)}</strong>
        <span>${escapeHtml(record.subject)}</span>
        <span>${escapeHtml(record.dci_work_unit)} | ${escapeHtml(record.status)}</span>
        <div class="deadline-meta">
          <span>${formatDate(record.due_date)}</span>
          <span>${escapeHtml(timing)}</span>
        </div>
      </li>
    `;
  }).join("");
}

function renderRequestTable() {
  if (!state.filteredRecords.length) {
    elements.foiaTableBody.innerHTML = "<tr><td colspan=\"9\">No FOIA records match your search.</td></tr>";
    return;
  }

  elements.foiaTableBody.innerHTML = state.filteredRecords.map((record) => {
    const selectedClass = record.request_id === state.selectedId ? "selected" : "";
    const statusClass = `status-${statusClassSuffix(record.status)}`;
    const completed = isCompletedStatus(record.status);
    const daysOpen = daysInProgress(record);
    const daysClose = daysToClose(record);
    const dueDelta = daysUntil(record.due_date);
    const timingText = completed
      ? formatNumericMetric(daysClose)
      : `${formatNumericMetric(daysOpen)}${Number.isFinite(dueDelta) ? ` (${formatDueInShort(dueDelta)})` : ""}`;

    return `
      <tr data-request-id="${escapeHtml(record.request_id)}" class="${selectedClass}" tabindex="0">
        <td>${escapeHtml(record.request_id)}</td>
        <td>${escapeHtml(record.subject)}</td>
        <td>${escapeHtml(record.requester)}</td>
        <td>${escapeHtml(record.dci_work_unit)}</td>
        <td><span class="status-chip ${statusClass}">${escapeHtml(record.status)}</span></td>
        <td>${record.received_date ? formatDate(record.received_date) : ""}</td>
        <td>${completed ? "" : (record.due_date ? formatDate(record.due_date) : "")}</td>
        <td>${completed && record.closed_date ? formatDate(record.closed_date) : ""}</td>
        <td>${timingText}</td>
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
  const selected = state.filteredRecords.find((record) => record.request_id === state.selectedId);

  if (!selected) {
    elements.detailContent.innerHTML = "<p class=\"placeholder\">Select a FOIA request to view detailed information.</p>";
    return;
  }

  const completed = isCompletedStatus(selected.status);
  const dueInDays = daysUntil(selected.due_date);
  const closeDays = daysToClose(selected);
  const details = [
    ["FOIA Number", selected.request_id],
    ["Subject", selected.subject],
    ["Requester", selected.requester],
    ["DCI Work Unit", selected.dci_work_unit],
    ["Exact Status", selected.status]
  ];

  if (selected.received_date) {
    details.push(["Date DCI Received", formatDate(selected.received_date)]);
  }

  if (!completed && selected.due_date) {
    details.push(["Due Date", `${formatDate(selected.due_date)} (${formatDueIn(dueInDays)})`]);
  }
  if (completed && selected.closed_date) {
    details.push(["Date Closed", formatDate(selected.closed_date)]);
  }

  if (!completed) {
    details.push(["Days In Progress", formatNumericMetric(daysInProgress(selected))]);
  } else {
    details.push(["Days To Complete", formatNumericMetric(closeDays)]);
  }

  if (selected.assigned_to) {
    details.push(["Assigned To", selected.assigned_to]);
  }
  if (selected.response) {
    details.push(["Response", selected.response]);
  }
  if (selected.crimepad_case) {
    details.push(["CRIMEPAD Case #", selected.crimepad_case]);
  }
  if (selected.tracs_case) {
    details.push(["TRACS Case #", selected.tracs_case]);
  }
  if (selected.other_case) {
    details.push(["Other Case #", selected.other_case]);
  }
  if (selected.notes) {
    details.push(["Notes", selected.notes]);
  }
  if (selected.last_update) {
    details.push(["Last Update", formatDate(selected.last_update)]);
  }

  elements.detailContent.innerHTML = `
    <h3>${escapeHtml(selected.request_id)}</h3>
    <div class="detail-grid">
      ${details.map(([label, value]) => `<span class="detail-label">${escapeHtml(label)}</span><span>${escapeHtml(value)}</span>`).join("")}
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
    date_stamped: parseDate(record.date_stamped),
    created_date: parseDate(record.created_date),
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
  const value = String(status || "").trim();
  const normalized = value.toLowerCase();
  if (!value) {
    return "Unknown";
  }

  // Keep live SharePoint labels exact; map legacy demo labels into canonical STATUS choices.
  if (normalized === "closed" || normalized === "complete" || normalized === "completed") {
    return COMPLETED_STATUS;
  }
  if (normalized === "in review") {
    return "3. PENDING WITH LEGAL";
  }
  if (normalized === "open" || normalized === "pending" || normalized === "active") {
    return "2. IN PROGRESS";
  }
  return value;
}

function statusKey(value) {
  return String(value || "").trim().toLowerCase().replace(/\s+/g, " ");
}

function statusEquals(a, b) {
  return statusKey(a) === statusKey(b);
}

function isOpenStageStatus(status) {
  return Array.from(OPEN_STAGE_STATUSES).some((knownStatus) => statusEquals(status, knownStatus));
}

function isCompletedStatus(status) {
  return statusEquals(status, COMPLETED_STATUS);
}

function getIntakeDate(record) {
  return record.received_date || record.date_stamped || record.created_date || null;
}

function getProgressStartDate(record) {
  return record.received_date || record.date_stamped || record.created_date || null;
}

function getRecordTimeFilterDate(record) {
  if (isCompletedStatus(record.status)) {
    return record.closed_date || null;
  }
  if (isOpenStageStatus(record.status)) {
    return record.received_date || null;
  }
  return null;
}

function getTimePeriodRange(period) {
  if (!period || period === ALL_TIME_OPTION) {
    return { start: null, end: null };
  }

  const end = today;
  if (!end) {
    return { start: null, end: null };
  }

  if (period === "THIS_MONTH") {
    const start = new Date(end.getFullYear(), end.getMonth(), 1);
    return { start: startOfDay(start), end };
  }

  if (period === "LAST_30_DAYS") {
    return { start: addDays(end, -29), end };
  }

  if (period === "LAST_90_DAYS") {
    return { start: addDays(end, -89), end };
  }

  if (period === "THIS_YEAR") {
    const start = new Date(end.getFullYear(), 0, 1);
    return { start: startOfDay(start), end };
  }

  return { start: null, end: null };
}

function getCustomDateRange(startValue, endValue) {
  const start = parseDate(startValue);
  const end = parseDate(endValue);
  if (start && end && start > end) {
    return { start: end, end: start };
  }
  return { start, end };
}

function recordMatchesDateRange(record, start, end) {
  if (!start && !end) {
    return true;
  }
  const date = getRecordTimeFilterDate(record);
  if (!(date instanceof Date) || Number.isNaN(date.getTime())) {
    return false;
  }
  if (start && date < start) {
    return false;
  }
  if (end && date > end) {
    return false;
  }
  return true;
}

function addDays(date, amount) {
  const next = new Date(date);
  next.setDate(next.getDate() + amount);
  return startOfDay(next);
}

function daysToClose(record) {
  const opened = getProgressStartDate(record);
  const closed = record.closed_date;
  if (!(opened instanceof Date) || Number.isNaN(opened.getTime())) {
    return Number.NaN;
  }
  if (!(closed instanceof Date) || Number.isNaN(closed.getTime())) {
    return Number.NaN;
  }
  return daysBetween(opened, closed);
}

function isWithinMonth(date, referenceDate) {
  if (!(date instanceof Date) || Number.isNaN(date.getTime())) {
    return false;
  }
  if (!(referenceDate instanceof Date) || Number.isNaN(referenceDate.getTime())) {
    return false;
  }
  return date.getMonth() === referenceDate.getMonth() && date.getFullYear() === referenceDate.getFullYear();
}

function statusClassSuffix(status) {
  return String(status || "")
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, "-")
    .replace(/^-+|-+$/g, "") || "unknown";
}

function formatNumericMetric(value) {
  return Number.isFinite(value) ? String(value) : "N/A";
}

function countByStatus(records, statusLabel) {
  return records.filter((record) => {
    const statusText = String(record.status || "").trim();
    if (!statusText || statusKey(statusText) === "unknown") {
      return false;
    }
    return statusEquals(statusText, statusLabel);
  }).length;
}

function hasValidDueDate(record) {
  return record.due_date instanceof Date && !Number.isNaN(record.due_date.getTime());
}

function daysInProgress(record) {
  const startDate = getProgressStartDate(record);
  if (!(startDate instanceof Date) || Number.isNaN(startDate.getTime())) {
    return Number.NaN;
  }
  return daysBetween(startDate, today);
}

function getCurrentMonthMetricScopeRecords() {
  return state.records.filter((record) => {
    if (state.selectedWorkUnit !== ALL_WORK_UNITS_OPTION && record.dci_work_unit !== state.selectedWorkUnit) {
      return false;
    }
    if (state.selectedStatus !== ALL_STATUSES_OPTION && !statusEquals(record.status, state.selectedStatus)) {
      return false;
    }
    return true;
  });
}

function getTimePeriodLabel(value) {
  if (value === "THIS_MONTH") {
    return "the current calendar month";
  }
  if (value === "LAST_30_DAYS") {
    return "the last 30 days";
  }
  if (value === "LAST_90_DAYS") {
    return "the last 90 days";
  }
  if (value === "THIS_YEAR") {
    return "the current calendar year";
  }
  return "all time";
}

function buildFilterSummaryText() {
  const statusLabel = state.selectedStatus === ALL_STATUSES_OPTION ? "all statuses" : state.selectedStatus;
  const workUnitLabel = state.selectedWorkUnit === ALL_WORK_UNITS_OPTION ? "all DCI work units" : state.selectedWorkUnit;

  let periodLabel = getTimePeriodLabel(state.selectedTimePeriod);
  if (state.customStartDate || state.customEndDate) {
    const start = state.customStartDate || "earliest";
    const end = state.customEndDate || "today";
    periodLabel = `${start} to ${end}`;
  }

  return `Showing ${statusLabel} requests for ${workUnitLabel} during ${periodLabel}.`;
}

function buildTrendRows() {
  const periodRange = getTimePeriodRange(state.selectedTimePeriod);
  const customRange = getCustomDateRange(state.customStartDate, state.customEndDate);

  const baseRecords = state.records.filter((record) => {
    if (state.selectedWorkUnit !== ALL_WORK_UNITS_OPTION && record.dci_work_unit !== state.selectedWorkUnit) {
      return false;
    }
    if (state.selectedStatus !== ALL_STATUSES_OPTION && !statusEquals(record.status, state.selectedStatus)) {
      return false;
    }
    return true;
  });

  const buckets = new Map();
  const upsertBucket = (date) => {
    const key = `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}`;
    if (!buckets.has(key)) {
      buckets.set(key, {
        key,
        label: date.toLocaleDateString("en-US", { month: "short", year: "numeric" }),
        received: 0,
        completed: 0
      });
    }
    return buckets.get(key);
  };

  baseRecords.forEach((record) => {
    const receivedDate = record.received_date;
    if (receivedDate instanceof Date && !Number.isNaN(receivedDate.getTime())) {
      if (recordMatchesDateRangeByDate(receivedDate, periodRange.start, periodRange.end) && recordMatchesDateRangeByDate(receivedDate, customRange.start, customRange.end)) {
        upsertBucket(receivedDate).received += 1;
      }
    }

    if (isCompletedStatus(record.status) && record.closed_date instanceof Date && !Number.isNaN(record.closed_date.getTime())) {
      if (recordMatchesDateRangeByDate(record.closed_date, periodRange.start, periodRange.end) && recordMatchesDateRangeByDate(record.closed_date, customRange.start, customRange.end)) {
        upsertBucket(record.closed_date).completed += 1;
      }
    }
  });

  return Array.from(buckets.values()).sort((a, b) => a.key.localeCompare(b.key));
}

function recordMatchesDateRangeByDate(date, start, end) {
  if (!(date instanceof Date) || Number.isNaN(date.getTime())) {
    return false;
  }
  if (start && date < start) {
    return false;
  }
  if (end && date > end) {
    return false;
  }
  return true;
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
