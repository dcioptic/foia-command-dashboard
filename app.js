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
const SHAREPOINT_HOSTNAME = "ilgov.sharepoint.com";
const SHAREPOINT_SITE_PATH = "/teams/ISP.DCI.Commanders";
const SHAREPOINT_FOIA_LIST_PATH = "/Lists/DCI%20FOIA/";
const SHAREPOINT_FOIA_LIST_NAME_FALLBACKS = ["DCI FOIA", "DCI FOIA TRACKER", "DCI FOIA Tracker"];

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
const devWarningFlags = new Set();

initialize();

async function initialize() {
  attachCoreEventListeners();
  setAuthControlsEnabled(false);
  renderDiagnostics();
  showStatusMessage("Demo data currently displayed. Sign in with Microsoft to load live SharePoint data.", "demo");
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
    await loadDemoData("Sign in with Microsoft to load live SharePoint data.");
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
    await loadDemoData(message);
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
    showStatusMessage(`DEMO FALLBACK: Demo data currently displayed. ${reason}`, "demo");
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
  const normalized = String(status || "").trim().toLowerCase();
  if (!normalized) {
    return "Unknown";
  }
  if (normalized === "closed" || normalized === "complete" || normalized === "completed") {
    return "Closed";
  }
  if (normalized === "in review") {
    return "In Review";
  }
  if (normalized === "open" || normalized === "pending" || normalized === "active") {
    return "Open";
  }
  return String(status).trim();
}

function isOpenStatus(status) {
  const normalized = String(status || "").trim().toLowerCase();
  return normalized === "open" || normalized === "in review" || normalized === "pending" || normalized === "active";
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
