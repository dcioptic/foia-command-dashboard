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

const OFFICIAL_WORK_UNIT_BY_KEY = new Map(
  DCI_WORK_UNITS.map((unit) => [normalizeWorkUnitKey(unit), unit])
);

const ALL_WORK_UNITS_OPTION = "ALL";
const ALL_STATUSES_OPTION = "ALL_STATUSES";
const ALL_TIME_OPTION = "ALL_TIME";
const LOCAL_RECEIVED_SORT_NEWEST = "RECEIVED_NEWEST";
const LOCAL_RECEIVED_SORT_OLDEST = "RECEIVED_OLDEST";
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
const PENDING_WITH_LEGAL_STATUS = "3. PENDING WITH LEGAL";
const COMPLETED_STATUS = "5. DCI COMPLETED";
const PENDING_WITH_LEGAL_FALLBACK_TO_MODIFIED = Boolean(window.GRAPH_CONFIG?.pendingWithLegalDateFallbackToModified);
const GRAPH_ROOT = "https://graph.microsoft.com/v1.0";
const DEFAULT_GRAPH_SCOPES = ["User.Read", "Sites.Read.All"];
const LIVE_REFRESH_INTERVAL_MS = 5 * 60 * 1000;
const SHAREPOINT_LOAD_TIMEOUT_MS = 45 * 1000;
const SHAREPOINT_HOSTNAME = "ilgov.sharepoint.com";
const SHAREPOINT_SITE_PATH = "/teams/ISP.DCI.Commanders";
const SHAREPOINT_FOIA_LIST_PATH = "/Lists/DCI%20FOIA/";
const SHAREPOINT_FOIA_LIST_NAME_FALLBACKS = ["DCI FOIA", "DCI FOIA TRACKER", "DCI FOIA Tracker"];
const AUTH_REDIRECT_PENDING_KEY = "authenticationRedirectPending";
const GRAPH_RETRYABLE_STATUS = new Set([429, 500, 502, 503, 504]);
const GRAPH_MAX_RETRIES = 2;

const state = {
  records: [],
  availableWorkUnits: [],
  baseFilteredRecords: [],
  scopedRecords: [],
  filteredRecords: [],
  selectedId: null,
  tableSelectedWorkUnit: ALL_WORK_UNITS_OPTION,
  tableSelectedStatus: ALL_STATUSES_OPTION,
  tableReceivedSort: LOCAL_RECEIVED_SORT_NEWEST,
  tableAvailableWorkUnits: [],
  tableAvailableStatuses: [],
  selectedWorkUnit: ALL_WORK_UNITS_OPTION,
  selectedStatus: ALL_STATUSES_OPTION,
  availableStatuses: [...FALLBACK_STATUS_CHOICES],
  selectedTimePeriod: ALL_TIME_OPTION,
  customStartDate: "",
  customEndDate: "",
  searchQuery: "",
  appView: "login",
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
  authShell: document.getElementById("authShell"),
  dashboardShell: document.getElementById("dashboardShell"),
  authStateHeading: document.getElementById("authStateHeading"),
  authPrimaryMessage: document.getElementById("authPrimaryMessage"),
  authSecondaryMessage: document.getElementById("authSecondaryMessage"),
  authSpinner: document.getElementById("authSpinner"),
  loginSignInButton: document.getElementById("loginSignInButton"),
  authSecondaryButton: document.getElementById("authSecondaryButton"),
  kpiStatusNew: document.getElementById("kpiStatusNew"),
  kpiStatusInProgress: document.getElementById("kpiStatusInProgress"),
  kpiStatusPendingLegal: document.getElementById("kpiStatusPendingLegal"),
  kpiAvgDaysToReceive: document.getElementById("kpiAvgDaysToReceive"),
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
  receivedCompletedPanel: document.getElementById("receivedCompletedPanel"),
  upcomingDeadlinesPanel: document.getElementById("upcomingDeadlinesPanel"),
  commandAlertsPanel: document.getElementById("commandAlertsPanel"),
  upcomingDeadlines: document.getElementById("upcomingDeadlines"),
  commandAlerts: document.getElementById("commandAlerts"),
  foiaTableBody: document.getElementById("foiaTableBody"),
  detailContent: document.getElementById("detailContent"),
  searchInput: document.getElementById("searchInput"),
  tableWorkUnitFilter: document.getElementById("tableWorkUnitFilter"),
  tableStatusFilter: document.getElementById("tableStatusFilter"),
  tableReceivedSort: document.getElementById("tableReceivedSort"),
  workUnitFilter: document.getElementById("workUnitFilter"),
  statusFilter: document.getElementById("statusFilter"),
  timePeriodFilter: document.getElementById("timePeriodFilter"),
  customStartDate: document.getElementById("customStartDate"),
  customEndDate: document.getElementById("customEndDate"),
  clearFiltersButton: document.getElementById("clearFiltersButton"),
  darkModeToggle: document.getElementById("darkModeToggle"),
  signOutButton: document.getElementById("signOutButton"),
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
let analyticsResizeObserver = null;
let analyticsSyncFrame = 0;
const devWarningFlags = new Set();
let initializationPromise = null;
let dataLoadPromise = null;

initialize();

async function initialize() {
  hydrateInitialFiltersFromUrl();
  attachCoreEventListeners();
  initializeAnalyticsPanelSync();
  clearDashboardData({ preserveFilters: true, render: true });
  setAppView("login");
  setAuthControlsEnabled({ loginEnabled: false, refreshEnabled: false, signOutEnabled: false });
  renderDiagnostics();

  console.log("MSAL Browser loaded:", Boolean(window.msal?.PublicClientApplication));

  console.info("Startup auth diagnostics", {
    graphConfigDetected: Boolean(window.GRAPH_CONFIG),
    msalLibraryDetected: Boolean(window.msal?.PublicClientApplication)
  });

  try {
    await startApplication();
  } catch (error) {
    console.error("MSAL initialization failed.", error);
    state.diagnostics.authStatus = "Initialization failed";
    renderDiagnostics();
    setAuthControlsEnabled({ loginEnabled: false, refreshEnabled: false, signOutEnabled: false });
    setAppView("error", {
      heading: "Authentication Unavailable",
      primaryMessage: "Microsoft authentication could not be initialized for this application.",
      secondaryMessage: "Please verify your Microsoft Entra configuration or contact the system administrator.",
      primaryAction: null,
      secondaryAction: null,
      showSpinner: false
    });
  }
}

function attachCoreEventListeners() {
  elements.loginSignInButton.addEventListener("click", async () => {
    await handleAuthAction(elements.loginSignInButton.dataset.action || "signin");
  });

  elements.authSecondaryButton.addEventListener("click", async () => {
    await handleAuthAction(elements.authSecondaryButton.dataset.action || "");
  });

  elements.searchInput.addEventListener("input", (event) => {
    state.searchQuery = event.target.value;
    applyFilters();
  });

  elements.tableWorkUnitFilter.addEventListener("change", (event) => {
    state.tableSelectedWorkUnit = event.target.value;
    applyFilters();
  });

  elements.tableStatusFilter.addEventListener("change", (event) => {
    state.tableSelectedStatus = event.target.value;
    applyFilters();
  });

  elements.tableReceivedSort.addEventListener("change", (event) => {
    state.tableReceivedSort = event.target.value;
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
    requestSyncAnalyticsPanelHeights();
  });

  elements.refreshNowButton.addEventListener("click", async () => {
    await refreshLiveData({ reason: "manual", showBusyState: true });
  });

  elements.signOutButton.addEventListener("click", async () => {
    await handleSignOutClick();
  });

  window.addEventListener("resize", () => {
    requestSyncAnalyticsPanelHeights();
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

async function initializeMsal() {
  if (initializationPromise) {
    return initializationPromise;
  }

  initializationPromise = (async () => {
    if (!window.msal || !window.msal.PublicClientApplication) {
      throw new Error("MSAL Browser failed to load.");
    }

    const config = window.GRAPH_CONFIG;
    if (!config?.clientId || !config?.tenantId || !config?.authority || !config?.redirectUri) {
      throw new Error("Microsoft Entra configuration is missing.");
    }

    if (msalInstance) {
      return msalInstance;
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

    console.info("MSAL initialized");
    return msalInstance;
  })();

  try {
    return await initializationPromise;
  } catch (error) {
    initializationPromise = null;
    throw error;
  }
}

async function startApplication() {
  await initializeMsal();

  let redirectResult = null;
  try {
    redirectResult = await msalInstance.handleRedirectPromise();
  } finally {
    sessionStorage.removeItem(AUTH_REDIRECT_PENDING_KEY);
  }

  console.info("Redirect result processed", {
    receivedRedirectAccount: Boolean(redirectResult?.account)
  });

  if (redirectResult?.account) {
    msalInstance.setActiveAccount(redirectResult.account);
  }

  const account = msalInstance.getActiveAccount() || msalInstance.getAllAccounts()[0];
  if (!account) {
    state.diagnostics.authStatus = "Ready to sign in";
    state.diagnostics.signedInUser = "N/A";
    renderDiagnostics();
    setAuthControlsEnabled({ loginEnabled: true, refreshEnabled: false, signOutEnabled: false });
    setAppView("login");
    return;
  }

  msalInstance.setActiveAccount(account);
  console.info("Active account detected", {
    username: account.username || "N/A"
  });
  updateAuthDiagnostics(account, "Signed in");
  setAuthControlsEnabled({ loginEnabled: false, refreshEnabled: true, signOutEnabled: true });

  setAppView("loading");
  const loaded = await refreshLiveData({ reason: "startup", showBusyState: false });
  if (loaded) {
    startLiveRefreshSchedule();
  }
}

function setAuthControlsEnabled({ loginEnabled, refreshEnabled, signOutEnabled }) {
  elements.loginSignInButton.disabled = !loginEnabled;
  elements.refreshNowButton.disabled = !refreshEnabled;
  elements.signOutButton.disabled = !signOutEnabled;
}

async function handleAuthAction(action) {
  if (action === "retry") {
    setAppView("loading");
    await refreshLiveData({ reason: "retry", showBusyState: false });
    return;
  }

  if (action === "signout") {
    await handleSignOutClick();
    return;
  }

  if (action === "continue-signin") {
    await continueMicrosoftSignIn();
    return;
  }

  if (action === "signin") {
    await handleSignInClick();
  }
}

function getRedirectStartPage() {
  return window.GRAPH_CONFIG?.redirectUri || "https://mdroot7282.github.io/foia-command-dashboard/";
}

async function continueMicrosoftSignIn() {
  if (!msalInstance) {
    return;
  }

  const account = getPreferredAccount();
  const button = elements.loginSignInButton;
  button.disabled = true;
  button.textContent = "Continuing Microsoft Sign-In...";
  setAuthControlsEnabled({ loginEnabled: false, refreshEnabled: false, signOutEnabled: Boolean(account) });

  try {
    sessionStorage.setItem(AUTH_REDIRECT_PENDING_KEY, "true");
    await msalInstance.acquireTokenRedirect({
      scopes: getGraphScopes(),
      account: account || undefined,
      redirectStartPage: getRedirectStartPage()
    });
  } catch (error) {
    sessionStorage.removeItem(AUTH_REDIRECT_PENDING_KEY);
    console.error("Microsoft continue sign-in redirect failed.", error);
    setAppView("error", {
      heading: "Authentication Unavailable",
      primaryMessage: "Microsoft sign-in could not continue.",
      secondaryMessage: "Please select Continue Microsoft Sign-In again.",
      primaryAction: { label: "Continue Microsoft Sign-In", action: "continue-signin" },
      secondaryAction: account ? { label: "Sign Out", action: "signout" } : null,
      showSpinner: false
    });
    setAuthControlsEnabled({ loginEnabled: true, refreshEnabled: Boolean(account), signOutEnabled: Boolean(account) });
  }
}

async function handleSignInClick() {
  if (!window.msal || !window.msal.PublicClientApplication) {
    throw new Error("MSAL Browser failed to load.");
  }

  const config = window.GRAPH_CONFIG;
  if (!config?.clientId || !config?.tenantId || !config?.authority || !config?.redirectUri) {
    throw new Error("Microsoft Entra configuration is missing.");
  }

  if (!msalInstance) {
    setAppView("error", {
      heading: "Authentication Unavailable",
      primaryMessage: "Microsoft sign-in is unavailable because authentication initialization did not complete.",
      secondaryMessage: "Please reload the page or contact the system administrator.",
      primaryAction: null,
      secondaryAction: null,
      showSpinner: false
    });
    return;
  }

  if (sessionStorage.getItem(AUTH_REDIRECT_PENDING_KEY) === "true") {
    setAppView("loading", {
      heading: "Microsoft Sign-In In Progress",
      primaryMessage: "",
      secondaryMessage: "Please complete the Microsoft sign-in process in this browser window.",
      primaryAction: null,
      secondaryAction: null,
      showSpinner: true
    });
    return;
  }

  elements.loginSignInButton.disabled = true;
  elements.loginSignInButton.textContent = "Redirecting to Microsoft...";
  setAppView("loading");

  try {
    sessionStorage.setItem(AUTH_REDIRECT_PENDING_KEY, "true");
    await msalInstance.loginRedirect({
      scopes: window.GRAPH_CONFIG.scopes,
      redirectStartPage: getRedirectStartPage(),
      prompt: "select_account"
    });
  } catch (error) {
    sessionStorage.removeItem(AUTH_REDIRECT_PENDING_KEY);
    console.error("Microsoft login redirect failed.", error);
    await handleLiveLoadFailure(error, { reason: "signin" });
    elements.loginSignInButton.disabled = false;
    elements.loginSignInButton.textContent = "Sign in with Microsoft";
    setAuthControlsEnabled({ loginEnabled: true, refreshEnabled: false, signOutEnabled: false });
  }
}

async function refreshLiveData(options = {}) {
  const {
    reason = "scheduled",
    showBusyState = false
  } = options;

  if (dataLoadPromise) {
    return dataLoadPromise;
  }

  dataLoadPromise = (async () => {
    if (showBusyState) {
      elements.refreshNowButton.disabled = true;
      elements.refreshNowButton.textContent = "Refreshing...";
    }

    try {
      await withTimeout(loadLiveSharePointData(), SHAREPOINT_LOAD_TIMEOUT_MS);
      setAppView("dashboard");
      setAuthControlsEnabled({ loginEnabled: false, refreshEnabled: true, signOutEnabled: true });
      return true;
    } catch (error) {
      await handleLiveLoadFailure(error, { reason });
      return false;
    } finally {
      if (showBusyState) {
        elements.refreshNowButton.disabled = false;
        elements.refreshNowButton.textContent = "Refresh Now";
      }
    }
  })().finally(() => {
    dataLoadPromise = null;
  });

  return dataLoadPromise;
}

async function loadLiveSharePointData() {
  const token = await acquireGraphAccessToken();
  console.info("Token acquired silently");

  const site = await resolveSharePointSite(token);
  state.liveConnection.site.id = site.id || "";
  state.liveConnection.site.displayName = site.displayName || "";
  state.liveConnection.site.webUrl = site.webUrl || "";
  state.diagnostics.siteResolved = true;
  console.info("Graph site resolved", {
    siteId: state.liveConnection.site.id,
    displayName: state.liveConnection.site.displayName
  });
  renderDiagnostics();

  const list = await resolveFoiaList(token, site.id);
  state.liveConnection.list.id = list.id || "";
  state.liveConnection.list.displayName = list.displayName || "";
  state.liveConnection.list.webUrl = list.webUrl || "";
  state.diagnostics.listResolved = true;
  console.info("Graph list resolved", {
    listId: state.liveConnection.list.id,
    displayName: state.liveConnection.list.displayName
  });
  renderDiagnostics();

  const columns = await fetchAllListColumns(token, site.id, list.id);
  const fieldMap = buildSharePointFieldMap(columns);
  state.liveConnection.fieldMap = fieldMap;
  state.availableStatuses = deriveStatusChoices(columns, fieldMap.status);

  const items = await fetchAllListItems(token, site.id, list.id);
  console.log("Number of list items returned:", items.length);
  items.slice(0, 3).forEach((item) => {
    console.log("SharePoint field keys:", Object.keys(item?.fields || {}));
  });

  const graphRecords = items.map((item) => mapGraphItemToRecord(item, fieldMap));
  logFieldMappingValidationSummary(items, graphRecords);
  logFoiaNumberValidationSummary(items, graphRecords);

  state.diagnostics.liveItemsLoaded = items.length;
  state.diagnostics.dataSource = "SharePoint";
  state.diagnostics.lastUpdated = new Date();
  console.info("Records loaded", {
    count: items.length
  });
  showStatusMessage("Live SharePoint data connected", "live");
  renderDiagnostics();

  setDashboardRecords(graphRecords);
}

async function handleLiveLoadFailure(error, options = {}) {
  const { reason = "scheduled" } = options;

  console.error("SharePoint authentication or loading failed.", error);
  clearDashboardData({ preserveFilters: true, render: true });
  renderDiagnostics();
  const account = getPreferredAccount();
  const hasAuthenticatedAccount = Boolean(account);

  if (isInteractionRequired(error)) {
    state.diagnostics.authStatus = "Authentication interaction required";
    if (hasAuthenticatedAccount) {
      updateAuthDiagnostics(account, "Signed in");
    } else {
      state.diagnostics.signedInUser = "N/A";
      renderDiagnostics();
    }

    setAppView("error", {
      heading: "Microsoft Session Verification Required",
      primaryMessage: "Your Microsoft session has expired. Please sign in again.",
      secondaryMessage: "Select Continue Microsoft Sign-In to complete authentication.",
      primaryAction: { label: "Continue Microsoft Sign-In", action: "continue-signin" },
      secondaryAction: hasAuthenticatedAccount ? { label: "Sign Out", action: "signout" } : null,
      showSpinner: false
    });
    setAuthControlsEnabled({ loginEnabled: true, refreshEnabled: hasAuthenticatedAccount, signOutEnabled: hasAuthenticatedAccount });
    return;
  }

  if (isSessionExpiredError(error)) {
    stopLiveRefreshSchedule();
    if (msalInstance && typeof msalInstance.setActiveAccount === "function") {
      msalInstance.setActiveAccount(null);
    }
    state.diagnostics.authStatus = "Session expired";
    state.diagnostics.signedInUser = "N/A";
    renderDiagnostics();
    setAppView("login", {
      heading: "Session Expired",
      primaryMessage: "Your Microsoft session has expired.",
      secondaryMessage: "Sign in with your Illinois.gov Microsoft account to continue.",
      primaryAction: { label: "Sign in with Microsoft", action: "signin" },
      secondaryAction: null,
      showSpinner: false
    });
    return;
  }

  if (isAccessDeniedError(error)) {
    state.diagnostics.authStatus = "Access denied";
    renderDiagnostics();
    setAppView("access-denied", {
      primaryMessage: "You are signed in but do not have permission to access the DCI FOIA SharePoint list.",
      secondaryAction: hasAuthenticatedAccount ? { label: "Sign Out", action: "signout" } : null
    });
    setAuthControlsEnabled({ loginEnabled: false, refreshEnabled: false, signOutEnabled: hasAuthenticatedAccount });
    return;
  }

  const graphStatus = getGraphStatusFromError(error);
  if (graphStatus === 401) {
    setAppView("error", {
      heading: "Session Expired",
      primaryMessage: "Your Microsoft session has expired. Please sign in again.",
      secondaryMessage: "Select Continue Microsoft Sign-In to renew your Microsoft session.",
      primaryAction: { label: "Continue Microsoft Sign-In", action: "continue-signin" },
      secondaryAction: hasAuthenticatedAccount ? { label: "Sign Out", action: "signout" } : null,
      showSpinner: false
    });
    setAuthControlsEnabled({ loginEnabled: true, refreshEnabled: hasAuthenticatedAccount, signOutEnabled: hasAuthenticatedAccount });
    return;
  }

  if (graphStatus === 404) {
    setAppView("error", {
      heading: "SharePoint Resource Not Found",
      primaryMessage: "The DCI FOIA SharePoint site or list could not be located.",
      secondaryMessage: "Verify the configured site and list paths, then select Try Again.",
      primaryAction: { label: "Try Again", action: "retry" },
      secondaryAction: hasAuthenticatedAccount ? { label: "Sign Out", action: "signout" } : null,
      showSpinner: false
    });
    setAuthControlsEnabled({ loginEnabled: false, refreshEnabled: hasAuthenticatedAccount, signOutEnabled: hasAuthenticatedAccount });
    return;
  }

  showStatusMessage("Live SharePoint data is temporarily unavailable. Select Try Again.", "warning");
  setAppView("error", {
    heading: "Unable to Retrieve Live SharePoint Data",
    primaryMessage: "Live SharePoint data is temporarily unavailable. Select Try Again.",
    secondaryMessage: "Your Microsoft session is still active.",
    primaryAction: { label: "Try Again", action: "retry" },
    secondaryAction: hasAuthenticatedAccount ? { label: "Sign Out", action: "signout" } : null,
    showSpinner: false
  });
  setAuthControlsEnabled({ loginEnabled: false, refreshEnabled: hasAuthenticatedAccount, signOutEnabled: hasAuthenticatedAccount });
}

async function acquireGraphAccessToken() {
  const account = getPreferredAccount();
  if (!account) {
    const error = new Error("No signed-in account. Sign in with Microsoft.");
    error.authInteractionRequired = true;
    throw error;
  }

  const scopes = getGraphScopes();
  const tokenRequest = {
    scopes,
    account: msalInstance.getActiveAccount()
  };
  console.log("Graph token scopes requested:", scopes);

  try {
    const tokenResult = await msalInstance.acquireTokenSilent(tokenRequest);
    return tokenResult.accessToken;
  } catch (error) {
    if (isInteractionRequired(error)) {
      error.authInteractionRequired = true;
    }
    throw error;
  }
}

async function graphFetch(token, url) {
  for (let attempt = 0; attempt <= GRAPH_MAX_RETRIES; attempt += 1) {
    let response;

    try {
      response = await fetch(url, {
        headers: {
          Authorization: `Bearer ${token}`,
          Accept: "application/json"
        }
      });
    } catch (networkError) {
      if (attempt < GRAPH_MAX_RETRIES) {
        await waitForMs(getRetryDelayMs(null, attempt));
        continue;
      }
      const error = new Error("Network request failed");
      error.graphTransient = true;
      error.cause = networkError;
      throw error;
    }

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
      const retryable = GRAPH_RETRYABLE_STATUS.has(status);
      if (retryable && attempt < GRAPH_MAX_RETRIES) {
        await waitForMs(getRetryDelayMs(response.headers.get("Retry-After"), attempt));
        continue;
      }

      console.error("Graph request failed", {
        url,
        status,
        error: graphError
      });
      const error = new Error(`Graph request failed (${status})`);
      error.graphStatus = status;
      error.graphError = graphError;
      error.graphTransient = retryable;
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

  throw new Error("Graph request retry limit reached.");
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
  const foiaNumber = normalizeGraphFieldText(fields.Title) || "Not Assigned";
  const statusFromField = readMappedChoice(fields, fieldMap.status);
  const status = statusFromField || "Unknown";

  if (!statusFromField) {
    warnOnce("status-mapping-missing-value", "Status field mapping returned no value for one or more records. Status mapping may need attention.");
  }

  const workUnits = getWorkUnits({
    workUnit: readMappedValue(fields, fieldMap.workUnit)
  });
  const workUnit = getWorkUnitDisplay(workUnits);
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
    workUnits,
    workUnit,
    assignedTo,
    subject,
    crimepadCase: readMappedText(fields, fieldMap.crimepadCase),
    tracsCase: readMappedText(fields, fieldMap.tracsCase),
    otherCase: readMappedText(fields, fieldMap.otherCase),
    modifiedDate: readMappedDate(fields, fieldMap.modified),
    dueDate,
    closedDate: readMappedDate(fields, fieldMap.closedDate),
    pendingWithLegalDate: readMappedDate(fields, fieldMap.pendingWithLegalDate),
    notes: readMappedText(fields, fieldMap.notes)
  };

  return {
    id: normalizedRecord.id,
    request_id: normalizedRecord.foiaNumber,
    dci_work_unit: normalizedRecord.workUnit,
    workUnits: normalizedRecord.workUnits,
    workUnit: normalizedRecord.workUnit,
    requester,
    subject: normalizedRecord.subject,
    received_date: normalizedRecord.receivedDate,
    due_date: normalizedRecord.dueDate,
    status: normalizedRecord.status,
    assigned_to: normalizedRecord.assignedTo,
    last_update: normalizedRecord.modifiedDate,
    closed_date: normalizedRecord.closedDate,
    pending_with_legal_date: normalizedRecord.pendingWithLegalDate,
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
    title: "Title",
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
    pendingWithLegalDate: findInternalName(
      [
        "PENDING WITH LEGAL DATE",
        "Pending With Legal Date",
        "DATE PENDING WITH LEGAL",
        "Pending Legal Date",
        "LEGAL PENDING DATE",
        "DATE SENT TO LEGAL"
      ],
      internalNameByDisplayName,
      fallbackByInternalName
    ),
    notes: findInternalName(["NOTES", "Notes", "COMMENTS", "Comments"], internalNameByDisplayName, fallbackByInternalName)
  };

  const configuredPendingLegalInternalName = String(window.GRAPH_CONFIG?.pendingWithLegalDateInternalName || "").trim();
  if (configuredPendingLegalInternalName) {
    const resolvedConfiguredPendingLegal = columns.find((column) => column.name === configuredPendingLegalInternalName);
    if (resolvedConfiguredPendingLegal) {
      fieldMap.pendingWithLegalDate = resolvedConfiguredPendingLegal.name;
    } else {
      console.warn("Configured pendingWithLegalDateInternalName was not found in SharePoint columns.", {
        configuredPendingWithLegalDateInternalName: configuredPendingLegalInternalName
      });
    }
  }

  const mappingTable = [
    ["Title", fieldMap.title],
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
    ["Modified", fieldMap.modified],
    ["Pending With Legal Date", fieldMap.pendingWithLegalDate]
  ];

  console.table(
    mappingTable.map(([visibleName, internalName]) => ({
      visibleName,
      internalName: internalName || "(unmapped)"
    }))
  );

  return fieldMap;
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

function normalizeWorkUnitKey(value) {
  return String(value || "").trim().toLowerCase().replace(/\s+/g, " ");
}

function splitDelimitedWorkUnitText(value) {
  const normalized = String(value || "").replace(/\r?\n/g, " ").trim();
  if (!normalized) {
    return [];
  }

  if (normalized.includes(";#")) {
    const parts = normalized.split(";#").map((entry) => entry.trim()).filter(Boolean);
    return parts.filter((part, index) => !(index < parts.length - 1 && /^\d+$/.test(part)));
  }

  return normalized.split(/[;,]+/).map((entry) => entry.trim()).filter(Boolean);
}

function collectWorkUnitTokens(value, tokens) {
  if (value === undefined || value === null) {
    return;
  }

  if (Array.isArray(value)) {
    value.forEach((entry) => collectWorkUnitTokens(entry, tokens));
    return;
  }

  if (typeof value === "object") {
    if (Array.isArray(value.results)) {
      collectWorkUnitTokens(value.results, tokens);
      return;
    }

    const candidateFields = [
      value.displayName,
      value.LookupValue,
      value.title,
      value.Label,
      value.value,
      value.name
    ];

    candidateFields.forEach((entry) => {
      if (entry !== undefined && entry !== null && entry !== "") {
        collectWorkUnitTokens(entry, tokens);
      }
    });
    return;
  }

  splitDelimitedWorkUnitText(value).forEach((token) => {
    if (!/^\d+$/.test(token)) {
      tokens.push(token);
    }
  });
}

function normalizeWorkUnitLabel(value) {
  const normalized = String(value || "").trim().replace(/\s+/g, " ");
  if (!normalized) {
    return "";
  }

  return OFFICIAL_WORK_UNIT_BY_KEY.get(normalizeWorkUnitKey(normalized)) || normalized;
}

function getWorkUnits(record) {
  const candidates = [];
  if (record && typeof record === "object" && !Array.isArray(record)) {
    candidates.push(record.workUnits, record.workUnit, record.dci_work_unit, record.work_unit);
  } else {
    candidates.push(record);
  }

  const uniqueUnits = new Map();
  candidates.forEach((candidate) => {
    const tokens = [];
    collectWorkUnitTokens(candidate, tokens);
    tokens.forEach((token) => {
      const label = normalizeWorkUnitLabel(token);
      if (!label) {
        return;
      }

      const key = normalizeWorkUnitKey(label);
      if (!uniqueUnits.has(key)) {
        uniqueUnits.set(key, label);
      }
    });
  });

  if (!uniqueUnits.size) {
    return ["Unknown"];
  }

  return Array.from(uniqueUnits.values());
}

function getWorkUnitDisplay(value) {
  const workUnits = Array.isArray(value) ? value : getWorkUnits(value);
  return workUnits.length ? workUnits.join(", ") : "Unknown";
}

function recordHasWorkUnit(record, workUnit) {
  const targetKey = normalizeWorkUnitKey(workUnit);
  return getWorkUnits(record).some((unit) => normalizeWorkUnitKey(unit) === targetKey);
}

function orderWorkUnits(values) {
  const uniqueValues = Array.from(new Set(values.filter(Boolean)));
  const orderedKnown = DCI_WORK_UNITS.filter((unit) => uniqueValues.includes(unit));
  const extras = uniqueValues
    .filter((unit) => !DCI_WORK_UNITS.includes(unit))
    .sort((a, b) => a.localeCompare(b));

  return [...orderedKnown, ...extras];
}

function getRecordIdentity(record) {
  return String(record.id || record.item_id || record.request_id || record.foiaNumber || "").trim();
}

function logWorkUnitValidationSummary(records, generatedRowCount) {
  const uniqueRecordIds = new Set();
  let singleUnitCount = 0;
  let multiUnitCount = 0;
  let maxWorkUnits = 0;

  records.forEach((record) => {
    const recordId = getRecordIdentity(record);
    if (recordId) {
      uniqueRecordIds.add(recordId);
    }

    const resolvedUnits = getWorkUnits(record).filter((unit) => normalizeWorkUnitKey(unit) !== "unknown");
    if (resolvedUnits.length === 1) {
      singleUnitCount += 1;
    }
    if (resolvedUnits.length > 1) {
      multiUnitCount += 1;
    }
    maxWorkUnits = Math.max(maxWorkUnits, resolvedUnits.length);
  });

  console.info("Work unit validation summary", {
    totalUniqueFoiaRecords: uniqueRecordIds.size || records.length,
    foiasAssignedToOneWorkUnit: singleUnitCount,
    foiasAssignedToMultipleWorkUnits: multiUnitCount,
    maximumWorkUnitsAssignedToOneFoia: maxWorkUnits,
    generatedIndividualWorkUnitRows: generatedRowCount
  });
}

function logFieldMappingValidationSummary(items, records) {
  const summary = {
    totalItemsRetrieved: items.length,
    recordsWithMappedWorkUnits: records.filter((record) => getWorkUnits(record).some((unit) => normalizeWorkUnitKey(unit) !== "unknown")).length,
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

function setAppView(view, overrides = {}) {
  state.appView = view;

  const baseConfig = {
    heading: "Secure Access Required",
    primaryMessage: "This application contains official law enforcement information.",
    secondaryMessage: "You must sign in with your Illinois.gov Microsoft account to continue.",
    primaryAction: { label: "Sign in with Microsoft", action: "signin" },
    secondaryAction: null,
    showSpinner: false,
    showAuth: view !== "dashboard",
    showDashboard: view === "dashboard"
  };

  const viewConfig = {
    login: baseConfig,
    loading: {
      ...baseConfig,
      heading: "Loading Live FOIA Data...",
      primaryMessage: "",
      secondaryMessage: "Authenticating your session and retrieving the required SharePoint records.",
      primaryAction: null,
      secondaryAction: null,
      showSpinner: true
    },
    error: {
      ...baseConfig,
      heading: "Unable to Retrieve Live SharePoint Data",
      primaryMessage: "Unable to retrieve live SharePoint data.",
      secondaryMessage: "Please verify your network connection or contact the system administrator.",
      primaryAction: { label: "Try Again", action: "retry" },
      secondaryAction: { label: "Sign Out", action: "signout" },
      showSpinner: false
    },
    "access-denied": {
      ...baseConfig,
      heading: "Access Denied",
      primaryMessage: "Your Illinois.gov Microsoft account is authenticated, but you do not have permission to access the DCI FOIA SharePoint list.",
      secondaryMessage: "If you believe you should have access, contact the system administrator.",
      primaryAction: { label: "Sign Out", action: "signout" },
      secondaryAction: null,
      showSpinner: false
    },
    dashboard: {
      ...baseConfig,
      showAuth: false,
      showDashboard: true
    }
  }[view] || baseConfig;

  const config = { ...viewConfig, ...overrides };

  elements.authShell.hidden = !config.showAuth;
  elements.dashboardShell.hidden = !config.showDashboard;
  elements.authStateHeading.textContent = config.heading;
  elements.authPrimaryMessage.textContent = config.primaryMessage;
  elements.authSecondaryMessage.textContent = config.secondaryMessage;
  elements.authSpinner.hidden = !config.showSpinner;

  if (config.primaryAction) {
    elements.loginSignInButton.hidden = false;
    elements.loginSignInButton.textContent = config.primaryAction.label;
    elements.loginSignInButton.dataset.action = config.primaryAction.action;
  } else {
    elements.loginSignInButton.hidden = true;
    elements.loginSignInButton.dataset.action = "";
  }

  if (config.secondaryAction) {
    elements.authSecondaryButton.hidden = false;
    elements.authSecondaryButton.textContent = config.secondaryAction.label;
    elements.authSecondaryButton.dataset.action = config.secondaryAction.action;
  } else {
    elements.authSecondaryButton.hidden = true;
    elements.authSecondaryButton.dataset.action = "";
  }

  if (config.showDashboard) {
    requestSyncAnalyticsPanelHeights();
  }
}

function initializeAnalyticsPanelSync() {
  if (!elements.receivedCompletedPanel || analyticsResizeObserver || typeof ResizeObserver !== "function") {
    return;
  }

  analyticsResizeObserver = new ResizeObserver(() => {
    requestSyncAnalyticsPanelHeights();
  });

  analyticsResizeObserver.observe(elements.receivedCompletedPanel);
}

function requestSyncAnalyticsPanelHeights() {
  if (analyticsSyncFrame) {
    window.cancelAnimationFrame(analyticsSyncFrame);
  }

  analyticsSyncFrame = window.requestAnimationFrame(() => {
    analyticsSyncFrame = 0;
    syncAnalyticsPanelHeights();
  });
}

function syncAnalyticsPanelHeights() {
  const referencePanel = elements.receivedCompletedPanel;
  const deadlinePanel = elements.upcomingDeadlinesPanel;
  const alertsPanel = elements.commandAlertsPanel;

  if (!referencePanel || !deadlinePanel || !alertsPanel || elements.dashboardShell.hidden) {
    return;
  }

  deadlinePanel.style.height = "";
  alertsPanel.style.height = "";

  const targetHeight = referencePanel.getBoundingClientRect().height;
  if (!Number.isFinite(targetHeight) || targetHeight <= 0) {
    return;
  }

  const pixelHeight = `${Math.round(targetHeight)}px`;
  deadlinePanel.style.height = pixelHeight;
  alertsPanel.style.height = pixelHeight;
}

function stopLiveRefreshSchedule() {
  if (!liveRefreshTimerId) {
    return;
  }
  window.clearInterval(liveRefreshTimerId);
  liveRefreshTimerId = null;
}

function resetLiveConnectionState() {
  state.liveConnection.site.id = "";
  state.liveConnection.site.displayName = "";
  state.liveConnection.site.webUrl = "";
  state.liveConnection.list.id = "";
  state.liveConnection.list.displayName = "";
  state.liveConnection.list.webUrl = "";
  state.liveConnection.fieldMap = {};
  state.diagnostics.siteResolved = false;
  state.diagnostics.listResolved = false;
  state.diagnostics.liveItemsLoaded = 0;
  state.diagnostics.dataSource = "None";
  state.diagnostics.lastUpdated = null;
}

function clearDashboardData(options = {}) {
  const { preserveFilters = false, render = false } = options;

  state.records = [];
  state.availableWorkUnits = [];
  state.baseFilteredRecords = [];
  state.scopedRecords = [];
  state.filteredRecords = [];
  state.selectedId = null;
  state.tableAvailableWorkUnits = [];
  state.tableAvailableStatuses = [];
  state.availableStatuses = [...FALLBACK_STATUS_CHOICES];

  if (!preserveFilters) {
    resetFiltersToDefault();
  }

  resetLiveConnectionState();

  populateWorkUnitFilter();
  populateStatusFilter();
  elements.workUnitFilter.value = state.selectedWorkUnit;
  elements.statusFilter.value = state.selectedStatus;
  elements.timePeriodFilter.value = state.selectedTimePeriod;
  elements.customStartDate.value = state.customStartDate;
  elements.customEndDate.value = state.customEndDate;
  elements.searchInput.value = state.searchQuery;
  populateTableLocalFilters(state.filteredRecords);

  if (render) {
    renderAll();
  }
}

function isSessionExpiredError(error) {
  return Boolean(error?.authInteractionRequired) || String(error?.errorCode || "").toLowerCase().includes("interaction");
}

function getGraphStatusFromError(error) {
  return error?.graphStatus || error?.cause?.graphStatus || error?.cause?.cause?.graphStatus || 0;
}

function isAccessDeniedError(error) {
  const status = getGraphStatusFromError(error);
  if (status === 401 || status === 403) {
    return true;
  }
  const message = String(error?.graphError || error?.cause?.graphError || error?.message || "").toLowerCase();
  return message.includes("access denied") || message.includes("forbidden") || message.includes("insufficient privileges");
}

async function handleSignOutClick() {
  stopLiveRefreshSchedule();
  setAuthControlsEnabled({ loginEnabled: false, refreshEnabled: false, signOutEnabled: false });
  clearDashboardData({ preserveFilters: false, render: true });

  try {
    if (msalInstance) {
      await msalInstance.logoutRedirect({
        account: msalInstance.getActiveAccount(),
        postLogoutRedirectUri: getRedirectStartPage()
      });
      if (typeof msalInstance.setActiveAccount === "function") {
        msalInstance.setActiveAccount(null);
      }
    }
  } catch (error) {
    console.error("Microsoft sign-out failed.", error);
  } finally {
    state.diagnostics.authStatus = "Not signed in";
    state.diagnostics.signedInUser = "N/A";
    renderDiagnostics();
    setAppView("login");
    setAuthControlsEnabled({ loginEnabled: true, refreshEnabled: false, signOutEnabled: false });
  }
}

function getRetryDelayMs(retryAfterHeader, attempt) {
  const baseDelay = 1000 * Math.pow(2, attempt);
  if (!retryAfterHeader) {
    return baseDelay;
  }

  const retryAfterSeconds = Number(retryAfterHeader);
  if (Number.isFinite(retryAfterSeconds) && retryAfterSeconds >= 0) {
    return Math.max(baseDelay, retryAfterSeconds * 1000);
  }

  const retryAfterDate = new Date(retryAfterHeader).getTime();
  if (Number.isFinite(retryAfterDate)) {
    const delta = retryAfterDate - Date.now();
    if (delta > 0) {
      return Math.max(baseDelay, delta);
    }
  }

  return baseDelay;
}

async function waitForMs(delayMs) {
  return new Promise((resolve) => {
    window.setTimeout(resolve, Math.max(0, delayMs));
  });
}

async function withTimeout(promise, timeoutMs) {
  let timeoutId = 0;
  const timeoutPromise = new Promise((_, reject) => {
    timeoutId = window.setTimeout(() => {
      const timeoutError = new Error("SharePoint load timed out");
      timeoutError.graphTransient = true;
      timeoutError.loadTimedOut = true;
      reject(timeoutError);
    }, timeoutMs);
  });

  try {
    return await Promise.race([promise, timeoutPromise]);
  } finally {
    if (timeoutId) {
      window.clearTimeout(timeoutId);
    }
  }
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
  logWorkUnitValidationSummary(state.records, state.availableWorkUnits.length);
  logDueDateValidationSummary(state.records);
  logOperationalClosureValidationSummary(state.records);
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
      showBusyState: false
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
  return orderWorkUnits(
    records.flatMap((record) => getWorkUnits(record))
  );
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

function buildPresentStatuses(records) {
  const normalizedToDisplay = new Map();
  records.forEach((record) => {
    const statusText = String(record.status || "").trim();
    if (!statusText) {
      return;
    }
    const key = statusKey(statusText);
    if (!normalizedToDisplay.has(key)) {
      normalizedToDisplay.set(key, statusText);
    }
  });

  const remaining = new Map(normalizedToDisplay);
  const orderedKnown = [];

  KPI_STATUS_ORDER.forEach((preferredStatus) => {
    const preferredKey = statusKey(preferredStatus);
    if (!remaining.has(preferredKey)) {
      return;
    }
    orderedKnown.push(remaining.get(preferredKey));
    remaining.delete(preferredKey);
  });

  const extras = Array.from(remaining.values()).sort((a, b) => a.localeCompare(b));
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

function populateTableLocalFilters(records) {
  state.tableAvailableWorkUnits = buildAvailableWorkUnits(records);
  state.tableAvailableStatuses = buildPresentStatuses(records);

  if (
    state.tableSelectedWorkUnit !== ALL_WORK_UNITS_OPTION
    && !state.tableAvailableWorkUnits.includes(state.tableSelectedWorkUnit)
  ) {
    state.tableSelectedWorkUnit = ALL_WORK_UNITS_OPTION;
  }

  if (
    state.tableSelectedStatus !== ALL_STATUSES_OPTION
    && !state.tableAvailableStatuses.some((status) => statusEquals(status, state.tableSelectedStatus))
  ) {
    state.tableSelectedStatus = ALL_STATUSES_OPTION;
  }

  elements.tableWorkUnitFilter.innerHTML = [
    `<option value="${ALL_WORK_UNITS_OPTION}">All Work Units</option>`,
    ...state.tableAvailableWorkUnits.map(
      (workUnit) => `<option value="${escapeHtml(workUnit)}">${escapeHtml(workUnit)}</option>`
    )
  ].join("");

  elements.tableStatusFilter.innerHTML = [
    `<option value="${ALL_STATUSES_OPTION}">${STATUS_FILTER_ALL_LABEL}</option>`,
    ...state.tableAvailableStatuses.map((status) => `<option value="${escapeHtml(status)}">${escapeHtml(status)}</option>`)
  ].join("");

  elements.tableWorkUnitFilter.value = state.tableSelectedWorkUnit;
  elements.tableStatusFilter.value = resolveStatusOptionValue(state.tableSelectedStatus, state.tableAvailableStatuses);
  elements.tableReceivedSort.value = state.tableReceivedSort;
}

function getReceivedDateSortValue(record) {
  const rawValue = record?.received_date;
  if (!rawValue) {
    return null;
  }

  const timeValue = rawValue instanceof Date ? rawValue.getTime() : new Date(rawValue).getTime();
  return Number.isFinite(timeValue) ? timeValue : null;
}

function sortRecordsByReceivedDate(records, direction) {
  const sortDirection = direction === LOCAL_RECEIVED_SORT_OLDEST
    ? LOCAL_RECEIVED_SORT_OLDEST
    : LOCAL_RECEIVED_SORT_NEWEST;

  return [...records].sort((a, b) => {
    const aDate = getReceivedDateSortValue(a);
    const bDate = getReceivedDateSortValue(b);

    if (aDate === null && bDate === null) {
      return String(a.request_id || "").localeCompare(String(b.request_id || ""));
    }
    if (aDate === null) {
      return 1;
    }
    if (bDate === null) {
      return -1;
    }

    if (sortDirection === LOCAL_RECEIVED_SORT_OLDEST) {
      return aDate - bDate;
    }

    return bDate - aDate;
  });
}

function applyFilters() {
  const normalizedQuery = state.searchQuery.trim().toLowerCase();

  const periodRange = getTimePeriodRange(state.selectedTimePeriod);
  const customRange = getCustomDateRange(state.customStartDate, state.customEndDate);

  state.baseFilteredRecords = state.records.filter((record) => {
    if (state.selectedWorkUnit !== ALL_WORK_UNITS_OPTION && !recordHasWorkUnit(record, state.selectedWorkUnit)) {
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

  const searchedRecords = state.baseFilteredRecords.filter((record) => {
    if (!normalizedQuery) {
      return true;
    }

    const searchable = [
      record.request_id,
      record.subject,
      record.dci_work_unit,
      record.status
    ].join(" ").toLowerCase();

    return searchable.includes(normalizedQuery);
  });

  populateTableLocalFilters(searchedRecords);

  const locallyFilteredRecords = searchedRecords.filter((record) => {
    if (
      state.tableSelectedWorkUnit !== ALL_WORK_UNITS_OPTION
      && !recordHasWorkUnit(record, state.tableSelectedWorkUnit)
    ) {
      return false;
    }

    if (
      state.tableSelectedStatus !== ALL_STATUSES_OPTION
      && !statusEquals(record.status, state.tableSelectedStatus)
    ) {
      return false;
    }

    return true;
  });

  state.filteredRecords = sortRecordsByReceivedDate(locallyFilteredRecords, state.tableReceivedSort);

  if (!state.filteredRecords.some((item) => getRecordIdentity(item) === state.selectedId)) {
    state.selectedId = getRecordIdentity(state.filteredRecords[0]) || null;
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
  requestSyncAnalyticsPanelHeights();
}

function renderKpis() {
  const records = state.scopedRecords;
  const now = new Date();

  const statusCounts = KPI_STATUS_ORDER.map((statusLabel) => countByStatus(records, statusLabel));
  const inProgressRecords = records.filter((record) => isInProgressMetricEligible(record));
  const dueToday = records.filter((record) => isDueToday(record));
  const due5 = records.filter((record) => isDueInFiveDays(record));
  const due10 = records.filter((record) => isDueInTenDays(record));
  const overdue = records.filter((record) => isOverdue(record));
  const avgDaysToReceive = calculateAverageDaysToReceive(records);

  const inProgressDurations = inProgressRecords
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
  elements.kpiAvgDaysToReceive.textContent = `${avgDaysToReceive.averageDays} Days`;
  elements.kpiStatusCompleted.textContent = String(statusCounts[4]);
  elements.kpiDue10.textContent = String(due10.length);
  elements.kpiDue5.textContent = String(due5.length);
  elements.kpiDueToday.textContent = String(dueToday.length);
  elements.kpiOverdue.textContent = String(overdue.length);
  elements.kpiAvgDaysInProgress.textContent = String(avgDaysInProgress);
  elements.kpiReceivedThisMonth.textContent = String(receivedThisMonth);
  elements.kpiCompletedThisMonth.textContent = String(completedThisMonth);

  logAverageDaysToReceiveSummary(avgDaysToReceive);
  elements.kpiScopeSummary.textContent = buildFilterSummaryText();
}

function calculateAverageDaysToReceive(records) {
  let includedRecords = 0;
  let skippedStampedByLegalBlank = 0;
  let skippedDateDciReceivedBlank = 0;
  let totalDays = 0;

  records.forEach((record) => {
    const stampedByLegalDate = record.date_stamped;
    const dateDciReceived = record.received_date;

    if (!(stampedByLegalDate instanceof Date) || Number.isNaN(stampedByLegalDate.getTime())) {
      skippedStampedByLegalBlank += 1;
      return;
    }

    if (!(dateDciReceived instanceof Date) || Number.isNaN(dateDciReceived.getTime())) {
      skippedDateDciReceivedBlank += 1;
      return;
    }

    const dayDifference = daysBetween(stampedByLegalDate, dateDciReceived);
    if (!Number.isFinite(dayDifference) || dayDifference < 0) {
      return;
    }

    totalDays += dayDifference;
    includedRecords += 1;
  });

  const averageDays = includedRecords ? Math.round(totalDays / includedRecords) : 0;
  return {
    includedRecords,
    skippedStampedByLegalBlank,
    skippedDateDciReceivedBlank,
    averageDays
  };
}

function logAverageDaysToReceiveSummary(summary) {
  console.info("Average days to receive validation summary", {
    recordsIncludedInCalculation: summary.includedRecords,
    recordsSkippedStampedByLegalBlank: summary.skippedStampedByLegalBlank,
    recordsSkippedDateDciReceivedBlank: summary.skippedDateDciReceivedBlank,
    calculatedAverageDaysToReceive: summary.averageDays
  });
}

function renderWorkUnitSummary() {
  const statusColumns = [...KPI_STATUS_ORDER];
  const records = state.scopedRecords;
  const monthReference = new Date();
  const monthScopeRecords = getCurrentMonthMetricScopeRecords();
  const unitsInScope = state.selectedWorkUnit === ALL_WORK_UNITS_OPTION
    ? buildAvailableWorkUnits(records)
    : [state.selectedWorkUnit];

  const rows = unitsInScope
    .map((workUnit) => {
      const unitRecords = records.filter((record) => recordHasWorkUnit(record, workUnit));
      const statusCounts = statusColumns.map((status) => unitRecords.filter((record) => statusEquals(record.status, status)).length);
      const inProgressRecords = unitRecords.filter((record) => isInProgressMetricEligible(record));
      const due10 = unitRecords.filter((record) => isDueInTenDays(record)).length;
      const due5 = unitRecords.filter((record) => isDueInFiveDays(record)).length;
      const dueToday = unitRecords.filter((record) => isDueToday(record)).length;
      const overdue = unitRecords.filter((record) => isOverdue(record)).length;
      const durations = inProgressRecords.map((record) => daysInProgress(record)).filter((days) => Number.isFinite(days));
      const avgDays = durations.length ? Math.round(durations.reduce((sum, days) => sum + days, 0) / durations.length) : 0;

      const monthUnitScope = monthScopeRecords.filter((record) => recordHasWorkUnit(record, workUnit));
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
  const monthScopeRecords = getCurrentMonthMetricScopeRecords();

  const due10 = records.filter((record) => isDueInTenDays(record)).length;
  const due5 = records.filter((record) => isDueInFiveDays(record)).length;
  const dueToday = records.filter((record) => isDueToday(record)).length;
  const overdue = records.filter((record) => isOverdue(record)).length;
  const respondedWorkUnits = new Set(
    records
      .filter((record) => statusEquals(record.status, "4. WORK UNIT RESPONDED"))
      .flatMap((record) => getWorkUnits(record))
      .filter(Boolean)
  );
  const completedThisMonth = monthScopeRecords.filter((record) => isCompletedStatus(record.status) && isWithinMonth(record.closed_date, new Date())).length;

  elements.commandAlerts.innerHTML = `
    <li class="alert-item"><div class="alert-title">New Requests Awaiting Action</div><div class="alert-value">${countByStatus(records, "1. NEW")}</div></li>
    <li class="alert-item"><div class="alert-title">Requests Currently In Progress</div><div class="alert-value">${countByStatus(records, "2. IN PROGRESS")}</div></li>
    <li class="alert-item"><div class="alert-title">Requests Pending With Legal</div><div class="alert-value">${countByStatus(records, "3. PENDING WITH LEGAL")}</div></li>
    <li class="alert-item"><div class="alert-title">Work Units That Have Responded</div><div class="alert-value">${respondedWorkUnits.size}</div></li>
    <li class="alert-item"><div class="alert-title">Due in 10 Days</div><div class="alert-value">${due10}</div></li>
    <li class="alert-item"><div class="alert-title">Due in 5 Days</div><div class="alert-value">${due5}</div></li>
    <li class="alert-item"><div class="alert-title">Due Today</div><div class="alert-value">${dueToday}</div></li>
    <li class="alert-item"><div class="alert-title">Overdue</div><div class="alert-value">${overdue}</div></li>
    <li class="alert-item"><div class="alert-title">Closed This Month</div><div class="alert-value">${completedThisMonth}</div></li>
  `;
}

function renderUpcomingDeadlines() {
  const upcoming = state.scopedRecords
    .filter((record) => isDeadlineEligible(record))
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
    elements.foiaTableBody.innerHTML = "<tr><td colspan=\"8\">No FOIA records match your search.</td></tr>";
    return;
  }

  elements.foiaTableBody.innerHTML = state.filteredRecords.map((record) => {
    const recordId = getRecordIdentity(record);
    const selectedClass = recordId === state.selectedId ? "selected" : "";
    const statusClass = `status-${statusClassSuffix(record.status)}`;
    const completed = isDciCompleted(record);
    const operationallyClosed = isOperationallyClosed(record);
    const daysOpen = daysInProgress(record);
    const daysToMilestone = daysToOperationalMilestone(record);
    const dueDelta = daysUntil(record.due_date);
    const timingText = operationallyClosed
      ? formatNumericMetric(daysToMilestone)
      : `${formatNumericMetric(daysOpen)}${hasValidDueDate(record) ? ` (${formatDueInShort(record, dueDelta)})` : ""}`;

    return `
      <tr data-record-id="${escapeHtml(recordId)}" class="${selectedClass}" tabindex="0">
        <td>${escapeHtml(record.request_id)}</td>
        <td>${escapeHtml(record.subject)}</td>
        <td>${escapeHtml(record.dci_work_unit)}</td>
        <td><span class="status-chip ${statusClass}">${escapeHtml(record.status)}</span></td>
        <td>${record.received_date ? formatDate(record.received_date) : ""}</td>
        <td>${operationallyClosed ? "" : (record.due_date ? formatDate(record.due_date) : "")}</td>
        <td>${completed && record.closed_date ? formatDate(record.closed_date) : ""}</td>
        <td>${timingText}</td>
      </tr>
    `;
  }).join("");

  Array.from(elements.foiaTableBody.querySelectorAll("tr[data-record-id]")).forEach((row) => {
    const select = () => {
      state.selectedId = row.getAttribute("data-record-id");
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
  const selected = state.filteredRecords.find((record) => getRecordIdentity(record) === state.selectedId);

  if (!selected) {
    elements.detailContent.innerHTML = "<p class=\"placeholder\">Select a FOIA request to view detailed information.</p>";
    return;
  }

  const completed = isDciCompleted(selected);
  const pendingWithLegal = isPendingWithLegal(selected);
  const operationallyClosed = isOperationallyClosed(selected);
  const dueInDays = daysUntil(selected.due_date);
  const milestoneDays = daysToOperationalMilestone(selected);
  const details = [
    ["FOIA Number", selected.request_id],
    ["Subject", selected.subject],
    ["DCI Work Unit", selected.dci_work_unit],
    ["Exact Status", selected.status]
  ];

  if (selected.received_date) {
    details.push(["Date DCI Received", formatDate(selected.received_date)]);
  }

  if (!operationallyClosed && selected.due_date) {
    details.push(["Due Date", `${formatDate(selected.due_date)} (${formatDueIn(selected, dueInDays)})`]);
  } else if (pendingWithLegal && selected.due_date) {
    details.push(["Due Date", formatDate(selected.due_date)]);
  }
  if (completed && selected.closed_date) {
    details.push(["Date Closed", formatDate(selected.closed_date)]);
  }

  if (pendingWithLegal) {
    details.push(["Days to Pending Legal", formatNumericMetric(milestoneDays)]);
  } else if (completed) {
    details.push(["Days to Complete", formatNumericMetric(milestoneDays)]);
  } else {
    details.push(["Days In Progress", formatNumericMetric(daysInProgress(selected))]);
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
  const workUnits = getWorkUnits(record);
  const workUnit = getWorkUnitDisplay(workUnits);
  return {
    ...record,
    id: String(record.id || record.item_id || "").trim(),
    request_id: String(record.request_id || "").trim() || "Not Assigned",
    workUnits,
    workUnit,
    dci_work_unit: workUnit,
    status: normalizeStatus(record.status),
    received_date: parseDate(record.received_date),
    date_stamped: parseDate(record.date_stamped),
    created_date: parseDate(record.created_date),
    due_date: parseDate(record.due_date),
    last_update: parseDate(record.last_update),
    closed_date: record.closed_date ? parseDate(record.closed_date) : null,
    pending_with_legal_date: parseDate(record.pending_with_legal_date)
  };
}

function normalizeWorkUnit(value) {
  return getWorkUnitDisplay(getWorkUnits(value));
}

function logFoiaNumberValidationSummary(items, records) {
  const recordsWithPopulatedTitle = items.filter((item) => {
    const title = normalizeGraphFieldText(item?.fields?.Title);
    return Boolean(title);
  }).length;
  const blankTitleRecords = items.filter((item) => !normalizeGraphFieldText(item?.fields?.Title)).length;
  const recordsShowingNotAssigned = records.filter((record) => String(record.request_id || "").trim() === "Not Assigned").length;
  const recordsUsingGeneratedPlaceholders = records.filter((record) => /^ITEM-\d+$/i.test(String(record.request_id || "").trim())).length;

  console.info("FOIA number validation summary", {
    totalRecords: items.length,
    recordsWithTitlePopulated: recordsWithPopulatedTitle,
    blankTitleRecords,
    recordsShowingNotAssigned,
    recordsStillUsingItemGeneratedPlaceholders: recordsUsingGeneratedPlaceholders
  });
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
    return PENDING_WITH_LEGAL_STATUS;
  }
  if (normalized === "open" || normalized === "pending" || normalized === "active") {
    return "2. IN PROGRESS";
  }
  return value;
}

function statusKey(value) {
  return String(value || "")
    .trim()
    .toLowerCase()
    .replace(/^(\d+)\s*\.\s*/, "$1. ")
    .replace(/\s+/g, " ");
}

function statusEquals(a, b) {
  return statusKey(a) === statusKey(b);
}

function normalizeStatusForComparison(status) {
  return String(status || "")
    .trim()
    .toLowerCase()
    .replace(/^\d+\s*\.\s*/, "")
    .replace(/\s+/g, " ");
}

function startOfToday() {
  return today;
}

function isDciCompleted(record) {
  const status = typeof record === "object" && record !== null ? record.status : record;
  return normalizeStatusForComparison(status) === "dci completed";
}

function isPendingWithLegal(record) {
  const status = typeof record === "object" && record !== null ? record.status : record;
  return normalizeStatusForComparison(status) === "pending with legal";
}

function isOperationallyClosed(record) {
  return isPendingWithLegal(record) || isDciCompleted(record);
}

function isOpenStageStatus(status) {
  return Array.from(OPEN_STAGE_STATUSES).some((knownStatus) => statusEquals(status, knownStatus));
}

function isCompletedStatus(status) {
  return isDciCompleted(status);
}

function isDeadlineEligible(record) {
  return hasValidDueDate(record) && !isOperationallyClosed(record);
}

function isDueToday(record) {
  return isDeadlineEligible(record) && daysUntil(record.due_date) === 0;
}

function isDueInFiveDays(record) {
  if (!isDeadlineEligible(record)) {
    return false;
  }
  const delta = daysUntil(record.due_date);
  return delta >= 1 && delta <= 5;
}

function isDueInTenDays(record) {
  if (!isDeadlineEligible(record)) {
    return false;
  }
  const delta = daysUntil(record.due_date);
  return delta >= 6 && delta <= 10;
}

function isOverdue(record) {
  const dueDate = parseDate(record?.dueDate || record?.due_date);
  return Boolean(dueDate) && dueDate < startOfToday() && !isOperationallyClosed(record);
}

function isInProgressMetricEligible(record) {
  return !isOperationallyClosed(record);
}

function getIntakeDate(record) {
  return record.received_date || record.date_stamped || record.created_date || null;
}

function getProgressStartDate(record) {
  return record.received_date || record.date_stamped || record.created_date || null;
}

function getRecordTimeFilterDate(record) {
  if (isDciCompleted(record)) {
    return record.closed_date || null;
  }
  return getIntakeDate(record);
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

function getPendingWithLegalDate(record) {
  const explicitPendingWithLegalDate = record?.pending_with_legal_date;
  if (explicitPendingWithLegalDate instanceof Date && !Number.isNaN(explicitPendingWithLegalDate.getTime())) {
    return explicitPendingWithLegalDate;
  }

  if (PENDING_WITH_LEGAL_FALLBACK_TO_MODIFIED) {
    const modifiedDate = record?.last_update;
    if (modifiedDate instanceof Date && !Number.isNaN(modifiedDate.getTime())) {
      return modifiedDate;
    }
  }

  return null;
}

function daysToOperationalMilestone(record) {
  const opened = getProgressStartDate(record);
  if (!(opened instanceof Date) || Number.isNaN(opened.getTime())) {
    return Number.NaN;
  }

  if (isDciCompleted(record)) {
    const closed = record.closed_date;
    if (!(closed instanceof Date) || Number.isNaN(closed.getTime())) {
      return Number.NaN;
    }
    return Math.max(0, daysBetween(opened, closed));
  }

  if (isPendingWithLegal(record)) {
    const pendingWithLegalDate = getPendingWithLegalDate(record);
    if (!(pendingWithLegalDate instanceof Date) || Number.isNaN(pendingWithLegalDate.getTime())) {
      return Number.NaN;
    }
    return Math.max(0, daysBetween(opened, pendingWithLegalDate));
  }

  return Number.NaN;
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
  if (!isInProgressMetricEligible(record)) {
    return Number.NaN;
  }
  const startDate = getProgressStartDate(record);
  if (!(startDate instanceof Date) || Number.isNaN(startDate.getTime())) {
    return Number.NaN;
  }
  return Math.max(0, daysBetween(startDate, today));
}

function getCurrentMonthMetricScopeRecords() {
  return state.records.filter((record) => {
    if (state.selectedWorkUnit !== ALL_WORK_UNITS_OPTION && !recordHasWorkUnit(record, state.selectedWorkUnit)) {
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
  const currentMonthEnd = getCurrentMonthEnd();

  const baseRecords = state.records.filter((record) => {
    if (state.selectedWorkUnit !== ALL_WORK_UNITS_OPTION && !recordHasWorkUnit(record, state.selectedWorkUnit)) {
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
      if (
        recordMatchesDateRangeByDate(receivedDate, periodRange.start, periodRange.end) &&
        recordMatchesDateRangeByDate(receivedDate, customRange.start, customRange.end) &&
        recordMatchesDateRangeByDate(receivedDate, null, currentMonthEnd)
      ) {
        upsertBucket(receivedDate).received += 1;
      }
    }

    if (isCompletedStatus(record.status) && record.closed_date instanceof Date && !Number.isNaN(record.closed_date.getTime())) {
      if (
        recordMatchesDateRangeByDate(record.closed_date, periodRange.start, periodRange.end) &&
        recordMatchesDateRangeByDate(record.closed_date, customRange.start, customRange.end) &&
        recordMatchesDateRangeByDate(record.closed_date, null, currentMonthEnd)
      ) {
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

function getCurrentMonthEnd() {
  return startOfDay(new Date(today.getFullYear(), today.getMonth() + 1, 0));
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
  return daysBetween(startOfToday(), date);
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

function formatDueIn(record, days) {
  if (!Number.isFinite(days)) {
    return "Due date unavailable";
  }
  if (isOverdue(record)) {
    return `${Math.abs(days)} day${Math.abs(days) === 1 ? "" : "s"} overdue`;
  }
  if (days === 0) {
    return "Due today";
  }
  return `Due in ${days} day${days === 1 ? "" : "s"}`;
}

function formatDueInShort(record, days) {
  if (!Number.isFinite(days)) {
    return "N/A";
  }
  if (isOverdue(record)) {
    return `${Math.abs(days)} overdue`;
  }
  if (days === 0) {
    return "Due today";
  }
  return String(days);
}

function logDueDateValidationSummary(records) {
  const recordsWithValidDueDates = records.filter((record) => hasValidDueDate(record)).length;
  const dciCompletedRecords = records.filter((record) => isDciCompleted(record)).length;
  const pendingWithLegalRecords = records.filter((record) => isPendingWithLegal(record)).length;
  const nonCompletedPastDueRecords = records.filter((record) => {
    if (!hasValidDueDate(record) || isOperationallyClosed(record)) {
      return false;
    }
    return parseDate(record.due_date) < startOfToday();
  }).length;
  const finalOverdueCount = records.filter((record) => isOverdue(record)).length;

  console.info("Due date validation summary", {
    totalRecordsWithValidDueDates: recordsWithValidDueDates,
    recordsWithStatusPendingWithLegal: pendingWithLegalRecords,
    recordsWithStatusDciCompleted: dciCompletedRecords,
    nonOperationallyClosedRecordsWithPastDueDates: nonCompletedPastDueRecords,
    finalOverdueCount
  });
}

function logOperationalClosureValidationSummary(records) {
  const pendingWithLegalRecords = records.filter((record) => isPendingWithLegal(record));
  const dciCompletedRecords = records.filter((record) => isDciCompleted(record));
  const operationallyClosedRecords = records.filter((record) => isOperationallyClosed(record));
  const overdueAfterExclusions = records.filter((record) => isOverdue(record));

  const upcomingEligibleRecords = records.filter((record) => isDeadlineEligible(record));
  const pendingWithLegalRemovedFromUpcoming = pendingWithLegalRecords.filter((record) => hasValidDueDate(record)).length;
  const pendingWithLegalWithValidDate = pendingWithLegalRecords.filter((record) => {
    const pendingWithLegalDate = getPendingWithLegalDate(record);
    return pendingWithLegalDate instanceof Date && !Number.isNaN(pendingWithLegalDate.getTime());
  }).length;
  const pendingWithLegalMissingDate = pendingWithLegalRecords.length - pendingWithLegalWithValidDate;

  console.info("Operational status validation summary", {
    pendingWithLegalRecords: pendingWithLegalRecords.length,
    dciCompletedRecords: dciCompletedRecords.length,
    operationallyClosedRecords: operationallyClosedRecords.length,
    overdueAfterOperationalClosedExclusions: overdueAfterExclusions.length,
    pendingWithLegalRemovedFromUpcomingDeadlines: pendingWithLegalRemovedFromUpcoming,
    pendingWithLegalRecordsWithValidPendingLegalDate: pendingWithLegalWithValidDate,
    pendingWithLegalRecordsMissingPendingLegalDate: pendingWithLegalMissingDate,
    upcomingDeadlineEligibleRecordsAfterExclusions: upcomingEligibleRecords.length
  });
}

function escapeHtml(value) {
  return String(value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/\"/g, "&quot;")
    .replace(/'/g, "&#39;");
}
