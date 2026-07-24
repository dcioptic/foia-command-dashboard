window.SHAREPOINT_GRAPH_CONFIG = {
  msalConfig: {
    auth: {
      clientId: "73d22f27-20f1-491a-9476-0075390d474f",
      authority: "https://login.microsoftonline.com/6b22089f-8eba-463c-84f5-3171210f9005",
      redirectUri: window.location.origin + window.location.pathname
    },
    cache: {
      cacheLocation: "sessionStorage"
    }
  }
};
