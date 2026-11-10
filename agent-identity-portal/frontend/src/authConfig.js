import { LogLevel } from "@azure/msal-browser";

export const msalConfig = {
  auth: {
    clientId: process.env.REACT_APP_AZURE_CLIENT_ID || "",
    authority: `https://login.microsoftonline.com/${process.env.REACT_APP_AZURE_TENANT_ID || "common"}`,
    redirectUri: process.env.REACT_APP_AZURE_REDIRECT_URI || "http://localhost:3000",
  },
  cache: {
    cacheLocation: "sessionStorage",
    storeAuthStateInCookie: false,
  },
  system: {
    loggerOptions: {
      logLevel: LogLevel.Warning,
    },
  },
};

const apiClientId = process.env.REACT_APP_API_CLIENT_ID || "";

export const loginRequest = {
  scopes: [`api://${apiClientId}/access_as_user`],
};

export const apiScopes = {
  scopes: [`api://${apiClientId}/access_as_user`],
};
