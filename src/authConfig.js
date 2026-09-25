const tenantId = "d7cbbb08-47a3-4bd7-8347-5018f2744cfb";
const clientId = "78ca0754-7f04-47cc-aedd-9c6b9d44e9c4";
const redirectUri = new URL(import.meta.env.BASE_URL, window.location.origin).href;

export const isAuthConfigured =
  tenantId !== "YOUR_TENANT_ID" && clientId !== "YOUR_CLIENT_ID";

export const msalConfig = {
  auth: {
    clientId,
    authority: `https://login.microsoftonline.com/${tenantId}`,
    redirectUri,
    postLogoutRedirectUri: redirectUri,
  },
  cache: {
    cacheLocation: "sessionStorage",
  },
};

export const loginRequest = {
  scopes: ["User.Read"],
};