import { PublicClientApplication } from "@azure/msal-browser";
import { msalConfig, powerBiRequest } from "../msalConfig";
import { CLIENT_SECRET, MSAL_CLIENT_ID, MSAL_TENANT_ID } from "../../constants/servicePrinciple";
/**
 * MSAL instance configuration and authentication helpers for Azure AD and Power BI API.
 *
 * @module msalInstance
 *
 * @remarks
 * - Initializes the MSAL PublicClientApplication with the provided configuration.
 * - Provides functions to initialize MSAL, retrieve the MSAL instance, and acquire access tokens for Power BI API.
 */
const msalInstance = new PublicClientApplication(msalConfig);
/**
 * Initializes the MSAL instance asynchronously. Should be called at app startup.
 *
 * @returns {Promise<void>} Resolves when initialization is complete.
 */
export const initializeMsal = async () => {
  try {
    await msalInstance.initialize();
  } catch (error) {
    console.error("Error initializing MSAL:", error);
  }
};
/**
 * Returns the initialized MSAL instance for authentication operations.
 *
 * @returns {PublicClientApplication} The MSAL instance used for authentication.
 */
export const getMsalInstance = () => msalInstance;
/**
 * Retrieves the access token for Power BI API using MSAL.
 *
 * @async
 * @returns {Promise<string | null>} The access token if available, or null if no account is signed in.
 * @throws Will throw if both silent and popup token acquisition fail.
 */
export const getAccessToken = async () => {
  const accounts = msalInstance.getAllAccounts();

  if (accounts.length === 0) {
    console.warn("No accounts found. User not signed in.");
    return null;
  }

  const request = {
    scopes: powerBiRequest.scopes,
    account: accounts[0],
  };

  try {
    const response = await msalInstance.acquireTokenSilent(request);
    return response.accessToken;
  } catch (silentError) {
    console.warn("Silent token acquisition failed:", silentError);

    try {
      const response = await msalInstance.acquireTokenPopup(request);
      // Only log access tokens in non-production environments for debugging.

      return response.accessToken;
    } catch (popupError) {
      console.error("Popup token acquisition failed:", popupError);
      throw popupError; // Let the caller decide how to handle it
    }
  }
};
export const getAppModeAccessKey = async () => {
  return await getAppModeAccessKeyUtil(MSAL_TENANT_ID, MSAL_CLIENT_ID, CLIENT_SECRET, "https://analysis.windows.net/powerbi/api/.default");
};

export const getAppModeAccessKeyUtil = async (tenantId: string, clientId: string, clientSecret: string, scope = "https://management.azure.com/.default") => {
  const url = `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`;

  const params = new URLSearchParams();
  params.append("grant_type", "client_credentials");
  params.append("client_id", clientId);
  params.append("client_secret", clientSecret);
  params.append("scope", scope);

  const response = await fetch(url, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: params,
  });

  if (!response.ok) {
    const error = await response.text();
    throw new Error(`Failed to get access token: ${error}`);
  }

  const data = await response.json();
  return data.access_token; // ⬅️ This is your Microsoft Access Token
};
