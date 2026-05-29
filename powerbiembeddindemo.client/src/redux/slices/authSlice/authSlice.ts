/**
 * Redux slice for authentication, including async actions for login and logout using MSAL and Microsoft Graph.
 *
 * @module authSlice
 *
 * @remarks
 * - Handles authentication state, user info, and account info.
 * - Uses MSAL for Azure AD authentication and Microsoft Graph for user data.
 * - Persists backend access token in local storage after login.
 */

import { createSlice, createAsyncThunk } from "@reduxjs/toolkit";
import { InteractionRequiredAuthError } from "@azure/msal-browser";
import type { AccountInfo } from "@azure/msal-browser";
import { loginRequest, graphConfig } from "../../../configs/msalConfig";
import { getMsalInstance } from "../../../configs/msalInstance";
import type { AuthState, SerializableAccount, User } from "./@types";

// Helper function to convert AccountInfo to SerializableAccount
/**
 * Helper function to convert an MSAL AccountInfo object to a serializable account structure.
 *
 * @param {AccountInfo} account - The MSAL account information.
 * @returns {SerializableAccount} The serializable account object.
 */
const serializeAccount = (account: AccountInfo): SerializableAccount => {
  return {
    homeAccountId: account.homeAccountId,
    localAccountId: account.localAccountId,
    environment: account.environment,
    tenantId: account.tenantId,
    username: account.username,
  };
};
/**
 * Async thunk for logging in the user using MSAL and Microsoft Graph.
 *
 * @function
 * @async
 * @returns {Promise<{ user: User; account: SerializableAccount } | undefined>} The user and account info, or undefined if redirected.
 * @throws {string} If login fails or access token is not obtained.
 */
export const login = createAsyncThunk<{ user: User; account: SerializableAccount } | undefined, void, { rejectValue: string }>(
  "auth/login",
  async (_, { rejectWithValue }) => {
    try {
      console.log("[login] Starting login flow");
      const instance = getMsalInstance();
      const account = instance.getAllAccounts()[0];
      console.log("[login] Accounts found:", instance.getAllAccounts().length, "Current account:", account?.username);
      let accessToken: string | undefined;

      try {
        if (account) {
          console.log("[login] Found account, attempting silent token acquisition");
          const silentRequest = {
            ...loginRequest,
            account,
          };
          const response = await instance.acquireTokenSilent(silentRequest);
          accessToken = response.accessToken;
          console.log("[login] Silent token acquisition succeeded");
        } else {
          console.log("[login] No account found, redirecting to login");
          await instance.loginRedirect(loginRequest);
          return undefined;
        }
      } catch (error) {
        console.error("[login] Silent token acquisition failed:", error);
        if (error instanceof InteractionRequiredAuthError) {
          console.log("[login] InteractionRequired error, redirecting to login");
          await instance.loginRedirect(loginRequest);
          return undefined;
        } else {
          throw error;
        }
      }

      if (!accessToken) {
        throw new Error("Access token not obtained");
      }

      console.log("[login] Getting user info from Microsoft Graph");
      // Get user info from Microsoft Graph
      const response = await fetch(graphConfig.graphMeEndpoint, {
        headers: {
          Authorization: `Bearer ${accessToken}`,
        },
      });

      if (!response.ok) {
        console.error("[login] Failed to fetch user data, status:", response.status);
        throw new Error("Failed to fetch user data");
      }

      const userData = await response.json();
      console.log("[login] User data received:", { displayName: userData.displayName, email: userData.mail });
      const currentAccount = instance.getAllAccounts()[0];

      // Fetch profile image as a data URL
      let profileImageUrl: string | undefined = undefined;
      try {
        const photoResponse = await fetch(graphConfig.graphPhotoEndpoint, {
          headers: {
            Authorization: `Bearer ${accessToken}`,
          },
        });
        if (photoResponse.ok) {
          const blob = await photoResponse.blob();
          const reader = new FileReader();
          profileImageUrl = await new Promise<string>((resolve, reject) => {
            reader.onloadend = () => resolve(reader.result as string);
            reader.onerror = reject;
            reader.readAsDataURL(blob);
          });
        }
      } catch (e: unknown) {
        console.log(e);
        profileImageUrl = undefined; // Ignore errors if photo is not available
      }

      // return the user and account information
      console.log("[login] Returning user and account data");
      return {
        user: {
          id: userData.id,
          email: userData.mail || userData.userPrincipalName,
          displayName: userData.displayName,
          givenName: userData.givenName,
          surname: userData.surname,
          userPrincipalName: userData.userPrincipalName,
          profileImage: profileImageUrl,
          jobTitle: userData.jobTitle,
        },
        account: serializeAccount(currentAccount),
      };
    } catch (error) {
      console.error("[login] Error in login thunk:", error);
      if (error instanceof Error) {
        return rejectWithValue(error.message);
      }
      return rejectWithValue("An unknown error occurred during login");
    }
  }
);
/**
 * Async thunk for logging out the user and clearing authentication state.
 *
 * @function
 * @async
 * @returns {Promise<void>} Resolves when logout is complete.
 * @throws {string} If logout fails.
 */
export const logout = createAsyncThunk<void, void, { rejectValue: string }>("auth/logout", async (_, { rejectWithValue }) => {
  try {
    const instance = getMsalInstance();
    const account = instance.getAllAccounts()[0];
    if (account) {
      await instance.logoutRedirect({
        account,
      });
    } else {
      await instance.logoutRedirect();
    }
  } catch (error) {
    if (error instanceof Error) {
      return rejectWithValue(error.message);
    }
    return rejectWithValue("An unknown error occurred during logout");
  }
});

/**
 * The initial authentication state.
 *
 * @type {AuthState}
 */
const initialState: AuthState = {
  user: null,
  account: null,
  isLoading: false,
  error: null,
};

/**
 * The authentication slice, including reducers and extraReducers for login/logout.
 *
 * @type {import('@reduxjs/toolkit').Slice}
 */
const AuthSlice = createSlice({
  name: "auth",
  initialState,
  reducers: {
    setUser: (
      state,
      action: {
        payload: { user: User | null; account: SerializableAccount | null };
      }
    ) => {
      state.user = action.payload.user;
      state.account = action.payload.account;
    },
    clearUser: (state) => {
      state.user = null;
      state.account = null;
    },
  },
  extraReducers: (builder) => {
    builder
      .addCase(login.pending, (state) => {
        console.log("[auth/login pending]");
        state.isLoading = true;
        state.error = null;
      })
      .addCase(login.fulfilled, (state, action) => {
        console.log("Login fulfilled with payload:", action.payload);
        state.isLoading = false;
        if (!action.payload) {
          state.error = "Login was redirected, please check your browser.";
          console.log("Login redirected - no payload returned");
          return;
        }
        state.user = action.payload.user;
        state.account = action.payload.account;
        state.error = null;
        console.log("Login successful - user and account set:", { user: action.payload.user.displayName, account: action.payload.account.username });
      })
      .addCase(login.rejected, (state, action) => {
        console.log("[auth/login rejected] Error:", action.payload);
        state.isLoading = false;
        state.error = action.payload ?? "An unknown error occurred";
        state.user = null;
        state.account = null;
      });

    builder
      .addCase(logout.pending, (state) => {
        state.isLoading = true;
        state.error = null;
      })
      .addCase(logout.fulfilled, (state) => {
        state.isLoading = false;
        state.user = null;
        state.account = null;
        state.error = null;
      })
      .addCase(logout.rejected, (state, action) => {
        state.isLoading = false;
        state.error = action.payload ?? "An unknown error occurred";
      });
  },
});

export const { setUser, clearUser } = AuthSlice.actions;
export default AuthSlice.reducer;
