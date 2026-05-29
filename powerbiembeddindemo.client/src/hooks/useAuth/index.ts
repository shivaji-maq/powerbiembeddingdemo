// src/hooks/useAuth.ts
import { useSelector, useDispatch } from "react-redux";
import type { RootState, AppDispatch } from "../../redux/store";
import { login, logout, setUser } from "../../redux/slices/authSlice/authSlice";
import { useEffect, useState } from "react";
import { getMsalInstance, initializeMsal } from "../../configs/msalInstance";
import type { SerializableAccount, User } from "../../redux/slices/authSlice/@types";
/**
 * Custom React hook for authentication state and actions using MSAL and Redux.
 *
 * Initializes MSAL on mount, handles redirect responses, and provides helpers for login, logout, and setting user/account.
 *
 * @returns {object} An object containing authentication state and actions:
 * - `user`: The current authenticated user (or null).
 * - `account`: The current MSAL account (or null).
 * - `isLoading`: Whether authentication is in progress.
 * - `error`: Any authentication error.
 * - `loginUser`: Function to trigger login.
 * - `logoutUser`: Function to trigger logout.
 * - `setAuthUser`: Function to set the user and account in Redux.
 * - `isAuthenticated`: Boolean indicating if the user is authenticated and MSAL is ready.
 *
 * @example
 * ```typescript
 * const { user, account, isAuthenticated, loginUser, logoutUser } = useAuth();
 * if (!isAuthenticated) loginUser();
 * ```
 *
 * @see {@link User}, {@link SerializableAccount}
 */
export const useAuth = () => {
    const dispatch = useDispatch<AppDispatch>();
    const { user, account, isLoading, error } = useSelector((state: RootState) => state.auth);
    const [msalReady, setMsalReady] = useState(false);

    useEffect(() => {
        const initMsal = async () => {
            await initializeMsal();
            console.log("[useAuth] MSAL initialized");

            const instance = getMsalInstance();
            console.log("[useAuth] Handling redirect promise");
            instance
                .handleRedirectPromise()
                .then(response => {
                    console.log("[useAuth] Redirect promise resolved with response:", !!response, response?.account?.username);
                    if (response && response.account) {
                        instance.setActiveAccount(response.account);
                        console.log("[useAuth] Active account set to:", response.account.username);
                    } else {
                        const allAccounts = instance.getAllAccounts();
                        console.log("[useAuth] No redirect response, checking for existing accounts:", allAccounts.length);
                        if (allAccounts.length > 0) {
                            instance.setActiveAccount(allAccounts[0]);
                            console.log("[useAuth] Active account set to existing:", allAccounts[0].username);
                        }
                    }
                })
                .catch(error => {
                    console.error("[useAuth] Error handling redirect:", error);
                })
                .finally(() => {
                    console.log("[useAuth] MSAL ready");
                    setMsalReady(true);
                });
        };
        initMsal();
    }, []);

    useEffect(() => {
        if (msalReady && !(user && account)) {
            console.log("[useAuth] MSAL ready but user/account missing, dispatching login");
            dispatch(login());
        }
    }, [msalReady, user, account, dispatch]);

    const loginUser = async () => {
        await dispatch(login());
    };

    const logoutUser = async () => {
        await dispatch(logout());
    };

    const setAuthUser = (user: User | null, account: SerializableAccount | null) => {
        dispatch(setUser({ user, account }));
    };

    return {
        user,
        account,
        isLoading,
        error,
        loginUser,
        logoutUser,
        setAuthUser,
        isAuthenticated: msalReady && user && account ? true : false,
    };
};
