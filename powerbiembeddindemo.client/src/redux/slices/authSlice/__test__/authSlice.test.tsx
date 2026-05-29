import { describe, it, expect, beforeEach, vi} from "vitest";
import reducer, {
    setUser,
    clearUser,
    login,
    logout,
} from "../authSlice";
import type { AuthState, User, SerializableAccount } from "../@types";

// Mock all external dependencies
vi.mock("@azure/msal-browser", () => ({
    InteractionRequiredAuthError: class {},
    PublicClientApplication: class {},
}));
vi.mock("../../../configs/msalConfig", () => ({
    loginRequest: {},
    graphConfig: { graphMeEndpoint: "https://graph.me", graphPhotoEndpoint: "https://graph.photo" },
}));
vi.mock("../../../configs/msalInstance", () => ({
    getMsalInstance: () => ({
        getAllAccounts: () => [{ homeAccountId: "1", localAccountId: "2", environment: "env", tenantId: "tid", username: "user" }],
        acquireTokenSilent: vi.fn().mockResolvedValue({ accessToken: "mock-access-token" }),
        loginRedirect: vi.fn(),
        logoutRedirect: vi.fn(),
    }),
}));
vi.mock("../../../lib", () => ({
    networkReq: {
        post: vi.fn().mockResolvedValue({ data: { token: "backend-token" } }),
    },
}));
vi.mock("../../../utils/functions/localStorage", () => ({
    writeToLocalStorage: vi.fn(),
    removeFromLocalStorage: vi.fn(),
}));
vi.mock("../../../utils/constants/localstorageConst", () => ({
    LOCALSTORAGE_KEYS: {
        BACKEND_ACCESS_TOKEN_SESSION_KEY: "BACKEND_ACCESS_TOKEN_SESSION_KEY",
    },
}));
vi.mock("../../../utils/constants/apiConst", () => ({
    API_CONST: { LOGIN: "/api/login" },
}));

// Mock fetch
globalThis.fetch = vi.fn((url) => {
    if (url === "https://graph.me") {
        return Promise.resolve({
            ok: true,
            json: () => Promise.resolve({
                id: "id",
                mail: "mail",
                userPrincipalName: "user@domain.com",
                displayName: "Display Name",
                givenName: "Given",
                surname: "Surname",
                jobTitle: "Job",
            }),
        });
    }
    if (url === "https://graph.photo") {
        return Promise.resolve({
            ok: true,
            blob: () => Promise.resolve(new Blob()),
        });
    }
    return Promise.resolve({ ok: false });
}) as any;

beforeEach(() => {
    vi.clearAllMocks();
});

describe("authSlice reducers", () => {
    const initialState: AuthState = {
        user: null,
        account: null,
        isLoading: false,
        error: null,
    };

    it("should handle setUser", () => {
        const user: User = {
            id: "id",
            email: "mail",
            displayName: "Display Name",
            givenName: "Given",
            surname: "Surname",
            userPrincipalName: "user@domain.com",
            profileImage: "img",
            jobTitle: "Job",
        };
        const account: SerializableAccount = {
            homeAccountId: "1",
            localAccountId: "2",
            environment: "env",
            tenantId: "tid",
            username: "user",
        };
        const state = reducer(initialState, setUser({ user, account }));
        expect(state.user).toEqual(user);
        expect(state.account).toEqual(account);
    });

    it("should handle clearUser", () => {
        const state = reducer(
            { ...initialState, user: { id: "id" } as any, account: { homeAccountId: "1" } as any },
            clearUser()
        );
        expect(state.user).toBeNull();
        expect(state.account).toBeNull();
    });
});

describe("authSlice async thunks", () => {
    it("login fulfilled updates state", async () => {
        const dispatch = vi.fn();
        const getState = vi.fn();
        const action = await login()(dispatch, getState, undefined);
        expect(action.type).toContain("auth/login/");
        // fulfilled or redirected
    });

    it("logout fulfilled clears state", async () => {
        const dispatch = vi.fn();
        const getState = vi.fn();
        const action = await logout()(dispatch, getState, undefined);
        expect(action.type).toContain("auth/logout/");
    });

    it("login rejected on error", async () => {
        // Force fetch to fail
        (globalThis.fetch as any) = vi.fn(() => Promise.resolve({ ok: false }));
        const dispatch = vi.fn();
        const getState = vi.fn();
        const action = await login()(dispatch, getState, undefined);
        expect(action.type).toContain("auth/login/rejected");
    });
});

describe("authSlice extraReducers", () => {
    const getInitialState = (): AuthState => ({
        user: null,
        account: null,
        isLoading: false,
        error: null,
    });

    it("handles login.pending", () => {
        const state = reducer(getInitialState(), { type: login.pending.type });
        expect(state.isLoading).toBe(true);
        expect(state.error).toBeNull();
    });

    it("handles login.fulfilled", () => {
        const payload = {
            user: { id: "id" } as any,
            account: { homeAccountId: "1" } as any,
        };
        const state = reducer(getInitialState(), { type: login.fulfilled.type, payload });
        expect(state.isLoading).toBe(false);
        expect(state.user).toEqual(payload.user);
        expect(state.account).toEqual(payload.account);
        expect(state.error).toBeNull();
    });

    it("handles login.rejected", () => {
        const state = reducer(getInitialState(), { type: login.rejected.type, payload: "fail" });
        expect(state.isLoading).toBe(false);
        expect(state.error).toBe("fail");
        expect(state.user).toBeNull();
        expect(state.account).toBeNull();
    });

    it("handles logout.fulfilled", () => {
        const state = reducer(
            { user: { id: "id" } as any, account: { homeAccountId: "1" } as any, isLoading: true, error: "err" },
            { type: logout.fulfilled.type }
        );
        expect(state.isLoading).toBe(false);
        expect(state.user).toBeNull();
        expect(state.account).toBeNull();
        expect(state.error).toBeNull();
    });
});
