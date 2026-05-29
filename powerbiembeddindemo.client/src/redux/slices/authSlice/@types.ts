/**
 * @file  This file contains TypeScript interfaces that define the structure of user and authentication state data.
 * These interfaces are used to manage user information and authentication state in the application.
 */

/**
 * Represents a user in the authentication system.
 *
 * @property {string} id - The unique identifier for the user.
 * @property {string} email - The user's email address.
 * @property {string} displayName - The user's display name.
 * @property {string} givenName - The user's given name.
 * @property {string} surname - The user's surname.
 * @property {string} userPrincipalName - The user's principal name.
 * @property {string | undefined} profileImage - The user's profile image URL (optional).
 * @property {string} [jobTitle] - The user's job title (optional).
 */
export interface User {
    id: string;
    email: string;
    displayName: string;
    givenName: string;
    surname: string;
    userPrincipalName: string;
    profileImage: string | undefined;
    jobTitle?: string;
}

/**
 * Represents the structure of an account information object (serializable version of AccountInfo).
 *
 * @property {string} homeAccountId - The home account ID.
 * @property {string} localAccountId - The local account ID.
 * @property {string} environment - The environment for the account.
 * @property {string} tenantId - The tenant ID.
 * @property {string} username - The username for the account.
 */
export interface SerializableAccount {
    homeAccountId: string;
    localAccountId: string;
    environment: string;
    tenantId: string;
    username: string;
}

/**
 * Represents the structure of the authentication state in the application.
 *
 * @property {User | null} user - The current authenticated user, or null if not authenticated.
 * @property {SerializableAccount | null} account - The current account, or null if not authenticated.
 * @property {boolean} isLoading - Whether authentication is in progress.
 * @property {string | null} error - Any authentication error message.
 */
export interface AuthState {
    user: User | null;
    account: SerializableAccount | null;
    isLoading: boolean;
    error: string | null;
}
