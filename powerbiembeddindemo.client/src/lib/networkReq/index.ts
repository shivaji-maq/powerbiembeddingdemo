import axios from "axios";
/**
 * Network request utilities for backend and external API calls.
 *
 * @module networkReq
 *
 * @remarks
 * - Provides a pre-configured Axios instance for backend API requests with authorization headers.
 * - Exports the default Axios instance for external requests.
 */
/**
 * Network Request Library
 * This library provides functions to make network requests to the backend API.
 */
export const networkReq = axios.create({
  baseURL: "",
});

export const networkReqToExternal = axios;
