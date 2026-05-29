import { configureStore } from "@reduxjs/toolkit";
import rootReducer from "./slices";
/**
 * @module redux/store
 *
 * Configures and exports the Redux store for the application.
 * Sets up the root reducer, middleware, Redux DevTools, and preloaded state.
 *
 * @remarks
 * The store is the central state container for the app. Use {@link RootState} and {@link AppDispatch} for type safety in selectors and dispatch calls.
 *
 * @example
 * import store, { RootState, AppDispatch } from './redux/store';
 * store.dispatch(someAction());
 */
const store = configureStore({
  reducer: rootReducer,
  middleware: (getDefaultMiddleware) => getDefaultMiddleware(),
  // devTools: process.env.NODE_ENV !== "production",
  preloadedState: {},
});

export default store;
/**
 * The root state type derived from the store's reducer.
 * @typedef {object} RootState
 */
export type RootState = ReturnType<typeof store.getState>;
/**
 * The dispatch type for the Redux store.
 * @typedef {function} AppDispatch
 */
export type AppDispatch = typeof store.dispatch;
