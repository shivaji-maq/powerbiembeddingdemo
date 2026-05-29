import AuthReducer from "./authSlice/authSlice";
import { combineReducers } from "redux";
import powerBIReducer from "./powerBISlice/powerBISlice";
import bookmarkReducer from "./bookmarkSlice/bookmarkSlice";
/**
 * @module redux/slices/index
 *
 * Combines all Redux slices into a single root reducer for the application state.
 * Imports individual reducers and combines them using Redux's {@link combineReducers} function.
 *
 * @remarks
 * This root reducer is typically used when configuring the Redux store.
 *
 * @example
 * import rootReducer from './redux/slices';
 * const store = configureStore({ reducer: rootReducer });
 */
const rootReducer = combineReducers({
  auth: AuthReducer,
  powerBI: powerBIReducer,
  bookmarks: bookmarkReducer,
});

export default rootReducer;
