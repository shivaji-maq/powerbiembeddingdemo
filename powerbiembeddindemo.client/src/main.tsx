import { StrictMode } from "react";
import { createRoot, type Container } from "react-dom/client";
import "./index.css";
import App from "./App.jsx";

import { Provider } from "react-redux";
import store from "./redux/store.tsx";

createRoot(document.getElementById("root") as Container).render(
  <StrictMode>
    <Provider store={store}>
      <App />
    </Provider>
  </StrictMode>,
);
