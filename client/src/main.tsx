import { createRoot } from "react-dom/client";
import App from "./App";
import "./index.css";

const root = createRoot(document.getElementById("root")!);

const renderApp = () => {
  root.render(<App />);
};

if (typeof Office !== "undefined") {
  Office.onReady(renderApp);
} else {
  window.addEventListener("load", () => {
    if (typeof Office !== "undefined") {
      Office.onReady(renderApp);
    } else {
      renderApp();
    }
  });
}
