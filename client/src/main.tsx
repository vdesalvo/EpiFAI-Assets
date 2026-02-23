import { createRoot } from "react-dom/client";
import App from "./App";
import "./index.css";

const root = createRoot(document.getElementById("root")!);

const renderApp = () => {
  try {
    root.render(<App />);
  } catch (e) {
    console.error("Failed to render app:", e);
    document.getElementById("root")!.innerHTML = `<div style="padding:20px;font-family:sans-serif"><h3>Failed to load</h3><p>${e}</p></div>`;
  }
};

const initWithTimeout = () => {
  let rendered = false;

  const doRender = () => {
    if (rendered) return;
    rendered = true;
    renderApp();
  };

  if (typeof Office !== "undefined" && Office.onReady) {
    Office.onReady(() => doRender());
    setTimeout(doRender, 5000);
  } else {
    doRender();
  }
};

if (document.readyState === "loading") {
  document.addEventListener("DOMContentLoaded", initWithTimeout);
} else {
  initWithTimeout();
}
