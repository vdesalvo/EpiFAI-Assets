import { createRoot } from "react-dom/client";
import App from "./App";
import "./index.css";

const root = createRoot(document.getElementById("root")!);

const renderApp = () => {
  try {
    root.render(<App />);
  } catch (e) {
    console.error("Failed to render app:", e);
    const container = document.getElementById("root")!;
    const wrapper = document.createElement("div");
    wrapper.style.cssText = "padding:20px;font-family:sans-serif";
    const heading = document.createElement("h3");
    heading.textContent = "Failed to load";
    const message = document.createElement("p");
    message.textContent = String(e);
    wrapper.appendChild(heading);
    wrapper.appendChild(message);
    container.appendChild(wrapper);
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
