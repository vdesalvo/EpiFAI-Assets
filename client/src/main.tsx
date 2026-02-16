import { createRoot } from "react-dom/client";
import App from "./App";
import "./index.css";

const root = createRoot(document.getElementById("root")!);

if (window.Office) {
  Office.onReady(() => {
    root.render(<App />);
  });
} else {
  root.render(<App />);
}
