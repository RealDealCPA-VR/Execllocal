import * as React from "react";
import { createRoot, Root } from "react-dom/client";
import App from "./components/App";
import "./taskpane.css";

/* global document, Office, window, module, require, HTMLElement */

const rootElement: HTMLElement | null = document.getElementById("container");
const root: Root | undefined = rootElement ? createRoot(rootElement) : undefined;

/**
 * A task pane that fails to boot renders as a blank white rectangle with no
 * way to tell why. Put the reason on screen instead.
 */
function showFatal(reason: string): void {
  if (!rootElement) {
    return;
  }
  root?.unmount();
  rootElement.innerHTML = "";
  const box = document.createElement("div");
  box.setAttribute("role", "alert");
  box.style.cssText = "padding:16px;font:13px 'Segoe UI',sans-serif;color:#a4262c;line-height:1.5";
  const title = document.createElement("strong");
  title.textContent = "ExcelLocal could not start";
  const detail = document.createElement("div");
  detail.style.cssText = "margin-top:8px;color:#424242;white-space:pre-wrap;word-break:break-word";
  detail.textContent = reason;
  const hint = document.createElement("div");
  hint.style.cssText = "margin-top:8px;color:#616161";
  hint.textContent = "Check that the dev server is running, then close and reopen the pane.";
  box.append(title, detail, hint);
  rootElement.appendChild(box);
}

if (typeof Office === "undefined") {
  // office.js is loaded from the Office CDN; without it no host API exists.
  showFatal("Office.js did not load. The pane needs a connection to appsforoffice.microsoft.com on first load.");
} else {
  Office.onReady((info) => {
    if (info.host && info.host !== Office.HostType.Excel) {
      showFatal("This add-in only runs in Excel (detected host: " + info.host + ").");
      return;
    }
    try {
      root?.render(<App />);
    } catch (e) {
      showFatal(String((e as Error)?.message ?? e));
    }
  }).catch((e: unknown) => showFatal("Office.onReady failed: " + String((e as Error)?.message ?? e)));
}

if ((module as any).hot) {
  (module as any).hot.accept("./components/App", () => {
    const NextApp = require("./components/App").default;
    root?.render(<NextApp />);
  });
}
