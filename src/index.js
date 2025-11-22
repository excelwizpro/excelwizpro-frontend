// index.js
/* global Office, Excel */
import "../taskpane.css";
import { EXWZ_VERSION } from "./core/config.js";
import {
  officeReady,
  ensureExcelHost,
  waitForExcelApi,
  attachWorkbookChangeListeners
} from "./core/excelApi.js";
import { initUI } from "./UI/mainUI.js";

console.log(`🧠 ExcelWizPro Taskpane v${EXWZ_VERSION} starting…`);

if (typeof Office !== "undefined" && Office && Office.config) {
  Office.config = { extendedErrorLogging: true };
}

// -------------------------------------------------------------
// FIX: Ensure Office + DOM are both ready before attaching UI
// -------------------------------------------------------------
async function fullyReady() {
  // Wait for Office runtime
  const info = await officeReady();
  if (!(await ensureExcelHost(info))) return false;
  if (!(await waitForExcelApi())) return false;

  // Wait for DOM
  await new Promise((resolve) => {
    if (document.readyState === "complete" || document.readyState === "interactive") {
      resolve();
    } else {
      document.addEventListener("DOMContentLoaded", resolve);
    }
  });

  return true;
}

(async function boot() {
  console.log("🧠 Boot sequence…");

  if (!(await fullyReady())) {
    console.warn("Office not ready.");
    return;
  }

  // Attach workbook listeners AFTER Excel is valid
  await attachWorkbookChangeListeners();

  // Attach UI handlers AFTER DOM is fully loaded
  await initUI();

  const { showToast } = await import("./UI/toast.js");
  showToast("✅ ExcelWizPro ready!", "success");
})();
