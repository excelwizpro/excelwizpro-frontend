// index.js
/* global Office, Excel */
import "../taskpane.css";
import { EXWZ_VERSION } from "./core/config.js";
import { officeReady, ensureExcelHost, waitForExcelApi, attachWorkbookChangeListeners } from "./core/excelApi.js";
import { initUI } from "./UI/mainUI.js";

console.log(`🧠 ExcelWizPro Taskpane v${EXWZ_VERSION} starting…`);

if (typeof Office !== "undefined" && Office && Office.config) {
  Office.config = { extendedErrorLogging: true };
}

(async function boot() {
  console.log("🧠 Boot sequence…");
  const info = await officeReady();
  if (!(await ensureExcelHost(info))) return;
  if (!(await waitForExcelApi())) return;

  await attachWorkbookChangeListeners();
  await initUI();

  const { showToast } = await import("./UI/toast.js");
  showToast("✅ ExcelWizPro ready!", "success");
})();
