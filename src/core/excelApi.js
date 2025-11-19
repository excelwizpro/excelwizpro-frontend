// core/excelApi.js
import { delay } from "./utils.js";
import { emit } from "./eventBus.js";
import { getOfficeDiagnostics } from "./diagnostics.js";

export function officeReady() {
  return new Promise((resolve) => {
    if (window.Office && Office.onReady) {
      Office.onReady(resolve);
    } else {
      let tries = 0;
      const timer = setInterval(() => {
        tries++;
        if (window.Office && Office.onReady) {
          clearInterval(timer);
          Office.onReady(resolve);
        }
        if (tries > 40) {
          clearInterval(timer);
          resolve({ host: "unknown" });
        }
      }, 500);
    }
  });
}

export async function ensureExcelHost(info) {
  if (!info || info.host !== Office.HostType.Excel) {
    console.warn("⚠️ Not Excel host:", info && info.host);
    emit("ui:toast", { message: "⚠️ Excel host not detected.", kind: "warn" });
    return false;
  }
  console.log("🟢 Excel host OK");
  return true;
}

export async function waitForExcelApi(maxAttempts = 20) {
  for (let i = 1; i <= maxAttempts; i++) {
    try {
      await Excel.run(async (ctx) => {
        ctx.workbook.properties.load("title");
        await ctx.sync();
      });
      return true;
    } catch {
      await delay(350 + i * 120);
    }
  }
  emit("ui:toast", {
    message: "⚠️ Excel not ready — try reopening the add-in.",
    kind: "warn"
  });
  return false;
}

export async function safeExcelRun(cb) {
  try {
    return await Excel.run(cb);
  } catch (err) {
    console.warn("Excel.run failed:", err);
    emit("ui:toast", { message: "⚠️ Excel not ready", kind: "error" });
    throw err;
  }
}

export async function refreshSheetDropdown(el) {
  return safeExcelRun(async (ctx) => {
    const sheets = ctx.workbook.worksheets;
    sheets.load("items/name");
    await ctx.sync();

    el.innerHTML = "";
    sheets.items.forEach((s) => {
      const opt = document.createElement("option");
      opt.value = s.name;
      opt.textContent = s.name;
      el.appendChild(opt);
    });
  });
}

// Workbook structure change awareness (best-effort)
export async function attachWorkbookChangeListeners() {
  try {
    await Excel.run(async (ctx) => {
      const sheets = ctx.workbook.worksheets;
      sheets.load("items/name");
      await ctx.sync();

      sheets.items.forEach((sheet) => {
        try {
          const onChanged = sheet.onChanged;
          if (onChanged && onChanged.add) {
            onChanged.add(async () => {
              emit("workbook:changed");
            });
          }
        } catch (err) {
          console.warn("Sheet change listener unsupported:", err);
        }
      });
    });
  } catch (err) {
    console.warn("Workbook change listeners failed:", err);
  }
}

export function getExcelVersion() {
  const diag = getOfficeDiagnostics();
  return diag.version || "unknown";
}
