// ui/mainUI.js
import { getEl } from "../core/utils.js";
import { generateFormulaFromBackend, warmUpBackend } from "../core/backendClient.js";
import {
  refreshSheetDropdown,
  getExcelVersion
} from "../core/excelApi.js";
import { autoRefreshColumnMap, getCurrentColumnMap } from "../core/columnMapper.js";
import { showToast } from "./toast.js";
import { showError, clearError } from "./errorPanel.js";
import { on } from "../core/eventBus.js";

let lastFormula = "";
let isGenerating = false;

export async function initUI() {
  const sheetSelect = getEl("sheetSelect");
  const queryInput = getEl("query");
  const output = getEl("output");
  const genBtn = getEl("generateBtn");
  const clearBtn = getEl("clearBtn");
  const insertBtn = document.getElementById("insertBtn"); // recommended separate button

  await refreshSheetDropdown(sheetSelect);
  await autoRefreshColumnMap(true);
  warmUpBackend();

  // Toast bridge
  on("ui:toast", ({ message, kind }) => showToast(message, kind));

  // Column map updated hook (optional)
  on("columnMap:updated", () => {
    // Could update a status indicator if desired
  });

  genBtn.addEventListener("click", async () => {
    if (isGenerating) return;
    const query = queryInput.value.trim();
    if (!query) return showToast("⚠️ Enter a request", "warn");
    if (!navigator.onLine) return showToast("📴 Offline", "warn");

    isGenerating = true;
    genBtn.disabled = true;
    clearError();
    output.textContent = "⏳ Generating…";

    try {
      await autoRefreshColumnMap(false);
      if (!getCurrentColumnMap()) {
        await autoRefreshColumnMap(true);
      }

      const version = getExcelVersion();
      const payload = {
        query,
        columnMap: getCurrentColumnMap(),
        excelVersion: version,
        mainSheet: sheetSelect.value
      };

      const formula = await generateFormulaFromBackend(payload);
      lastFormula = formula;
      output.textContent = formula;
      showToast("✅ Formula ready", "success");
    } catch (err) {
      console.error("Generation failed:", err);
      output.textContent = "❌ Error — see details";
      showError("Formula generation failed.", err?.message || String(err));
    } finally {
      isGenerating = false;
      genBtn.disabled = false;
    }
  });

  clearBtn.addEventListener("click", () => {
    output.textContent = "";
    queryInput.value = "";
    clearError();
  });

  if (insertBtn) {
    insertBtn.addEventListener("click", async () => {
      if (!lastFormula) {
        return showToast("⚠️ No formula to insert", "warn");
      }
      try {
        await Excel.run(async (ctx) => {
          const range = ctx.workbook.getSelectedRange();
          range.load("rowCount,columnCount");
          await ctx.sync();

          if (range.rowCount !== 1 || range.columnCount !== 1) {
            const err = new Error("MULTI_CELL_SELECTION");
            err.code = "MULTI_CELL_SELECTION";
            throw err;
          }

          range.formulas = [[lastFormula]];
          await ctx.sync();
        });
        showToast("✅ Inserted", "success");
      } catch (err) {
        console.warn("Insert failed:", err);
        if (err && err.code === "MULTI_CELL_SELECTION") {
          showToast("⚠️ Select a single cell first", "warn");
        } else {
          showToast("⚠️ Select a cell first", "warn");
        }
      }
    });
  } else {
    console.warn("No #insertBtn found — insert will not be available.");
  }

  window.addEventListener("online", () => {
    if (lastFormula) showToast("🌐 Back online — formula restored", "info");
  });

  console.log("🟢 ExcelWizPro UI ready");
}
