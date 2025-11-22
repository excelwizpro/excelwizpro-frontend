// frontend/src/UI/mainUI.js — v18 (uses backendClient)
/* global Excel, Office */

import {
  autoRefreshColumnMap,
  getCurrentColumnMap,
} from "../core/columnMapper.js";

import {
  getExcelVersion,
  safeExcelRun,
  refreshSheetDropdown,
} from "../core/excelApi.js";

import { generateFormulaFromBackend } from "../core/backendClient.js";
import { showToast } from "./toast.js";
import { on } from "../core/eventBus.js";

let lastFormula = "";

/**
 * Additional web-side clean up in case anything weird slipped through.
 * (This is intentionally conservative; backend already does most of the work.)
 */
function sanitizeForDisplay(str = "") {
  return String(str)
    // Drop zero-width, NBSP, etc.
    .replace(/[\u200B\u200C\u200D\u200E\u200F\uFEFF]/g, "")
    .replace(/\u00A0/g, " ")
    // Normalise whitespace
    .replace(/[\r\n\t]+/g, " ")
    .replace(/ {2,}/g, " ")
    .trim();
}

export async function initUI() {
  const sheetSelect = document.getElementById("sheetSelect");
  const queryBox = document.getElementById("query");
  const outputBox = document.getElementById("output");
  const generateBtn = document.getElementById("generateBtn");
  const clearBtn = document.getElementById("clearBtn");
  const insertBtn = document.getElementById("insertBtn");

  if (
    !sheetSelect ||
    !queryBox ||
    !outputBox ||
    !generateBtn ||
    !clearBtn ||
    !insertBtn
  ) {
    console.error("❌ Missing one or more UI elements in taskpane.html");
    return;
  }

  // Pipe EventBus toast events into the UI toast renderer
  on("ui:toast", ({ message, kind }) => showToast(message, kind || "info"));

  // Populate the sheet dropdown once at startup…
  await refreshSheetDropdown(sheetSelect);

  // …and prime the column map cache
  await autoRefreshColumnMap(true);

  // ------------------
  // GENERATE FORMULA
  // ------------------
  generateBtn.addEventListener("click", async () => {
    const prompt = queryBox.value.trim();
    if (!prompt) {
      showToast("Enter what you want the formula to do.", "warn");
      return;
    }

    generateBtn.disabled = true;
    insertBtn.disabled = true;
    clearBtn.disabled = true;

    outputBox.textContent = "🔍 Analysing your workbook and building a formula…";

    try {
      // Make sure we have a valid, recent Smart Column Map
      await autoRefreshColumnMap(false);
      let columnMap = getCurrentColumnMap();

      if (!columnMap) {
        // Force a rebuild if cache was empty or expired
        await autoRefreshColumnMap(true);
        columnMap = getCurrentColumnMap();
      }

      if (!columnMap) {
        showToast("Could not read workbook structure.", "error");
        outputBox.textContent =
          "ERROR: Smart Column Map not available. Try re-opening the add-in.";
        return;
      }

      const excelVersion = getExcelVersion();
      const mainSheet = sheetSelect.value || null;

      const formula = await generateFormulaFromBackend({
        query: prompt,
        columnMap,
        excelVersion,
        mainSheet,
      });

      lastFormula = sanitizeForDisplay(formula);
      outputBox.textContent = lastFormula;

      console.log("🧮 Engine formula:", lastFormula);
      showToast("Formula ready.", "success");
    } catch (err) {
      console.error("❌ Generation failed:", err);
      outputBox.textContent = `ERROR: ${
        err && err.message ? err.message : "Formula generation failed."
      }`;
      showToast("Generation failed.", "error");
    } finally {
      generateBtn.disabled = false;
      clearBtn.disabled = false;
      insertBtn.disabled = !lastFormula;
    }
  });

  // -----------
  // CLEAR UI
  // -----------
  clearBtn.addEventListener("click", () => {
    lastFormula = "";
    outputBox.textContent = "";
    queryBox.value = "";
    insertBtn.disabled = true;
  });

  // -----------
  // INSERT INTO SHEET
  // -----------
  insertBtn.addEventListener("click", async () => {
    if (!lastFormula) {
      showToast("Generate a formula first.", "warn");
      return;
    }

    try {
      await safeExcelRun(async (ctx) => {
        const sheet = ctx.workbook.worksheets.getActiveWorksheet();
        const range = sheet.getSelectedRange();

        range.load("rowCount,columnCount,address");
        await ctx.sync();

        if (range.rowCount !== 1 || range.columnCount !== 1) {
          const err = new Error("Please select exactly one cell.");
          err.code = "MULTI";
          throw err;
        }

        range.values = [[lastFormula]];
        await ctx.sync();
      });

      showToast("Formula inserted.", "success");
    } catch (err) {
      console.error("❌ Insert failed:", err);
      if (err.code === "MULTI") {
        showToast("Select exactly one cell before inserting.", "warn");
      } else {
        showToast("Could not insert formula into the sheet.", "error");
      }
    }
  });

  // Start with Insert disabled until we have a formula
  insertBtn.disabled = !lastFormula;
}
