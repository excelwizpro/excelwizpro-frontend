// src/UI/mainUI.js — v17.0 (Safe Insert + Debug)
/* global Excel, Office */

import {
  autoRefreshColumnMap,
  getCurrentColumnMap
} from "../core/columnMapper.js";

import { getExcelVersion, safeExcelRun } from "../core/excelApi.js";
import { DEFAULT_API_BASE } from "../core/config.js";
import { showToast } from "./toast.js";
import { emit, on } from "../core/eventBus.js";

let lastFormula = "";

/** Excel Web sanitiser */
function sanitizeForExcel(str = "") {
  let f = String(str)

    // Remove backslash escapes
    .replace(/\\+"/g, '"')
    .replace(/\\+'/g, "'")
    .replace(/\\/g, "")

    // Zero-width junk
    .replace(/[\u200B\u200C\u200D\u200E\u200F\uFEFF\u00A0]/g, "")

    // Newlines / tabs
    .replace(/[\r\n\t]+/g, " ")

    // Smart quotes
    .replace(/[“”]/g, '"')
    .replace(/[‘’]/g, "'")

    // Unicode math
    .replace(/×/g, "*").replace(/÷/g, "/")
    .replace(/[–—−]/g, "-")

    .replace(/ {2,}/g, " ")
    .trim();

  if (!f.startsWith("=")) f = "=" + f;
  return f;
}

function resolveApiBase() {
  try {
    const params = new URLSearchParams(window.location.search);
    if (params.get("apiBase")) return params.get("apiBase");

    const saved = Office?.context?.roamingSettings?.get("excelwizpro_api_base");
    if (saved) return saved;
  } catch {}

  return DEFAULT_API_BASE;
}

export async function initUI() {
  const sheetSelect = document.getElementById("sheetSelect");
  const queryBox = document.getElementById("query");
  const outputBox = document.getElementById("output");
  const generateBtn = document.getElementById("generateBtn");
  const clearBtn = document.getElementById("clearBtn");
  const insertBtn = document.getElementById("insertBtn");

  // Load sheet list
  await safeExcelRun(async (ctx) => {
    const sheets = ctx.workbook.worksheets;
    sheets.load("items/name");
    await ctx.sync();

    sheetSelect.innerHTML = "";
    sheets.items.forEach((s) => {
      const opt = document.createElement("option");
      opt.value = s.name;
      opt.textContent = s.name;
      sheetSelect.appendChild(opt);
    });
  });

  await autoRefreshColumnMap(true);
  on("ui:toast", ({ message, kind }) => showToast(message, kind));

  // ------------------
  // GENERATE
  // ------------------
  generateBtn.addEventListener("click", async () => {
    const prompt = queryBox.value.trim();
    if (!prompt) return showToast("Enter a request.", "warn");

    generateBtn.disabled = true;
    outputBox.textContent = "Working…";

    try {
      await autoRefreshColumnMap(false);
      let columnMap = getCurrentColumnMap();
      if (!columnMap) {
        await autoRefreshColumnMap(true);
        columnMap = getCurrentColumnMap();
      }

      const res = await fetch(`${resolveApiBase()}/generate`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          query: prompt,
          columnMap,
          excelVersion: getExcelVersion(),
          mainSheet: sheetSelect.value
        })
      });

      const data = await res.json();
      lastFormula = data.formula || "=ERROR(\"No formula\")";

      console.log("RAW backend EXACT:", lastFormula, [...lastFormula]);

      outputBox.textContent = lastFormula;
      showToast("Ready.", "success");

    } catch (err) {
      console.error(err);
      showToast("Generation failed.", "error");
    } finally {
      generateBtn.disabled = false;
    }
  });

  // CLEAR
  clearBtn.addEventListener("click", () => {
    lastFormula = "";
    outputBox.textContent = "";
    queryBox.value = "";
  });

  // ------------------
  // INSERT → Excel
  // ------------------
  insertBtn.addEventListener("click", async () => {
    if (!lastFormula) return showToast("Nothing to insert.", "warn");

    const cleaned = sanitizeForExcel(lastFormula);
    console.log("Cleaned before insert:", cleaned);

    try {
      await Excel.run(async (ctx) => {
        const rng = ctx.workbook.getSelectedRange();
        rng.load("rowCount,columnCount");
        await ctx.sync();

        if (rng.rowCount !== 1 || rng.columnCount !== 1) {
          const e = new Error("MULTI");
          e.code = "MULTI";
          throw e;
        }

        rng.formulas = [[cleaned]];
        await ctx.sync();
      });

      showToast("Inserted.", "success");

    } catch (err) {
      console.error(err);
      if (err.code === "MULTI")
        showToast("Select one cell.", "warn");
      else
        showToast("Could not insert.", "error");
    }
  });
}
