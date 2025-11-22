// frontend/src/core/backendClient.js
/* global Office */

import { DEFAULT_API_BASE } from "./config.js";
import { parseQueryParams } from "./utils.js";
import { fetchWithRetry } from "./network.js";
import { emit } from "./eventBus.js";

let apiBase = DEFAULT_API_BASE;

/**
 * Clean up formulas so Excel doesn't choke on weird Unicode etc.
 */
function sanitizeFormula(str = "") {
  return String(str)
    // Zero-width junk
    .replace(/[\u200B\u200C\u200D\u200E\u200F\uFEFF]/g, "")
    // Non-breaking space → normal space
    .replace(/\u00A0/g, " ")
    // Smart quotes → regular
    .replace(/[“”]/g, '"')
    .replace(/[‘’]/g, "'")
    // Unicode math → ASCII
    .replace(/×/g, "*")
    .replace(/÷/g, "/")
    .replace(/[–—−]/g, "-")
    // Newlines & tabs → single space
    .replace(/[\r\n\t]+/g, " ")
    // Collapse multiple spaces
    .replace(/ {2,}/g, " ")
    .trim();
}

/**
 * Decide which backend base URL to use:
 *  1. ?apiBase= query param (for dev / staging)
 *  2. Office roamingSettings (user override)
 *  3. DEFAULT_API_BASE from config
 */
export function resolveApiBase() {
  try {
    const params = parseQueryParams();
    if (params.apiBase) {
      apiBase = params.apiBase;
      return apiBase;
    }

    if (Office?.context?.roamingSettings) {
      const stored = Office.context.roamingSettings.get(
        "excelwizpro_api_base"
      );
      if (stored) {
        apiBase = stored;
        return apiBase;
      }
    }
  } catch {
    // fall through to default
  }

  return apiBase;
}

/**
 * Optional: allow UI / admin screen to override the API base,
 * and persist it into roaming settings.
 */
export function setApiBase(nextBase) {
  apiBase = nextBase || DEFAULT_API_BASE;

  try {
    if (Office?.context?.roamingSettings) {
      Office.context.roamingSettings.set("excelwizpro_api_base", apiBase);
      Office.context.roamingSettings.saveAsync();
    }
  } catch {
    // non-fatal
  }

  return apiBase;
}

/**
 * Main call into the AI formula engine.
 *
 * Payload shape matches backend/src/routes/generate.js:
 *   { query, columnMap, excelVersion, mainSheet }
 *
 * Returns the final, sanitized Excel formula string (always starting with "=").
 */
export async function generateFormulaFromBackend(payload) {
  const base = resolveApiBase();

  try {
    const res = await fetchWithRetry(`${base}/generate`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload),
      cache: "no-store",
    });

    if (!res.ok) {
      const err = new Error(`Backend returned ${res.status}`);
      err.status = res.status;
      throw err;
    }

    const data = await res.json();

    let formula = data.formula || '=ERROR("No formula returned")';

    // Final frontend sanitiser pass
    formula = sanitizeFormula(formula);

    // Guarantee a leading "=" so users can paste safely
    if (!formula.startsWith("=")) {
      formula = "=" + formula;
    }

    return formula;
  } catch (err) {
    console.error("❌ Backend /generate failed:", err);

    emit("ui:toast", {
      message: "⚠️ Could not reach the ExcelWizPro formula engine",
      kind: "error",
    });

    throw err;
  }
}
