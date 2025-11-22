// core/backendClient.js
import { DEFAULT_API_BASE } from "./config.js";
import { parseQueryParams } from "./utils.js";
import { fetchWithRetry } from "./network.js";
import { emit } from "./eventBus.js";

let apiBase = DEFAULT_API_BASE;

// Clean Excel formula (frontend)
function sanitizeFormula(str = "") {
  return String(str)
    .replace(/[\u200B\u200C\u200D\u200E\u200F\uFEFF]/g, "") // zero-width
    .replace(/\u00A0/g, " ")                                // NBSP
    .replace(/[“”]/g, '"')                                  // smart quotes
    .replace(/[‘’]/g, "'")
    .replace(/×/g, "*")
    .replace(/÷/g, "/")
    .replace(/[–—−]/g, "-")
    .replace(/\r?\n/g, " ")                                 // newlines → space
    .replace(/ {2,}/g, " ")                                 // collapse spaces
    .trim();
}

export function resolveApiBase() {
  try {
    const params = parseQueryParams();
    if (params.apiBase) {
      apiBase = params.apiBase;
      return apiBase;
    }

    if (Office?.context?.roamingSettings) {
      const stored = Office.context.roamingSettings.get("excelwizpro_api_base");
      if (stored) {
        apiBase = stored;
        return apiBase;
      }
    }
  } catch {}

  return apiBase;
}

export function setApiBase(value) {
  apiBase = value || DEFAULT_API_BASE;
  try {
    if (Office?.context?.roamingSettings) {
      Office.context.roamingSettings.set("excelwizpro_api_base", apiBase);
      Office.context.roamingSettings.saveAsync();
    }
  } catch {}
}

// Generate Formula — patched
export async function generateFormulaFromBackend(payload) {
  const base = resolveApiBase();

  try {
    const res = await fetchWithRetry(`${base}/generate`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload),
      cache: "no-store"
    });

    const data = await res.json();

    let formula = data.formula || '=ERROR("No formula returned")';

    // 🔥 Apply frontend sanitiser here too
    formula = sanitizeFormula(formula);

    // Ensure it always starts with "="
    if (!formula.startsWith("=")) {
      formula = "=" + formula;
    }

    return formula;

  } catch (err) {
    console.error("❌ Backend /generate failed:", err);

    emit("ui:toast", {
      message: "⚠️ Could not reach formula engine",
      kind: "error"
    });

    throw err;
  }
}
