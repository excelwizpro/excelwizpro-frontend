// core/backendClient.js
import { DEFAULT_API_BASE } from "./config.js";
import { parseQueryParams } from "./utils.js";
import { fetchWithRetry } from "./network.js";
import { emit } from "./eventBus.js";

let apiBase = DEFAULT_API_BASE;

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
  } catch (err) {
    console.warn("API base resolution error:", err);
  }
  return apiBase;
}

export function setApiBase(value) {
  apiBase = value || DEFAULT_API_BASE;
  try {
    if (Office?.context?.roamingSettings) {
      Office.context.roamingSettings.set("excelwizpro_api_base", apiBase);
      Office.context.roamingSettings.saveAsync();
    }
  } catch (err) {
    console.warn("Failed to persist API base:", err);
  }
}

export async function warmUpBackend(maxAttempts = 5) {
  const container = document.querySelector("main.container");
  const status = document.createElement("div");
  Object.assign(status.style, {
    padding: "6px",
    marginBottom: "8px",
    borderRadius: "6px",
    fontSize: "0.9rem",
    textAlign: "center"
  });
  if (container) container.prepend(status);

  const base = resolveApiBase();

  for (let i = 1; i <= maxAttempts; i++) {
    try {
      const r = await fetchWithRetry(`${base}/health`, { timeout: 4000 });
      if (r.ok) {
        status.textContent = "✅ Backend ready";
        status.style.background = "#e6ffed";
        status.style.color = "#0c7a0c";
        setTimeout(() => status.remove(), 2000);
        emit("backend:ready");
        return;
      }
    } catch {
      // ignore; will retry
    }

    status.textContent = `⏳ Waking backend… (${i}/${maxAttempts})`;
    status.style.background = "#fff4ce";
    status.style.color = "#976f00";
  }

  status.textContent = "❌ Backend unreachable";
  status.style.background = "#fde7e9";
  status.style.color = "#c22";
  emit("backend:error");
}

export async function generateFormulaFromBackend(payload) {
  const base = resolveApiBase();
  const r = await fetchWithRetry(`${base}/generate`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(payload),
    cache: "no-store"
  });
  const data = await r.json();
  return data.formula || '=ERROR("No formula returned")';
}
