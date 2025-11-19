// core/network.js
import { DEFAULT_REQUEST_TIMEOUT, MAX_BACKEND_RETRIES } from "./config.js";
import { delay } from "./utils.js";

function timeoutSignal(ms) {
  if (typeof AbortController === "undefined") return undefined;
  const ctrl = new AbortController();
  const id = setTimeout(() => ctrl.abort(), ms);
  ctrl.signal.addEventListener("abort", () => clearTimeout(id));
  return ctrl.signal;
}

export async function safeFetch(url, { timeout = DEFAULT_REQUEST_TIMEOUT, ...opts } = {}) {
  if (!navigator.onLine) {
    const err = new Error("offline");
    err.code = "OFFLINE";
    throw err;
  }
  const signal = opts.signal || timeoutSignal(timeout);
  return fetch(url, { ...opts, signal });
}

export async function fetchWithRetry(url, options = {}) {
  let lastErr = null;
  for (let attempt = 1; attempt <= MAX_BACKEND_RETRIES; attempt++) {
    try {
      const res = await safeFetch(url, options);
      if (!res.ok) {
        lastErr = new Error(`HTTP ${res.status}`);
        lastErr.status = res.status;
        continue;
      }
      return res;
    } catch (err) {
      lastErr = err;
      if (err.name === "AbortError" || err.code === "OFFLINE") {
        // retry for timeouts/offline (in case network comes back)
      }
    }
    await delay(300 * attempt);
  }
  throw lastErr || new Error("fetchWithRetry failed");
}
