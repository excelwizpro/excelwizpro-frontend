(function polyfill() {
  const relList = document.createElement("link").relList;
  if (relList && relList.supports && relList.supports("modulepreload")) {
    return;
  }
  for (const link of document.querySelectorAll('link[rel="modulepreload"]')) {
    processPreload(link);
  }
  new MutationObserver((mutations) => {
    for (const mutation of mutations) {
      if (mutation.type !== "childList") {
        continue;
      }
      for (const node of mutation.addedNodes) {
        if (node.tagName === "LINK" && node.rel === "modulepreload")
          processPreload(node);
      }
    }
  }).observe(document, { childList: true, subtree: true });
  function getFetchOpts(link) {
    const fetchOpts = {};
    if (link.integrity) fetchOpts.integrity = link.integrity;
    if (link.referrerPolicy) fetchOpts.referrerPolicy = link.referrerPolicy;
    if (link.crossOrigin === "use-credentials")
      fetchOpts.credentials = "include";
    else if (link.crossOrigin === "anonymous") fetchOpts.credentials = "omit";
    else fetchOpts.credentials = "same-origin";
    return fetchOpts;
  }
  function processPreload(link) {
    if (link.ep)
      return;
    link.ep = true;
    const fetchOpts = getFetchOpts(link);
    fetch(link.href, fetchOpts);
  }
})();
const scriptRel = "modulepreload";
const assetsURL = function(dep) {
  return "/excelwizpro-frontend/" + dep;
};
const seen = {};
const __vitePreload = function preload(baseModule, deps, importerUrl) {
  let promise = Promise.resolve();
  if (deps && deps.length > 0) {
    document.getElementsByTagName("link");
    const cspNonceMeta = document.querySelector(
      "meta[property=csp-nonce]"
    );
    const cspNonce = (cspNonceMeta == null ? void 0 : cspNonceMeta.nonce) || (cspNonceMeta == null ? void 0 : cspNonceMeta.getAttribute("nonce"));
    promise = Promise.allSettled(
      deps.map((dep) => {
        dep = assetsURL(dep);
        if (dep in seen) return;
        seen[dep] = true;
        const isCss = dep.endsWith(".css");
        const cssSelector = isCss ? '[rel="stylesheet"]' : "";
        if (document.querySelector(`link[href="${dep}"]${cssSelector}`)) {
          return;
        }
        const link = document.createElement("link");
        link.rel = isCss ? "stylesheet" : scriptRel;
        if (!isCss) {
          link.as = "script";
        }
        link.crossOrigin = "";
        link.href = dep;
        if (cspNonce) {
          link.setAttribute("nonce", cspNonce);
        }
        document.head.appendChild(link);
        if (isCss) {
          return new Promise((res, rej) => {
            link.addEventListener("load", res);
            link.addEventListener(
              "error",
              () => rej(new Error(`Unable to preload CSS for ${dep}`))
            );
          });
        }
      })
    );
  }
  function handlePreloadError(err) {
    const e = new Event("vite:preloadError", {
      cancelable: true
    });
    e.payload = err;
    window.dispatchEvent(e);
    if (!e.defaultPrevented) {
      throw err;
    }
  }
  return promise.then((res) => {
    for (const item of res || []) {
      if (item.status !== "rejected") continue;
      handlePreloadError(item.reason);
    }
    return baseModule().catch(handlePreloadError);
  });
};
const EXWZ_VERSION = "13.0.0";
const DEFAULT_API_BASE = "https://excelwizpro-backend.onrender.com";
const COLUMN_MAP_TTL_MS = 90 * 1e3;
const MAX_DATA_ROWS_PER_COLUMN = 5e4;
function delay(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}
function columnIndexToLetter(index) {
  let n = index + 1;
  let letters = "";
  while (n > 0) {
    const rem = (n - 1) % 26;
    letters = String.fromCharCode(65 + rem) + letters;
    n = Math.floor((n - 1) / 26);
  }
  return letters;
}
function normalizeName(name) {
  return String(name || "").trim().toLowerCase().replace(/\s+/g, "_");
}
const listeners = /* @__PURE__ */ new Map();
function on(event, handler) {
  if (!listeners.has(event)) listeners.set(event, /* @__PURE__ */ new Set());
  listeners.get(event).add(handler);
  return () => off(event, handler);
}
function off(event, handler) {
  const set = listeners.get(event);
  if (!set) return;
  set.delete(handler);
}
function emit(event, payload) {
  const set = listeners.get(event);
  if (!set) return;
  for (const handler of set) {
    try {
      handler(payload);
    } catch (err) {
      console.warn(`EventBus handler for '${event}' failed:`, err);
    }
  }
}
function getOfficeDiagnostics() {
  var _a, _b, _c, _d, _e, _f, _g;
  try {
    return {
      host: ((_a = Office.context) == null ? void 0 : _a.host) || "unknown",
      platform: ((_c = (_b = Office.context) == null ? void 0 : _b.diagnostics) == null ? void 0 : _c.platform) || "unknown",
      version: ((_e = (_d = Office.context) == null ? void 0 : _d.diagnostics) == null ? void 0 : _e.version) || "unknown",
      build: ((_g = (_f = Office.context) == null ? void 0 : _f.diagnostics) == null ? void 0 : _g.build) || "n/a"
    };
  } catch {
    return { host: "unknown", platform: "unknown", version: "unknown" };
  }
}
function officeReady() {
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
async function ensureExcelHost(info) {
  if (!info || info.host !== Office.HostType.Excel) {
    console.warn("⚠️ Not Excel host:", info && info.host);
    emit("ui:toast", { message: "⚠️ Excel host not detected.", kind: "warn" });
    return false;
  }
  console.log("🟢 Excel host OK");
  return true;
}
async function waitForExcelApi(maxAttempts = 20) {
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
async function safeExcelRun(cb) {
  try {
    return await Excel.run(cb);
  } catch (err) {
    console.warn("Excel.run failed:", err);
    emit("ui:toast", { message: "⚠️ Excel not ready", kind: "error" });
    throw err;
  }
}
async function attachWorkbookChangeListeners() {
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
function getExcelVersion() {
  const diag = getOfficeDiagnostics();
  return diag.version || "unknown";
}
let columnMapCache = "";
let lastColumnMapBuild = 0;
let refreshInProgress = false;
async function buildColumnMapInternal() {
  return safeExcelRun(async (ctx) => {
    var _a;
    const wb = ctx.workbook;
    const sheets = wb.worksheets;
    sheets.load("items/name,items/visibility");
    await ctx.sync();
    const lines = [];
    const globalNameCounts = /* @__PURE__ */ Object.create(null);
    for (const sheet of sheets.items) {
      const vis = sheet.visibility;
      const visText = vis !== "Visible" ? ` (${vis.toLowerCase()})` : "";
      lines.push(`Sheet: ${sheet.name}${visText}`);
      const used = sheet.getUsedRangeOrNullObject();
      used.load("rowCount,columnCount,rowIndex,columnIndex,isNullObject");
      await ctx.sync();
      if (used.isNullObject || used.rowCount < 1) continue;
      const headerRange = sheet.getRangeByIndexes(
        used.rowIndex,
        // top row
        used.columnIndex,
        1,
        // ALWAYS 1 header row
        used.columnCount
      );
      headerRange.load("values");
      const tables = sheet.tables;
      tables.load("items/name");
      const pivots = sheet.pivotTables;
      pivots.load("items/name");
      await ctx.sync();
      const headers = headerRange.values;
      const startRow = 2;
      const lastRowCandidate = used.rowIndex + used.rowCount;
      const maxAllowed = startRow + MAX_DATA_ROWS_PER_COLUMN;
      const lastRow = Math.min(lastRowCandidate, maxAllowed);
      if (lastRow >= startRow) {
        for (let col = 0; col < used.columnCount; col++) {
          const primaryHeader = String(headers[0][col] ?? "").trim();
          if (!primaryHeader) continue;
          let normalized = normalizeName(primaryHeader);
          if (globalNameCounts[normalized]) {
            globalNameCounts[normalized] += 1;
            normalized = `${normalized}__${globalNameCounts[normalized]}`;
          } else {
            globalNameCounts[normalized] = 1;
          }
          const colLetter = columnIndexToLetter(used.columnIndex + col);
          const safeSheet = sheet.name.replace(/'/g, "''");
          lines.push(
            `${normalized} = '${safeSheet}'!${colLetter}${startRow}:${colLetter}${lastRow}`
          );
        }
      }
      const tableMeta = tables.items.map((table) => {
        return {
          table,
          header: table.getHeaderRowRange(),
          body: table.getDataBodyRange()
        };
      });
      tableMeta.forEach((m) => {
        m.header.load("values");
        m.body.load("address,rowCount,columnCount");
      });
      await ctx.sync();
      for (const { table, header } of tableMeta) {
        lines.push(`Table: ${table.name}`);
        const headerVals = ((_a = header.values) == null ? void 0 : _a[0]) || [];
        headerVals.forEach((h) => {
          if (!h) return;
          let norm = normalizeName(`${table.name}.${h}`);
          if (globalNameCounts[norm]) {
            globalNameCounts[norm] += 1;
            norm = `${norm}__${globalNameCounts[norm]}`;
          } else {
            globalNameCounts[norm] = 1;
          }
          const structured = `${table.name}[${h}]`;
          lines.push(`${norm} = ${structured}`);
        });
      }
      pivots.items.forEach((p) => lines.push(`PivotSource: ${p.name}`));
    }
    const names = wb.names;
    names.load("items/name");
    await ctx.sync();
    const namedMeta = [];
    for (const n of names.items) {
      const r = n.getRange();
      r.load("address");
      namedMeta.push({ name: n.name, range: r });
    }
    await ctx.sync();
    namedMeta.forEach(({ name, range }) => {
      lines.push(`NamedRange: ${name}`);
      let norm = normalizeName(name);
      if (globalNameCounts[norm]) {
        globalNameCounts[norm] += 1;
        norm = `${norm}__${globalNameCounts[norm]}`;
      } else {
        globalNameCounts[norm] = 1;
      }
      lines.push(`${norm} = ${range.address}`);
    });
    const preview = lines.join("\n").slice(0, 600);
    console.log("🔍 Column map preview (v15.4 forced row2):\n", preview);
    return lines.join("\n");
  });
}
async function autoRefreshColumnMap(force = false) {
  if (refreshInProgress) return;
  try {
    const now = Date.now();
    if (!force && columnMapCache && now - lastColumnMapBuild < COLUMN_MAP_TTL_MS) {
      console.log("🔄 Using cached Smart Column Map (recent)");
      return;
    }
    refreshInProgress = true;
    console.log("🔄 Refreshing Smart Column Map (v15.4)...");
    columnMapCache = await buildColumnMapInternal();
    lastColumnMapBuild = Date.now();
    console.log("✅ Updated Smart Column Map (v15.4)");
    emit("columnMap:updated", { columnMap: columnMapCache });
  } catch (err) {
    console.warn("Auto-refresh failed:", err);
    emit("ui:toast", { message: "⚠️ Could not refresh column map", kind: "error" });
  } finally {
    refreshInProgress = false;
  }
}
function getCurrentColumnMap() {
  return columnMapCache;
}
on("workbook:changed", () => {
  columnMapCache = "";
  lastColumnMapBuild = 0;
  console.log("🧹 Column map cache invalidated due to workbook change");
});
let container = null;
function getContainer() {
  if (!container) {
    container = document.createElement("div");
    container.className = "exwz-toast-container";
    Object.assign(container.style, {
      position: "fixed",
      bottom: "14px",
      right: "14px",
      zIndex: 99999,
      display: "flex",
      flexDirection: "column",
      gap: "6px",
      maxWidth: "260px",
      pointerEvents: "none"
    });
    document.body.appendChild(container);
  }
  return container;
}
function showToast(msg, kind = "info") {
  const c = getContainer();
  const t = document.createElement("div");
  t.textContent = msg;
  t.className = "exwz-toast";
  const base = {
    padding: "8px 10px",
    borderRadius: "6px",
    fontSize: "0.85rem",
    fontFamily: "Inter, sans-serif",
    boxShadow: "0 2px 10px rgba(0,0,0,0.25)",
    pointerEvents: "auto"
  };
  const style = {
    info: { background: "#e5f1ff", color: "#084f94" },
    success: { background: "#e6ffed", color: "#0c7a0c" },
    warn: { background: "#fff4ce", color: "#976f00" },
    error: { background: "#fde7e9", color: "#c22" }
  };
  Object.assign(t.style, base, style[kind] || style.info);
  c.appendChild(t);
  setTimeout(() => t.remove(), 2400);
}
const toast = /* @__PURE__ */ Object.freeze(/* @__PURE__ */ Object.defineProperty({
  __proto__: null,
  showToast
}, Symbol.toStringTag, { value: "Module" }));
let lastFormula = "";
function resolveApiBase() {
  var _a;
  try {
    const params = new URLSearchParams(window.location.search);
    if (params.get("apiBase")) return params.get("apiBase");
    if ((_a = Office == null ? void 0 : Office.context) == null ? void 0 : _a.roamingSettings) {
      const saved = Office.context.roamingSettings.get("excelwizpro_api_base");
      if (saved) return saved;
    }
  } catch (err) {
    console.warn("API base resolution error:", err);
  }
  return DEFAULT_API_BASE;
}
async function initUI() {
  const sheetSelect = document.getElementById("sheetSelect");
  const queryBox = document.getElementById("query");
  const outputBox = document.getElementById("output");
  const generateBtn = document.getElementById("generateBtn");
  const clearBtn = document.getElementById("clearBtn");
  const insertBtn = document.getElementById("insertBtn");
  if (!sheetSelect || !queryBox || !outputBox) {
    console.error("❌ UI elements missing");
    return;
  }
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
  pingBackend(resolveApiBase());
  generateBtn.addEventListener("click", async () => {
    const prompt = queryBox.value.trim();
    if (!prompt) {
      showToast("⚠️ Enter a request", "warn");
      return;
    }
    if (!navigator.onLine) {
      showToast("📴 Offline", "warn");
      return;
    }
    generateBtn.disabled = true;
    outputBox.textContent = "⏳ Generating…";
    try {
      await autoRefreshColumnMap(false);
      let columnMap = getCurrentColumnMap();
      if (!columnMap) {
        await autoRefreshColumnMap(true);
        columnMap = getCurrentColumnMap();
      }
      const excelVersion = getExcelVersion();
      const payload = {
        query: prompt,
        columnMap,
        excelVersion,
        mainSheet: sheetSelect.value
      };
      const apiBase = resolveApiBase();
      const res = await fetch(`${apiBase}/generate`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(payload)
      });
      const data = await res.json();
      lastFormula = data.formula || '=ERROR("No formula returned")';
      outputBox.textContent = lastFormula;
      showToast("✅ Formula ready", "success");
    } catch (err) {
      console.error("Generation error:", err);
      outputBox.textContent = "❌ Error — check console";
      showToast("⚠️ Formula generation failed", "error");
    } finally {
      generateBtn.disabled = false;
    }
  });
  clearBtn.addEventListener("click", () => {
    outputBox.textContent = "";
    queryBox.value = "";
  });
  insertBtn.addEventListener("click", async () => {
    if (!lastFormula) {
      showToast("⚠️ Nothing to insert", "warn");
      return;
    }
    let formula = lastFormula.replace(/[\u200B\u200C\u200D\u200E\u200F\uFEFF\u00A0]/g, "").replace(/[\r\n]+/g, " ").replace(/\s+/g, " ").replace(/''/g, "'").trim();
    if (!formula.startsWith("=")) {
      formula = "=" + formula;
    }
    try {
      await Excel.run(async (ctx) => {
        const rng = ctx.workbook.getSelectedRange();
        rng.load("rowCount,columnCount");
        await ctx.sync();
        if (rng.rowCount !== 1 || rng.columnCount !== 1) {
          const e = new Error("MULTI_CELL");
          e.code = "MULTI_CELL";
          throw e;
        }
        rng.formulas = [[formula]];
        await ctx.sync();
      });
      showToast("✅ Inserted", "success");
    } catch (err) {
      console.error("Insert error:", err);
      if ((err == null ? void 0 : err.code) === "MULTI_CELL") {
        showToast("⚠️ Select a single cell", "warn");
      } else {
        showToast("⚠️ Could not insert formula", "error");
      }
    }
  });
}
async function pingBackend(apiBase) {
  try {
    const res = await fetch(`${apiBase}/health`, { method: "GET" });
    if (res.ok) emit("backend:ready");
    else emit("backend:error");
  } catch {
    emit("backend:error");
  }
}
console.log(`🧠 ExcelWizPro Taskpane v${EXWZ_VERSION} starting…`);
if (typeof Office !== "undefined" && Office && Office.config) {
  Office.config = { extendedErrorLogging: true };
}
async function fullyReady() {
  const info = await officeReady();
  if (!await ensureExcelHost(info)) return false;
  if (!await waitForExcelApi()) return false;
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
  if (!await fullyReady()) {
    console.warn("Office not ready.");
    return;
  }
  await attachWorkbookChangeListeners();
  await initUI();
  const { showToast: showToast2 } = await __vitePreload(async () => {
    const { showToast: showToast3 } = await Promise.resolve().then(() => toast);
    return { showToast: showToast3 };
  }, true ? void 0 : void 0);
  showToast2("✅ ExcelWizPro ready!", "success");
})();
//# sourceMappingURL=taskpane-DXF9yk1d.js.map
