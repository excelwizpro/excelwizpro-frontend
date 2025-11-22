// src/core/columnMapper.js
import {
  COLUMN_MAP_TTL_MS,
  MAX_DATA_ROWS_PER_COLUMN
} from "./config.js";
import { columnIndexToLetter, normalizeName } from "./utils.js";
import { safeExcelRun } from "./excelApi.js";
import { emit, on } from "./eventBus.js";

let columnMapCache = "";
let lastColumnMapBuild = 0;
let refreshInProgress = false;

/**
 * Build the Smart Column Map (V15.4)
 * ---------------------------------
 * Forced assumptions:
 *  - Exactly 1 header row
 *  - Data ALWAYS starts at row 2
 *  - This fixes all misalignment issues
 */
async function buildColumnMapInternal() {
  return safeExcelRun(async (ctx) => {
    const wb = ctx.workbook;
    const sheets = wb.worksheets;

    sheets.load("items/name,items/visibility");
    await ctx.sync();

    const lines = [];
    const globalNameCounts = Object.create(null);

    for (const sheet of sheets.items) {
      const vis = sheet.visibility;
      const visText = vis !== "Visible" ? ` (${vis.toLowerCase()})` : "";
      lines.push(`Sheet: ${sheet.name}${visText}`);

      const used = sheet.getUsedRangeOrNullObject();
      used.load("rowCount,columnCount,rowIndex,columnIndex,isNullObject");
      await ctx.sync();

      if (used.isNullObject || used.rowCount < 1) continue;

      // ---------------------------------------------------------
      // FORCE: Exactly 1 header row
      // ---------------------------------------------------------
      const headerRows = 1;

      const headerRange = sheet.getRangeByIndexes(
        used.rowIndex,         // top row
        used.columnIndex,
        1,                     // ALWAYS 1 header row
        used.columnCount
      );
      headerRange.load("values");

      const tables = sheet.tables;
      tables.load("items/name");

      const pivots = sheet.pivotTables;
      pivots.load("items/name");

      await ctx.sync();

      const headers = headerRange.values;

      // ---------------------------------------------------------
      // FORCE: Data ALWAYS starts row 2
      // ---------------------------------------------------------
      const startRow = 2; // 1-based index

      const lastRowCandidate = used.rowIndex + used.rowCount; // 1-based
      const maxAllowed = startRow + MAX_DATA_ROWS_PER_COLUMN;
      const lastRow = Math.min(lastRowCandidate, maxAllowed);

      // ---------------------------------------------------------
      // Column mapping (Sheet fields)
      // ---------------------------------------------------------
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

      // ---------------------------------------------------------
      // Tables
      // ---------------------------------------------------------
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

        const headerVals = header.values?.[0] || [];
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

      // ---------------------------------------------------------
      // Pivot markers
      // ---------------------------------------------------------
      pivots.items.forEach((p) => lines.push(`PivotSource: ${p.name}`));
    }

    // ---------------------------------------------------------
    // Named ranges
    // ---------------------------------------------------------
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

// ---------------------------------------------------------
// Auto-refresh caching
// ---------------------------------------------------------
export async function autoRefreshColumnMap(force = false) {
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

export function getCurrentColumnMap() {
  return columnMapCache;
}

// Invalidate cache on workbook change
on("workbook:changed", () => {
  columnMapCache = "";
  lastColumnMapBuild = 0;
  console.log("🧹 Column map cache invalidated due to workbook change");
});
