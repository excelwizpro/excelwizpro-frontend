// core/columnMapper.js
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

      if (used.isNullObject || used.rowCount < 2) continue;

      const headerRows = Math.min(3, used.rowCount);
      const headerRange = sheet.getRangeByIndexes(
        used.rowIndex,
        used.columnIndex,
        headerRows,
        used.columnCount
      );
      headerRange.load("values");

      const tables = sheet.tables;
      tables.load("items/name");

      const pivots = sheet.pivotTables;
      pivots.load("items/name");

      await ctx.sync();

      const headers = headerRange.values;
      const dataStartRowIndex = used.rowIndex + headerRows;
      const dataLastRowIndex = used.rowIndex + used.rowCount - 1;
      const startRow = dataStartRowIndex + 1; // 1-based

      const maxLastRow = startRow + MAX_DATA_ROWS_PER_COLUMN - 1;
      const lastRowCandidate = dataLastRowIndex + 1; // 1-based
      const lastRow = Math.min(lastRowCandidate, maxLastRow);

      if (lastRow >= startRow) {
        for (let col = 0; col < used.columnCount; col++) {
          const headerTexts = [];
          for (let r = 0; r < headerRows; r++) {
            const v = headers[r][col];
            headerTexts[r] =
              v !== null && v !== "" && v !== undefined ? String(v).trim() : "";
          }

          let primary = "";
          for (let r = headerRows - 1; r >= 0; r--) {
            if (headerTexts[r]) {
              primary = headerTexts[r];
              break;
            }
          }
          if (!primary) continue;

          let combined = primary;
          for (let r = 0; r < headerRows - 1; r++) {
            if (headerTexts[r] && headerTexts[r] !== primary) {
              combined = `${headerTexts[r]} - ${combined}`;
              break;
            }
          }

          let normalized = normalizeName(combined);

          if (globalNameCounts[normalized]) {
            globalNameCounts[normalized] += 1;
            normalized = `${normalized}__${globalNameCounts[normalized]}`;
          } else {
            globalNameCounts[normalized] = 1;
          }

          const colLetter = columnIndexToLetter(used.columnIndex + col);
          const safeSheetName = sheet.name.replace(/'/g, "''");

          lines.push(
            `${normalized} = '${safeSheetName}'!${colLetter}${startRow}:${colLetter}${lastRow}`
          );
        }
      }

      // Tables
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

        const headerVals = (header.values && header.values[0]) || [];
        headerVals.forEach((h) => {
          if (!h) return;
          let norm = normalizeName(`${table.name}.${h}`);
          if (globalNameCounts[norm]) {
            globalNameCounts[norm] += 1;
            norm = `${norm}__${globalNameCounts[norm]}`;
          } else {
            globalNameCounts[norm] = 1;
          }

          // NOTE: header text may contain special chars; in practice Excel tolerates most,
          // but if needed, further escaping logic could be added.
          const structuredRef = `${table.name}[${h}]`;
          lines.push(`${norm} = ${structuredRef}`);
        });
      }

      // Pivot markers
      pivots.items.forEach((p) => lines.push(`PivotSource: ${p.name}`));
    }

    // Named ranges
    const names = wb.names;
    names.load("items/name");
    await ctx.sync();
    const meta = [];

    for (const n of names.items) {
      const r = n.getRange();
      r.load("address");
      meta.push({ name: n.name, range: r });
    }
    await ctx.sync();

    meta.forEach(({ name, range }) => {
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

    return lines.join("\n");
  });
}

export async function autoRefreshColumnMap(force = false) {
  if (refreshInProgress) return;
  try {
    const now = Date.now();
    if (!force && columnMapCache && now - lastColumnMapBuild < COLUMN_MAP_TTL_MS) {
      console.log("🔄 Using cached Smart Column Map (recent)");
      return;
    }

    refreshInProgress = true;
    console.log("🔄 Refreshing Smart Column Map…");
    columnMapCache = await buildColumnMapInternal();
    lastColumnMapBuild = Date.now();
    console.log("✅ Updated Smart Column Map");
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

// React to workbook structure changes (invalidate cache)
on("workbook:changed", () => {
  columnMapCache = "";
  lastColumnMapBuild = 0;
  console.log("🧹 Column map cache invalidated due to workbook change");
});
