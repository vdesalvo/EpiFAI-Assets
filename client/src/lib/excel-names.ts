/// <reference path="./excel-types.d.ts" />

export const EPIFAI_TAG = "[Epifai]";
const SKIP_TAG_RE = /\[skip:(\d+),(\d+)\]/;
const FIXREF_TAG_RE = /\[fixref:([^\]]+)\]/;
const DYNREF_TAG_RE = /\[dynref:([^\]]+)\]/;
const LASTCOL_TAG = "[lastcol]";
const LASTROW_TAG = "[lastrow]";
const ORIGRANGE_TAG_RE = /\[origrange:([^\]]+)\]/;
const EXPANDROWS_TAG = "[expandrows]";
const EXPANDCOLS_TAG = "[expandcols]";
const SKIPCIDX_TAG_RE = /\[skipcidx:([^\]]+)\]/;
const SKIPRIDX_TAG_RE = /\[skipridx:([^\]]+)\]/;
const MULTIAREAS_TAG_RE = /\[multiareas:([^\]]+)\]/;

export function parseSkipTag(comment: string): { skipRows: number; skipCols: number } {
  const m = comment.match(SKIP_TAG_RE);
  return m ? { skipRows: parseInt(m[1], 10), skipCols: parseInt(m[2], 10) } : { skipRows: 0, skipCols: 0 };
}

export function parseFixedRefTag(comment: string): string {
  const m = comment.match(FIXREF_TAG_RE);
  return m ? m[1] : "";
}

export function parseDynamicRefTag(comment: string): string {
  const m = comment.match(DYNREF_TAG_RE);
  return m ? m[1] : "";
}

export function buildSkipTag(skipRows: number, skipCols: number): string {
  if (skipRows === 0 && skipCols === 0) return "";
  return `[skip:${skipRows},${skipCols}]`;
}

export function buildFixedRefTag(ref: string): string {
  if (!ref) return "";
  return `[fixref:${ref}]`;
}

export function buildDynamicRefTag(ref: string): string {
  if (!ref) return "";
  return `[dynref:${ref}]`;
}

export function parseLastColTag(comment: string): boolean {
  return comment.includes(LASTCOL_TAG);
}

export function buildLastColTag(lastColOnly: boolean): string {
  return lastColOnly ? LASTCOL_TAG : "";
}

export function parseLastRowTag(comment: string): boolean {
  return comment.includes(LASTROW_TAG);
}

export function buildLastRowTag(lastRowOnly: boolean): string {
  return lastRowOnly ? LASTROW_TAG : "";
}

export function parseOrigRangeTag(comment: string): string {
  const m = comment.match(ORIGRANGE_TAG_RE);
  return m ? m[1] : "";
}
export function buildOrigRangeTag(address: string): string {
  if (!address) return "";
  return `[origrange:${address}]`;
}
export function parseExpandRowsTag(comment: string): boolean {
  return comment.includes(EXPANDROWS_TAG);
}
export function parseExpandColsTag(comment: string): boolean {
  return comment.includes(EXPANDCOLS_TAG);
}
export function parseSkipColIndicesTag(comment: string): number[] {
  const m = comment.match(SKIPCIDX_TAG_RE);
  return m ? m[1].split(",").map(Number) : [];
}
export function parseSkipRowIndicesTag(comment: string): number[] {
  const m = comment.match(SKIPRIDX_TAG_RE);
  return m ? m[1].split(",").map(Number) : [];
}
export function buildExpandRowsTag(on: boolean): string { return on ? EXPANDROWS_TAG : ""; }
export function buildExpandColsTag(on: boolean): string { return on ? EXPANDCOLS_TAG : ""; }
export function buildSkipColIndicesTag(indices: number[]): string {
  if (!indices.length) return "";
  return `[skipcidx:${indices.join(",")}]`;
}
export function buildSkipRowIndicesTag(indices: number[]): string {
  if (!indices.length) return "";
  return `[skipridx:${indices.join(",")}]`;
}
export function parseMultiAreasTag(comment: string): string[] {
  const m = comment.match(MULTIAREAS_TAG_RE);
  if (!m) return [];
  return m[1].split("|").filter(Boolean);
}
export function buildMultiAreasTag(areas: string[]): string {
  if (!areas.length) return "";
  return `[multiareas:${areas.join("|")}]`;
}

export function stripMetaTags(comment: string): string {
  return comment
    .replace(EPIFAI_TAG, "")
    .replace(SKIP_TAG_RE, "")
    .replace(FIXREF_TAG_RE, "")
    .replace(DYNREF_TAG_RE, "")
    .replace(LASTCOL_TAG, "")
    .replace(LASTROW_TAG, "")
    .replace(ORIGRANGE_TAG_RE, "")
    .replace(EXPANDROWS_TAG, "")
    .replace(EXPANDCOLS_TAG, "")
    .replace(SKIPCIDX_TAG_RE, "")
    .replace(SKIPRIDX_TAG_RE, "")
    .replace(MULTIAREAS_TAG_RE, "")
    .trim();
}

export interface ExcelName {
  name: string;
  type: string;
  value: any;
  formula: string;
  comment: string;
  visible: boolean;
  scope: string;
  address: string;
  values?: any[][];
  status: "valid" | "broken" | "unknown";
  origin: "epifai" | "excel";
  skipRows: number;
  skipCols: number;
  fixedRef: string;
  dynamicRef: string;
  lastColOnly: boolean;
  lastRowOnly: boolean;
  origRange: string;
  expandRows: boolean;
  expandCols: boolean;
  skippedColIndices: number[];
  skippedRowIndices: number[];
  multiAreas: string[];
}

export interface UpdateNameParams {
  newName?: string;
  refersTo?: string;
  comment?: string;
  skipRows?: number;
  skipCols?: number;
  fixedRef?: string;
  dynamicRef?: string;
  lastColOnly?: boolean;
  lastRowOnly?: boolean;
  origRange?: string;
  expandRows?: boolean;
  expandCols?: boolean;
  skippedColIndices?: number[];
  skippedRowIndices?: number[];
  multiAreas?: string[];
}

export interface SelectionData {
  address: string;       // Full address like "Sheet1!$A$1:$H$20"
  sheet: string;         // Sheet name
  values: any[][];       // 2D array of cell values
  startCol: string;      // First column letter e.g. "A"
  endCol: string;        // Last column letter e.g. "H"
  startRow: number;      // First row number e.g. 1
  endRow: number;        // Last row number e.g. 20
  colCount: number;      // Number of columns
  rowCount: number;      // Number of rows
  columns: string[];     // Array of column letters ["A","B","C",...]
}

function isDynamicFormula(formula: string): boolean {
  const f = formula.toUpperCase();
  return f.includes("OFFSET(") || f.includes("INDIRECT(") || f.includes("INDEX(");
}

function isUnionFormula(formula: string): boolean {
  const raw = formula.replace(/^=/, "");
  return splitFormulaTopLevel(raw).length > 1;
}

function isBrokenValue(value: any): boolean {
  if (typeof value === "string") {
    const v = value.toUpperCase();
    return v === "#REF!" || v.includes("#REF!");
  }
  return false;
}

function buildExcelName(item: any, scope: string, address: string, values?: any[][]): ExcelName {
  const rawComment = item.comment || "";
  const isEpifai = rawComment.includes(EPIFAI_TAG);
  const skip = parseSkipTag(rawComment);

  const hasMultiAreas = parseMultiAreasTag(rawComment).length > 1;
  let status: ExcelName["status"] = "valid";
  if (!address && !isDynamicFormula(item.formula || "") && !isUnionFormula(item.formula || "") && !hasMultiAreas) {
    status = "broken";
  } else if (isBrokenValue(item.value) && !hasMultiAreas) {
    status = "broken";
  }

  return {
    name: item.name,
    type: item.type,
    value: item.value,
    formula: item.formula,
    comment: stripMetaTags(rawComment),
    visible: item.visible,
    scope,
    address,
    values,
    status,
    origin: isEpifai ? "epifai" : "excel",
    skipRows: skip.skipRows,
    skipCols: skip.skipCols,
    fixedRef: parseFixedRefTag(rawComment),
    dynamicRef: parseDynamicRefTag(rawComment),
    lastColOnly: parseLastColTag(rawComment),
    lastRowOnly: parseLastRowTag(rawComment),
    origRange: parseOrigRangeTag(rawComment),
    expandRows: parseExpandRowsTag(rawComment),
    expandCols: parseExpandColsTag(rawComment),
    skippedColIndices: parseSkipColIndicesTag(rawComment),
    skippedRowIndices: parseSkipRowIndicesTag(rawComment),
    multiAreas: parseMultiAreasTag(rawComment),
  };
}

export async function getAllNames(): Promise<ExcelName[]> {
  try {
    return await Excel.run(async (ctx) => {
      const wb = ctx.workbook;
      const names = wb.names;
      names.load("items/name,items/type,items/value,items/formula,items/visible,items/comment");
      const sheets = wb.worksheets;
      sheets.load("items/name");
      await ctx.sync();

      const results: ExcelName[] = [];

      const hasNullObjectApi = typeof names.items[0]?.getRangeOrNullObject === "function";

      const wbPending: { item: any; range: any | null; skip: boolean }[] = [];

      for (const item of names.items) {
        const f = item.formula || "";
        const val = item.value;
        if (isBrokenValue(val) || isDynamicFormula(f) || isUnionFormula(f)) {
          wbPending.push({ item, range: null, skip: true });
        } else if (hasNullObjectApi) {
          const r = item.getRangeOrNullObject();
          r.load("address,values,isNullObject");
          wbPending.push({ item, range: r, skip: false });
        } else {
          wbPending.push({ item, range: null, skip: false });
        }
      }

      for (const sheet of sheets.items) {
        sheet.names.load("items/name,items/type,items/value,items/formula,items/visible,items/comment");
      }

      await ctx.sync();

      for (const { item, range, skip } of wbPending) {
        if (skip) {
          results.push(buildExcelName(item, "Workbook", ""));
        } else if (range) {
          if (range.isNullObject) {
            results.push(buildExcelName(item, "Workbook", ""));
          } else {
            results.push(buildExcelName(item, "Workbook", range.address || "", range.values));
          }
        } else {
          let address = "";
          let values: any[][] | undefined;
          try {
            const r = item.getRange();
            r.load("address,values");
            await ctx.sync();
            address = r.address || "";
            values = r.values;
          } catch {
          }
          results.push(buildExcelName(item, "Workbook", address, values));
        }
      }

      const allSheetPending: { item: any; range: any | null; skip: boolean; sheetName: string }[] = [];

      for (const sheet of sheets.items) {
        for (const item of sheet.names.items) {
          const f = item.formula || "";
          const val = item.value;
          if (isBrokenValue(val) || isDynamicFormula(f) || isUnionFormula(f)) {
            allSheetPending.push({ item, range: null, skip: true, sheetName: sheet.name });
          } else if (hasNullObjectApi) {
            const r = item.getRangeOrNullObject();
            r.load("address,isNullObject");
            allSheetPending.push({ item, range: r, skip: false, sheetName: sheet.name });
          } else {
            allSheetPending.push({ item, range: null, skip: false, sheetName: sheet.name });
          }
        }
      }

      if (allSheetPending.some(p => !p.skip && p.range)) {
        await ctx.sync();
      }

      for (const { item, range, skip, sheetName } of allSheetPending) {
        if (skip) {
          results.push(buildExcelName(item, sheetName, ""));
        } else if (range) {
          if (range.isNullObject) {
            results.push(buildExcelName(item, sheetName, ""));
          } else {
            results.push(buildExcelName(item, sheetName, range.address || ""));
          }
        } else {
          let address = "";
          try {
            const r = item.getRange();
            r.load("address");
            await ctx.sync();
            address = r.address || "";
          } catch {
          }
          results.push(buildExcelName(item, sheetName, address));
        }
      }

      return results;
    });
  } catch (error) {
    console.error("Error fetching names:", error);
    // Return empty array or throw depending on desired behavior
    // For mock purposes in browser without Excel, we might want to return dummy data
    if (import.meta.env.DEV && !window.hasOwnProperty('Excel')) {
       console.warn("Mocking Excel Data for Development");
       return [
         { name: "Revenue_2024", type: "Range", value: 1000, formula: "=Sheet1!$A$1", comment: "Total Revenue", visible: true, scope: "Workbook", address: "Sheet1!$A$1", status: "valid", origin: "epifai" as const, skipRows: 0, skipCols: 0, fixedRef: "", dynamicRef: "", lastColOnly: false, lastRowOnly: false, origRange: "", expandRows: false, expandCols: false, skippedColIndices: [], skippedRowIndices: [] },
         { name: "Expenses_Q1", type: "Range", value: 500, formula: "=Sheet1!$B$2", comment: "", visible: true, scope: "Workbook", address: "Sheet1!$B$2", status: "valid", origin: "excel" as const, skipRows: 0, skipCols: 0, fixedRef: "", dynamicRef: "", lastColOnly: false, lastRowOnly: false, origRange: "", expandRows: false, expandCols: false, skippedColIndices: [], skippedRowIndices: [] },
         { name: "Broken_Ref", type: "Range", value: "#REF!", formula: "=Sheet1!$Z$99", comment: "Old ref", visible: true, scope: "Workbook", address: "", status: "broken", origin: "excel" as const, skipRows: 0, skipCols: 0, fixedRef: "", dynamicRef: "", lastColOnly: false, lastRowOnly: false, origRange: "", expandRows: false, expandCols: false, skippedColIndices: [], skippedRowIndices: [] }
       ];
    }
    throw error;
  }
}

export async function claimAsEpifai(name: string, scope: string): Promise<void> {
  return Excel.run(async (ctx) => {
    const item = scope && scope !== "Workbook"
      ? ctx.workbook.worksheets.getItem(scope).names.getItem(name)
      : ctx.workbook.names.getItem(name);
    item.load("comment");
    await ctx.sync();
    const current = item.comment || "";
    if (!current.includes(EPIFAI_TAG)) {
      item.comment = current ? `${EPIFAI_TAG} ${current}` : EPIFAI_TAG;
      await ctx.sync();
    }
  });
}

export async function addName(name: string, formula: string, comment = "", scope = "Workbook", skipRows = 0, skipCols = 0, fixedRef = "", dynamicRef = "", lastColOnly = false, lastRowOnly = false, origRange = "", expandRows = false, expandCols = false, skippedColIndices: number[] = [], skippedRowIndices: number[] = [], multiAreas: string[] = []): Promise<void> {
  return Excel.run(async (ctx) => {
    const raw = formula.replace(/^=/, "");
    const hasFunction = /[A-Z]+\(/.test(raw.toUpperCase());
    const isUnion = splitFormulaTopLevel(raw).length > 1;

    let item;
    if (hasFunction || isUnion) {
      const ref = raw.startsWith("=") ? raw : `=${raw}`;
      item = scope === "Workbook"
        ? ctx.workbook.names.add(name, ref)
        : ctx.workbook.worksheets.getItem(scope).names.add(name, ref);
    } else {
      let sheetName = "";
      let cellRef = raw;
      if (raw.includes("!")) {
        sheetName = raw.split("!")[0].replace(/'/g, "");
        cellRef = raw.split("!")[1];
      }

      let range;
      if (sheetName) {
        range = ctx.workbook.worksheets.getItem(sheetName).getRange(cellRef);
      } else {
        range = ctx.workbook.worksheets.getActiveWorksheet().getRange(cellRef);
      }

      item = scope === "Workbook"
        ? ctx.workbook.names.add(name, range)
        : ctx.workbook.worksheets.getItem(scope).names.add(name, range);
    }

    const skipTag = buildSkipTag(skipRows, skipCols);
    const fixRefTag = buildFixedRefTag(fixedRef);
    const dynRefTag = buildDynamicRefTag(dynamicRef);
    const lastColTag = buildLastColTag(lastColOnly);
    const lastRowTag = buildLastRowTag(lastRowOnly);
    const origRangeTag = buildOrigRangeTag(origRange);
    const expandRowsTag = buildExpandRowsTag(expandRows);
    const expandColsTag = buildExpandColsTag(expandCols);
    const skipCIdxTag = buildSkipColIndicesTag(skippedColIndices);
    const skipRIdxTag = buildSkipRowIndicesTag(skippedRowIndices);
    const multiAreasTag = buildMultiAreasTag(multiAreas);
    const parts = [EPIFAI_TAG, comment, skipTag, fixRefTag, dynRefTag, lastColTag, lastRowTag, origRangeTag, expandRowsTag, expandColsTag, skipCIdxTag, skipRIdxTag, multiAreasTag].filter(Boolean);
    item.comment = parts.join(" ");
    await ctx.sync();
  });
}

export async function updateName(name: string, updates: UpdateNameParams): Promise<void> {
  return Excel.run(async (ctx) => {
    const item = ctx.workbook.names.getItem(name);
    item.load("name,formula,comment");
    await ctx.sync();

    const oldRaw = item.comment || "";
    const hadEpifaiTag = oldRaw.includes(EPIFAI_TAG);
    
    if (updates.refersTo) item.formula = updates.refersTo;

    const userComment = updates.comment !== undefined ? updates.comment : stripMetaTags(oldRaw);
    const skipR = updates.skipRows !== undefined ? updates.skipRows : parseSkipTag(oldRaw).skipRows;
    const skipC = updates.skipCols !== undefined ? updates.skipCols : parseSkipTag(oldRaw).skipCols;
    const fRef = updates.fixedRef !== undefined ? updates.fixedRef : parseFixedRefTag(oldRaw);
    const dRef = updates.dynamicRef !== undefined ? updates.dynamicRef : parseDynamicRefTag(oldRaw);
    const lCol = updates.lastColOnly !== undefined ? updates.lastColOnly : parseLastColTag(oldRaw);
    const lRow = updates.lastRowOnly !== undefined ? updates.lastRowOnly : parseLastRowTag(oldRaw);
    const origR = updates.origRange !== undefined ? updates.origRange : parseOrigRangeTag(oldRaw);
    const expR = updates.expandRows !== undefined ? updates.expandRows : parseExpandRowsTag(oldRaw);
    const expC = updates.expandCols !== undefined ? updates.expandCols : parseExpandColsTag(oldRaw);
    const scIdx = updates.skippedColIndices !== undefined ? updates.skippedColIndices : parseSkipColIndicesTag(oldRaw);
    const srIdx = updates.skippedRowIndices !== undefined ? updates.skippedRowIndices : parseSkipRowIndicesTag(oldRaw);
    const skipTag = buildSkipTag(skipR, skipC);
    const fixRefTag = buildFixedRefTag(fRef);
    const dynRefTag = buildDynamicRefTag(dRef);
    const lastColTag = buildLastColTag(lCol);
    const lastRowTag = buildLastRowTag(lRow);
    const origRangeTag = buildOrigRangeTag(origR);
    const expandRowsTag = buildExpandRowsTag(expR);
    const expandColsTag = buildExpandColsTag(expC);
    const skipCIdxTag = buildSkipColIndicesTag(scIdx);
    const skipRIdxTag = buildSkipRowIndicesTag(srIdx);
    const mAreas = updates.multiAreas !== undefined ? updates.multiAreas : parseMultiAreasTag(oldRaw);
    const multiAreasTag = buildMultiAreasTag(mAreas);
    const parts = [hadEpifaiTag ? EPIFAI_TAG : "", userComment, skipTag, fixRefTag, dynRefTag, lastColTag, lastRowTag, origRangeTag, expandRowsTag, expandColsTag, skipCIdxTag, skipRIdxTag, multiAreasTag].filter(Boolean);
    item.comment = parts.join(" ");
    
    if (updates.newName && updates.newName !== item.name) {
      const f = updates.refersTo || item.formula;
      const c = item.comment;
      item.delete();
      const n = ctx.workbook.names.add(updates.newName, f);
      n.comment = c;
    }
    await ctx.sync();
  });
}

export async function deleteName(params: { name: string; scope: string }): Promise<void> {
  return Excel.run(async (ctx) => {
    let namedItem;
    if (params.scope && params.scope !== "Workbook") {
      namedItem = ctx.workbook.worksheets.getItem(params.scope).names.getItem(params.name);
    } else {
      namedItem = ctx.workbook.names.getItem(params.name);
    }
    namedItem.delete();
    await ctx.sync();
  });
}

export async function deleteBrokenNames(): Promise<number> {
  return Excel.run(async (ctx) => {
    const wb = ctx.workbook;
    const names = wb.names;
    names.load("items/name,items/type,items/value,items/formula,items/comment");
    const sheets = wb.worksheets;
    sheets.load("items/name");
    await ctx.sync();

    const toDelete: any[] = [];
    const hasNullObjectApi = names.items.length > 0 && typeof names.items[0].getRangeOrNullObject === "function";

    const wbCheck: { item: any; range: any | null }[] = [];
    for (const item of names.items) {
      const f = (item.formula || "").toUpperCase();
      if (isDynamicFormula(f)) continue;
      if (isBrokenValue(item.value)) {
        toDelete.push(item);
      } else if (hasNullObjectApi) {
        const r = item.getRangeOrNullObject();
        r.load("isNullObject");
        wbCheck.push({ item, range: r });
      }
    }

    for (const sheet of sheets.items) {
      sheet.names.load("items/name,items/type,items/value,items/formula,items/comment");
    }
    await ctx.sync();

    const sheetCheck: { item: any; range: any | null }[] = [];
    for (const sheet of sheets.items) {
      for (const item of sheet.names.items) {
        const f = (item.formula || "").toUpperCase();
        if (isDynamicFormula(f)) continue;
        if (isBrokenValue(item.value)) {
          toDelete.push(item);
        } else if (hasNullObjectApi) {
          const r = item.getRangeOrNullObject();
          r.load("isNullObject");
          sheetCheck.push({ item, range: r });
        }
      }
    }

    if (wbCheck.length > 0 || sheetCheck.length > 0) {
      await ctx.sync();
    }

    for (const { item, range } of wbCheck) {
      if (range && range.isNullObject) {
        toDelete.push(item);
      }
    }
    for (const { item, range } of sheetCheck) {
      if (range && range.isNullObject) {
        toDelete.push(item);
      }
    }

    for (const item of toDelete) {
      item.delete();
    }

    if (toDelete.length > 0) {
      await ctx.sync();
    }

    return toDelete.length;
  });
}

export async function goToName(params: { name: string; scope: string }): Promise<void> {
  return Excel.run(async (ctx) => {
    try {
      let namedItem;
      if (params.scope && params.scope !== "Workbook") {
        namedItem = ctx.workbook.worksheets.getItem(params.scope).names.getItem(params.name);
      } else {
        namedItem = ctx.workbook.names.getItem(params.name);
      }

      try {
        const range = namedItem.getRange();
        range.load("worksheet");
        await ctx.sync();
        range.worksheet.activate();
        await ctx.sync();
        range.select();
        await ctx.sync();
        return;
      } catch (_) {
      }

      namedItem.load("formula");
      await ctx.sync();

      const formula = namedItem.formula.replace(/^=/, "");
      const refMatch = formula.match(/(?:'([^']+)'|([A-Za-z0-9_]+))!\$?([A-Z]+)\$?(\d+)(?::\$?([A-Z]+)\$?(\d+))?/);
      if (refMatch) {
        const sheetName = refMatch[1] || refMatch[2];
        const cellRef = refMatch[5]
          ? `$${refMatch[3]}$${refMatch[4]}:$${refMatch[5]}$${refMatch[6]}`
          : `$${refMatch[3]}$${refMatch[4]}`;
        const sheet = ctx.workbook.worksheets.getItem(sheetName);
        sheet.activate();
        await ctx.sync();
        const range = sheet.getRange(cellRef);
        range.select();
        await ctx.sync();
      }
    } catch (err) {
      console.error("goToName error:", err);
      throw err;
    }
  });
}

export async function selectNameRange(params: { name: string; scope: string }): Promise<void> {
  return Excel.run(async (ctx) => {
    try {
      let namedItem;
      if (params.scope && params.scope !== "Workbook") {
        namedItem = ctx.workbook.worksheets.getItem(params.scope).names.getItem(params.name);
      } else {
        namedItem = ctx.workbook.names.getItem(params.name);
      }

      try {
        const range = namedItem.getRange();
        range.load("worksheet");
        await ctx.sync();
        range.worksheet.activate();
        await ctx.sync();
        range.select();
        await ctx.sync();
        return;
      } catch (_) {
      }

      namedItem.load("formula");
      await ctx.sync();
      const formula = namedItem.formula.replace(/^=/, "");
      const refMatch = formula.match(/(?:'([^']+)'|([A-Za-z0-9_]+))!\$?([A-Z]+)\$?(\d+)(?::\$?([A-Z]+)\$?(\d+))?/);
      if (refMatch) {
        const sheetName = refMatch[1] || refMatch[2];
        const cellRef = refMatch[5]
          ? `$${refMatch[3]}$${refMatch[4]}:$${refMatch[5]}$${refMatch[6]}`
          : `$${refMatch[3]}$${refMatch[4]}`;
        const sheet = ctx.workbook.worksheets.getItem(sheetName);
        sheet.activate();
        await ctx.sync();
        const range = sheet.getRange(cellRef);
        range.select();
        await ctx.sync();
      }
    } catch (err) {
      console.error("selectNameRange error:", err);
    }
  });
}

export async function exportNameToSheet(params: { name: string; scope: string }): Promise<{ rowCount: number; colCount: number }> {
  if (import.meta.env.DEV && !window.hasOwnProperty('Excel')) {
    throw new Error("Export requires Excel");
  }
  return Excel.run(async (ctx) => {
    let namedItem;
    if (params.scope && params.scope !== "Workbook") {
      namedItem = ctx.workbook.worksheets.getItem(params.scope).names.getItem(params.name);
    } else {
      namedItem = ctx.workbook.names.getItem(params.name);
    }

    namedItem.load("formula,comment");
    await ctx.sync();

    const formula = namedItem.formula;
    const rawFormula = formula.replace(/^=/, "");
    const rawComment = namedItem.comment || "";

    const storedAreas = parseMultiAreasTag(rawComment);
    const formulaParts = storedAreas.length > 1 ? storedAreas : splitFormulaTopLevel(rawFormula);

    const skip = parseSkipTag(rawComment);
    const skippedColIdx = new Set(parseSkipColIndicesTag(rawComment));
    const skippedRowIdx = new Set(parseSkipRowIndicesTag(rawComment));

    let partValues: any[][][] = [];
    const tempNames: string[] = [];

    try {
      const directRanges: { idx: number; range: any }[] = [];
      const fallbackParts: { idx: number; part: string }[] = [];

      for (let i = 0; i < formulaParts.length; i++) {
        const part = formulaParts[i].trim();
        if (!part) continue;

        const sheetMatch = part.match(/^(?:'([^']+)'|([A-Za-z0-9_]+))!/);
        const hasDynFunc = /OFFSET\(|INDIRECT\(|INDEX\(/i.test(part);

        if (!hasDynFunc && sheetMatch) {
          try {
            const sheetName = sheetMatch[1] || sheetMatch[2];
            const ref = part.substring(sheetMatch[0].length);
            const partRange = ctx.workbook.worksheets.getItem(sheetName).getRange(ref);
            partRange.load("values,address");
            directRanges.push({ idx: i, range: partRange });
          } catch (e) {
            fallbackParts.push({ idx: i, part });
          }
        } else {
          fallbackParts.push({ idx: i, part });
        }
      }

      if (directRanges.length > 0) {
        await ctx.sync();
      }

      const indexedValues: { idx: number; values: any[][] }[] = [];

      for (const dr of directRanges) {
        try {
          if (dr.range.values) {
            indexedValues.push({ idx: dr.idx, values: dr.range.values });
          } else {
            const part = formulaParts[dr.idx].trim();
            fallbackParts.push({ idx: dr.idx, part });
          }
        } catch (e) {
          const part = formulaParts[dr.idx].trim();
          fallbackParts.push({ idx: dr.idx, part });
        }
      }

      for (const fb of fallbackParts) {
        try {
          const tempName = `_epifai_tmp_${Date.now()}_${fb.idx}`;
          tempNames.push(tempName);
          ctx.workbook.names.add(tempName, `=${fb.part}`);
          await ctx.sync();
          const tempItem = ctx.workbook.names.getItem(tempName);
          const tempRange = tempItem.getRangeOrNullObject();
          tempRange.load("isNullObject,values");
          await ctx.sync();
          if (!tempRange.isNullObject) {
            indexedValues.push({ idx: fb.idx, values: tempRange.values });
          }
          tempItem.delete();
          await ctx.sync();
        } catch (e) {
          console.warn(`Export: fallback failed for part "${fb.part}"`, e);
        }
      }

      indexedValues.sort((a, b) => a.idx - b.idx);
      partValues = indexedValues.map(iv => iv.values);

      if (partValues.length === 0) {
        throw new Error("No data could be read from the named range.");
      }
    } catch (err: any) {
      for (const tn of tempNames) {
        try {
          const c = ctx.workbook.names.getItem(tn);
          c.delete();
          await ctx.sync();
        } catch (_) {}
      }
      if (err.message === "No data could be read from the named range.") throw err;
      throw new Error("Could not resolve this named range for export.");
    }

    let exportSheet = ctx.workbook.worksheets.getItemOrNullObject("Epifai_Export");
    exportSheet.load("isNullObject");
    await ctx.sync();

    if (exportSheet.isNullObject) {
      exportSheet = ctx.workbook.worksheets.add("Epifai_Export");
    }

    const usedRange = exportSheet.getUsedRangeOrNullObject();
    usedRange.load("isNullObject,columnCount,columnIndex");
    await ctx.sync();

    let startCol = 0;
    if (!usedRange.isNullObject) {
      startCol = usedRange.columnIndex + usedRange.columnCount + 1;
    }

    const filteredPartValues: any[][][] = [];
    for (const pv of partValues) {
      const pvCols = pv[0]?.length || 0;

      const hasSkips = skip.skipRows > 0 || skip.skipCols > 0 || skippedColIdx.size > 0 || skippedRowIdx.size > 0;
      if (!hasSkips || pvCols <= 1) {
        filteredPartValues.push(pv);
        continue;
      }

      const keepCols: number[] = [];
      for (let c = 0; c < pvCols; c++) {
        if (c < skip.skipCols) continue;
        if (skippedColIdx.has(c)) continue;
        keepCols.push(c);
      }

      const filtered: any[][] = [];
      for (let r = 0; r < pv.length; r++) {
        if (r < skip.skipRows) continue;
        if (skippedRowIdx.has(r)) continue;
        filtered.push(keepCols.map(c => pv[r][c]));
      }

      if (filtered.length > 0 && filtered[0].length > 0) {
        filteredPartValues.push(filtered);
      }
    }

    const finalParts = filteredPartValues.length > 0 ? filteredPartValues : partValues;

    const allSameWidth = finalParts.length > 1 &&
      finalParts.every(pv => (pv[0]?.length || 0) === (finalParts[0][0]?.length || 0));

    let exportData: any[][];
    let dataCols: number;

    if (allSameWidth) {
      exportData = [];
      for (const pv of finalParts) {
        for (const row of pv) {
          exportData.push(row);
        }
      }
      dataCols = finalParts[0][0]?.length || 0;
    } else {
      dataCols = finalParts.reduce((sum, pv) => sum + (pv[0]?.length || 0), 0);
      const maxRows = Math.max(...finalParts.map(v => v.length));
      exportData = [];
      for (let r = 0; r < maxRows; r++) {
        const row: any[] = [];
        for (const pv of finalParts) {
          const pvCols = pv[0]?.length || 0;
          if (r < pv.length) {
            row.push(...pv[r]);
          } else {
            row.push(...new Array(pvCols).fill(""));
          }
        }
        exportData.push(row);
      }
    }

    const totalRows = exportData.length;

    if (dataCols === 0 || totalRows === 0) {
      throw new Error("No data could be read from the named range.");
    }

    const headerRow: any[] = [params.name];
    for (let i = 1; i < dataCols; i++) headerRow.push("");
    const headerCell = exportSheet.getRangeByIndexes(0, startCol, 1, dataCols);
    headerCell.values = [headerRow];
    headerCell.format.font.bold = true;

    if (totalRows > 0 && dataCols > 0) {
      const dataRange = exportSheet.getRangeByIndexes(1, startCol, totalRows, dataCols);
      dataRange.values = exportData;
    }

    const fullRange = exportSheet.getRangeByIndexes(0, startCol, totalRows + 1, dataCols);
    fullRange.format.autofitColumns();

    exportSheet.activate();
    await ctx.sync();

    fullRange.select();
    await ctx.sync();

    return { rowCount: totalRows, colCount: dataCols };
  });
}

export function splitFormulaTopLevel(formula: string): string[] {
  const parts: string[] = [];
  let depth = 0;
  let inQuote = false;
  let current = "";
  for (let i = 0; i < formula.length; i++) {
    const ch = formula[i];
    if (inQuote) {
      current += ch;
      if (ch === "'" && i + 1 < formula.length && formula[i + 1] === "'") {
        current += formula[i + 1];
        i++;
      } else if (ch === "'") {
        inQuote = false;
      }
    } else if (ch === "'") {
      inQuote = true;
      current += ch;
    } else if (ch === "(") {
      depth++;
      current += ch;
    } else if (ch === ")") {
      depth--;
      current += ch;
    } else if (ch === "," && depth === 0) {
      parts.push(current);
      current = "";
    } else {
      current += ch;
    }
  }
  if (current) parts.push(current);
  return parts;
}

function colToNum(col: string): number {
  let num = 0;
  for (let i = 0; i < col.length; i++) {
    num = num * 26 + (col.charCodeAt(i) - 64);
  }
  return num;
}

function numToCol(num: number): string {
  let col = "";
  while (num > 0) {
    const rem = (num - 1) % 26;
    col = String.fromCharCode(65 + rem) + col;
    num = Math.floor((num - 1) / 26);
  }
  return col;
}

function parseRangeAddress(address: string): { sheet: string; startCol: string; startRow: number; endCol: string; endRow: number } {
  const match = address.match(/^(.+)!\$?([A-Z]+)\$?(\d+)(?::\$?([A-Z]+)\$?(\d+))?$/);
  
  if (!match) {
    throw new Error(`Invalid address format: ${address}`);
  }

  let sheet = match[1];
  if (sheet.startsWith("'") && sheet.endsWith("'")) {
    sheet = sheet.slice(1, -1);
  }
  const startCol = match[2];
  const startRow = parseInt(match[3], 10);
  const endCol = match[4] || startCol;
  const endRow = match[5] ? parseInt(match[5], 10) : startRow;

  return { sheet, startCol, startRow, endCol, endRow };
}

export async function getSelection(): Promise<string> {
  return Excel.run(async (ctx) => {
    const r = ctx.workbook.getSelectedRange();
    r.load("address");
    await ctx.sync();
    return r.address;
  });
}

export async function readRangeData(rangeAddress: string): Promise<SelectionData> {
  try {
    return await Excel.run(async (ctx) => {
      const { sheet, startCol, startRow, endCol, endRow } = parseRangeAddress(rangeAddress);
      const ws = ctx.workbook.worksheets.getItem(sheet);
      const cellRef = rangeAddress.includes("!") ? rangeAddress.split("!")[1] : rangeAddress;
      const range = ws.getRange(cellRef);
      range.load("address,values");
      await ctx.sync();

      const values = range.values;
      const rowCount = values ? values.length : (endRow - startRow + 1);
      const colCount = values && values.length > 0 ? values[0].length : (colToNum(endCol) - colToNum(startCol) + 1);
      const columns: string[] = [];
      const startColNum = colToNum(startCol);
      const endColNum = colToNum(endCol);
      for (let i = startColNum; i <= endColNum; i++) {
        columns.push(numToCol(i));
      }

      return { address: range.address, sheet, values, startCol, endCol, startRow, endRow, colCount, rowCount, columns };
    });
  } catch (error) {
    console.error("Error reading range data:", error);
    if (import.meta.env.DEV && !window.hasOwnProperty('Excel')) {
      console.warn("Mocking readRangeData for Development");
      return {
        address: rangeAddress,
        sheet: "Sheet1",
        values: [
          ["Header1", "Header2", "Header3", "Header4", "Header5", "Header6", "Header7", "Header8"],
          [1, 2, 3, 4, 5, 6, 7, 8],
          [10, 20, 30, 40, 50, 60, 70, 80],
        ],
        startCol: "A",
        endCol: "H",
        startRow: 1,
        endRow: 20,
        colCount: 8,
        rowCount: 20,
        columns: ["A", "B", "C", "D", "E", "F", "G", "H"],
      };
    }
    throw error;
  }
}

export async function getSelectionData(): Promise<SelectionData> {
  try {
    return await Excel.run(async (ctx) => {
      const worksheet = ctx.workbook.worksheets.getActiveWorksheet();
      const range = ctx.workbook.getSelectedRange();
      
      worksheet.load("name");
      range.load("address,values");
      await ctx.sync();

      const address = range.address;
      const sheet = worksheet.name;
      const values = range.values;

      // Parse address to extract startCol, endCol, startRow, endRow
      const { startCol, startRow, endCol, endRow } = parseRangeAddress(address);

      // Calculate row and column counts from the values array
      const rowCount = values ? values.length : (endRow - startRow + 1);
      const colCount = values && values.length > 0 ? values[0].length : (colToNum(endCol) - colToNum(startCol) + 1);

      // Build columns array
      const columns: string[] = [];
      const startColNum = colToNum(startCol);
      const endColNum = colToNum(endCol);
      for (let i = startColNum; i <= endColNum; i++) {
        columns.push(numToCol(i));
      }

      return {
        address,
        sheet,
        values,
        startCol,
        endCol,
        startRow,
        endRow,
        colCount,
        rowCount,
        columns,
      };
    });
  } catch (error) {
    console.error("Error fetching selection data:", error);
    // Return mock data for dev/fallback
    if (import.meta.env.DEV && !window.hasOwnProperty('Excel')) {
      console.warn("Mocking Excel Selection Data for Development");
      return {
        address: "Sheet1!$A$1:$H$20",
        sheet: "Sheet1",
        values: [
          ["Header1", "Header2", "Header3", "Header4", "Header5", "Header6", "Header7", "Header8"],
          [1, 2, 3, 4, 5, 6, 7, 8],
          [10, 20, 30, 40, 50, 60, 70, 80],
        ],
        startCol: "A",
        endCol: "H",
        startRow: 1,
        endRow: 20,
        colCount: 8,
        rowCount: 20,
        columns: ["A", "B", "C", "D", "E", "F", "G", "H"],
      };
    }
    throw error;
  }
}

export interface MultiAreaItem {
  address: string;
  rows: number;
  cols: number;
}

export async function getMultiAreaSelectionData(): Promise<MultiAreaItem[]> {
  try {
    return await Excel.run(async (ctx) => {
      const worksheet = ctx.workbook.worksheets.getActiveWorksheet();
      worksheet.load("name");

      const rangeAreas = ctx.workbook.getSelectedRanges();
      rangeAreas.load("areaCount");
      rangeAreas.areas.load("items");
      await ctx.sync();

      const results: MultiAreaItem[] = [];

      for (const area of rangeAreas.areas.items) {
        area.load("address,rowCount,columnCount");
      }
      await ctx.sync();

      const sheet = worksheet.name;
      const quotedSheet = /^[A-Za-z_][A-Za-z0-9_]*$/.test(sheet) ? sheet : `'${sheet.replace(/'/g, "''")}'`;
      for (const area of rangeAreas.areas.items) {
        const addr = area.address.includes("!") ? area.address : `${quotedSheet}!${area.address}`;
        results.push({
          address: addr,
          rows: area.rowCount,
          cols: area.columnCount,
        });
      }

      return results;
    });
  } catch (error) {
    console.error("Error fetching multi-area selection data:", error);
    if (import.meta.env.DEV && !window.hasOwnProperty('Excel')) {
      console.warn("Mocking multi-area selection data for development");
      return [
        { address: "Sheet1!$A$1:$B$5", rows: 5, cols: 2 },
        { address: "Sheet1!$D$1:$E$5", rows: 5, cols: 2 },
      ];
    }
    throw error;
  }
}

export async function onSelectionChange(cb: (address: string) => void): Promise<() => Promise<void>> {
  let active = true;
  let lastAddr = "";

  const poll = async () => {
    while (active) {
      try {
        const addr = await Excel.run(async (ctx) => {
          const sheet = ctx.workbook.worksheets.getActiveWorksheet();
          const range = ctx.workbook.getSelectedRange();
          sheet.load("name");
          range.load("address");
          await ctx.sync();
          const cellRef = range.address.includes("!") ? range.address.split("!")[1] : range.address;
          return `${sheet.name}!${cellRef}`;
        });
        if (addr !== lastAddr) {
          lastAddr = addr;
          cb(addr);
        }
      } catch (e) {
        console.error("Selection poll error:", e);
      }
      await new Promise(r => setTimeout(r, 500));
    }
  };

  poll();

  return async () => {
    active = false;
  };
}
