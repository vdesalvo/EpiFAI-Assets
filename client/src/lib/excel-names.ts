/// <reference path="./excel-types.d.ts" />

export const EPIFAI_TAG = "[Epifai]";
const SKIP_TAG_RE = /\[skip:(\d+),(\d+)\]/;
const FIXREF_TAG_RE = /\[fixref:([^\]]+)\]/;
const DYNREF_TAG_RE = /\[dynref:([^\]]+)\]/;
const LASTCOL_TAG = "[lastcol]";

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

export function stripMetaTags(comment: string): string {
  return comment
    .replace(EPIFAI_TAG, "")
    .replace(SKIP_TAG_RE, "")
    .replace(FIXREF_TAG_RE, "")
    .replace(DYNREF_TAG_RE, "")
    .replace(LASTCOL_TAG, "")
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

  let status: ExcelName["status"] = "valid";
  if (!address && !isDynamicFormula(item.formula || "")) {
    status = "broken";
  } else if (isBrokenValue(item.value)) {
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
        const f = (item.formula || "").toUpperCase();
        const val = item.value;
        if (isBrokenValue(val) || isDynamicFormula(f)) {
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
          const f = (item.formula || "").toUpperCase();
          const val = item.value;
          if (isBrokenValue(val) || isDynamicFormula(f)) {
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
         { name: "Revenue_2024", type: "Range", value: 1000, formula: "=Sheet1!$A$1", comment: "Total Revenue", visible: true, scope: "Workbook", address: "Sheet1!$A$1", status: "valid", origin: "epifai" as const, skipRows: 0, skipCols: 0, fixedRef: "", dynamicRef: "", lastColOnly: false },
         { name: "Expenses_Q1", type: "Range", value: 500, formula: "=Sheet1!$B$2", comment: "", visible: true, scope: "Workbook", address: "Sheet1!$B$2", status: "valid", origin: "excel" as const, skipRows: 0, skipCols: 0, fixedRef: "", dynamicRef: "", lastColOnly: false },
         { name: "Broken_Ref", type: "Range", value: "#REF!", formula: "=Sheet1!$Z$99", comment: "Old ref", visible: true, scope: "Workbook", address: "", status: "broken", origin: "excel" as const, skipRows: 0, skipCols: 0, fixedRef: "", dynamicRef: "", lastColOnly: false }
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
      item.comment = current ? `${current} ${EPIFAI_TAG}` : EPIFAI_TAG;
      await ctx.sync();
    }
  });
}

export async function addName(name: string, formula: string, comment = "", scope = "Workbook", skipRows = 0, skipCols = 0, fixedRef = "", dynamicRef = "", lastColOnly = false): Promise<void> {
  return Excel.run(async (ctx) => {
    const raw = formula.replace(/^=/, "");
    const hasFunction = /[A-Z]+\(/.test(raw.toUpperCase());

    let item;
    if (hasFunction) {
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
    const parts = [comment, skipTag, fixRefTag, dynRefTag, lastColTag, EPIFAI_TAG].filter(Boolean);
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
    const skipTag = buildSkipTag(skipR, skipC);
    const fixRefTag = buildFixedRefTag(fRef);
    const dynRefTag = buildDynamicRefTag(dRef);
    const lastColTag = buildLastColTag(lCol);
    const parts = [userComment, skipTag, fixRefTag, dynRefTag, lastColTag, hadEpifaiTag ? EPIFAI_TAG : ""].filter(Boolean);
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
  // Address format: "Sheet1!$A$1:$H$20" or "Sheet1!$A$1"
  const match = address.match(/^([^!]+)!\$([A-Z]+)\$(\d+)(?::\$([A-Z]+)\$(\d+))?$/);
  
  if (!match) {
    throw new Error(`Invalid address format: ${address}`);
  }

  const sheet = match[1];
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
