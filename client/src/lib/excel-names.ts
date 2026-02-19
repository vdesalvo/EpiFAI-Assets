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

      for (const item of names.items) {
        const rawComment = item.comment || "";
        const isEpifai = rawComment.includes(EPIFAI_TAG);
        const skip = parseSkipTag(rawComment);
        const cleanComment = stripMetaTags(rawComment);

        let status: ExcelName["status"] = "valid";
        const val = String(item.value || "");
        if (val.includes("#REF!")) {
          status = "broken";
        }

        results.push({
          name: item.name,
          type: item.type,
          value: item.value,
          formula: item.formula,
          comment: cleanComment,
          visible: item.visible,
          scope: "Workbook",
          address: "",
          values: undefined,
          status,
          origin: isEpifai ? "epifai" : "excel",
          skipRows: skip.skipRows,
          skipCols: skip.skipCols,
          fixedRef: parseFixedRefTag(rawComment),
          dynamicRef: parseDynamicRefTag(rawComment),
          lastColOnly: parseLastColTag(rawComment),
        });
      }

      for (const sheet of sheets.items) {
        const sn = sheet.names;
        sn.load("items/name,items/type,items/value,items/formula,items/visible,items/comment");
        await ctx.sync();

        for (const item of sn.items) {
          const rawComment = item.comment || "";
          const isEpifai = rawComment.includes(EPIFAI_TAG);
          const skip = parseSkipTag(rawComment);
          const cleanComment = stripMetaTags(rawComment);

          let status: ExcelName["status"] = "valid";
          const val = String(item.value || "");
          if (val.includes("#REF!")) {
            status = "broken";
          }

          results.push({
            name: item.name,
            type: item.type,
            value: item.value,
            formula: item.formula,
            comment: cleanComment,
            visible: item.visible,
            scope: sheet.name,
            address: "",
            values: undefined,
            status,
            origin: isEpifai ? "epifai" : "excel",
            skipRows: skip.skipRows,
            skipCols: skip.skipCols,
            fixedRef: parseFixedRefTag(rawComment),
            dynamicRef: parseDynamicRefTag(rawComment),
            lastColOnly: parseLastColTag(rawComment),
          });
        }
      }
      return results;
    });
  } catch (error) {
    console.error("Error fetching names:", error);
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

export async function resolveNameAddress(name: string, scope: string): Promise<string> {
  return Excel.run(async (ctx) => {
    try {
      let namedItem;
      if (scope && scope !== "Workbook") {
        namedItem = ctx.workbook.worksheets.getItem(scope).names.getItem(name);
      } else {
        namedItem = ctx.workbook.names.getItem(name);
      }
      const r = namedItem.getRange();
      r.load("address");
      await ctx.sync();
      return r.address;
    } catch {
      return "";
    }
  });
}

export async function tagAsEpifai(name: string, scope: string): Promise<void> {
  return Excel.run(async (ctx) => {
    let namedItem;
    if (scope && scope !== "Workbook") {
      namedItem = ctx.workbook.worksheets.getItem(scope).names.getItem(name);
    } else {
      namedItem = ctx.workbook.names.getItem(name);
    }
    namedItem.load("comment");
    await ctx.sync();
    if (!namedItem.comment.includes(EPIFAI_TAG)) {
      namedItem.comment = (namedItem.comment + " " + EPIFAI_TAG).trim();
      await ctx.sync();
    }
  });
}

export async function untagFromEpifai(name: string, scope: string): Promise<void> {
  return Excel.run(async (ctx) => {
    let namedItem;
    if (scope && scope !== "Workbook") {
      namedItem = ctx.workbook.worksheets.getItem(scope).names.getItem(name);
    } else {
      namedItem = ctx.workbook.names.getItem(name);
    }
    namedItem.load("comment");
    await ctx.sync();
    namedItem.comment = namedItem.comment.replace(EPIFAI_TAG, "").trim();
    await ctx.sync();
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

export async function getSelection(): Promise<string> {
  return Excel.run(async (ctx) => {
    const r = ctx.workbook.getSelectedRange();
    r.load("address");
    await ctx.sync();
    return r.address;
  });
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
