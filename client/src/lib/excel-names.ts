/// <reference path="./excel-types.d.ts" />

export const EPIFAI_TAG = "[Epifai]";
const SKIP_TAG_RE = /\[skip:(\d+),(\d+)\]/;
const FIXCOL_TAG_RE = /\[fixcols:(\d+)\]/;

export function parseSkipTag(comment: string): { skipRows: number; skipCols: number } {
  const m = comment.match(SKIP_TAG_RE);
  return m ? { skipRows: parseInt(m[1], 10), skipCols: parseInt(m[2], 10) } : { skipRows: 0, skipCols: 0 };
}

export function parseFixedColsTag(comment: string): number {
  const m = comment.match(FIXCOL_TAG_RE);
  return m ? parseInt(m[1], 10) : 0;
}

export function buildSkipTag(skipRows: number, skipCols: number): string {
  if (skipRows === 0 && skipCols === 0) return "";
  return `[skip:${skipRows},${skipCols}]`;
}

export function buildFixedColsTag(fixedCols: number): string {
  if (fixedCols <= 0) return "";
  return `[fixcols:${fixedCols}]`;
}

export function stripMetaTags(comment: string): string {
  return comment.replace(EPIFAI_TAG, "").replace(SKIP_TAG_RE, "").replace(FIXCOL_TAG_RE, "").trim();
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
  fixedCols: number;
}

export interface UpdateNameParams {
  newName?: string;
  refersTo?: string;
  comment?: string;
  skipRows?: number;
  skipCols?: number;
  fixedCols?: number;
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
        let address = "";
        let values: any[][] | undefined = undefined;
        let status: ExcelName["status"] = "valid";
        try {
          const r = item.getRange();
          r.load("address,values");
          await ctx.sync();
          address = r.address;
          values = r.values;
        } catch {
          status = "broken";
        }

        const rawComment = item.comment || "";
        const isEpifai = rawComment.includes(EPIFAI_TAG);
        const skip = parseSkipTag(rawComment);
        const cleanComment = stripMetaTags(rawComment);

        results.push({
          name: item.name,
          type: item.type,
          value: item.value,
          formula: item.formula,
          comment: cleanComment,
          visible: item.visible,
          scope: "Workbook",
          address,
          values,
          status,
          origin: isEpifai ? "epifai" : "excel",
          skipRows: skip.skipRows,
          skipCols: skip.skipCols,
          fixedCols: parseFixedColsTag(rawComment),
        });
      }

      for (const sheet of sheets.items) {
        const sn = sheet.names;
        sn.load("items/name,items/type,items/value,items/formula,items/visible,items/comment");
        await ctx.sync();
        for (const item of sn.items) {
          let address = "";
          let status: ExcelName["status"] = "valid";
          try {
            const r = item.getRange();
            r.load("address");
            await ctx.sync();
            address = r.address;
          } catch {
            status = "broken";
          }
          const rawComment2 = item.comment || "";
          const isEpifai2 = rawComment2.includes(EPIFAI_TAG);
          const skip2 = parseSkipTag(rawComment2);
          const cleanComment2 = stripMetaTags(rawComment2);

          results.push({
            name: item.name,
            type: item.type,
            value: item.value,
            formula: item.formula,
            comment: cleanComment2,
            visible: item.visible,
            scope: sheet.name,
            address,
            status,
            origin: isEpifai2 ? "epifai" : "excel",
            skipRows: skip2.skipRows,
            skipCols: skip2.skipCols,
            fixedCols: parseFixedColsTag(rawComment2),
          });
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
         { name: "Revenue_2024", type: "Range", value: 1000, formula: "=Sheet1!$A$1", comment: "Total Revenue", visible: true, scope: "Workbook", address: "Sheet1!$A$1", status: "valid", origin: "epifai" as const, skipRows: 0, skipCols: 0, fixedCols: 0 },
         { name: "Expenses_Q1", type: "Range", value: 500, formula: "=Sheet1!$B$2", comment: "", visible: true, scope: "Workbook", address: "Sheet1!$B$2", status: "valid", origin: "excel" as const, skipRows: 0, skipCols: 0, fixedCols: 0 },
         { name: "Broken_Ref", type: "Range", value: "#REF!", formula: "=Sheet1!$Z$99", comment: "Old ref", visible: true, scope: "Workbook", address: "", status: "broken", origin: "excel" as const, skipRows: 0, skipCols: 0, fixedCols: 0 }
       ];
    }
    throw error;
  }
}

export async function addName(name: string, formula: string, comment = "", scope = "Workbook", skipRows = 0, skipCols = 0, fixedCols = 0): Promise<void> {
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
    const fixColTag = buildFixedColsTag(fixedCols);
    const parts = [comment, skipTag, fixColTag, EPIFAI_TAG].filter(Boolean);
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
    const fixC = updates.fixedCols !== undefined ? updates.fixedCols : parseFixedColsTag(oldRaw);
    const skipTag = buildSkipTag(skipR, skipC);
    const fixColTag = buildFixedColsTag(fixC);
    const parts = [userComment, skipTag, fixColTag, hadEpifaiTag ? EPIFAI_TAG : ""].filter(Boolean);
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
      namedItem.load("formula");
      await ctx.sync();

      const formula = namedItem.formula.replace(/^=/, "");
      let sheetName = "";
      let cellRef = formula;
      if (formula.includes("!")) {
        sheetName = formula.split("!")[0].replace(/'/g, "");
        cellRef = formula.split("!")[1];
      }

      if (sheetName) {
        const sheet = ctx.workbook.worksheets.getItem(sheetName);
        sheet.activate();
        await ctx.sync();
        const range = sheet.getRange(cellRef);
        range.select();
        await ctx.sync();
      } else {
        const range = namedItem.getRange();
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
      const range = namedItem.getRange();
      range.load("worksheet");
      await ctx.sync();
      range.worksheet.activate();
      await ctx.sync();
      range.select();
      await ctx.sync();
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
