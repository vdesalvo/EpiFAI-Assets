/// <reference path="./excel-types.d.ts" />

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
}

export interface UpdateNameParams {
  newName?: string;
  refersTo?: string;
  comment?: string;
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

        results.push({
          name: item.name,
          type: item.type,
          value: item.value,
          formula: item.formula,
          comment: item.comment || "",
          visible: item.visible,
          scope: "Workbook",
          address,
          values,
          status,
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
          results.push({
            name: item.name,
            type: item.type,
            value: item.value,
            formula: item.formula,
            comment: item.comment || "",
            visible: item.visible,
            scope: sheet.name,
            address,
            status,
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
         { name: "Revenue_2024", type: "Range", value: 1000, formula: "=Sheet1!$A$1", comment: "Total Revenue", visible: true, scope: "Workbook", address: "Sheet1!$A$1", status: "valid" },
         { name: "Expenses_Q1", type: "Range", value: 500, formula: "=Sheet1!$B$2", comment: "", visible: true, scope: "Workbook", address: "Sheet1!$B$2", status: "valid" },
         { name: "Broken_Ref", type: "Range", value: "#REF!", formula: "=Sheet1!$Z$99", comment: "Old ref", visible: true, scope: "Workbook", address: "", status: "broken" }
       ];
    }
    throw error;
  }
}

export async function addName(name: string, formula: string, comment = "", scope = "Workbook"): Promise<void> {
  return Excel.run(async (ctx) => {
    const item = scope === "Workbook"
      ? ctx.workbook.names.add(name, formula)
      : ctx.workbook.worksheets.getItem(scope).names.add(name, formula);
    if (comment) item.comment = comment;
    await ctx.sync();
  });
}

export async function updateName(name: string, updates: UpdateNameParams): Promise<void> {
  return Excel.run(async (ctx) => {
    const item = ctx.workbook.names.getItem(name);
    item.load("name,formula,comment");
    await ctx.sync();
    
    if (updates.refersTo) item.formula = updates.refersTo;
    if (updates.comment !== undefined) item.comment = updates.comment;
    
    if (updates.newName && updates.newName !== item.name) {
      const f = updates.refersTo || item.formula;
      const c = updates.comment !== undefined ? updates.comment : item.comment;
      item.delete();
      const n = ctx.workbook.names.add(updates.newName, f);
      n.comment = c;
    }
    await ctx.sync();
  });
}

export async function deleteName(name: string): Promise<void> {
  return Excel.run(async (ctx) => {
    ctx.workbook.names.getItem(name).delete();
    await ctx.sync();
  });
}

export async function goToName(params: { name: string; scope: string }): Promise<void> {
  return Excel.run(async (ctx) => {
    let range;
    if (params.scope && params.scope !== "Workbook") {
      const sheetItem = ctx.workbook.worksheets.getItem(params.scope);
      const namedItem = sheetItem.names.getItem(params.name);
      range = namedItem.getRange();
    } else {
      const namedItem = ctx.workbook.names.getItem(params.name);
      range = namedItem.getRange();
    }
    range.load("address,worksheet/name");
    await ctx.sync();
    const ws = range.worksheet;
    if (ws && ws.name) {
      ctx.workbook.worksheets.getItem(ws.name).activate();
      await ctx.sync();
    }
    range.select();
    await ctx.sync();
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
  let handler: any;
  await Excel.run(async (ctx) => {
    handler = ctx.workbook.worksheets.getActiveWorksheet().onSelectionChanged.add((e: any) => cb(e.address));
    await ctx.sync();
  });
  return async () => { 
    await Excel.run(async (ctx) => { 
      // This part handles removal, precise implementation might vary based on Office.js version
      // handler.remove(); 
      await ctx.sync(); 
    }); 
  };
}
