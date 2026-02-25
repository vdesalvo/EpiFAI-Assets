/// <reference path="./excel-types.d.ts" />

export interface ExcelTable {
  id: number;
  name: string;
  sheet: string;
  address: string;
  rowCount: number;
  colCount: number;
  hasHeaders: boolean;
}

export async function getAllTables(): Promise<ExcelTable[]> {
  try {
    return await Excel.run(async (ctx) => {
      const sheets = ctx.workbook.worksheets;
      sheets.load("items/name");
      await ctx.sync();
      const tables: ExcelTable[] = [];
      for (const sheet of sheets.items) {
        try {
          const tablesCol = sheet.tables;
          tablesCol.load("items/id,items/name,items/showHeaders");
          await ctx.sync();
          for (const table of tablesCol.items) {
            const range = table.getRange();
            range.load("address,rowCount,columnCount");
            await ctx.sync();
            tables.push({
              id: table.id,
              name: table.name,
              sheet: sheet.name,
              address: range.address,
              rowCount: range.rowCount,
              colCount: range.columnCount,
              hasHeaders: table.showHeaders,
            });
          }
        } catch (sheetErr) {
          console.warn(`Could not load tables for sheet "${sheet.name}":`, sheetErr);
        }
      }
      return tables;
    });
  } catch (error) {
    console.error("getAllTables error:", error);
    if (import.meta.env.DEV && !window.hasOwnProperty('Excel')) {
      return [
        { id: 1, name: "Table1", sheet: "Sheet1", address: "Sheet1!A1:D10", rowCount: 10, colCount: 4, hasHeaders: true },
      ];
    }
    throw error;
  }
}

export async function createTable(sheetName: string, rangeAddress: string, tableName: string, hasHeaders: boolean): Promise<string> {
  return Excel.run(async (ctx) => {
    const sheet = ctx.workbook.worksheets.getItem(sheetName);
    const range = sheet.getRange(rangeAddress);
    const table = sheet.tables.add(range, hasHeaders);
    table.name = tableName;
    table.load("name");
    await ctx.sync();
    return table.name;
  });
}

export async function deleteTable(tableName: string): Promise<void> {
  return Excel.run(async (ctx) => {
    const table = ctx.workbook.tables.getItem(tableName);
    table.convertToRange();
    await ctx.sync();
  });
}

export async function renameTable(oldName: string, newName: string): Promise<void> {
  return Excel.run(async (ctx) => {
    const table = ctx.workbook.tables.getItem(oldName);
    table.name = newName;
    await ctx.sync();
  });
}

export async function goToTable(sheetName: string, tableName: string): Promise<void> {
  return Excel.run(async (ctx) => {
    const sheet = ctx.workbook.worksheets.getItem(sheetName);
    sheet.activate();
    const table = ctx.workbook.tables.getItem(tableName);
    const range = table.getRange();
    range.select();
    await ctx.sync();
  });
}
