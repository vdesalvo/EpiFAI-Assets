/// <reference path="./excel-types.d.ts" />
import { addName } from "./excel-names";

export interface ExcelChart {
  id: string;
  name: string;
  title: string;
  sheet: string;
  isDefault?: boolean;
}

export async function getAllCharts(): Promise<ExcelChart[]> {
  try {
    return await Excel.run(async (ctx) => {
      const sheets = ctx.workbook.worksheets;
      sheets.load("items/name,items/charts");
      await ctx.sync();
      const charts: ExcelChart[] = [];
      for (const sheet of sheets.items) {
        sheet.charts.load("items/name,items/id,items/title/text");
        await ctx.sync();
        for (const chart of sheet.charts.items) {
          charts.push({
            id: chart.id, 
            name: chart.name, 
            title: chart.title?.text || "(No Title)", 
            sheet: sheet.name
          });
        }
      }
      return charts;
    });
  } catch (error) {
    if (import.meta.env.DEV && !window.hasOwnProperty('Excel')) {
      return [
        { id: "c1", name: "Chart 1", title: "Sales 2024", sheet: "Sheet1" },
        { id: "c2", name: "Chart 2", title: "Growth", sheet: "Sheet1" }
      ];
    }
    throw error;
  }
}

export async function renameChart(sheetName: string, oldName: string, newName: string): Promise<void> {
  return Excel.run(async (ctx) => {
    const sheet = ctx.workbook.worksheets.getItem(sheetName);
    const chart = sheet.charts.getItem(oldName);
    chart.name = newName;
    await ctx.sync();
  });
}

export async function getChartImage(sheetName: string, chartName: string): Promise<string> {
  return Excel.run(async (ctx) => {
    const sheet = ctx.workbook.worksheets.getItem(sheetName);
    const chart = sheet.charts.getItem(chartName);
    const img = chart.getImage();
    await ctx.sync();
    return img.value;
  });
}

export async function goToChart(sheetName: string, chartName: string): Promise<void> {
  return Excel.run(async (ctx) => {
    const sheet = ctx.workbook.worksheets.getItem(sheetName);
    sheet.activate();
    const chart = sheet.charts.getItem(chartName);
    chart.activate();
    await ctx.sync();
  });
}

export function isDefaultName(name: string): boolean {
  return /^Chart \d+$/.test(name);
}

export function sanitizeChartTitle(title: string): string {
  return title
    .replace(/[^A-Za-z0-9_.\\]/g, "_")
    .replace(/^(\d)/, "_$1")
    .replace(/_+/g, "_")
    .replace(/^_|_$/g, "") || "ChartName";
}

export async function createNameFromChart(sheetName: string, chartName: string, title: string): Promise<string> {
  const baseName = sanitizeChartTitle(title);

  const existingNames = await Excel.run(async (ctx) => {
    const names = ctx.workbook.names;
    names.load("items/name");
    await ctx.sync();
    return names.items.map(n => n.name.toLowerCase());
  });

  let rangeName = baseName;
  let suffix = 1;
  while (existingNames.includes(rangeName.toLowerCase())) {
    suffix++;
    rangeName = `${baseName}_${suffix}`;
  }

  const address = await Excel.run(async (ctx) => {
    const sheet = ctx.workbook.worksheets.getItem(sheetName);
    const chart = sheet.charts.getItem(chartName);
    const seriesCollection = chart.series;
    seriesCollection.load("count");
    await ctx.sync();

    const rangeAddresses: string[] = [];

    for (let i = 0; i < seriesCollection.count; i++) {
      const series = seriesCollection.getItemAt(i);
      const valuesSource = series.getDimensionDataSourceString("Values");
      await ctx.sync();
      if (valuesSource.value) {
        rangeAddresses.push(valuesSource.value);
      }
    }

    if (rangeAddresses.length === 0) {
      throw new Error("Could not read chart data range");
    }

    const allRefs: string[] = [];
    for (const addr of rangeAddresses) {
      for (const part of addr.split(",")) {
        const trimmed = part.trim();
        if (trimmed && !allRefs.includes(trimmed)) {
          allRefs.push(trimmed);
        }
      }
    }

    try {
      const firstRange = sheet.getRange(allRefs[0]);
      firstRange.load("address");
      const lastRange = sheet.getRange(allRefs[allRefs.length - 1]);
      lastRange.load("address");
      await ctx.sync();

      const combined = firstRange.getBoundingRect(lastRange);
      combined.load("address");
      await ctx.sync();
      return combined.address;
    } catch {
      return allRefs.join(",");
    }
  });

  await addName(rangeName, `=${address}`, `Source: ${title}`);
  return rangeName;
}
