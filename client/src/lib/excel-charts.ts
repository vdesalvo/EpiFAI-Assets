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
      sheets.load("items/name");
      await ctx.sync();
      const charts: ExcelChart[] = [];
      for (const sheet of sheets.items) {
        try {
          const chartsCol = sheet.charts;
          chartsCol.load("items/name,items/id");
          await ctx.sync();
          for (const chart of chartsCol.items) {
            let titleText = "(No Title)";
            try {
              chart.load("title/text");
              await ctx.sync();
              titleText = chart.title?.text || "(No Title)";
            } catch { /* title not available */ }
            charts.push({
              id: chart.id,
              name: chart.name,
              title: titleText,
              sheet: sheet.name
            });
          }
        } catch (sheetErr) {
          console.warn(`Could not load charts for sheet "${sheet.name}":`, sheetErr);
        }
      }
      return charts;
    });
  } catch (error) {
    console.error("getAllCharts error:", error);
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

export async function createNameFromChart(sheetName: string, chartName: string, fallbackTitle: string): Promise<string> {
  const result = await Excel.run(async (ctx) => {
    const sheet = ctx.workbook.worksheets.getItem(sheetName);
    const chart = sheet.charts.getItem(chartName);
    chart.load("title/text");
    const names = ctx.workbook.names;
    names.load("items/name");
    const seriesCollection = chart.series;
    seriesCollection.load("count");
    await ctx.sync();

    const title = chart.title?.text || fallbackTitle;
    const existing = names.items.map(n => n.name.toLowerCase());

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

    let address: string;
    try {
      const firstRange = sheet.getRange(allRefs[0]);
      const lastRange = sheet.getRange(allRefs[allRefs.length - 1]);
      const combined = firstRange.getBoundingRect(lastRange);
      combined.load("address");
      await ctx.sync();
      address = combined.address;
    } catch {
      address = allRefs.join(",");
    }

    return { title, address, existing };
  });

  const baseName = sanitizeChartTitle(result.title);
  let rangeName = baseName;
  let suffix = 1;
  while (result.existing.includes(rangeName.toLowerCase())) {
    suffix++;
    rangeName = `${baseName}_${suffix}`;
  }

  await addName(rangeName, `=${result.address}`, `Source: ${result.title}`);
  return rangeName;
}
