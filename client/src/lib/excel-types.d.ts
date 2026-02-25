// Minimal type definitions for Office.js if not available
// This prevents compilation errors if the types aren't installed
// In a real project, you would install @types/office-js

declare namespace Excel {
  function run<T>(callback: (context: RequestContext) => Promise<T>): Promise<T>;
  
  interface RequestContext {
    workbook: Workbook;
    sync(): Promise<void>;
  }

  interface Workbook {
    names: NamedItemCollection;
    worksheets: WorksheetCollection;
    getSelectedRange(): Range;
  }

  interface WorksheetCollection {
    items: Worksheet[];
    getItem(name: string): Worksheet;
    getItemOrNullObject(name: string): Worksheet;
    getActiveWorksheet(): Worksheet;
    add(name?: string): Worksheet;
    load(propertyNames?: string | string[]): void;
  }

  interface Worksheet {
    name: string;
    names: NamedItemCollection;
    charts: ChartCollection;
    shapes: ShapeCollection;
    isNullObject: boolean;
    activate(): void;
    getRange(address?: string): Range;
    getRangeByIndexes(startRow: number, startColumn: number, rowCount: number, columnCount: number): Range;
    getUsedRangeOrNullObject(valuesOnly?: boolean): Range;
    onSelectionChanged: EventHandlers;
    load(propertyNames?: string | string[]): void;
  }

  interface ShapeCollection {
    items: Shape[];
    count: number;
    getItem(name: string): Shape;
    load(propertyNames?: string | string[]): void;
  }

  interface Shape {
    id: string;
    name: string;
    type: string;
    width: number;
    height: number;
    top: number;
    left: number;
    geometricShapeType: string;
    load(propertyNames?: string | string[]): void;
  }

  interface NamedItemCollection {
    items: NamedItem[];
    add(name: string, reference: string | Range): NamedItem;
    getItem(name: string): NamedItem;
    load(propertyNames?: string | string[]): void;
  }

  interface NamedItem {
    name: string;
    type: string;
    value: any;
    formula: string;
    comment: string;
    visible: boolean;
    getRange(): Range;
    getRangeOrNullObject(): Range;
    delete(): void;
    load(propertyNames?: string | string[]): void;
  }

  interface RangeFormat {
    font: RangeFont;
    autofitColumns(): void;
    autofitRows(): void;
  }

  interface RangeFont {
    bold: boolean;
    color: string;
    size: number;
  }

  interface Range {
    address: string;
    values: any[][];
    worksheet: Worksheet;
    isNullObject: boolean;
    columnCount: number;
    columnIndex: number;
    rowCount: number;
    rowIndex: number;
    format: RangeFormat;
    load(propertyNames?: string | string[]): void;
    select(): void;
    getBoundingRect(anotherRange: Range | string): Range;
  }

  interface ChartCollection {
    items: Chart[];
    count: number;
    getItem(name: string): Chart;
    load(propertyNames?: string | string[]): void;
  }

  interface ChartSeriesCollection {
    items: ChartSeries[];
    getItemAt(index: number): ChartSeries;
    load(propertyNames?: string | string[]): void;
    count: number;
  }

  interface ChartSeries {
    name: string;
    getDimensionDataSourceString(dimension: string): { value: string };
    getDimensionDataSourceType(dimension: string): { value: string };
    load(propertyNames?: string | string[]): void;
  }

  interface ChartTitle {
    text: string;
    load(propertyNames?: string | string[]): void;
  }

  interface Chart {
    id: string;
    name: string;
    title: ChartTitle;
    series: ChartSeriesCollection;
    activate(): void;
    getImage(): { value: string };
    load(propertyNames?: string | string[]): void;
  }

  interface EventHandlers {
    add(handler: (args: any) => void): any;
  }
}

declare var Office: {
  onReady(callback?: (info: { host: string; platform: string }) => void): Promise<{ host: string; platform: string }>;
};

interface Window {
  Office?: typeof Office;
}
