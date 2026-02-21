import { useState, useCallback, useMemo } from "react";
import { SelectionData } from "@/lib/excel-names";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import { Switch } from "@/components/ui/switch";
import { MousePointerClick, Loader2, ArrowLeft } from "lucide-react";
import { cn } from "@/lib/utils";
import { useToast } from "@/hooks/use-toast";

interface RangePickerProps {
  onSave: (data: {
    name: string;
    refersTo: string;
    comment: string;
    skipRows?: number;
    skipCols?: number;
    fixedRef?: string;
    dynamicRef?: string;
    lastColOnly?: boolean;
  }) => void;
  onCancel: () => void;
  onPickSelection: () => Promise<SelectionData | undefined>;
  isPicking?: boolean;
}

const MAX_PREVIEW_ROWS = 8;
const MAX_PREVIEW_COLS = 10;

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

function quoteSheet(name: string): string {
  if (!name) return "";
  if (/^[A-Za-z_][A-Za-z0-9_]*$/.test(name)) return name;
  return `'${name.replace(/'/g, "''")}'`;
}

function sheetPrefix(sheet: string): string {
  if (!sheet) return "";
  return `${quoteSheet(sheet)}!`;
}

export function RangePicker({ onSave, onCancel, onPickSelection, isPicking }: RangePickerProps) {
  const [selectionData, setSelectionData] = useState<SelectionData | null>(null);
  const [loading, setLoading] = useState(false);
  const [name, setName] = useState("");
  const [comment, setComment] = useState("");
  const [nameError, setNameError] = useState("");
  const { toast } = useToast();

  const [skippedRows, setSkippedRows] = useState<Set<number>>(new Set());
  const [skippedCols, setSkippedCols] = useState<Set<number>>(new Set());
  const [fixedBoundary, setFixedBoundary] = useState<number>(0);
  const [lastColOnly, setLastColOnly] = useState(false);

  const handlePick = async () => {
    setLoading(true);
    try {
      const data = await onPickSelection();
      if (data) {
        setSelectionData(data);
        setSkippedRows(new Set());
        setSkippedCols(new Set());
        setFixedBoundary(0);
        setLastColOnly(false);
      }
    } catch (e: any) {
      console.error("Failed to get selection:", e);
      toast({
        title: "Could not read selection",
        description: e?.message || "Make sure you have a range selected in Excel and try again.",
        variant: "destructive",
      });
    } finally {
      setLoading(false);
    }
  };

  const toggleRowSkip = useCallback((rowIdx: number) => {
    setSkippedRows(prev => {
      const next = new Set(prev);
      if (next.has(rowIdx)) next.delete(rowIdx);
      else next.add(rowIdx);
      return next;
    });
  }, []);

  const toggleColSkip = useCallback((colIdx: number) => {
    setSkippedCols(prev => {
      const next = new Set(prev);
      if (next.has(colIdx)) next.delete(colIdx);
      else next.add(colIdx);
      return next;
    });
  }, []);

  const setDividerAt = useCallback((colIdx: number) => {
    setFixedBoundary(prev => prev === colIdx ? 0 : colIdx);
  }, []);

  const activeColIndices = useMemo(() => {
    if (!selectionData) return [];
    return Array.from({ length: selectionData.colCount }, (_, i) => i)
      .filter(i => !skippedCols.has(i));
  }, [selectionData, skippedCols]);

  const activeRowIndices = useMemo(() => {
    if (!selectionData) return [];
    return Array.from({ length: selectionData.rowCount }, (_, i) => i)
      .filter(i => !skippedRows.has(i));
  }, [selectionData, skippedRows]);

  const topSkipCount = useMemo(() => {
    if (!selectionData) return 0;
    let count = 0;
    for (let i = 0; i < selectionData.rowCount; i++) {
      if (skippedRows.has(i)) count++;
      else break;
    }
    return count;
  }, [selectionData, skippedRows]);

  const leftSkipCount = useMemo(() => {
    if (!selectionData) return 0;
    let count = 0;
    for (let i = 0; i < selectionData.colCount; i++) {
      if (skippedCols.has(i)) count++;
      else break;
    }
    return count;
  }, [selectionData, skippedCols]);

  const isHybrid = fixedBoundary > 0;

  const summary = useMemo(() => {
    if (!selectionData) return null;
    const cols = selectionData.columns;
    const activeCols = activeColIndices.map(i => cols[i]);

    if (isHybrid) {
      const fixedCols = activeCols.slice(0, fixedBoundary);
      const dynCols = activeCols.slice(fixedBoundary);
      return {
        type: "hybrid" as const,
        fixedCols,
        dynamicCols: dynCols,
        skipRows: topSkipCount,
        skipCols: leftSkipCount,
        lastColOnly,
      };
    }

    return {
      type: "simple" as const,
      cols: activeCols,
      skipRows: topSkipCount,
      skipCols: leftSkipCount,
      lastColOnly: false,
    };
  }, [selectionData, activeColIndices, fixedBoundary, topSkipCount, leftSkipCount, isHybrid, lastColOnly]);

  const validateName = (n: string): string => {
    if (!n.trim()) return "Name is required";
    if (/\s/.test(n)) return "No spaces allowed. Use underscores.";
    if (/^\d/.test(n)) return "Cannot start with a number";
    if (!/^[A-Za-z_][A-Za-z0-9_.]*$/.test(n)) return "Invalid characters";
    if (/^[A-Za-z]{1,3}\d+$/.test(n)) return "Looks like a cell reference";
    return "";
  };

  const buildResult = useCallback(() => {
    if (!selectionData || !summary) return null;

    const sp = sheetPrefix(selectionData.sheet);
    const startRowNum = selectionData.startRow;

    if (summary.type === "simple") {
      const activeCols = summary.cols;
      if (activeCols.length === 0) return null;
      const firstCol = activeCols[0];
      const lastCol = activeCols[activeCols.length - 1];
      const dataStartRow = startRowNum + summary.skipRows;
      const dataStartCol = firstCol;
      const ref = `${sp}$${dataStartCol}$${dataStartRow}:$${lastCol}$${selectionData.endRow}`;
      return {
        refersTo: `=${ref}`,
        skipRows: summary.skipRows,
        skipCols: summary.skipCols,
        fixedRef: "",
        dynamicRef: "",
        lastColOnly: false,
      };
    }

    if (summary.fixedCols.length === 0 || summary.dynamicCols.length === 0) return null;

    const dataStartRow = startRowNum + summary.skipRows;
    const fixedFirst = summary.fixedCols[0];
    const fixedLast = summary.fixedCols[summary.fixedCols.length - 1];
    const dynFirst = summary.dynamicCols[0];
    const dynLast = summary.dynamicCols[summary.dynamicCols.length - 1];

    const qSheet = quoteSheet(selectionData.sheet);
    const sheetRef = qSheet ? `${qSheet}!` : "";
    const fixedRefStr = `${sheetRef}$${fixedFirst}$${dataStartRow}:$${fixedLast}$${selectionData.endRow}`;
    const dynamicRefStr = `${sheetRef}$${dynFirst}$${dataStartRow}:$${dynLast}$${selectionData.endRow}`;

    const bufferColNum = colToNum(dynLast) + 20;
    const bufferCol = numToCol(Math.min(bufferColNum, 16384));
    const bufferRowNum = Math.min(selectionData.endRow + 20, 1048576);

    const fixedPart = `${sp}$${fixedFirst}$${dataStartRow}:$${fixedLast}$${selectionData.endRow}`;
    const dynAnchor = `${sp}$${dynFirst}$${dataStartRow}`;
    const rowCountRange = `${sp}$${dynFirst}$${dataStartRow}:$${dynFirst}$${bufferRowNum}`;
    const colCountRange = `${sp}$${dynFirst}$${dataStartRow}:$${bufferCol}$${dataStartRow}`;

    let formula: string;
    if (summary.lastColOnly) {
      formula = `=${fixedPart},OFFSET(${dynAnchor},0,COUNTA(${colCountRange})-1,COUNTA(${rowCountRange}),1)`;
    } else {
      formula = `=${fixedPart},OFFSET(${dynAnchor},0,0,COUNTA(${rowCountRange}),COUNTA(${colCountRange}))`;
    }

    return {
      refersTo: formula,
      skipRows: summary.skipRows,
      skipCols: summary.skipCols,
      fixedRef: fixedRefStr,
      dynamicRef: dynamicRefStr,
      lastColOnly: summary.lastColOnly,
    };
  }, [selectionData, summary]);

  const handleSave = () => {
    const err = validateName(name);
    if (err) {
      setNameError(err);
      return;
    }
    const result = buildResult();
    if (!result) {
      setNameError("Please select a range first");
      return;
    }
    setNameError("");
    onSave({
      name,
      comment,
      ...result,
    });
  };

  const previewRows = selectionData ? Math.min(selectionData.values.length, MAX_PREVIEW_ROWS) : 0;
  const previewCols = selectionData ? Math.min(selectionData.colCount, MAX_PREVIEW_COLS) : 0;
  const hasMoreRows = selectionData ? selectionData.values.length > MAX_PREVIEW_ROWS : false;
  const hasMoreCols = selectionData ? selectionData.colCount > MAX_PREVIEW_COLS : false;

  return (
    <div className="flex flex-col h-full bg-background p-4 animate-in slide-in-from-right-4 duration-300">
      <div className="mb-4">
        <button
          onClick={onCancel}
          className="flex items-center gap-1 text-xs text-muted-foreground hover:text-foreground transition-colors mb-2"
          data-testid="button-back-range-picker"
        >
          <ArrowLeft className="w-3 h-3" /> Back
        </button>
        <h2 className="text-lg font-bold text-foreground">Visual Range Picker</h2>
        <p className="text-xs text-muted-foreground">Select a range in Excel, then configure it visually.</p>
      </div>

      <div className="space-y-4 flex-1 overflow-y-auto pr-1">
        <div className="space-y-2">
          <Label htmlFor="vp-name" className="text-xs uppercase tracking-wider text-muted-foreground font-semibold">Name</Label>
          <Input
            id="vp-name"
            value={name}
            onChange={e => { setName(e.target.value); setNameError(""); }}
            placeholder="e.g. Revenue_2024"
            className={cn("font-medium", nameError && "border-destructive")}
            data-testid="input-vp-name"
          />
          {nameError && (
            <p className="text-[11px] text-destructive font-medium">{nameError}</p>
          )}
        </div>

        <div className="space-y-2">
          <Label className="text-xs uppercase tracking-wider text-muted-foreground font-semibold">Excel Range</Label>
          <Button
            variant="outline"
            className="w-full justify-center gap-2"
            onClick={handlePick}
            disabled={loading}
            data-testid="button-pick-range"
          >
            {loading ? (
              <Loader2 className="w-4 h-4 animate-spin" />
            ) : (
              <MousePointerClick className="w-4 h-4" />
            )}
            {selectionData ? "Re-pick Selection" : "Pick Current Selection"}
          </Button>
          {selectionData && (
            <p className="text-[10px] text-muted-foreground font-mono text-center">
              {selectionData.address} ({selectionData.rowCount} rows × {selectionData.colCount} cols)
            </p>
          )}
        </div>

        {selectionData && (
          <>
            <div className="space-y-2">
              <div className="flex items-center justify-between">
                <Label className="text-xs uppercase tracking-wider text-muted-foreground font-semibold">Range Preview</Label>
                <span className="text-[9px] text-muted-foreground">Click headers to dim/skip</span>
              </div>
              <div className="border rounded-lg overflow-hidden bg-card">
                <div className="overflow-x-auto">
                  <table className="w-full text-[10px] border-collapse" data-testid="table-range-preview">
                    <thead>
                      <tr>
                        <th className="w-8 min-w-[32px] bg-muted/50 border-b border-r p-0" />
                        {Array.from({ length: previewCols }, (_, ci) => {
                          const colLetter = selectionData.columns[ci];
                          const isSkipped = skippedCols.has(ci);
                          const activeIdx = activeColIndices.indexOf(ci);
                          const isFixed = !isSkipped && isHybrid && activeIdx >= 0 && activeIdx < fixedBoundary;
                          const isDynamic = !isSkipped && isHybrid && activeIdx >= fixedBoundary;

                          return (
                            <th
                              key={ci}
                              className={cn(
                                "min-w-[48px] px-1 py-1.5 text-center font-bold cursor-pointer select-none border-b border-r transition-all",
                                isSkipped
                                  ? "bg-muted/30 text-muted-foreground/30 line-through"
                                  : isFixed
                                    ? "bg-blue-100 text-blue-700 dark:bg-blue-900/30 dark:text-blue-300"
                                    : isDynamic
                                      ? "bg-emerald-100 text-emerald-700 dark:bg-emerald-900/30 dark:text-emerald-300"
                                      : "bg-muted/50 text-muted-foreground hover:bg-muted"
                              )}
                              onClick={() => toggleColSkip(ci)}
                              title={isSkipped ? `Click to include column ${colLetter}` : `Click to skip column ${colLetter}`}
                              data-testid={`col-header-${colLetter}`}
                            >
                              {colLetter}
                            </th>
                          );
                        })}
                        {hasMoreCols && (
                          <th className="min-w-[32px] px-1 py-1.5 text-center bg-muted/30 text-muted-foreground/50 border-b text-[9px]">
                            +{selectionData.colCount - MAX_PREVIEW_COLS}
                          </th>
                        )}
                      </tr>
                    </thead>
                    <tbody>
                      {Array.from({ length: previewRows }, (_, ri) => {
                        const rowNum = selectionData.startRow + ri;
                        const isRowSkipped = skippedRows.has(ri);

                        return (
                          <tr key={ri}>
                            <td
                              className={cn(
                                "text-center font-bold px-1 py-1 border-r border-b cursor-pointer select-none transition-all",
                                isRowSkipped
                                  ? "bg-muted/30 text-muted-foreground/30 line-through"
                                  : "bg-muted/50 text-muted-foreground hover:bg-muted"
                              )}
                              onClick={() => toggleRowSkip(ri)}
                              title={isRowSkipped ? `Click to include row ${rowNum}` : `Click to skip row ${rowNum}`}
                              data-testid={`row-header-${rowNum}`}
                            >
                              {rowNum}
                            </td>
                            {Array.from({ length: previewCols }, (_, ci) => {
                              const isColSkipped = skippedCols.has(ci);
                              const isDimmed = isRowSkipped || isColSkipped;
                              const activeIdx = activeColIndices.indexOf(ci);
                              const isFixed = !isDimmed && isHybrid && activeIdx >= 0 && activeIdx < fixedBoundary;
                              const isDynamic = !isDimmed && isHybrid && activeIdx >= fixedBoundary;
                              const cellVal = selectionData.values[ri]?.[ci];
                              const displayVal = cellVal === null || cellVal === undefined || cellVal === "" ? "" : String(cellVal);

                              return (
                                <td
                                  key={ci}
                                  className={cn(
                                    "px-1.5 py-1 border-b border-r truncate max-w-[80px] transition-all",
                                    isDimmed
                                      ? "bg-muted/10 text-muted-foreground/20"
                                      : isFixed
                                        ? "bg-blue-50/50 dark:bg-blue-950/20"
                                        : isDynamic
                                          ? "bg-emerald-50/50 dark:bg-emerald-950/20"
                                          : ""
                                  )}
                                  title={displayVal}
                                  data-testid={`cell-${ri}-${ci}`}
                                >
                                  {displayVal}
                                </td>
                              );
                            })}
                            {hasMoreCols && (
                              <td className="px-1 py-1 border-b text-center text-muted-foreground/30 text-[9px]">…</td>
                            )}
                          </tr>
                        );
                      })}
                      {hasMoreRows && (
                        <tr>
                          <td className="text-center text-muted-foreground/40 text-[9px] py-1 border-r" colSpan={1}>…</td>
                          <td className="text-center text-muted-foreground/40 text-[9px] py-1" colSpan={previewCols + (hasMoreCols ? 1 : 0)}>
                            +{selectionData.values.length - MAX_PREVIEW_ROWS} more rows
                          </td>
                        </tr>
                      )}
                    </tbody>
                  </table>
                </div>
              </div>
            </div>

            {activeColIndices.length > 1 && (
              <div className="space-y-2">
                <Label className="text-xs uppercase tracking-wider text-muted-foreground font-semibold">
                  Fixed / Dynamic Split
                </Label>
                <p className="text-[10px] text-muted-foreground leading-snug">
                  Click a column below to set where fixed columns end and dynamic columns begin. Columns to the left are fixed (blue), to the right are dynamic (green).
                </p>
                <div className="flex gap-0.5 flex-wrap">
                  {activeColIndices.slice(0, previewCols).map((colIdx, activePos) => {
                    const colLetter = selectionData.columns[colIdx];
                    const isBeforeBoundary = isHybrid && activePos < fixedBoundary;
                    const isAtBoundary = isHybrid && activePos === fixedBoundary - 1;

                    return (
                      <button
                        key={colIdx}
                        onClick={() => setDividerAt(activePos + 1)}
                        className={cn(
                          "px-2 py-1 text-[10px] font-bold rounded border transition-all",
                          isBeforeBoundary
                            ? "bg-blue-100 text-blue-700 border-blue-300 dark:bg-blue-900/30 dark:text-blue-300 dark:border-blue-700"
                            : isHybrid
                              ? "bg-emerald-100 text-emerald-700 border-emerald-300 dark:bg-emerald-900/30 dark:text-emerald-300 dark:border-emerald-700"
                              : "bg-muted text-muted-foreground border-border hover:bg-accent",
                          isAtBoundary && "ring-2 ring-blue-400 ring-offset-1"
                        )}
                        title={`Set split after column ${colLetter}`}
                        data-testid={`split-col-${colLetter}`}
                      >
                        {colLetter}
                      </button>
                    );
                  })}
                </div>
                {isHybrid && (
                  <div className="flex gap-3 text-[10px]">
                    <span className="flex items-center gap-1">
                      <span className="w-2.5 h-2.5 rounded-sm bg-blue-200 border border-blue-300" /> Fixed ({summary?.type === "hybrid" ? summary.fixedCols.join(", ") : ""})
                    </span>
                    <span className="flex items-center gap-1">
                      <span className="w-2.5 h-2.5 rounded-sm bg-emerald-200 border border-emerald-300" /> Dynamic ({summary?.type === "hybrid" ? summary.dynamicCols.join(", ") : ""})
                    </span>
                  </div>
                )}
              </div>
            )}

            {isHybrid && (
              <div className="bg-muted/30 border rounded-md p-3">
                <div className="flex items-center justify-between">
                  <div>
                    <Label htmlFor="vp-last-col" className="text-[11px] text-muted-foreground">Last column only</Label>
                    <p className="text-[9px] text-muted-foreground leading-snug">Only include the rightmost dynamic column</p>
                  </div>
                  <Switch
                    id="vp-last-col"
                    checked={lastColOnly}
                    onCheckedChange={setLastColOnly}
                    data-testid="switch-vp-last-col"
                  />
                </div>
              </div>
            )}

            {summary && (
              <div className="bg-muted/50 border rounded-md p-2.5 space-y-1">
                <p className="text-[10px] text-muted-foreground font-semibold uppercase tracking-wider">Summary</p>
                <div className="text-[11px] text-foreground space-y-0.5">
                  {summary.skipRows > 0 && (
                    <p>Skip: {summary.skipRows} row{summary.skipRows > 1 ? "s" : ""} from top</p>
                  )}
                  {summary.skipCols > 0 && (
                    <p>Skip: {summary.skipCols} col{summary.skipCols > 1 ? "s" : ""} from left</p>
                  )}
                  {summary.type === "hybrid" ? (
                    <>
                      <p>
                        <span className="text-blue-600 font-semibold">Fixed:</span> {summary.fixedCols.join(", ")}
                      </p>
                      <p>
                        <span className="text-emerald-600 font-semibold">Dynamic:</span> {summary.dynamicCols.join(", ")}
                        {summary.lastColOnly ? " (last col only)" : " (auto-expand)"}
                      </p>
                    </>
                  ) : (
                    <p>
                      <span className="font-semibold">Range:</span> {summary.cols.join(", ")} (fixed)
                    </p>
                  )}
                </div>
              </div>
            )}

            <div className="space-y-2">
              <Label htmlFor="vp-comment" className="text-xs uppercase tracking-wider text-muted-foreground font-semibold">Description</Label>
              <Input
                id="vp-comment"
                value={comment}
                onChange={e => setComment(e.target.value)}
                placeholder="What is this range used for?"
                className="text-sm"
                data-testid="input-vp-comment"
              />
            </div>
          </>
        )}
      </div>

      <div className="flex gap-3 mt-4 pt-4 border-t">
        <Button
          className="flex-1 bg-primary"
          onClick={handleSave}
          disabled={!selectionData}
          data-testid="button-vp-save"
        >
          Create Range
        </Button>
        <Button variant="outline" className="flex-1" onClick={onCancel} data-testid="button-vp-cancel">
          Cancel
        </Button>
      </div>
    </div>
  );
}
