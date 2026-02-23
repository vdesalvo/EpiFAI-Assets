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
    newName?: string;
    skipRows?: number;
    skipCols?: number;
    fixedRef?: string;
    dynamicRef?: string;
    lastColOnly?: boolean;
    lastRowOnly?: boolean;
  }) => void;
  onCancel: () => void;
  onPickSelection: () => Promise<SelectionData | undefined>;
  isPicking?: boolean;
  editTarget?: { name: string; comment: string; formula: string } | null;
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

export function RangePicker({ onSave, onCancel, onPickSelection, isPicking, editTarget }: RangePickerProps) {
  const isEditing = !!editTarget;
  const [selectionData, setSelectionData] = useState<SelectionData | null>(null);
  const [loading, setLoading] = useState(false);
  const [name, setName] = useState(editTarget?.name || "");
  const [comment, setComment] = useState(editTarget?.comment || "");
  const [nameError, setNameError] = useState("");
  const { toast } = useToast();

  const [skippedRows, setSkippedRows] = useState<Set<number>>(new Set());
  const [skippedCols, setSkippedCols] = useState<Set<number>>(new Set());
  const [expandCols, setExpandCols] = useState(false);
  const [expandRows, setExpandRows] = useState(false);

  const handlePick = async () => {
    setLoading(true);
    try {
      const data = await onPickSelection();
      if (data) {
        setSelectionData(data);
        setSkippedRows(new Set());
        setSkippedCols(new Set());
        setExpandCols(false);
        setExpandRows(false);
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

  const validateName = (n: string): string => {
    if (!n.trim()) return "Name is required";
    if (/\s/.test(n)) return "No spaces allowed. Use underscores.";
    if (/^\d/.test(n)) return "Cannot start with a number";
    if (!/^[A-Za-z_][A-Za-z0-9_.]*$/.test(n)) return "Invalid characters";
    if (/^[A-Za-z]{1,3}\d+$/.test(n)) return "Looks like a cell reference";
    return "";
  };

  const buildResult = useCallback(() => {
    if (!selectionData || activeColIndices.length === 0 || activeRowIndices.length === 0) return null;

    const sp = sheetPrefix(selectionData.sheet);
    const activeCols = activeColIndices.map(i => selectionData.columns[i]);

    const colGroups: string[][] = [];
    for (const col of activeCols) {
      const last = colGroups[colGroups.length - 1];
      if (last && colToNum(col) === colToNum(last[last.length - 1]) + 1) {
        last.push(col);
      } else {
        colGroups.push([col]);
      }
    }

    const activeRows = activeRowIndices.map(i => selectionData.startRow + i);
    const rowGroups: number[][] = [];
    for (const row of activeRows) {
      const last = rowGroups[rowGroups.length - 1];
      if (last && row === last[last.length - 1] + 1) {
        last.push(row);
      } else {
        rowGroups.push([row]);
      }
    }

    const endRow = selectionData.endRow;
    const firstCol = activeCols[0];
    const dataStartRow = activeRows[0];

    if (!expandRows && !expandCols) {
      const parts: string[] = [];
      for (const cg of colGroups) {
        const c1 = cg[0];
        const c2 = cg[cg.length - 1];
        for (const rg of rowGroups) {
          const r1 = rg[0];
          const r2 = rg[rg.length - 1];
          parts.push(`${sp}$${c1}$${r1}:$${c2}$${r2}`);
        }
      }
      return {
        refersTo: `=${parts.join(",")}`,
        skipRows: topSkipCount,
        skipCols: leftSkipCount,
        fixedRef: "",
        dynamicRef: "",
        lastColOnly: false,
        lastRowOnly: false,
      };
    }

    const bufferRow = Math.min(endRow + 500, 1048576);
    const hasColGaps = colGroups.length > 1;
    const hasRowGaps = rowGroups.length > 1;

    const parts: string[] = [];
    for (let cgIdx = 0; cgIdx < colGroups.length; cgIdx++) {
      const cg = colGroups[cgIdx];
      const c1 = cg[0];
      const c2 = cg[cg.length - 1];
      const cgWidth = cg.length;
      const isLastColGroup = cgIdx === colGroups.length - 1;

      for (let rgIdx = 0; rgIdx < rowGroups.length; rgIdx++) {
        const rg = rowGroups[rgIdx];
        const r1 = rg[0];
        const r2 = rg[rg.length - 1];
        const rgHeight = rg.length;
        const isLastRowGroup = rgIdx === rowGroups.length - 1;

        const useExpandHeight = expandRows && (!hasRowGaps || isLastRowGroup);
        const useExpandWidth = expandCols && (!hasColGaps || isLastColGroup);

        if (useExpandHeight || useExpandWidth) {
          const anchor = `${sp}$${c1}$${r1}`;
          let height: string;
          if (useExpandHeight) {
            const rowCountRange = `${sp}$${c1}$${r1}:$${c1}$${bufferRow}`;
            height = `COUNTA(${rowCountRange})`;
          } else {
            height = String(rgHeight);
          }
          let width: string;
          if (useExpandWidth) {
            const bufferCol = numToCol(Math.min(colToNum(c2) + 500, 16384));
            const colCountRange = `${sp}$${c1}$${r1}:$${bufferCol}$${r1}`;
            width = `COUNTA(${colCountRange})`;
          } else {
            width = String(cgWidth);
          }
          parts.push(`OFFSET(${anchor},0,0,${height},${width})`);
        } else {
          parts.push(`${sp}$${c1}$${r1}:$${c2}$${r2}`);
        }
      }
    }

    return {
      refersTo: `=${parts.join(",")}`,
      skipRows: topSkipCount,
      skipCols: leftSkipCount,
      fixedRef: "",
      dynamicRef: "",
      lastColOnly: false,
      lastRowOnly: false,
    };
  }, [selectionData, activeColIndices, activeRowIndices, topSkipCount, leftSkipCount, expandRows, expandCols]);

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
      name: isEditing ? editTarget!.name : name,
      comment,
      ...(isEditing && name !== editTarget!.name ? { newName: name } : {}),
      ...result,
    });
  };

  const previewRows = selectionData ? Math.min(selectionData.values.length, MAX_PREVIEW_ROWS) : 0;
  const previewCols = selectionData ? Math.min(selectionData.colCount, MAX_PREVIEW_COLS) : 0;
  const hasMoreRows = selectionData ? selectionData.values.length > MAX_PREVIEW_ROWS : false;
  const hasMoreCols = selectionData ? selectionData.colCount > MAX_PREVIEW_COLS : false;

  const lastActiveColIdx = activeColIndices.length > 0 ? activeColIndices[activeColIndices.length - 1] : -1;
  const lastActiveRowIdx = activeRowIndices.length > 0 ? activeRowIndices[activeRowIndices.length - 1] : -1;

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
        <h2 className="text-lg font-bold text-foreground">{isEditing ? "Edit Named Range" : "Visual Range Picker"}</h2>
        <p className="text-xs text-muted-foreground">{isEditing ? "Re-pick the range in Excel and adjust settings." : "Select a range in Excel, then configure it visually."}</p>
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

        {isEditing && editTarget?.formula && (
          <div className="bg-muted/30 border rounded-md p-2.5">
            <p className="text-[10px] text-muted-foreground font-semibold uppercase tracking-wider mb-1">Current Formula</p>
            <p className="text-[11px] font-mono text-foreground/70 break-all">{editTarget.formula}</p>
          </div>
        )}

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
                          const isLastActive = expandCols && ci === lastActiveColIdx;

                          return (
                            <th key={ci} className="min-w-[48px] p-0 border-b border-r">
                              <button
                                type="button"
                                className={cn(
                                  "w-full px-1 py-1.5 text-center font-bold cursor-pointer select-none transition-all text-[10px]",
                                  isSkipped
                                    ? "bg-red-100 text-red-500 line-through dark:bg-red-950/30 dark:text-red-400"
                                    : isLastActive
                                      ? "bg-emerald-50 text-emerald-700 dark:bg-emerald-950/20 dark:text-emerald-300"
                                      : "bg-muted/50 text-muted-foreground hover:bg-muted"
                                )}
                                onClick={() => toggleColSkip(ci)}
                                title={isSkipped ? `Click to include column ${colLetter}` : `Click to skip column ${colLetter}`}
                                data-testid={`col-header-${colLetter}`}
                              >
                                {colLetter}{isSkipped ? " ✕" : ""}
                              </button>
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
                        const isLastActiveRow = expandRows && ri === lastActiveRowIdx;

                        return (
                          <tr key={ri}>
                            <td className="p-0 border-r border-b">
                              <button
                                type="button"
                                className={cn(
                                  "w-full text-center font-bold px-1 py-1 cursor-pointer select-none transition-all text-[10px]",
                                  isRowSkipped
                                    ? "bg-red-100 text-red-500 line-through dark:bg-red-950/30 dark:text-red-400"
                                    : isLastActiveRow
                                      ? "bg-emerald-50 text-emerald-700 dark:bg-emerald-950/20 dark:text-emerald-300"
                                      : "bg-muted/50 text-muted-foreground hover:bg-muted"
                                )}
                                onClick={() => toggleRowSkip(ri)}
                                title={isRowSkipped ? `Click to include row ${rowNum}` : `Click to skip row ${rowNum}`}
                                data-testid={`row-header-${rowNum}`}
                              >
                                {rowNum}{isRowSkipped ? " ✕" : ""}
                              </button>
                            </td>
                            {Array.from({ length: previewCols }, (_, ci) => {
                              const isDimmed = isRowSkipped || skippedCols.has(ci);
                              const cellVal = selectionData.values[ri]?.[ci];
                              const displayVal = cellVal === null || cellVal === undefined || cellVal === "" ? "" : String(cellVal);

                              return (
                                <td
                                  key={ci}
                                  className={cn(
                                    "px-1.5 py-1 border-b border-r truncate max-w-[80px] transition-all",
                                    isDimmed ? "bg-muted/10 text-muted-foreground/20" : ""
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

            <div className="bg-muted/30 border rounded-md p-3 space-y-3">
              <div className="flex items-center justify-between">
                <div>
                  <Label htmlFor="vp-expand-cols" className="text-[11px] text-muted-foreground">Expand Columns</Label>
                  <p className="text-[9px] text-muted-foreground leading-snug">New columns added to the right will be included automatically</p>
                </div>
                <Switch
                  id="vp-expand-cols"
                  checked={expandCols}
                  onCheckedChange={setExpandCols}
                  data-testid="switch-vp-expand-cols"
                />
              </div>
              <div className="flex items-center justify-between">
                <div>
                  <Label htmlFor="vp-expand-rows" className="text-[11px] text-muted-foreground">Expand Rows</Label>
                  <p className="text-[9px] text-muted-foreground leading-snug">New rows added to the bottom will be included automatically</p>
                </div>
                <Switch
                  id="vp-expand-rows"
                  checked={expandRows}
                  onCheckedChange={setExpandRows}
                  data-testid="switch-vp-expand-rows"
                />
              </div>
            </div>

            <div className="bg-muted/50 border rounded-md p-2.5 space-y-1">
              <p className="text-[10px] text-muted-foreground font-semibold uppercase tracking-wider">Formula Preview</p>
              <div className="text-[11px] text-foreground space-y-0.5">
                {(() => { const r = buildResult(); return r ? (
                  <p className="font-mono text-[10px] break-all" data-testid="text-formula-preview">{r.refersTo}</p>
                ) : null; })()}
                <p>{activeColIndices.length} active column{activeColIndices.length !== 1 ? "s" : ""}{skippedCols.size > 0 ? ` (${skippedCols.size} skipped)` : ""}, {activeRowIndices.length} active row{activeRowIndices.length !== 1 ? "s" : ""}{skippedRows.size > 0 ? ` (${skippedRows.size} skipped)` : ""}</p>
                {expandCols && <p className="text-emerald-600 dark:text-emerald-400">↔ Columns will auto-expand</p>}
                {expandRows && <p className="text-emerald-600 dark:text-emerald-400">↕ Rows will auto-expand</p>}
              </div>
            </div>

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
          {isEditing ? "Update Range" : "Create Range"}
        </Button>
        <Button variant="outline" className="flex-1" onClick={onCancel} data-testid="button-vp-cancel">
          Cancel
        </Button>
      </div>
    </div>
  );
}
