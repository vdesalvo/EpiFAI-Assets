import { useState, useEffect, useRef } from "react";
import { ExcelName, onSelectionChange } from "@/lib/excel-names";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import { RadioGroup, RadioGroupItem } from "@/components/ui/radio-group";
import { Textarea } from "@/components/ui/textarea";
import { Grid, ArrowDownUp, MousePointerClick } from "lucide-react";
import { cn } from "@/lib/utils";

interface NameEditorProps {
  initialData?: ExcelName;
  onSave: (data: { name: string; refersTo: string; comment: string; newName?: string; skipRows?: number; skipCols?: number }) => void;
  onCancel: () => void;
}

function detectIsDynamic(formula: string): boolean {
  const u = formula.toUpperCase();
  return u.includes("OFFSET(") || u.includes("INDIRECT(") || u.includes("INDEX(");
}

function parseRangeRef(ref: string): { sheet: string; startCol: string; startRow: string; endCol: string; endRow: string } | null {
  const clean = ref.replace(/^=/, "").trim();
  const match = clean.match(/^(?:'?([^'!]+)'?!)?(\$?)([A-Za-z]+)(\$?)(\d+):(\$?)([A-Za-z]+)(\$?)(\d+)$/);
  if (!match) return null;
  return {
    sheet: match[1] || "",
    startCol: match[3].toUpperCase(),
    startRow: match[5],
    endCol: match[7].toUpperCase(),
    endRow: match[9],
  };
}

interface DynamicOptions {
  skipRows: number;
  skipCols: number;
}

function buildDynamicFormula(ref: string, options?: DynamicOptions): string {
  const parsed = parseRangeRef(ref);
  if (!parsed) return `=${ref.replace(/^=/, "")}`;

  const { sheet, startCol, startRow, endCol, endRow } = parsed;
  const sheetPrefix = sheet ? `${sheet}!` : "";
  const skipRows = options?.skipRows || 0;
  const skipCols = options?.skipCols || 0;

  const startColNum = colToNum(startCol);
  const endColNum = colToNum(endCol);
  const maxSkipCols = Math.max(0, endColNum - startColNum);
  const clampedSkipCols = Math.min(skipCols, maxSkipCols);
  const adjustedStartCol = numToCol(startColNum + clampedSkipCols);

  const startRowNum = parseInt(startRow, 10);
  const endRowNum = parseInt(endRow, 10);
  const maxSkipRows = Math.max(0, endRowNum - startRowNum);
  const clampedSkipRows = Math.min(skipRows, maxSkipRows);
  const adjustedStartRow = startRowNum + clampedSkipRows;

  const bufferColNum = endColNum + 20;
  const bufferCol = numToCol(Math.min(bufferColNum, 16384));

  const bufferRowNum = Math.min(endRowNum + 20, 1048576);

  const anchorCell = `${sheetPrefix}$${adjustedStartCol}$${adjustedStartRow}`;
  const rowCountRange = `${sheetPrefix}$${adjustedStartCol}$${adjustedStartRow}:$${adjustedStartCol}$${bufferRowNum}`;
  const colCountRange = `${sheetPrefix}$${adjustedStartCol}$${adjustedStartRow}:$${bufferCol}$${adjustedStartRow}`;

  return `=OFFSET(${anchorCell},0,0,COUNTA(${rowCountRange}),COUNTA(${colCountRange}))`;
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

export function NameEditor({ initialData, onSave, onCancel }: NameEditorProps) {
  const [name, setName] = useState(initialData?.name || "");
  const [refersTo, setRefersTo] = useState(initialData?.formula.replace(/^=/, "") || "");
  const [comment, setComment] = useState(initialData?.comment || "");
  const [type, setType] = useState<"fixed" | "dynamic">(
    initialData ? (detectIsDynamic(initialData.formula) ? "dynamic" : "fixed") : "fixed"
  );
  const [skipRows, setSkipRows] = useState(initialData?.skipRows || 0);
  const [skipCols, setSkipCols] = useState(initialData?.skipCols || 0);
  const [picking, setPicking] = useState(false);
  const unregRef = useRef<(() => Promise<void>) | null>(null);

  useEffect(() => {
    return () => {
      if (unregRef.current) {
        unregRef.current();
      }
    };
  }, []);

  const togglePicker = async () => {
    if (picking) {
      setPicking(false);
      if (unregRef.current) {
        await unregRef.current();
        unregRef.current = null;
      }
    } else {
      setPicking(true);
      try {
        const unreg = await onSelectionChange((address) => {
          setRefersTo(address);
        });
        unregRef.current = unreg;
      } catch (e) {
        console.error("Failed to start picker", e);
        setPicking(false);
      }
    }
  };

  const [nameError, setNameError] = useState("");

  const validateName = (n: string): string => {
    if (!n.trim()) return "Name is required";
    if (/\s/.test(n)) return "Name cannot contain spaces. Use underscores instead.";
    if (/^\d/.test(n)) return "Name cannot start with a number";
    if (!/^[A-Za-z_\\][A-Za-z0-9_.\\]*$/.test(n)) return "Name contains invalid characters";
    if (/^[A-Za-z]{1,3}\d+$/.test(n)) return "Name looks like a cell reference (e.g. A1)";
    return "";
  };

  const handleSave = () => {
    const err = validateName(name);
    if (err) {
      setNameError(err);
      return;
    }
    if (!refersTo.trim()) {
      setNameError("Reference is required");
      return;
    }
    setNameError("");

    let finalRef: string;
    if (type === "dynamic" && !detectIsDynamic(refersTo)) {
      finalRef = buildDynamicFormula(refersTo, { skipRows, skipCols });
    } else {
      finalRef = `=${refersTo.replace(/^=/, "")}`;
    }

    onSave({
      name: initialData?.name || name,
      newName: name !== initialData?.name ? name : undefined,
      refersTo: finalRef,
      comment,
      skipRows: type === "dynamic" ? skipRows : undefined,
      skipCols: type === "dynamic" ? skipCols : undefined,
    });
  };

  const previewFormula = type === "dynamic" && !detectIsDynamic(refersTo) && refersTo.trim()
    ? buildDynamicFormula(refersTo, { skipRows, skipCols })
    : null;

  return (
    <div className="flex flex-col h-full bg-background p-4 animate-in slide-in-from-right-4 duration-300">
      <div className="mb-6">
        <h2 className="text-lg font-bold text-foreground">
          {initialData ? "Edit Named Range" : "Create New Range"}
        </h2>
        <p className="text-xs text-muted-foreground">Configure the properties of your named range.</p>
      </div>

      <div className="space-y-5 flex-1 overflow-y-auto pr-1">
        <div className="space-y-2">
          <Label htmlFor="name" className="text-xs uppercase tracking-wider text-muted-foreground font-semibold">Name</Label>
          <Input
            id="name"
            value={name}
            onChange={e => { setName(e.target.value); setNameError(""); }}
            placeholder="e.g. Revenue_2024"
            className={cn("font-medium", nameError && "border-destructive")}
            data-testid="input-name"
          />
          {nameError && (
            <p className="text-[11px] text-destructive font-medium">{nameError}</p>
          )}
        </div>

        <div className="space-y-2">
          <Label className="text-xs uppercase tracking-wider text-muted-foreground font-semibold">Refers To</Label>
          <div className="flex gap-2">
            <Input
              value={refersTo}
              onChange={e => setRefersTo(e.target.value)}
              className={cn("font-mono text-xs", picking && "border-primary ring-1 ring-primary/20 bg-primary/5")}
              placeholder="Sheet1!$A$1:$B$10"
              data-testid="input-refers-to"
            />
            <Button
              type="button"
              variant={picking ? "default" : "outline"}
              size="icon"
              className={cn("shrink-0", picking ? "animate-pulse" : "")}
              onClick={togglePicker}
              title="Pick range from Excel"
              data-testid="button-range-picker"
            >
              <MousePointerClick className="w-4 h-4" />
            </Button>
          </div>
          {picking && (
            <p className="text-[10px] text-primary font-medium animate-pulse">
              Select cells in Excel to update the reference...
            </p>
          )}
        </div>

        <div className="space-y-3">
          <Label className="text-xs uppercase tracking-wider text-muted-foreground font-semibold">Range Type</Label>
          <RadioGroup
            value={type}
            onValueChange={(v) => setType(v as "fixed" | "dynamic")}
            className="grid grid-cols-2 gap-3"
          >
            <div className={cn(
              "flex items-center space-x-2 border rounded-md p-3 transition-colors cursor-pointer",
              type === "fixed" ? "bg-accent/50 border-primary/30" : "bg-card"
            )}>
              <RadioGroupItem value="fixed" id="fixed" data-testid="radio-fixed" />
              <Label htmlFor="fixed" className="cursor-pointer">
                <div className="flex items-center font-semibold text-sm mb-1">
                  <Grid className="w-3.5 h-3.5 mr-1.5" /> Fixed
                </div>
                <div className="text-[10px] text-muted-foreground leading-tight">
                  Static reference to specific cells.
                </div>
              </Label>
            </div>

            <div className={cn(
              "flex items-center space-x-2 border rounded-md p-3 transition-colors cursor-pointer",
              type === "dynamic" ? "bg-accent/50 border-primary/30" : "bg-card"
            )}>
              <RadioGroupItem value="dynamic" id="dynamic" data-testid="radio-dynamic" />
              <Label htmlFor="dynamic" className="cursor-pointer">
                <div className="flex items-center font-semibold text-sm mb-1">
                  <ArrowDownUp className="w-3.5 h-3.5 mr-1.5" /> Dynamic
                </div>
                <div className="text-[10px] text-muted-foreground leading-tight">
                  Auto-expands when rows or columns are added.
                </div>
              </Label>
            </div>
          </RadioGroup>
          {type === "dynamic" && (
            <div className="bg-muted/30 border rounded-md p-3 space-y-3">
              <p className="text-[10px] text-muted-foreground font-semibold uppercase tracking-wider">Skip Options</p>
              <div className="grid grid-cols-2 gap-3">
                <div className="space-y-1">
                  <Label htmlFor="skipRows" className="text-[11px] text-muted-foreground">Skip rows from top</Label>
                  <Input
                    id="skipRows"
                    type="number"
                    min={0}
                    value={skipRows}
                    onChange={e => setSkipRows(Math.max(0, parseInt(e.target.value) || 0))}
                    className="h-8 text-xs font-mono"
                    data-testid="input-skip-rows"
                  />
                </div>
                <div className="space-y-1">
                  <Label htmlFor="skipCols" className="text-[11px] text-muted-foreground">Skip columns from left</Label>
                  <Input
                    id="skipCols"
                    type="number"
                    min={0}
                    value={skipCols}
                    onChange={e => setSkipCols(Math.max(0, parseInt(e.target.value) || 0))}
                    className="h-8 text-xs font-mono"
                    data-testid="input-skip-cols"
                  />
                </div>
              </div>
              <p className="text-[10px] text-muted-foreground leading-snug">
                Use these to exclude header rows or label columns from the dynamic range. The range will start counting after the skipped rows/columns.
              </p>
            </div>
          )}
          {previewFormula && (
            <div className="bg-muted/50 border rounded-md p-2">
              <p className="text-[10px] text-muted-foreground mb-1 font-medium">Generated formula:</p>
              <p className="text-[11px] font-mono break-all text-foreground">{previewFormula}</p>
            </div>
          )}
        </div>

        <div className="space-y-2">
          <Label htmlFor="comment" className="text-xs uppercase tracking-wider text-muted-foreground font-semibold">Description</Label>
          <Textarea
            id="comment"
            value={comment}
            onChange={e => setComment(e.target.value)}
            placeholder="What is this range used for?"
            className="resize-none text-sm h-20"
            data-testid="input-comment"
          />
        </div>
      </div>

      <div className="flex gap-3 mt-6 pt-4 border-t">
        <Button className="flex-1 bg-primary" onClick={handleSave} data-testid="button-save">
          Save Changes
        </Button>
        <Button variant="outline" className="flex-1" onClick={onCancel} data-testid="button-cancel">
          Cancel
        </Button>
      </div>
    </div>
  );
}
