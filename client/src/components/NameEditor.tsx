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
  onSave: (data: { name: string; refersTo: string; comment: string; newName?: string }) => void;
  onCancel: () => void;
}

export function NameEditor({ initialData, onSave, onCancel }: NameEditorProps) {
  const [name, setName] = useState(initialData?.name || "");
  const [refersTo, setRefersTo] = useState(initialData?.formula.replace(/^=/, "") || "");
  const [comment, setComment] = useState(initialData?.comment || "");
  const [type, setType] = useState<"fixed" | "dynamic">("fixed");
  const [picking, setPicking] = useState(false);
  const unregRef = useRef<(() => Promise<void>) | null>(null);

  useEffect(() => {
    // Basic detection of dynamic range
    const isDyn = refersTo.toUpperCase().includes("OFFSET(") || 
                 refersTo.toUpperCase().includes("INDIRECT(");
    setType(isDyn ? "dynamic" : "fixed");
  }, []); // Only run once on mount for initial detection

  // Cleanup selection listener
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
    onSave({
      name: initialData?.name || name,
      newName: name !== initialData?.name ? name : undefined,
      refersTo: `=${refersTo.replace(/^=/, "")}`,
      comment
    });
  };

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
            />
            <Button
              type="button"
              variant={picking ? "default" : "outline"}
              size="icon"
              className={cn("shrink-0", picking ? "animate-pulse" : "")}
              onClick={togglePicker}
              title="Pick range from Excel"
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
              type === "fixed" ? "bg-accent/50 border-primary/30" : "bg-card hover:bg-muted/50"
            )}>
              <RadioGroupItem value="fixed" id="fixed" />
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
              type === "dynamic" ? "bg-accent/50 border-primary/30" : "bg-card hover:bg-muted/50"
            )}>
              <RadioGroupItem value="dynamic" id="dynamic" />
              <Label htmlFor="dynamic" className="cursor-pointer">
                <div className="flex items-center font-semibold text-sm mb-1">
                  <ArrowDownUp className="w-3.5 h-3.5 mr-1.5" /> Dynamic
                </div>
                <div className="text-[10px] text-muted-foreground leading-tight">
                  Auto-expands with data (OFFSET/INDEX).
                </div>
              </Label>
            </div>
          </RadioGroup>
        </div>

        <div className="space-y-2">
          <Label htmlFor="comment" className="text-xs uppercase tracking-wider text-muted-foreground font-semibold">Description</Label>
          <Textarea 
            id="comment" 
            value={comment} 
            onChange={e => setComment(e.target.value)} 
            placeholder="What is this range used for?"
            className="resize-none text-sm h-20"
          />
        </div>
      </div>

      <div className="flex gap-3 mt-6 pt-4 border-t">
        <Button className="flex-1 bg-primary hover:bg-primary/90" onClick={handleSave}>
          Save Changes
        </Button>
        <Button variant="outline" className="flex-1" onClick={onCancel}>
          Cancel
        </Button>
      </div>
    </div>
  );
}
