import { useState } from "react";
import { ExcelName, selectNameRange } from "@/lib/excel-names";
import { Badge } from "@/components/ui/badge";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { 
  Trash2, 
  ExternalLink, 
  Edit, 
  Search, 
  AlertCircle,
  CheckCircle2,
  HelpCircle,
  LayoutGrid,
  Maximize,
  Columns,
  Loader2,
  FileSpreadsheet
} from "lucide-react";
import { cn } from "@/lib/utils";

interface NameListProps {
  names: ExcelName[];
  onEdit: (name: ExcelName) => void;
  onDelete: (name: ExcelName) => void;
  onGoTo: (name: ExcelName) => void;
  onCreate: () => void;
  onClaim?: (name: ExcelName) => void;
  onExport?: (name: ExcelName) => void;
  isExporting?: boolean;
  onDeleteBroken?: () => void;
  isDeletingBroken?: boolean;
  pendingDeleteName?: string | null;
}

export function NameList({ names, onEdit, onDelete, onGoTo, onCreate, onClaim, onExport, isExporting, onDeleteBroken, isDeletingBroken, pendingDeleteName }: NameListProps) {
  const [search, setSearch] = useState("");
  const [filter, setFilter] = useState<"all" | "epifai" | "excel" | "broken">("epifai");
  const [selectedId, setSelectedId] = useState<string | null>(null);
  const [confirmDeleteBroken, setConfirmDeleteBroken] = useState(false);

  const stats = {
    total: names.length,
    epifai: names.filter(n => n.origin === "epifai").length,
    excel: names.filter(n => n.origin === "excel" && n.status !== "broken").length,
    broken: names.filter(n => n.status === "broken").length,
  };

  const filteredNames = names.filter(n => {
    if (filter === "epifai" && n.origin !== "epifai") return false;
    if (filter === "excel" && (n.origin !== "excel" || n.status === "broken")) return false;
    if (filter === "broken" && n.status !== "broken") return false;
    if (search) {
      const q = search.toLowerCase();
      return (
        n.name.toLowerCase().includes(q) || 
        n.comment.toLowerCase().includes(q) ||
        n.formula.toLowerCase().includes(q)
      );
    }
    return true;
  });

  const getStatusBadge = (status: string) => {
    switch (status) {
      case "valid": return <Badge variant="success" className="gap-1"><CheckCircle2 className="w-3 h-3" /> Valid</Badge>;
      case "broken": return <Badge variant="error" className="gap-1"><AlertCircle className="w-3 h-3" /> Broken</Badge>;
      default: return <Badge variant="secondary" className="gap-1"><HelpCircle className="w-3 h-3" /> {status}</Badge>;
    }
  };

  const getRangeType = (n: ExcelName): "fixed" | "dynamic" | "hybrid" => {
    const u = n.formula.toUpperCase();
    const hasOffset = u.includes("OFFSET(") || u.includes("INDIRECT(") || u.includes("INDEX(");
    if (!hasOffset) return "fixed";
    if (n.fixedRef && n.dynamicRef) return "hybrid";
    return "dynamic";
  };

  return (
    <div className="flex flex-col h-full bg-background">
      {/* Search & Filters */}
      <div className="p-3 border-b space-y-3 bg-muted/20">
        <div className="flex gap-1 p-1 bg-muted rounded-lg">
          {(["epifai", "excel", "all", "broken"] as const).map(key => {
            const labels: Record<string, string> = { all: "ALL", epifai: "EPIFAI", excel: "EXCEL", broken: "BROKEN" };
            return (
              <button
                key={key}
                onClick={() => setFilter(key)}
                data-testid={`filter-${key}`}
                className={cn(
                  "flex-1 py-1 px-1.5 rounded-md text-xs font-medium transition-all",
                  filter === key 
                    ? "bg-background text-foreground shadow-sm ring-1 ring-black/5" 
                    : "text-muted-foreground hover:bg-background/50"
                )}
              >
                <span className="block text-sm font-bold">
                  {key === 'all' ? stats.total : stats[key]}
                </span>
                <span className="uppercase text-[10px] opacity-70">{labels[key]}</span>
              </button>
            );
          })}
        </div>

        <div className="relative">
          <Search className="absolute left-3 top-2.5 h-4 w-4 text-muted-foreground" />
          <Input 
            placeholder="Search names..." 
            value={search}
            onChange={(e) => setSearch(e.target.value)}
            className="pl-9 h-9 bg-background border-border/60 focus:border-primary"
          />
        </div>
      </div>

      {/* List */}
      <div className="flex-1 min-h-0 overflow-y-auto">
        <div className="flex flex-col">
          {filteredNames.length === 0 && (
            <div className="p-8 text-center text-muted-foreground text-sm">
              No names found matching your criteria.
            </div>
          )}
          
          {filteredNames.map((n) => {
            const isSelected = selectedId === n.name;
            const rangeType = getRangeType(n);
            const dynamic = rangeType !== "fixed";
            
            return (
              <div
                key={n.name + n.scope}
                onClick={() => {
                  const newId = isSelected ? null : n.name;
                  setSelectedId(newId);
                  if (newId && n.status === "valid") {
                    selectNameRange({ name: n.name, scope: n.scope }).catch(() => {});
                  }
                }}
                className={cn(
                  "border-b border-border/40 transition-colors cursor-pointer group",
                  isSelected ? "bg-accent/30 border-l-4 border-l-primary" : "hover-elevate border-l-4 border-l-transparent"
                )}
              >
                <div className="p-3">
                  <div className="flex items-center justify-between gap-1 mb-1 flex-wrap">
                    <div className="flex items-center gap-1.5 truncate pr-2">
                      <span className="font-semibold text-sm truncate text-foreground/90">
                        {n.name}
                      </span>
                      {n.origin === "epifai" ? (
                        <span className="inline-flex items-center text-[9px] font-bold px-1.5 py-0.5 rounded bg-primary/15 text-primary border border-primary/20 shrink-0" data-testid={`origin-epifai-${n.name}`}>
                          Epifai
                        </span>
                      ) : (
                        <>
                          <span className="inline-flex items-center text-[9px] font-bold px-1.5 py-0.5 rounded bg-muted text-muted-foreground border border-border/50 shrink-0" data-testid={`origin-excel-${n.name}`}>
                            Excel
                          </span>
                          {onClaim && (filter === "excel" || filter === "all") && (
                            <button
                              onClick={(e) => { e.stopPropagation(); onClaim(n); }}
                              className="inline-flex items-center justify-center px-1.5 py-0.5 rounded text-[9px] font-bold bg-primary/10 text-primary border border-primary/20 shrink-0 transition-colors hover:bg-primary/20"
                              title="Add to Epifai"
                              data-testid={`button-claim-${n.name}`}
                            >
                              Add
                            </button>
                          )}
                        </>
                      )}
                    </div>
                    <div className="flex items-center gap-1.5 shrink-0">
                      <Badge 
                        variant={rangeType === "fixed" ? "secondary" : "info"} 
                        className="text-[10px] h-5 px-1.5"
                      >
                        {rangeType === "hybrid" ? (
                          <><Columns className="w-2.5 h-2.5 mr-1"/> Hybrid</>
                        ) : rangeType === "dynamic" ? (
                          <><Maximize className="w-2.5 h-2.5 mr-1"/> Dynamic</>
                        ) : (
                          <><LayoutGrid className="w-2.5 h-2.5 mr-1"/> Fixed</>
                        )}
                      </Badge>
                      {getStatusBadge(n.status)}
                    </div>
                  </div>
                  
                  <div className="text-xs font-mono text-muted-foreground truncate mb-1 bg-muted/30 p-1 rounded">
                    {dynamic ? n.formula.replace(/^=/, "") : (n.address || n.formula.replace(/^=/, ""))}
                  </div>
                  {dynamic && n.address && (
                    <div className="text-[10px] text-muted-foreground mb-1">
                      Currently resolves to: <span className="font-mono font-medium text-foreground/70">{n.address}</span>
                    </div>
                  )}

                  {n.scope !== "Workbook" && (
                    <span className="inline-block text-[10px] text-purple-600 dark:text-purple-400 bg-purple-50 dark:bg-purple-900/30 px-1.5 py-0.5 rounded border border-purple-100 dark:border-purple-800 mb-1">
                      Scope: {n.scope}
                    </span>
                  )}

                  {isSelected && (
                    <div className="pt-3 mt-2 border-t border-border/50">
                      {n.comment && (
                        <div className="text-xs text-muted-foreground mb-3 italic">
                          "{n.comment}"
                        </div>
                      )}
                      
                      <div className="flex items-center gap-2 mt-2">
                        <Button 
                          size="sm" 
                          variant="default"
                          className="text-xs"
                          onClick={(e) => { e.stopPropagation(); onGoTo(n); }}
                          data-testid={`button-goto-${n.name}`}
                        >
                          <ExternalLink className="w-3 h-3 mr-1.5" /> Go To
                        </Button>
                        <Button 
                          size="sm" 
                          variant="outline"
                          className="text-xs"
                          onClick={(e) => { e.stopPropagation(); onEdit(n); }}
                          data-testid={`button-edit-${n.name}`}
                        >
                          <Edit className="w-3 h-3 mr-1.5" /> Edit
                        </Button>
                        <Button 
                          size="sm" 
                          variant="outline"
                          className="text-xs"
                          onClick={(e) => { e.stopPropagation(); onExport?.(n); }}
                          disabled={isExporting || n.status === "broken"}
                          data-testid={`button-export-${n.name}`}
                        >
                          {isExporting ? <Loader2 className="w-3 h-3 mr-1.5 animate-spin" /> : <FileSpreadsheet className="w-3 h-3 mr-1.5" />} Export
                        </Button>
                        <Button 
                          size="sm" 
                          variant={pendingDeleteName === n.name ? "destructive" : "ghost"}
                          className={cn("text-xs", pendingDeleteName !== n.name && "text-destructive")}
                          onClick={(e) => { e.stopPropagation(); onDelete(n); }}
                          data-testid={`button-delete-${n.name}`}
                        >
                          <Trash2 className="w-3.5 h-3.5" />
                          {pendingDeleteName === n.name && <span className="ml-1">Confirm?</span>}
                        </Button>
                      </div>
                    </div>
                  )}
                </div>
              </div>
            );
          })}
        </div>
      </div>

      {/* Footer */}
      <div className="p-3 border-t bg-muted/10 space-y-2">
        <div className="flex items-center justify-between gap-2 flex-wrap">
          <span className="text-xs text-muted-foreground font-medium">
            {filteredNames.length} names
          </span>
          <Button 
            variant="ghost" 
            className="text-primary h-auto p-0 text-xs font-semibold"
            onClick={onCreate}
            data-testid="button-new-named-range"
          >
            + New Named Range
          </Button>
        </div>
        {stats.broken > 0 && onDeleteBroken && filter === "broken" && (
          <div className="flex items-center justify-between gap-2">
            {confirmDeleteBroken ? (
              <>
                <span className="text-xs text-destructive font-medium">
                  Delete {stats.broken} broken names?
                </span>
                <div className="flex items-center gap-1">
                  <Button
                    size="sm"
                    variant="destructive"
                    className="text-xs"
                    onClick={() => { onDeleteBroken(); setConfirmDeleteBroken(false); }}
                    disabled={isDeletingBroken}
                    data-testid="button-confirm-delete-broken"
                  >
                    {isDeletingBroken ? <Loader2 className="w-3 h-3 animate-spin mr-1" /> : <Trash2 className="w-3 h-3 mr-1" />}
                    {isDeletingBroken ? "Deleting..." : "Yes, delete all"}
                  </Button>
                  <Button
                    size="sm"
                    variant="outline"
                    className="text-xs"
                    onClick={() => setConfirmDeleteBroken(false)}
                    disabled={isDeletingBroken}
                    data-testid="button-cancel-delete-broken"
                  >
                    Cancel
                  </Button>
                </div>
              </>
            ) : (
              <Button
                size="sm"
                variant="ghost"
                className="text-xs text-destructive w-full"
                onClick={() => setConfirmDeleteBroken(true)}
                data-testid="button-delete-broken"
              >
                <Trash2 className="w-3 h-3 mr-1" /> Delete All Broken ({stats.broken})
              </Button>
            )}
          </div>
        )}
      </div>
    </div>
  );
}
