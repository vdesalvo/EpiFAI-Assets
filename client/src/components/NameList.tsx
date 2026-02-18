import { useState } from "react";
import { ExcelName } from "@/lib/excel-names";
import { Badge } from "@/components/ui/badge";
import { Button } from "@/components/ui/button";
import { ScrollArea } from "@/components/ui/scroll-area";
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
  Maximize
} from "lucide-react";
import { cn } from "@/lib/utils";

interface NameListProps {
  names: ExcelName[];
  onEdit: (name: ExcelName) => void;
  onDelete: (name: ExcelName) => void;
  onGoTo: (name: ExcelName) => void;
  onCreate: () => void;
}

export function NameList({ names, onEdit, onDelete, onGoTo, onCreate }: NameListProps) {
  const [search, setSearch] = useState("");
  const [filter, setFilter] = useState<"all" | "valid" | "broken" | "unused">("all");
  const [selectedId, setSelectedId] = useState<string | null>(null);

  const stats = {
    total: names.length,
    valid: names.filter(n => n.status === "valid").length,
    broken: names.filter(n => n.status === "broken").length,
    unused: names.filter(n => false).length // "unused" logic wasn't fully implemented in service, placeholder
  };

  const filteredNames = names.filter(n => {
    if (filter !== "all" && n.status !== filter) return false;
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

  const isDynamic = (formula: string) => {
    const u = formula.toUpperCase();
    return u.includes("OFFSET(") || u.includes("INDIRECT(") || u.includes("INDEX(");
  };

  return (
    <div className="flex flex-col h-full bg-background">
      {/* Search & Filters */}
      <div className="p-3 border-b space-y-3 bg-muted/20">
        <div className="flex gap-2 p-1 bg-muted rounded-lg">
          {(["all", "valid", "broken"] as const).map(key => (
            <button
              key={key}
              onClick={() => setFilter(key)}
              className={cn(
                "flex-1 py-1 px-2 rounded-md text-xs font-medium transition-all",
                filter === key 
                  ? "bg-background text-foreground shadow-sm ring-1 ring-black/5" 
                  : "text-muted-foreground hover:bg-background/50"
              )}
            >
              <span className="block text-sm font-bold">
                {key === 'all' ? stats.total : stats[key]}
              </span>
              <span className="uppercase text-[10px] opacity-70">{key}</span>
            </button>
          ))}
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
      <ScrollArea className="flex-1">
        <div className="flex flex-col">
          {filteredNames.length === 0 && (
            <div className="p-8 text-center text-muted-foreground text-sm">
              No names found matching your criteria.
            </div>
          )}
          
          {filteredNames.map((n) => {
            const isSelected = selectedId === n.name;
            const dynamic = isDynamic(n.formula);
            
            return (
              <div
                key={n.name + n.scope}
                onClick={() => setSelectedId(isSelected ? null : n.name)}
                className={cn(
                  "border-b border-border/40 transition-colors cursor-pointer group",
                  isSelected ? "bg-accent/30 border-l-4 border-l-primary" : "hover-elevate border-l-4 border-l-transparent"
                )}
              >
                <div className="p-3">
                  <div className="flex items-center justify-between gap-1 mb-1 flex-wrap">
                    <div className="font-semibold text-sm truncate pr-2 text-foreground/90">
                      {n.name}
                    </div>
                    <div className="flex items-center gap-1.5 shrink-0">
                      <Badge variant={dynamic ? "info" : "secondary"} className="text-[10px] h-5 px-1.5">
                        {dynamic ? <Maximize className="w-2.5 h-2.5 mr-1"/> : <LayoutGrid className="w-2.5 h-2.5 mr-1"/>}
                        {dynamic ? "Dynamic" : "Fixed"}
                      </Badge>
                      {getStatusBadge(n.status)}
                    </div>
                  </div>
                  
                  <div className="text-xs font-mono text-muted-foreground truncate mb-1 bg-muted/30 p-1 rounded">
                    {n.formula.replace(/^=/, "")}
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
                          className="text-xs flex-1"
                          onClick={(e) => { e.stopPropagation(); onGoTo(n); }}
                          data-testid={`button-goto-${n.name}`}
                        >
                          <ExternalLink className="w-3 h-3 mr-1.5" /> Go To
                        </Button>
                        <Button 
                          size="sm" 
                          variant="outline"
                          className="text-xs flex-1"
                          onClick={(e) => { e.stopPropagation(); onEdit(n); }}
                          data-testid={`button-edit-${n.name}`}
                        >
                          <Edit className="w-3 h-3 mr-1.5" /> Edit
                        </Button>
                        <Button 
                          size="sm" 
                          variant="ghost"
                          className="text-xs text-destructive"
                          onClick={(e) => { e.stopPropagation(); onDelete(n); }}
                          data-testid={`button-delete-${n.name}`}
                        >
                          <Trash2 className="w-3.5 h-3.5" />
                        </Button>
                      </div>
                    </div>
                  )}
                </div>
              </div>
            );
          })}
        </div>
      </ScrollArea>

      {/* Footer */}
      <div className="p-3 border-t bg-muted/10 flex items-center justify-between">
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
    </div>
  );
}
