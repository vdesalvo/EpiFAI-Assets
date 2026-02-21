import { useState, useEffect } from "react";
import { Tabs, TabsContent, TabsList, TabsTrigger } from "@/components/ui/tabs";
import { useNames, useCharts, useAddName, useUpdateName, useDeleteName, useGoToName, useClaimName, useDeleteBrokenNames, useRenameChart, useGoToChart, useCreateNameFromChart } from "@/hooks/use-excel";
import { NameList } from "@/components/NameList";
import { NameEditor } from "@/components/NameEditor";
import { RangePicker } from "@/components/RangePicker";
import { FullPageLoader, LoadingSpinner } from "@/components/ui/loading-spinner";
import { useToast } from "@/hooks/use-toast";
import { ExcelName, getSelectionData } from "@/lib/excel-names";
import { RefreshCw, Table2, BarChart3, Info, Download, Plus, Check, Loader2 } from "lucide-react";
import { Button } from "@/components/ui/button";
import { motion, AnimatePresence } from "framer-motion";

function BuildTimestamp() {
  const [info, setInfo] = useState<{ buildTime: string; version: string } | null>(null);
  useEffect(() => {
    fetch("/api/build-info").then(r => r.json()).then(setInfo).catch(() => {});
  }, []);
  if (!info) return null;
  const d = new Date(info.buildTime);
  const formatted = d.toLocaleString("en-US", { month: "short", day: "numeric", year: "numeric", hour: "numeric", minute: "2-digit", hour12: true });
  return (
    <div className="px-4 pb-1 text-[9px] text-muted-foreground/60 text-right shrink-0" data-testid="text-build-time">
      v{info.version} &middot; Published: {formatted}
    </div>
  );
}

// Simple Chart List Item Component (Internal)
function ChartListItem({ chart, onRename, onGoTo, onCreateName }: { chart: any, onRename: (id: string, name: string) => void, onGoTo: (c: any) => void, onCreateName: (c: any) => Promise<void> }) {
  const [isEditing, setIsEditing] = useState(false);
  const [tempName, setTempName] = useState(chart.name);
  const [creating, setCreating] = useState<"idle" | "loading" | "done">("idle");

  const handleSave = () => {
    if (tempName !== chart.name) {
      onRename(chart.id, tempName);
    }
    setIsEditing(false);
  };

  const handleCreate = async () => {
    setCreating("loading");
    try {
      await onCreateName(chart);
      setCreating("done");
      setTimeout(() => setCreating("idle"), 2000);
    } catch {
      setCreating("idle");
    }
  };

  return (
    <div className="border-b p-4 flex items-center justify-between hover:bg-muted/20 transition-colors">
      <div className="flex items-center gap-3 overflow-hidden">
        <div className="w-10 h-10 rounded bg-blue-50 text-blue-600 flex items-center justify-center shrink-0 border border-blue-100">
          <BarChart3 className="w-5 h-5" />
        </div>
        <div className="min-w-0">
          {isEditing ? (
            <input 
              className="text-sm font-semibold border rounded px-1 w-full"
              value={tempName}
              onChange={e => setTempName(e.target.value)}
              onBlur={handleSave}
              onKeyDown={e => e.key === 'Enter' && handleSave()}
              autoFocus
              data-testid={`input-chart-rename-${chart.id}`}
            />
          ) : (
            <div className="text-sm font-semibold text-foreground truncate cursor-text" onClick={() => setIsEditing(true)} title="Click to rename">
              {chart.title || "(No Title)"}
            </div>
          )}
          <div className="text-xs text-muted-foreground truncate">{chart.name} • {chart.sheet}</div>
        </div>
      </div>
      <div className="flex items-center gap-1 shrink-0 ml-2">
        {chart.title && chart.title !== "(No Title)" && (
          <Button
            size="icon"
            variant="ghost"
            onClick={handleCreate}
            disabled={creating === "loading"}
            title="Create named range from chart title"
            data-testid={`button-create-name-${chart.id}`}
          >
            {creating === "loading" ? (
              <Loader2 className="w-4 h-4 animate-spin" />
            ) : creating === "done" ? (
              <Check className="w-4 h-4 text-green-600" />
            ) : (
              <Plus className="w-4 h-4" />
            )}
          </Button>
        )}
        <Button size="sm" variant="ghost" onClick={() => onGoTo(chart)} data-testid={`button-goto-chart-${chart.id}`}>
          Go To
        </Button>
      </div>
    </div>
  );
}

export default function Home() {
  const [init, setInit] = useState(false);
  const { toast } = useToast();
  
  // Data Hooks
  const { data: names = [], isLoading: loadingNames, refetch: refetchNames, error: namesError } = useNames();
  const { data: charts = [], isLoading: loadingCharts, refetch: refetchCharts } = useCharts();
  
  // Mutations
  const addName = useAddName();
  const updateName = useUpdateName();
  const deleteName = useDeleteName();
  const goToName = useGoToName();
  const claimName = useClaimName();
  const deleteBroken = useDeleteBrokenNames();
  const renameChart = useRenameChart();
  const goToChart = useGoToChart();
  const createNameFromChart = useCreateNameFromChart();

  // UI State
  const [view, setView] = useState<"list" | "edit" | "visual-picker">("list");
  const [editTarget, setEditTarget] = useState<ExcelName | undefined>(undefined);
  const [activeTab, setActiveTab] = useState("names");

  useEffect(() => {
    setInit(true);
  }, []);

  const handleRefresh = () => {
    refetchNames();
    refetchCharts();
    toast({ description: "Synced with Excel workbook" });
  };

  // --- Handlers ---

  const handleCreateName = () => {
    setEditTarget(undefined);
    setView("edit");
  };

  const handleVisualPicker = () => {
    setView("visual-picker");
  };

  const handleEditName = (name: ExcelName) => {
    setEditTarget(name);
    setView("edit");
  };

  const handleCreateNameFromChart = async (chart: any) => {
    try {
      const createdName = await createNameFromChart.mutateAsync({ sheet: chart.sheet, chartName: chart.name, title: chart.title });
      toast({ title: "Created", description: `Named range "${createdName}" created from chart data` });
      await new Promise(r => setTimeout(r, 300));
      await refetchNames();
      setActiveTab("names");
    } catch (e: any) {
      toast({ variant: "destructive", title: "Error", description: e.message || "Could not create name from chart" });
    }
  };

  const handleSaveName = async (data: { name: string; refersTo: string; comment: string; newName?: string; skipRows?: number; skipCols?: number; fixedRef?: string; dynamicRef?: string; lastColOnly?: boolean; lastRowOnly?: boolean }) => {
    try {
      if (editTarget) {
        await updateName.mutateAsync({ name: editTarget.name, updates: { ...data, skipRows: data.skipRows, skipCols: data.skipCols, fixedRef: data.fixedRef, dynamicRef: data.dynamicRef, lastColOnly: data.lastColOnly, lastRowOnly: data.lastRowOnly } });
        toast({ title: "Updated", description: `Updated range "${data.newName || data.name}"` });
      } else {
        await addName.mutateAsync({ name: data.name, formula: data.refersTo, comment: data.comment, skipRows: data.skipRows || 0, skipCols: data.skipCols || 0, fixedRef: data.fixedRef || "", dynamicRef: data.dynamicRef || "", lastColOnly: data.lastColOnly || false, lastRowOnly: data.lastRowOnly || false });
        toast({ title: "Created", description: `Created range "${data.name}"` });
      }
      setView("list");
    } catch (e: any) {
      toast({ variant: "destructive", title: "Error", description: e.message });
    }
  };

  const [pendingDelete, setPendingDelete] = useState<ExcelName | null>(null);

  const handleDeleteName = async (name: ExcelName) => {
    if (!pendingDelete || pendingDelete.name !== name.name) {
      setPendingDelete(name);
      return;
    }
    try {
      await deleteName.mutateAsync({ name: name.name, scope: name.scope });
      toast({ description: `Deleted "${name.name}"` });
      setPendingDelete(null);
    } catch (e: any) {
      toast({ variant: "destructive", title: "Error", description: e.message });
      setPendingDelete(null);
    }
  };

  const handleGoToName = (name: ExcelName) => {
    goToName.mutate({ name: name.name, scope: name.scope }, {
      onError: (err: any) => {
        toast({ variant: "destructive", title: "Go To failed", description: err?.message || "Could not navigate to this name" });
      }
    });
  };

  const handleDeleteBroken = () => {
    deleteBroken.mutate(undefined, {
      onSuccess: (count) => {
        toast({ description: `Deleted ${count} broken name${count !== 1 ? "s" : ""}` });
      },
      onError: (err: any) => {
        toast({ variant: "destructive", title: "Error", description: err?.message || "Could not delete broken names" });
      }
    });
  };

  const handleClaimName = (name: ExcelName) => {
    claimName.mutate({ name: name.name, scope: name.scope }, {
      onSuccess: () => {
        toast({ description: `"${name.name}" added to Epifai` });
      },
      onError: (err: any) => {
        toast({ variant: "destructive", title: "Error", description: err?.message || "Could not add to Epifai" });
      }
    });
  };

  const handleRenameChart = async (id: string, newName: string) => {
    const chart = charts.find(c => c.id === id);
    if (!chart) return;
    try {
      await renameChart.mutateAsync({ sheet: chart.sheet, oldName: chart.name, newName });
      toast({ description: "Chart renamed" });
    } catch (e: any) {
      toast({ variant: "destructive", title: "Error", description: e.message });
    }
  };

  if (!init || (loadingNames && !names.length)) return <FullPageLoader />;

  return (
    <div className="flex flex-col h-screen w-full bg-background text-foreground overflow-hidden">
      {/* Header */}
      <div className="flex items-center justify-between p-4 border-b bg-gradient-to-r from-background to-muted/20 shrink-0">
        <div className="flex items-center gap-2">
          <div className="w-8 h-8 rounded-lg bg-primary flex items-center justify-center text-primary-foreground shadow-lg shadow-primary/20">
            <Table2 className="w-5 h-5" />
          </div>
          <div>
            <h1 className="font-bold text-sm leading-none">Epifai</h1>
            <p className="text-[10px] text-muted-foreground font-medium uppercase tracking-wider mt-0.5">Manager</p>
          </div>
        </div>
        <Button variant="outline" size="icon" onClick={handleRefresh} title="Sync from Excel">
          <RefreshCw className={`w-3.5 h-3.5 ${loadingNames ? 'animate-spin' : ''}`} />
        </Button>
      </div>
      <BuildTimestamp />

      {namesError ? (
        <div className="flex flex-col items-center justify-center flex-1 p-6 gap-4">
          <div className="p-4 border border-destructive/20 bg-destructive/10 rounded-lg text-sm text-destructive flex items-start gap-2 w-full">
            <Info className="w-4 h-4 shrink-0 mt-0.5" />
            <p>{(namesError as Error).message}. Make sure you are running this inside Excel.</p>
          </div>
          <div className="text-center text-sm text-muted-foreground">
            <p className="mb-3">To get started, download the manifest and upload it into Excel.</p>
            <Button
              data-testid="button-download-manifest"
              variant="default"
              onClick={() => {
                const link = document.createElement("a");
                link.href = "/manifest.xml";
                link.download = "manifest.xml";
                document.body.appendChild(link);
                link.click();
                document.body.removeChild(link);
              }}
            >
              <Download className="w-4 h-4 mr-2" />
              Download Manifest
            </Button>
          </div>
        </div>
      ) : (
        <Tabs value={activeTab} onValueChange={setActiveTab} className="flex-1 flex flex-col min-h-0">
          <TabsList className="grid w-full grid-cols-2 p-1 bg-muted/30 border-b rounded-none h-11 shrink-0">
            <TabsTrigger value="names" className="data-[state=active]:bg-background data-[state=active]:shadow-sm rounded-md transition-all">
              <Table2 className="w-3.5 h-3.5 mr-2" /> Names ({names.length})
            </TabsTrigger>
            <TabsTrigger value="charts" className="data-[state=active]:bg-background data-[state=active]:shadow-sm rounded-md transition-all">
              <BarChart3 className="w-3.5 h-3.5 mr-2" /> Charts ({charts.length})
            </TabsTrigger>
          </TabsList>

          <TabsContent value="names" className="flex-1 min-h-0 m-0 relative">
            <AnimatePresence mode="wait">
              {view === "list" ? (
                <motion.div 
                  key="list" 
                  initial={{ opacity: 0, x: -20 }} 
                  animate={{ opacity: 1, x: 0 }} 
                  exit={{ opacity: 0, x: -20 }}
                  className="absolute inset-0"
                >
                  <NameList 
                    names={names} 
                    onCreate={handleCreateName}
                    onVisualPicker={handleVisualPicker}
                    onEdit={handleEditName}
                    onDelete={handleDeleteName}
                    onGoTo={handleGoToName}
                    onClaim={handleClaimName}
                    onDeleteBroken={handleDeleteBroken}
                    isDeletingBroken={deleteBroken.isPending}
                    pendingDeleteName={pendingDelete?.name || null}
                  />
                </motion.div>
              ) : view === "edit" ? (
                <motion.div 
                  key="edit"
                  initial={{ opacity: 0, x: 20 }} 
                  animate={{ opacity: 1, x: 0 }} 
                  exit={{ opacity: 0, x: 20 }}
                  className="absolute inset-0"
                >
                  <NameEditor 
                    key={editTarget?.name || "new"}
                    initialData={editTarget}
                    onSave={handleSaveName} 
                    onCancel={() => setView("list")} 
                  />
                </motion.div>
              ) : (
                <motion.div
                  key="visual-picker"
                  initial={{ opacity: 0, x: 20 }}
                  animate={{ opacity: 1, x: 0 }}
                  exit={{ opacity: 0, x: 20 }}
                  className="absolute inset-0"
                >
                  <RangePicker
                    onSave={handleSaveName}
                    onCancel={() => setView("list")}
                    onPickSelection={getSelectionData}
                  />
                </motion.div>
              )}
            </AnimatePresence>
          </TabsContent>

          <TabsContent value="charts" className="flex-1 min-h-0 m-0 overflow-y-auto">
             {charts.length === 0 ? (
               <div className="flex flex-col items-center justify-center h-full text-muted-foreground text-sm p-8 text-center">
                 <BarChart3 className="w-12 h-12 mb-3 opacity-20" />
                 <p>No charts found in this workbook.</p>
               </div>
             ) : (
               <div className="divide-y">
                 {charts.map(c => (
                   <ChartListItem 
                    key={c.id} 
                    chart={c} 
                    onRename={handleRenameChart}
                    onGoTo={(c) => goToChart.mutate({sheet: c.sheet, name: c.name})}
                    onCreateName={handleCreateNameFromChart}
                  />
                 ))}
               </div>
             )}
          </TabsContent>
        </Tabs>
      )}
    </div>
  );
}
