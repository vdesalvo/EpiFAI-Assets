import { useState, useEffect } from "react";
import { Tabs, TabsContent, TabsList, TabsTrigger } from "@/components/ui/tabs";
import { useNames, useCharts, useAddName, useUpdateName, useDeleteName, useGoToName, useRenameChart, useGoToChart } from "@/hooks/use-excel";
import { NameList } from "@/components/NameList";
import { NameEditor } from "@/components/NameEditor";
import { FullPageLoader, LoadingSpinner } from "@/components/ui/loading-spinner";
import { useToast } from "@/hooks/use-toast";
import { ExcelName } from "@/lib/excel-names";
import { RefreshCw, Table2, BarChart3, Info, Download } from "lucide-react";
import { Button } from "@/components/ui/button";
import { motion, AnimatePresence } from "framer-motion";

// Simple Chart List Item Component (Internal)
function ChartListItem({ chart, onRename, onGoTo }: { chart: any, onRename: (id: string, name: string) => void, onGoTo: (c: any) => void }) {
  const [isEditing, setIsEditing] = useState(false);
  const [tempName, setTempName] = useState(chart.name);

  const handleSave = () => {
    if (tempName !== chart.name) {
      onRename(chart.id, tempName);
    }
    setIsEditing(false);
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
            />
          ) : (
            <div className="text-sm font-semibold text-foreground truncate cursor-text" onClick={() => setIsEditing(true)} title="Click to rename">
              {chart.name}
            </div>
          )}
          <div className="text-xs text-muted-foreground truncate">{chart.title} • {chart.sheet}</div>
        </div>
      </div>
      <Button size="sm" variant="ghost" className="shrink-0 ml-2" onClick={() => onGoTo(chart)}>
        Go To
      </Button>
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
  const renameChart = useRenameChart();
  const goToChart = useGoToChart();

  // UI State
  const [view, setView] = useState<"list" | "edit">("list");
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

  const handleEditName = (name: ExcelName) => {
    setEditTarget(name);
    setView("edit");
  };

  const handleSaveName = async (data: { name: string; refersTo: string; comment: string; newName?: string }) => {
    try {
      if (editTarget) {
        await updateName.mutateAsync({ name: editTarget.name, updates: data });
        toast({ title: "Updated", description: `Updated range "${data.newName || data.name}"` });
      } else {
        await addName.mutateAsync({ name: data.name, formula: data.refersTo, comment: data.comment });
        toast({ title: "Created", description: `Created range "${data.name}"` });
      }
      setView("list");
    } catch (e: any) {
      toast({ variant: "destructive", title: "Error", description: e.message });
    }
  };

  const handleDeleteName = async (name: ExcelName) => {
    if (!confirm(`Are you sure you want to delete "${name.name}"?`)) return;
    try {
      await deleteName.mutateAsync(name.name);
      toast({ description: `Deleted "${name.name}"` });
    } catch (e: any) {
      toast({ variant: "destructive", title: "Error", description: e.message });
    }
  };

  const handleGoToName = (name: ExcelName) => {
    goToName.mutate(name.name);
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
        <Button variant="outline" size="sm" className="h-8 w-8 p-0" onClick={handleRefresh} title="Sync from Excel">
          <RefreshCw className={`w-3.5 h-3.5 ${loadingNames ? 'animate-spin' : ''}`} />
        </Button>
      </div>

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
                    onEdit={handleEditName}
                    onDelete={handleDeleteName}
                    onGoTo={handleGoToName}
                  />
                </motion.div>
              ) : (
                <motion.div 
                  key="edit"
                  initial={{ opacity: 0, x: 20 }} 
                  animate={{ opacity: 1, x: 0 }} 
                  exit={{ opacity: 0, x: 20 }}
                  className="absolute inset-0"
                >
                  <NameEditor 
                    initialData={editTarget} 
                    onSave={handleSaveName} 
                    onCancel={() => setView("list")} 
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
