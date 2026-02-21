import { useQuery, useMutation, useQueryClient } from "@tanstack/react-query";
import * as namesSvc from "@/lib/excel-names";
import * as chartsSvc from "@/lib/excel-charts";

// === NAMES HOOKS ===

export function useNames() {
  return useQuery({
    queryKey: ["excel-names"],
    queryFn: namesSvc.getAllNames,
    refetchOnWindowFocus: false, // Excel state doesn't change just because window focused
  });
}

export function useAddName() {
  const queryClient = useQueryClient();
  return useMutation({
    mutationFn: ({ name, formula, comment, skipRows, skipCols, fixedRef, dynamicRef, lastColOnly, lastRowOnly }: { name: string; formula: string; comment?: string; skipRows?: number; skipCols?: number; fixedRef?: string; dynamicRef?: string; lastColOnly?: boolean; lastRowOnly?: boolean }) => 
      namesSvc.addName(name, formula, comment, "Workbook", skipRows || 0, skipCols || 0, fixedRef || "", dynamicRef || "", lastColOnly || false, lastRowOnly || false),
    onSuccess: () => queryClient.invalidateQueries({ queryKey: ["excel-names"] }),
  });
}

export function useUpdateName() {
  const queryClient = useQueryClient();
  return useMutation({
    mutationFn: ({ name, updates }: { name: string; updates: namesSvc.UpdateNameParams }) => 
      namesSvc.updateName(name, updates),
    onSuccess: () => queryClient.invalidateQueries({ queryKey: ["excel-names"] }),
  });
}

export function useDeleteName() {
  const queryClient = useQueryClient();
  return useMutation({
    mutationFn: namesSvc.deleteName,
    onSuccess: () => queryClient.invalidateQueries({ queryKey: ["excel-names"] }),
  });
}

export function useGoToName() {
  return useMutation({
    mutationFn: namesSvc.goToName,
  });
}

export function useDeleteBrokenNames() {
  const queryClient = useQueryClient();
  return useMutation({
    mutationFn: namesSvc.deleteBrokenNames,
    onSuccess: () => queryClient.invalidateQueries({ queryKey: ["excel-names"] }),
  });
}

export function useClaimName() {
  const queryClient = useQueryClient();
  return useMutation({
    mutationFn: ({ name, scope }: { name: string; scope: string }) =>
      namesSvc.claimAsEpifai(name, scope),
    onSuccess: () => queryClient.invalidateQueries({ queryKey: ["excel-names"] }),
  });
}

export function useSelectionData() {
  const queryClient = useQueryClient();
  return useMutation({
    mutationFn: namesSvc.getSelectionData,
  });
}

export function useExportName() {
  return useMutation({
    mutationFn: ({ name, scope }: { name: string; scope: string }) =>
      namesSvc.exportNameToSheet({ name, scope }),
  });
}

// === CHARTS HOOKS ===

export function useCharts() {
  return useQuery({
    queryKey: ["excel-charts"],
    queryFn: async () => {
      const data = await chartsSvc.getAllCharts();
      return data.map(c => ({
        ...c,
        isDefault: chartsSvc.isDefaultName(c.name)
      }));
    },
    refetchOnWindowFocus: false,
  });
}

export function useChartImage(sheet: string, name: string) {
  return useQuery({
    queryKey: ["excel-chart-img", sheet, name],
    queryFn: () => chartsSvc.getChartImage(sheet, name),
    staleTime: 1000 * 60 * 5, // Cache images for 5 mins
    enabled: !!sheet && !!name,
  });
}

export function useRenameChart() {
  const queryClient = useQueryClient();
  return useMutation({
    mutationFn: ({ sheet, oldName, newName }: { sheet: string; oldName: string; newName: string }) => 
      chartsSvc.renameChart(sheet, oldName, newName),
    onSuccess: () => queryClient.invalidateQueries({ queryKey: ["excel-charts"] }),
  });
}

export function useGoToChart() {
  return useMutation({
    mutationFn: ({ sheet, name }: { sheet: string; name: string }) => 
      chartsSvc.goToChart(sheet, name),
  });
}

export function useCreateNameFromChart() {
  const queryClient = useQueryClient();
  return useMutation({
    mutationFn: ({ sheet, chartName, title }: { sheet: string; chartName: string; title: string }) =>
      chartsSvc.createNameFromChart(sheet, chartName, title),
    onSuccess: () => queryClient.invalidateQueries({ queryKey: ["excel-names"] }),
  });
}
