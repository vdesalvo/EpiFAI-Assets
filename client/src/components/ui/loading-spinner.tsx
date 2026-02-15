import { cn } from "@/lib/utils";
import { Loader2 } from "lucide-react";

export function LoadingSpinner({ className }: { className?: string }) {
  return (
    <Loader2 className={cn("h-4 w-4 animate-spin text-primary", className)} />
  );
}

export function FullPageLoader() {
  return (
    <div className="flex h-[50vh] flex-col items-center justify-center space-y-4">
      <div className="relative h-10 w-10">
        <div className="absolute inset-0 rounded-full border-2 border-primary/20"></div>
        <div className="absolute inset-0 animate-spin rounded-full border-2 border-primary border-t-transparent"></div>
      </div>
      <p className="text-sm font-medium text-muted-foreground animate-pulse">Syncing with Excel...</p>
    </div>
  );
}
