import { Badge } from "@/components/ui/badge";
import { Card, CardContent, CardHeader, CardTitle } from "@/components/ui/card";

type ProgressPanelProps = {
  status: string;
  progressStep: string;
  lastFilled: string;
};

export function ProgressPanel({ status, progressStep, lastFilled }: ProgressPanelProps) {
  return (
    <Card>
      <CardHeader>
        <CardTitle>Pipeline Status</CardTitle>
      </CardHeader>
      <CardContent className="space-y-3 text-sm">
        <div className="flex flex-wrap items-center gap-2">
          <span className="text-muted-foreground">Status:</span>
          <Badge variant="secondary">{status || "Idle"}</Badge>
        </div>
        <p className="text-muted-foreground">{progressStep || "No active run."}</p>
        <p className="text-muted-foreground">{lastFilled || "No previous filled month."}</p>
      </CardContent>
    </Card>
  );
}
