import { Card, CardContent } from "@/components/ui/card";
import { type LucideIcon } from "lucide-react";

interface KpiCardProps {
  title: string;
  value: string;
  subtitle?: string;
  icon: LucideIcon;
  trend?: number;
  color?: "primary" | "success" | "destructive" | "warning";
}

const colorMap = {
  primary: "bg-primary/10 text-primary",
  success: "bg-accent/10 text-accent",
  destructive: "bg-destructive/10 text-destructive",
  warning: "bg-[hsl(var(--warning))]/10 text-[hsl(var(--warning))]",
};

export function KpiCard({ title, value, subtitle, icon: Icon, trend, color = "primary" }: KpiCardProps) {
  return (
    <Card className="overflow-hidden">
      <CardContent className="p-5">
        <div className="flex items-start justify-between">
          <div className="space-y-1">
            <p className="text-xs font-medium uppercase tracking-wider text-muted-foreground">
              {title}
            </p>
            <p className="text-2xl font-bold text-foreground">{value}</p>
            {subtitle && (
              <p className="text-xs text-muted-foreground">{subtitle}</p>
            )}
          </div>
          <div className={`rounded-lg p-2.5 ${colorMap[color]}`}>
            <Icon className="h-5 w-5" />
          </div>
        </div>
        {trend !== undefined && (
          <div className="mt-3 flex items-center gap-1 text-xs font-medium">
            <span className={trend >= 0 ? "text-accent" : "text-destructive"}>
              {trend >= 0 ? "▲" : "▼"} {Math.abs(trend).toFixed(1)}%
            </span>
            <span className="text-muted-foreground">vs período anterior</span>
          </div>
        )}
      </CardContent>
    </Card>
  );
}
