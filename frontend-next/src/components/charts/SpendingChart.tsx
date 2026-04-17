"use client";

import { Card, CardContent, CardDescription, CardHeader, CardTitle } from "@/components/ui/card";
import { ExpensesSummary } from "@/lib/api/types";
import { money, safeNumber } from "@/lib/format";
import { useCallback, useRef, useState } from "react";
import { Cell, Pie, PieChart, ResponsiveContainer, Tooltip } from "recharts";

const CHART_COLORS = ["#6366f1", "#8b5cf6", "#06b6d4", "#22c55e", "#eab308", "#f97316", "#f43f5e", "#84cc16", "#14b8a6"];

type SpendingChartProps = {
  summary: ExpensesSummary;
  title?: string;
  description?: string;
  totalLabel?: string;
  loading?: boolean;
};

function TruncatedCategoryLabel({ text }: { text: string }) {
  const labelRef = useRef<HTMLSpanElement>(null);
  const [isTruncated, setIsTruncated] = useState(false);

  const updateTruncation = useCallback(() => {
    if (!labelRef.current) {
      return;
    }
    setIsTruncated(labelRef.current.scrollWidth > labelRef.current.clientWidth);
  }, []);

  return (
    <span
      ref={labelRef}
      className="truncate"
      title={isTruncated ? text : undefined}
      onMouseEnter={updateTruncation}
      onFocus={updateTruncation}
    >
      {text}
    </span>
  );
}

export function SpendingChart({
  summary,
  title = "Spending Breakdown",
  description = "Top categories for selected month.",
  totalLabel = "Total Spendings",
  loading = false,
}: SpendingChartProps) {
  const sorted = Object.entries(summary)
    .map(([name, value]) => ({ name, value: safeNumber(value) }))
    .sort((a, b) => b.value - a.value);

  const top = sorted.slice(0, 8);
  const other = sorted.slice(8).reduce((acc, item) => acc + item.value, 0);
  const data = other > 0 ? [...top, { name: "Other", value: other }] : top;
  const total = data.reduce((acc, item) => acc + item.value, 0);

  return (
    <Card>
      <CardHeader>
        <CardTitle>{title}</CardTitle>
        <CardDescription>{description}</CardDescription>
      </CardHeader>
      <CardContent className="space-y-4">
        {loading ? (
          <>
            <div className="h-[280px] rounded-md bg-muted/70 animate-pulse" />
            <div className="space-y-2">
              {Array.from({ length: 5 }).map((_, idx) => (
                <div key={`skeleton-${idx}`} className="flex items-center justify-between">
                  <div className="h-4 w-32 rounded bg-muted/70 animate-pulse" />
                  <div className="h-4 w-24 rounded bg-muted/70 animate-pulse" />
                </div>
              ))}
              <div className="mt-3 flex items-center justify-between border-t pt-3">
                <div className="h-4 w-28 rounded bg-muted/70 animate-pulse" />
                <div className="h-4 w-20 rounded bg-muted/70 animate-pulse" />
              </div>
            </div>
          </>
        ) : data.length === 0 ? (
          <p className="text-sm text-muted-foreground">No summary available yet for this month.</p>
        ) : (
          <>
            <div className="h-[280px]">
              <ResponsiveContainer width="100%" height="100%">
                <PieChart>
                  <Pie data={data} dataKey="value" nameKey="name" outerRadius={110}>
                    {data.map((_, index) => (
                      <Cell key={`cell-${index}`} fill={CHART_COLORS[index % CHART_COLORS.length]} />
                    ))}
                  </Pie>
                  <Tooltip />
                </PieChart>
              </ResponsiveContainer>
            </div>
            <div className="space-y-2">
              {data.map((item, index) => {
                const percent = total > 0 ? (item.value / total) * 100 : 0;
                return (
                  <div key={item.name} className="flex items-center justify-between text-sm">
                    <div className="flex min-w-0 items-center gap-2">
                      <span
                        aria-hidden
                        className="inline-block h-2.5 w-2.5 shrink-0 rounded-full"
                        style={{ backgroundColor: CHART_COLORS[index % CHART_COLORS.length] }}
                      />
                      <TruncatedCategoryLabel text={item.name} />
                    </div>
                    <span className="shrink-0 text-muted-foreground">{money(item.value)} ({percent.toFixed(1)}%)</span>
                  </div>
                );
              })}
              <div className="mt-3 flex items-center justify-between border-t pt-3 text-sm font-semibold">
                <span>{totalLabel}</span>
                <span>{money(total)}</span>
              </div>
            </div>
          </>
        )}
      </CardContent>
    </Card>
  );
}
