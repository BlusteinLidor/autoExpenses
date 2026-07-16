"use client";

import { Card, CardContent, CardDescription, CardHeader, CardTitle } from "@/components/ui/card";
import { CategoryTimelinePoint, TotalsTimelinePoint } from "@/lib/api/types";
import { money } from "@/lib/format";
import { CartesianGrid, Legend, Line, LineChart, ResponsiveContainer, Tooltip, XAxis, YAxis } from "recharts";

type TotalsTimelineChartProps = {
  view?: "totals";
  points: TotalsTimelinePoint[];
  title?: string;
  description?: string;
  loading?: boolean;
};

type CategoryTimelineChartProps = {
  view: "category";
  points: CategoryTimelinePoint[];
  category: string;
  stroke?: string;
  title?: string;
  description?: string;
  loading?: boolean;
};

type TimelineChartProps = TotalsTimelineChartProps | CategoryTimelineChartProps;

export function TimelineChart(props: TimelineChartProps) {
  const {
    title = props.view === "category"
      ? `Category Timeline: ${props.category}`
      : "Income / Spending / Investment Timeline",
    description = props.view === "category"
      ? "Track how this category changes over time."
      : "Track how totals change over time.",
    loading = false,
    points,
  } = props;

  return (
    <Card>
      <CardHeader>
        <CardTitle>{title}</CardTitle>
        <CardDescription>{description}</CardDescription>
      </CardHeader>
      <CardContent>
        {loading ? (
          <div className="space-y-3">
            <div className="h-[320px] w-full rounded-md bg-muted/70 animate-pulse" />
            <div className="grid grid-cols-3 gap-2">
              <div className="h-4 rounded bg-muted/70 animate-pulse" />
              <div className="h-4 rounded bg-muted/70 animate-pulse" />
              <div className="h-4 rounded bg-muted/70 animate-pulse" />
            </div>
          </div>
        ) : points.length === 0 ? (
          <p className="text-sm text-muted-foreground">No monthly output data available yet for this range.</p>
        ) : (
          <div className="h-[320px] w-full">
            <ResponsiveContainer width="100%" height="100%">
              {props.view === "category" ? (
                <LineChart data={props.points} margin={{ top: 10, right: 18, left: 6, bottom: 8 }}>
                  <CartesianGrid strokeDasharray="3 3" />
                  <XAxis dataKey="label" tick={{ fontSize: 12 }} />
                  <YAxis tickFormatter={(value) => money(Number(value))} tick={{ fontSize: 12 }} />
                  <Tooltip formatter={(value) => money(Number(value))} />
                  <Legend />
                  <Line
                    type="monotone"
                    dataKey="amount"
                    name={props.category}
                    stroke={props.stroke ?? "#0ea5e9"}
                    strokeWidth={2}
                    dot={false}
                  />
                </LineChart>
              ) : (
                <LineChart data={props.points} margin={{ top: 10, right: 18, left: 6, bottom: 8 }}>
                  <CartesianGrid strokeDasharray="3 3" />
                  <XAxis dataKey="label" tick={{ fontSize: 12 }} />
                  <YAxis tickFormatter={(value) => money(Number(value))} tick={{ fontSize: 12 }} />
                  <Tooltip formatter={(value) => money(Number(value))} />
                  <Legend />
                  <Line type="monotone" dataKey="spending" stroke="#f43f5e" strokeWidth={2} dot={false} />
                  <Line type="monotone" dataKey="income" stroke="#22c55e" strokeWidth={2} dot={false} />
                  <Line type="monotone" dataKey="investments" stroke="#6366f1" strokeWidth={2} dot={false} />
                </LineChart>
              )}
            </ResponsiveContainer>
          </div>
        )}
      </CardContent>
    </Card>
  );
}
