"use client";

import { Card, CardContent, CardDescription, CardHeader, CardTitle } from "@/components/ui/card";
import { TotalsTimelinePoint } from "@/lib/api/types";
import { money } from "@/lib/format";
import { CartesianGrid, Legend, Line, LineChart, ResponsiveContainer, Tooltip, XAxis, YAxis } from "recharts";

type TimelineChartProps = {
  points: TotalsTimelinePoint[];
  title?: string;
  description?: string;
};

export function TimelineChart({
  points,
  title = "Income / Spending / Investment Timeline",
  description = "Track how totals change over time.",
}: TimelineChartProps) {
  return (
    <Card>
      <CardHeader>
        <CardTitle>{title}</CardTitle>
        <CardDescription>{description}</CardDescription>
      </CardHeader>
      <CardContent>
        {points.length === 0 ? (
          <p className="text-sm text-muted-foreground">No monthly output data available yet for this range.</p>
        ) : (
          <div className="h-[320px] w-full">
            <ResponsiveContainer width="100%" height="100%">
              <LineChart data={points} margin={{ top: 10, right: 18, left: 6, bottom: 8 }}>
                <CartesianGrid strokeDasharray="3 3" />
                <XAxis dataKey="label" tick={{ fontSize: 12 }} />
                <YAxis tickFormatter={(value) => money(Number(value))} tick={{ fontSize: 12 }} />
                <Tooltip formatter={(value) => money(Number(value))} />
                <Legend />
                <Line type="monotone" dataKey="spending" stroke="#f43f5e" strokeWidth={2} dot={false} />
                <Line type="monotone" dataKey="income" stroke="#22c55e" strokeWidth={2} dot={false} />
                <Line type="monotone" dataKey="investments" stroke="#6366f1" strokeWidth={2} dot={false} />
              </LineChart>
            </ResponsiveContainer>
          </div>
        )}
      </CardContent>
    </Card>
  );
}
