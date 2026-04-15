"use client";

import { Button } from "@/components/ui/button";
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from "@/components/ui/card";
import { Label } from "@/components/ui/label";
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from "@/components/ui/select";

type RunFormProps = {
  year: string;
  month: number;
  years: string[];
  busy: boolean;
  onYearChange: (year: string) => void;
  onMonthChange: (month: number) => void;
  onRun: () => void;
};

export function RunForm(props: RunFormProps) {
  const {
    year,
    month,
    years,
    busy,
    onYearChange,
    onMonthChange,
    onRun,
  } = props;

  return (
    <Card>
      <CardHeader>
        <CardTitle>Monthly Run</CardTitle>
        <CardDescription>Prepare, review, and apply expense categories.</CardDescription>
      </CardHeader>
      <CardContent className="space-y-4">
        <div className="grid gap-3 sm:grid-cols-2">
          <div className="space-y-2">
            <Label>Run Year</Label>
            <Select value={year} onValueChange={(value) => onYearChange(value ?? year)}>
              <SelectTrigger>
                <SelectValue placeholder="Year" />
              </SelectTrigger>
              <SelectContent>
                {years.map((itemYear) => (
                  <SelectItem key={itemYear} value={itemYear}>
                    {itemYear}
                  </SelectItem>
                ))}
              </SelectContent>
            </Select>
          </div>
          <div className="space-y-2">
            <Label>Run Month</Label>
            <Select value={String(month)} onValueChange={(value) => onMonthChange(Number(value ?? month))}>
              <SelectTrigger>
                <SelectValue placeholder="Month" />
              </SelectTrigger>
              <SelectContent>
                {Array.from({ length: 12 }).map((_, idx) => (
                  <SelectItem key={idx + 1} value={String(idx + 1)}>
                    {idx + 1}
                  </SelectItem>
                ))}
              </SelectContent>
            </Select>
          </div>
        </div>

        <Button className="w-full" onClick={onRun} disabled={busy}>
          {busy ? "Running..." : "Run Month"}
        </Button>
      </CardContent>
    </Card>
  );
}
