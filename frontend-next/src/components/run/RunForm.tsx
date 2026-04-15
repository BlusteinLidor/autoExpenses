"use client";

import { Button } from "@/components/ui/button";
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from "@/components/ui/card";
import { Checkbox } from "@/components/ui/checkbox";
import { Label } from "@/components/ui/label";
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from "@/components/ui/select";

type RunFormProps = {
  year: string;
  month: number;
  years: string[];
  includeLeumi: boolean;
  busy: boolean;
  onYearChange: (year: string) => void;
  onMonthChange: (month: number) => void;
  onIncludeLeumiChange: (checked: boolean) => void;
  onRun: () => void;
};

export function RunForm(props: RunFormProps) {
  const {
    year,
    month,
    years,
    includeLeumi,
    busy,
    onYearChange,
    onMonthChange,
    onIncludeLeumiChange,
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
            <Label>Year</Label>
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
            <Label>Month</Label>
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

        <div className="flex items-center gap-2">
          <Checkbox
            id="includeLeumi"
            checked={includeLeumi}
            onCheckedChange={(checked) => onIncludeLeumiChange(Boolean(checked))}
          />
          <Label htmlFor="includeLeumi">Include Leumi transactions</Label>
        </div>

        <Button className="w-full" onClick={onRun} disabled={busy}>
          {busy ? "Running..." : "Run Month"}
        </Button>
      </CardContent>
    </Card>
  );
}
