"use client";

import { AssetsPanel } from "@/components/assets/AssetsPanel";
import { SpendingChart } from "@/components/charts/SpendingChart";
import { TimelineChart } from "@/components/charts/TimelineChart";
import { ReviewTable } from "@/components/review/ReviewTable";
import { ProgressPanel } from "@/components/run/ProgressPanel";
import { RunForm } from "@/components/run/RunForm";
import { Alert, AlertDescription, AlertTitle } from "@/components/ui/alert";
import { Label } from "@/components/ui/label";
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from "@/components/ui/select";
import {
  finalizeRun,
  getAssets,
  getExpensesSummary,
  getIncomeSummary,
  getInvestmentsSummary,
  getTotalsTimeline,
  getRunProgress,
  getState,
  prepareRun,
  saveAssets,
} from "@/lib/api/client";
import { AssetsResponse, FinalizeRunItem, PrepareRunResponse } from "@/lib/api/types";
import { generateRunToken, monthName, safeNumber } from "@/lib/format";
import { useMutation, useQuery, useQueryClient } from "@tanstack/react-query";
import { useMemo, useState } from "react";

function nextMonthFromState(state: Awaited<ReturnType<typeof getState>> | undefined) {
  const lastYear = Number(state?.last_filled_year);
  const lastMonth = Number(state?.last_filled_month);
  if (!Number.isFinite(lastYear) || !Number.isFinite(lastMonth) || lastMonth < 1 || lastMonth > 12) {
    const now = new Date();
    return { year: String(now.getFullYear()), month: now.getMonth() + 1 };
  }
  const nextDate = new Date(lastYear, lastMonth, 1);
  return { year: String(nextDate.getFullYear()), month: nextDate.getMonth() + 1 };
}

function toTimelineComparable(year: string, month: number) {
  return Number(year) * 100 + month;
}

export default function DashboardPage() {
  const now = new Date();
  const queryClient = useQueryClient();
  const stateQuery = useQuery({ queryKey: ["state"], queryFn: getState });
  const assetsQuery = useQuery({ queryKey: ["assets"], queryFn: getAssets });
  const [runYear, setRunYear] = useState(String(now.getFullYear()));
  const [runMonth, setRunMonth] = useState(now.getMonth() + 1);
  const [breakdownYear, setBreakdownYear] = useState(String(now.getFullYear()));
  const [breakdownMonth, setBreakdownMonth] = useState(now.getMonth() + 1);
  const [timelineYear, setTimelineYear] = useState(String(now.getFullYear()));
  const [timelineMode, setTimelineMode] = useState<"year" | "trailing" | "range">("year");
  const [timelineTrailingMonths, setTimelineTrailingMonths] = useState(12);
  const [timelineStartYear, setTimelineStartYear] = useState(String(now.getFullYear()));
  const [timelineStartMonth, setTimelineStartMonth] = useState(1);
  const [timelineEndYear, setTimelineEndYear] = useState(String(now.getFullYear()));
  const [timelineEndMonth, setTimelineEndMonth] = useState(now.getMonth() + 1);
  const [runToken, setRunToken] = useState("");
  const [status, setStatus] = useState("Ready");
  const [review, setReview] = useState<PrepareRunResponse | null>(null);
  const [error, setError] = useState("");

  const summaryQuery = useQuery({
    queryKey: ["summary", breakdownYear, breakdownMonth],
    queryFn: () => getExpensesSummary(breakdownYear, breakdownMonth),
  });
  const incomeSummaryQuery = useQuery({
    queryKey: ["income-summary", breakdownYear, breakdownMonth],
    queryFn: () => getIncomeSummary(breakdownYear, breakdownMonth),
  });
  const investmentsSummaryQuery = useQuery({
    queryKey: ["investments-summary", breakdownYear, breakdownMonth],
    queryFn: () => getInvestmentsSummary(breakdownYear, breakdownMonth),
  });
  const timelineQuery = useQuery({
    queryKey: [
      "totals-timeline",
      timelineMode,
      timelineYear,
      timelineTrailingMonths,
      timelineStartYear,
      timelineStartMonth,
      timelineEndYear,
      timelineEndMonth,
    ],
    queryFn: () =>
      getTotalsTimeline(timelineMode, {
        year: timelineMode === "year" ? timelineYear : undefined,
        trailingMonths: timelineMode === "trailing" ? timelineTrailingMonths : undefined,
        startYear: timelineMode === "range" ? normalizedTimelineRange.start.year : undefined,
        startMonth: timelineMode === "range" ? normalizedTimelineRange.start.month : undefined,
        endYear: timelineMode === "range" ? normalizedTimelineRange.end.year : undefined,
        endMonth: timelineMode === "range" ? normalizedTimelineRange.end.month : undefined,
      }),
  });

  const progressQuery = useQuery({
    queryKey: ["progress", runToken],
    queryFn: () => getRunProgress(runToken),
    enabled: Boolean(runToken),
    refetchInterval: (query) => (query.state.data?.done ? false : 1200),
  });

  const prepareMutation = useMutation({
    mutationFn: prepareRun,
    onSuccess: (data) => {
      setReview(data);
      setStatus("Run prepared. Review items below.");
      setError("");
    },
    onError: (mutationError) => {
      setError(mutationError instanceof Error ? mutationError.message : String(mutationError));
    },
  });

  const finalizeMutation = useMutation({
    mutationFn: finalizeRun,
    onSuccess: () => {
      setReview(null);
      setStatus("Workbook filled successfully.");
      setError("");
      void queryClient.invalidateQueries({ queryKey: ["state"] });
      void queryClient.invalidateQueries({ queryKey: ["assets"] });
      void queryClient.invalidateQueries({ queryKey: ["summary"] });
      void queryClient.invalidateQueries({ queryKey: ["income-summary"] });
      void queryClient.invalidateQueries({ queryKey: ["investments-summary"] });
      void queryClient.invalidateQueries({ queryKey: ["totals-timeline"] });
    },
    onError: (mutationError) => {
      setError(mutationError instanceof Error ? mutationError.message : String(mutationError));
    },
  });

  const saveAssetsMutation = useMutation({
    mutationFn: (payload: Pick<AssetsResponse, "cars" | "investments">) => saveAssets(payload),
    onSuccess: () => {
      void queryClient.invalidateQueries({ queryKey: ["assets"] });
      setStatus("Assets saved.");
    },
    onError: (mutationError) => {
      setError(mutationError instanceof Error ? mutationError.message : String(mutationError));
    },
  });

  const suggestion = nextMonthFromState(stateQuery.data);
  const selectableYears = useMemo(() => {
    const nowYear = new Date().getFullYear();
    const years = new Set([String(nowYear), String(nowYear - 1), String(nowYear - 2), suggestion.year]);
    return [...years].sort((a, b) => Number(a) - Number(b));
  }, [suggestion.year]);
  const normalizedTimelineRange = useMemo(() => {
    const start = { year: timelineStartYear, month: timelineStartMonth };
    const end = { year: timelineEndYear, month: timelineEndMonth };
    if (toTimelineComparable(start.year, start.month) <= toTimelineComparable(end.year, end.month)) {
      return { start, end };
    }
    return { start: end, end: start };
  }, [timelineStartYear, timelineStartMonth, timelineEndYear, timelineEndMonth]);

  const lastFilled = stateQuery.data?.last_filled_year && stateQuery.data?.last_filled_month
    ? `Last filled: ${monthName(Number(stateQuery.data.last_filled_month))} ${stateQuery.data.last_filled_year}`
    : "";

  return (
    <main className="mx-auto flex w-full max-w-7xl flex-col gap-6 p-4 sm:p-8">
      <div className="flex flex-wrap items-center justify-between gap-3 rounded-lg border border-border/70 bg-card/60 px-4 py-3 shadow-sm">
        <div>
          <h1 className="text-2xl font-semibold tracking-tight sm:text-[1.7rem]">AutoExpenses</h1>
          <p className="text-sm text-muted-foreground/90">Dashboard for monthly expense workflows.</p>
        </div>
        {/* <Badge variant="secondary">Dark Theme</Badge> */}
      </div>

      {error ? (
        <Alert variant="destructive">
          <AlertTitle>Something went wrong</AlertTitle>
          <AlertDescription>{error}</AlertDescription>
        </Alert>
      ) : null}

      <section className="grid gap-4 lg:grid-cols-2">
        <RunForm
          year={runYear}
          month={runMonth}
          years={selectableYears}
          busy={prepareMutation.isPending || finalizeMutation.isPending}
          onYearChange={setRunYear}
          onMonthChange={setRunMonth}
          onRun={() => {
            const token = generateRunToken();
            setRunToken(token);
            setStatus(`Running ${monthName(runMonth)} ${runYear}...`);
            setReview(null);
            prepareMutation.mutate({
              year: runYear,
              month: runMonth,
              include_leumi: true,
              run_token: token,
            });
          }}
        />
        <ProgressPanel
          status={status}
          progressStep={progressQuery.data?.step || ""}
          lastFilled={lastFilled}
        />
      </section>

      {review ? (
        <ReviewTable
          items={review.items}
          categories={review.allowed_categories}
          busy={finalizeMutation.isPending}
          onCancel={() => {
            setReview(null);
            setStatus("Review canceled.");
          }}
          onApply={async (reviewedItems: FinalizeRunItem[]) => {
            await finalizeMutation.mutateAsync({
              run_token: runToken,
              review_token: review.review_token,
              reviewed_items: reviewedItems,
            });
          }}
        />
      ) : null}

      <section className="rounded-lg border border-border/70 bg-card/40 p-4 shadow-sm">
        <div className="flex flex-wrap items-end gap-3">
          <div className="space-y-2">
            <Label>Breakdown Year</Label>
            <Select value={breakdownYear} onValueChange={(value) => setBreakdownYear(value ?? breakdownYear)}>
              <SelectTrigger className="w-[140px]">
                <SelectValue placeholder="Year" />
              </SelectTrigger>
              <SelectContent>
                {selectableYears.map((itemYear) => (
                  <SelectItem key={itemYear} value={itemYear}>
                    {itemYear}
                  </SelectItem>
                ))}
              </SelectContent>
            </Select>
          </div>
          <div className="space-y-2">
            <Label>Breakdown Month</Label>
            <Select
              value={String(breakdownMonth)}
              onValueChange={(value) => setBreakdownMonth(Number(value ?? breakdownMonth))}
            >
              <SelectTrigger className="w-[140px]">
                <SelectValue placeholder="Month" />
              </SelectTrigger>
              <SelectContent>
                {Array.from({ length: 12 }).map((_, idx) => (
                  <SelectItem key={idx + 1} value={String(idx + 1)}>
                    {monthName(idx + 1)}
                  </SelectItem>
                ))}
              </SelectContent>
            </Select>
          </div>
          <p className="pb-2 text-sm text-muted-foreground">
            Showing breakdown for {monthName(breakdownMonth)} {breakdownYear}
          </p>
        </div>
      </section>

      <section className="grid gap-4 xl:grid-cols-4">
        <AssetsPanel
          cars={assetsQuery.data?.cars ?? []}
          investments={assetsQuery.data?.investments ?? []}
          leumiBalance={safeNumber(assetsQuery.data?.leumi_balance)}
          onChange={async (nextCars, nextInvestments) => {
            await saveAssetsMutation.mutateAsync({ cars: nextCars, investments: nextInvestments });
          }}
        />
        <SpendingChart
          summary={summaryQuery.data ?? {}}
          title="Spending Breakdown"
          totalLabel="Total Spendings"
          loading={summaryQuery.isLoading || summaryQuery.isFetching}
        />
        <SpendingChart
          summary={incomeSummaryQuery.data ?? {}}
          title="Income Breakdown"
          description="Income categories for selected month."
          totalLabel="Total Income"
          loading={incomeSummaryQuery.isLoading || incomeSummaryQuery.isFetching}
        />
        <SpendingChart
          summary={investmentsSummaryQuery.data ?? {}}
          title="Investment Breakdown"
          description="Investment categories for selected month."
          totalLabel="Total Investments"
          loading={investmentsSummaryQuery.isLoading || investmentsSummaryQuery.isFetching}
        />
      </section>

      <section className="rounded-lg border border-border/70 bg-card/40 p-4 shadow-sm">
        <div className="flex flex-wrap items-end gap-3">
          <div className="space-y-2">
            <Label>Timeline Year</Label>
            <Select value={timelineYear} onValueChange={(value) => setTimelineYear(value ?? timelineYear)}>
              <SelectTrigger className="w-[140px]">
                <SelectValue placeholder="Year" />
              </SelectTrigger>
              <SelectContent>
                {selectableYears.map((itemYear) => (
                  <SelectItem key={`timeline-${itemYear}`} value={itemYear}>
                    {itemYear}
                  </SelectItem>
                ))}
              </SelectContent>
            </Select>
          </div>
          <div className="space-y-2">
            <Label>Timeline Range</Label>
            <Select
              value={
                timelineMode === "year"
                  ? "selected-year"
                  : timelineMode === "range"
                    ? "custom-range"
                    : `trailing-${timelineTrailingMonths}`
              }
              onValueChange={(value) => {
                if (!value) {
                  setTimelineMode("year");
                  return;
                }
                if (value === "selected-year") {
                  setTimelineMode("year");
                  return;
                }
                if (value === "custom-range") {
                  setTimelineMode("range");
                  return;
                }
                const trailingMonths = Number(value.replace("trailing-", ""));
                setTimelineMode("trailing");
                setTimelineTrailingMonths(Number.isFinite(trailingMonths) ? trailingMonths : 12);
              }}
            >
              <SelectTrigger className="w-[220px]">
                <SelectValue placeholder="Timeline range" />
              </SelectTrigger>
              <SelectContent>
                <SelectItem value="selected-year">Selected year (Jan-Dec)</SelectItem>
                <SelectItem value="custom-range">Custom start/end month range</SelectItem>
                <SelectItem value="trailing-12">Last 12 months</SelectItem>
                <SelectItem value="trailing-24">Last 24 months</SelectItem>
                <SelectItem value="trailing-36">Last 36 months</SelectItem>
              </SelectContent>
            </Select>
          </div>
          {timelineMode === "range" ? (
            <>
              <div className="space-y-2">
                <Label>Start</Label>
                <div className="flex gap-2">
                  <Select value={timelineStartYear} onValueChange={(value) => setTimelineStartYear(value ?? timelineStartYear)}>
                    <SelectTrigger className="w-[100px]">
                      <SelectValue placeholder="Year" />
                    </SelectTrigger>
                    <SelectContent>
                      {selectableYears.map((itemYear) => (
                        <SelectItem key={`timeline-start-year-${itemYear}`} value={itemYear}>
                          {itemYear}
                        </SelectItem>
                      ))}
                    </SelectContent>
                  </Select>
                  <Select
                    value={String(timelineStartMonth)}
                    onValueChange={(value) => setTimelineStartMonth(Number(value ?? timelineStartMonth))}
                  >
                    <SelectTrigger className="w-[120px]">
                      <SelectValue placeholder="Month" />
                    </SelectTrigger>
                    <SelectContent>
                      {Array.from({ length: 12 }).map((_, idx) => (
                        <SelectItem key={`timeline-start-month-${idx + 1}`} value={String(idx + 1)}>
                          {monthName(idx + 1)}
                        </SelectItem>
                      ))}
                    </SelectContent>
                  </Select>
                </div>
              </div>
              <div className="space-y-2">
                <Label>End</Label>
                <div className="flex gap-2">
                  <Select value={timelineEndYear} onValueChange={(value) => setTimelineEndYear(value ?? timelineEndYear)}>
                    <SelectTrigger className="w-[100px]">
                      <SelectValue placeholder="Year" />
                    </SelectTrigger>
                    <SelectContent>
                      {selectableYears.map((itemYear) => (
                        <SelectItem key={`timeline-end-year-${itemYear}`} value={itemYear}>
                          {itemYear}
                        </SelectItem>
                      ))}
                    </SelectContent>
                  </Select>
                  <Select
                    value={String(timelineEndMonth)}
                    onValueChange={(value) => setTimelineEndMonth(Number(value ?? timelineEndMonth))}
                  >
                    <SelectTrigger className="w-[120px]">
                      <SelectValue placeholder="Month" />
                    </SelectTrigger>
                    <SelectContent>
                      {Array.from({ length: 12 }).map((_, idx) => (
                        <SelectItem key={`timeline-end-month-${idx + 1}`} value={String(idx + 1)}>
                          {monthName(idx + 1)}
                        </SelectItem>
                      ))}
                    </SelectContent>
                  </Select>
                </div>
              </div>
            </>
          ) : null}
          <p className="pb-2 text-sm text-muted-foreground">
            {timelineMode === "year"
              ? `Timeline for ${timelineYear}`
              : timelineMode === "range"
                ? `Timeline from ${monthName(normalizedTimelineRange.start.month)} ${normalizedTimelineRange.start.year} to ${monthName(normalizedTimelineRange.end.month)} ${normalizedTimelineRange.end.year}`
                : `Timeline for latest ${timelineTrailingMonths} months`}
          </p>
        </div>
      </section>

      <TimelineChart
        points={timelineQuery.data?.points ?? []}
        loading={timelineQuery.isLoading || timelineQuery.isFetching}
      />
    </main>
  );
}
