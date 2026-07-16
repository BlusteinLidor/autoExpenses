"use client";

import { AssetsPanel } from "@/components/assets/AssetsPanel";
import { SpendingChart } from "@/components/charts/SpendingChart";
import { TimelineChart } from "@/components/charts/TimelineChart";
import { CategorySelectItems } from "@/components/review/CategorySelectItems";
import { ReviewTable } from "@/components/review/ReviewTable";
import { ProgressPanel } from "@/components/run/ProgressPanel";
import { RunForm } from "@/components/run/RunForm";
import { Alert, AlertDescription, AlertTitle } from "@/components/ui/alert";
import { Button } from "@/components/ui/button";
import { Label } from "@/components/ui/label";
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from "@/components/ui/select";
import {
  disconnectDrive,
  finalizeRun,
  getDriveStatus,
  getAssets,
  getCategories,
  getCategoryTimeline,
  getExpensesSummary,
  getIncomeSummary,
  getInvestmentsSummary,
  getTotalsTimeline,
  getRunProgress,
  getState,
  prepareRun,
  saveAssets,
  startDriveConnect,
} from "@/lib/api/client";
import { AssetsResponse, DriveUploadResult, FinalizeRunItem, PrepareRunResponse } from "@/lib/api/types";
import { generateRunToken, monthName, safeNumber } from "@/lib/format";
import { useMutation, useQuery, useQueryClient } from "@tanstack/react-query";
import { useEffect, useMemo, useState } from "react";

const CATEGORY_TIMELINE_STROKES: Record<string, string> = {
  income: "#22c55e",
  necessary_expenses: "#f43f5e",
  luxury_expenses: "#fb7185",
  investments: "#6366f1",
};

function previousCalendarMonth(from = new Date()) {
  const date = new Date(from.getFullYear(), from.getMonth() - 1, 1);
  return { year: String(date.getFullYear()), month: date.getMonth() + 1 };
}

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
  const defaultPeriod = previousCalendarMonth(now);
  const queryClient = useQueryClient();
  const stateQuery = useQuery({ queryKey: ["state"], queryFn: getState });
  const assetsQuery = useQuery({ queryKey: ["assets"], queryFn: getAssets });
  const [runYear, setRunYear] = useState(defaultPeriod.year);
  const [runMonth, setRunMonth] = useState(defaultPeriod.month);
  const [breakdownYear, setBreakdownYear] = useState(defaultPeriod.year);
  const [breakdownMonth, setBreakdownMonth] = useState(defaultPeriod.month);
  const [timelineYear, setTimelineYear] = useState(String(now.getFullYear()));
  const [timelineMode, setTimelineMode] = useState<"year" | "trailing" | "range">("year");
  const [timelineTrailingMonths, setTimelineTrailingMonths] = useState(12);
  const [timelineStartYear, setTimelineStartYear] = useState(String(now.getFullYear()));
  const [timelineStartMonth, setTimelineStartMonth] = useState(1);
  const [timelineEndYear, setTimelineEndYear] = useState(String(now.getFullYear()));
  const [timelineEndMonth, setTimelineEndMonth] = useState(now.getMonth() + 1);
  const [timelineView, setTimelineView] = useState<"totals" | "category">("totals");
  const [timelineCategory, setTimelineCategory] = useState("");
  const [runToken, setRunToken] = useState("");
  const [status, setStatus] = useState("Ready");
  const [review, setReview] = useState<PrepareRunResponse | null>(null);
  const [error, setError] = useState("");
  const [driveUploadResult, setDriveUploadResult] = useState<DriveUploadResult | null>(null);
  const [driveConnectResult, setDriveConnectResult] = useState<{
    connected: boolean | null;
    error: string | null;
  }>({ connected: null, error: null });

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
  const categoriesQuery = useQuery({
    queryKey: ["categories"],
    queryFn: getCategories,
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
    enabled: timelineView === "totals",
  });
  const categoryTimelineQuery = useQuery({
    queryKey: [
      "category-timeline",
      timelineCategory,
      timelineMode,
      timelineYear,
      timelineTrailingMonths,
      timelineStartYear,
      timelineStartMonth,
      timelineEndYear,
      timelineEndMonth,
    ],
    queryFn: () =>
      getCategoryTimeline(timelineCategory, timelineMode, {
        year: timelineMode === "year" ? timelineYear : undefined,
        trailingMonths: timelineMode === "trailing" ? timelineTrailingMonths : undefined,
        startYear: timelineMode === "range" ? normalizedTimelineRange.start.year : undefined,
        startMonth: timelineMode === "range" ? normalizedTimelineRange.start.month : undefined,
        endYear: timelineMode === "range" ? normalizedTimelineRange.end.year : undefined,
        endMonth: timelineMode === "range" ? normalizedTimelineRange.end.month : undefined,
      }),
    enabled: timelineView === "category" && Boolean(timelineCategory),
  });

  const progressQuery = useQuery({
    queryKey: ["progress", runToken],
    queryFn: () => getRunProgress(runToken),
    enabled: Boolean(runToken),
    refetchInterval: (query) => (query.state.data?.done ? false : 1200),
  });
  const driveStatusQuery = useQuery({ queryKey: ["drive-status"], queryFn: getDriveStatus });

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
    onSuccess: (data) => {
      setReview(null);
      setStatus("Workbook filled successfully.");
      setError("");
      setDriveUploadResult(data.drive_upload ?? null);
      void queryClient.invalidateQueries({ queryKey: ["state"] });
      void queryClient.invalidateQueries({ queryKey: ["assets"] });
      void queryClient.invalidateQueries({ queryKey: ["summary"] });
      void queryClient.invalidateQueries({ queryKey: ["income-summary"] });
      void queryClient.invalidateQueries({ queryKey: ["investments-summary"] });
      void queryClient.invalidateQueries({ queryKey: ["totals-timeline"] });
      void queryClient.invalidateQueries({ queryKey: ["category-timeline"] });
      void queryClient.invalidateQueries({ queryKey: ["categories"] });
      void queryClient.invalidateQueries({ queryKey: ["drive-status"] });
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
  const driveConnectMutation = useMutation({
    mutationFn: startDriveConnect,
    onSuccess: (data) => {
      window.location.assign(data.authorization_url);
    },
    onError: (mutationError) => {
      setError(mutationError instanceof Error ? mutationError.message : String(mutationError));
    },
  });
  const driveDisconnectMutation = useMutation({
    mutationFn: disconnectDrive,
    onSuccess: () => {
      setDriveUploadResult(null);
      setStatus("Google Drive disconnected.");
      void queryClient.invalidateQueries({ queryKey: ["drive-status"] });
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
  const normalizedTimelineRange = (() => {
    const start = { year: timelineStartYear, month: timelineStartMonth };
    const end = { year: timelineEndYear, month: timelineEndMonth };
    if (toTimelineComparable(start.year, start.month) <= toTimelineComparable(end.year, end.month)) {
      return { start, end };
    }
    return { start: end, end: start };
  })();
  const categoryGroups = categoriesQuery.data?.groups ?? [];
  const availableCategories = categoriesQuery.data?.categories ?? [];

  useEffect(() => {
    if (!availableCategories.length) {
      return;
    }
    if (timelineCategory && availableCategories.includes(timelineCategory)) {
      return;
    }
    const preferred =
      categoryGroups.find((group) => group.id === "necessary_expenses")?.categories[0] ??
      categoryGroups.find((group) => group.id === "luxury_expenses")?.categories[0] ??
      availableCategories[0] ??
      "";
    setTimelineCategory(preferred);
  }, [availableCategories, categoryGroups, timelineCategory]);

  const timelineCategoryStroke = useMemo(() => {
    const groupId = categoryGroups.find((group) => group.categories.includes(timelineCategory))?.id;
    return (groupId && CATEGORY_TIMELINE_STROKES[groupId]) || "#0ea5e9";
  }, [categoryGroups, timelineCategory]);

  const lastFilled = stateQuery.data?.last_filled_year && stateQuery.data?.last_filled_month
    ? `Last filled: ${monthName(Number(stateQuery.data.last_filled_month))} ${stateQuery.data.last_filled_year}`
    : "";
  useEffect(() => {
    const params = new URLSearchParams(window.location.search);
    const connected = params.get("drive_connected");
    const driveError = params.get("drive_error");
    if (connected === "1") {
      setDriveConnectResult({ connected: true, error: null });
      return;
    }
    if (connected === "0") {
      setDriveConnectResult({
        connected: false,
        error: driveError || "Google Drive connection failed.",
      });
      return;
    }
    setDriveConnectResult({ connected: null, error: null });
  }, []);

  useEffect(() => {
    if (driveConnectResult.connected === true || driveConnectResult.connected === false) {
      void queryClient.invalidateQueries({ queryKey: ["drive-status"] });
      window.history.replaceState({}, "", window.location.pathname);
    }
  }, [driveConnectResult.connected, queryClient]);

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
      {driveConnectResult.connected === true ? (
        <Alert>
          <AlertTitle>Google Drive connected</AlertTitle>
          <AlertDescription>Your dashboard is now connected to Google Drive.</AlertDescription>
        </Alert>
      ) : null}
      {driveConnectResult.connected === false && driveConnectResult.error ? (
        <Alert variant="destructive">
          <AlertTitle>Google Drive connection failed</AlertTitle>
          <AlertDescription>{driveConnectResult.error}</AlertDescription>
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

      <section className="rounded-lg border border-border/70 bg-card/40 p-4 shadow-sm">
        <div className="flex flex-wrap items-center justify-between gap-3">
          <div className="space-y-1">
            <h2 className="text-lg font-medium">Google Drive</h2>
            {driveStatusQuery.isError ? (
              <p className="text-sm text-destructive">
                Could not load Google Drive status. Check that the API server is running.
              </p>
            ) : !driveStatusQuery.data?.configured ? (
              <p className="text-sm text-muted-foreground">
                OAuth is not configured on the server yet.
              </p>
            ) : driveStatusQuery.data?.connected ? (
              <p className="text-sm text-muted-foreground">
                Connected as {driveStatusQuery.data.email || "your Google account"}.
              </p>
            ) : (
              <p className="text-sm text-muted-foreground">
                {driveStatusQuery.data.last_error
                  ? driveStatusQuery.data.last_error
                  : `Not connected. Connect to auto-upload files to AutoExpenses/${"{year}"}/${"{month}"}.`}
              </p>
            )}
          </div>
          <div className="flex gap-2">
            {driveStatusQuery.data?.connected ? (
              <Button
                variant="outline"
                onClick={() => driveDisconnectMutation.mutate()}
                disabled={driveDisconnectMutation.isPending}
              >
                {driveDisconnectMutation.isPending ? "Disconnecting..." : "Disconnect"}
              </Button>
            ) : (
              <Button
                onClick={() => driveConnectMutation.mutate()}
                disabled={!driveStatusQuery.data?.configured || driveConnectMutation.isPending}
              >
                {driveConnectMutation.isPending ? "Opening Google..." : "Connect Google Drive"}
              </Button>
            )}
          </div>
        </div>
      </section>

      {driveUploadResult ? (
        <Alert variant={driveUploadResult.success ? "default" : "destructive"}>
          <AlertTitle>
            {driveUploadResult.success
              ? "Google Drive upload completed"
              : "Google Drive upload did not complete"}
          </AlertTitle>
          <AlertDescription>
            {driveUploadResult.success ? (
              <>
                Uploaded to <strong>{driveUploadResult.folder_path}</strong>.
                {"  "}
                {driveUploadResult.monthly_file?.web_view_link ? (
                  <a
                    className="underline"
                    href={driveUploadResult.monthly_file.web_view_link}
                    target="_blank"
                    rel="noreferrer"
                  >
                    Monthly file
                  </a>
                ) : null}
                {"  "}
                {driveUploadResult.yearly_file?.web_view_link ? (
                  <a
                    className="underline"
                    href={driveUploadResult.yearly_file.web_view_link}
                    target="_blank"
                    rel="noreferrer"
                  >
                    Yearly file
                  </a>
                ) : null}
              </>
            ) : (
              driveUploadResult.error || "Local run succeeded, but Google Drive upload failed."
            )}
          </AlertDescription>
        </Alert>
      ) : null}

      {review ? (
        <ReviewTable
          items={review.items}
          categories={review.allowed_categories}
          categoryGroups={review.category_groups}
          warnings={review.warnings}
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
            <Label>Timeline View</Label>
            <Select
              value={timelineView}
              onValueChange={(value) => setTimelineView(value === "category" ? "category" : "totals")}
            >
              <SelectTrigger className="w-[260px]">
                <SelectValue placeholder="Timeline view" />
              </SelectTrigger>
              <SelectContent>
                <SelectItem value="totals">Income / Spending / Investments</SelectItem>
                <SelectItem value="category">By category</SelectItem>
              </SelectContent>
            </Select>
          </div>
          {timelineView === "category" ? (
            <div className="space-y-2">
              <Label>Category</Label>
              <Select
                value={timelineCategory || undefined}
                onValueChange={(value) => setTimelineCategory(value ?? "")}
                disabled={!availableCategories.length}
              >
                <SelectTrigger className="w-[240px]">
                  <SelectValue placeholder={categoriesQuery.isLoading ? "Loading…" : "Select category"} />
                </SelectTrigger>
                <SelectContent>
                  <CategorySelectItems
                    categories={availableCategories}
                    categoryGroups={categoryGroups}
                  />
                </SelectContent>
              </Select>
            </div>
          ) : null}
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
            {timelineView === "category" && timelineCategory ? `${timelineCategory} · ` : ""}
            {timelineMode === "year"
              ? `Timeline for ${timelineYear}`
              : timelineMode === "range"
                ? `Timeline from ${monthName(normalizedTimelineRange.start.month)} ${normalizedTimelineRange.start.year} to ${monthName(normalizedTimelineRange.end.month)} ${normalizedTimelineRange.end.year}`
                : `Timeline for latest ${timelineTrailingMonths} months`}
          </p>
        </div>
      </section>

      {timelineView === "category" ? (
        <TimelineChart
          view="category"
          category={timelineCategory || "Category"}
          points={categoryTimelineQuery.data?.points ?? []}
          stroke={timelineCategoryStroke}
          loading={
            categoriesQuery.isLoading ||
            categoryTimelineQuery.isLoading ||
            categoryTimelineQuery.isFetching
          }
        />
      ) : (
        <TimelineChart
          points={timelineQuery.data?.points ?? []}
          loading={timelineQuery.isLoading || timelineQuery.isFetching}
        />
      )}
    </main>
  );
}
