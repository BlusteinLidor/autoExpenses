"use client";

import { AssetsPanel } from "@/components/assets/AssetsPanel";
import { SpendingChart } from "@/components/charts/SpendingChart";
import { ReviewTable } from "@/components/review/ReviewTable";
import { ProgressPanel } from "@/components/run/ProgressPanel";
import { RunForm } from "@/components/run/RunForm";
import { Alert, AlertDescription, AlertTitle } from "@/components/ui/alert";
import { Badge } from "@/components/ui/badge";
import {
  finalizeRun,
  getAssets,
  getExpensesSummary,
  getIncomeSummary,
  getInvestmentsSummary,
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

export default function DashboardPage() {
  const queryClient = useQueryClient();
  const stateQuery = useQuery({ queryKey: ["state"], queryFn: getState });
  const assetsQuery = useQuery({ queryKey: ["assets"], queryFn: getAssets });
  const [year, setYear] = useState(String(new Date().getFullYear()));
  const [month, setMonth] = useState(new Date().getMonth() + 1);
  const [includeLeumi, setIncludeLeumi] = useState(true);
  const [runToken, setRunToken] = useState("");
  const [status, setStatus] = useState("Ready");
  const [review, setReview] = useState<PrepareRunResponse | null>(null);
  const [error, setError] = useState("");

  const summaryQuery = useQuery({
    queryKey: ["summary", year, month],
    queryFn: () => getExpensesSummary(year, month),
  });
  const incomeSummaryQuery = useQuery({
    queryKey: ["income-summary", year, month],
    queryFn: () => getIncomeSummary(year, month),
  });
  const investmentsSummaryQuery = useQuery({
    queryKey: ["investments-summary", year, month],
    queryFn: () => getInvestmentsSummary(year, month),
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
      void queryClient.invalidateQueries({ queryKey: ["summary", year, month] });
      void queryClient.invalidateQueries({ queryKey: ["income-summary", year, month] });
      void queryClient.invalidateQueries({ queryKey: ["investments-summary", year, month] });
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
          year={year}
          month={month}
          years={selectableYears}
          includeLeumi={includeLeumi}
          busy={prepareMutation.isPending || finalizeMutation.isPending}
          onYearChange={setYear}
          onMonthChange={setMonth}
          onIncludeLeumiChange={setIncludeLeumi}
          onRun={() => {
            const token = generateRunToken();
            setRunToken(token);
            setStatus(`Running ${monthName(month)} ${year}...`);
            setReview(null);
            prepareMutation.mutate({
              year,
              month,
              include_leumi: includeLeumi,
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
          selectedMonthLabel={`${monthName(month)} ${year}`}
          title="Spending Breakdown"
          totalLabel="Total Spendings"
        />
        <SpendingChart
          summary={incomeSummaryQuery.data ?? {}}
          selectedMonthLabel={`${monthName(month)} ${year}`}
          title="Income Breakdown"
          description="Income categories for selected month."
          totalLabel="Total Income"
        />
        <SpendingChart
          summary={investmentsSummaryQuery.data ?? {}}
          selectedMonthLabel={`${monthName(month)} ${year}`}
          title="Investment Breakdown"
          description="Investment categories for selected month."
          totalLabel="Total Investments"
        />
      </section>
    </main>
  );
}
