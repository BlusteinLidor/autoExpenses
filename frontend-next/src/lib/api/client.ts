import {
  AppState,
  AssetsResponse,
  ExpensesSummary,
  TotalsTimelineResponse,
  FinalizeRunPayload,
  FinalizeRunResponse,
  PrepareRunPayload,
  PrepareRunResponse,
  ProgressResponse,
} from "./types";

const API_BASE_URL = process.env.NEXT_PUBLIC_API_BASE_URL?.trim() ?? "";

async function apiFetch<T>(path: string, init?: RequestInit): Promise<T> {
  const response = await fetch(`${API_BASE_URL}${path}`, {
    ...init,
    headers: {
      "Content-Type": "application/json",
      ...(init?.headers ?? {}),
    },
    cache: "no-store",
  });

  let payload: unknown;
  try {
    payload = await response.json();
  } catch {
    payload = null;
  }

  if (!response.ok) {
    const detail =
      payload && typeof payload === "object" && "detail" in payload
        ? String((payload as { detail?: string }).detail)
        : `Request failed (${response.status})`;
    throw new Error(detail);
  }
  return payload as T;
}

export function getState() {
  return apiFetch<AppState>("/state");
}

export function getAssets() {
  return apiFetch<AssetsResponse>("/assets");
}

export function saveAssets(payload: Pick<AssetsResponse, "cars" | "investments">) {
  return apiFetch<AppState>("/assets", {
    method: "POST",
    body: JSON.stringify(payload),
  });
}

export function getExpensesSummary(year: string, month: number) {
  return apiFetch<ExpensesSummary>(
    `/expenses/summary?year=${encodeURIComponent(year)}&month=${encodeURIComponent(month)}`,
  );
}

export function getIncomeSummary(year: string, month: number) {
  return apiFetch<ExpensesSummary>(
    `/income/summary?year=${encodeURIComponent(year)}&month=${encodeURIComponent(month)}`,
  );
}

export function getInvestmentsSummary(year: string, month: number) {
  return apiFetch<ExpensesSummary>(
    `/investments/summary?year=${encodeURIComponent(year)}&month=${encodeURIComponent(month)}`,
  );
}

export function getTotalsTimeline(mode: "year" | "trailing", year?: string, trailingMonths?: number) {
  const params = new URLSearchParams({ mode });
  if (year) {
    params.set("year", year);
  }
  if (typeof trailingMonths === "number") {
    params.set("trailing_months", String(trailingMonths));
  }
  return apiFetch<TotalsTimelineResponse>(`/totals/timeline?${params.toString()}`);
}

export function prepareRun(payload: PrepareRunPayload) {
  return apiFetch<PrepareRunResponse>("/run-month/prepare", {
    method: "POST",
    body: JSON.stringify(payload),
  });
}

export function finalizeRun(payload: FinalizeRunPayload) {
  return apiFetch<FinalizeRunResponse>("/run-month/finalize", {
    method: "POST",
    body: JSON.stringify(payload),
  });
}

export function getRunProgress(runToken: string) {
  return apiFetch<ProgressResponse>(`/run-month/progress/${encodeURIComponent(runToken)}`);
}
