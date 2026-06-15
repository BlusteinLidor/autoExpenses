export function monthName(month: number) {
  const names = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
  return names[month - 1] ?? String(month);
}

export function safeNumber(value: unknown) {
  if (typeof value === "number") {
    return Number.isFinite(value) ? value : 0;
  }
  if (typeof value === "string") {
    const direct = Number(value.replaceAll(",", "").trim());
    if (Number.isFinite(direct)) {
      return direct;
    }
    const token = value.match(/[-+]?\d[\d,]*(?:\.\d+)?/);
    if (!token) {
      return 0;
    }
    const parsed = Number(token[0].replaceAll(",", ""));
    return Number.isFinite(parsed) ? parsed : 0;
  }
  return 0;
}

export function money(value: unknown) {
  return new Intl.NumberFormat(undefined, {
    maximumFractionDigits: 2,
  }).format(safeNumber(value));
}

export function isSplitExpenseName(name: string) {
  return name.replace(/\s+/g, " ").trim() === "העברה דיגיטל";
}

export type TransactionKind = "expense" | "investment" | "income";

const DEFAULT_INCOME_CATEGORY = "הכנסה אחרת / חד פעמית";

export function incomeCategories(categoryGroups?: { id: string; categories: string[] }[]) {
  const incomeGroup = categoryGroups?.find((group) => group.id === "income");
  return incomeGroup?.categories ?? [];
}

export function investmentCategories(categoryGroups?: { id: string; categories: string[] }[]) {
  const investmentGroup = categoryGroups?.find((group) => group.id === "investments");
  return investmentGroup?.categories ?? [];
}

export function spendingCategories(
  categories: string[],
  categoryGroups?: { id: string; categories: string[] }[],
) {
  const excluded = new Set([
    ...incomeCategories(categoryGroups),
    ...investmentCategories(categoryGroups),
  ]);
  if (excluded.size === 0) {
    return categories;
  }
  return categories.filter((category) => !excluded.has(category));
}

/** All non-income categories (spending + investments). */
export function expenseCategories(
  categories: string[],
  categoryGroups?: { id: string; categories: string[] }[],
) {
  const income = new Set(incomeCategories(categoryGroups));
  if (income.size === 0) {
    return categories;
  }
  return categories.filter((category) => !income.has(category));
}

export function categoryGroupsForTransactionKind(
  kind: TransactionKind,
  categoryGroups?: { id: string; label: string; categories: string[] }[],
) {
  if (!categoryGroups?.length) {
    return undefined;
  }
  if (kind === "income") {
    return categoryGroups.filter((group) => group.id === "income");
  }
  if (kind === "investment") {
    return categoryGroups.filter((group) => group.id === "investments");
  }
  return categoryGroups.filter(
    (group) => group.id === "necessary_expenses" || group.id === "luxury_expenses",
  );
}

export function categoriesForTransactionKind(
  kind: TransactionKind,
  categories: string[],
  categoryGroups?: { id: string; categories: string[] }[],
) {
  if (kind === "income") {
    const incomeOnly = incomeCategories(categoryGroups);
    return incomeOnly.length > 0 ? incomeOnly : categories;
  }
  if (kind === "investment") {
    const investmentOnly = investmentCategories(categoryGroups);
    return investmentOnly.length > 0 ? investmentOnly : [];
  }
  const spendingOnly = spendingCategories(categories, categoryGroups);
  return spendingOnly.length > 0 ? spendingOnly : expenseCategories(categories, categoryGroups);
}

export function inferTransactionKind(
  item: { is_income?: boolean; category?: string },
  categoryGroups?: { id: string; categories: string[] }[],
): TransactionKind {
  if (item.is_income) {
    return "income";
  }
  const investments = new Set(investmentCategories(categoryGroups));
  if (item.category && investments.has(item.category)) {
    return "investment";
  }
  return "expense";
}

export function defaultCategoryForTransactionKind(
  kind: TransactionKind,
  item: { category?: string; needs_manual_review?: boolean },
  categories: string[],
  categoryGroups?: { id: string; categories: string[] }[],
) {
  const allowed = categoriesForTransactionKind(kind, categories, categoryGroups);
  if (item.category && allowed.includes(item.category)) {
    return item.category;
  }
  if (kind === "income") {
    return allowed.includes(DEFAULT_INCOME_CATEGORY)
      ? DEFAULT_INCOME_CATEGORY
      : allowed[0] || "";
  }
  return item.needs_manual_review ? "" : allowed[0] || "";
}

export function generateRunToken() {
  if (typeof crypto !== "undefined" && typeof crypto.randomUUID === "function") {
    return crypto.randomUUID();
  }
  return `${Date.now()}-${Math.random().toString(16).slice(2)}`;
}
