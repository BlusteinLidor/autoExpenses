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

export function incomeCategories(categoryGroups?: { id: string; categories: string[] }[]) {
  const incomeGroup = categoryGroups?.find((group) => group.id === "income");
  return incomeGroup?.categories ?? [];
}

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

export function generateRunToken() {
  if (typeof crypto !== "undefined" && typeof crypto.randomUUID === "function") {
    return crypto.randomUUID();
  }
  return `${Date.now()}-${Math.random().toString(16).slice(2)}`;
}
