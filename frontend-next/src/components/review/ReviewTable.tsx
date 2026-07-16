"use client";

import { CategorySelectItems } from "@/components/review/CategorySelectItems";
import { SplitEditor, SplitPart } from "@/components/review/SplitEditor";
import { Alert, AlertDescription, AlertTitle } from "@/components/ui/alert";
import { Badge } from "@/components/ui/badge";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from "@/components/ui/select";
import { Table, TableBody, TableCell, TableHead, TableHeader, TableRow } from "@/components/ui/table";
import { CategoryGroup, FinalizeRunItem, ReviewItem } from "@/lib/api/types";
import {
  categoriesForTransactionKind,
  categoryGroupsForTransactionKind,
  defaultCategoryForTransactionKind,
  formatExpenseDate,
  inferTransactionKind,
  isSplitExpenseName,
  TransactionKind,
} from "@/lib/format";
import { Plus } from "lucide-react";
import { useMemo, useState } from "react";

const TRANSACTION_KIND_OPTIONS: { value: TransactionKind; label: string }[] = [
  { value: "expense", label: "Expense" },
  { value: "investment", label: "Investment" },
  { value: "income", label: "Income" },
];

function pickCategoryForKind(
  kind: TransactionKind,
  currentCategory: string,
  categories: string[],
  categoryGroups?: CategoryGroup[],
  needsManualReview?: boolean,
) {
  const allowed = categoriesForTransactionKind(kind, categories, categoryGroups);
  if (currentCategory && allowed.includes(currentCategory)) {
    return currentCategory;
  }
  return defaultCategoryForTransactionKind(
    kind,
    { category: currentCategory, needs_manual_review: needsManualReview },
    categories,
    categoryGroups,
  );
}

type RowState = {
  include: boolean;
  name: string;
  cost: number;
  transactionKind: TransactionKind;
  category: string;
  source: string;
  date: string;
  splitEnabled: boolean;
  splitParts: SplitPart[];
  isManual: boolean;
  needsManualReview: boolean;
  isPossibleDuplicate: boolean;
  errorReason: string | null;
};

function buildInitialRow(
  item: ReviewItem,
  categories: string[],
  categoryGroups?: CategoryGroup[],
  isManual = false,
): RowState {
  const transactionKind = inferTransactionKind(item, categoryGroups);
  const defaultCategory = defaultCategoryForTransactionKind(
    transactionKind,
    item,
    categories,
    categoryGroups,
  );
  const cost = Math.abs(Number(item.cost || 0));
  const splitEnabled = isSplitExpenseName(item.name);
  return {
    include: item.needs_manual_review ? true : !item.is_possible_duplicate,
    name: item.name,
    cost,
    transactionKind,
    category: defaultCategory,
    source: isManual ? "Manual" : String(item.source || "").trim(),
    date: isManual ? "" : String(item.date || "").trim(),
    splitEnabled,
    splitParts: [
      { amount: cost, transactionKind, category: defaultCategory },
      { amount: 0, transactionKind, category: defaultCategory },
    ],
    isManual,
    needsManualReview: Boolean(item.needs_manual_review),
    isPossibleDuplicate: Boolean(item.is_possible_duplicate),
    errorReason: item.error_reason ?? null,
  };
}

type ReviewTableProps = {
  items: ReviewItem[];
  categories: string[];
  categoryGroups?: CategoryGroup[];
  warnings?: string[];
  onApply: (items: FinalizeRunItem[]) => Promise<void>;
  onCancel: () => void;
  busy: boolean;
};

export function ReviewTable({
  items,
  categories,
  categoryGroups,
  warnings = [],
  onApply,
  onCancel,
  busy,
}: ReviewTableProps) {
  const [rows, setRows] = useState<RowState[]>(() =>
    items.map((item) => buildInitialRow(item, categories, categoryGroups)),
  );
  const [error, setError] = useState("");

  const hasSplitRows = useMemo(() => rows.some((row) => row.splitEnabled), [rows]);
  const manualReviewCount = useMemo(
    () => rows.filter((row) => row.needsManualReview).length,
    [rows],
  );
  const kindCounts = useMemo(
    () =>
      rows.reduce(
        (acc, row) => {
          if (row.splitEnabled) {
            for (const part of row.splitParts) {
              if (Math.abs(Number(part.amount || 0)) > 0) {
                acc[part.transactionKind] += 1;
              }
            }
          } else {
            acc[row.transactionKind] += 1;
          }
          return acc;
        },
        { expense: 0, investment: 0, income: 0 } as Record<TransactionKind, number>,
      ),
    [rows],
  );

  const updateRow = (idx: number, patch: Partial<RowState>) => {
    setRows((current) => {
      const nextRows = [...current];
      nextRows[idx] = { ...nextRows[idx], ...patch };
      return nextRows;
    });
  };

  const updateTransactionKind = (idx: number, kind: TransactionKind) => {
    const row = rows[idx];
    const nextCategory = pickCategoryForKind(
      kind,
      row.category,
      categories,
      categoryGroups,
      row.needsManualReview,
    );
    if (row.splitEnabled) {
      // Row type is the default for new split parts; leave existing parts alone.
      updateRow(idx, {
        transactionKind: kind,
        category: nextCategory,
      });
      return;
    }
    updateRow(idx, {
      transactionKind: kind,
      category: nextCategory,
      splitParts: row.splitParts.map((part) => ({
        ...part,
        transactionKind: kind,
        category: pickCategoryForKind(kind, part.category, categories, categoryGroups),
      })),
    });
  };

  const setSplitEnabled = (idx: number, enabled: boolean) => {
    const row = rows[idx];
    if (enabled) {
      updateRow(idx, {
        splitEnabled: true,
        splitParts:
          row.splitParts.length > 0
            ? row.splitParts.map((part) => ({
                ...part,
                transactionKind: part.transactionKind || row.transactionKind,
              }))
            : [
                {
                  amount: Math.abs(Number(row.cost || 0)),
                  transactionKind: row.transactionKind,
                  category: row.category,
                },
                {
                  amount: 0,
                  transactionKind: row.transactionKind,
                  category: row.category,
                },
              ],
      });
      return;
    }
    const firstPart = row.splitParts.find((part) => Math.abs(Number(part.amount || 0)) > 0);
    updateRow(idx, {
      splitEnabled: false,
      transactionKind: firstPart?.transactionKind ?? row.transactionKind,
      category: firstPart?.category || row.category,
    });
  };

  const addManualExpense = () => {
    setRows((current) => [
      ...current,
      buildInitialRow(
        {
          name: "",
          cost: 0,
          category: "",
          is_possible_duplicate: false,
          needs_manual_review: true,
        },
        categories,
        categoryGroups,
        true,
      ),
    ]);
  };

  const removeManualRow = (idx: number) => {
    setRows((current) => current.filter((_, index) => index !== idx));
  };

  const buildFinalizeItems = () => {
    const finalizeItems: FinalizeRunItem[] = [];
    for (let idx = 0; idx < rows.length; idx += 1) {
      const row = rows[idx];
      if (!row?.include) continue;
      const displayName = row.name.replace(/\s+/g, " ").trim();
      if (!displayName) {
        throw new Error("Enter a name for each included transaction.");
      }
      if (row.needsManualReview && !row.splitEnabled && !row.category) {
        throw new Error(`Choose a category for '${displayName}' before applying.`);
      }
      if (row.splitEnabled) {
        let total = 0;
        let partNo = 1;
        for (const part of row.splitParts) {
          const amount = Math.abs(Number(part.amount || 0));
          if (amount <= 0) {
            partNo += 1;
            continue;
          }
          if (!part.category) {
            throw new Error(`Choose a category for split part ${partNo} of '${displayName}'.`);
          }
          finalizeItems.push({
            name: `${displayName} חלק ${partNo}`,
            cost: amount,
            category: part.category,
          });
          total += amount;
          partNo += 1;
        }
        const expectedTotal = Math.abs(Number(row.cost || 0));
        if (Math.abs(total - expectedTotal) > 0.01) {
          throw new Error(`Split total for '${displayName}' must equal ${expectedTotal}.`);
        }
        if (total <= 0) {
          throw new Error(`Split for '${displayName}' must include at least one positive amount.`);
        }
      } else {
        const cost = Math.abs(Number(row.cost || 0));
        if (cost <= 0) {
          throw new Error(`Amount for '${displayName}' must be greater than zero.`);
        }
        if (!row.category) {
          throw new Error(`Choose a category for '${displayName}' before applying.`);
        }
        finalizeItems.push({
          name: displayName,
          cost,
          category: row.category,
        });
      }
    }
    return finalizeItems;
  };

  return (
    <div className="space-y-4 rounded-xl border bg-card p-4">
      <div className="flex items-center justify-between gap-2">
        <div>
          <h2 className="text-lg font-semibold">Review Transactions</h2>
          <p className="text-sm text-muted-foreground">
            Edit amounts, transaction type, and categories before applying. Split any row by amount
            and type, or add a transaction manually. {kindCounts.expense} expenses
            {kindCounts.investment > 0 ? `, ${kindCounts.investment} investments` : ""}
            {kindCounts.income > 0 ? `, ${kindCounts.income} income` : ""} in this review.
          </p>
        </div>
        <div className="flex flex-wrap gap-2">
          {kindCounts.income > 0 ? (
            <Badge className="border-emerald-500/40 bg-emerald-500/15 text-emerald-200">
              {kindCounts.income} income
            </Badge>
          ) : null}
          {kindCounts.investment > 0 ? (
            <Badge className="border-indigo-500/40 bg-indigo-500/15 text-indigo-200">
              {kindCounts.investment} investments
            </Badge>
          ) : null}
          {manualReviewCount > 0 ? (
            <Badge variant="destructive">{manualReviewCount} need manual review</Badge>
          ) : null}
          {hasSplitRows ? <Badge variant="secondary">Split enabled</Badge> : null}
        </div>
      </div>
      {warnings.length > 0 ? (
        <Alert>
          <AlertTitle>Attention needed</AlertTitle>
          <AlertDescription>
            <ul className="list-disc space-y-1 pl-5">
              {warnings.map((warning) => (
                <li key={warning}>{warning}</li>
              ))}
            </ul>
          </AlertDescription>
        </Alert>
      ) : null}
      {error ? <p className="text-sm text-destructive">{error}</p> : null}

      <div className="overflow-auto rounded-md border">
        <Table>
          <TableHeader>
            <TableRow>
              <TableHead>Transaction</TableHead>
              <TableHead>Amount</TableHead>
              <TableHead>Type</TableHead>
              <TableHead>Category / Split</TableHead>
              <TableHead>Include</TableHead>
            </TableRow>
          </TableHeader>
          <TableBody>
            {rows.map((row, idx) => {
              const rowCategories = categoriesForTransactionKind(
                row.transactionKind,
                categories,
                categoryGroups,
              );
              const rowCategoryGroups = categoryGroupsForTransactionKind(
                row.transactionKind,
                categoryGroups,
              );
              const rowAccentClass =
                row.transactionKind === "income"
                  ? "border-l-2 border-l-emerald-500/70 bg-emerald-500/5"
                  : row.transactionKind === "investment"
                    ? "border-l-2 border-l-indigo-500/70 bg-indigo-500/5"
                    : "";
              const manualReviewClass = row.needsManualReview ? "bg-destructive/5" : "";
              return (
                <TableRow
                  key={row.isManual ? `manual-${idx}` : `${row.name}-${idx}`}
                  className={[rowAccentClass, manualReviewClass].filter(Boolean).join(" ") || undefined}
                >
                  <TableCell>
                    <div className="space-y-1">
                      {row.isManual ? (
                        <Input
                          value={row.name}
                          placeholder="Transaction name"
                          onChange={(event) => {
                            updateRow(idx, { name: event.target.value });
                          }}
                        />
                      ) : (
                        <p
                          className={
                            row.transactionKind === "income"
                              ? "font-medium text-emerald-200"
                              : row.transactionKind === "investment"
                                ? "font-medium text-indigo-200"
                                : undefined
                          }
                        >
                          {row.name}
                        </p>
                      )}
                      <div className="flex flex-wrap gap-1">
                        {row.date ? (
                          <Badge variant="outline" title="Transaction date">
                            {formatExpenseDate(row.date)}
                          </Badge>
                        ) : null}
                        {row.source ? (
                          <Badge variant="secondary" title="Transaction source">
                            {row.source}
                          </Badge>
                        ) : null}
                        {row.needsManualReview ? (
                          <Badge variant="destructive">Needs categorization</Badge>
                        ) : null}
                        {row.isManual ? <Badge variant="outline">Manual</Badge> : null}
                        {row.splitEnabled ? <Badge variant="secondary">Split</Badge> : null}
                        {row.isPossibleDuplicate ? (
                          <Badge variant="outline">Possible duplicate</Badge>
                        ) : null}
                      </div>
                      {row.errorReason ? (
                        <p className="text-xs text-muted-foreground">{row.errorReason}</p>
                      ) : null}
                      <div className="flex flex-wrap gap-2 pt-1">
                        <Button
                          type="button"
                          size="sm"
                          variant={row.splitEnabled ? "default" : "outline"}
                          onClick={() => {
                            setSplitEnabled(idx, !row.splitEnabled);
                          }}
                        >
                          {row.splitEnabled ? "Unsplit" : "Split"}
                        </Button>
                        {row.isManual ? (
                          <Button
                            type="button"
                            size="sm"
                            variant="ghost"
                            onClick={() => {
                              removeManualRow(idx);
                            }}
                          >
                            Remove
                          </Button>
                        ) : null}
                      </div>
                    </div>
                  </TableCell>
                  <TableCell className="w-[120px]">
                    <Input
                      type="number"
                      step="0.01"
                      min="0"
                      className={
                        row.transactionKind === "income"
                          ? "border-emerald-500/40 focus-visible:border-emerald-500"
                          : row.transactionKind === "investment"
                            ? "border-indigo-500/40 focus-visible:border-indigo-500"
                            : undefined
                      }
                      value={String(row.cost)}
                      onChange={(event) => {
                        updateRow(idx, { cost: Number(event.target.value) });
                      }}
                    />
                  </TableCell>
                  <TableCell className="w-[140px]">
                    <div className="space-y-1">
                      {row.splitEnabled ? (
                        <p className="text-xs text-muted-foreground">Default for new parts</p>
                      ) : null}
                      <Select
                        value={row.transactionKind}
                        onValueChange={(value) => {
                          if (!value) return;
                          updateTransactionKind(idx, value as TransactionKind);
                        }}
                      >
                        <SelectTrigger className="w-full">
                          <SelectValue placeholder="Type" />
                        </SelectTrigger>
                        <SelectContent>
                          {TRANSACTION_KIND_OPTIONS.map((option) => (
                            <SelectItem key={option.value} value={option.value}>
                              {option.label}
                            </SelectItem>
                          ))}
                        </SelectContent>
                      </Select>
                    </div>
                  </TableCell>
                  <TableCell className="min-w-[420px]">
                    {row.splitEnabled ? (
                      <SplitEditor
                        categories={categories}
                        categoryGroups={categoryGroups}
                        defaultKind={row.transactionKind}
                        parts={row.splitParts}
                        onChange={(splitParts) => {
                          updateRow(idx, { splitParts });
                        }}
                      />
                    ) : (
                      <Select
                        value={row.category || undefined}
                        onValueChange={(value) => {
                          updateRow(idx, { category: value ?? "" });
                        }}
                      >
                        <SelectTrigger
                          className={[
                            "h-auto w-full min-w-0 py-2 *:data-[slot=select-value]:line-clamp-none *:data-[slot=select-value]:whitespace-normal",
                            row.transactionKind === "income" ? "border-emerald-500/40" : "",
                            row.transactionKind === "investment" ? "border-indigo-500/40" : "",
                          ]
                            .filter(Boolean)
                            .join(" ")}
                        >
                          <SelectValue placeholder="Category" />
                        </SelectTrigger>
                        <SelectContent
                          alignItemWithTrigger={false}
                          className="w-max min-w-[var(--anchor-width)] max-w-[min(calc(100vw-2rem),32rem)]"
                        >
                          <CategorySelectItems
                            categories={rowCategories}
                            categoryGroups={rowCategoryGroups}
                          />
                        </SelectContent>
                      </Select>
                    )}
                  </TableCell>
                  <TableCell>
                    <input
                      checked={row.include}
                      onChange={(event) => {
                        updateRow(idx, { include: event.target.checked });
                      }}
                      type="checkbox"
                      className="size-4"
                    />
                  </TableCell>
                </TableRow>
              );
            })}
          </TableBody>
        </Table>
      </div>
      <div className="flex flex-wrap justify-between gap-2">
        <Button type="button" variant="outline" onClick={addManualExpense} disabled={busy}>
          <Plus className="mr-2 size-4" />
          Add transaction
        </Button>
        <div className="flex flex-wrap gap-2">
          <Button type="button" variant="secondary" onClick={onCancel} disabled={busy}>
            Cancel
          </Button>
          <Button
            type="button"
            disabled={busy}
            onClick={async () => {
              try {
                setError("");
                const finalizeItems = buildFinalizeItems();
                if (finalizeItems.length === 0) {
                  throw new Error("Nothing selected to include.");
                }
                await onApply(finalizeItems);
              } catch (err) {
                setError(err instanceof Error ? err.message : String(err));
              }
            }}
          >
            {busy ? "Applying..." : "Apply and Fill"}
          </Button>
        </div>
      </div>
    </div>
  );
}
