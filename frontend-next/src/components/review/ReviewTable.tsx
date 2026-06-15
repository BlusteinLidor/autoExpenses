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
  inferTransactionKind,
  isSplitExpenseName,
  TransactionKind,
} from "@/lib/format";
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
  cost: number;
  transactionKind: TransactionKind;
  category: string;
  splitParts: SplitPart[];
};

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
  const [rows, setRows] = useState<RowState[]>(
    items.map((item) => {
      const transactionKind = inferTransactionKind(item, categoryGroups);
      const defaultCategory = defaultCategoryForTransactionKind(
        transactionKind,
        item,
        categories,
        categoryGroups,
      );
      return {
        include: item.needs_manual_review ? true : !item.is_possible_duplicate,
        cost: Math.abs(Number(item.cost || 0)),
        transactionKind,
        category: defaultCategory,
        splitParts: [
          { amount: Math.abs(Number(item.cost || 0)), category: defaultCategory },
          { amount: 0, category: defaultCategory },
        ],
      };
    }),
  );
  const [error, setError] = useState("");

  const hasSplitRows = useMemo(() => items.some((item) => isSplitExpenseName(item.name)), [items]);
  const manualReviewCount = useMemo(
    () => items.filter((item) => item.needs_manual_review).length,
    [items],
  );
  const kindCounts = useMemo(
    () =>
      rows.reduce(
        (acc, row) => {
          acc[row.transactionKind] += 1;
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
    const source = items[idx];
    const row = rows[idx];
    const nextCategory = pickCategoryForKind(
      kind,
      row.category,
      categories,
      categoryGroups,
      source.needs_manual_review,
    );
    updateRow(idx, {
      transactionKind: kind,
      category: nextCategory,
      splitParts: row.splitParts.map((part) => ({
        ...part,
        category: pickCategoryForKind(kind, part.category, categories, categoryGroups),
      })),
    });
  };

  const buildFinalizeItems = () => {
    const finalizeItems: FinalizeRunItem[] = [];
    for (let idx = 0; idx < items.length; idx += 1) {
      const source = items[idx];
      const row = rows[idx];
      if (!row?.include) continue;
      if (source.needs_manual_review && !row.category) {
        throw new Error(`Choose a category for '${source.name}' before applying.`);
      }
      if (isSplitExpenseName(source.name)) {
        let total = 0;
        let partNo = 1;
        for (const part of row.splitParts) {
          const amount = Math.abs(Number(part.amount || 0));
          if (amount <= 0) {
            partNo += 1;
            continue;
          }
          finalizeItems.push({
            name: `${source.name} חלק ${partNo}`,
            cost: amount,
            category: part.category,
          });
          total += amount;
          partNo += 1;
        }
        const expectedTotal = Math.abs(Number(row.cost || 0));
        if (Math.abs(total - expectedTotal) > 0.01) {
          throw new Error(`Split total for '${source.name}' must equal ${expectedTotal}.`);
        }
      } else {
        const cost = Math.abs(Number(row.cost || 0));
        if (cost <= 0) {
          throw new Error(`Amount for '${source.name}' must be greater than zero.`);
        }
        finalizeItems.push({
          name: source.name,
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
            Edit amounts, transaction type, and categories before applying. {kindCounts.expense}{" "}
            expenses
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
          {hasSplitRows ? <Badge variant="secondary">Split flow enabled</Badge> : null}
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
            {items.map((item, idx) => {
              const row = rows[idx];
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
              const manualReviewClass = item.needs_manual_review ? "bg-destructive/5" : "";
              return (
                <TableRow
                  key={`${item.name}-${idx}`}
                  className={[rowAccentClass, manualReviewClass].filter(Boolean).join(" ") || undefined}
                >
                  <TableCell>
                    <div className="space-y-1">
                      <p
                        className={
                          row.transactionKind === "income"
                            ? "font-medium text-emerald-200"
                            : row.transactionKind === "investment"
                              ? "font-medium text-indigo-200"
                              : undefined
                        }
                      >
                        {item.name}
                      </p>
                      {item.needs_manual_review ? (
                        <Badge variant="destructive">Needs categorization</Badge>
                      ) : null}
                      {item.is_possible_duplicate ? <Badge variant="outline">Possible duplicate</Badge> : null}
                      {item.error_reason ? (
                        <p className="text-xs text-muted-foreground">{item.error_reason}</p>
                      ) : null}
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
                  </TableCell>
                  <TableCell className="min-w-[360px]">
                    {isSplitExpenseName(item.name) ? (
                      <SplitEditor
                        categories={rowCategories}
                        categoryGroups={rowCategoryGroups}
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
      <div className="flex flex-wrap justify-end gap-2">
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
  );
}
