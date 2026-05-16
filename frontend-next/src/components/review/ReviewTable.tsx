"use client";

import { CategorySelectItems } from "@/components/review/CategorySelectItems";
import { SplitEditor, SplitPart } from "@/components/review/SplitEditor";
import { Alert, AlertDescription, AlertTitle } from "@/components/ui/alert";
import { Badge } from "@/components/ui/badge";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Select, SelectContent, SelectTrigger, SelectValue } from "@/components/ui/select";
import { Table, TableBody, TableCell, TableHead, TableHeader, TableRow } from "@/components/ui/table";
import { CategoryGroup, FinalizeRunItem, ReviewItem } from "@/lib/api/types";
import { expenseCategories, incomeCategories, isSplitExpenseName } from "@/lib/format";
import { useMemo, useState } from "react";

const DEFAULT_INCOME_CATEGORY = "הכנסה אחרת / חד פעמית";

function categoriesForItem(
  item: ReviewItem,
  categories: string[],
  categoryGroups?: CategoryGroup[],
) {
  if (!item.is_income) {
    return expenseCategories(categories, categoryGroups);
  }
  const incomeOnly = incomeCategories(categoryGroups);
  return incomeOnly.length > 0 ? incomeOnly : categories;
}

function defaultCategoryForItem(
  item: ReviewItem,
  categories: string[],
  categoryGroups?: CategoryGroup[],
) {
  const allowed = categoriesForItem(item, categories, categoryGroups);
  if (item.category && allowed.includes(item.category)) {
    return item.category;
  }
  if (item.is_income) {
    return allowed.includes(DEFAULT_INCOME_CATEGORY)
      ? DEFAULT_INCOME_CATEGORY
      : allowed[0] || "";
  }
  return item.needs_manual_review ? "" : allowed[0] || "";
}

type RowState = {
  include: boolean;
  cost: number;
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
      const defaultCategory = defaultCategoryForItem(item, categories, categoryGroups);
      return {
        include: item.needs_manual_review ? true : !item.is_possible_duplicate,
        cost: Math.abs(Number(item.cost || 0)),
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
  const incomeCount = useMemo(() => items.filter((item) => item.is_income).length, [items]);
  const expenseCount = items.length - incomeCount;

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
            Edit amounts and categories before applying. {expenseCount} expenses
            {incomeCount > 0 ? `, ${incomeCount} income` : ""} ready for review.
          </p>
        </div>
        <div className="flex flex-wrap gap-2">
          {incomeCount > 0 ? (
            <Badge className="border-emerald-500/40 bg-emerald-500/15 text-emerald-200">
              {incomeCount} income
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
              <TableHead>Category / Split</TableHead>
              <TableHead>Include</TableHead>
            </TableRow>
          </TableHeader>
          <TableBody>
            {items.map((item, idx) => {
              const row = rows[idx];
              const rowCategories = categoriesForItem(item, categories, categoryGroups);
              const incomeRowClass = item.is_income
                ? "border-l-2 border-l-emerald-500/70 bg-emerald-500/5"
                : "";
              const manualReviewClass = item.needs_manual_review ? "bg-destructive/5" : "";
              return (
                <TableRow
                  key={`${item.name}-${idx}`}
                  className={[incomeRowClass, manualReviewClass].filter(Boolean).join(" ") || undefined}
                >
                  <TableCell>
                    <div className="space-y-1">
                      <p className={item.is_income ? "font-medium text-emerald-200" : undefined}>
                        {item.name}
                      </p>
                      {item.is_income ? (
                        <Badge className="border-emerald-500/40 bg-emerald-500/15 text-emerald-200">
                          Income
                        </Badge>
                      ) : null}
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
                      className={item.is_income ? "border-emerald-500/40 focus-visible:border-emerald-500" : undefined}
                      value={String(row.cost)}
                      onChange={(event) => {
                        const nextRows = [...rows];
                        nextRows[idx] = { ...nextRows[idx], cost: Number(event.target.value) };
                        setRows(nextRows);
                      }}
                    />
                  </TableCell>
                  <TableCell className="min-w-[360px]">
                    {isSplitExpenseName(item.name) ? (
                      <SplitEditor
                        categories={rowCategories}
                        categoryGroups={
                          item.is_income
                            ? categoryGroups?.filter((group) => group.id === "income")
                            : categoryGroups?.filter((group) => group.id !== "income")
                        }
                        parts={row.splitParts}
                        onChange={(splitParts) => {
                          const nextRows = [...rows];
                          nextRows[idx] = { ...nextRows[idx], splitParts };
                          setRows(nextRows);
                        }}
                      />
                    ) : (
                      <Select
                        value={row.category || undefined}
                        onValueChange={(value) => {
                          const nextRows = [...rows];
                          nextRows[idx] = { ...nextRows[idx], category: value ?? "" };
                          setRows(nextRows);
                        }}
                      >
                        <SelectTrigger
                          className={[
                            "h-auto w-full min-w-0 py-2 *:data-[slot=select-value]:line-clamp-none *:data-[slot=select-value]:whitespace-normal",
                            item.is_income ? "border-emerald-500/40" : "",
                          ]
                            .filter(Boolean)
                            .join(" ")}
                        >
                          <SelectValue placeholder={item.is_income ? "Income category" : "Category"} />
                        </SelectTrigger>
                        <SelectContent
                          alignItemWithTrigger={false}
                          className="w-max min-w-[var(--anchor-width)] max-w-[min(calc(100vw-2rem),32rem)]"
                        >
                          <CategorySelectItems
                            categories={rowCategories}
                            categoryGroups={
                              item.is_income
                                ? categoryGroups?.filter((group) => group.id === "income")
                                : categoryGroups?.filter((group) => group.id !== "income")
                            }
                          />
                        </SelectContent>
                      </Select>
                    )}
                  </TableCell>
                  <TableCell>
                    <input
                      checked={row.include}
                      onChange={(event) => {
                        const nextRows = [...rows];
                        nextRows[idx] = { ...nextRows[idx], include: event.target.checked };
                        setRows(nextRows);
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
