"use client";

import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { CategorySelectItems } from "@/components/review/CategorySelectItems";
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from "@/components/ui/select";
import { CategoryGroup } from "@/lib/api/types";
import {
  categoriesForTransactionKind,
  categoryGroupsForTransactionKind,
  TransactionKind,
} from "@/lib/format";
import { Plus, Trash2 } from "lucide-react";

export type SplitPart = {
  amount: number;
  transactionKind: TransactionKind;
  category: string;
};

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
) {
  const allowed = categoriesForTransactionKind(kind, categories, categoryGroups);
  if (currentCategory && allowed.includes(currentCategory)) {
    return currentCategory;
  }
  return allowed[0] ?? "";
}

type SplitEditorProps = {
  parts: SplitPart[];
  categories: string[];
  categoryGroups?: CategoryGroup[];
  defaultKind?: TransactionKind;
  onChange: (parts: SplitPart[]) => void;
};

export function SplitEditor({
  parts,
  categories,
  categoryGroups,
  defaultKind = "expense",
  onChange,
}: SplitEditorProps) {
  return (
    <div className="space-y-2 rounded-md border p-2">
      {parts.map((part, idx) => {
        const partCategories = categoriesForTransactionKind(
          part.transactionKind,
          categories,
          categoryGroups,
        );
        const partCategoryGroups = categoryGroupsForTransactionKind(
          part.transactionKind,
          categoryGroups,
        );
        return (
          <div key={idx} className="grid gap-2 sm:grid-cols-[100px_120px_1fr_auto]">
            <Input
              type="number"
              step="0.01"
              min="0"
              value={String(part.amount)}
              onChange={(event) => {
                const next = [...parts];
                next[idx] = { ...next[idx], amount: Number(event.target.value) };
                onChange(next);
              }}
            />
            <Select
              value={part.transactionKind}
              onValueChange={(value) => {
                if (!value) return;
                const kind = value as TransactionKind;
                const next = [...parts];
                next[idx] = {
                  ...next[idx],
                  transactionKind: kind,
                  category: pickCategoryForKind(kind, next[idx].category, categories, categoryGroups),
                };
                onChange(next);
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
            <Select
              value={part.category || undefined}
              onValueChange={(value) => {
                const next = [...parts];
                next[idx] = { ...next[idx], category: value ?? "" };
                onChange(next);
              }}
            >
              <SelectTrigger className="h-auto w-full min-w-0 py-2 *:data-[slot=select-value]:line-clamp-none *:data-[slot=select-value]:whitespace-normal">
                <SelectValue placeholder="Category" />
              </SelectTrigger>
              <SelectContent
                alignItemWithTrigger={false}
                className="w-max min-w-[var(--anchor-width)] max-w-[min(calc(100vw-2rem),32rem)]"
              >
                <CategorySelectItems categories={partCategories} categoryGroups={partCategoryGroups} />
              </SelectContent>
            </Select>
            <Button
              type="button"
              variant="ghost"
              size="icon"
              onClick={() => {
                if (parts.length <= 1) return;
                onChange(parts.filter((_, index) => index !== idx));
              }}
            >
              <Trash2 className="size-4" />
            </Button>
          </div>
        );
      })}
      <Button
        type="button"
        variant="secondary"
        className="w-full"
        onClick={() => {
          const category = pickCategoryForKind(defaultKind, "", categories, categoryGroups);
          onChange([
            ...parts,
            { amount: 0, transactionKind: defaultKind, category },
          ]);
        }}
      >
        <Plus className="mr-2 size-4" />
        Add split row
      </Button>
    </div>
  );
}
