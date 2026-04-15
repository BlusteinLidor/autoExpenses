"use client";

import { SplitEditor, SplitPart } from "@/components/review/SplitEditor";
import { Badge } from "@/components/ui/badge";
import { Button } from "@/components/ui/button";
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from "@/components/ui/select";
import { Table, TableBody, TableCell, TableHead, TableHeader, TableRow } from "@/components/ui/table";
import { FinalizeRunItem, ReviewItem } from "@/lib/api/types";
import { isSplitExpenseName, money } from "@/lib/format";
import { useMemo, useState } from "react";

type RowState = {
  include: boolean;
  category: string;
  splitParts: SplitPart[];
};

type ReviewTableProps = {
  items: ReviewItem[];
  categories: string[];
  onApply: (items: FinalizeRunItem[]) => Promise<void>;
  onCancel: () => void;
  busy: boolean;
};

export function ReviewTable({ items, categories, onApply, onCancel, busy }: ReviewTableProps) {
  const [rows, setRows] = useState<RowState[]>(
    items.map((item) => ({
      include: !item.is_possible_duplicate,
      category: item.category || categories[0] || "",
      splitParts: [
        { amount: Math.abs(Number(item.cost || 0)), category: item.category || categories[0] || "" },
        { amount: 0, category: item.category || categories[0] || "" },
      ],
    })),
  );
  const [error, setError] = useState("");

  const hasSplitRows = useMemo(() => items.some((item) => isSplitExpenseName(item.name)), [items]);

  const buildFinalizeItems = () => {
    const finalizeItems: FinalizeRunItem[] = [];
    for (let idx = 0; idx < items.length; idx += 1) {
      const source = items[idx];
      const row = rows[idx];
      if (!row?.include) continue;
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
        if (Math.abs(total - Math.abs(Number(source.cost || 0))) > 0.01) {
          throw new Error(`Split total for '${source.name}' must equal the original amount.`);
        }
      } else {
        finalizeItems.push({
          name: source.name,
          cost: Math.abs(Number(source.cost || 0)),
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
          <h2 className="text-lg font-semibold">Review Categories</h2>
          <p className="text-sm text-muted-foreground">{items.length} items ready for review.</p>
        </div>
        {hasSplitRows ? <Badge variant="secondary">Split flow enabled</Badge> : null}
      </div>
      {error ? <p className="text-sm text-destructive">{error}</p> : null}

      <div className="overflow-auto rounded-md border">
        <Table>
          <TableHeader>
            <TableRow>
              <TableHead>Expense</TableHead>
              <TableHead>Amount</TableHead>
              <TableHead>Category / Split</TableHead>
              <TableHead>Include</TableHead>
            </TableRow>
          </TableHeader>
          <TableBody>
            {items.map((item, idx) => {
              const row = rows[idx];
              return (
                <TableRow key={`${item.name}-${idx}`}>
                  <TableCell>
                    <div className="space-y-1">
                      <p>{item.name}</p>
                      {item.is_possible_duplicate ? <Badge variant="outline">Possible duplicate</Badge> : null}
                    </div>
                  </TableCell>
                  <TableCell>{money(item.cost)}</TableCell>
                  <TableCell className="min-w-[280px]">
                    {isSplitExpenseName(item.name) ? (
                      <SplitEditor
                        categories={categories}
                        parts={row.splitParts}
                        onChange={(splitParts) => {
                          const nextRows = [...rows];
                          nextRows[idx] = { ...nextRows[idx], splitParts };
                          setRows(nextRows);
                        }}
                      />
                    ) : (
                      <Select
                        value={row.category}
                        onValueChange={(value) => {
                          const nextRows = [...rows];
                          nextRows[idx] = { ...nextRows[idx], category: value ?? "" };
                          setRows(nextRows);
                        }}
                      >
                        <SelectTrigger>
                          <SelectValue placeholder="Category" />
                        </SelectTrigger>
                        <SelectContent>
                          {categories.map((category) => (
                            <SelectItem key={category} value={category}>
                              {category}
                            </SelectItem>
                          ))}
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
