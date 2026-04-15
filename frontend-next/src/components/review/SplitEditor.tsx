"use client";

import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from "@/components/ui/select";
import { Plus, Trash2 } from "lucide-react";

export type SplitPart = {
  amount: number;
  category: string;
};

type SplitEditorProps = {
  parts: SplitPart[];
  categories: string[];
  onChange: (parts: SplitPart[]) => void;
};

export function SplitEditor({ parts, categories, onChange }: SplitEditorProps) {
  return (
    <div className="space-y-2 rounded-md border p-2">
      {parts.map((part, idx) => (
        <div key={idx} className="grid gap-2 sm:grid-cols-[100px_1fr_auto]">
          <Input
            type="number"
            step="0.01"
            value={String(part.amount)}
            onChange={(event) => {
              const next = [...parts];
              next[idx] = { ...next[idx], amount: Number(event.target.value) };
              onChange(next);
            }}
          />
          <Select
            value={part.category}
            onValueChange={(value) => {
              const next = [...parts];
              next[idx] = { ...next[idx], category: value ?? "" };
              onChange(next);
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
      ))}
      <Button
        type="button"
        variant="secondary"
        className="w-full"
        onClick={() => onChange([...parts, { amount: 0, category: categories[0] ?? "" }])}
      >
        <Plus className="mr-2 size-4" />
        Add split row
      </Button>
    </div>
  );
}
