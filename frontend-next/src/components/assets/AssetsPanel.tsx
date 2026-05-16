"use client";

import { Button } from "@/components/ui/button";
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from "@/components/ui/card";
import { Input } from "@/components/ui/input";
import { AssetItem } from "@/lib/api/types";
import { money } from "@/lib/format";
import { Check, Pencil, Trash2, X } from "lucide-react";
import { useState } from "react";

type AssetsPanelProps = {
  cars: AssetItem[];
  investments: AssetItem[];
  leumiBalance: number;
  onChange: (nextCars: AssetItem[], nextInvestments: AssetItem[]) => Promise<void>;
};

function parseAssetValue(raw: string) {
  const numeric = Number(raw);
  return Number.isFinite(numeric) && numeric > 0 ? numeric : null;
}

function AssetList({
  title,
  data,
  onAdd,
  onUpdate,
  onRemove,
}: {
  title: string;
  data: AssetItem[];
  onAdd: (item: AssetItem) => void;
  onUpdate: (index: number, item: AssetItem) => void;
  onRemove: (index: number) => void;
}) {
  const [name, setName] = useState("");
  const [value, setValue] = useState("");
  const [editingIndex, setEditingIndex] = useState<number | null>(null);
  const [editName, setEditName] = useState("");
  const [editValue, setEditValue] = useState("");

  function startEdit(index: number) {
    const item = data[index];
    if (!item) {
      return;
    }
    setEditingIndex(index);
    setEditName(item.name);
    setEditValue(String(item.value));
  }

  function cancelEdit() {
    setEditingIndex(null);
    setEditName("");
    setEditValue("");
  }

  function saveEdit(index: number) {
    const numeric = parseAssetValue(editValue);
    if (!editName.trim() || numeric === null) {
      return;
    }
    onUpdate(index, { name: editName.trim(), value: numeric });
    cancelEdit();
  }

  return (
    <div className="space-y-3">
      <h3 className="text-sm font-semibold">{title}</h3>
      <div className="space-y-2">
        {data.map((item, index) =>
          editingIndex === index ? (
            <div key={`${item.name}-${index}-edit`} className="space-y-2 rounded-md border p-2">
              <div className="grid gap-2 sm:grid-cols-2">
                <Input
                  value={editName}
                  onChange={(e) => setEditName(e.target.value)}
                  placeholder="Name"
                  aria-label="Asset name"
                />
                <Input
                  value={editValue}
                  onChange={(e) => setEditValue(e.target.value)}
                  placeholder="Value"
                  type="number"
                  aria-label="Asset value"
                />
              </div>
              <div className="flex justify-end gap-1">
                <Button type="button" variant="ghost" size="icon" onClick={cancelEdit} aria-label="Cancel edit">
                  <X className="size-4" />
                </Button>
                <Button
                  type="button"
                  variant="ghost"
                  size="icon"
                  onClick={() => saveEdit(index)}
                  aria-label="Save edit"
                >
                  <Check className="size-4" />
                </Button>
              </div>
            </div>
          ) : (
            <div key={`${item.name}-${index}`} className="flex items-center justify-between rounded-md border p-2">
              <div>
                <p className="text-sm">{item.name}</p>
                <p className="text-xs text-muted-foreground">{money(item.value)}</p>
              </div>
              <div className="flex items-center">
                <Button
                  type="button"
                  variant="ghost"
                  size="icon"
                  onClick={() => startEdit(index)}
                  aria-label={`Edit ${item.name}`}
                >
                  <Pencil className="size-4" />
                </Button>
                <Button
                  type="button"
                  variant="ghost"
                  size="icon"
                  onClick={() => onRemove(index)}
                  aria-label={`Remove ${item.name}`}
                >
                  <Trash2 className="size-4" />
                </Button>
              </div>
            </div>
          ),
        )}
      </div>
      <div className="grid gap-2 sm:grid-cols-[1fr_140px_auto]">
        <Input value={name} onChange={(e) => setName(e.target.value)} placeholder="Name" />
        <Input value={value} onChange={(e) => setValue(e.target.value)} placeholder="Value" type="number" />
        <Button
          type="button"
          onClick={() => {
            const numeric = parseAssetValue(value);
            if (!name.trim() || numeric === null) {
              return;
            }
            onAdd({ name: name.trim(), value: numeric });
            setName("");
            setValue("");
          }}
        >
          Add
        </Button>
      </div>
    </div>
  );
}

export function AssetsPanel({ cars, investments, leumiBalance, onChange }: AssetsPanelProps) {
  const carsTotal = cars.reduce((acc, item) => acc + Number(item.value || 0), 0);
  const invTotal = investments.reduce((acc, item) => acc + Number(item.value || 0), 0);
  const total = leumiBalance + carsTotal + invTotal;

  function updateAtIndex(items: AssetItem[], index: number, next: AssetItem) {
    return items.map((item, i) => (i === index ? next : item));
  }

  return (
    <Card>
      <CardHeader>
        <CardTitle>Assets & Net Worth</CardTitle>
        <CardDescription>Manage tracked assets used in your net-worth overview.</CardDescription>
      </CardHeader>
      <CardContent className="space-y-5">
        <div className="grid gap-3 text-sm sm:grid-cols-2">
          <div className="rounded-md border p-3">
            <p className="text-muted-foreground">Bank Balance</p>
            <p className="text-xl font-semibold">{money(leumiBalance)}</p>
          </div>
          <div className="rounded-md border p-3">
            <p className="text-muted-foreground">Net Worth</p>
            <p className="text-xl font-semibold">{money(total)}</p>
          </div>
        </div>

        <AssetList
          title={`Cars (${money(carsTotal)})`}
          data={cars}
          onAdd={(item) => void onChange([...cars, item], investments)}
          onUpdate={(index, item) => void onChange(updateAtIndex(cars, index, item), investments)}
          onRemove={(index) => void onChange(cars.filter((_, i) => i !== index), investments)}
        />

        <AssetList
          title={`Investments (${money(invTotal)})`}
          data={investments}
          onAdd={(item) => void onChange(cars, [...investments, item])}
          onUpdate={(index, item) => void onChange(cars, updateAtIndex(investments, index, item))}
          onRemove={(index) => void onChange(cars, investments.filter((_, i) => i !== index))}
        />
      </CardContent>
    </Card>
  );
}
