"use client";

import { Button } from "@/components/ui/button";
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from "@/components/ui/card";
import { Input } from "@/components/ui/input";
import { AssetItem } from "@/lib/api/types";
import { money } from "@/lib/format";
import { Trash2 } from "lucide-react";
import { useState } from "react";

type AssetsPanelProps = {
  cars: AssetItem[];
  investments: AssetItem[];
  leumiBalance: number;
  onChange: (nextCars: AssetItem[], nextInvestments: AssetItem[]) => Promise<void>;
};

function AssetList({
  title,
  data,
  onAdd,
  onRemove,
}: {
  title: string;
  data: AssetItem[];
  onAdd: (item: AssetItem) => void;
  onRemove: (index: number) => void;
}) {
  const [name, setName] = useState("");
  const [value, setValue] = useState("");

  return (
    <div className="space-y-3">
      <h3 className="text-sm font-semibold">{title}</h3>
      <div className="space-y-2">
        {data.map((item, index) => (
          <div key={`${item.name}-${index}`} className="flex items-center justify-between rounded-md border p-2">
            <div>
              <p className="text-sm">{item.name}</p>
              <p className="text-xs text-muted-foreground">{money(item.value)}</p>
            </div>
            <Button type="button" variant="ghost" size="icon" onClick={() => onRemove(index)}>
              <Trash2 className="size-4" />
            </Button>
          </div>
        ))}
      </div>
      <div className="grid gap-2 sm:grid-cols-[1fr_140px_auto]">
        <Input value={name} onChange={(e) => setName(e.target.value)} placeholder="Name" />
        <Input value={value} onChange={(e) => setValue(e.target.value)} placeholder="Value" type="number" />
        <Button
          type="button"
          onClick={() => {
            const numeric = Number(value);
            if (!name.trim() || !Number.isFinite(numeric) || numeric <= 0) {
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
          onRemove={(index) => void onChange(cars.filter((_, i) => i !== index), investments)}
        />

        <AssetList
          title={`Investments (${money(invTotal)})`}
          data={investments}
          onAdd={(item) => void onChange(cars, [...investments, item])}
          onRemove={(index) => void onChange(cars, investments.filter((_, i) => i !== index))}
        />
      </CardContent>
    </Card>
  );
}
