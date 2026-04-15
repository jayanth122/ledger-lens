"use client";

import { useState, useRef, type DragEvent, type ChangeEvent } from "react";
import * as XLSX from "xlsx";
import {
  AreaChart,
  Area,
  BarChart,
  Bar,
  XAxis,
  YAxis,
  CartesianGrid,
  Tooltip,
  ResponsiveContainer,
  ReferenceLine,
  Cell,
} from "recharts";

// ── types ─────────────────────────────────────────────────────────────────────

interface CategoryAmount {
  name: string;
  amount: number;
}

interface MonthlyPoint {
  month: string;
  monthKey: string;
  income: number;
  expense: number;
  netSavings: number;
  cumSavings: number;
  expenseCategories: CategoryAmount[];
  incomeCategories: CategoryAmount[];
}

interface TooltipEntry {
  payload: MonthlyPoint;
  value: number;
}

interface ChartTooltipProps {
  active?: boolean;
  payload?: TooltipEntry[];
  label?: string;
}

interface BarTooltipEntry {
  payload: CategoryAmount;
  value: number;
}

interface BarTooltipProps {
  active?: boolean;
  payload?: BarTooltipEntry[];
  label?: string;
}

// ── helpers ───────────────────────────────────────────────────────────────────

function toMonthKey(d: Date): string {
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}`;
}

function toMonthLabel(d: Date): string {
  return d.toLocaleDateString("en-CA", { year: "numeric", month: "short" });
}

function parseDate(val: unknown): Date | null {
  if (val instanceof Date) return isNaN(val.getTime()) ? null : val;
  if (typeof val === "number") {
    const d = new Date(Math.round((val - 25569) * 86_400_000));
    return isNaN(d.getTime()) ? null : d;
  }
  if (typeof val === "string" && val.trim()) {
    const d = new Date(val);
    return isNaN(d.getTime()) ? null : d;
  }
  return null;
}

// ── core computation ──────────────────────────────────────────────────────────

function computeSavings(file: File): Promise<MonthlyPoint[]> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onerror = () => reject(new Error("Could not read file"));
    reader.onload = (e) => {
      try {
        const wb = XLSX.read(e.target!.result as ArrayBuffer, {
          type: "array",
          cellDates: true,
        });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const raw = XLSX.utils.sheet_to_json(ws, {
          header: 1,
          raw: true,
        }) as unknown[][];

        if (!raw.length) throw new Error("Spreadsheet is empty");

        const headers = (raw[0] as unknown[]).map((h) =>
          h != null ? String(h).trim() : ""
        );

        const idxPeriod = headers.indexOf("Period");
        const idxAmount = headers.indexOf("Amount");
        const idxCategory = headers.indexOf("Category");

        let idxIE = headers.indexOf("IncomeExpense");
        if (idxIE === -1) {
          idxIE = headers.findIndex(
            (h) =>
              h.toLowerCase().includes("income") &&
              h.toLowerCase().includes("expense")
          );
        }

        if (idxPeriod === -1) throw new Error("'Period' column not found");
        if (idxAmount === -1) throw new Error("'Amount' column not found");
        if (idxIE === -1) throw new Error("'IncomeExpense' column not found");

        const monthly = new Map<
          string,
          {
            label: string;
            income: number;
            expense: number;
            expCats: Map<string, number>;
            incCats: Map<string, number>;
          }
        >();

        for (let i = 1; i < raw.length; i++) {
          const row = raw[i];
          if (!row || row.every((v) => v == null)) continue;
          if (typeof row[0] === "string" && row[0].startsWith("TITLE ")) continue;

          const date = parseDate(row[idxPeriod]);
          if (!date) continue;

          const rawAmt = row[idxAmount];
          const amount =
            typeof rawAmt === "number"
              ? rawAmt
              : parseFloat(String(rawAmt ?? "").replace(/[^0-9.-]/g, ""));
          if (isNaN(amount)) continue;

          const ie = String(row[idxIE] ?? "").trim();
          if (/transfer/i.test(ie)) continue;

          const category =
            idxCategory !== -1 && row[idxCategory] != null
              ? String(row[idxCategory]).trim() || "Uncategorized"
              : "Uncategorized";

          const key = toMonthKey(date);
          if (!monthly.has(key)) {
            monthly.set(key, {
              label: toMonthLabel(date),
              income: 0,
              expense: 0,
              expCats: new Map(),
              incCats: new Map(),
            });
          }
          const m = monthly.get(key)!;

          if (/income/i.test(ie)) {
            m.income += amount;
            m.incCats.set(category, (m.incCats.get(category) ?? 0) + amount);
          } else if (/exp/i.test(ie)) {
            m.expense += amount;
            m.expCats.set(category, (m.expCats.get(category) ?? 0) + amount);
          }
        }

        const sorted = Array.from(monthly.entries()).sort(([a], [b]) =>
          a.localeCompare(b)
        );
        let cum = 0;
        resolve(
          sorted.map(([key, m]) => {
            const net = m.income - m.expense;
            cum += net;
            return {
              month: m.label,
              monthKey: key,
              income: m.income,
              expense: m.expense,
              netSavings: net,
              cumSavings: cum,
              expenseCategories: Array.from(m.expCats.entries())
                .map(([name, amount]) => ({ name, amount }))
                .sort((a, b) => b.amount - a.amount),
              incomeCategories: Array.from(m.incCats.entries())
                .map(([name, amount]) => ({ name, amount }))
                .sort((a, b) => b.amount - a.amount),
            };
          })
        );
      } catch (err) {
        reject(err instanceof Error ? err : new Error(String(err)));
      }
    };
    reader.readAsArrayBuffer(file);
  });
}

// ── formatting ────────────────────────────────────────────────────────────────

const cad = (v: number) =>
  v.toLocaleString("en-CA", {
    style: "currency",
    currency: "CAD",
    maximumFractionDigits: 0,
  });

const BAR_COLORS = [
  "#2563eb", "#0891b2", "#7c3aed", "#db2777", "#ea580c",
  "#ca8a04", "#16a34a", "#0f766e", "#9333ea", "#dc2626",
];

// ── tooltips ──────────────────────────────────────────────────────────────────

function AreaTooltip({ active, payload, label }: ChartTooltipProps) {
  if (!active || !payload?.length) return null;
  const d = payload[0].payload;
  return (
    <div className="rounded-lg border border-gray-200 bg-white px-3 py-2 shadow-md text-sm">
      <p className="font-semibold text-gray-800 mb-1">{label}</p>
      <p className="text-blue-600">Cumulative: {cad(d.cumSavings)}</p>
      <p className={d.netSavings >= 0 ? "text-green-600" : "text-red-500"}>
        Net this month: {cad(d.netSavings)}
      </p>
    </div>
  );
}

function CategoryTooltip({ active, payload, label }: BarTooltipProps) {
  if (!active || !payload?.length) return null;
  return (
    <div className="rounded-lg border border-gray-200 bg-white px-3 py-2 shadow-md text-sm">
      <p className="font-semibold text-gray-800 mb-1">{label}</p>
      <p className="text-gray-700">{cad(payload[0].value)}</p>
    </div>
  );
}

// ── page ──────────────────────────────────────────────────────────────────────

export default function Home() {
  const [data, setData] = useState<MonthlyPoint[] | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [fileName, setFileName] = useState<string | null>(null);
  const [dragging, setDragging] = useState(false);
  const [selectedMonthKey, setSelectedMonthKey] = useState<string>("");
  const inputRef = useRef<HTMLInputElement>(null);

  async function handleFile(f: File) {
    if (!/\.(xlsx|xls)$/i.test(f.name)) {
      setError("Please upload an .xlsx or .xls file");
      return;
    }
    setError(null);
    setFileName(f.name);
    try {
      const pts = await computeSavings(f);
      if (!pts.length) {
        setError(
          "No data found — make sure your file has Period, Amount, and IncomeExpense columns"
        );
        return;
      }
      setData(pts);
      setSelectedMonthKey(pts[pts.length - 1].monthKey);
    } catch (err) {
      setError(err instanceof Error ? err.message : "Failed to parse file");
    }
  }

  function onFilePick(e: ChangeEvent<HTMLInputElement>) {
    if (e.target.files?.[0]) handleFile(e.target.files[0]);
  }

  function onDrop(e: DragEvent<HTMLDivElement>) {
    e.preventDefault();
    setDragging(false);
    if (e.dataTransfer.files?.[0]) handleFile(e.dataTransfer.files[0]);
  }

  const hasNeg = data?.some((d) => d.cumSavings < 0) ?? false;
  const selected = data?.find((d) => d.monthKey === selectedMonthKey) ?? null;

  return (
    <main className="min-h-screen bg-gray-50 px-4 py-12 flex flex-col items-center">
      <div className="w-full max-w-4xl">
        <h1 className="text-2xl font-bold text-gray-900 mb-1">Savings Over Time</h1>
        <p className="text-sm text-gray-500 mb-8">
          Upload your finance spreadsheet to see cumulative net savings by month.
          Transfers excluded. Formula: income&nbsp;&minus;&nbsp;expenses.
        </p>

        {/* Upload area */}
        <div
          role="button"
          tabIndex={0}
          aria-label="Upload Excel file"
          onClick={() => inputRef.current?.click()}
          onKeyDown={(e) => {
            if (e.key === "Enter" || e.key === " ") inputRef.current?.click();
          }}
          onDragOver={(e) => { e.preventDefault(); setDragging(true); }}
          onDragLeave={() => setDragging(false)}
          onDrop={onDrop}
          className={`cursor-pointer rounded-xl border-2 border-dashed p-10 text-center transition-colors ${
            dragging
              ? "border-blue-500 bg-blue-50"
              : "border-gray-300 bg-white hover:border-blue-400"
          }`}
        >
          <input
            ref={inputRef}
            type="file"
            accept=".xlsx,.xls"
            className="hidden"
            onChange={onFilePick}
          />
          <svg
            className="mx-auto mb-3 h-10 w-10 text-gray-400"
            fill="none"
            stroke="currentColor"
            strokeWidth={1.5}
            viewBox="0 0 24 24"
          >
            <path
              strokeLinecap="round"
              strokeLinejoin="round"
              d="M3 16.5v2.25A2.25 2.25 0 0 0 5.25 21h13.5A2.25 2.25 0 0 0 21 18.75V16.5m-13.5-9L12 3m0 0 4.5 4.5M12 3v13.5"
            />
          </svg>
          <p className="font-medium text-gray-700">
            {fileName ?? "Drop your Excel file here, or click to browse"}
          </p>
          <p className="mt-1 text-xs text-gray-400">.xlsx or .xls</p>
        </div>

        {error && (
          <div className="mt-4 rounded-lg border border-red-200 bg-red-50 px-4 py-3 text-sm text-red-700">
            {error}
          </div>
        )}

        {data && (
          <div className="mt-10 space-y-6">

            {/* ── Savings chart ── */}
            <div className="rounded-xl border bg-white p-6 shadow-sm">
              <h2 className="text-base font-semibold text-gray-900">Savings Over Time</h2>
              <p className="mb-6 mt-0.5 text-xs text-gray-400">
                {fileName} &middot; cumulative net savings by month
              </p>
              <ResponsiveContainer width="100%" height={340}>
                <AreaChart data={data} margin={{ top: 8, right: 16, left: 0, bottom: 0 }}>
                  <defs>
                    <linearGradient id="grad" x1="0" y1="0" x2="0" y2="1">
                      <stop offset="5%" stopColor="#2563eb" stopOpacity={0.18} />
                      <stop offset="95%" stopColor="#2563eb" stopOpacity={0} />
                    </linearGradient>
                  </defs>
                  <CartesianGrid strokeDasharray="3 3" stroke="#f0f0f0" />
                  <XAxis dataKey="month" tick={{ fontSize: 12 }} />
                  <YAxis
                    tickFormatter={(v: number) => `$${(v / 1000).toFixed(0)}k`}
                    tick={{ fontSize: 12 }}
                  />
                  {hasNeg && <ReferenceLine y={0} stroke="#ef4444" strokeDasharray="4 4" />}
                  <Tooltip content={<AreaTooltip />} />
                  <Area
                    type="monotone"
                    dataKey="cumSavings"
                    name="Cumulative savings"
                    stroke="#2563eb"
                    strokeWidth={2.5}
                    fill="url(#grad)"
                    dot={{ r: 4, fill: "#2563eb", strokeWidth: 0 }}
                    activeDot={{ r: 6 }}
                  />
                </AreaChart>
              </ResponsiveContainer>
            </div>

            {/* ── Month detail card ── */}
            <div className="rounded-xl border bg-white p-6 shadow-sm">
              <div className="flex flex-wrap items-center justify-between gap-3 mb-6">
                <div>
                  <h2 className="text-base font-semibold text-gray-900">Monthly Breakdown</h2>
                  <p className="mt-0.5 text-xs text-gray-400">
                    Income vs expenses and spending by category
                  </p>
                </div>
                <select
                  value={selectedMonthKey}
                  onChange={(e) => setSelectedMonthKey(e.target.value)}
                  className="rounded-lg border border-gray-300 bg-white px-3 py-1.5 text-sm text-gray-700 shadow-sm focus:outline-none focus:ring-2 focus:ring-blue-400"
                >
                  {[...data].reverse().map((d) => (
                    <option key={d.monthKey} value={d.monthKey}>
                      {d.month}
                    </option>
                  ))}
                </select>
              </div>

              {selected && (
                <>
                  {/* Summary chips */}
                  <div className="grid grid-cols-3 gap-3 mb-6">
                    <div className="rounded-lg bg-green-50 border border-green-100 px-4 py-3">
                      <p className="text-xs text-green-600 font-medium">Income</p>
                      <p className="mt-0.5 text-lg font-bold text-green-700">{cad(selected.income)}</p>
                    </div>
                    <div className="rounded-lg bg-red-50 border border-red-100 px-4 py-3">
                      <p className="text-xs text-red-500 font-medium">Expenses</p>
                      <p className="mt-0.5 text-lg font-bold text-red-600">{cad(selected.expense)}</p>
                    </div>
                    <div className={`rounded-lg border px-4 py-3 ${selected.netSavings >= 0 ? "bg-blue-50 border-blue-100" : "bg-orange-50 border-orange-100"}`}>
                      <p className={`text-xs font-medium ${selected.netSavings >= 0 ? "text-blue-600" : "text-orange-500"}`}>
                        Net savings
                      </p>
                      <p className={`mt-0.5 text-lg font-bold ${selected.netSavings >= 0 ? "text-blue-700" : "text-orange-600"}`}>
                        {cad(selected.netSavings)}
                      </p>
                    </div>
                  </div>

                  {/* Expense categories chart */}
                  {selected.expenseCategories.length > 0 && (
                    <div className="mb-6">
                      <h3 className="text-sm font-semibold text-gray-700 mb-3">Spending by Category</h3>
                      <ResponsiveContainer width="100%" height={selected.expenseCategories.length * 36 + 20}>
                        <BarChart
                          data={selected.expenseCategories}
                          layout="vertical"
                          margin={{ top: 0, right: 60, left: 0, bottom: 0 }}
                        >
                          <CartesianGrid strokeDasharray="3 3" stroke="#f0f0f0" horizontal={false} />
                          <XAxis
                            type="number"
                            tickFormatter={(v: number) => `$${(v / 1000).toFixed(0)}k`}
                            tick={{ fontSize: 11 }}
                          />
                          <YAxis type="category" dataKey="name" width={130} tick={{ fontSize: 12 }} />
                          <Tooltip content={<CategoryTooltip />} />
                          <Bar dataKey="amount" radius={[0, 4, 4, 0]}>
                            {selected.expenseCategories.map((_, i) => (
                              <Cell key={i} fill={BAR_COLORS[i % BAR_COLORS.length]} />
                            ))}
                          </Bar>
                        </BarChart>
                      </ResponsiveContainer>
                    </div>
                  )}

                  {/* Income sources */}
                  {selected.incomeCategories.length > 0 && (
                    <div>
                      <h3 className="text-sm font-semibold text-gray-700 mb-3">Income Sources</h3>
                      <div className="overflow-hidden rounded-lg border border-gray-100">
                        <table className="w-full text-sm">
                          <tbody className="divide-y divide-gray-100">
                            {selected.incomeCategories.map((c, i) => (
                              <tr key={i} className="flex items-center justify-between px-4 py-2.5 hover:bg-gray-50">
                                <td className="text-gray-700">{c.name}</td>
                                <td className="font-mono text-green-600">{cad(c.amount)}</td>
                              </tr>
                            ))}
                          </tbody>
                        </table>
                      </div>
                    </div>
                  )}
                </>
              )}
            </div>

            {/* ── Monthly summary table (latest → oldest) ── */}
            <div className="overflow-hidden rounded-xl border bg-white shadow-sm">
              <div className="px-4 py-3 border-b bg-gray-50">
                <p className="text-xs text-gray-400 mt-0.5">Click a row to see its breakdown above</p>
              </div>
              <table className="w-full text-sm">
                <thead className="border-b bg-gray-50">
                  <tr>
                    <th className="px-4 py-3 text-left font-medium text-gray-600">Month</th>
                    <th className="px-4 py-3 text-right font-medium text-gray-600">Income</th>
                    <th className="px-4 py-3 text-right font-medium text-gray-600">Expenses</th>
                    <th className="px-4 py-3 text-right font-medium text-gray-600">Net savings</th>
                    <th className="px-4 py-3 text-right font-medium text-gray-600">Cumulative</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-gray-100">
                  {[...data].reverse().map((row) => (
                    <tr
                      key={row.monthKey}
                      onClick={() => setSelectedMonthKey(row.monthKey)}
                      className={`cursor-pointer hover:bg-gray-50 ${row.monthKey === selectedMonthKey ? "bg-blue-50" : ""}`}
                    >
                      <td className="px-4 py-2.5 text-gray-700">{row.month}</td>
                      <td className="px-4 py-2.5 text-right font-mono text-green-600">{cad(row.income)}</td>
                      <td className="px-4 py-2.5 text-right font-mono text-red-500">{cad(row.expense)}</td>
                      <td className={`px-4 py-2.5 text-right font-mono ${row.netSavings >= 0 ? "text-green-600" : "text-red-500"}`}>
                        {cad(row.netSavings)}
                      </td>
                      <td className={`px-4 py-2.5 text-right font-mono ${row.cumSavings >= 0 ? "text-blue-700" : "text-red-500"}`}>
                        {cad(row.cumSavings)}
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>

          </div>
        )}
      </div>
    </main>
  );
}
