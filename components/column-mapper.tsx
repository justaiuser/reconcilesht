"use client";

import { Button } from "@/components/ui/button";
import {
  Select,
  SelectContent,
  SelectItem,
  SelectTrigger,
  SelectValue,
} from "@/components/ui/select";
import type { Cell, ColumnMap, ParsedAct, SideColumns } from "@/lib/types";
import { headerCells, mapFromHeaderRow } from "@/lib/excel-act";
import { cellText, formatDateRu, formatMoney, lineLabel } from "@/lib/text";

type ColumnMapperProps = {
  rows: Cell[][];
  map: ColumnMap;
  act: ParsedAct;
  sheetNames: string[];
  sheetName: string;
  onSheet: (name: string) => void;
  onMap: (map: ColumnMap) => void;
};

export function ColumnMapper({
  rows,
  map,
  act,
  sheetNames,
  sheetName,
  onSheet,
  onMap,
}: ColumnMapperProps) {
  const columnCount = rows.reduce((max, row) => Math.max(max, row.length), 0);
  const header = headerCells(rows, map.headerRow);
  const options = Array.from({ length: columnCount }, (_, index) => {
    const label = cellText(header[index]) || `Столбец ${index + 1}`;
    return { index, label: `${index + 1}. ${label}` };
  });
  const headerChoices = rows
    .slice(0, 25)
    .map((row, index) => ({
      index,
      label: row.map((cell) => cellText(cell)).filter(Boolean).join(" · ") || "Пустая строка",
    }))
    .filter((row) => row.label !== "Пустая строка" || row.index === map.headerRow);

  const updateSide = (
    side: "ours" | "theirs",
    key: keyof SideColumns,
    value: number | null,
  ) => {
    const current = map[side];
    if (!current) return;
    onMap({ ...map, [side]: { ...current, [key]: value } });
  };

  return (
    <section className="min-w-0 rounded-2xl border border-border bg-card p-4 sm:p-5">
      <div className="flex flex-wrap items-end justify-between gap-3">
        <div>
          <h2 className="text-base font-semibold">Колонки Excel</h2>
          <p className="mt-1 max-w-2xl text-sm leading-6 text-muted-foreground">
            Даты и суммы должны стоять в своих колонках. Если программа выгрузила акт
            иначе, поправьте поля ниже — предпросмотр обновится сразу.
          </p>
        </div>
        {sheetNames.length > 1 ? (
          <label className="grid gap-1 text-sm">
            <span className="text-muted-foreground">Лист</span>
            <Select
              items={sheetNames.map((name) => ({ value: name, label: name }))}
              value={sheetName}
              onValueChange={(value) => {
                if (value) onSheet(value);
              }}
            >
              <SelectTrigger className="w-56">
                <SelectValue />
              </SelectTrigger>
              <SelectContent>
                {sheetNames.map((name) => (
                  <SelectItem key={name} value={name}>
                    {name}
                  </SelectItem>
                ))}
              </SelectContent>
            </Select>
          </label>
        ) : null}
      </div>

      <div className="mt-4 grid min-w-0 gap-3 sm:grid-cols-2">
        <label className="grid min-w-0 gap-1 text-sm">
          <span className="text-muted-foreground">Строка шапки</span>
          <Select
            items={headerChoices.map((row) => ({
              value: String(row.index),
              label: `${row.index + 1}. ${trim(row.label, 80)}`,
            }))}
            value={String(map.headerRow)}
            onValueChange={(value) => {
              if (!value) return;
              const headerRow = Number(value);
              onMap(
                mapFromHeaderRow(rows, headerRow) ?? {
                  ...map,
                  headerRow,
                  dataStartRow: headerRow + 1,
                },
              );
            }}
          >
            <SelectTrigger className="w-full">
              <SelectValue />
            </SelectTrigger>
            <SelectContent>
              {headerChoices.map((row) => (
                <SelectItem key={row.index} value={String(row.index)}>
                  {row.index + 1}. {trim(row.label, 80)}
                </SelectItem>
              ))}
            </SelectContent>
          </Select>
        </label>
        <div className="flex items-end">
          <Button
            type="button"
            variant="outline"
            onClick={() =>
              onMap(
                map.theirs
                  ? { ...map, theirs: null }
                  : {
                      ...map,
                      theirs: { date: 4, document: 5, debit: 6, credit: 7 },
                    },
              )
            }
          >
            {map.theirs ? "В файле одна колонка" : "В файле две колонки"}
          </Button>
        </div>
      </div>

      <SideFields
        title={map.theirs ? "Наша колонка" : "Колонки"}
        columns={map.ours}
        options={options}
        onChange={(key, value) => updateSide("ours", key, value)}
      />
      {map.theirs ? (
        <SideFields
          title="Колонка контрагента"
          columns={map.theirs}
          options={options}
          onChange={(key, value) => updateSide("theirs", key, value)}
        />
      ) : null}

      {act.warnings.map((warning) => (
        <p key={warning} className="mt-4 text-sm leading-6 text-destructive">
          {warning}
        </p>
      ))}

      <div className="mt-4 overflow-x-auto">
        <table className="w-full min-w-[40rem] text-sm">
          <caption className="mb-2 text-left text-muted-foreground">
            Разобрано строк: {act.lines.length}
            {act.period ? ` · ${act.period}` : ""}
            {act.partyA && act.partyB ? ` · ${act.partyA} и ${act.partyB}` : ""}
          </caption>
          <thead className="text-left text-muted-foreground">
            <tr className="border-b border-border">
              <th className="py-2 pr-3 font-medium">Колонка</th>
              <th className="py-2 pr-3 font-medium">Дата</th>
              <th className="py-2 pr-3 font-medium">Документ</th>
              <th className="py-2 pr-3 text-right font-medium">Дебет</th>
              <th className="py-2 text-right font-medium">Кредит</th>
            </tr>
          </thead>
          <tbody>
            {act.lines.slice(0, 8).map((line) => (
              <tr key={line.id} className="border-b border-border/70">
                <td className="py-2 pr-3 text-muted-foreground">
                  {line.side === "theirs" ? "Контрагент" : line.side === "ours" ? "Наша" : "Строка"}
                </td>
                <td className="py-2 pr-3 tabular-nums">{formatDateRu(line.date)}</td>
                <td className="py-2 pr-3">{lineLabel(line.description, line.docNumber)}</td>
                <td className="py-2 pr-3 text-right tabular-nums">{formatMoney(line.debit)}</td>
                <td className="py-2 text-right tabular-nums">{formatMoney(line.credit)}</td>
              </tr>
            ))}
          </tbody>
        </table>
        {act.lines.length > 8 ? (
          <p className="mt-2 text-sm text-muted-foreground">
            И ещё {act.lines.length - 8}. В сверку попадут все.
          </p>
        ) : null}
      </div>
    </section>
  );
}

function SideFields({
  title,
  columns,
  options,
  onChange,
}: {
  title: string;
  columns: SideColumns;
  options: { index: number; label: string }[];
  onChange: (key: keyof SideColumns, value: number | null) => void;
}) {
  const fields: { key: keyof SideColumns; label: string }[] = [
    { key: "date", label: "Дата" },
    { key: "document", label: "Документ" },
    { key: "debit", label: "Дебет" },
    { key: "credit", label: "Кредит" },
  ];
  return (
    <fieldset className="mt-4">
      <legend className="text-sm font-medium">{title}</legend>
      <div className="mt-2 grid min-w-0 gap-3 sm:grid-cols-2 lg:grid-cols-4">
        {fields.map((field) => (
          <label key={field.key} className="grid min-w-0 gap-1 text-sm">
            <span className="text-muted-foreground">{field.label}</span>
            <Select
              items={[
                { value: "none", label: "Не использовать" },
                ...options.map((option) => ({ value: String(option.index), label: option.label })),
              ]}
              value={columns[field.key] == null ? "none" : String(columns[field.key])}
              onValueChange={(value) =>
                onChange(field.key, !value || value === "none" ? null : Number(value))
              }
            >
              <SelectTrigger className="w-full">
                <SelectValue />
              </SelectTrigger>
              <SelectContent>
                <SelectItem value="none">Не использовать</SelectItem>
                {options.map((option) => (
                  <SelectItem key={option.index} value={String(option.index)}>
                    {option.label}
                  </SelectItem>
                ))}
              </SelectContent>
            </Select>
          </label>
        ))}
      </div>
    </fieldset>
  );
}

function trim(value: string, max: number): string {
  return value.length > max ? `${value.slice(0, max - 1)}…` : value;
}
