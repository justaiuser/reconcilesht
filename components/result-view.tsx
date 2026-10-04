"use client";

import { useState, type ReactNode } from "react";
import { Download } from "lucide-react";
import { Button } from "@/components/ui/button";
import type { CompareLabels, CompareMode, CompareResult, MirrorMode, Pair, SideReport } from "@/lib/types";
import { problemCount } from "@/lib/compare";
import { buildAnalysis } from "@/lib/quality";
import { downloadReport } from "@/lib/report";
import { formatDateRu, formatMoney, lineLabel } from "@/lib/text";
import { cn } from "@/lib/utils";

type ResultViewProps = {
  result: CompareResult;
  mode: CompareMode;
  mirror: MirrorMode;
  rawText: string;
  confidence: number | null;
  warnings: string[];
  onMode: (mode: CompareMode) => void;
  onMirror: (mirror: MirrorMode) => void;
};

export function ResultView({
  result,
  mode,
  mirror,
  rawText,
  confidence,
  warnings,
  onMode,
  onMirror,
}: ResultViewProps) {
  const problems = result.reports.reduce((sum, report) => sum + problemCount(report), 0);
  const analysis = buildAnalysis(result);
  const [exporting, setExporting] = useState(false);
  const [exportError, setExportError] = useState<string | null>(null);

  return (
    <section className="grid min-w-0 gap-4 [&>*]:min-w-0">
      <div className="rounded-2xl border border-border bg-card p-4 sm:p-5">
        <div className="flex flex-wrap items-start justify-between gap-3">
          <div>
            <h2 className="text-lg font-semibold">
              {problems === 0 ? "Расхождений нет" : `Нужно разобрать: ${problems}`}
            </h2>
            <p className="mt-1 max-w-3xl text-sm leading-6 text-muted-foreground">
              {result.interpretation}
            </p>
            {analysis.length > 0 ? (
              <ul className="mt-3 max-w-3xl list-disc space-y-1 pl-5 text-sm leading-6">
                {analysis.map((line) => (
                  <li key={line}>{line}</li>
                ))}
              </ul>
            ) : null}
          </div>
          <Button
            type="button"
            variant="outline"
            disabled={exporting}
            onClick={() => {
              setExporting(true);
              setExportError(null);
              downloadReport(result)
                .catch(() => setExportError("Не удалось сохранить отчёт."))
                .finally(() => setExporting(false));
            }}
          >
            <Download />
            Скачать отчёт
          </Button>
        </div>
        {exportError ? <p className="mt-3 text-sm text-destructive">{exportError}</p> : null}
        {confidence != null && confidence < 60 ? (
          <p className="mt-3 text-sm leading-6 text-destructive">
            Скан прочитан неуверенно ({confidence}%). Сверьте отмеченные строки с бумагой:
            часть расхождений может быть ошибкой чтения, а не учёта.
          </p>
        ) : null}
        {warnings.map((warning) => (
          <p key={warning} className="mt-3 text-sm leading-6 text-destructive">
            {warning}
          </p>
        ))}

        <div className="mt-4 flex flex-wrap gap-2">
          {result.columnsOnly ? null : (
            <Choice
              label="Что сверять"
              value={mode}
              options={[
                ["auto", "Авто"],
                ["both", "Обе колонки"],
                ["ours", "Наша"],
                ["theirs", "Контрагент"],
              ]}
              onChange={onMode}
            />
          )}
          <Choice
            label="Суммы"
            value={mirror}
            options={[
              ["auto", "Авто"],
              ["off", "Напрямую"],
              ["on", "Зеркально"],
            ]}
            onChange={onMirror}
          />
        </div>
      </div>

      {result.reports.map((report) => (
        <ReportCard
          key={`${report.title}-${report.side}-${report.mirror}`}
          report={report}
          labels={result.labels}
        />
      ))}

      {rawText.trim() ? (
        <details className="rounded-2xl border border-border bg-card p-4 sm:p-5">
          <summary className="cursor-pointer text-sm font-medium">Текст, прочитанный со скана</summary>
          <pre className="mt-3 max-h-80 overflow-auto whitespace-pre-wrap text-sm leading-6 text-muted-foreground">
            {rawText}
          </pre>
        </details>
      ) : null}
    </section>
  );
}

function ReportCard({ report, labels }: { report: SideReport; labels: CompareLabels }) {
  const [filter, setFilter] = useState<"problems" | "all" | "matches">("problems");
  const problems = report.pairs.filter((pair) => pair.status !== "match");
  const matches = report.pairs.filter((pair) => pair.status === "match");
  const visible =
    filter === "all" ? report.pairs : filter === "matches" ? matches : problems;

  return (
    <article className="min-w-0 rounded-2xl border border-border bg-card p-4 sm:p-5">
      <h3 className="text-base font-semibold">{report.title}</h3>
      {report.mirror ? (
        <p className="mt-1 text-sm text-muted-foreground">{labels.mirrorHint}</p>
      ) : null}

      {report.summary.length > 0 ? (
        <div className="mt-4 overflow-x-auto">
          <table className="w-full min-w-[36rem] text-sm">
            <caption className="sr-only">Итоги: {report.title}</caption>
            <thead className="text-left text-muted-foreground">
              <tr className="border-b border-border">
                <th className="py-2 pr-3 font-medium">Показатель</th>
                <th className="py-2 pr-3 text-right font-medium">{labels.left}</th>
                <th className="py-2 pr-3 text-right font-medium">{labels.right}</th>
                <th className="py-2 font-medium">Статус</th>
              </tr>
            </thead>
            <tbody>
              {report.summary.map((row) => (
                <tr key={row.label} className="border-b border-border/70">
                  <td className="py-2 pr-3">{row.label}</td>
                  <td className="py-2 pr-3 text-right tabular-nums">{showAmount(row.excel)}</td>
                  <td className="py-2 pr-3 text-right tabular-nums">{showAmount(row.scan)}</td>
                  <td className="py-2">
                    <Status
                      tone={row.status === "match" ? "ok" : row.status === "diff" ? "bad" : "muted"}
                      label={
                        row.status === "match"
                          ? "Совпадает"
                          : row.status === "diff"
                            ? "Расхождение"
                            : "Не найдено"
                      }
                    />
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      ) : (
        <p className="mt-3 text-sm text-muted-foreground">
          Итоги — сальдо и обороты — в этой колонке не найдены.
        </p>
      )}

      <div className="mt-5 flex flex-wrap gap-2">
        <FilterButton active={filter === "problems"} onClick={() => setFilter("problems")}>
          Расхождения ({problems.length})
        </FilterButton>
        <FilterButton active={filter === "matches"} onClick={() => setFilter("matches")}>
          Совпадения ({matches.length})
        </FilterButton>
        <FilterButton active={filter === "all"} onClick={() => setFilter("all")}>
          Все ({report.pairs.length})
        </FilterButton>
      </div>

      {visible.length === 0 ? (
        <p className="mt-4 text-sm leading-6 text-muted-foreground">
          {filter === "problems"
            ? "В строках этой колонки расхождений нет."
            : "Строк для этого фильтра нет."}
        </p>
      ) : (
        <div className="mt-3 overflow-x-auto">
          <table className="w-full min-w-[52rem] text-sm">
            <caption className="sr-only">Строки: {report.title}</caption>
            <thead className="text-left text-muted-foreground">
              <tr className="border-b border-border">
                <th className="py-2 pr-3 font-medium">Статус</th>
                <th className="py-2 pr-3 font-medium">Дата</th>
                <th className="py-2 pr-3 font-medium">Документ</th>
                <th className="py-2 pr-3 text-right font-medium">{labels.debitLeft}</th>
                <th className="py-2 pr-3 text-right font-medium">{labels.creditLeft}</th>
                <th className="py-2 pr-3 text-right font-medium">{labels.debitRight}</th>
                <th className="py-2 text-right font-medium">{labels.creditRight}</th>
              </tr>
            </thead>
            <tbody>
              {visible.map((pair, index) => (
                <PairRow key={`${pair.status}-${index}`} pair={pair} labels={labels} />
              ))}
            </tbody>
          </table>
        </div>
      )}
    </article>
  );
}

function PairRow({ pair, labels }: { pair: Pair; labels: CompareLabels }) {
  const line = pair.excel ?? pair.scan;
  return (
    <tr className="border-b border-border/70 align-top">
      <td className="py-2 pr-3">
        <Status
          tone={pair.status === "match" ? "ok" : pair.status === "amount" ? "bad" : "muted"}
          label={
            pair.status === "match"
              ? "Совпадает"
              : pair.status === "amount"
                ? "Сумма"
                : pair.status === "excel-only"
                  ? labels.onlyLeft
                  : labels.onlyRight
          }
        />
        {pair.notes.length > 0 ? (
          <p className="mt-1 max-w-40 text-xs leading-5 text-muted-foreground">{pair.notes.join(". ")}</p>
        ) : null}
      </td>
      <td className="py-2 pr-3 tabular-nums">{formatDateRu(line?.date ?? null)}</td>
      <td className="py-2 pr-3">{line ? lineLabel(line.description, line.docNumber) : "—"}</td>
      <td className="py-2 pr-3 text-right tabular-nums">{formatMoney(pair.excel?.debit ?? null)}</td>
      <td className="py-2 pr-3 text-right tabular-nums">{formatMoney(pair.excel?.credit ?? null)}</td>
      <td className="py-2 pr-3 text-right tabular-nums">{formatMoney(pair.scan?.debit ?? null)}</td>
      <td className="py-2 text-right tabular-nums">{formatMoney(pair.scan?.credit ?? null)}</td>
    </tr>
  );
}

function showAmount(value: number | null): string {
  return value == null ? "не найдено" : formatMoney(value);
}

function Status({ tone, label }: { tone: "ok" | "bad" | "muted"; label: string }) {
  return (
    <span
      className={cn(
        "inline-flex rounded-md px-1.5 py-0.5 text-xs font-medium",
        tone === "ok" && "bg-emerald-700/10 text-emerald-800 dark:text-emerald-300",
        tone === "bad" && "bg-destructive/10 text-destructive",
        tone === "muted" && "bg-muted text-muted-foreground",
      )}
    >
      {label}
    </span>
  );
}

function FilterButton({
  active,
  onClick,
  children,
}: {
  active: boolean;
  onClick: () => void;
  children: ReactNode;
}) {
  return (
    <Button type="button" size="sm" variant={active ? "default" : "outline"} onClick={onClick}>
      {children}
    </Button>
  );
}

function Choice<T extends string>({
  label,
  value,
  options,
  onChange,
}: {
  label: string;
  value: T;
  options: [T, string][];
  onChange: (value: T) => void;
}) {
  return (
    <div className="flex flex-wrap items-center gap-2">
      <span className="text-sm text-muted-foreground">{label}</span>
      {options.map(([option, caption]) => (
        <Button
          key={option}
          type="button"
          size="sm"
          variant={value === option ? "default" : "outline"}
          aria-pressed={value === option}
          onClick={() => onChange(option)}
        >
          {caption}
        </Button>
      ))}
    </div>
  );
}
