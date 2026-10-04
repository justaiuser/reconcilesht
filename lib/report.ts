import type { CompareResult, Pair } from "@/lib/types";
import { formatDateRu, lineLabel } from "@/lib/text";

export async function downloadReport(result: CompareResult): Promise<void> {
  const { default: writeExcelFile } = await import("write-excel-file/browser");
  const labels = result.labels;
  const summary: (string | number | null)[][] = [
    ["Колонка", "Показатель", labels.left, labels.right, "Статус"],
  ];
  const lines: (string | number | null)[][] = [
    [
      "Колонка",
      "Статус",
      "Дата",
      "Документ",
      labels.debitLeft,
      labels.creditLeft,
      labels.debitRight,
      labels.creditRight,
      "Комментарий",
    ],
  ];

  for (const report of result.reports) {
    for (const row of report.summary) {
      summary.push([
        report.title,
        row.label,
        row.excel,
        row.scan,
        row.status === "match" ? "Совпадает" : row.status === "diff" ? "Расхождение" : "Не найдено",
      ]);
    }
    for (const pair of report.pairs) {
      lines.push(pairRow(report.title, pair, result));
    }
  }

  const blob = await writeExcelFile([
    { data: summary, sheet: "Итоги" },
    { data: lines, sheet: "Строки" },
  ]).toBlob();
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = "Расхождения сверки.xlsx";
  link.click();
  URL.revokeObjectURL(url);
}

function pairRow(title: string, pair: Pair, result: CompareResult): (string | number | null)[] {
  const line = pair.excel ?? pair.scan;
  const status =
    pair.status === "match"
      ? "Совпадает"
      : pair.status === "amount"
        ? "Расхождение суммы"
        : pair.status === "excel-only"
          ? result.labels.onlyLeft
          : result.labels.onlyRight;
  return [
    title,
    status,
    formatDateRu(line?.date ?? null),
    line ? lineLabel(line.description, line.docNumber) : "—",
    pair.excel?.debit ?? null,
    pair.excel?.credit ?? null,
    pair.scan?.debit ?? null,
    pair.scan?.credit ?? null,
    pair.notes.join("; "),
  ];
}
