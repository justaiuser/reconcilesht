import type { CompareResult, Pair } from "@/lib/types";
import { formatDateRu, lineLabel } from "@/lib/text";

const STATUS_LABEL = {
  match: "Совпадает",
  amount: "Расхождение суммы",
  "excel-only": "Только в Excel",
  "scan-only": "Только в скане",
} as const;

export async function downloadReport(result: CompareResult): Promise<void> {
  const { default: writeExcelFile } = await import("write-excel-file/browser");
  const summary: (string | number | null)[][] = [
    ["Колонка", "Показатель", "Excel", "Скан", "Статус"],
  ];
  const lines: (string | number | null)[][] = [
    [
      "Колонка",
      "Статус",
      "Дата",
      "Документ",
      "Дебет Excel",
      "Кредит Excel",
      "Дебет скан",
      "Кредит скан",
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
      lines.push(pairRow(report.title, pair));
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

function pairRow(title: string, pair: Pair): (string | number | null)[] {
  const line = pair.excel ?? pair.scan;
  return [
    title,
    STATUS_LABEL[pair.status],
    formatDateRu(line?.date ?? null),
    line ? lineLabel(line.description, line.docNumber) : "—",
    pair.excel?.debit ?? null,
    pair.excel?.credit ?? null,
    pair.scan?.debit ?? null,
    pair.scan?.credit ?? null,
    pair.notes.join("; "),
  ];
}
