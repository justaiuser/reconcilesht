import type { CompareResult, ParsedAct, SideReport } from "@/lib/types";

/** OCR / row parse looks too thin to trust for reconciliation. */
export function isWeakRecognition(act: ParsedAct): boolean {
  const withAmount = act.lines.filter((line) => line.debit != null || line.credit != null);
  if (withAmount.length === 0) return true;
  if (act.confidence != null && act.confidence < 45) return true;
  // One vague amount without a document number and almost no text is usually OCR noise.
  const hasDoc = withAmount.some((line) => Boolean(line.docNumber));
  if (withAmount.length === 1 && !hasDoc && (act.rawText?.length ?? 0) < 40) return true;
  return false;
}

export type AnalysisCounts = {
  matches: number;
  amountDiffs: number;
  onlyLeft: number;
  onlyRight: number;
  summaryDiffs: number;
  summaryMissing: number;
};

export function analysisCounts(result: CompareResult): AnalysisCounts {
  const counts: AnalysisCounts = {
    matches: 0,
    amountDiffs: 0,
    onlyLeft: 0,
    onlyRight: 0,
    summaryDiffs: 0,
    summaryMissing: 0,
  };
  for (const report of result.reports) {
    addReportCounts(counts, report);
  }
  return counts;
}

function addReportCounts(counts: AnalysisCounts, report: SideReport) {
  for (const pair of report.pairs) {
    if (pair.status === "match") counts.matches += 1;
    else if (pair.status === "amount") counts.amountDiffs += 1;
    else if (pair.status === "excel-only") counts.onlyLeft += 1;
    else counts.onlyRight += 1;
  }
  for (const row of report.summary) {
    if (row.status === "diff") counts.summaryDiffs += 1;
    else if (row.status === "missing") counts.summaryMissing += 1;
  }
}

/** Short Russian analysis for the accountant. */
export function buildAnalysis(result: CompareResult): string[] {
  const counts = analysisCounts(result);
  const lines: string[] = [];
  const totalPairs = counts.matches + counts.amountDiffs + counts.onlyLeft + counts.onlyRight;

  if (totalPairs === 0 && counts.summaryDiffs === 0 && counts.summaryMissing === 0) {
    return ["Строк для сравнения не нашлось. Проверьте файл скана и сопоставление колонок Excel."];
  }

  if (counts.matches > 0) {
    lines.push(
      counts.amountDiffs + counts.onlyLeft + counts.onlyRight === 0
        ? `По строкам всё сходится: ${ruPairs(counts.matches)}.`
        : `Совпало ${ruPairs(counts.matches)}.`,
    );
  }

  if (counts.amountDiffs > 0) {
    lines.push(
      `Разные суммы по ${ruDocs(counts.amountDiffs)} — сверьте дебет и кредит с бумагой и проводками.`,
    );
  }
  if (counts.onlyLeft > 0) {
    lines.push(
      result.columnsOnly
        ? `Только в левой колонке: ${ruDocs(counts.onlyLeft)}.`
        : `Есть в Excel, нет на скане: ${ruDocs(counts.onlyLeft)}.`,
    );
  }
  if (counts.onlyRight > 0) {
    lines.push(
      result.columnsOnly
        ? `Только в правой колонке: ${ruDocs(counts.onlyRight)}.`
        : `Есть на скане, нет в Excel: ${ruDocs(counts.onlyRight)}.`,
    );
  }

  if (counts.summaryDiffs > 0) {
    lines.push(
      `Сальдо или обороты расходятся по ${ruIndicators(counts.summaryDiffs)} — смотрите блок итогов.`,
    );
  } else if (totalPairs > 0 && counts.summaryMissing === 0) {
    const hasSummary = result.reports.some((report) => report.summary.length > 0);
    if (hasSummary) lines.push("Начальное и конечное сальдо, а также обороты сходятся.");
  } else if (counts.summaryMissing > 0 && counts.summaryDiffs === 0) {
    lines.push("Часть итогов (сальдо/обороты) на скане не прочиталась — сверьте их вручную.");
  }

  return lines;
}

function ruPairs(n: number): string {
  return `${n} ${plural(n, "строка", "строки", "строк")}`;
}

function ruDocs(n: number): string {
  return `${n} ${plural(n, "документу", "документам", "документам")}`;
}

function ruIndicators(n: number): string {
  return `${n} ${plural(n, "показателю", "показателям", "показателям")}`;
}

function plural(n: number, one: string, few: string, many: string): string {
  const mod10 = n % 10;
  const mod100 = n % 100;
  if (mod10 === 1 && mod100 !== 11) return one;
  if (mod10 >= 2 && mod10 <= 4 && (mod100 < 12 || mod100 > 14)) return few;
  return many;
}
