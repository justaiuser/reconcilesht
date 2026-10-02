import type {
  ActLine,
  ActSummary,
  CompareMode,
  CompareResult,
  MirrorMode,
  Pair,
  ParsedAct,
  SideName,
  SideReport,
  SummaryRow,
} from "@/lib/types";
import { emptySummary } from "@/lib/excel-act";
import { normalizeDoc, sameMoney } from "@/lib/text";

const SUMMARY_FIELDS: { key: keyof ActSummary; label: string; kind: "debit" | "credit" }[] = [
  { key: "openingDebit", label: "Сальдо начальное, дебет", kind: "debit" },
  { key: "openingCredit", label: "Сальдо начальное, кредит", kind: "credit" },
  { key: "turnoverDebit", label: "Обороты по дебету", kind: "debit" },
  { key: "turnoverCredit", label: "Обороты по кредиту", kind: "credit" },
  { key: "closingDebit", label: "Сальдо конечное, дебет", kind: "debit" },
  { key: "closingCredit", label: "Сальдо конечное, кредит", kind: "credit" },
];

function amount(value: number | null): number {
  return value == null ? 0 : value;
}

function excelSides(act: ParsedAct): SideName[] {
  const sides = new Set(act.lines.map((line) => line.side));
  if (sides.has("ours") && sides.has("theirs")) return ["ours", "theirs"];
  if (sides.has("theirs")) return ["theirs"];
  if (sides.has("ours")) return ["ours"];
  return ["single"];
}

function linesFor(act: ParsedAct, side: SideName): ActLine[] {
  const present = new Set(act.lines.map((line) => line.side));
  if (side === "single" || (!present.has("ours") && !present.has("theirs"))) return act.lines;
  if (present.has(side)) return act.lines.filter((line) => line.side === side);
  return [];
}

function summaryFor(act: ParsedAct, side: SideName): ActSummary {
  if (side === "theirs") return act.theirsSummary ?? emptySummary();
  return act.summary;
}

function scorePair(excel: ActLine, scan: ActLine, mirror: boolean): number {
  const excelDoc = normalizeDoc(excel.docNumber);
  const scanDoc = normalizeDoc(scan.docNumber);
  if (excelDoc && scanDoc && excelDoc !== scanDoc && !excelDoc.includes(scanDoc) && !scanDoc.includes(excelDoc)) {
    return -1;
  }

  let score = 0;
  if (excelDoc && scanDoc && excelDoc === scanDoc) score += 6;
  else if (excelDoc && scanDoc) score += 3;

  if (excel.date && scan.date && excel.date === scan.date) score += 3;
  else if (excel.date && scan.date) score -= 1;

  const excelDebit = amount(excel.debit);
  const excelCredit = amount(excel.credit);
  const scanDebit = amount(mirror ? scan.credit : scan.debit);
  const scanCredit = amount(mirror ? scan.debit : scan.credit);
  const debitMatches = sameMoney(excelDebit, scanDebit);
  const creditMatches = sameMoney(excelCredit, scanCredit);
  if (debitMatches && creditMatches && (excelDebit !== 0 || excelCredit !== 0)) score += 4;
  else if (debitMatches || creditMatches) score += 1;
  return score;
}

function pairNotes(excel: ActLine, scan: ActLine, mirror: boolean): string[] {
  const excelDebit = amount(excel.debit);
  const excelCredit = amount(excel.credit);
  const scanDebit = amount(mirror ? scan.credit : scan.debit);
  const scanCredit = amount(mirror ? scan.debit : scan.credit);
  const notes: string[] = [];
  if (!sameMoney(excelDebit, scanDebit)) {
    notes.push(mirror ? "Дебет Excel не равен кредиту скана" : "Дебет не сходится");
  }
  if (!sameMoney(excelCredit, scanCredit)) {
    notes.push(mirror ? "Кредит Excel не равен дебету скана" : "Кредит не сходится");
  }
  return notes;
}

export function matchLines(excelLines: ActLine[], scanLines: ActLine[], mirror: boolean): Pair[] {
  const candidates: { excelIndex: number; scanIndex: number; score: number }[] = [];
  excelLines.forEach((excelLine, excelIndex) => {
    scanLines.forEach((scanLine, scanIndex) => {
      const score = scorePair(excelLine, scanLine, mirror);
      if (score >= 5) candidates.push({ excelIndex, scanIndex, score });
    });
  });
  candidates.sort(
    (a, b) => b.score - a.score || a.excelIndex - b.excelIndex || a.scanIndex - b.scanIndex,
  );

  const usedExcel = new Set<number>();
  const usedScan = new Set<number>();
  const pairs: Pair[] = [];
  for (const candidate of candidates) {
    if (usedExcel.has(candidate.excelIndex) || usedScan.has(candidate.scanIndex)) continue;
    usedExcel.add(candidate.excelIndex);
    usedScan.add(candidate.scanIndex);
    const excel = excelLines[candidate.excelIndex];
    const scan = scanLines[candidate.scanIndex];
    const notes = pairNotes(excel, scan, mirror);
    pairs.push({
      status: notes.length === 0 ? "match" : "amount",
      excel,
      scan,
      notes,
    });
  }

  excelLines.forEach((line, index) => {
    if (!usedExcel.has(index)) {
      pairs.push({ status: "excel-only", excel: line, scan: null, notes: ["Нет такой строки в скане"] });
    }
  });
  scanLines.forEach((line, index) => {
    if (!usedScan.has(index)) {
      pairs.push({ status: "scan-only", excel: null, scan: line, notes: ["Нет такой строки в Excel"] });
    }
  });

  const order = { amount: 0, "excel-only": 1, "scan-only": 2, match: 3 };
  pairs.sort((a, b) => {
    const byStatus = order[a.status] - order[b.status];
    if (byStatus !== 0) return byStatus;
    const aDate = a.excel?.date ?? a.scan?.date ?? "";
    const bDate = b.excel?.date ?? b.scan?.date ?? "";
    return aDate.localeCompare(bDate);
  });
  return pairs;
}

function compareSummaries(excel: ActSummary, scan: ActSummary, mirror: boolean): SummaryRow[] {
  const scanValue = (kind: "debit" | "credit", key: keyof ActSummary) => {
    if (!mirror) return scan[key];
    const swapped: keyof ActSummary =
      kind === "debit"
        ? (key.replace("Debit", "Credit") as keyof ActSummary)
        : (key.replace("Credit", "Debit") as keyof ActSummary);
    return scan[swapped];
  };

  return SUMMARY_FIELDS.flatMap((field) => {
    const excelValue = excel[field.key];
    const value = scanValue(field.kind, field.key);
    if (excelValue == null && value == null) return [];
    let status: SummaryRow["status"] = "missing";
    if (excelValue != null && value != null) {
      status = sameMoney(excelValue, value) ? "match" : "diff";
    }
    return [{ label: field.label, excel: excelValue, scan: value, status }];
  });
}

function sideTitle(side: SideName): string {
  if (side === "ours") return "По нашим данным";
  if (side === "theirs") return "По данным контрагента";
  return "Строки акта";
}

function quality(pairs: Pair[]): [number, number] {
  const paired = pairs.filter((pair) => pair.excel && pair.scan).length;
  const matched = pairs.filter((pair) => pair.status === "match").length;
  return [paired, matched];
}

function better(a: [number, number], b: [number, number]): boolean {
  return a[0] > b[0] || (a[0] === b[0] && a[1] > b[1]);
}

function buildReport(
  excel: ParsedAct,
  scan: ParsedAct,
  side: SideName,
  mirror: boolean,
): SideReport {
  return {
    title: sideTitle(side),
    side,
    mirror,
    pairs: matchLines(linesFor(excel, side), linesFor(scan, side === "single" ? "single" : side), mirror),
    summary: compareSummaries(summaryFor(excel, side), summaryFor(scan, side === "single" ? "single" : side), mirror),
  };
}

function scanLinesForSide(scan: ParsedAct, side: SideName): ActLine[] {
  const present = new Set(scan.lines.map((line) => line.side));
  if (!present.has("ours") && !present.has("theirs")) return scan.lines;
  if (present.has(side)) return scan.lines.filter((line) => line.side === side);
  return scan.lines;
}

function buildLooseReport(
  excel: ParsedAct,
  scan: ParsedAct,
  side: SideName,
  mirror: boolean,
): SideReport {
  const excelLines = linesFor(excel, side);
  const scanLines = scanLinesForSide(scan, side);
  const scanSummary =
    side === "theirs" && scan.theirsSummary
      ? scan.theirsSummary
      : scan.twoSided
        ? summaryFor(scan, side)
        : scan.summary;
  return {
    title: sideTitle(side),
    side,
    mirror,
    pairs: matchLines(excelLines, scanLines, mirror),
    summary: compareSummaries(summaryFor(excel, side), scanSummary, mirror),
  };
}

export function compareActs(
  excel: ParsedAct,
  scan: ParsedAct,
  options?: { mode?: CompareMode; mirror?: MirrorMode },
): CompareResult {
  const mode = options?.mode ?? "auto";
  const mirrorMode = options?.mirror ?? "auto";
  const sides = excelSides(excel).filter((side) => side !== "single");
  const excelHasBoth = sides.includes("ours") && sides.includes("theirs");
  const scanHasBoth = scan.twoSided;

  if ((mode === "both" || (mode === "auto" && excelHasBoth && scanHasBoth)) && excelHasBoth) {
    const mirror = mirrorMode === "on";
    return {
      reports: [
        buildReport(excel, scan, "ours", mirror),
        buildReport(excel, scan, "theirs", mirror),
      ],
      interpretation: mirror
        ? "Сравниваю обе колонки и переворачиваю суммы: дебет одного файла сверяется с кредитом другого."
        : "В обоих файлах две колонки. Сверяю нашу колонку с нашей, а колонку контрагента — с колонкой контрагента.",
    };
  }

  const candidateSides: SideName[] = mode === "ours" || mode === "theirs"
    ? [mode]
    : excelHasBoth
      ? ["ours", "theirs"]
      : ["single"];
  const mirrors = mirrorMode === "auto" ? [false, true] : [mirrorMode === "on"];

  let best: { side: SideName; mirror: boolean; report: SideReport; score: [number, number] } | null =
    null;
  for (const side of candidateSides) {
    for (const mirror of mirrors) {
      const report = side === "single"
        ? buildLooseReport(excel, scan, "single", mirror)
        : buildLooseReport(excel, scan, side, mirror);
      const score = quality(report.pairs);
      if (!best || better(score, best.score)) best = { side, mirror, report, score };
    }
  }

  const chosen = best ?? {
    side: "single" as const,
    mirror: false,
    report: buildLooseReport(excel, scan, "single", false),
    score: [0, 0] as [number, number],
  };
  const where =
    chosen.side === "ours"
      ? "с нашей колонкой Excel"
      : chosen.side === "theirs"
        ? "с колонкой контрагента в Excel"
        : "со строками Excel";
  const mirrorText = chosen.mirror
    ? " Суммы сопоставлены зеркально: дебет одного файла с кредитом другого."
    : " Суммы сравниваются напрямую, дебет с дебетом.";
  return {
    reports: [chosen.report],
    interpretation: `Сверяю скан ${where}.${mirrorText}`,
  };
}

export function problemCount(report: SideReport): number {
  return (
    report.pairs.filter((pair) => pair.status !== "match").length +
    report.summary.filter((row) => row.status === "diff").length
  );
}
