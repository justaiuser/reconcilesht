import type {
  ActLine,
  ActSummary,
  CompareLabels,
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

type NoteVoice = "files" | "columns";

const FILE_LABELS: CompareLabels = {
  left: "Excel",
  right: "Скан",
  onlyLeft: "Только Excel",
  onlyRight: "Только скан",
  debitLeft: "Дебет Excel",
  creditLeft: "Кредит Excel",
  debitRight: "Дебет скан",
  creditRight: "Кредит скан",
  mirrorHint: "Дебет этой колонки сравнивается с кредитом скана и наоборот.",
};

const COLUMN_LABELS: CompareLabels = {
  left: "Левая колонка",
  right: "Правая колонка",
  onlyLeft: "Только слева",
  onlyRight: "Только справа",
  debitLeft: "Дебет слева",
  creditLeft: "Кредит слева",
  debitRight: "Дебет справа",
  creditRight: "Кредит справа",
  mirrorHint: "Дебет левой колонки сравнивается с кредитом правой и наоборот.",
};

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

function pairNotes(excel: ActLine, scan: ActLine, mirror: boolean, voice: NoteVoice): string[] {
  const excelDebit = amount(excel.debit);
  const excelCredit = amount(excel.credit);
  const scanDebit = amount(mirror ? scan.credit : scan.debit);
  const scanCredit = amount(mirror ? scan.debit : scan.credit);
  const notes: string[] = [];
  if (!sameMoney(excelDebit, scanDebit)) {
    notes.push(
      mirror
        ? voice === "columns"
          ? "Дебет левой колонки не равен кредиту правой"
          : "Дебет Excel не равен кредиту скана"
        : "Дебет не сходится",
    );
  }
  if (!sameMoney(excelCredit, scanCredit)) {
    notes.push(
      mirror
        ? voice === "columns"
          ? "Кредит левой колонки не равен дебету правой"
          : "Кредит Excel не равен дебету скана"
        : "Кредит не сходится",
    );
  }
  return notes;
}

function missingNote(missingFrom: "left" | "right", voice: NoteVoice): string {
  if (voice === "columns") {
    return missingFrom === "right"
      ? "Нет такой строки в правой колонке"
      : "Нет такой строки в левой колонке";
  }
  return missingFrom === "right" ? "Нет такой строки в скане" : "Нет такой строки в Excel";
}

export function matchLines(
  excelLines: ActLine[],
  scanLines: ActLine[],
  mirror: boolean,
  voice: NoteVoice = "files",
): Pair[] {
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
    const notes = pairNotes(excel, scan, mirror, voice);
    pairs.push({
      status: notes.length === 0 ? "match" : "amount",
      excel,
      scan,
      notes,
    });
  }

  excelLines.forEach((line, index) => {
    if (!usedExcel.has(index)) {
      pairs.push({
        status: "excel-only",
        excel: line,
        scan: null,
        notes: [missingNote("right", voice)],
      });
    }
  });
  scanLines.forEach((line, index) => {
    if (!usedScan.has(index)) {
      pairs.push({
        status: "scan-only",
        excel: null,
        scan: line,
        notes: [missingNote("left", voice)],
      });
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
  voice: NoteVoice,
): SideReport {
  return {
    title: sideTitle(side),
    side,
    mirror,
    pairs: matchLines(
      linesFor(excel, side),
      linesFor(scan, side === "single" ? "single" : side),
      mirror,
      voice,
    ),
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
  voice: NoteVoice,
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
    pairs: matchLines(excelLines, scanLines, mirror, voice),
    summary: compareSummaries(summaryFor(excel, side), scanSummary, mirror),
  };
}

function withLabels(result: Omit<CompareResult, "labels" | "columnsOnly">, voice: NoteVoice): CompareResult {
  return {
    ...result,
    labels: voice === "columns" ? COLUMN_LABELS : FILE_LABELS,
    columnsOnly: voice === "columns",
  };
}

export function compareActs(
  excel: ParsedAct,
  scan: ParsedAct,
  options?: { mode?: CompareMode; mirror?: MirrorMode; voice?: NoteVoice },
): CompareResult {
  const mode = options?.mode ?? "auto";
  const mirrorMode = options?.mirror ?? "auto";
  const voice = options?.voice ?? "files";
  const sides = excelSides(excel).filter((side) => side !== "single");
  const excelHasBoth = sides.includes("ours") && sides.includes("theirs");
  const scanHasBoth = scan.twoSided;

  if ((mode === "both" || (mode === "auto" && excelHasBoth && scanHasBoth)) && excelHasBoth) {
    const mirror = mirrorMode === "on";
    return withLabels(
      {
        reports: [
          buildReport(excel, scan, "ours", mirror, voice),
          buildReport(excel, scan, "theirs", mirror, voice),
        ],
        interpretation: mirror
          ? "Сравниваю обе колонки и переворачиваю суммы: дебет одного файла сверяется с кредитом другого."
          : "В обоих файлах две колонки. Сверяю нашу колонку с нашей, а колонку контрагента — с колонкой контрагента.",
      },
      voice,
    );
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
        ? buildLooseReport(excel, scan, "single", mirror, voice)
        : buildLooseReport(excel, scan, side, mirror, voice);
      const score = quality(report.pairs);
      if (!best || better(score, best.score)) best = { side, mirror, report, score };
    }
  }

  const chosen = best ?? {
    side: "single" as const,
    mirror: false,
    report: buildLooseReport(excel, scan, "single", false, voice),
    score: [0, 0] as [number, number],
  };
  const mirrorText = chosen.mirror
    ? voice === "columns"
      ? " Суммы сопоставлены зеркально: дебет одной колонки с кредитом другой."
      : " Суммы сопоставлены зеркально: дебет одного файла с кредитом другого."
    : " Суммы сравниваются напрямую, дебет с дебетом.";
  if (voice === "columns") {
    return withLabels(
      {
        reports: [chosen.report],
        interpretation: `Сверяю левую колонку с правой.${mirrorText}`,
      },
      voice,
    );
  }
  const where =
    chosen.side === "ours"
      ? "с нашей колонкой Excel"
      : chosen.side === "theirs"
        ? "с колонкой контрагента в Excel"
        : "со строками Excel";
  return withLabels(
    {
      reports: [chosen.report],
      interpretation: `Сверяю скан ${where}.${mirrorText}`,
    },
    voice,
  );
}

function columnAct(act: ParsedAct, side: "ours" | "theirs"): ParsedAct {
  const lines = act.lines
    .filter((line) => (side === "ours" ? line.side !== "theirs" : line.side === "theirs"))
    .map((line) => ({ ...line, side: "single" as const }));
  return {
    title: act.title,
    period: act.period,
    partyA: act.partyA,
    partyB: act.partyB,
    lines,
    summary: side === "theirs" ? (act.theirsSummary ?? emptySummary()) : act.summary,
    theirsSummary: null,
    warnings: [],
    confidence: act.confidence,
    rawText: act.rawText,
    twoSided: false,
  };
}

function actsWord(count: number): string {
  const mod10 = count % 10;
  const mod100 = count % 100;
  if (mod10 >= 2 && mod10 <= 4 && (mod100 < 12 || mod100 > 14)) return "акта";
  return "актов";
}

function documentTitle(act: ParsedAct, index: number): string {
  const who = [act.partyA, act.partyB].filter(Boolean).join(" — ");
  const period = act.period ? ` (${act.period})` : "";
  if (who) return `Акт ${index + 1}: ${who}${period}`;
  return `Акт ${index + 1}`;
}

export function compareDocuments(
  acts: ParsedAct[],
  options?: { mirror?: MirrorMode },
): CompareResult | null {
  const sided = acts.filter((act) => act.twoSided);
  if (sided.length === 0) return null;
  const pieces = sided.map((act, index) => {
    const compared = compareActs(columnAct(act, "ours"), columnAct(act, "theirs"), {
      mode: "auto",
      mirror: options?.mirror ?? "auto",
      voice: "columns",
    });
    const title = sided.length > 1 ? documentTitle(act, index) : "Левая и правая колонки";
    return {
      ...compared,
      reports: compared.reports.map((report) => ({ ...report, title })),
    };
  });
  const skipped = acts.length - sided.length;
  const skippedText =
    skipped > 0 ? " Часть страниц без двух колонок в сверку не вошла." : "";
  const interpretation =
    pieces.length === 1
      ? `${pieces[0].interpretation}${skippedText}`
      : `В файле ${pieces.length} ${actsWord(pieces.length)}: в каждом левая колонка сверяется с правой.${skippedText}`;
  return {
    reports: pieces.flatMap((piece) => piece.reports),
    interpretation,
    labels: COLUMN_LABELS,
    columnsOnly: true,
  };
}

export function problemCount(report: SideReport): number {
  return (
    report.pairs.filter((pair) => pair.status !== "match").length +
    report.summary.filter((row) => row.status === "diff").length
  );
}
