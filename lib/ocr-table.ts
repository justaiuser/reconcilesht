import type { ActLine, ActSummary, OcrWord, ParsedAct, SideName } from "@/lib/types";
import { emptySummary } from "@/lib/excel-act";
import { headerRole } from "@/lib/excel-act";
import {
  docFromDocumentCell,
  extractDocNumber,
  extractParties,
  extractPeriod,
  normalizeHeader,
  parseDate,
  parseMoney,
} from "@/lib/text";

type Role = "date" | "document" | "debit" | "credit";

type ColumnBound = {
  role: Role;
  side: SideName;
  left: number;
  right: number;
};

function summaryKind(text: string): "opening" | "turnover" | "closing" | null {
  const normalized = normalizeHeader(text);
  if (/сальдо началь/.test(normalized)) return "opening";
  if (/обороты/.test(normalized)) return "turnover";
  if (/сальдо конеч/.test(normalized)) return "closing";
  return null;
}

export function groupLines(words: OcrWord[]): OcrWord[][] {
  const items = words.filter((word) => word.text.trim());
  items.sort((a, b) => a.y0 - b.y0 || a.x0 - b.x0);
  const lines: { y: number; height: number; words: OcrWord[] }[] = [];

  for (const word of items) {
    const center = (word.y0 + word.y1) / 2;
    const height = Math.max(8, word.y1 - word.y0);
    let best = -1;
    let bestDistance = Infinity;
    for (let index = 0; index < lines.length; index += 1) {
      const distance = Math.abs(lines[index].y - center);
      const limit = Math.max(lines[index].height, height) * 0.65;
      if (distance <= limit && distance < bestDistance) {
        best = index;
        bestDistance = distance;
      }
    }
    if (best < 0) {
      lines.push({ y: center, height, words: [word] });
      continue;
    }
    const line = lines[best];
    line.words.push(word);
    const count = line.words.length;
    line.y = (line.y * (count - 1) + center) / count;
    line.height = Math.max(line.height, height);
  }

  return lines
    .sort((a, b) => a.y - b.y)
    .map((line) => line.words.sort((a, b) => a.x0 - b.x0));
}

function lineText(line: OcrWord[]): string {
  return line
    .map((word) => word.text.trim())
    .filter(Boolean)
    .join(" ");
}

function headerScore(line: OcrWord[]): number {
  let score = 0;
  let dates = 0;
  let amounts = 0;
  for (const word of line) {
    const role = headerRole(word.text);
    if (role === "date") {
      score += 3;
      dates += 1;
    } else if (role === "document") score += 2;
    else if (role === "debit" || role === "credit") {
      score += 3;
      amounts += 1;
    }
  }
  if (dates === 0 || amounts === 0) return 0;
  return score;
}

function boundsFromHeader(line: OcrWord[]): ColumnBound[] {
  const columns: { role: Role; center: number }[] = [];
  for (const word of line) {
    const role = headerRole(word.text);
    if (role === "other") continue;
    columns.push({ role, center: (word.x0 + word.x1) / 2 });
  }
  columns.sort((a, b) => a.center - b.center);

  const sides: { role: Role; center: number; side: SideName }[] = [];
  let side: SideName = "ours";
  let seenDate = false;
  for (const column of columns) {
    if (column.role === "date" && seenDate) side = "theirs";
    if (column.role === "date") seenDate = true;
    sides.push({ ...column, side: side === "theirs" ? "theirs" : "ours" });
  }
  if (!sides.some((column) => column.side === "theirs")) {
    for (const column of sides) column.side = "single";
  }

  return sides.map((column, index) => {
    const previous = sides[index - 1]?.center;
    const next = sides[index + 1]?.center;
    return {
      role: column.role,
      side: column.side,
      left: previous == null ? Number.NEGATIVE_INFINITY : (previous + column.center) / 2,
      right: next == null ? Number.POSITIVE_INFINITY : (column.center + next) / 2,
    };
  });
}

function joinWords(words: OcrWord[]): string {
  const texts = words.map((word) => word.text.trim()).filter(Boolean);
  if (texts.length === 0) return "";
  if (texts.every((text) => /^[\d.,\-–—]+$/.test(text))) return texts.join("");
  return texts.join(" ");
}

function wordsFor(
  line: OcrWord[],
  bounds: ColumnBound[],
  side: SideName,
  role: Role,
): OcrWord[] {
  const column = bounds.find((item) => item.side === side && item.role === role);
  if (!column) return [];
  return line.filter((word) => {
    const center = (word.x0 + word.x1) / 2;
    return center >= column.left && center < column.right;
  });
}

function applySummary(
  summary: ActSummary,
  kind: "opening" | "turnover" | "closing",
  debit: number | null,
  credit: number | null,
) {
  if (kind === "opening") {
    if (debit != null) summary.openingDebit = debit;
    if (credit != null) summary.openingCredit = credit;
  } else if (kind === "turnover") {
    if (debit != null) summary.turnoverDebit = debit;
    if (credit != null) summary.turnoverCredit = credit;
  } else {
    if (debit != null) summary.closingDebit = debit;
    if (credit != null) summary.closingCredit = credit;
  }
}

function parseWithColumns(lines: OcrWord[][], bounds: ColumnBound[]): {
  actLines: ActLine[];
  summary: ActSummary;
  theirsSummary: ActSummary | null;
} {
  const sides: SideName[] = bounds.some((column) => column.side === "theirs")
    ? ["ours", "theirs"]
    : ["single"];
  const summary = emptySummary();
  const theirsSummary = sides.includes("theirs") ? emptySummary() : null;
  const actLines: ActLine[] = [];

  lines.forEach((line, index) => {
    const text = lineText(line);
    const kind = summaryKind(text);
    if (kind) {
      for (const side of sides) {
        const debit = parseMoney(joinWords(wordsFor(line, bounds, side, "debit")));
        const credit = parseMoney(joinWords(wordsFor(line, bounds, side, "credit")));
        applySummary(side === "theirs" && theirsSummary ? theirsSummary : summary, kind, debit, credit);
      }
      return;
    }

    for (const side of sides) {
      const dateRaw = joinWords(wordsFor(line, bounds, side, "date"));
      const documentRaw = joinWords(wordsFor(line, bounds, side, "document"));
      const debit = parseMoney(joinWords(wordsFor(line, bounds, side, "debit")));
      const credit = parseMoney(joinWords(wordsFor(line, bounds, side, "credit")));
      const date = parseDate(dateRaw);
      const docNumber = docFromDocumentCell(documentRaw);
      if (!date && debit == null && credit == null) continue;
      if (!date && !docNumber) continue;
      actLines.push({
        id: `${side}-${index}`,
        side,
        date,
        docNumber,
        description: documentRaw,
        debit,
        credit,
      });
    }
  });

  return { actLines, summary, theirsSummary };
}

function moneyTokens(line: OcrWord[]): { value: number; x: number }[] {
  const tokens: { value: number; x: number }[] = [];
  let buffer: OcrWord[] = [];
  const flush = () => {
    if (buffer.length === 0) return;
    const value = parseMoney(joinWords(buffer));
    if (value != null) {
      const x0 = Math.min(...buffer.map((word) => word.x0));
      const x1 = Math.max(...buffer.map((word) => word.x1));
      tokens.push({ value, x: (x0 + x1) / 2 });
    }
    buffer = [];
  };
  for (const word of line) {
    if (/^[\d.,\-–—]+$/.test(word.text.trim())) buffer.push(word);
    else flush();
  }
  flush();
  return tokens;
}

function parsePlain(lines: OcrWord[][]): ActLine[] {
  const amountCenters = lines.flatMap((line) => moneyTokens(line).map((token) => token.x));
  const sorted = [...amountCenters].sort((a, b) => a - b);
  let split: number | null = null;
  let widest = 0;
  for (let index = 1; index < sorted.length; index += 1) {
    const gap = sorted[index] - sorted[index - 1];
    if (gap > widest) {
      widest = gap;
      split = (sorted[index] + sorted[index - 1]) / 2;
    }
  }
  if (widest < 40) split = null;

  const actLines: ActLine[] = [];
  lines.forEach((line, index) => {
    const text = lineText(line);
    if (summaryKind(text) || headerScore(line) > 0) return;
    const date = parseDate(text);
    if (!date) return;
    const tokens = moneyTokens(line);
    let debit: number | null = null;
    let credit: number | null = null;
    if (split == null) {
      if (tokens[0]) debit = tokens[0].value;
      if (tokens[1]) credit = tokens[1].value;
    } else {
      for (const token of tokens) {
        if (token.x < split) debit = token.value;
        else credit = token.value;
      }
    }
    if (debit == null && credit == null) return;
    const description = text
      .replace(/(\d{1,2}[./-]\d{1,2}[./-]\d{2,4})/, "")
      .replace(/[\d\s.,\-–—]+$/g, "")
      .trim();
    actLines.push({
      id: `plain-${index}`,
      side: "single",
      date,
      docNumber: extractDocNumber(text),
      description,
      debit,
      credit,
    });
  });
  return actLines;
}

function meanConfidence(words: OcrWord[]): number | null {
  if (words.length === 0) return null;
  const total = words.reduce((sum, word) => sum + word.confidence, 0);
  return Math.round(total / words.length);
}

export function parseOcrWords(words: OcrWord[], rawText = ""): ParsedAct {
  const lines = groupLines(words);
  const header = [...lines].sort((a, b) => headerScore(b) - headerScore(a))[0];
  const bounds = header && headerScore(header) > 0 ? boundsFromHeader(header) : [];
  const columnParsed = bounds.length
    ? parseWithColumns(lines, bounds)
    : { actLines: [], summary: emptySummary(), theirsSummary: null };
  const plain = parsePlain(lines);
  const usePlain = columnParsed.actLines.length === 0 && plain.length > 0;
  const actLines = usePlain ? plain : columnParsed.actLines;
  const warnings: string[] = [];
  if (bounds.length === 0) {
    warnings.push(
      "На скане не нашлась шапка таблицы с колонками «Дата» и «Дебет». Строки собраны по тексту, суммы стоит проверить.",
    );
  }
  if (actLines.length === 0) {
    warnings.push(
      "Таблица на скане не разобралась. Ниже есть распознанный текст — по нему можно сверить итоги. Если скан бледный или перекошен, пересканируйте его при 300 dpi.",
    );
  }
  const fullText = rawText || lines.map(lineText).join("\n");
  const parties = extractParties(fullText);
  return {
    title: null,
    period: extractPeriod(fullText),
    partyA: parties.partyA,
    partyB: parties.partyB,
    lines: actLines,
    summary: usePlain ? emptySummary() : columnParsed.summary,
    theirsSummary: usePlain ? null : columnParsed.theirsSummary,
    warnings,
    confidence: meanConfidence(words),
    rawText: fullText,
    twoSided: actLines.some((line) => line.side === "theirs"),
  };
}

export function parseOcrPages(pages: OcrWord[][], rawText: string): ParsedAct {
  if (pages.length === 0) return parseOcrWords([], rawText);
  const parsed = pages.map((page, index) =>
    parseOcrWords(page, index === pages.length - 1 ? rawText : ""),
  );
  const summary = emptySummary();
  let theirsSummary: ActSummary | null = null;
  const lines: ActLine[] = [];
  const warnings = new Set<string>();
  let confidenceSum = 0;
  let confidenceCount = 0;

  for (const page of parsed) {
    lines.push(
      ...page.lines.map((line, index) => ({
        ...line,
        id: `${lines.length + index}-${line.id}`,
      })),
    );
    for (const key of Object.keys(summary) as (keyof ActSummary)[]) {
      if (page.summary[key] != null) summary[key] = page.summary[key];
    }
    if (page.theirsSummary) {
      theirsSummary ??= emptySummary();
      for (const key of Object.keys(theirsSummary) as (keyof ActSummary)[]) {
        if (page.theirsSummary[key] != null) theirsSummary[key] = page.theirsSummary[key];
      }
    }
    for (const warning of page.warnings) warnings.add(warning);
    if (page.confidence != null) {
      confidenceSum += page.confidence;
      confidenceCount += 1;
    }
  }

  const first = parsed.find((page) => page.partyA) ?? parsed[0];
  const withLines = lines.length > 0;
  return {
    title: first.title,
    period: parsed.find((page) => page.period)?.period ?? null,
    partyA: first.partyA,
    partyB: first.partyB,
    lines,
    summary,
    theirsSummary,
    warnings: withLines
      ? [...warnings].filter((warning) => !warning.startsWith("Таблица на скане"))
      : [...warnings],
    confidence: confidenceCount ? Math.round(confidenceSum / confidenceCount) : null,
    rawText,
    twoSided: lines.some((line) => line.side === "theirs"),
  };
}
