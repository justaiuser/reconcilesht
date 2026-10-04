import type { ActLine, ActSummary, ParsedAct } from "@/lib/types";
import { emptySummary } from "@/lib/excel-act";
import { docFromDocumentCell, parseDate, parseMoney, summaryKind } from "@/lib/text";

export type PrintedLine = {
  date: string | null;
  number: string | null;
  debit: number | null;
  credit: number | null;
  description: string;
  side: "ours" | "theirs" | "single";
};

export type PrintedPage = {
  lines: PrintedLine[];
  summary: ActSummary;
  theirsSummary: ActSummary | null;
  rawText: string;
};

type SummaryKey = "opening" | "turnover" | "closing";

export function readPrintedPage(texts: string[]): PrintedPage {
  const lines: PrintedLine[] = [];
  const summary = emptySummary();
  const right = emptySummary();
  let rightUsed = false;
  let pending: SummaryKey | null = null;
  let pendingRight: SummaryKey | null = null;

  const take = (text: string, side: PrintedLine["side"], label: SummaryKey | null) => {
    const kind = labelKind(text);
    const amounts = moneyAmounts(text);
    if (kind && amounts.length === 0) return kind;
    const target = side === "theirs" ? right : summary;
    const waiting = side === "theirs" ? pendingRight : pending;
    if (!invoiceNumber(text) && !firstDate(text) && amounts.length > 0 && waiting) {
      writeSummary(target, waiting, amounts);
      if (side === "theirs") {
        pendingRight = null;
        rightUsed = true;
      } else pending = null;
      return null;
    }
    const date = firstDate(text);
    const number = invoiceNumber(text);
    if (number && date && amounts.length > 0) {
      lines.push({
        side,
        date,
        number,
        debit: amounts[0] ?? null,
        credit: amounts[1] ?? null,
        description: number,
      });
    }
    return label;
  };

  for (const text of texts) {
    const parts = splitSides(text);
    if (parts) {
      const leftKind = take(parts[0], "ours", pending);
      if (leftKind) pending = leftKind;
      const rightKind = take(parts[1], "theirs", pendingRight);
      if (rightKind) pendingRight = rightKind;
      if (rightKind || invoiceNumber(parts[1])) rightUsed = true;
      continue;
    }
    const kind = take(text, "single", pending);
    if (kind) pending = kind;
  }

  return {
    lines,
    summary,
    theirsSummary: rightUsed ? right : null,
    rawText: texts.map((text) => text.trim()).filter(Boolean).join("\n"),
  };
}

export function printedPagesToRecognition(pages: PrintedPage[]): { act: ParsedAct; documents: ParsedAct[] } {
  const act = emptyAct();
  for (const page of pages) {
    act.lines.push(
      ...page.lines.map((line, index) => ({
        id: `${act.lines.length + index}`,
        side: line.side,
        date: line.date,
        docNumber: line.number,
        description: line.description,
        debit: line.debit,
        credit: line.credit,
      })),
    );
    fillSummary(act.summary, page.summary);
    if (page.theirsSummary) {
      act.theirsSummary ??= emptySummary();
      fillSummary(act.theirsSummary, page.theirsSummary);
      act.twoSided = true;
    }
    act.rawText = [act.rawText, page.rawText].filter(Boolean).join("\n");
  }
  act.twoSided = act.twoSided || act.lines.some((line) => line.side === "theirs");
  return { act, documents: [act] };
}

function labelKind(text: string): SummaryKey | null {
  const known = summaryKind(text);
  if (known) return known;
  if (/итого за период|обороты по/i.test(text)) return "turnover";
  return null;
}

function writeSummary(summary: ActSummary, kind: SummaryKey, amounts: number[]) {
  const debitKey = `${kind}Debit` as keyof ActSummary;
  const creditKey = `${kind}Credit` as keyof ActSummary;
  if (summary[debitKey] == null && amounts[0] != null) summary[debitKey] = amounts[0];
  if (summary[creditKey] == null && amounts[1] != null) summary[creditKey] = amounts[1];
}

function fillSummary(target: ActSummary, extra: ActSummary) {
  (Object.keys(target) as (keyof ActSummary)[]).forEach((key) => {
    if (target[key] == null && extra[key] != null) target[key] = extra[key];
  });
}

function splitSides(text: string): [string, string] | null {
  const dates = [...text.matchAll(/\d{2}[./]\d{2}[./]\d{2,4}/g)];
  if (dates.length < 2 || dates[1].index == null) return null;
  const between = text.slice((dates[0].index ?? 0) + dates[0][0].length, dates[1].index);
  if (!/[a-zа-яё]/i.test(between)) return null;
  const left = text.slice(0, dates[1].index).trim();
  const right = text.slice(dates[1].index).trim();
  if (moneyAmounts(left).length === 0 || moneyAmounts(right).length === 0) return null;
  return [left, right];
}

function invoiceNumber(text: string): string | null {
  return docFromDocumentCell(text);
}

function firstDate(text: string): string | null {
  const match = text.match(/\d{2}[./]\d{2}[./]\d{2,4}/);
  return match ? parseDate(match[0]) : null;
}

function moneyAmounts(text: string): number[] {
  const withoutDates = text.replace(/\d{2}[./]\d{2}[./]\d{2,4}/g, " ");
  const matches = withoutDates.match(/\d{1,3}(?:[ \u00a0]\d{3})+(?:[.,]\d{2})|\d+[.,]\d{2}/g) ?? [];
  return matches.map((token) => parseMoney(token)).filter((value): value is number => value != null);
}

function emptyAct(): ParsedAct {
  return {
    title: null,
    period: null,
    partyA: null,
    partyB: null,
    lines: [] as ActLine[],
    summary: emptySummary(),
    theirsSummary: null,
    warnings: [],
    confidence: null,
    rawText: "",
    twoSided: false,
  };
}
