import type { ActLine, ActSummary, ParsedAct } from "@/lib/types";
import { emptySummary } from "@/lib/excel-act";
import { extractDocNumber, parseDate, parseMoney, round2 } from "@/lib/text";

export type QwenPage = {
  newAct: boolean;
  partyLeft: string | null;
  partyRight: string | null;
  period: string | null;
  rows: QwenRow[];
  left: ActSummary;
  right: ActSummary;
  note: string | null;
};

export type QwenRow = {
  date: string | null;
  document: string;
  number: string | null;
  leftDebit: number | null;
  leftCredit: number | null;
  rightDebit: number | null;
  rightCredit: number | null;
};

export const QWEN_PROMPT = `Это страница акта сверки двух организаций. Слева одна колонка, справа другая. Верни только JSON без пояснений:
{"newAct":true,"partyLeft":"","partyRight":"","period":"","rows":[{"date":"12.01.2026","document":"Продажа (1 от 12.01.2026)","number":"1","leftDebit":471487.25,"leftCredit":null,"rightDebit":null,"rightCredit":471487.25}],"left":{"openingDebit":null,"openingCredit":null,"turnoverDebit":null,"turnoverCredit":null,"closingDebit":null,"closingCredit":null},"right":{"openingDebit":null,"openingCredit":null,"turnoverDebit":null,"turnoverCredit":null,"closingDebit":null,"closingCredit":null}}
newAct равен true только если страница начинается с заголовка «Акт сверки». Продолжение таблицы — false. Суммы — числа с точкой, пустая ячейка — null. Не пропускай строки и не складывай соседние суммы. Итоги сальдо и оборотов положи в left и right.`;

export function parseQwenResponse(raw: string): QwenPage {
  const start = raw.indexOf("{");
  const end = raw.lastIndexOf("}");
  if (start < 0 || end <= start) {
    throw new Error("Модель не вернула таблицу акта. Повторите чтение страницы.");
  }
  let data: unknown;
  try {
    data = JSON.parse(raw.slice(start, end + 1));
  } catch {
    throw new Error("Модель оборвала таблицу. Повторите чтение страницы.");
  }
  if (!data || typeof data !== "object") {
    throw new Error("Модель не вернула таблицу акта. Повторите чтение страницы.");
  }
  const record = data as Record<string, unknown>;
  const rows = Array.isArray(record.rows) ? record.rows : Array.isArray(record.lines) ? record.lines : [];
  return {
    newAct: record.newAct !== false,
    partyLeft: textOrNull(record.partyLeft ?? record.partyA),
    partyRight: textOrNull(record.partyRight ?? record.partyB),
    period: textOrNull(record.period),
    rows: rows.map(normalizeRow),
    left: normalizeSummary(record.left),
    right: normalizeSummary(record.right),
    note: textOrNull(record.note),
  };
}

export function qwenPagesToRecognition(pages: QwenPage[]): { act: ParsedAct; documents: ParsedAct[] } {
  const documents: ParsedAct[] = [];
  pages.forEach((page, index) => {
    const next = pageToAct(page, `${index}`);
    if (documents.length === 0 || (index > 0 && page.newAct)) documents.push(next);
    else appendAct(documents[documents.length - 1], next);
  });
  if (documents.length === 0) {
    const empty = pageToAct(
      {
        newAct: true,
        partyLeft: null,
        partyRight: null,
        period: null,
        rows: [],
        left: emptySummary(),
        right: emptySummary(),
        note: null,
      },
      "0",
    );
    return { act: empty, documents: [empty] };
  }
  const act = documents.length === 1 ? documents[0] : documents.map(cloneAct).reduce(appendAct);
  return { act, documents };
}

function cloneAct(act: ParsedAct): ParsedAct {
  return {
    ...act,
    lines: act.lines.map((line) => ({ ...line })),
    summary: { ...act.summary },
    theirsSummary: act.theirsSummary ? { ...act.theirsSummary } : null,
    warnings: [...act.warnings],
  };
}

function pageToAct(page: QwenPage, prefix: string): ParsedAct {
  const lines: ActLine[] = [];
  page.rows.forEach((row, index) => {
    pushSide(lines, "ours", row, row.leftDebit, row.leftCredit, `${prefix}-l-${index}`);
    pushSide(lines, "theirs", row, row.rightDebit, row.rightCredit, `${prefix}-r-${index}`);
  });
  const rightHasSummary = hasAmount(page.right);
  const warnings: string[] = [];
  if (page.rows.length === 0) warnings.push("На странице модель не нашла строк таблицы.");
  if (page.note) warnings.push(page.note);
  return {
    title: null,
    period: page.period,
    partyA: page.partyLeft,
    partyB: page.partyRight,
    lines,
    summary: page.left,
    theirsSummary: rightHasSummary || lines.some((line) => line.side === "theirs") ? page.right : null,
    warnings,
    confidence: null,
    rawText: rawText(page),
    twoSided: lines.some((line) => line.side === "theirs") || rightHasSummary,
  };
}

function appendAct(target: ParsedAct, extra: ParsedAct): ParsedAct {
  const offset = target.lines.length;
  target.lines.push(
    ...extra.lines.map((line, index) => ({ ...line, id: `${offset + index}-${line.id}` })),
  );
  fillSummary(target.summary, extra.summary);
  if (extra.theirsSummary) {
    target.theirsSummary ??= emptySummary();
    fillSummary(target.theirsSummary, extra.theirsSummary);
  }
  target.period ??= extra.period;
  target.partyA ??= extra.partyA;
  target.partyB ??= extra.partyB;
  target.rawText = [target.rawText, extra.rawText].filter(Boolean).join("\n\n");
  target.warnings.push(...extra.warnings.filter((warning) => !target.warnings.includes(warning)));
  target.twoSided = target.twoSided || extra.twoSided;
  return target;
}

function pushSide(
  lines: ActLine[],
  side: "ours" | "theirs",
  row: QwenRow,
  debit: number | null,
  credit: number | null,
  id: string,
) {
  if (debit == null && credit == null) return;
  const number = row.number || extractDocNumber(row.document);
  lines.push({
    id,
    side,
    date: parseDate(row.date ?? ""),
    docNumber: number,
    description: row.document,
    debit,
    credit,
  });
}

function normalizeRow(value: unknown): QwenRow {
  const row = value && typeof value === "object" ? (value as Record<string, unknown>) : {};
  const document = textOrNull(row.document ?? row.description) ?? "";
  return {
    date: textOrNull(row.date),
    document,
    number: textOrNull(row.number ?? row.docNumber),
    leftDebit: asMoney(row.leftDebit ?? row.debit),
    leftCredit: asMoney(row.leftCredit ?? row.credit),
    rightDebit: asMoney(row.rightDebit),
    rightCredit: asMoney(row.rightCredit),
  };
}

function normalizeSummary(value: unknown): ActSummary {
  const source = value && typeof value === "object" ? (value as Record<string, unknown>) : {};
  return {
    openingDebit: asMoney(source.openingDebit),
    openingCredit: asMoney(source.openingCredit),
    turnoverDebit: asMoney(source.turnoverDebit),
    turnoverCredit: asMoney(source.turnoverCredit),
    closingDebit: asMoney(source.closingDebit),
    closingCredit: asMoney(source.closingCredit),
  };
}

function asMoney(value: unknown): number | null {
  if (typeof value === "number" && Number.isFinite(value)) return round2(value);
  if (typeof value === "string") return parseMoney(value);
  return null;
}

function textOrNull(value: unknown): string | null {
  if (typeof value !== "string") return null;
  const text = value.replace(/\s+/g, " ").trim();
  return text || null;
}

function hasAmount(summary: ActSummary): boolean {
  return Object.values(summary).some((value) => value != null);
}

function fillSummary(target: ActSummary, extra: ActSummary) {
  (Object.keys(target) as (keyof ActSummary)[]).forEach((key) => {
    if (target[key] == null && extra[key] != null) target[key] = extra[key];
  });
}

function rawText(page: QwenPage): string {
  const head = [page.partyLeft, page.partyRight, page.period].filter(Boolean).join(" · ");
  const rows = page.rows.map((row) =>
    [
      row.date,
      row.document,
      row.leftDebit != null ? `лев.деб ${row.leftDebit}` : "",
      row.leftCredit != null ? `лев.кред ${row.leftCredit}` : "",
      row.rightDebit != null ? `прав.деб ${row.rightDebit}` : "",
      row.rightCredit != null ? `прав.кред ${row.rightCredit}` : "",
    ]
      .filter(Boolean)
      .join(" "),
  );
  return [head, ...rows].filter(Boolean).join("\n");
}
