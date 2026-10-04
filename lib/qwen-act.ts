import type { ActLine, ActSummary, ParsedAct } from "@/lib/types";
import { emptySummary } from "@/lib/excel-act";
import { extractDocNumber, parseDate, parseMoney, round2, summaryKind } from "@/lib/text";

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

export const QWEN_PROMPT = `Перед тобой страница акта сверки взаимных расчетов. Перепиши таблицу в JSON.

Правила:
- Перепиши каждую строку таблицы по порядку сверху вниз, ровно один раз. Ничего не придумывай и не повторяй.
- date — дата строки, как на картинке.
- document — название операции, как написано в строке.
- number — номер документа из этой строки, все цифры и дефисы подряд. Если номера нет, null.
- В акте бывает одна таблица сумм или две рядом: слева одна организация, справа другая. Суммы левой таблицы — leftDebit и leftCredit. Суммы правой — rightDebit и rightCredit. Если таблица одна, right-поля всегда null.
- Сумму пиши числом: без пробелов между разрядами, копейки через точку. Пустая ячейка — null.
- Переписывай только печатный текст. Рукописные пометки, подписи и печати не переписывай.
- Сальдо и обороты не пиши в rows. Входящее или начальное сальдо — opening, обороты и «Итого за период» — turnover, исходящее или конечное сальдо — closing. Суммы левой таблицы — в left, правой — в right.
- partyLeft и partyRight — названия организаций из шапки, period — период акта, иначе null.
- newAct — true, если вверху страницы заголовок «Акт сверки» и шапка с организациями. Если страница начинается с «Лист 2», «Лист 3» или продолжает таблицу, false.`;

const MONEY = { type: ["number", "null"] };
const TEXT = { type: ["string", "null"] };
const SUMMARY_SCHEMA = {
  type: "object",
  properties: {
    openingDebit: MONEY,
    openingCredit: MONEY,
    turnoverDebit: MONEY,
    turnoverCredit: MONEY,
    closingDebit: MONEY,
    closingCredit: MONEY,
  },
};

export const QWEN_SCHEMA = {
  type: "object",
  properties: {
    newAct: { type: "boolean" },
    partyLeft: TEXT,
    partyRight: TEXT,
    period: TEXT,
    left: SUMMARY_SCHEMA,
    right: SUMMARY_SCHEMA,
    rows: {
      type: "array",
      items: {
        type: "object",
        properties: {
          date: TEXT,
          document: { type: "string" },
          number: TEXT,
          leftDebit: MONEY,
          leftCredit: MONEY,
          rightDebit: MONEY,
          rightCredit: MONEY,
        },
        required: ["date", "document", "number", "leftDebit", "leftCredit", "rightDebit", "rightCredit"],
      },
    },
  },
  required: ["newAct", "left", "right", "rows"],
};

export function looksLooping(content: string): boolean {
  if (content.length < 1600) return false;
  const tail = content.slice(-400).replace(/\s+/g, " ");
  const flat = content.replace(/\s+/g, " ");
  let count = 0;
  for (let index = flat.indexOf(tail); index >= 0; index = flat.indexOf(tail, index + 1)) {
    count += 1;
    if (count >= 3) return true;
  }
  return false;
}

function parseLenient(text: string): unknown {
  try {
    return JSON.parse(text);
  } catch {
    // The model may stop mid-row (length limit or a loop we cut off). Keep every complete row.
    for (let end = text.lastIndexOf("}"), tries = 0; end > 0 && tries < 400; end = text.lastIndexOf("}", end - 1), tries += 1) {
      for (const tail of ["]}", "}", "]}}", ""]) {
        try {
          return JSON.parse(text.slice(0, end + 1) + tail);
        } catch {
          // try a shorter prefix
        }
      }
    }
    throw new Error("unparsable");
  }
}

export function parseQwenResponse(raw: string): QwenPage {
  const start = raw.indexOf("{");
  if (start < 0) {
    throw new Error("Модель не вернула таблицу акта. Повторите чтение страницы.");
  }
  let data: unknown;
  try {
    data = parseLenient(raw.slice(start).replace(/\s*```\s*$/, ""));
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
    rows: dropRepeats(rows.map(normalizeRow)).filter((row) => !/акт\s+сверки/i.test(row.document)),
    left: normalizeSummary(record.left),
    right: normalizeSummary(record.right),
    note: textOrNull(record.note),
  };
}

export function qwenPagesToRecognition(pages: QwenPage[]): { act: ParsedAct; documents: ParsedAct[] } {
  const documents: ParsedAct[] = [];
  pages.forEach((page, index) => {
    const next = pageToAct(page, `${index}`);
    const startsAct = index > 0 && page.newAct && next.lines.length > 0;
    if (documents.length === 0 || startsAct) documents.push(next);
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
  const left = { ...page.left };
  const right = { ...page.right };
  page.rows.forEach((row, index) => {
    const balance = balanceKind(row.document, lines.length === 0);
    if (balance) {
      applyBalance(left, balance, row.leftDebit, row.leftCredit);
      applyBalance(right, balance, row.rightDebit, row.rightCredit);
      return;
    }
    pushSide(lines, "ours", row, row.leftDebit, row.leftCredit, `${prefix}-l-${index}`);
    pushSide(lines, "theirs", row, row.rightDebit, row.rightCredit, `${prefix}-r-${index}`);
  });
  const rightHasSummary =
    hasAmount(right) && (lines.some((line) => line.side === "theirs") || !echoesLeft(right, left));
  const warnings: string[] = [];
  if (page.rows.length === 0) warnings.push("На странице модель не нашла строк таблицы.");
  if (page.note) warnings.push(page.note);
  return {
    title: null,
    period: page.period,
    partyA: page.partyLeft,
    partyB: page.partyRight,
    lines,
    summary: left,
    theirsSummary: rightHasSummary || lines.some((line) => line.side === "theirs") ? right : null,
    warnings,
    confidence: null,
    rawText: rawText(page),
    twoSided: lines.some((line) => line.side === "theirs") || rightHasSummary,
  };
}

function echoesLeft(right: ActSummary, left: ActSummary): boolean {
  return (Object.keys(right) as (keyof ActSummary)[]).every((key) => {
    const value = right[key];
    return value == null || value === 0 || value === left[key];
  });
}

function balanceKind(document: string, beforeLines: boolean): "opening" | "turnover" | "closing" | null {
  const kind = summaryKind(document);
  if (kind) return kind;
  if (!/сальдо/i.test(document)) return null;
  return beforeLines ? "opening" : "closing";
}

function applyBalance(
  summary: ActSummary,
  kind: "opening" | "turnover" | "closing",
  debit: number | null,
  credit: number | null,
) {
  const debitKey = `${kind}Debit` as keyof ActSummary;
  const creditKey = `${kind}Credit` as keyof ActSummary;
  if (debit != null && summary[debitKey] == null) summary[debitKey] = debit;
  if (credit != null && summary[creditKey] == null) summary[creditKey] = credit;
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

// Three or more identical rows in a row is the model looping, not real data: keep one.
function dropRepeats(rows: QwenRow[]): QwenRow[] {
  const keys = rows.map((row) => JSON.stringify(row));
  return rows.filter((_, index) => {
    if (index === 0 || keys[index] !== keys[index - 1]) return true;
    let start = index;
    while (start > 0 && keys[start - 1] === keys[index]) start -= 1;
    let end = index;
    while (end + 1 < keys.length && keys[end + 1] === keys[index]) end += 1;
    return end - start + 1 < 3;
  });
}

function docNumberText(value: unknown): string | null {
  const text = textOrNull(value);
  if (!text) return null;
  if (/^\d{1,3}(?:[\s\u00a0]\d{3})+(?:[,.]\d{2})?$|^\d+[,.]\d{2}$/.test(text)) return null;
  return text;
}

function normalizeRow(value: unknown): QwenRow {
  const row = value && typeof value === "object" ? (value as Record<string, unknown>) : {};
  const document = textOrNull(row.document ?? row.description) ?? "";
  return {
    date: textOrNull(row.date),
    document,
    number: docNumberText(row.number ?? row.docNumber),
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

// The opening balance is printed on the first page of an act and the totals on the last one.
// The model fills fields it cannot see with 0 or a guess, so trust each figure only from its page.
function fillSummary(target: ActSummary, extra: ActSummary) {
  (Object.keys(target) as (keyof ActSummary)[]).forEach((key) => {
    const value = extra[key];
    if (value == null) return;
    if (target[key] == null) target[key] = value;
    else if (!key.startsWith("opening") && value !== 0) target[key] = value;
  });
}

function rawText(page: QwenPage): string {
  const head = [page.partyLeft, page.partyRight, page.period].filter(Boolean).join(" · ");
  const rows = page.rows.map((row) =>
    [
      row.date,
      row.number ? `№${row.number}` : "",
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
