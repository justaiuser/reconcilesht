import type { ActLine, ActSummary, Cell, ParsedAct } from "@/lib/types";
import { parseDate, round2 } from "@/lib/text";

export type SampleLine = {
  date: string;
  description: string;
  debit: number | null;
  credit: number | null;
};

export const EXCEL_OURS: SampleLine[] = [
  { date: "12.01.2026", description: "Реализация №15", debit: 120000, credit: null },
  { date: "28.01.2026", description: "Оплата п/п №42", debit: null, credit: 100000 },
  { date: "03.02.2026", description: "Реализация №27", debit: 45500, credit: null },
  { date: "19.03.2026", description: "Оплата п/п №88", debit: null, credit: 40000 },
  { date: "25.03.2026", description: "Реализация №31", debit: 8200, credit: null },
];

export const EXCEL_THEIRS: SampleLine[] = [
  { date: "12.01.2026", description: "Поступление №15", debit: null, credit: 120000 },
  { date: "28.01.2026", description: "Оплата п/п №42", debit: 100000, credit: null },
  { date: "03.02.2026", description: "Поступление №27", debit: null, credit: 45500 },
  { date: "19.03.2026", description: "Оплата п/п №88", debit: 40000, credit: null },
  { date: "25.03.2026", description: "Поступление №31", debit: null, credit: 8200 },
];

export const SCAN_OURS: SampleLine[] = EXCEL_OURS;

export const SCAN_THEIRS: SampleLine[] = [
  { date: "12.01.2026", description: "Поступление №15", debit: null, credit: 120000 },
  { date: "28.01.2026", description: "Оплата п/п №42", debit: 100000, credit: null },
  { date: "03.02.2026", description: "Поступление №27", debit: null, credit: 45000 },
  { date: "02.03.2026", description: "Корректировка №5", debit: null, credit: 3000 },
  { date: "19.03.2026", description: "Оплата п/п №88", debit: 45500, credit: null },
];

export const SAMPLE_PARTY_A = 'ООО "Север"';
export const SAMPLE_PARTY_B = 'ООО "Луч"';
export const SAMPLE_PERIOD = "1 квартал 2026 г.";

export function sideTotals(
  openingDebit: number,
  openingCredit: number,
  lines: SampleLine[],
): ActSummary {
  const turnoverDebit = round2(lines.reduce((sum, line) => sum + (line.debit ?? 0), 0));
  const turnoverCredit = round2(lines.reduce((sum, line) => sum + (line.credit ?? 0), 0));
  const net = round2(openingDebit - openingCredit + turnoverDebit - turnoverCredit);
  return {
    openingDebit: openingDebit || null,
    openingCredit: openingCredit || null,
    turnoverDebit,
    turnoverCredit,
    closingDebit: net > 0 ? net : net === 0 ? 0 : null,
    closingCredit: net < 0 ? round2(-net) : null,
  };
}

export function buildExcelMatrix(): Cell[][] {
  const ours = sideTotals(10000, 0, EXCEL_OURS);
  const theirs = sideTotals(0, 10000, EXCEL_THEIRS);
  const width = Math.max(EXCEL_OURS.length, EXCEL_THEIRS.length);
  const data = Array.from({ length: width }, (_, index) => {
    const left = EXCEL_OURS[index];
    const right = EXCEL_THEIRS[index];
    return [
      left?.date ?? null,
      left?.description ?? null,
      left?.debit ?? null,
      left?.credit ?? null,
      right?.date ?? null,
      right?.description ?? null,
      right?.debit ?? null,
      right?.credit ?? null,
    ] satisfies Cell[];
  });

  return [
    ["Акт сверки взаимных расчетов"],
    [`за период: ${SAMPLE_PERIOD}`],
    [`между ${SAMPLE_PARTY_A} и ${SAMPLE_PARTY_B}`],
    [],
    [`По данным ${SAMPLE_PARTY_A}, руб.`, null, null, null, `По данным ${SAMPLE_PARTY_B}, руб.`],
    ["Дата", "Документ", "Дебет", "Кредит", "Дата", "Документ", "Дебет", "Кредит"],
    [
      "Сальдо начальное",
      null,
      ours.openingDebit,
      ours.openingCredit,
      null,
      null,
      theirs.openingDebit,
      theirs.openingCredit,
    ],
    ...data,
    [
      "Обороты за период",
      null,
      ours.turnoverDebit,
      ours.turnoverCredit,
      null,
      null,
      theirs.turnoverDebit,
      theirs.turnoverCredit,
    ],
    [
      "Сальдо конечное",
      null,
      ours.closingDebit,
      ours.closingCredit,
      null,
      null,
      theirs.closingDebit,
      theirs.closingCredit,
    ],
  ];
}

function toActLines(lines: SampleLine[], side: ActLine["side"]): ActLine[] {
  return lines.map((line, index) => ({
    id: `${side}-${index}`,
    side,
    date: parseDate(line.date),
    docNumber: line.description.match(/№\s*(\d+)/)?.[1] ?? null,
    description: line.description,
    debit: line.debit,
    credit: line.credit,
  }));
}

export function buildSampleScanAct(): ParsedAct {
  return {
    title: "Акт сверки взаимных расчетов",
    period: SAMPLE_PERIOD,
    partyA: SAMPLE_PARTY_A,
    partyB: SAMPLE_PARTY_B,
    lines: [...toActLines(SCAN_OURS, "ours"), ...toActLines(SCAN_THEIRS, "theirs")],
    summary: sideTotals(10000, 0, SCAN_OURS),
    theirsSummary: sideTotals(0, 10000, SCAN_THEIRS),
    warnings: [],
    confidence: 90,
    rawText: "",
    twoSided: true,
  };
}
