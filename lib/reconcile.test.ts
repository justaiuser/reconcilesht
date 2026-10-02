import { describe, it } from "node:test";
import assert from "node:assert/strict";
import { parseMoney, parseDate, extractDocNumber, normalizeDoc } from "./text";
import { detectColumns, parseExcelMatrix } from "./excel-act";
import { parseOcrWords } from "./ocr-table";
import { compareActs } from "./compare";
import { buildExcelMatrix, buildSampleScanAct } from "./sample-data";
import type { OcrWord } from "./types";

describe("деньги и даты", () => {
  it("понимает русскую запись суммы", () => {
    assert.equal(parseMoney("120 000,50"), 120000.5);
    assert.equal(parseMoney("1.250.000,50"), 1250000.5);
    assert.equal(parseMoney("(1 000,00)"), -1000);
    assert.equal(parseMoney("—"), null);
    assert.equal(parseMoney("45500"), 45500);
  });

  it("понимает дату акта", () => {
    assert.equal(parseDate("03.02.2026"), "2026-02-03");
    assert.equal(parseDate("3/2/26"), "2026-02-03");
  });

  it("достаёт номер документа", () => {
    assert.equal(extractDocNumber("Реализация №15"), "15");
    assert.equal(extractDocNumber("Оплата п/п №42"), "42");
    assert.equal(normalizeDoc("015"), normalizeDoc("15"));
  });
});

describe("Excel акта сверки", () => {
  it("находит две колонки и строки примера", () => {
    const matrix = buildExcelMatrix();
    const map = detectColumns(matrix);
    assert.ok(map);
    assert.equal(map.headerRow, 5);
    assert.equal(map.ours.date, 0);
    assert.equal(map.ours.debit, 2);
    assert.equal(map.theirs?.date, 4);
    assert.equal(map.theirs?.credit, 7);

    const act = parseExcelMatrix(matrix, map);
    assert.equal(act.partyA, 'ООО "Север"');
    assert.equal(act.partyB, 'ООО "Луч"');
    assert.equal(act.period, "1 квартал 2026 г.");
    assert.equal(act.lines.filter((line) => line.side === "ours").length, 5);
    assert.equal(act.lines.filter((line) => line.side === "theirs").length, 5);
    assert.equal(act.summary.openingDebit, 10000);
    assert.equal(act.summary.turnoverDebit, 173700);
    assert.equal(act.summary.closingDebit, 43700);
    assert.equal(act.theirsSummary?.openingCredit, 10000);
    assert.equal(act.theirsSummary?.closingCredit, 43700);
  });
});

describe("разбор скана по координатам слов", () => {
  it("разносит слова по дебету и кредиту двух сторон", () => {
    const words: OcrWord[] = [
      word("Дата", 40, 100),
      word("Документ", 220, 100),
      word("Дебет", 520, 100),
      word("Кредит", 700, 100),
      word("Дата", 900, 100),
      word("Документ", 1100, 100),
      word("Дебет", 1420, 100),
      word("Кредит", 1620, 100),
      word("12.01.2026", 40, 160),
      word("Реализация", 220, 160),
      word("№15", 340, 160),
      word("120000,00", 500, 160),
      word("12.01.2026", 900, 160),
      word("Поступление", 1100, 160),
      word("№15", 1260, 160),
      word("120000,00", 1600, 160),
      word("Сальдо", 40, 220),
      word("конечное", 140, 220),
      word("43700,00", 500, 220),
      word("43700,00", 1600, 220),
    ];
    const act = parseOcrWords(words, 'между ООО "Север" и ООО "Луч"');
    const ours = act.lines.find((line) => line.side === "ours");
    const theirs = act.lines.find((line) => line.side === "theirs");
    assert.equal(ours?.docNumber, "15");
    assert.equal(ours?.debit, 120000);
    assert.equal(ours?.credit, null);
    assert.equal(theirs?.docNumber, "15");
    assert.equal(theirs?.credit, 120000);
    assert.equal(act.summary.closingDebit, 43700);
    assert.equal(act.theirsSummary?.closingCredit, 43700);
    assert.equal(act.partyB, 'ООО "Луч"');
  });
});

describe("сравнение", () => {
  it("показывает расхождения примера в колонке контрагента", () => {
    const matrix = buildExcelMatrix();
    const map = detectColumns(matrix);
    assert.ok(map);
    const excel = parseExcelMatrix(matrix, map);
    const scan = buildSampleScanAct();
    const result = compareActs(excel, scan);
    assert.equal(result.reports.length, 2);

    const ours = result.reports[0];
    assert.equal(ours.pairs.every((pair) => pair.status === "match"), true);
    assert.equal(ours.summary.every((row) => row.status === "match"), true);

    const theirs = result.reports[1];
    const byDoc = (number: string) =>
      theirs.pairs.find((pair) => pair.excel?.docNumber === number || pair.scan?.docNumber === number);
    assert.equal(byDoc("15")?.status, "match");
    assert.equal(byDoc("42")?.status, "match");
    assert.equal(byDoc("27")?.status, "amount");
    assert.equal(byDoc("88")?.status, "amount");
    assert.equal(byDoc("31")?.status, "excel-only");
    assert.equal(byDoc("5")?.status, "scan-only");
    assert.equal(theirs.summary.some((row) => row.label.startsWith("Сальдо конечное") && row.status === "diff"), true);
  });
});

function word(text: string, x: number, y: number): OcrWord {
  return { text, confidence: 90, x0: x, y0: y, x1: x + text.length * 12, y1: y + 24 };
}
