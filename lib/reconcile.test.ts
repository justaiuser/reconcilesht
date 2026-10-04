import { describe, it } from "node:test";
import assert from "node:assert/strict";
import { parseMoney, parseDate, extractDocNumber, docFromDocumentCell, normalizeDoc } from "./text";
import { detectColumns, parseExcelMatrix } from "./excel-act";
import { groupLines, parseOcrDocuments, parseOcrPages, parseOcrWords } from "./ocr-table";
import { pageStrips } from "./page-strips";
import { compareActs, compareDocuments } from "./compare";
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
    assert.equal(extractDocNumber("Продажа (1 от 12.01.2026)"), "1");
    assert.equal(docFromDocumentCell("1800065755-254"), "1800065755-254");
    assert.equal(docFromDocumentCell("1800065755 1800065755-254"), "1800065755-254");
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

describe("односторонний скан с двумя датами", () => {
  it("берёт номер счёта, дебет и входящее сальдо со следующей строки", () => {
    const words: OcrWord[] = [
      word("Дата", 420, 80),
      word("Дата", 580, 80),
      word("Документа", 40, 112),
      word("документа", 220, 112),
      word("дебету", 820, 112),
      word("кредиту", 1040, 112),
      word("Входящее", 40, 160),
      word("сальдо", 160, 160),
      word("0,00", 800, 200),
      word("0,00", 1040, 200),
      word("1800065755", 40, 250),
      word("1800065755-254", 220, 250),
      word("24.09.2025", 420, 250),
      word("24.09.2025", 580, 250),
      word("1", 760, 250),
      word("330", 800, 250),
      word("442,88", 860, 250),
      word("Обороты", 40, 320),
      word("по", 160, 320),
      word("дебету", 220, 320),
      word("26", 740, 360),
      word("306", 800, 360),
      word("875,68", 860, 360),
      word("0,00", 1040, 360),
    ];
    const act = parseOcrWords(words, "");
    assert.equal(act.twoSided, false);
    assert.equal(act.lines.length, 1);
    assert.equal(act.lines[0]?.docNumber, "1800065755-254");
    assert.equal(act.lines[0]?.date, "2025-09-24");
    assert.equal(act.lines[0]?.debit, 1330442.88);
    assert.equal(act.lines[0]?.credit, null);
    assert.equal(act.summary.openingDebit, 0);
    assert.equal(act.summary.openingCredit, 0);
    assert.equal(act.summary.turnoverDebit, 26306875.68);
    assert.equal(act.summary.turnoverCredit, 0);
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

  it("зеркалит кредит Excel и дебет скана по номеру счёта", () => {
    const matrix = [
      ["Дата", "Документ", "Дебет", "Кредит"],
      ["Сальдо начальное", null, null, 26306875.68],
      ["24.09.2025", "1800065755-254", null, 1330442.88],
      ["на 30.09.2026 задолженность 517 309 230,15"],
      ["Сальдо конечное", null, null, 26306875.68],
    ];
    const map = detectColumns(matrix);
    assert.ok(map);
    const excel = parseExcelMatrix(matrix, map);
    assert.equal(excel.lines.length, 1);
    assert.equal(excel.lines[0]?.docNumber, "1800065755-254");
    assert.equal(excel.lines[0]?.credit, 1330442.88);

    const scan = parseOcrWords(
      [
        word("Дата", 200, 40),
        word("Документ", 40, 70),
        word("дебету", 400, 70),
        word("кредиту", 620, 70),
        word("1800065755-254", 40, 120),
        word("24.09.2025", 200, 120),
        word("1330442,88", 400, 120),
      ],
      "",
    );
    const result = compareActs(excel, scan);
    assert.equal(result.reports[0]?.mirror, true);
    assert.equal(result.reports[0]?.pairs[0]?.status, "match");
  });

  it("сверяет левую колонку скана с правой без Excel", () => {
    const act = parseOcrWords(
      [
        word("Дата", 40, 80),
        word("Документ", 180, 80),
        word("Дебет", 420, 80),
        word("Кредит", 620, 80),
        word("Дата", 820, 80),
        word("Документ", 980, 80),
        word("Дебет", 1240, 80),
        word("Кредит", 1460, 80),
        word("12.01.2026", 40, 140),
        word("Продажа", 180, 140),
        word("№1", 280, 140),
        word("471487,25", 400, 140),
        word("12.01.2026", 820, 140),
        word("Приход", 980, 140),
        word("№1", 1100, 140),
        word("471487,25", 1440, 140),
        word("20.01.2026", 40, 190),
        word("Продажа", 180, 190),
        word("№2", 280, 190),
        word("1000,00", 400, 190),
        word("20.01.2026", 820, 190),
        word("Приход", 980, 190),
        word("№2", 1100, 190),
        word("900,00", 1440, 190),
      ],
      "",
    );
    const result = compareDocuments([act]);
    assert.ok(result);
    assert.equal(result.columnsOnly, true);
    assert.equal(result.labels.left, "Левая колонка");
    assert.equal(result.reports.length, 1);
    assert.equal(result.reports[0]?.mirror, true);
    assert.equal(result.reports[0]?.title, "Левая и правая колонки");
    const first = result.reports[0]?.pairs.find((pair) => pair.excel?.docNumber === "1");
    const second = result.reports[0]?.pairs.find((pair) => pair.excel?.docNumber === "2");
    assert.equal(first?.status, "match");
    assert.equal(second?.status, "amount");
    assert.equal(second?.notes[0], "Дебет левой колонки не равен кредиту правой");
  });
});

describe("полосы плотной страницы", () => {
  it("режет частые строки и оставляет редкую страницу целиком", () => {
    const dense = paintRows(48, 220, 20, 3);
    const strips = pageStrips(dense, 48, 220, 1);
    assert.ok(strips.length >= 3);
    for (let index = 1; index < strips.length; index += 1) {
      assert.ok(strips[index].top >= strips[index - 1].bottom);
    }
    assert.equal(strips[0].top >= 0, true);
    assert.equal(strips[strips.length - 1].bottom <= 220, true);

    const sparse = paintRows(48, 220, 4, 28);
    assert.equal(pageStrips(sparse, 48, 220, 1).length, 1);
  });

  it("не склеивает строки, которые распознаватель уже разделил", () => {
    const lines = groupLines([
      word("12.01.2026", 40, 100, 1),
      word("471487,25", 400, 108, 1),
      word("471487,25", 1400, 110, 1),
      word("13.01.2026", 40, 112, 2),
      word("1000,00", 400, 116, 2),
      word("900,00", 1400, 118, 2),
    ]);
    assert.equal(lines.length, 2);
    assert.equal(lines[0].some((item) => item.text.startsWith("471487")), true);
    assert.equal(lines[1].some((item) => item.text.startsWith("1000")), true);
  });
});

describe("две колонки в одном скане", () => {
  it("делит стороны, даже если левая дата в шапке не прочиталась", () => {
    const words: OcrWord[] = [
      word("Документ", 180, 80),
      word("Дебет", 420, 80),
      word("Кредит", 620, 80),
      word("Дата", 820, 78),
      word("Документ", 980, 78),
      word("Дебет", 1240, 78),
      word("Кредит", 1460, 78),
      word("12.01.2026", 40, 140),
      word("Продажа", 170, 140),
      word("№1", 250, 140),
      word("471487,25", 430, 140),
      word("12.01.2026", 820, 140),
      word("Приход", 980, 140),
      word("№1", 1120, 140),
      word("471487,25", 1460, 140),
    ];
    const act = parseOcrWords(words, "");
    assert.equal(act.twoSided, true);
    const left = act.lines.find((line) => line.side === "ours");
    const right = act.lines.find((line) => line.side === "theirs");
    assert.equal(left?.docNumber, "1");
    assert.equal(left?.date, "2026-01-12");
    assert.equal(left?.debit, 471487.25);
    assert.equal(right?.credit, 471487.25);
    const compared = compareDocuments([act]);
    assert.equal(compared?.reports[0]?.pairs[0]?.status, "match");
  });

  it("продолжает колонки на следующей странице без шапки", () => {
    const header = twoSidedPage("15", "120000,00");
    const continuation: OcrWord[] = [
      word("03.02.2026", 40, 80),
      word("Продажа", 180, 80),
      word("№16", 280, 80),
      word("5000,00", 400, 80),
      word("03.02.2026", 820, 80),
      word("Приход", 980, 80),
      word("№16", 1100, 80),
      word("5000,00", 1440, 80),
    ];
    const act = parseOcrPages([header, continuation], "");
    assert.equal(act.twoSided, true);
    assert.equal(act.lines.filter((line) => line.docNumber === "16" && line.side === "ours").length, 1);
    assert.equal(act.lines.filter((line) => line.docNumber === "16" && line.side === "theirs").length, 1);
  });
});

describe("несколько актов в одном скане", () => {
  it("начинает новый акт, когда следующая страница снова открывается заголовком", () => {
    const first = twoSidedPage("15", "120000,00");
    const second = twoSidedPage("99", "5000,00");
    const documents = parseOcrDocuments(
      [first, second],
      ["продолжение таблицы без заголовка", "Акт сверки взаимных расчетов между ООО Альфа и ООО Бета"],
    );
    assert.equal(documents.length, 2);
    assert.equal(documents[0]?.lines.some((line) => line.docNumber === "15"), true);
    assert.equal(documents[0]?.lines.some((line) => line.docNumber === "99"), false);
    assert.equal(documents[1]?.lines.some((line) => line.docNumber === "99"), true);

    const compared = compareDocuments(documents);
    assert.equal(compared?.reports.length, 2);
    assert.match(compared?.reports[0]?.title ?? "", /^Акт 1/);
    assert.match(compared?.reports[1]?.title ?? "", /^Акт 2/);
  });
});

function twoSidedPage(number: string, amount: string): OcrWord[] {
  return [
    word("Дата", 40, 80),
    word("Документ", 180, 80),
    word("Дебет", 420, 80),
    word("Кредит", 620, 80),
    word("Дата", 820, 80),
    word("Документ", 980, 80),
    word("Дебет", 1240, 80),
    word("Кредит", 1460, 80),
    word("12.01.2026", 40, 140),
    word("Продажа", 180, 140),
    word(`№${number}`, 280, 140),
    word(amount, 400, 140),
    word("12.01.2026", 820, 140),
    word("Приход", 980, 140),
    word(`№${number}`, 1100, 140),
    word(amount, 1440, 140),
  ];
}

function word(text: string, x: number, y: number, line?: number): OcrWord {
  return { text, confidence: 90, x0: x, y0: y, x1: x + text.length * 12, y1: y + 24, line };
}

function paintRows(width: number, height: number, count: number, gap: number): Uint8Array {
  const pixels = new Uint8Array(width * height).fill(255);
  const rowHeight = 4;
  for (let index = 0; index < count; index += 1) {
    const top = 8 + index * (rowHeight + gap);
    for (let y = top; y < top + rowHeight && y < height; y += 1) {
      for (let x = 2; x < width - 2; x += 1) pixels[y * width + x] = 0;
    }
  }
  return pixels;
}
