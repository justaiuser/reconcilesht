import { describe, it } from "node:test";
import assert from "node:assert/strict";
import { parseMoney, parseDate, extractDocNumber, docFromDocumentCell, lineLabel, normalizeDoc } from "./text";
import { detectColumns, emptySummary, parseExcelMatrix } from "./excel-act";
import { groupLines, parseOcrDocuments, parseOcrPages, parseOcrWords } from "./ocr-table";
import { pageStrips } from "./page-strips";
import { compareActs, compareDocuments } from "./compare";
import { QWEN_PROMPT, looksLooping, parseQwenResponse, qwenPagesToRecognition, type QwenPage } from "./qwen-act";
import { readPrintedPage } from "./printed-row";
import { chooseQwenModel, readStream } from "./qwen-ollama";
import { buildExcelMatrix, buildSampleScanAct } from "./sample-data";
import type { ActLine, ParsedAct, OcrWord } from "./types";

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

describe("ответ Qwen", () => {
  it("собирает двусторонний акт из JSON, даже если вокруг него есть текст", () => {
    const page = parseQwenResponse(
      "Таблица готова:\n```json\n" + JSON.stringify(qwenPayload("1", 471487.25)) + "\n```",
    );
    assert.equal(page.newAct, true);
    assert.equal(page.rows[0]?.leftDebit, 471487.25);
    assert.equal(page.rows[0]?.rightCredit, 471487.25);
    assert.equal(page.left.openingDebit, 6817644.88);

    const recognized = qwenPagesToRecognition([page]);
    assert.equal(recognized.documents.length, 1);
    assert.equal(recognized.act.twoSided, true);
    assert.equal(recognized.act.partyA, "ООО ТИМПРО");
    assert.equal(recognized.act.lines.find((line) => line.side === "ours")?.debit, 471487.25);
    const compared = compareDocuments(recognized.documents);
    assert.equal(compared?.reports[0]?.pairs[0]?.status, "match");
  });

  it("дописывает продолжение таблицы и отделяет новый акт", () => {
    const first = parseQwenResponse(JSON.stringify(qwenPayload("1", 1000)));
    const continuation = parseQwenResponse(JSON.stringify({ ...qwenPayload("2", 2500), newAct: false }));
    const joined = qwenPagesToRecognition([first, continuation]);
    assert.equal(joined.documents.length, 1);
    assert.equal(joined.act.lines.filter((line) => line.side === "ours").length, 2);

    const another = parseQwenResponse(JSON.stringify({ ...qwenPayload("9", 3000), newAct: true, partyLeft: "ООО Велес" }));
    const split = qwenPagesToRecognition([first, another]);
    assert.equal(split.documents.length, 2);
    assert.equal(split.documents[1]?.partyA, "ООО Велес");
    assert.equal(split.documents[0]?.lines.some((line) => line.docNumber === "9"), false);
  });

  it("принимает сумму строкой и имя lines", () => {
    const page = parseQwenResponse(
      JSON.stringify({
        newAct: true,
        lines: [{ date: "12.01.2026", document: "Продажа (7 от 12.01.2026)", leftDebit: "471 487,25", rightCredit: "471487.25" }],
      }),
    );
    assert.equal(page.rows[0]?.leftDebit, 471487.25);
    assert.equal(page.rows[0]?.rightCredit, 471487.25);
    const recognized = qwenPagesToRecognition([page]);
    assert.equal(recognized.act.lines.find((line) => line.side === "ours")?.docNumber, "7");
  });
});

describe("одна колонка Excel и две колонки ответа Qwen", () => {
  it("сверяет номера счетов один раз и не берет сальдо в документы", () => {
    const page = parseQwenResponse(
      JSON.stringify({
        newAct: true,
        rows: [
          echoed("24.09.2025", "1800065755-254", 1330442.88),
          echoed("26.09.2025", "1800066630-254", 3615754.13),
          { date: "01.10.2025", document: "Сальдо на 01.10.2025", leftDebit: 471487.25, rightCredit: 471487.25 },
        ],
      }),
    );
    const recognized = qwenPagesToRecognition([page]);
    assert.equal(recognized.act.lines.some((line) => /сальдо/i.test(line.description)), false);
    assert.equal(recognized.act.summary.closingDebit, 471487.25);
    assert.equal(recognized.act.theirsSummary?.closingCredit, 471487.25);
    assert.match(recognized.act.rawText, /№1800065755-254/);

    const excel = bareAct([
      invoice("e1", "single", "2025-09-24", "1800065755-254", null, 1330442.88),
      invoice("e2", "single", "2025-09-26", "1800066630-254", null, 3617522.5),
    ]);
    const result = compareActs(excel, recognized.act);
    const pairs = result.reports[0]?.pairs ?? [];
    assert.equal(pairs.filter((pair) => pair.status === "scan-only").length, 0);
    assert.equal(pairs.find((pair) => pair.excel?.docNumber === "1800065755-254")?.status, "match");
    assert.equal(pairs.find((pair) => pair.excel?.docNumber === "1800066630-254")?.status, "amount");
    assert.equal(pairs.length, 2);
    assert.equal(lineLabel("Исходящие счета", "1800065755-254"), "Исходящие счета № 1800065755-254");
    assert.equal(lineLabel("1800065755-254", "1800065755-254"), "1800065755-254");
  });
});

describe("ошибки маленькой модели", () => {
  it("дочитывает зациклившийся ответ и оставляет одну строку из повторов", () => {
    const row = '{"date":"01.10.2025","document":"Акт сверки","number":"26 306 875,68","leftDebit":26306875.68,"leftCredit":0,"rightDebit":0,"rightCredit":0}';
    const invoiceRow = '{"date":"24.09.2025","document":"Исходящие счета","number":"1800065755-254","leftDebit":1330442.88,"leftCredit":null,"rightDebit":null,"rightCredit":null}';
    const looped = `{"newAct":true,"left":{"closingDebit":26306875.68},"right":{},"rows":[${invoiceRow},${Array(6).fill(row).join(",")},{"date":"01.1`;
    assert.equal(looksLooping(looped.repeat(2)), true);
    const page = parseQwenResponse(looped);
    assert.equal(page.rows.length, 1);
    assert.equal(page.rows[0]?.number, "1800065755-254");
    assert.equal(page.left.closingDebit, 26306875.68);

    const repeated = parseQwenResponse(
      JSON.stringify({ newAct: true, rows: Array(4).fill(JSON.parse(invoiceRow)) }),
    );
    assert.equal(repeated.rows.length, 1);
    const twice = parseQwenResponse(JSON.stringify({ newAct: true, rows: Array(2).fill(JSON.parse(invoiceRow)) }));
    assert.equal(twice.rows.length, 2);
  });

  it("берёт входящее сальдо с первой страницы, а итоги — с последней", () => {
    const first = parseQwenResponse(
      JSON.stringify({
        newAct: true,
        left: { openingDebit: 0, openingCredit: 0, turnoverDebit: 0, closingDebit: 0 },
        rows: [JSON.parse('{"date":"24.09.2025","document":"Исходящие счета","number":"1800065755-254","leftDebit":1330442.88}')],
      }),
    );
    const totals = parseQwenResponse(
      JSON.stringify({
        newAct: true,
        left: { openingDebit: 26306875.68, turnoverDebit: 26306875.68, closingDebit: 26306875.68 },
        right: { openingDebit: 0, closingDebit: 0 },
        rows: [],
      }),
    );
    const { act, documents } = qwenPagesToRecognition([first, totals]);
    assert.equal(documents.length, 1);
    assert.equal(act.summary.openingDebit, 0);
    assert.equal(act.summary.turnoverDebit, 26306875.68);
    assert.equal(act.summary.closingDebit, 26306875.68);
    assert.equal(act.theirsSummary, null);
    assert.equal(act.twoSided, false);
  });
});

describe("строка печатного акта", () => {
  it("берёт номер с дефисом и сумму из своей строки, не путая даты с деньгами", () => {
    const page = readPrintedPage([
      "Входящее сальдо: по дебету по кредиту",
      "0,00 0,00 руб",
      "1800065755 — 1800065755-254 24.09.2025 24.09.2025 1 330 442,88 руб",
      "1800066630 — 1800066630-254 26.09.2025 26.09.2025 3 617 522,50 руб",
      "1800066631 1800066631-254 26.09.2025 26.09.2025 3 616 777,92 руб",
      "какая-то рукописная помета 24.09.2025",
    ]);
    assert.equal(page.summary.openingDebit, 0);
    assert.equal(page.summary.openingCredit, 0);
    assert.deepEqual(
      page.lines.map((line) => [line.number, line.date, line.debit]),
      [
        ["1800065755-254", "2025-09-24", 1330442.88],
        ["1800066630-254", "2025-09-26", 3617522.5],
        ["1800066631-254", "2025-09-26", 3616777.92],
      ],
    );
  });

  it("относит итог к оборотам, а следующий такой же итог к исходящему сальдо", () => {
    const page = readPrintedPage([
      "Обороты по дебету: Обороты по кредиту:",
      "26 306 875,68 0.00 руб",
      "Исходящее сальдо: по дебету по кредиту",
      "26 306 875,68 0,00 руб",
    ]);
    assert.equal(page.summary.turnoverDebit, 26306875.68);
    assert.equal(page.summary.turnoverCredit, 0);
    assert.equal(page.summary.closingDebit, 26306875.68);
    assert.equal(page.summary.closingCredit, 0);
    assert.equal(page.lines.length, 0);
  });

  it("не записывает продажу с датой во входящее сальдо, даже если номер не прочитался", () => {
    const page = readPrintedPage([
      "Сальдо начальное",
      "12.01.2026 Продажа (1 от 14.01.2026) 471 487,25",
    ]);
    assert.equal(page.summary.openingDebit, null);
    assert.equal(page.lines[0]?.number, "1");
    assert.equal(page.lines[0]?.debit, 471487.25);
  });
});

describe("выбор модели Qwen", () => {
  it("предпочитает instruct и иначе берёт любую qwen3-vl", () => {
    assert.equal(
      chooseQwenModel(["qwen3-vl:2b", "qwen3-vl:2b-instruct"], "qwen3-vl:2b-instruct"),
      "qwen3-vl:2b-instruct",
    );
    assert.equal(
      chooseQwenModel(["qwen3-vl:2b-instruct:latest"], "qwen3-vl:2b-instruct"),
      "qwen3-vl:2b-instruct:latest",
    );
    assert.equal(chooseQwenModel(["llama3", "qwen3-vl:2b"], "qwen3-vl:2b-instruct"), "qwen3-vl:2b");
    assert.equal(chooseQwenModel(["llama3"], "qwen3-vl:2b-instruct"), null);
    assert.equal(
      chooseQwenModel(["qwen3-vl:2b-instruct", "qwen3-vl:8b", "qwen3-vl:8b-instruct", "qwen3-vl:4b-instruct"], undefined),
      "qwen3-vl:8b-instruct",
    );
  });

  it("оставляет уже полученный текст, если модель оборвалась", async () => {
    const encoder = new TextEncoder();
    let sent = false;
    const stream = new ReadableStream<Uint8Array>({
      pull(controller) {
        if (sent) {
          controller.error(new Error("abort"));
          return;
        }
        sent = true;
        controller.enqueue(encoder.encode(`${JSON.stringify({ message: { content: "{\"rows\":[" } })}\n`));
      },
    });
    assert.match(await readStream(stream), /\{"rows":\[/);
  });

  it("не даёт модели готовых номеров и сумм, которые можно переписать", () => {
    assert.equal(/\d{4,}/.test(QWEN_PROMPT), false);
  });
});

function echoed(date: string, number: string, amount: number) {
  return {
    date,
    document: "Исходящие счета",
    number,
    leftDebit: amount,
    leftCredit: null,
    rightDebit: null,
    rightCredit: amount,
  };
}

function invoice(
  id: string,
  side: ActLine["side"],
  date: string,
  docNumber: string,
  debit: number | null,
  credit: number | null,
): ActLine {
  return { id, side, date, docNumber, description: docNumber, debit, credit };
}

function bareAct(lines: ActLine[]): ParsedAct {
  return {
    title: null,
    period: null,
    partyA: null,
    partyB: null,
    lines,
    summary: emptySummary(),
    theirsSummary: null,
    warnings: [],
    confidence: null,
    rawText: "",
    twoSided: false,
  };
}

function qwenPayload(number: string, amount: number): QwenPage {
  return {
    newAct: true,
    partyLeft: "ООО ТИМПРО",
    partyRight: "АО ТЭМПО",
    period: "1 квартал 2026",
    rows: [
      {
        date: "12.01.2026",
        document: `Продажа (${number} от 12.01.2026)`,
        number,
        leftDebit: amount,
        leftCredit: null,
        rightDebit: null,
        rightCredit: amount,
      },
    ],
    left: {
      openingDebit: 6817644.88,
      openingCredit: null,
      turnoverDebit: amount,
      turnoverCredit: null,
      closingDebit: null,
      closingCredit: null,
    },
    right: {
      openingDebit: null,
      openingCredit: 6817644.88,
      turnoverDebit: null,
      turnoverCredit: amount,
      closingDebit: null,
      closingCredit: null,
    },
    note: null,
  };
}

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
