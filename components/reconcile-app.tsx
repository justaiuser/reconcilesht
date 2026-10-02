"use client";

import { useMemo, useRef, useState } from "react";
import { LoaderCircle } from "lucide-react";
import { Button } from "@/components/ui/button";
import { ColumnMapper } from "@/components/column-mapper";
import { FileDrop } from "@/components/file-drop";
import { ResultView } from "@/components/result-view";
import { ThemeToggle } from "@/components/theme-toggle";
import { compareActs } from "@/lib/compare";
import { blankMap, detectColumns, parseExcelMatrix } from "@/lib/excel-act";
import { countPdfPages, recognizePdf } from "@/lib/recognize";
import { buildSampleFiles } from "@/lib/sample-files";
import type { Cell, ColumnMap, CompareMode, MirrorMode, ParsedAct } from "@/lib/types";

type SheetModel = { name: string; rows: Cell[][] };

type Phase =
  | { kind: "idle" }
  | { kind: "working"; label: string; percent: number }
  | { kind: "error"; message: string }
  | { kind: "done" };

export function ReconcileApp() {
  const [excelName, setExcelName] = useState<string | null>(null);
  const [sheets, setSheets] = useState<SheetModel[]>([]);
  const [sheetName, setSheetName] = useState("");
  const [map, setMap] = useState<ColumnMap | null>(null);
  const [excelError, setExcelError] = useState<string | null>(null);

  const [pdfName, setPdfName] = useState<string | null>(null);
  const [pdfFile, setPdfFile] = useState<File | null>(null);
  const [pdfPages, setPdfPages] = useState<number | null>(null);
  const [pdfError, setPdfError] = useState<string | null>(null);

  const [scanAct, setScanAct] = useState<ParsedAct | null>(null);
  const [phase, setPhase] = useState<Phase>({ kind: "idle" });
  const [mode, setMode] = useState<CompareMode>("auto");
  const [mirror, setMirror] = useState<MirrorMode>("auto");
  const [sampleBusy, setSampleBusy] = useState(false);
  const cancelRef = useRef(false);

  const sheet = sheets.find((item) => item.name === sheetName) ?? null;
  const excelAct = useMemo(
    () => (sheet && map ? parseExcelMatrix(sheet.rows, map) : null),
    [sheet, map],
  );
  const result = useMemo(
    () => (excelAct && scanAct ? compareActs(excelAct, scanAct, { mode, mirror }) : null),
    [excelAct, scanAct, mode, mirror],
  );
  const working = phase.kind === "working";

  async function loadExcel(file: File) {
    setExcelError(null);
    setScanAct(null);
    setPhase({ kind: "idle" });
    if (!file.name.toLowerCase().endsWith(".xlsx")) {
      setExcelName(null);
      setSheets([]);
      setMap(null);
      setExcelError("Нужен файл .xlsx. Если акт в старом .xls, откройте его и сохраните как книгу Excel.");
      return;
    }
    try {
      const { default: readXlsxFile } = await import("read-excel-file/browser");
      const loaded = await readXlsxFile(file);
      const nextSheets = loaded.map((item) => ({
        name: item.sheet,
        rows: item.data as Cell[][],
      }));
      if (nextSheets.length === 0) throw new Error("empty");
      const preferred =
        nextSheets.find((item) => /сверк|акт/i.test(item.name)) ?? nextSheets[0];
      setSheets(nextSheets);
      setSheetName(preferred.name);
      setMap(detectColumns(preferred.rows) ?? blankMap(columnCount(preferred.rows)));
      setExcelName(file.name);
    } catch {
      setExcelName(null);
      setSheets([]);
      setMap(null);
      setExcelError("Не удалось прочитать Excel. Сохраните копию в .xlsx без пароля.");
    }
  }

  function chooseSheet(name: string) {
    const next = sheets.find((item) => item.name === name);
    if (!next) return;
    setSheetName(name);
    setMap(detectColumns(next.rows) ?? blankMap(columnCount(next.rows)));
  }

  async function loadPdf(file: File) {
    setPdfError(null);
    setScanAct(null);
    setPhase({ kind: "idle" });
    if (file.type && file.type !== "application/pdf" && !file.name.toLowerCase().endsWith(".pdf")) {
      setPdfFile(null);
      setPdfName(null);
      setPdfPages(null);
      setPdfError("Нужен скан в PDF.");
      return;
    }
    if (file.size > 80 * 1024 * 1024) {
      setPdfFile(null);
      setPdfName(null);
      setPdfPages(null);
      setPdfError("Файл больше 80 МБ. Оставьте в нём только страницы акта.");
      return;
    }
    try {
      const pages = await countPdfPages(file);
      setPdfFile(file);
      setPdfName(file.name);
      setPdfPages(pages);
    } catch (error) {
      setPdfFile(null);
      setPdfName(null);
      setPdfPages(null);
      setPdfError(pdfMessage(error));
    }
  }

  async function compare() {
    if (!pdfFile || working) return;
    cancelRef.current = false;
    setScanAct(null);
    setPhase({ kind: "working", label: "Готовим чтение скана…", percent: 2 });
    try {
      const act = await recognizePdf(
        pdfFile,
        (progress) => {
          if (!cancelRef.current) setPhase({ kind: "working", ...progress });
        },
        () => {
          if (cancelRef.current) throw new Error("cancelled");
        },
      );
      if (cancelRef.current) return;
      setScanAct(act);
      setPhase({ kind: "done" });
    } catch (error) {
      if (cancelRef.current) {
        setPhase({ kind: "idle" });
        return;
      }
      setPhase({ kind: "error", message: pdfMessage(error) });
    }
  }

  async function loadSample() {
    setSampleBusy(true);
    setExcelError(null);
    setPdfError(null);
    try {
      const sample = await buildSampleFiles();
      await loadExcel(sample.excel);
      await loadPdf(sample.pdf);
    } catch {
      setExcelError("Не удалось собрать пример. Обновите страницу и попробуйте снова.");
    } finally {
      setSampleBusy(false);
    }
  }

  function reset() {
    cancelRef.current = true;
    setExcelName(null);
    setSheets([]);
    setSheetName("");
    setMap(null);
    setExcelError(null);
    setPdfName(null);
    setPdfFile(null);
    setPdfPages(null);
    setPdfError(null);
    setScanAct(null);
    setPhase({ kind: "idle" });
    setMode("auto");
    setMirror("auto");
  }

  return (
    <div className="min-h-full bg-background">
      <header className="border-b border-border bg-card">
        <div className="mx-auto flex max-w-6xl items-start justify-between gap-4 px-4 py-4 sm:px-6">
          <div>
            <h1 className="text-xl font-semibold tracking-tight">Сверка актов</h1>
            <p className="mt-1 max-w-2xl text-sm leading-6 text-muted-foreground">
              Скан акта сверки и файл из учётной программы. Документы читаются в этом
              браузере и никуда не отправляются.
            </p>
          </div>
          <ThemeToggle />
        </div>
      </header>

      <main className="mx-auto grid max-w-6xl gap-5 px-4 py-5 sm:px-6 sm:py-6 [&>*]:min-w-0">
        <div className="grid gap-4 lg:grid-cols-2">
          <FileDrop
            title="Скан акта"
            hint="PDF, в котором страницы — изображения. Текст внутри файла не обязателен."
            accept="application/pdf,.pdf"
            fileName={pdfName}
            meta={pdfPages == null ? null : pagesLabel(pdfPages)}
            disabled={working || sampleBusy}
            onFile={loadPdf}
          />
          <FileDrop
            title="Акт из программы"
            hint="Книга Excel .xlsx, как её выгружает учётная программа."
            accept=".xlsx,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            fileName={excelName}
            disabled={working || sampleBusy}
            onFile={loadExcel}
          />
        </div>

        {pdfError ? <p className="text-sm text-destructive">{pdfError}</p> : null}
        {excelError ? <p className="text-sm text-destructive">{excelError}</p> : null}
        {pdfPages != null && pdfPages > 8 ? (
          <p className="text-sm leading-6 text-muted-foreground">
            В скане {pagesLabel(pdfPages)}. Чтение займёт несколько минут, и компьютер
            может притормаживать — так и должно быть.
          </p>
        ) : null}

        {sheet && map && excelAct ? (
          <ColumnMapper
            rows={sheet.rows}
            map={map}
            act={excelAct}
            sheetNames={sheets.map((item) => item.name)}
            sheetName={sheetName}
            onSheet={chooseSheet}
            onMap={setMap}
          />
        ) : null}

        <div className="flex flex-wrap items-center gap-2">
          <Button type="button" disabled={!pdfFile || !excelAct || working || sampleBusy} onClick={compare}>
            {working ? <LoaderCircle className="animate-spin" /> : null}
            Сравнить
          </Button>
          {working ? (
            <Button
              type="button"
              variant="outline"
              onClick={() => {
                cancelRef.current = true;
              }}
            >
              Остановить
            </Button>
          ) : null}
          <Button type="button" variant="outline" disabled={working || sampleBusy} onClick={loadSample}>
            {sampleBusy ? <LoaderCircle className="animate-spin" /> : null}
            Открыть пример
          </Button>
          {excelName || pdfName ? (
            <Button type="button" variant="ghost" disabled={working} onClick={reset}>
              Сбросить
            </Button>
          ) : null}
        </div>

        {phase.kind === "working" ? (
          <div
            role="status"
            aria-live="polite"
            className="rounded-2xl border border-border bg-card p-4"
          >
            <div className="h-1.5 overflow-hidden rounded-full bg-muted">
              <div
                className="h-full bg-primary transition-[width]"
                style={{ width: `${phase.percent}%` }}
              />
            </div>
            <p className="mt-3 text-sm">{phase.label}</p>
          </div>
        ) : null}
        {phase.kind === "error" ? (
          <p role="alert" className="text-sm leading-6 text-destructive">
            {phase.message}
          </p>
        ) : null}

        {!excelName && !pdfName && phase.kind === "idle" ? (
          <p className="max-w-3xl text-sm leading-6 text-muted-foreground">
            В примере — акт ООО «Север» и ООО «Луч» за 1 квартал 2026. Наша колонка
            сходится, а у контрагента на скане другие суммы по №27 и №88, нет
            реализации №31 и есть корректировка №5.
          </p>
        ) : null}

        {result && scanAct ? (
          <ResultView
            result={result}
            mode={mode}
            mirror={mirror}
            rawText={scanAct.rawText}
            confidence={scanAct.confidence}
            warnings={scanAct.warnings}
            onMode={setMode}
            onMirror={setMirror}
          />
        ) : null}
      </main>
    </div>
  );
}

function columnCount(rows: Cell[][]): number {
  return rows.reduce((max, row) => Math.max(max, row.length), 0);
}

function pagesLabel(count: number): string {
  const mod10 = count % 10;
  const mod100 = count % 100;
  const word =
    mod10 === 1 && mod100 !== 11
      ? "страница"
      : mod10 >= 2 && mod10 <= 4 && (mod100 < 12 || mod100 > 14)
        ? "страницы"
        : "страниц";
  return `${count} ${word}`;
}

function pdfMessage(error: unknown): string {
  const message = error instanceof Error ? error.message : "";
  if (/password/i.test(message)) return "PDF закрыт паролем. Сохраните копию без пароля.";
  if (/cancelled/i.test(message)) return "Чтение остановлено.";
  if (/страниц/i.test(message)) return message;
  return "Не удалось прочитать скан. Проверьте, что это обычный PDF, и попробуйте ещё раз.";
}
