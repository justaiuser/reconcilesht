"use client";

import { useEffect, useMemo, useRef, useState } from "react";
import { LoaderCircle } from "lucide-react";
import { Button } from "@/components/ui/button";
import { ColumnMapper } from "@/components/column-mapper";
import { FileDrop } from "@/components/file-drop";
import { ResultView } from "@/components/result-view";
import { ThemeToggle } from "@/components/theme-toggle";
import { compareActs, compareDocuments } from "@/lib/compare";
import { blankMap, detectColumns, parseExcelMatrix } from "@/lib/excel-act";
import { recognizeRows } from "@/lib/row-client";
import { recognizeWithQwen } from "@/lib/qwen-client";
import { countPdfPages } from "@/lib/recognize";
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
  const [documents, setDocuments] = useState<ParsedAct[] | null>(null);
  const [singleFile, setSingleFile] = useState(false);
  const [qwenReady, setQwenReady] = useState(false);
  const [qwenHint, setQwenHint] = useState("");
  const [useQwen, setUseQwen] = useState(false);
  const [phase, setPhase] = useState<Phase>({ kind: "idle" });
  const [mode, setMode] = useState<CompareMode>("auto");
  const [mirror, setMirror] = useState<MirrorMode>("auto");
  const [sampleBusy, setSampleBusy] = useState(false);
  const cancelRef = useRef(false);
  const singleFileRef = useRef(false);
  const useQwenRef = useRef(false);

  useEffect(() => {
    let cancelled = false;
    fetch("/api/qwen")
      .then((response) => response.json())
      .then((body: { ok?: boolean; message?: string }) => {
        if (cancelled || body.ok !== true) return;
        setQwenReady(true);
        setQwenHint(typeof body.message === "string" ? body.message : "");
      })
      .catch(() => {
        if (!cancelled) setQwenReady(false);
      });
    return () => {
      cancelled = true;
    };
  }, []);

  function setSingle(value: boolean) {
    singleFileRef.current = value;
    setSingleFile(value);
  }

  function setQwen(value: boolean) {
    useQwenRef.current = value;
    setUseQwen(value);
  }

  const sheet = sheets.find((item) => item.name === sheetName) ?? null;
  const excelAct = useMemo(
    () => (sheet && map ? parseExcelMatrix(sheet.rows, map) : null),
    [sheet, map],
  );
  const result = useMemo(() => {
    if (singleFile) return documents ? compareDocuments(documents, { mirror }) : null;
    return excelAct && scanAct ? compareActs(excelAct, scanAct, { mode, mirror }) : null;
  }, [singleFile, documents, excelAct, scanAct, mode, mirror]);
  const working = phase.kind === "working";

  async function loadExcel(file: File) {
    setExcelError(null);
    setScanAct(null);
    setDocuments(null);
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
    setDocuments(null);
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
    setDocuments(null);
    setPhase({ kind: "working", label: "Готовим чтение скана…", percent: 2 });
    try {
      const onProgress = (progress: { label: string; percent: number }) => {
        if (!cancelRef.current) setPhase({ kind: "working", ...progress });
      };
      const shouldCancel = () => {
        if (cancelRef.current) throw new Error("cancelled");
      };
      const recognized = useQwenRef.current
        ? await recognizeWithQwen(pdfFile, onProgress, shouldCancel, { columns: singleFileRef.current })
        : await recognizeRows(pdfFile, onProgress, shouldCancel);
      if (cancelRef.current) return;
      setScanAct(recognized.act);
      setDocuments(recognized.documents);
      if (singleFileRef.current && !recognized.documents.some((act) => act.twoSided)) {
        setPhase({
          kind: "error",
          message:
            "В скане не нашлись две колонки. Снимите галочку «Прикрепить один файл» и приложите Excel.",
        });
        return;
      }
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
      setSingle(false);
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
    setDocuments(null);
    setSingle(false);
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
              Скан акта сверки и файл из учётной программы. Если в скане уже две
              колонки, Excel не нужен.{" "}
              {useQwen
                ? "Скан читает модель Qwen. Она может перепутать цифру или строку, поэтому для точной сверки галочку лучше снять."
                : "Каждая строка скана читается отдельно на компьютере, где запущена программа, и цифра не может переехать в соседнюю строку. В интернет файл не отправляется."}
            </p>
          </div>
          <ThemeToggle />
        </div>
      </header>

      <main className="mx-auto grid max-w-6xl gap-5 px-4 py-5 sm:px-6 sm:py-6 [&>*]:min-w-0">
        <label className="flex max-w-3xl items-start gap-3 text-sm leading-6">
          <input
            type="checkbox"
            className="mt-1 size-4 accent-primary"
            checked={singleFile}
            disabled={working || sampleBusy}
            onChange={(event) => {
              setSingle(event.target.checked);
              setPhase((current) => (current.kind === "error" ? { kind: "idle" } : current));
            }}
          />
          <span>
            <span className="font-medium">Прикрепить один файл</span>
            <span className="mt-0.5 block text-muted-foreground">
              В скане уже две колонки: слева одна компания, справа другая. Программа
              сравнит их между собой, Excel прикладывать не нужно.
            </span>
          </span>
        </label>

        {qwenReady ? (
          <label className="flex max-w-3xl items-start gap-3 text-sm leading-6">
            <input
              type="checkbox"
              className="mt-1 size-4 accent-primary"
              checked={useQwen}
              disabled={working || sampleBusy}
              onChange={(event) => setQwen(event.target.checked)}
            />
            <span>
              <span className="font-medium">Читать через Qwen</span>
              <span className="mt-0.5 block text-muted-foreground">
                {qwenHint ? `${qwenHint}. ` : ""}
                Модель переписывает таблицу своими словами и путает похожие суммы. Без
                галочки программа читает каждую строку отдельно и берёт цифры из неё.
              </span>
            </span>
          </label>
        ) : null}

        <div className={singleFile ? "grid gap-4" : "grid gap-4 lg:grid-cols-2"}>
          <FileDrop
            title="Скан акта"
            hint={
              singleFile
                ? "Один PDF, в котором слева и справа уже стоят обе компании."
                : "PDF, в котором страницы — изображения. Текст внутри файла не обязателен."
            }
            accept="application/pdf,.pdf"
            fileName={pdfName}
            meta={pdfPages == null ? null : pagesLabel(pdfPages)}
            disabled={working || sampleBusy}
            onFile={loadPdf}
          />
          {singleFile ? null : (
            <FileDrop
              title="Акт из программы"
              hint="Книга Excel .xlsx, как её выгружает учётная программа."
              accept=".xlsx,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
              fileName={excelName}
              disabled={working || sampleBusy}
              onFile={loadExcel}
            />
          )}
        </div>

        {pdfError ? <p className="text-sm text-destructive">{pdfError}</p> : null}
        {!singleFile && excelError ? <p className="text-sm text-destructive">{excelError}</p> : null}
        {pdfPages != null && pdfPages > 8 ? (
          <p className="text-sm leading-6 text-muted-foreground">
            {useQwen
              ? `В скане ${pagesLabel(pdfPages)}. Qwen читает страницы по одной, это займёт несколько минут.`
              : `В скане ${pagesLabel(pdfPages)}. Строки читаются по одной на компьютере, где запущена программа.`}
          </p>
        ) : null}

        {!singleFile && sheet && map && excelAct ? (
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
          <Button
            type="button"
            disabled={!pdfFile || working || sampleBusy || (!singleFile && !excelAct)}
            onClick={compare}
          >
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

        {result && scanAct && phase.kind !== "error" ? (
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
  if (/[а-яё]/i.test(message)) return message;
  return "Не удалось прочитать скан. Проверьте, что это обычный PDF, и попробуйте ещё раз.";
}
