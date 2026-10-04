import { emptySummary } from "@/lib/excel-act";
import { stretchContrast } from "@/lib/image";
import { loadPdfJs } from "@/lib/pdfjs";
import { columnSplit, pageStrips } from "@/lib/page-strips";
import { asRightTable, joinTables, qwenPagesToRecognition, type QwenPage } from "@/lib/qwen-act";
import type { ParsedAct } from "@/lib/types";
import type { RecognizeProgress } from "@/lib/recognize";

const MAX_PAGES = 40;
const PAGE_SIDE = 2000;
const DETAIL_SIDE = 3000;
// About a quarter of a dense act page: the model writes it out in a couple of minutes.
const ROWS_PER_PIECE = 22;

export async function recognizeWithQwen(
  file: File,
  onProgress: (progress: RecognizeProgress) => void,
  shouldCancel: () => void,
  options: { columns?: boolean } = {},
): Promise<{ act: ParsedAct; documents: ParsedAct[] }> {
  onProgress({ label: "Открываем скан для Qwen…", percent: 4 });
  const pdfjs = await loadPdfJs();
  const data = new Uint8Array(await file.arrayBuffer());
  const loadingTask = pdfjs.getDocument({ data });
  try {
    const pdf = await loadingTask.promise;
    if (pdf.numPages > MAX_PAGES) {
      throw new Error(
        `В скане ${pdf.numPages} страниц. Оставьте в файле только акт, не больше ${MAX_PAGES} страниц.`,
      );
    }
    const pages: QwenPage[] = [];
    for (let pageNumber = 1; pageNumber <= pdf.numPages; pageNumber += 1) {
      shouldCancel();
      onProgress({
        label: `Qwen читает страницу ${pageNumber} из ${pdf.numPages}`,
        percent: 8 + ((pageNumber - 1) / pdf.numPages) * 86,
      });
      const page = await pdf.getPage(pageNumber);
      const unscaled = page.getViewport({ scale: 1 });
      const longest = Math.max(unscaled.width, unscaled.height);
      const viewport = page.getViewport({ scale: Math.min(3, DETAIL_SIDE / longest) });
      const canvas = window.document.createElement("canvas");
      canvas.width = Math.ceil(viewport.width);
      canvas.height = Math.ceil(viewport.height);
      await page.render({ canvas, viewport }).promise;
      page.cleanup();
      stretchContrast(canvas);
      const context = canvas.getContext("2d", { willReadFrequently: true });
      const pixels = context?.getImageData(0, 0, canvas.width, canvas.height);
      const pieces = pixels
        ? pageStrips(pixels.data, canvas.width, canvas.height, 4, ROWS_PER_PIECE)
        : [{ top: 0, bottom: canvas.height }];

      const split =
        options.columns && pixels ? columnSplit(pixels.data, canvas.width, canvas.height) : null;
      const parts = split == null ? 1 : 2;
      const steps = pieces.length * parts;

      for (let pieceIndex = 0; pieceIndex < pieces.length; pieceIndex += 1) {
        const piece = pieces[pieceIndex];
        const read = async (left: number, right: number, part: number) => {
          shouldCancel();
          const step = pieceIndex * parts + part;
          onProgress({
            label:
              steps > 1
                ? `Qwen читает страницу ${pageNumber} из ${pdf.numPages}, часть ${step + 1} из ${steps}`
                : `Qwen читает страницу ${pageNumber} из ${pdf.numPages}`,
            percent: 8 + ((pageNumber - 1 + step / steps) / pdf.numPages) * 86,
          });
          const crop = cropCanvas(canvas, piece.top, piece.bottom, left, right);
          return readPiece(await canvasToBase64(scaleCanvas(crop, PAGE_SIDE)));
        };
        try {
          const result =
            split == null
              ? await read(0, canvas.width, 0)
              : joinTables(await read(0, split, 0), asRightTable(await read(split, canvas.width, 1)));
          const newAct = pieceIndex > 0 ? false : pageNumber > 1 ? result.newAct : true;
          pages.push({ ...result, newAct });
        } catch (error) {
          const message = error instanceof Error ? error.message : "Модель не прочитала часть страницы.";
          pages.push(blankPage(pages.length === 0, message));
        }
      }
    }
    if (pages.every((page) => page.rows.length === 0)) {
      throw new Error(pages.at(-1)?.note || "Модель не прочитала страницу.");
    }
    onProgress({ label: "Сопоставляем строки…", percent: 96 });
    return qwenPagesToRecognition(pages);
  } finally {
    await loadingTask.destroy();
  }
}

function blankPage(newAct: boolean, note: string): QwenPage {
  return {
    newAct,
    partyLeft: null,
    partyRight: null,
    period: null,
    rows: [],
    left: emptySummary(),
    right: emptySummary(),
    note,
  };
}

async function readPiece(image: string): Promise<QwenPage> {
  const response = await fetch("/api/qwen", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ image }),
  });
  const body = (await response.json().catch(() => null)) as { page?: QwenPage; message?: string } | null;
  if (!response.ok || !body?.page) {
    throw new Error(body?.message || "Модель Qwen не смогла прочитать страницу.");
  }
  return body.page;
}

function scaleCanvas(source: HTMLCanvasElement, side: number): HTMLCanvasElement {
  const ratio = Math.min(1, side / Math.max(source.width, source.height));
  if (ratio === 1) return source;
  const canvas = window.document.createElement("canvas");
  canvas.width = Math.round(source.width * ratio);
  canvas.height = Math.round(source.height * ratio);
  canvas.getContext("2d")?.drawImage(source, 0, 0, canvas.width, canvas.height);
  return canvas;
}

function cropCanvas(
  source: HTMLCanvasElement,
  top: number,
  bottom: number,
  left: number,
  right: number,
): HTMLCanvasElement {
  if (top === 0 && left === 0 && bottom === source.height && right === source.width) return source;
  const canvas = window.document.createElement("canvas");
  canvas.width = Math.max(1, right - left);
  canvas.height = Math.max(1, bottom - top);
  canvas
    .getContext("2d")
    ?.drawImage(source, left, top, canvas.width, canvas.height, 0, 0, canvas.width, canvas.height);
  return canvas;
}

function canvasToBase64(canvas: HTMLCanvasElement): Promise<string> {
  return new Promise((resolve, reject) => {
    canvas.toBlob(
      (blob) => {
        if (!blob) {
          reject(new Error("Не удалось подготовить страницу для модели."));
          return;
        }
        const reader = new FileReader();
        reader.onload = () => {
          const value = String(reader.result ?? "");
          resolve(value.replace(/^data:image\/jpeg;base64,/, ""));
        };
        reader.onerror = () => reject(new Error("Не удалось подготовить страницу для модели."));
        reader.readAsDataURL(blob);
      },
      "image/jpeg",
      0.9,
    );
  });
}
