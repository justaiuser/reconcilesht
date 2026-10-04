import type { OcrWord, ParsedAct } from "@/lib/types";
import { stretchContrast } from "@/lib/image";
import { parseOcrDocuments, parseOcrPages } from "@/lib/ocr-table";
import { pageStrips } from "@/lib/page-strips";

export type RecognizeProgress = {
  label: string;
  percent: number;
};

type OcrBlock = {
  paragraphs?: {
    lines?: {
      words?: {
        text: string;
        confidence: number;
        bbox: { x0: number; y0: number; x1: number; y1: number };
      }[];
    }[];
  }[];
};

const MAX_PAGES = 40;

export type Recognition = {
  act: ParsedAct;
  documents: ParsedAct[];
};

export async function recognizePdf(
  file: File,
  onProgress: (progress: RecognizeProgress) => void,
  shouldCancel: () => void,
): Promise<Recognition> {
  onProgress({ label: "Открываем скан…", percent: 4 });
  const pdfjs = await import("pdfjs-dist");
  pdfjs.GlobalWorkerOptions.workerSrc = "/pdf.worker.min.mjs";
  const data = new Uint8Array(await file.arrayBuffer());
  const loadingTask = pdfjs.getDocument({ data });
  try {
  const pdf = await loadingTask.promise;
  if (pdf.numPages > MAX_PAGES) {
    throw new Error(
      `В скане ${pdf.numPages} страниц. Оставьте в файле только акт, не больше ${MAX_PAGES} страниц.`,
    );
  }

  onProgress({ label: "Загружаем распознавание текста…", percent: 8 });
  const tesseract = await import("tesseract.js");
  const worker = await tesseract.createWorker("rus", 1, {
    workerPath: "/tesseract/worker.min.js",
    corePath: "/tesseract/tesseract-core-lstm.wasm.js",
    langPath: "/tessdata",
    workerBlobURL: false,
  });

  try {
    await worker.setParameters({
      tessedit_pageseg_mode: tesseract.PSM.AUTO,
      preserve_interword_spaces: "1",
      user_defined_dpi: "300",
    });

    const pages: OcrWord[][] = [];
    const texts: string[] = [];
    for (let pageNumber = 1; pageNumber <= pdf.numPages; pageNumber += 1) {
      shouldCancel();
      onProgress({
        label: `Читаем страницу ${pageNumber} из ${pdf.numPages}`,
        percent: 12 + ((pageNumber - 1) / pdf.numPages) * 80,
      });
      const page = await pdf.getPage(pageNumber);
      const unscaled = page.getViewport({ scale: 1 });
      const longest = Math.max(unscaled.width, unscaled.height);
      const viewport = page.getViewport({ scale: Math.min(3, Math.max(2, 4200 / longest)) });
      const canvas = window.document.createElement("canvas");
      canvas.width = Math.ceil(viewport.width);
      canvas.height = Math.ceil(viewport.height);
      await page.render({ canvas, viewport }).promise;
      stretchContrast(canvas);
      const context = canvas.getContext("2d", { willReadFrequently: true });
      const image = context?.getImageData(0, 0, canvas.width, canvas.height);
      const strips = image
        ? pageStrips(image.data, canvas.width, canvas.height)
        : [{ top: 0, bottom: canvas.height }];
      const lineIds = { next: 1 };
      const pageWords: OcrWord[] = [];
      const pageText: string[] = [];
      for (let stripIndex = 0; stripIndex < strips.length; stripIndex += 1) {
        shouldCancel();
        const strip = strips[stripIndex];
        const pad = strips.length === 1 ? 0 : 18;
        const cropTop = Math.max(0, strip.top - pad);
        const cropBottom = Math.min(canvas.height, strip.bottom + pad);
        onProgress({
          label:
            strips.length === 1
              ? `Читаем страницу ${pageNumber} из ${pdf.numPages}`
              : `Читаем страницу ${pageNumber} из ${pdf.numPages}, полоса ${stripIndex + 1} из ${strips.length}`,
          percent:
            12 +
            ((pageNumber - 1 + (stripIndex + 1) / strips.length) / pdf.numPages) * 80,
        });
        const target = strips.length === 1 ? canvas : cropCanvas(canvas, cropTop, cropBottom);
        const result = await worker.recognize(target, {}, { text: true, blocks: true });
        pageText.push(result.data.text ?? "");
        const stripWords = wordsFromResult(
          result.data.blocks as OcrBlock[] | null,
          result.data.text ?? "",
          cropTop,
          lineIds,
        ).filter((word) => {
          const center = (word.y0 + word.y1) / 2;
          return center >= strip.top && center < strip.bottom;
        });
        pageWords.push(...stripWords);
      }
      texts.push(pageText.join("\n"));
      pages.push(pageWords);
      page.cleanup();
    }

    onProgress({ label: "Сопоставляем строки…", percent: 96 });
    const rawText = texts.join("\n");
    return {
      act: parseOcrPages(pages, rawText),
      documents: parseOcrDocuments(pages, texts),
    };
  } finally {
    await worker.terminate();
  }
  } finally {
    await loadingTask.destroy();
  }
}

function cropCanvas(source: HTMLCanvasElement, top: number, bottom: number): HTMLCanvasElement {
  const canvas = document.createElement("canvas");
  canvas.width = source.width;
  canvas.height = Math.max(1, bottom - top);
  const context = canvas.getContext("2d");
  context?.drawImage(source, 0, top, source.width, canvas.height, 0, 0, source.width, canvas.height);
  return canvas;
}

function wordsFromResult(
  blocks: OcrBlock[] | null,
  text: string,
  yOffset: number,
  lineIds: { next: number },
): OcrWord[] {
  const words: OcrWord[] = [];
  for (const block of blocks ?? []) {
    for (const paragraph of block.paragraphs ?? []) {
      for (const line of paragraph.lines ?? []) {
        const lineId = lineIds.next;
        let added = false;
        for (const word of line.words ?? []) {
          if (!word.text.trim()) continue;
          added = true;
          words.push({
            text: word.text,
            confidence: word.confidence,
            x0: word.bbox.x0,
            y0: word.bbox.y0 + yOffset,
            x1: word.bbox.x1,
            y1: word.bbox.y1 + yOffset,
            line: lineId,
          });
        }
        if (added) lineIds.next += 1;
      }
    }
  }
  if (words.length > 0) return words;
  return wordsFromPlainText(text, yOffset, lineIds);
}

function wordsFromPlainText(text: string, yOffset: number, lineIds: { next: number }): OcrWord[] {
  const words: OcrWord[] = [];
  text.split(/\n/).forEach((line) => {
    const tokens = line.split(/\s+/).filter(Boolean);
    if (tokens.length === 0) return;
    const lineId = lineIds.next;
    lineIds.next += 1;
    let x = 0;
    const y = yOffset + (lineId - 1) * 36;
    for (const token of tokens) {
      const width = Math.max(12, token.length * 10);
      words.push({
        text: token,
        confidence: 40,
        x0: x,
        y0: y,
        x1: x + width,
        y1: y + 20,
        line: lineId,
      });
      x += width + 14;
    }
  });
  return words;
}

export async function countPdfPages(file: File): Promise<number> {
  const pdfjs = await import("pdfjs-dist");
  pdfjs.GlobalWorkerOptions.workerSrc = "/pdf.worker.min.mjs";
  const data = new Uint8Array(await file.arrayBuffer());
  const loadingTask = pdfjs.getDocument({ data });
  try {
    const pdf = await loadingTask.promise;
    return pdf.numPages;
  } finally {
    await loadingTask.destroy();
  }
}
