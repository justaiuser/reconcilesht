import type { OcrWord, ParsedAct } from "@/lib/types";
import { stretchContrast } from "@/lib/image";
import { parseOcrDocuments, parseOcrPages } from "@/lib/ocr-table";

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
      const result = await worker.recognize(canvas, {}, { text: true, blocks: true });
      texts.push(result.data.text ?? "");
      pages.push(wordsFromResult(result.data.blocks as OcrBlock[] | null, result.data.text ?? ""));
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

function wordsFromResult(blocks: OcrBlock[] | null, text: string): OcrWord[] {
  const words: OcrWord[] = [];
  for (const block of blocks ?? []) {
    for (const paragraph of block.paragraphs ?? []) {
      for (const line of paragraph.lines ?? []) {
        for (const word of line.words ?? []) {
          if (!word.text.trim()) continue;
          words.push({
            text: word.text,
            confidence: word.confidence,
            x0: word.bbox.x0,
            y0: word.bbox.y0,
            x1: word.bbox.x1,
            y1: word.bbox.y1,
          });
        }
      }
    }
  }
  if (words.length > 0) return words;
  return wordsFromPlainText(text);
}

function wordsFromPlainText(text: string): OcrWord[] {
  const words: OcrWord[] = [];
  text.split(/\n/).forEach((line, lineIndex) => {
    let x = 0;
    const y = lineIndex * 36;
    for (const token of line.split(/\s+/).filter(Boolean)) {
      const width = Math.max(12, token.length * 10);
      words.push({ text: token, confidence: 40, x0: x, y0: y, x1: x + width, y1: y + 20 });
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
