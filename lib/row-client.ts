import { pageTextLines, textLayerLooksUseful } from "@/lib/pdf-text";
import { loadPdfJs } from "@/lib/pdfjs";
import { printedPagesToRecognition, readPrintedPage, type PrintedPage } from "@/lib/printed-row";
import type { ParsedAct } from "@/lib/types";
import type { RecognizeProgress } from "@/lib/recognize";

const MAX_PAGES = 40;

export async function recognizeRows(
  file: File,
  onProgress: (progress: RecognizeProgress) => void,
  shouldCancel: () => void,
): Promise<{ act: ParsedAct; documents: ParsedAct[] }> {
  onProgress({ label: "Открываем скан…", percent: 4 });
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
    const pages: PrintedPage[] = [];
    let usedTextLayer = false;
    for (let pageNumber = 1; pageNumber <= pdf.numPages; pageNumber += 1) {
      shouldCancel();
      onProgress({
        label: `Читаем строки страницы ${pageNumber} из ${pdf.numPages}`,
        percent: 8 + ((pageNumber - 1) / pdf.numPages) * 84,
      });
      const source = await pdf.getPage(pageNumber);
      const textLines = await pageTextLines(source);
      if (textLayerLooksUseful(textLines)) {
        usedTextLayer = true;
        pages.push(readPrintedPage(textLines));
        source.cleanup();
        continue;
      }
      const unscaled = source.getViewport({ scale: 1 });
      const longest = Math.max(unscaled.width, unscaled.height);
      const viewport = source.getViewport({ scale: Math.min(4, 3400 / longest) });
      const canvas = window.document.createElement("canvas");
      canvas.width = Math.ceil(viewport.width);
      canvas.height = Math.ceil(viewport.height);
      await source.render({ canvas, viewport }).promise;
      source.cleanup();
      const image = await canvasToBase64(canvas);
      const response = await fetch("/api/rows", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ image }),
      });
      const body = (await response.json().catch(() => null)) as { page?: PrintedPage; message?: string } | null;
      if (!response.ok || !body?.page) {
        throw new Error(body?.message || "Не удалось прочитать строки страницы.");
      }
      pages.push(body.page);
    }
    onProgress({ label: "Сопоставляем строки…", percent: 96 });
    const recognized = printedPagesToRecognition(pages);
    if (usedTextLayer) {
      recognized.act.confidence = 92;
      recognized.act.warnings = [
        ...recognized.act.warnings,
        "Часть страниц прочитана из текстового слоя PDF (без OCR).",
      ];
      for (const document of recognized.documents) {
        document.confidence = document.confidence ?? 92;
      }
    }
    return recognized;
  } finally {
    await loadingTask.destroy();
  }
}

function canvasToBase64(canvas: HTMLCanvasElement): Promise<string> {
  return new Promise((resolve, reject) => {
    canvas.toBlob(
      (blob) => {
        if (!blob) {
          reject(new Error("Не удалось подготовить страницу."));
          return;
        }
        const reader = new FileReader();
        reader.onload = () => resolve(String(reader.result ?? "").replace(/^data:image\/jpeg;base64,/, ""));
        reader.onerror = () => reject(new Error("Не удалось подготовить страницу."));
        reader.readAsDataURL(blob);
      },
      "image/jpeg",
      0.92,
    );
  });
}
