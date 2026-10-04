import { qwenPagesToRecognition, type QwenPage } from "@/lib/qwen-act";
import type { ParsedAct } from "@/lib/types";
import type { RecognizeProgress } from "@/lib/recognize";

const MAX_PAGES = 40;

export async function recognizeWithQwen(
  file: File,
  onProgress: (progress: RecognizeProgress) => void,
  shouldCancel: () => void,
): Promise<{ act: ParsedAct; documents: ParsedAct[] }> {
  onProgress({ label: "Открываем скан для Qwen…", percent: 4 });
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
      const viewport = page.getViewport({ scale: Math.min(3, 2000 / longest) });
      const canvas = window.document.createElement("canvas");
      canvas.width = Math.ceil(viewport.width);
      canvas.height = Math.ceil(viewport.height);
      await page.render({ canvas, viewport }).promise;
      const image = await canvasToBase64(canvas);
      page.cleanup();
      const response = await fetch("/api/qwen", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ image }),
      });
      const body = (await response.json().catch(() => null)) as { page?: QwenPage; message?: string } | null;
      if (!response.ok || !body?.page) {
        throw new Error(body?.message || "Модель Qwen не смогла прочитать страницу.");
      }
      pages.push({ ...body.page, newAct: pageNumber > 1 ? body.page.newAct : true });
    }
    onProgress({ label: "Сопоставляем строки…", percent: 96 });
    return qwenPagesToRecognition(pages);
  } finally {
    await loadingTask.destroy();
  }
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
