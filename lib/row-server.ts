import path from "node:path";
import { createWorker, PSM, type Worker } from "tesseract.js";
import { decode as decodeJpeg, encode as encodeJpeg } from "jpeg-js";
import { textRows } from "@/lib/page-strips";
import { readPrintedPage, type PrintedPage } from "@/lib/printed-row";

let workerPromise: Promise<Worker> | null = null;

function worker(): Promise<Worker> {
  workerPromise ??= createWorker("rus", 1, {
    langPath: path.join(process.cwd(), "public/tessdata"),
    gzip: true,
  }).then(async (created) => {
    await created.setParameters({
      tessedit_pageseg_mode: PSM.SINGLE_LINE,
      preserve_interword_spaces: "1",
      user_defined_dpi: "300",
    });
    return created;
  });
  return workerPromise;
}

export async function readPageImage(image: Buffer): Promise<PrintedPage> {
  const decoded = decodeJpeg(image, { useTArray: true, formatAsRGBA: true });
  const bands = textRows(decoded.data, decoded.width, decoded.height, 4).filter((band) => {
    const height = band.bottom - band.top;
    return height >= 14 && height <= 100;
  });
  const reader = await worker();
  const texts: string[] = [];
  for (const band of bands) {
    const top = Math.max(0, band.top - 3);
    const bottom = Math.min(decoded.height, band.bottom + 3);
    const crop = cropJpeg(decoded.data, decoded.width, top, bottom);
    const result = await reader.recognize(crop);
    const text = (result.data.text ?? "").replace(/\s+/g, " ").trim();
    if (text) texts.push(text);
  }
  return readPrintedPage(texts);
}

function cropJpeg(pixels: Uint8Array, width: number, top: number, bottom: number): Buffer {
  const height = Math.max(1, bottom - top);
  const data = Buffer.from(pixels.subarray(top * width * 4, (top + height) * width * 4));
  return Buffer.from(encodeJpeg({ data, width, height }, 90).data);
}
