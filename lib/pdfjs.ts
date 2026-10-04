/** Legacy pdf.js build includes Map/WeakMap polyfills for older Chromium/Electron. */

export type PdfJs = typeof import("pdfjs-dist");

let loading: Promise<PdfJs> | null = null;

export function loadPdfJs(): Promise<PdfJs> {
  loading ??= import("pdfjs-dist/legacy/build/pdf.mjs").then((pdfjs) => {
    pdfjs.GlobalWorkerOptions.workerSrc = "/pdf.worker.min.mjs";
    return pdfjs as unknown as PdfJs;
  });
  return loading;
}
