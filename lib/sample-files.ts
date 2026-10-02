import {
  buildExcelMatrix,
  SAMPLE_PARTY_A,
  SAMPLE_PARTY_B,
  SAMPLE_PERIOD,
  SCAN_OURS,
  SCAN_THEIRS,
  sideTotals,
  type SampleLine,
} from "@/lib/sample-data";

function money(value: number | null | undefined): string | null {
  if (value == null) return null;
  return value.toFixed(2).replace(".", ",");
}

export async function buildSampleFiles(): Promise<{ excel: File; pdf: File }> {
  const { default: writeExcelFile } = await import("write-excel-file/browser");
  const excelBlob = await writeExcelFile(buildExcelMatrix(), { sheet: "Акт сверки" }).toBlob();
  const pdfBlob = await drawSamplePdf();
  return {
    excel: new File([excelBlob], "Акт сверки Север - Луч.xlsx", {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    }),
    pdf: new File([pdfBlob], "Акт сверки скан.pdf", { type: "application/pdf" }),
  };
}

async function loadActFont(): Promise<void> {
  const regular = new FontFace("ActSans", "url(/fonts/LiberationSans-Regular.ttf)", { weight: "400" });
  const bold = new FontFace("ActSans", "url(/fonts/LiberationSans-Bold.ttf)", { weight: "700" });
  const faces = await Promise.all([regular.load(), bold.load()]);
  for (const face of faces) document.fonts.add(face);
}

async function drawSamplePdf(): Promise<Blob> {
  await loadActFont();
  const width = 3200;
  const height = 1560;
  const canvas = document.createElement("canvas");
  canvas.width = width;
  canvas.height = height;
  const context = canvas.getContext("2d", { willReadFrequently: true });
  if (!context) throw new Error("Браузер не смог подготовить пример скана.");

  context.fillStyle = "#ffffff";
  context.fillRect(0, 0, width, height);
  context.fillStyle = "#000000";
  context.font = "700 54px ActSans, sans-serif";
  context.fillText("Акт сверки взаимных расчетов", 70, 100);
  context.font = "400 36px ActSans, sans-serif";
  context.fillText(`за период: ${SAMPLE_PERIOD}`, 70, 164);
  context.fillText(`между ${SAMPLE_PARTY_A} и ${SAMPLE_PARTY_B}`, 70, 216);

  const columns = [
    { x: 70, w: 460, align: "left" as const },
    { x: 550, w: 520, align: "left" as const },
    { x: 1090, w: 280, align: "right" as const },
    { x: 1390, w: 280, align: "right" as const },
    { x: 1780, w: 280, align: "left" as const },
    { x: 2080, w: 500, align: "left" as const },
    { x: 2600, w: 260, align: "right" as const },
    { x: 2880, w: 250, align: "right" as const },
  ];

  context.font = "700 32px ActSans, sans-serif";
  context.fillText(`По данным ${SAMPLE_PARTY_A}, руб.`, 70, 290);
  context.fillText(`По данным ${SAMPLE_PARTY_B}, руб.`, 1780, 290);

  let y = 370;
  const paint = (cells: (string | null)[], bold = false) => {
    context.font = `${bold ? 700 : 400} 40px ActSans, sans-serif`;
    context.fillStyle = "#000000";
    cells.forEach((cell, index) => {
      if (!cell) return;
      const column = columns[index];
      const measured = context.measureText(cell).width;
      const x = column.align === "right" ? column.x + column.w - measured : column.x;
      context.fillText(cell, Math.round(x), y);
    });
    y += 96;
  };

  paint(["Дата", "Документ", "Дебет", "Кредит", "Дата", "Документ", "Дебет", "Кредит"], true);
  const ours = sideTotals(10000, 0, SCAN_OURS);
  const theirs = sideTotals(0, 10000, SCAN_THEIRS);
  paint([
    "Сальдо начальное",
    null,
    money(ours.openingDebit),
    null,
    null,
    null,
    null,
    money(theirs.openingCredit),
  ]);

  const count = Math.max(SCAN_OURS.length, SCAN_THEIRS.length);
  for (let index = 0; index < count; index += 1) {
    paint(rowCells(SCAN_OURS[index], SCAN_THEIRS[index]));
  }
  paint([
    "Обороты за период",
    null,
    money(ours.turnoverDebit),
    money(ours.turnoverCredit),
    null,
    null,
    money(theirs.turnoverDebit),
    money(theirs.turnoverCredit),
  ]);
  paint([
    "Сальдо конечное",
    null,
    money(ours.closingDebit),
    null,
    null,
    null,
    null,
    money(theirs.closingCredit),
  ]);

  hardenInk(context, width, height);
  context.strokeStyle = "#d7dde6";
  context.lineWidth = 2;
  for (let line = 322; line < y; line += 96) {
    context.beginPath();
    context.moveTo(70, line);
    context.lineTo(width - 70, line);
    context.stroke();
  }

  const png = await new Promise<ArrayBuffer>((resolve, reject) => {
    canvas.toBlob(async (blob) => {
      if (!blob) {
        reject(new Error("Не удалось подготовить пример скана."));
        return;
      }
      resolve(await blob.arrayBuffer());
    }, "image/png");
  });

  const { PDFDocument } = await import("pdf-lib");
  const pdf = await PDFDocument.create();
  const image = await pdf.embedPng(png);
  const page = pdf.addPage([image.width, image.height]);
  page.drawImage(image, { x: 0, y: 0, width: image.width, height: image.height });
  const bytes = await pdf.save();
  return new Blob([Uint8Array.from(bytes)], { type: "application/pdf" });
}

function hardenInk(context: CanvasRenderingContext2D, width: number, height: number): void {
  const image = context.getImageData(0, 0, width, height);
  const { data } = image;
  for (let index = 0; index < data.length; index += 4) {
    const tone = data[index] * 0.2126 + data[index + 1] * 0.7152 + data[index + 2] * 0.0722;
    const ink = tone < 186 ? 0 : 255;
    data[index] = ink;
    data[index + 1] = ink;
    data[index + 2] = ink;
  }
  context.putImageData(image, 0, 0);
}

function rowCells(left?: SampleLine, right?: SampleLine): (string | null)[] {
  return [
    left?.date ?? null,
    left?.description ?? null,
    money(left?.debit),
    money(left?.credit),
    right?.date ?? null,
    right?.description ?? null,
    money(right?.debit),
    money(right?.credit),
  ];
}
