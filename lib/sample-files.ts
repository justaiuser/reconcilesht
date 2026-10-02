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

async function drawSamplePdf(): Promise<Blob> {
  await document.fonts.load("600 34px Manrope");
  await document.fonts.load("30px Manrope");
  const width = 3400;
  const height = 1680;
  const canvas = document.createElement("canvas");
  canvas.width = width;
  canvas.height = height;
  const context = canvas.getContext("2d");
  if (!context) throw new Error("Браузер не смог подготовить пример скана.");

  context.fillStyle = "#ffffff";
  context.fillRect(0, 0, width, height);
  context.fillStyle = "#1c2430";
  context.font = "600 48px Manrope, sans-serif";
  context.fillText("Акт сверки взаимных расчетов", 70, 90);
  context.font = "30px Manrope, sans-serif";
  context.fillText(`за период: ${SAMPLE_PERIOD}`, 70, 145);
  context.fillText(`между ${SAMPLE_PARTY_A} и ${SAMPLE_PARTY_B}`, 70, 190);

  const columns = [
    { x: 70, w: 280, align: "left" as const },
    { x: 370, w: 620, align: "left" as const },
    { x: 1060, w: 280, align: "right" as const },
    { x: 1440, w: 280, align: "right" as const },
    { x: 1840, w: 280, align: "left" as const },
    { x: 2140, w: 580, align: "left" as const },
    { x: 2780, w: 250, align: "right" as const },
    { x: 3100, w: 230, align: "right" as const },
  ];

  context.font = "600 28px Manrope, sans-serif";
  context.fillText(`По данным ${SAMPLE_PARTY_A}, руб.`, 70, 260);
  context.fillText(`По данным ${SAMPLE_PARTY_B}, руб.`, 1840, 260);

  let y = 330;
  const paint = (cells: (string | null)[], bold = false) => {
    context.strokeStyle = "#d5dbe3";
    context.beginPath();
    context.moveTo(70, y + 18);
    context.lineTo(width - 70, y + 18);
    context.stroke();
    context.font = `${bold ? 600 : 400} 32px Manrope, sans-serif`;
    context.fillStyle = "#1c2430";
    cells.forEach((cell, index) => {
      if (!cell) return;
      const column = columns[index];
      const measured = context.measureText(cell).width;
      const x = column.align === "right" ? column.x + column.w - measured : column.x;
      context.fillText(cell, x, y);
    });
    y += 72;
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
