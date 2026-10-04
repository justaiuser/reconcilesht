/** Extract printable lines from a PDF page text layer (no OCR). */

type Positioned = {
  text: string;
  x: number;
  y: number;
  height: number;
};

export async function pageTextLines(page: {
  // pdf.js returns TextItem | TextMarkedContent; we only need string items.
  getTextContent: () => Promise<{ items: unknown[] }>;
}): Promise<string[]> {
  const content = await page.getTextContent();
  const items: Positioned[] = [];
  for (const raw of content.items) {
    const item = raw as { str?: string; transform?: number[]; height?: number };
    const text = (item.str ?? "").replace(/\s+/g, " ").trim();
    if (!text || !item.transform || item.transform.length < 6) continue;
    items.push({
      text,
      x: item.transform[4] ?? 0,
      y: item.transform[5] ?? 0,
      height: Math.abs(item.height ?? item.transform[3] ?? 10) || 10,
    });
  }
  if (items.length === 0) return [];

  items.sort((a, b) => b.y - a.y || a.x - b.x);
  const lines: Positioned[][] = [];
  for (const item of items) {
    const last = lines.at(-1);
    const band = last ? last.reduce((sum, part) => sum + part.y, 0) / last.length : 0;
    const threshold = Math.max(item.height, last?.[0]?.height ?? 10) * 0.6;
    if (last && Math.abs(item.y - band) <= threshold) last.push(item);
    else lines.push([item]);
  }

  return lines
    .map((line) =>
      line
        .sort((a, b) => a.x - b.x)
        .map((part) => part.text)
        .join(" ")
        .replace(/\s+/g, " ")
        .trim(),
    )
    .filter(Boolean);
}

export function textLayerLooksUseful(lines: string[]): boolean {
  if (lines.length < 3) return false;
  const joined = lines.join("\n");
  const dates = (joined.match(/\d{1,2}[./]\d{1,2}[./]\d{2,4}/g) ?? []).length;
  const money = (joined.match(/\d+[.,]\d{2}/g) ?? []).length;
  return dates >= 1 && money >= 2;
}
