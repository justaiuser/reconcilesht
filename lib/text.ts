import type { Cell } from "@/lib/types";

export function round2(value: number): number {
  return Math.round(value * 100) / 100;
}

export function sameMoney(a: number, b: number): boolean {
  return Math.round(a * 100) === Math.round(b * 100);
}

export function parseMoney(raw: string): number | null {
  let source = raw.trim().replace(/\u00a0/g, "").replace(/\s/g, "");
  if (!source || source === "-" || source === "—" || source === "–") return null;

  let negative = false;
  if (source.startsWith("(") && source.endsWith(")")) {
    negative = true;
    source = source.slice(1, -1);
  }
  if (source.endsWith("-")) {
    negative = true;
    source = source.slice(0, -1);
  }
  if (source.startsWith("-")) {
    negative = true;
    source = source.slice(1);
  }
  if (/^\d{1,2}[./]\d{1,2}[./]\d{2,4}$/.test(source)) return null;

  if (source.includes(",") && source.includes(".")) {
    if (source.lastIndexOf(",") > source.lastIndexOf(".")) {
      source = source.replace(/\./g, "").replace(",", ".");
    } else {
      source = source.replace(/,/g, "");
    }
  } else if (source.includes(",")) {
    const parts = source.split(",");
    source =
      parts.length === 2 && parts[1].length <= 2
        ? `${parts[0].replace(/\./g, "")}.${parts[1]}`
        : source.replace(/,/g, "");
  } else if ((source.match(/\./g) ?? []).length > 1) {
    source = source.replace(/\./g, "");
  } else if (source.includes(".")) {
    const parts = source.split(".");
    if (parts[1] && parts[1].length === 3) source = source.replace(/\./g, "");
  }

  if (!/^\d+(\.\d+)?$/.test(source)) return null;
  const value = Number(source);
  if (!Number.isFinite(value)) return null;
  const rounded = round2(value);
  return negative ? -rounded : rounded;
}

export function formatMoney(value: number | null): string {
  if (value == null) return "—";
  return new Intl.NumberFormat("ru-RU", {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  }).format(value);
}

export function parseDate(raw: string): string | null {
  const match = raw.trim().match(/(\d{1,2})[./-](\d{1,2})[./-](\d{2,4})/);
  if (!match) return null;
  const day = Number(match[1]);
  const month = Number(match[2]);
  let year = Number(match[3]);
  if (year < 100) year += year >= 70 ? 1900 : 2000;
  if (month < 1 || month > 12 || day < 1 || day > 31) return null;
  return `${year}-${String(month).padStart(2, "0")}-${String(day).padStart(2, "0")}`;
}

export function formatDateRu(iso: string | null): string {
  if (!iso) return "—";
  const [year, month, day] = iso.split("-");
  return `${day}.${month}.${year}`;
}

export function normalizeHeader(value: string): string {
  return value
    .toLowerCase()
    .replace(/ё/g, "е")
    .replace(/\s+/g, " ")
    .trim();
}

export function summaryKind(text: string): "opening" | "turnover" | "closing" | null {
  const normalized = normalizeHeader(text);
  if (/сальдо началь|входящ[а-я]* сальдо/.test(normalized)) return "opening";
  if (/обороты/.test(normalized)) return "turnover";
  if (/сальдо конеч|исходящ[а-я]* сальдо/.test(normalized)) return "closing";
  return null;
}

export function extractDocNumber(text: string): string | null {
  if (!text.trim()) return null;
  const patterns = [
    /№\s*([0-9a-zа-яё][0-9a-zа-яё/-]{0,24})/i,
    /\(\s*(\d{1,6})\s+от(?:\s|$)/i,
    /(?:п\s*\/\s*п|пп)\s*№?\s*(\d{1,8})/i,
    /(?:упд|накладн\w*|акт|счет-фактур\w*|счёт-фактур\w*|счет|счёт)\s*№?\s*(\d{1,8})/i,
  ];
  for (const pattern of patterns) {
    const match = text.match(pattern);
    if (match?.[1]) return match[1];
  }
  return null;
}

export function docFromDocumentCell(text: string): string | null {
  const extracted = extractDocNumber(text);
  if (extracted) return extracted;
  const tokens = text.match(/\d{4,14}(?:-\d{1,6})?/g) ?? [];
  const withSuffix = tokens.find((token) => token.includes("-"));
  if (withSuffix) return withSuffix;
  const long = tokens.find((token) => token.length >= 5);
  if (long) return long;
  if (/^\d{1,8}$/.test(text.trim())) return text.trim();
  return null;
}

export function normalizeDoc(value: string | null): string | null {
  if (!value) return null;
  const compact = value
    .toLowerCase()
    .replace(/ё/g, "е")
    .replace(/[^0-9a-zа-я]/gi, "");
  const stripped = compact.replace(/^0+(?=\d)/, "");
  return stripped || null;
}

export function cellText(cell: Cell | undefined): string {
  if (cell == null || cell === false || cell === true) return "";
  if (cell instanceof Date) {
    const day = String(cell.getDate()).padStart(2, "0");
    const month = String(cell.getMonth() + 1).padStart(2, "0");
    return `${day}.${month}.${cell.getFullYear()}`;
  }
  if (typeof cell === "number") {
    if (Number.isInteger(cell)) return String(cell);
    return String(round2(cell)).replace(".", ",");
  }
  return String(cell).replace(/\u00a0/g, " ").trim();
}

export function cellAmount(cell: Cell | undefined): number | null {
  if (cell == null || cell === "" || typeof cell === "boolean") return null;
  if (cell instanceof Date) return null;
  if (typeof cell === "number") return Number.isFinite(cell) ? round2(cell) : null;
  return parseMoney(cell);
}

export function extractParties(text: string): {
  partyA: string | null;
  partyB: string | null;
} {
  const match = text.match(/между\s+(.+?)\s+и\s+([^\n]+)/i);
  if (!match) return { partyA: null, partyB: null };
  const clean = (value: string) => value.replace(/\s+/g, " ").replace(/[.,;]+$/, "").trim();
  return { partyA: clean(match[1]), partyB: clean(match[2]) };
}

export function extractPeriod(text: string): string | null {
  const labeled = text.match(/период[:\s]+([^\n]+)/i);
  if (labeled?.[1]) return labeled[1].replace(/\s+/g, " ").trim();
  const range = text.match(
    /с\s+\d{2}[./-]\d{2}[./-]\d{4}\s+(?:по|до)\s+\d{2}[./-]\d{2}[./-]\d{4}/i,
  );
  return range?.[0] ?? null;
}

export function lineLabel(description: string, docNumber: string | null): string {
  if (description.trim()) return description.trim();
  if (docNumber) return `№ ${docNumber}`;
  return "Без названия";
}
