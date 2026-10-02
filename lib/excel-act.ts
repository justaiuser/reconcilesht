import type {
  ActLine,
  ActSummary,
  Cell,
  ColumnMap,
  ParsedAct,
  SideColumns,
} from "@/lib/types";
import {
  cellAmount,
  cellText,
  docFromDocumentCell,
  extractParties,
  extractPeriod,
  normalizeHeader,
  parseDate,
} from "@/lib/text";

export function emptySummary(): ActSummary {
  return {
    openingDebit: null,
    openingCredit: null,
    turnoverDebit: null,
    turnoverCredit: null,
    closingDebit: null,
    closingCredit: null,
  };
}

type HeaderRole = "date" | "document" | "debit" | "credit" | "other";

function hasStem(text: string, stems: string[]): boolean {
  const parts = text.split(/[^a-zа-я0-9]+/i).filter(Boolean);
  return stems.some((stem) => parts.some((part) => part.startsWith(stem)));
}

export function headerRole(text: string): HeaderRole {
  const normalized = normalizeHeader(text);
  if (!normalized) return "other";
  if (hasStem(normalized, ["дата"])) return "date";
  if (hasStem(normalized, ["дебет", "дт"])) return "debit";
  if (hasStem(normalized, ["кредит", "кт"])) return "credit";
  if (hasStem(normalized, ["документ", "содержание", "операц", "наименован", "номер"])) {
    return "document";
  }
  return "other";
}

function scoreHeaderRow(row: Cell[] | undefined): number {
  if (!row) return 0;
  let score = 0;
  let dates = 0;
  let amounts = 0;
  for (const cell of row) {
    const role = headerRole(cellText(cell));
    if (role === "date") {
      score += 3;
      dates += 1;
    } else if (role === "document") score += 2;
    else if (role === "debit" || role === "credit") {
      score += 3;
      amounts += 1;
    }
  }
  if (dates === 0 || amounts === 0) return 0;
  return score;
}

function mergedHeader(matrix: Cell[][], rowIndex: number): { row: Cell[]; span: number } {
  const row = matrix[rowIndex] ?? [];
  const next = matrix[rowIndex + 1];
  if (!next) return { row, span: 1 };
  const merged = row.map((cell, index) => {
    if (headerRole(cellText(cell)) !== "other") return cell;
    const below = next[index];
    return below != null && headerRole(cellText(below)) !== "other" ? below : cell;
  });
  if (scoreHeaderRow(merged) > scoreHeaderRow(row)) return { row: merged, span: 2 };
  return { row, span: 1 };
}

function groupSides(
  columns: { index: number; role: Exclude<HeaderRole, "other"> }[],
): SideColumns[] {
  const groups: SideColumns[] = [];
  let current: SideColumns | null = null;

  const push = () => {
    if (!current) return;
    const useful =
      current.date != null ||
      current.document != null ||
      current.debit != null ||
      current.credit != null;
    if (useful) groups.push(current);
  };

  for (const column of columns) {
    if (column.role === "date" && current?.date != null) {
      push();
      current = null;
    }
    current ??= { date: null, document: null, debit: null, credit: null };
    if (current[column.role] == null) current[column.role] = column.index;
  }
  push();
  return groups.filter(
    (group) =>
      (group.date != null || group.document != null) &&
      (group.debit != null || group.credit != null),
  );
}

export function detectColumns(matrix: Cell[][]): ColumnMap | null {
  let bestIndex = -1;
  let bestScore = 0;
  let bestOwn = 0;
  const limit = Math.min(matrix.length, 40);
  for (let index = 0; index < limit; index += 1) {
    const own = scoreHeaderRow(matrix[index]);
    const score = Math.max(own, scoreHeaderRow(mergedHeader(matrix, index).row));
    if (score > bestScore || (score === bestScore && own > bestOwn)) {
      bestScore = score;
      bestOwn = own;
      bestIndex = index;
    }
  }
  if (bestIndex < 0) return null;
  return mapFromHeaderRow(matrix, bestIndex);
}

export function headerCells(matrix: Cell[][], headerRow: number): Cell[] {
  return mergedHeader(matrix, headerRow).row;
}

export function mapFromHeaderRow(matrix: Cell[][], headerRow: number): ColumnMap | null {
  const merged = mergedHeader(matrix, headerRow);
  const columns: { index: number; role: Exclude<HeaderRole, "other"> }[] = [];
  merged.row.forEach((cell, index) => {
    const role = headerRole(cellText(cell));
    if (role !== "other") columns.push({ index, role });
  });
  const groups = groupSides(columns);
  if (!groups[0]) return null;
  return {
    headerRow,
    dataStartRow: headerRow + merged.span,
    ours: groups[0],
    theirs: groups[1] ?? null,
  };
}

export function blankMap(columnCount = 4): ColumnMap {
  return {
    headerRow: 0,
    dataStartRow: 1,
    ours: {
      date: 0,
      document: columnCount > 1 ? 1 : null,
      debit: columnCount > 2 ? 2 : null,
      credit: columnCount > 3 ? 3 : null,
    },
    theirs: null,
  };
}

function summaryKind(text: string): "opening" | "turnover" | "closing" | null {
  const normalized = normalizeHeader(text);
  if (/сальдо началь/.test(normalized)) return "opening";
  if (/обороты/.test(normalized)) return "turnover";
  if (/сальдо конеч/.test(normalized)) return "closing";
  return null;
}

function applySummary(
  summary: ActSummary,
  kind: "opening" | "turnover" | "closing",
  debit: number | null,
  credit: number | null,
) {
  if (kind === "opening") {
    if (debit != null) summary.openingDebit = debit;
    if (credit != null) summary.openingCredit = credit;
  } else if (kind === "turnover") {
    if (debit != null) summary.turnoverDebit = debit;
    if (credit != null) summary.turnoverCredit = credit;
  } else {
    if (debit != null) summary.closingDebit = debit;
    if (credit != null) summary.closingCredit = credit;
  }
}

function readSideLine(
  row: Cell[],
  columns: SideColumns,
  side: ActLine["side"],
  index: number,
): ActLine | null {
  const dateRaw = columns.date == null ? "" : cellText(row[columns.date]);
  const documentRaw = columns.document == null ? "" : cellText(row[columns.document]);
  const debit = columns.debit == null ? null : cellAmount(row[columns.debit]);
  const credit = columns.credit == null ? null : cellAmount(row[columns.credit]);
  const date = parseDate(dateRaw);
  const description = documentRaw;
  const docNumber = docFromDocumentCell(documentRaw) ?? docFromDocumentCell(dateRaw);
  if (!date && !docNumber && debit == null && credit == null) return null;
  if (!date && debit == null && credit == null) return null;
  return {
    id: `${side}-${index}`,
    side,
    date,
    docNumber,
    description,
    debit,
    credit,
  };
}

export function parseExcelMatrix(matrix: Cell[][], map: ColumnMap): ParsedAct {
  const preamble = matrix
    .slice(0, map.headerRow)
    .map((row) => row.map((cell) => cellText(cell)).filter(Boolean).join(" "))
    .filter(Boolean);
  const preambleText = preamble.join("\n");
  const parties = extractParties(preambleText);
  const summary = emptySummary();
  const theirsSummary = map.theirs ? emptySummary() : null;
  const lines: ActLine[] = [];
  const warnings: string[] = [];

  for (let index = map.dataStartRow; index < matrix.length; index += 1) {
    const row = matrix[index] ?? [];
    const text = row.map((cell) => cellText(cell)).filter(Boolean).join(" ");
    if (!text) continue;
    const kind = summaryKind(text);
    if (kind) {
      applySummary(
        summary,
        kind,
        map.ours.debit == null ? null : cellAmount(row[map.ours.debit]),
        map.ours.credit == null ? null : cellAmount(row[map.ours.credit]),
      );
      if (map.theirs && theirsSummary) {
        applySummary(
          theirsSummary,
          kind,
          map.theirs.debit == null ? null : cellAmount(row[map.theirs.debit]),
          map.theirs.credit == null ? null : cellAmount(row[map.theirs.credit]),
        );
      }
      continue;
    }

    const ours = readSideLine(row, map.ours, map.theirs ? "ours" : "single", index);
    if (ours) lines.push(ours);
    if (map.theirs) {
      const theirs = readSideLine(row, map.theirs, "theirs", index);
      if (theirs) lines.push(theirs);
    }
  }

  if (lines.length === 0) {
    warnings.push(
      "В Excel не нашлось строк с датой и суммой. Проверьте, что колонки даты, документа и сумм указаны верно.",
    );
  }

  return {
    title: preamble[0] ?? null,
    period: extractPeriod(preambleText),
    partyA: parties.partyA,
    partyB: parties.partyB,
    lines,
    summary,
    theirsSummary,
    warnings,
    confidence: null,
    rawText: "",
    twoSided: map.theirs != null,
  };
}
