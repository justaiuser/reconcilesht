export type Cell = string | number | boolean | Date | null;

export type SideName = "ours" | "theirs" | "single";

export type SideColumns = {
  date: number | null;
  document: number | null;
  debit: number | null;
  credit: number | null;
};

export type ColumnMap = {
  headerRow: number;
  dataStartRow: number;
  ours: SideColumns;
  theirs: SideColumns | null;
};

export type ActLine = {
  id: string;
  side: SideName;
  date: string | null;
  docNumber: string | null;
  description: string;
  debit: number | null;
  credit: number | null;
};

export type ActSummary = {
  openingDebit: number | null;
  openingCredit: number | null;
  turnoverDebit: number | null;
  turnoverCredit: number | null;
  closingDebit: number | null;
  closingCredit: number | null;
};

export type ParsedAct = {
  title: string | null;
  period: string | null;
  partyA: string | null;
  partyB: string | null;
  lines: ActLine[];
  summary: ActSummary;
  theirsSummary: ActSummary | null;
  warnings: string[];
  confidence: number | null;
  rawText: string;
  twoSided: boolean;
};

export type PairStatus = "match" | "amount" | "excel-only" | "scan-only";

export type Pair = {
  status: PairStatus;
  excel: ActLine | null;
  scan: ActLine | null;
  notes: string[];
};

export type SummaryStatus = "match" | "diff" | "missing";

export type SummaryRow = {
  label: string;
  excel: number | null;
  scan: number | null;
  status: SummaryStatus;
};

export type SideReport = {
  title: string;
  side: SideName;
  mirror: boolean;
  pairs: Pair[];
  summary: SummaryRow[];
};

export type CompareMode = "auto" | "both" | "ours" | "theirs";
export type MirrorMode = "auto" | "on" | "off";

export type CompareLabels = {
  left: string;
  right: string;
  onlyLeft: string;
  onlyRight: string;
  debitLeft: string;
  creditLeft: string;
  debitRight: string;
  creditRight: string;
  mirrorHint: string;
};

export type CompareResult = {
  reports: SideReport[];
  interpretation: string;
  labels: CompareLabels;
  columnsOnly: boolean;
};

export type OcrWord = {
  text: string;
  confidence: number;
  x0: number;
  y0: number;
  x1: number;
  y1: number;
};
