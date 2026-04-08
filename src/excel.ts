import * as XLSX from "xlsx";
import { activityGroupPrefix } from "./activityGroups";
import type { DataRow, LoadedWorkbook, SheetModel, TransactionRow, TreeNode, WorkbookKind } from "./types";

function normalizeHeader(h: string): string {
  return h.toLowerCase().replace(/[^a-z0-9]/g, "");
}

export function detectColumn(headers: string[], candidates: string[]): string {
  const normalized = headers.map((h) => ({
    source: h,
    key: normalizeHeader(h),
  }));
  for (const c of candidates) {
    const m = normalized.find((h) => h.key === normalizeHeader(c) || h.key === c.replace(/[^a-z0-9]/g, ""));
    if (m) return m.source;
  }
  for (const c of candidates) {
    const cn = c.replace(/[^a-z0-9]/g, "");
    const m = normalized.find((h) => h.key.includes(cn) || cn.includes(h.key));
    if (m && m.key.length > 0) return m.source;
  }
  return "";
}

function isBankWorkbook(sheetNames: string[]): boolean {
  if (sheetNames.length < 2) return false;
  const monthRe =
    /(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec|january|february|march|april|june|july|august|september|october|november|december)/i;
  const hits = sheetNames.filter((n) => monthRe.test(n)).length;
  return hits / sheetNames.length >= 0.5;
}

function detectWorkbookKind(sheetNames: string[]): WorkbookKind {
  const joined = sheetNames.join("|").toLowerCase();
  if (
    sheetNames.some(
      (n) => normalizeHeader(n) === "chartofaccounts" || n.toLowerCase().includes("chart of account")
    )
  ) {
    return "coa";
  }
  if (isBankWorkbook(sheetNames)) {
    return "bank";
  }
  if (
    joined.includes("tb") ||
    joined.includes("balance sheet") ||
    joined.includes("profit and loss") ||
    joined.includes("transactions")
  ) {
    return "financial";
  }
  return "generic";
}

function parseSheet(ws: XLSX.WorkSheet, sheetName: string, fileId: string): SheetModel {
  const rows = XLSX.utils.sheet_to_json<TransactionRow>(ws, { defval: "" });
  const headers = rows.length > 0 ? Object.keys(rows[0]) : [];
  const dataRows: DataRow[] = rows.map((row, idx) => ({
    id: `${fileId}::${sheetName}::${idx}`,
    fileId,
    sheetName,
    rowIndex: idx,
    row,
    label: buildRowLabel(row, headers, sheetName, idx),
  }));
  return { sheetName, headers, rows, dataRows };
}

function findDateColumn(headers: string[]): string {
  for (const h of headers) {
    const k = normalizeHeader(h);
    if (
      k === "date" ||
      k === "transactiondate" ||
      k === "billdate" ||
      k === "invoicedate" ||
      k.includes("transactiondate") ||
      (k.includes("date") && !k.includes("update"))
    ) {
      return h;
    }
  }
  return "";
}

function buildRowLabel(row: TransactionRow, headers: string[], _sheetName: string, idx: number): string {
  const dateKey = findDateColumn(headers);
  const descKey = detectColumn(headers, [
    "Bank Description",
    "Description",
    "GL Name ",
    "GL Name",
    "Reference",
    "Invoice Number",
    "Bill Reference",
  ]);
  const particularKey = detectColumn(headers, ["Particulars", "Contact"]);
  const glNum = detectColumn(headers, ["GL Number"]);
  const glName = detectColumn(headers, ["GL Name ", "GL Name"]);
  const amount = formatNetAmount(row, headers);

  const date = dateKey ? String(row[dateKey] ?? "").trim() : "";
  let desc = String(row[descKey] ?? row[particularKey] ?? "").trim();
  if (!desc && (glNum || glName)) {
    desc = [glNum ? String(row[glNum] ?? "").trim() : "", glName ? String(row[glName] ?? "").trim() : ""]
      .filter(Boolean)
      .join(" — ");
  }
  if (!desc) desc = `Row ${idx + 1}`;
  const left = [date, desc].filter(Boolean).join(" · ");
  return `${left} (${amount})`;
}

function formatNetAmount(row: TransactionRow, headers: string[]): string {
  const debitKey = detectColumn(headers, ["Debit", "Debits"]);
  const creditKey = detectColumn(headers, ["Credit", "Credits"]);
  const amountKey = detectColumn(headers, ["Amount", "Gross", "Total"]);

  if (debitKey && creditKey) {
    const d = Number(row[debitKey]);
    const c = Number(row[creditKey]);
    if (!Number.isNaN(d) && d !== 0) return formatNum(d);
    if (!Number.isNaN(c) && c !== 0) return formatNum(-c);
    if (!Number.isNaN(d) || !Number.isNaN(c)) return formatNum((Number.isNaN(d) ? 0 : d) - (Number.isNaN(c) ? 0 : c));
  }
  const single = amountKey ? row[amountKey] : "";
  const n = Number(single);
  if (!Number.isNaN(n)) return formatNum(n);
  return single === "" || single === undefined ? "—" : String(single);
}

function formatNum(n: number): string {
  return n.toLocaleString(undefined, { minimumFractionDigits: 0, maximumFractionDigits: 2 });
}

let fileCounter = 0;
function newFileId(): string {
  fileCounter += 1;
  return `f${Date.now()}_${fileCounter}`;
}

export async function parseWorkbookFile(file: File): Promise<LoadedWorkbook> {
  const bytes = await file.arrayBuffer();
  const wb = XLSX.read(bytes, { type: "array" });
  const fileId = newFileId();
  const baseName = file.name.replace(/\.(xlsx|xlsm|xlsb|xls)$/i, "");
  const kind = detectWorkbookKind(wb.SheetNames);

  const sheets: SheetModel[] = wb.SheetNames.map((name) => parseSheet(wb.Sheets[name], name, fileId));

  return {
    id: fileId,
    fileName: file.name,
    baseName,
    kind,
    sheets,
  };
}

function groupKeyForRow(
  row: TransactionRow,
  headers: string[],
  sheetName: string,
  kind: WorkbookKind,
  useActivity: boolean
): string[] {
  const norm = (s: string) => String(s ?? "").trim() || "—";

  let base: string[];

  if (kind === "coa" || normalizeHeader(sheetName).includes("chartofaccounts")) {
    const g = detectColumn(headers, ["Group"]);
    const sg = detectColumn(headers, ["Sub Group", "SubGroup"]);
    if (g && sg) base = [norm(String(row[g])), norm(String(row[sg]))];
    else if (g) base = [norm(String(row[g]))];
    else base = ["Rows"];
    return base;
  }

  const account = detectColumn(headers, ["Account", "GL Number", "GL Name", "GL Name "]);
  if (account) base = [norm(String(row[account]))];
  else {
    const assets = detectColumn(headers, ["Assets"]);
    const acct = detectColumn(headers, ["Account"]);
    if (assets && acct) base = [norm(String(row[assets])), norm(String(row[acct]))];
    else if (acct) base = [norm(String(row[acct]))];
    else {
      const contact = detectColumn(headers, ["Contact"]);
      if (contact) base = [norm(String(row[contact]))];
      else {
        const particulars = detectColumn(headers, ["Particulars"]);
        if (particulars) base = [norm(String(row[particulars]))];
        else {
          const dateCol = detectColumn(headers, ["Date", "Transaction Date", "Bill Date", "Invoice Date"]);
          if (dateCol) {
            const raw = String(row[dateCol] ?? "").trim();
            const month = raw.length >= 7 ? raw.slice(0, 7) : raw || "Undated";
            base = [month];
          } else base = ["Rows"];
        }
      }
    }
  }

  if (!useActivity) return base;
  const prefix = activityGroupPrefix(sheetName, row, headers, kind);
  if (prefix.length === 0) return base;
  return [...prefix, ...base];
}

export function buildTreeForWorkbooks(
  workbooks: LoadedWorkbook[],
  options?: { groupByActivity?: boolean }
): TreeNode[] {
  const groupByActivity = options?.groupByActivity !== false;
  return workbooks.map((wb) => ({
    type: "file" as const,
    id: `file-${wb.id}`,
    label: formatFileTreeLabel(wb),
    children: wb.sheets.map((sheet) => buildSheetNode(wb, sheet, groupByActivity)),
  }));
}

function formatFileTreeLabel(wb: LoadedWorkbook): string {
  const kindLabel =
    wb.kind === "financial"
      ? "Financial"
      : wb.kind === "coa"
        ? "Chart of accounts"
        : wb.kind === "bank"
          ? "Bank statements"
          : "Workbook";
  return `${wb.baseName} · ${kindLabel}`;
}

function buildSheetNode(wb: LoadedWorkbook, sheet: SheetModel, groupByActivity: boolean): TreeNode {
  const sheetId = `sheet-${wb.id}-${escapeId(sheet.sheetName)}`;
  const children = buildNestedGroups(sheetId, wb, sheet, groupByActivity);
  return {
    type: "sheet",
    id: sheetId,
    label: `${sheet.sheetName} (${sheet.rows.length})`,
    children,
  };
}

function escapeId(s: string): string {
  return s.replace(/[^a-zA-Z0-9_-]/g, "_");
}

/** Multi-level groups: activity (optional) → COA / account / date → rows. */
function buildNestedGroups(
  sheetId: string,
  wb: LoadedWorkbook,
  sheet: SheetModel,
  groupByActivity: boolean
): TreeNode[] {
  const rowSubset = sheet.dataRows;
  if (rowSubset.length === 0) return [];

  const getPath = (dr: DataRow) =>
    groupKeyForRow(dr.row, sheet.headers, sheet.sheetName, wb.kind, groupByActivity);

  function recurse(rows: DataRow[], depth: number): TreeNode[] {
    if (rows.length === 0) return [];
    if (rows.every((dr) => {
      const p = getPath(dr);
      return p.length === 1 && p[0] === "Rows";
    })) {
      return rows.map(leafNode);
    }

    const buckets = new Map<string, DataRow[]>();
    for (const dr of rows) {
      const p = getPath(dr);
      const seg = depth < p.length ? p[depth] : p[p.length - 1] ?? "—";
      if (!buckets.has(seg)) buckets.set(seg, []);
      buckets.get(seg)!.push(dr);
    }

    const out: TreeNode[] = [];
    for (const [seg, subset] of [...buckets.entries()].sort(([a], [b]) => a.localeCompare(b))) {
      const anyDeeper = subset.some((dr) => getPath(dr).length > depth + 1);
      if (!anyDeeper) {
        out.push({
          type: "group",
          id: `${sheetId}-d${depth}-${escapeId(seg)}-${subset.length}`,
          label: `${seg} (${subset.length})`,
          children: subset.map(leafNode),
        });
      } else {
        out.push({
          type: "group",
          id: `${sheetId}-d${depth}-${escapeId(seg)}`,
          label: seg,
          children: recurse(subset, depth + 1),
        });
      }
    }
    return out;
  }

  return recurse(rowSubset, 0);
}

function leafNode(r: DataRow): TreeNode {
  return {
    type: "leaf",
    id: r.id,
    label: r.label,
    rowRef: r,
  };
}

export function findDataRow(workbooks: LoadedWorkbook[], rowId: string): DataRow | null {
  for (const wb of workbooks) {
    for (const sh of wb.sheets) {
      const found = sh.dataRows.find((r) => r.id === rowId);
      if (found) return found;
    }
  }
  return null;
}

export function updateDataRowLabel(dr: DataRow, headers: string[]): void {
  dr.label = buildRowLabel(dr.row, headers.length ? headers : Object.keys(dr.row), dr.sheetName, dr.rowIndex);
}

export function getSheetHeaders(workbooks: LoadedWorkbook[], fileId: string, sheetName: string): string[] {
  const wb = workbooks.find((w) => w.id === fileId);
  const sh = wb?.sheets.find((s) => s.sheetName === sheetName);
  return sh?.headers ?? [];
}

export function parseInput(value: string): string | number {
  const t = value.trim();
  if (t === "") return "";
  if (/^-?\d+(\.\d+)?$/.test(t)) return Number(t);
  return value;
}

function excelSafeSheetName(name: string): string {
  const cleaned = name.replace(/[:\\/?*[\]]/g, "_").slice(0, 31);
  return cleaned || "Sheet";
}

/** Export one workbook (original sheet names preserved where possible). */
export function exportWorkbook(wb: LoadedWorkbook): void {
  const out = XLSX.utils.book_new();
  for (const sheet of wb.sheets) {
    const ws = XLSX.utils.json_to_sheet(sheet.rows);
    XLSX.utils.book_append_sheet(out, ws, excelSafeSheetName(sheet.sheetName));
  }
  XLSX.writeFile(out, `${wb.baseName}_updated.xlsx`);
}

/** Export each workbook; stagger slightly so browsers allow multiple downloads. */
export function exportAllWorkbooks(workbooks: LoadedWorkbook[]): void {
  workbooks.forEach((wb, index) => {
    setTimeout(() => exportWorkbook(wb), index * 450);
  });
}
