import type { TransactionRow, WorkbookKind } from "./types";

function normHeader(h: string): string {
  return h.toLowerCase().replace(/[^a-z0-9]/g, "");
}

function pickColumn(headers: string[], candidates: string[]): string {
  const normalized = headers.map((h) => ({ source: h, key: normHeader(h) }));
  for (const c of candidates) {
    const ck = normHeader(c);
    const m = normalized.find((h) => h.key === ck || h.key.includes(ck) || ck.includes(h.key));
    if (m?.source) return m.source;
  }
  return "";
}

/** Ordered rules: first keyword match wins within the same section. */
const EXPENSE_BUCKETS: { bucket: string; keywords: string[] }[] = [
  { bucket: "Salary & payroll", keywords: ["salary", "salaries", "wage", "wages", "payroll", "paycheque", "paycheck", "compensation", "bonus", "employee", "employer", "fica", "401k", "benefit", "health ins", "dental", "pto", "severance"] },
  { bucket: "Contractors", keywords: ["contractor", "1099", "freelance", "upwork", "fiverr"] },
  { bucket: "Rent & facilities", keywords: ["rent", "lease", "landlord", "office space", "facilities", "utilities", "electric", "water", "internet", "wifi"] },
  { bucket: "Software & subscriptions", keywords: ["software", "saas", "subscription", "aws", "azure", "gcp", "google cloud", "github", "slack", "zoom", "notion", "hosting", "domain", "heroku", "vercel"] },
  { bucket: "Marketing & sales", keywords: ["marketing", "advert", "ads", "google ads", "facebook", "linkedin", "seo", "campaign", "event", "sponsor"] },
  { bucket: "Professional fees", keywords: ["legal", "attorney", "lawyer", "accounting", "cpa", "audit", "consult", "advisor"] },
  { bucket: "Travel & meals", keywords: ["travel", "flight", "airline", "hotel", "uber", "lyft", "taxi", "meal", "restaurant", "entertainment"] },
  { bucket: "Insurance", keywords: ["insurance", "premium", "coverage"] },
  { bucket: "Taxes & licenses", keywords: ["tax", "irs", "franchise", "license", "permit", "registration", "state tax"] },
  { bucket: "Bank & processing fees", keywords: ["bank fee", "wire fee", "service charge", "processing fee", "stripe", "paypal", "merchant", "interest expense"] },
];

const REVENUE_BUCKETS: { bucket: string; keywords: string[] }[] = [
  { bucket: "Subscription & SaaS revenue", keywords: ["subscription", "mrr", "recurring", "saas revenue", "license revenue"] },
  { bucket: "Product & service sales", keywords: ["sales", "revenue", "invoice", "product", "service revenue", "consulting revenue"] },
  { bucket: "Interest & other income", keywords: ["interest income", "dividend", "other income", "gain on"] },
];

function rowSearchBlob(row: TransactionRow, headers: string[]): string {
  const keys = [
    pickColumn(headers, ["Account", "GL Name ", "GL Name", "GL Number", "Particulars"]),
    pickColumn(headers, ["Bank Description", "Description"]),
    pickColumn(headers, ["Contact", "Reference"]),
    pickColumn(headers, ["Assets"]),
  ].filter(Boolean) as string[];
  const bits = keys.map((k) => String(row[k] ?? ""));
  return bits.join(" ").toLowerCase();
}

function matchBucket(text: string, buckets: { bucket: string; keywords: string[] }[]): string | null {
  for (const { bucket, keywords } of buckets) {
    if (keywords.some((kw) => text.includes(kw))) return bucket;
  }
  return null;
}

function sheetContext(sheetName: string): "pl" | "bs" | "tb" | "tx" | "aging_ar" | "aging_ap" | "opening" | "bank" | "other" {
  const s = sheetName.toLowerCase();
  if (
    s.includes("profit and loss") ||
    s.includes("profit & loss") ||
    s.includes("p&l") ||
    (s.includes("profit") && s.includes("loss"))
  ) {
    return "pl";
  }
  if (s.includes("balance sheet")) return "bs";
  if (s.includes("trial") || /\btb\b/i.test(sheetName) || (s.includes("tb") && s.match(/\d{4}/))) return "tb";
  if (s.includes("transaction")) return "tx";
  if (s.includes("ar aging") || s.includes("receivable")) return "aging_ar";
  if (s.includes("ap aging") || s.includes("payable")) return "aging_ap";
  if (s.includes("opening")) return "opening";
  return "other";
}

/**
 * Optional path prefix: e.g. ["Expenses", "Salary & payroll"] before account/line grouping.
 * Skipped for chart-of-accounts workbooks.
 */
export function activityGroupPrefix(
  sheetName: string,
  row: TransactionRow,
  headers: string[],
  kind: WorkbookKind
): string[] {
  if (kind === "coa") return [];

  const text = rowSearchBlob(row, headers);
  const ctx = sheetContext(sheetName);

  /* Balance sheet: excel.ts already groups Assets → Account; skip extra prefix to avoid duplicates. */
  if (ctx === "bs") return [];

  if (ctx === "pl") {
    const exp = matchBucket(text, EXPENSE_BUCKETS);
    if (exp) return ["Expenses", exp];
    const rev = matchBucket(text, REVENUE_BUCKETS);
    if (rev) return ["Revenue", rev];
    if (text.match(/\b(income|revenue|sales)\b/) && !text.match(/expense|cost|fee/)) {
      return ["Revenue", "Other revenue"];
    }
    if (text.match(/expense|cost|fee|charge|payroll|salary/)) {
      return ["Expenses", matchBucket(text, EXPENSE_BUCKETS) ?? "Other expenses"];
    }
    return ["Profit & loss", "Other lines"];
  }

  if (ctx === "tb") {
    const exp = matchBucket(text, EXPENSE_BUCKETS);
    if (exp) return ["Expenses (by activity)", exp];
    const rev = matchBucket(text, REVENUE_BUCKETS);
    if (rev) return ["Revenue (by activity)", rev];
    if (text.match(/\b(asset|cash|receivable|prepaid|fixed)\b/)) return ["Balance sheet (TB)", "Assets & prepaids"];
    if (text.match(/\b(liabilit|payable|debt|loan|credit card)\b/)) return ["Balance sheet (TB)", "Liabilities"];
    if (text.match(/\b(equity|retained|capital|stock)\b/)) return ["Balance sheet (TB)", "Equity"];
    return ["Trial balance", "Other accounts"];
  }

  if (ctx === "tx") {
    const exp = matchBucket(text, EXPENSE_BUCKETS);
    if (exp) return ["Transactions", "Likely expenses · " + exp];
    const rev = matchBucket(text, REVENUE_BUCKETS);
    if (rev) return ["Transactions", "Likely revenue · " + rev];
    return ["Transactions", "All activity"];
  }

  if (ctx === "aging_ar") return ["Receivables", "AR aging"];
  if (ctx === "aging_ap") return ["Payables", "AP aging"];
  if (ctx === "opening") return ["Opening balances", "Lines"];

  if (kind === "bank") {
    const exp = matchBucket(text, EXPENSE_BUCKETS);
    if (exp) return ["Bank activity", "Outflows · " + exp];
    if (text.match(/\bdeposit|incoming|credit|interest earned|refund from\b/)) return ["Bank activity", "Inflows"];
    return ["Bank activity", "Other"];
  }

  if (kind === "financial" || kind === "generic") {
    const exp = matchBucket(text, EXPENSE_BUCKETS);
    if (exp) return ["Activity", "Expenses · " + exp];
    const rev = matchBucket(text, REVENUE_BUCKETS);
    if (rev) return ["Activity", "Revenue · " + rev];
  }

  return [];
}
