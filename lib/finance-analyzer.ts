import { randomUUID } from "crypto";
import { existsSync } from "fs";
import { mkdir, readdir, readFile, rm, stat, writeFile } from "fs/promises";
import os from "os";
import path from "path";
import * as XLSX from "xlsx";

export type UploadKind = "credit" | "bank";

export type UploadedWorkbook = {
  buffer: Buffer;
  fileName: string;
  kind: UploadKind;
};

export type AnalyzeResult = {
  downloadUrl: string;
  jobId: string;
  rowCount: number;
};

type MappingRule = {
  category: string;
  keyword: string;
  status: string;
};

type NormalizedTransaction = {
  amount: number | null;
  date: string;
  description: string;
  source: string;
};

type ClassifiedTransaction = NormalizedTransaction & {
  category: string;
  status: string;
};

const jobsDir = getJobsDir();
const mappingPath = path.join(/*turbopackIgnore: true*/ process.cwd(), "data", "category-mapping.xlsx");
const retentionMs = 30 * 24 * 60 * 60 * 1000;
const excelExtensions = [".xls", ".xlsx", ".xlsm"];

const fallbackRules: MappingRule[] = [
  { keyword: "דלק", category: "רכב ונסיעות", status: "מוכר" },
  { keyword: "fuel", category: "רכב ונסיעות", status: "מוכר" },
  { keyword: "office", category: "ציוד משרדי", status: "מוכר" },
  { keyword: "software", category: "מערכות ותוכנה", status: "מוכר" },
  { keyword: "מע\"מ", category: "מסים ותשלומים", status: "מאומת" },
  { keyword: "vat", category: "מסים ותשלומים", status: "מאומת" },
  { keyword: "משכורת", category: "הכנסות", status: "מאומת" },
  { keyword: "salary", category: "הכנסות", status: "מאומת" },
];

const headerAliases = {
  amount: ["amount", "total", "סכום", "סכום עסקה", "סכום חיוב", "סכום בשח"],
  credit: ["credit", "deposit", "deposits", "זכות", "זיכוי", "הפקדה"],
  date: ["date", "transaction date", "posting date", "תאריך", "תאריך עסקה", "תאריך חיוב", "תאריך פעולה"],
  debit: ["debit", "withdrawal", "withdrawals", "חובה", "חיוב", "משיכה"],
  description: [
    "description",
    "details",
    "merchant",
    "payee",
    "business name",
    "תיאור",
    "תאור",
    "תיאור פעולה",
    "תאור פעולה",
    "פרטים",
    "פירוט",
    "בית עסק",
    "שם בית עסק",
  ],
};

export function isExcelFileName(fileName: string) {
  return excelExtensions.includes(path.extname(fileName).toLowerCase());
}

function getJobsDir() {
  if (process.env.FINANCE_DATA_DIR) {
    return path.join(process.env.FINANCE_DATA_DIR, "jobs");
  }

  const isServerless = Boolean(process.env.VERCEL || process.env.AWS_LAMBDA_FUNCTION_NAME || process.env.WEBSITE_INSTANCE_ID);
  return isServerless
    ? path.join(/*turbopackIgnore: true*/ os.tmpdir(), "benny-finance-classifier", "jobs")
    : path.join(/*turbopackIgnore: true*/ process.cwd(), ".data", "jobs");
}

export async function analyzeFinancialStatements(files: UploadedWorkbook[]): Promise<AnalyzeResult> {
  await mkdir(jobsDir, { recursive: true });
  await cleanupExpiredJobs();

  const jobId = randomUUID();
  const jobDir = path.join(jobsDir, jobId);
  const uploadsDir = path.join(jobDir, "uploads");
  await mkdir(uploadsDir, { recursive: true });

  for (const file of files) {
    await writeFile(path.join(uploadsDir, `${file.kind}-${sanitizeFileName(file.fileName)}`), file.buffer);
  }

  const mappings = await loadMappings();
  const transactions = files.flatMap((file) => parseWorkbook(file));

  if (transactions.length === 0) {
    throw new Error("לא נמצאו תנועות בקבצים שהועלו. יש לבדוק שהקבצים כוללים טבלת פעולות.");
  }

  const classifiedRows = transactions.map((transaction) => classifyTransaction(transaction, mappings));
  await writeFile(path.join(jobDir, "report.xlsx"), buildReportWorkbook(classifiedRows));

  return {
    downloadUrl: `/api/reports/${jobId}`,
    jobId,
    rowCount: classifiedRows.length,
  };
}

export async function readReport(jobId: string) {
  if (!/^[0-9a-f-]{36}$/i.test(jobId)) {
    return null;
  }

  const reportPath = path.join(jobsDir, jobId, "report.xlsx");

  if (!existsSync(reportPath)) {
    return null;
  }

  return readFile(reportPath);
}

async function cleanupExpiredJobs() {
  if (!existsSync(jobsDir)) {
    return;
  }

  const entries = await readdir(jobsDir, { withFileTypes: true });
  const now = Date.now();

  await Promise.all(
    entries
      .filter((entry) => entry.isDirectory())
      .map(async (entry) => {
        const jobPath = path.join(jobsDir, entry.name);
        const info = await stat(jobPath);

        if (now - info.mtimeMs > retentionMs) {
          await rm(jobPath, { force: true, recursive: true });
        }
      }),
  );
}

async function loadMappings(): Promise<MappingRule[]> {
  if (!existsSync(mappingPath)) {
    return fallbackRules;
  }

  const workbook = XLSX.read(await readFile(mappingPath), { type: "buffer" });
  const firstSheetName = workbook.SheetNames[0];

  if (!firstSheetName) {
    return fallbackRules;
  }

  const rows = XLSX.utils.sheet_to_json<Record<string, unknown>>(workbook.Sheets[firstSheetName], {
    defval: "",
  });
  const mappedRows = rows
    .map((row) => ({
      category: readObjectValue(row, ["category", "סיווג", "קטגוריה"]) || "לא מסווג",
      keyword: readObjectValue(row, ["keyword", "מילת מפתח", "מפתח"]),
      status: readObjectValue(row, ["status", "סטטוס"]) || "מוכר",
    }))
    .filter((row) => row.keyword.length > 0);

  return mappedRows.length > 0 ? mappedRows : fallbackRules;
}

function parseWorkbook(file: UploadedWorkbook): NormalizedTransaction[] {
  const workbook = XLSX.read(file.buffer, { cellDates: true, type: "buffer" });
  const firstSheetName = workbook.SheetNames[0];

  if (!firstSheetName) {
    return [];
  }

  const rows = XLSX.utils.sheet_to_json<unknown[]>(workbook.Sheets[firstSheetName], {
    blankrows: false,
    defval: "",
    header: 1,
    raw: false,
  });
  const headerIndex = findHeaderRowIndex(rows);

  if (headerIndex < 0) {
    return [];
  }

  const headers = rows[headerIndex].map((cell) => normalizeHeader(String(cell ?? "")));
  const indexes = detectColumnIndexes(headers);

  if (indexes.description === -1 && indexes.amount === -1 && indexes.debit === -1 && indexes.credit === -1) {
    return [];
  }

  return rows
    .slice(headerIndex + 1)
    .map((row) => normalizeTransaction(row, indexes, file.kind))
    .filter((transaction) => transaction.description || transaction.amount !== null || transaction.date);
}

function findHeaderRowIndex(rows: unknown[][]) {
  const firstNonEmpty = rows.findIndex((row) => row.some((cell) => String(cell ?? "").trim().length > 0));
  const likelyHeader = rows.slice(0, 12).findIndex((row) => {
    const normalizedCells = row.map((cell) => normalizeHeader(String(cell ?? "")));
    return (
      normalizedCells.some((cell) => headerAliases.date.includes(cell)) &&
      normalizedCells.some((cell) =>
        [...headerAliases.description, ...headerAliases.amount, ...headerAliases.debit, ...headerAliases.credit].includes(cell),
      )
    );
  });

  return likelyHeader >= 0 ? likelyHeader : firstNonEmpty;
}

function detectColumnIndexes(headers: string[]) {
  return {
    amount: findHeaderIndex(headers, headerAliases.amount),
    credit: findHeaderIndex(headers, headerAliases.credit),
    date: findHeaderIndex(headers, headerAliases.date),
    debit: findHeaderIndex(headers, headerAliases.debit),
    description: findHeaderIndex(headers, headerAliases.description),
  };
}

function findHeaderIndex(headers: string[], aliases: string[]) {
  return headers.findIndex((header) => aliases.includes(header));
}

function normalizeTransaction(
  row: unknown[],
  indexes: ReturnType<typeof detectColumnIndexes>,
  kind: UploadKind,
): NormalizedTransaction {
  const debit = readRowNumber(row, indexes.debit);
  const credit = readRowNumber(row, indexes.credit);
  const directAmount = readRowNumber(row, indexes.amount);

  return {
    amount: directAmount ?? (debit !== null ? -Math.abs(debit) : credit),
    date: readRowText(row, indexes.date),
    description: readRowText(row, indexes.description),
    source: kind === "credit" ? "כרטיס אשראי" : "חשבון בנק",
  };
}

function classifyTransaction(transaction: NormalizedTransaction, mappings: MappingRule[]): ClassifiedTransaction {
  const description = transaction.description.toLocaleLowerCase("he-IL");
  const match = mappings
    .slice()
    .sort((left, right) => right.keyword.length - left.keyword.length)
    .find((mapping) => description.includes(mapping.keyword.toLocaleLowerCase("he-IL")));

  return {
    ...transaction,
    category: match?.category ?? "לא מסווג",
    status: match?.status ?? "לבדיקה",
  };
}

function buildReportWorkbook(rows: ClassifiedTransaction[]) {
  const worksheetRows = [
    ["תאריך", "מקור", "תיאור פעולה", "סיווג", "סכום", "סטטוס"],
    ...rows.map((row) => [row.date, row.source, row.description, row.category, row.amount ?? "", row.status]),
  ];
  const workbook = XLSX.utils.book_new();
  const worksheet = XLSX.utils.aoa_to_sheet(worksheetRows);
  worksheet["!rtl"] = true;
  worksheet["!cols"] = [{ wch: 14 }, { wch: 16 }, { wch: 40 }, { wch: 22 }, { wch: 14 }, { wch: 14 }];
  workbook.Workbook = { Views: [{ RTL: true }] };
  XLSX.utils.book_append_sheet(workbook, worksheet, "סיווג פעולות");
  return XLSX.write(workbook, { bookType: "xlsx", type: "buffer" }) as Buffer;
}

function normalizeHeader(value: string) {
  return value.replace(/\s+/g, " ").trim().toLocaleLowerCase("he-IL");
}

function readObjectValue(row: Record<string, unknown>, candidates: string[]) {
  const entries = Object.entries(row);
  const match = entries.find(([key]) => candidates.includes(normalizeHeader(key)));
  return String(match?.[1] ?? "").trim();
}

function readRowText(row: unknown[], index: number) {
  if (index < 0) {
    return "";
  }

  return String(row[index] ?? "").trim();
}

function readRowNumber(row: unknown[], index: number) {
  if (index < 0) {
    return null;
  }

  const text = String(row[index] ?? "")
    .replace(/[₪,\s]/g, "")
    .replace(/[()]/g, "-")
    .trim();
  const parsed = Number.parseFloat(text);
  return Number.isFinite(parsed) ? parsed : null;
}

function sanitizeFileName(fileName: string) {
  return path.basename(fileName).replace(/[^a-zA-Z0-9._\-\u0590-\u05FF]/g, "_");
}
