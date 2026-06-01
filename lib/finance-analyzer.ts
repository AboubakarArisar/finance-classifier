import { randomUUID } from "crypto";
import { existsSync } from "fs";
import { mkdir, readdir, readFile, rm, stat, writeFile } from "fs/promises";
import os from "os";
import path from "path";
import ExcelJS from "exceljs";
import * as XLSX from "xlsx";

export type UploadKind = "credit" | "bank";

export type UploadedWorkbook = {
  buffer: Buffer;
  fileName: string;
  kind: UploadKind;
};

export type AnalyzeResult = {
  downloadUrl: string;
  fileName: string;
  jobId: string;
  reportBase64: string;
  rowCount: number;
};

type MappingRule = {
  direction?: "הוצאה" | "הכנסה";
  keyword: string;
  mainCategory: string;
  recurrence?: string;
  subCategory: string;
};

type ParsedSheetSummary = {
  maxDate: string;
  minDate: string;
  originalFileName: string;
  originalSheetName: string;
  sourceName: string;
  totalAmount: number;
  transactionCount: number;
};

type ParsedSheet = {
  summary: ParsedSheetSummary;
  transactions: NormalizedTransaction[];
};

type NormalizedTransaction = {
  amount: number;
  cardOrAccount: string;
  chargeCurrency: string;
  date: string;
  description: string;
  direction: "הוצאה" | "הכנסה";
  installmentNote: string;
  mainCategory: string;
  note: string;
  originalAmount: number | "";
  recurrence: string;
  sourceName: string;
  subCategory: string;
};

const jobsDir = getJobsDir();
const mappingPath = path.join(/*turbopackIgnore: true*/ process.cwd(), "data", "category-mapping.xlsx");
const retentionMs = 30 * 24 * 60 * 60 * 1000;
const excelExtensions = [".xls", ".xlsx", ".xlsm"];

const classificationHeaders = [
  "מקור",
  "תאריך",
  "תיאור / שם בית העסק",
  "סכום",
  "הוצאה/הכנסה\n(ברירת מחדל: הוצאה)",
  "מחזוריות\n(ברירת מחדל: חודשי/מזדמן)",
  "סעיף ראשי",
  "שם סעיף",
  "הערות (מלל חופשי)",
  "סכום עסקה",
  "מטבע לחיוב",
  "מספר כרטיס/בנק",
  "",
  "לממוצע חודשי (לשימוש פנימי)",
  "מס' חודשים (לשימוש פנימי)",
  "מס' מופעים (לשימוש פנימי)",
];

const categorySheetRows = [
  ["מזון ופארמה", "פנאי, בילוי ותחביבים", "ביגוד והנעלה", "תכולת בית", "אחזקת בית", "טיפוח", "חינוך", "אירועים, תרומות, צרכי דת", "בריאות", "תחבורה", "משפחה", "תקשורת", "דיור", "התחייבויות", "נכסים", "פיננסים", "שכר", "קצבאות", "לא לסיווג"],
  ["מזון", "מסעדה ואוכל בחוץ", "ביגוד הורים", "ריהוט", "חשמל", "מספרה", "בית ספר", "חגים וצרכי דת", "קופ\"ח תשלום קבוע", "דלק", "ארועי שמחות במשפחה", "טלפון נייד ונייח", "משכנתה", "החזר חובות חודשי (למעט משכנתה) - כללי", "הפקדות לחסכונות - כללי", "עמלות", "שכר עבודה 1", "קצבת ילדים", "לא לסיווג"],
  ["פארמה וטואלטיקה", "ספורט", "ביגוד ילדים", "מוצרי חשמל ואלקטרוניקה", "מים וביוב", "קוסמטיקה", "מסגרות צהריים", "אירוע בעבודה / לחברים", "ביטוח רפואי נוסף", "חניה", "דמי כיס", "טלויזיה ואינטרנט (ספק ותשתית)", "שכר דירה", "ריביות משיכת יתר", "", "ביטוח חיים", "שכר עבודה 2", "קצבאות - כללי", "תשלומים, שקול רישום כחוב"],
  ["בר מים", "חופשות", "נעליים", "משחקים, צעצועים וספרים", "גז", "טיפוח - כללי", "מסגרות יום", "תרומות", "טיפולים פרטיים", "כבישי אגרה", "עזרה למשפחה", "שירותי תוכן", "מיסי ישוב / ועד בית", "", "", "ביטוח לאומי (למי שלא עובד)", "שכר עבודה 3", "סיוע בשכר דירה", ""],
  ["אוכל מוכן / בעבודה", "בילויים ומופעים", "ביגוד והנעלה - כללי", "כלי בית", "ניקיון", "", "צהרון / מטפלת", "", "תרופות", "ביטוח רכב", "תשלום מזונות", "תקשורת - כללי", "ארנונה", "", "", "פיננסים - כללי", "שכר - כללי", "", ""],
];

const fallbackRules: MappingRule[] = [
  { keyword: "מקס איט פיננסי", mainCategory: "לא לסיווג", subCategory: "לא לסיווג" },
  { keyword: "דירקט- מצטבר", mainCategory: "לא לסיווג", subCategory: "לא לסיווג" },
  { keyword: "WOLT", mainCategory: "פנאי, בילוי ותחביבים", subCategory: "מסעדה ואוכל בחוץ" },
  { keyword: "פיצה", mainCategory: "פנאי, בילוי ותחביבים", subCategory: "מסעדה ואוכל בחוץ" },
  { keyword: "קפה", mainCategory: "פנאי, בילוי ותחביבים", subCategory: "מסעדה ואוכל בחוץ" },
  { keyword: "רולדין", mainCategory: "פנאי, בילוי ותחביבים", subCategory: "מסעדה ואוכל בחוץ" },
  { keyword: "מקדונלדס", mainCategory: "פנאי, בילוי ותחביבים", subCategory: "מסעדה ואוכל בחוץ" },
  { keyword: "ארומה", mainCategory: "פנאי, בילוי ותחביבים", subCategory: "מסעדה ואוכל בחוץ" },
  { keyword: "גולדה", mainCategory: "פנאי, בילוי ותחביבים", subCategory: "מסעדה ואוכל בחוץ" },
  { keyword: "פאפא", mainCategory: "פנאי, בילוי ותחביבים", subCategory: "מסעדה ואוכל בחוץ" },
  { keyword: "סופר פארם", mainCategory: "מזון ופארמה", subCategory: "פארמה וטואלטיקה" },
  { keyword: "סופר-פארם", mainCategory: "מזון ופארמה", subCategory: "פארמה וטואלטיקה" },
  { keyword: "סופרפארם", mainCategory: "מזון ופארמה", subCategory: "פארמה וטואלטיקה" },
  { keyword: "דראגסטורס", mainCategory: "מזון ופארמה", subCategory: "פארמה וטואלטיקה" },
  { keyword: "שופרסל", mainCategory: "מזון ופארמה", subCategory: "מזון" },
  { keyword: "ויקטורי", mainCategory: "מזון ופארמה", subCategory: "מזון" },
  { keyword: "מחסני השוק", mainCategory: "מזון ופארמה", subCategory: "מזון" },
  { keyword: "NETFLIX", mainCategory: "תקשורת", subCategory: "שירותי תוכן" },
  { keyword: "DISNEY", mainCategory: "תקשורת", subCategory: "שירותי תוכן" },
  { keyword: "YOUTUBEPREMIUM", mainCategory: "תקשורת", subCategory: "שירותי תוכן" },
  { keyword: "HBOMAX", mainCategory: "תקשורת", subCategory: "שירותי תוכן" },
  { keyword: "APPLE.COM/BILL", mainCategory: "תקשורת", subCategory: "שירותי תוכן" },
  { keyword: "קצבת ילדים", mainCategory: "קצבאות", subCategory: "קצבת ילדים", direction: "הכנסה" },
  { keyword: "ביטוח לאומי", mainCategory: "קצבאות", subCategory: "קצבאות - כללי", direction: "הכנסה" },
  { keyword: "חברת החשמל", mainCategory: "אחזקת בית", subCategory: "חשמל" },
  { keyword: "מאוחדת שיניים", mainCategory: "בריאות", subCategory: "טיפולי שיניים / אורטודנט" },
  { keyword: "קופת חולים", mainCategory: "בריאות", subCategory: "קופ\"ח תשלום קבוע" },
  { keyword: "פיס", mainCategory: "פנאי, בילוי ותחביבים", subCategory: "הגרלות" },
  { keyword: "GETT", mainCategory: "תחבורה", subCategory: "תחבורה ציבורית" },
  { keyword: "חניוני", mainCategory: "תחבורה", subCategory: "חניה" },
  { keyword: "דלק", mainCategory: "תחבורה", subCategory: "דלק" },
  { keyword: "דור אלון", mainCategory: "תחבורה", subCategory: "דלק" },
  { keyword: "טופ טן", mainCategory: "תחבורה", subCategory: "דלק" },
  { keyword: "ביוטיקייר", mainCategory: "טיפוח", subCategory: "קוסמטיקה" },
  { keyword: "אורבניקה", mainCategory: "ביגוד והנעלה", subCategory: "ביגוד והנעלה - כללי" },
  { keyword: "דקטלון", mainCategory: "ביגוד והנעלה", subCategory: "ביגוד והנעלה - כללי" },
  { keyword: "שילב", mainCategory: "ביגוד והנעלה", subCategory: "ביגוד והנעלה - כללי" },
  { keyword: "סינמה", mainCategory: "פנאי, בילוי ותחביבים", subCategory: "בילויים ומופעים" },
];

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
  let creditSheetNumber = 1;
  let bankSheetNumber = 1;
  const parsedWorkbooks = files.map((file) =>
    parseWorkbook(file, mappings, (kind) => {
      if (kind === "credit") {
        return `אשראי-${creditSheetNumber++}`;
      }

      return `בנק-${bankSheetNumber++}`;
    }),
  );
  const transactions = parsedWorkbooks.flatMap((workbook) => workbook.transactions);
  const sheetSummaries = parsedWorkbooks.flatMap((workbook) => workbook.summaries);

  if (transactions.length === 0) {
    throw new Error("לא נמצאו תנועות בקבצים שהועלו. יש לבדוק שהקבצים כוללים טבלת פעולות.");
  }

  const reportBuffer = await buildReportWorkbook(transactions, sheetSummaries);
  const fileName = `shamir-classified-${jobId}.xlsx`;
  await writeFile(path.join(jobDir, "report.xlsx"), reportBuffer);

  return {
    downloadUrl: `/api/reports/${jobId}`,
    fileName,
    jobId,
    reportBase64: reportBuffer.toString("base64"),
    rowCount: transactions.length,
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
      direction: readObjectValue(row, ["direction", "type", "הוצאה/הכנסה"]) as MappingRule["direction"],
      keyword: readObjectValue(row, ["keyword", "מילת מפתח", "מפתח"]),
      mainCategory: readObjectValue(row, ["mainCategory", "category", "סעיף ראשי", "סיווג", "קטגוריה"]),
      recurrence: readObjectValue(row, ["recurrence", "מחזוריות"]),
      subCategory: readObjectValue(row, ["subCategory", "שם סעיף", "תת קטגוריה", "status"]),
    }))
    .filter((row) => row.keyword && row.mainCategory && row.subCategory);

  return mappedRows.length > 0 ? [...mappedRows, ...fallbackRules] : fallbackRules;
}

function parseWorkbook(
  file: UploadedWorkbook,
  mappings: MappingRule[],
  getNextSourceName: (kind: UploadKind) => string,
) {
  const workbook = XLSX.read(file.buffer, { cellDates: true, type: "buffer" });
  const summaries: ParsedSheetSummary[] = [];
  const transactions: NormalizedTransaction[] = [];

  workbook.SheetNames.forEach((sheetName) => {
    const rows = XLSX.utils.sheet_to_json<unknown[]>(workbook.Sheets[sheetName], {
      blankrows: false,
      defval: "",
      header: 1,
      raw: false,
    });
    const creditParsed = parseCreditSheet(rows, sheetName, file, mappings, () => getNextSourceName("credit"));
    const parsed =
      creditParsed.transactions.length > 0
        ? creditParsed
        : parseBankSheet(rows, sheetName, file, mappings, () => getNextSourceName("bank"));

    if (parsed.transactions.length === 0) {
      return;
    }

    transactions.push(...parsed.transactions);
    summaries.push(parsed.summary);
  });

  return { summaries, transactions };
}

function parseCreditSheet(
  rows: unknown[][],
  sheetName: string,
  file: UploadedWorkbook,
  mappings: MappingRule[],
  getSourceName: () => string,
): ParsedSheet {
  const headerIndex = findHeaderRowIndex(rows, ["תאריך עסקה", "שם בית העסק", "סכום חיוב"]);

  if (headerIndex < 0) {
    return emptyParsedSheet(file, sheetName, "");
  }

  const headers = rows[headerIndex].map((cell) => normalizeHeader(String(cell ?? "")));
  const indexes = {
    amount: findHeaderIndex(headers, ["סכום חיוב"]),
    card: findHeaderIndex(headers, ["4 ספרות אחרונות של כרטיס האשראי"]),
    category: findHeaderIndex(headers, ["קטגוריה"]),
    chargeCurrency: findHeaderIndex(headers, ["מטבע חיוב"]),
    date: findHeaderIndex(headers, ["תאריך עסקה"]),
    description: findHeaderIndex(headers, ["שם בית העסק"]),
    note: findHeaderIndex(headers, ["הערות"]),
    originalAmount: findHeaderIndex(headers, ["סכום עסקה מקורי"]),
  };

  const transactions = rows
    .slice(headerIndex + 1)
    .map<NormalizedTransaction | null>((row) => {
      const description = readRowText(row, indexes.description);
      const amount = readRowNumber(row, indexes.amount);

      if (!description || amount === null) {
        return null;
      }

      const note = readRowText(row, indexes.note);
      const sourceCategory = readRowText(row, indexes.category);
      const classification = classifyTransaction(description, note, amount, mappings, sourceCategory);

      return {
        amount: Math.abs(amount),
        cardOrAccount: readRowText(row, indexes.card),
        chargeCurrency: readRowText(row, indexes.chargeCurrency) || "₪",
        date: formatDateText(readRowText(row, indexes.date)),
        description,
        direction: amount < 0 ? "הכנסה" : classification.direction,
        installmentNote: note.includes("תשלום") ? note : "",
        mainCategory: classification.mainCategory,
        note,
        originalAmount: readRowNumber(row, indexes.originalAmount) ?? "",
        recurrence: classification.recurrence,
        sourceName: "",
        subCategory: classification.subCategory,
      } satisfies NormalizedTransaction;
    })
    .filter((row): row is NormalizedTransaction => Boolean(row));
  const sourceName = transactions.length > 0 ? getSourceName() : "";
  transactions.forEach((transaction) => {
    transaction.sourceName = sourceName;
  });

  return {
    summary: buildSummary(file.fileName, sheetName, sourceName, transactions),
    transactions,
  };
}

function parseBankSheet(
  rows: unknown[][],
  sheetName: string,
  file: UploadedWorkbook,
  mappings: MappingRule[],
  getSourceName: () => string,
): ParsedSheet {
  const headerIndex = findHeaderRowIndex(rows, ["תאריך", "הפעולה", "חובה", "זכות"]);

  if (headerIndex < 0) {
    return emptyParsedSheet(file, sheetName, "");
  }

  const accountNumber = extractAccountNumber(rows);
  const headers = rows[headerIndex].map((cell) => normalizeHeader(String(cell ?? "")));
  const indexes = {
    credit: findHeaderIndex(headers, ["זכות"]),
    date: findHeaderIndex(headers, ["תאריך"]),
    debit: findHeaderIndex(headers, ["חובה"]),
    description: findHeaderIndex(headers, ["הפעולה"]),
    details: findHeaderIndex(headers, ["פרטים"]),
  };

  const transactions = rows
    .slice(headerIndex + 1)
    .map<NormalizedTransaction | null>((row) => {
      const description = readRowText(row, indexes.description);
      const debit = readRowNumber(row, indexes.debit);
      const credit = readRowNumber(row, indexes.credit);
      const amount = debit ?? credit;

      if (!description || amount === null) {
        return null;
      }

      const details = readRowText(row, indexes.details);
      const direction = debit !== null ? "הוצאה" : "הכנסה";
      const classification = classifyTransaction(description, details, amount, mappings);

      return {
        amount: Math.abs(amount),
        cardOrAccount: accountNumber,
        chargeCurrency: "",
        date: formatDateText(readRowText(row, indexes.date)),
        description,
        direction: classification.direction ?? direction,
        installmentNote: "",
        mainCategory: classification.mainCategory,
        note: details,
        originalAmount: "",
        recurrence: classification.recurrence,
        sourceName: "",
        subCategory: classification.subCategory,
      } satisfies NormalizedTransaction;
    })
    .filter((row): row is NormalizedTransaction => Boolean(row));
  const sourceName = transactions.length > 0 ? getSourceName() : "";
  transactions.forEach((transaction) => {
    transaction.sourceName = sourceName;
  });

  return {
    summary: buildSummary(file.fileName, sheetName, sourceName, transactions),
    transactions,
  };
}

function emptyParsedSheet(file: UploadedWorkbook, sheetName: string, sourceName: string): ParsedSheet {
  return {
    summary: buildSummary(file.fileName, sheetName, sourceName, []),
    transactions: [],
  };
}

function buildSummary(
  originalFileName: string,
  originalSheetName: string,
  sourceName: string,
  transactions: NormalizedTransaction[],
): ParsedSheetSummary {
  const dates = transactions.map((transaction) => transaction.date).filter(Boolean).sort(compareDateText);

  return {
    maxDate: dates.at(-1) ?? "",
    minDate: dates[0] ?? "",
    originalFileName,
    originalSheetName,
    sourceName,
    totalAmount: transactions.reduce((sum, transaction) => sum + transaction.amount, 0),
    transactionCount: transactions.length,
  };
}

function classifyTransaction(
  description: string,
  note: string,
  amount: number,
  mappings: MappingRule[],
  sourceCategory = "",
) {
  if (/תשלום\s+\d+\s+מתוך\s+\d+/.test(note)) {
    return {
      direction: "הוצאה" as const,
      mainCategory: "לא לסיווג",
      recurrence: "",
      subCategory: "תשלומים, שקול רישום כחוב",
    };
  }

  const searchable = `${description} ${note}`.toLocaleLowerCase("he-IL");
  const match = mappings
    .slice()
    .sort((left, right) => right.keyword.length - left.keyword.length)
    .find((mapping) => searchable.includes(mapping.keyword.toLocaleLowerCase("he-IL")));

  if (match) {
    return {
      direction: match.direction ?? (amount < 0 ? "הכנסה" : "הוצאה"),
      mainCategory: match.mainCategory,
      recurrence: match.recurrence ?? "",
      subCategory: match.subCategory,
    };
  }

  const categoryFallback = classifyBySourceCategory(sourceCategory);
  return {
    direction: amount < 0 ? ("הכנסה" as const) : ("הוצאה" as const),
    mainCategory: categoryFallback?.mainCategory ?? "",
    recurrence: "",
    subCategory: categoryFallback?.subCategory ?? "",
  };
}

function classifyBySourceCategory(sourceCategory: string) {
  if (sourceCategory.includes("מסעדות")) {
    return { mainCategory: "פנאי, בילוי ותחביבים", subCategory: "מסעדה ואוכל בחוץ" };
  }

  if (sourceCategory.includes("שירותי תקשורת") || sourceCategory.includes("פנאי, בידור")) {
    return { mainCategory: "תקשורת", subCategory: "שירותי תוכן" };
  }

  if (sourceCategory.includes("מזון")) {
    return { mainCategory: "מזון ופארמה", subCategory: "מזון" };
  }

  if (sourceCategory.includes("ביטוח")) {
    return { mainCategory: "", subCategory: "" };
  }

  return null;
}

async function buildReportWorkbook(transactions: NormalizedTransaction[], summaries: ParsedSheetSummary[]) {
  const sortedTransactions = transactions.slice().sort((left, right) => compareDateText(left.date, right.date));
  const workbook = new ExcelJS.Workbook();
  workbook.creator = "Finance Classifier";
  workbook.created = new Date();
  workbook.views = [{ activeTab: 0, firstSheet: 0, height: 12000, visibility: "visible", width: 20000, x: 0, y: 0 }];

  appendExcelClassificationSheet(workbook, sortedTransactions);
  appendExcelSummarySheet(workbook, summaries);
  appendExcelCategorySheet(workbook);
  appendExcelImportSheet(workbook, sortedTransactions);
  appendExcelChoicesSheet(workbook);

  const buffer = await workbook.xlsx.writeBuffer();
  return Buffer.from(buffer);
}

function appendExcelSummarySheet(workbook: ExcelJS.Workbook, summaries: ParsedSheetSummary[]) {
  const sheet = workbook.addWorksheet("שמיר", {
    views: [{ rightToLeft: true }],
  });
  sheet.addRows([
    ["נוצר באמצעות Finance Classifier", new Date().toLocaleDateString("he-IL")],
    ["שם קובץ מקורי", "שם לשונית מקורי", "שם לשונית חדש", "מספר רשומות", "סכום כולל", "תאריך מאוחר", "תאריך מוקדם"],
    ...summaries.map((summary, index) => [
      summary.originalFileName,
      summary.originalSheetName,
      summary.sourceName || `מקור-${index + 1}`,
      summary.transactionCount,
      summary.totalAmount,
      summary.maxDate,
      summary.minDate,
    ]),
  ]);
  sheet.columns = [
    { width: 34 },
    { width: 26 },
    { width: 15 },
    { width: 14 },
    { width: 14 },
    { width: 14 },
    { width: 14 },
  ];
  styleHeaderRow(sheet.getRow(2));
}

function appendExcelClassificationSheet(workbook: ExcelJS.Workbook, transactions: NormalizedTransaction[]) {
  const sheet = workbook.addWorksheet("שלב ב - סיווג תנועות", {
    views: [{ rightToLeft: true, state: "frozen", ySplit: 8 }],
  });
  const lastRow = transactions.length + 8;
  const monthCount = Math.max(1, countDistinctMonths(transactions));

  sheet.addRows([
    ["שם המשפחה", "ישראלי"],
    ["תקופת השיקוף", monthCount],
    ["ממוצע הוצאות בחודש", { formula: `IFERROR(SUMIF(E9:E${lastRow},"הוצאה",D9:D${lastRow})/B2,0)` }],
    ["ממוצע הכנסות בחודש", { formula: `IFERROR(ABS(SUMIF(E9:E${lastRow},"הכנסה",D9:D${lastRow}))/B2,0)` }],
    ["מאזן חודשי", { formula: "B4-B3" }],
    ["שימו לב! ניתן לבחור ולתקן סעיף ראשי, שם סעיף והוצאה/הכנסה בעזרת הרשימות הנפתחות."],
    ["תנועות מתוך דפי בנק וכרטיסי אשראי"],
    classificationHeaders,
    ...transactions.map((transaction) => [
      transaction.sourceName,
      transaction.date,
      transaction.description,
      transaction.direction === "הוצאה" ? transaction.amount : -transaction.amount,
      transaction.direction,
      transaction.recurrence,
      transaction.mainCategory,
      transaction.subCategory,
      transaction.note,
      transaction.originalAmount,
      transaction.chargeCurrency,
      transaction.cardOrAccount,
      "",
      "",
      "",
      "1",
    ]),
  ]);

  sheet.columns = [
    { width: 14 },
    { width: 12 },
    { width: 34 },
    { width: 14 },
    { width: 16 },
    { width: 18 },
    { width: 24 },
    { width: 28 },
    { width: 34 },
    { width: 14 },
    { width: 12 },
    { width: 16 },
    { width: 4 },
    { width: 20 },
    { width: 20 },
    { width: 20 },
  ];
  sheet.autoFilter = `A8:P${lastRow}`;
  styleHeaderRow(sheet.getRow(8));
  ["B3", "B4", "B5"].forEach((cellAddress) => {
    sheet.getCell(cellAddress).numFmt = '[$₪-40D]#,##0;[Red]-[$₪-40D]#,##0;[$₪-40D]-';
    sheet.getCell(cellAddress).font = { bold: true };
  });
  sheet.getCell("A5").font = { bold: true };
  sheet.getCell("B5").fill = { fgColor: { argb: "FFEFEFEA" }, pattern: "solid", type: "pattern" };

  for (let rowNumber = 9; rowNumber <= lastRow; rowNumber += 1) {
    sheet.getCell(`D${rowNumber}`).numFmt = '[$₪-40D]#,##0.00;[Red]-[$₪-40D]#,##0.00';
    sheet.getCell(`E${rowNumber}`).dataValidation = {
      allowBlank: false,
      formulae: ["בחירות!$C$2:$C$3"],
      showErrorMessage: true,
      type: "list",
    };
    sheet.getCell(`F${rowNumber}`).dataValidation = {
      allowBlank: true,
      formulae: ["בחירות!$D$2:$D$6"],
      type: "list",
    };
    sheet.getCell(`G${rowNumber}`).dataValidation = {
      allowBlank: true,
      formulae: ["בחירות!$A$2:$A$20"],
      showErrorMessage: true,
      type: "list",
    };
    sheet.getCell(`H${rowNumber}`).dataValidation = {
      allowBlank: true,
      formulae: ["בחירות!$B$2:$B$80"],
      showErrorMessage: true,
      type: "list",
    };
  }
}

function appendExcelCategorySheet(workbook: ExcelJS.Workbook) {
  const sheet = workbook.addWorksheet("רשימת הסעיפים", {
    views: [{ rightToLeft: true }],
  });
  sheet.addRows(categorySheetRows);
  sheet.columns = categorySheetRows[0].map(() => ({ width: 22 }));
  styleHeaderRow(sheet.getRow(1));
}

function appendExcelImportSheet(workbook: ExcelJS.Workbook, transactions: NormalizedTransaction[]) {
  const sheet = workbook.addWorksheet("ליבוא מהתוכנה", {
    views: [{ rightToLeft: true }],
  });
  sheet.addRows([
    ["[אין למחוק או לערוך גיליון זה]", "", "", "", "", "", "", "", "מוכן ליבוא"],
    ["תאריך", "תיאור", "סכום", "הוצאה/הכנסה", "סעיף ראשי", "שם סעיף", "הערות"],
    ...transactions.map((transaction) => [
      transaction.date,
      transaction.description,
      transaction.direction === "הוצאה" ? transaction.amount : -transaction.amount,
      transaction.direction,
      transaction.mainCategory,
      transaction.subCategory,
      transaction.note,
    ]),
  ]);
  sheet.columns = [{ width: 12 }, { width: 36 }, { width: 14 }, { width: 14 }, { width: 24 }, { width: 28 }, { width: 40 }];
  styleHeaderRow(sheet.getRow(2));
}

function appendExcelChoicesSheet(workbook: ExcelJS.Workbook) {
  const sheet = workbook.addWorksheet("בחירות", {
    state: "veryHidden",
    views: [{ rightToLeft: true }],
  });
  const mainCategories = categorySheetRows[0].filter(Boolean);
  const subCategories = [...new Set(categorySheetRows.slice(1).flat().filter(Boolean))];
  const directions = ["הוצאה", "הכנסה"];
  const recurrences = ["", "חודשי/מזדמן", "שנתי", "דו-חודשי", "רבעוני", "חלוקת הסכום ב X"];
  const maxRows = Math.max(mainCategories.length, subCategories.length, directions.length, recurrences.length);

  sheet.addRow(["סעיף ראשי", "שם סעיף", "הוצאה/הכנסה", "מחזוריות"]);

  for (let index = 0; index < maxRows; index += 1) {
    sheet.addRow([
      mainCategories[index] ?? "",
      subCategories[index] ?? "",
      directions[index] ?? "",
      recurrences[index] ?? "",
    ]);
  }
}

function styleHeaderRow(row: ExcelJS.Row) {
  row.font = { bold: true };
  row.fill = { fgColor: { argb: "FFEFEFEA" }, pattern: "solid", type: "pattern" };
  row.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
  row.eachCell((cell) => {
    cell.border = {
      bottom: { color: { argb: "FFD8D4C8" }, style: "thin" },
      left: { color: { argb: "FFD8D4C8" }, style: "thin" },
      right: { color: { argb: "FFD8D4C8" }, style: "thin" },
      top: { color: { argb: "FFD8D4C8" }, style: "thin" },
    };
  });
}

function countDistinctMonths(transactions: NormalizedTransaction[]) {
  const months = new Set(
    transactions
      .map((transaction) => transaction.date.match(/^\d{2}\/(\d{2})\/(\d{4})$/))
      .filter((match): match is RegExpMatchArray => Boolean(match))
      .map((match) => `${match[2]}-${match[1]}`),
  );
  return months.size;
}

function findHeaderRowIndex(rows: unknown[][], requiredHeaders: string[]) {
  return rows.slice(0, 20).findIndex((row) => {
    const normalizedCells = row.map((cell) => normalizeHeader(String(cell ?? "")));
    return requiredHeaders.every((header) => normalizedCells.includes(normalizeHeader(header)));
  });
}

function findHeaderIndex(headers: string[], aliases: string[]) {
  return headers.findIndex((header) => aliases.map(normalizeHeader).includes(header));
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

  const raw = String(row[index] ?? "").trim();

  if (!raw) {
    return null;
  }

  const text = raw.replace(/[₪,\s]/g, "").replace(/^\((.*)\)$/, "-$1");
  const parsed = Number.parseFloat(text);
  return Number.isFinite(parsed) ? parsed : null;
}

function formatDateText(value: string) {
  const match = value.match(/^(\d{1,2})[./-](\d{1,2})[./-](\d{4})$/);

  if (!match) {
    return value;
  }

  const [, day, month, year] = match;
  return `${day.padStart(2, "0")}/${month.padStart(2, "0")}/${year}`;
}

function compareDateText(left: string, right: string) {
  return toDateTimestamp(left) - toDateTimestamp(right);
}

function toDateTimestamp(value: string) {
  const match = value.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);

  if (!match) {
    return 0;
  }

  const [, day, month, year] = match;
  return new Date(Number(year), Number(month) - 1, Number(day)).getTime();
}

function extractAccountNumber(rows: unknown[][]) {
  const text = rows
    .slice(0, 3)
    .flat()
    .map((cell) => String(cell ?? ""))
    .join(" ");
  const match = text.match(/\d{2}-\d{3}-\d{6}/);
  return match?.[0] ?? "";
}

function sanitizeFileName(fileName: string) {
  return path.basename(fileName).replace(/[^a-zA-Z0-9._\-\u0590-\u05FF]/g, "_");
}
