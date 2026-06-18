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

type CategoryGroup = {
  mainCategory: string;
  namedRange: string;
  subCategories: string[];
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

const categoryCatalog: CategoryGroup[] = [
  { mainCategory: "מזון ופארמה", namedRange: "CategorySubList_1", subCategories: ["מזון", "פארמה וטואלטיקה", "בר מים", "אוכל מוכן / בעבודה", "עישון", "מזון ופארמה - כללי"] },
  { mainCategory: "פנאי, בילוי ותחביבים", namedRange: "CategorySubList_2", subCategories: ["מסעדה ואוכל בחוץ", "ספורט", "חופשות", "בילויים ומופעים", "חיות מחמד", "חוגי מבוגרים", "בייביסיטר", "הגרלות", "פנאי - כללי"] },
  { mainCategory: "ביגוד והנעלה", namedRange: "CategorySubList_3", subCategories: ["ביגוד הורים", "ביגוד ילדים", "נעליים", "ביגוד והנעלה - כללי"] },
  { mainCategory: "תכולת בית", namedRange: "CategorySubList_4", subCategories: ["ריהוט", "מוצרי חשמל ואלקטרוניקה", "משחקים, צעצועים וספרים", "כלי בית", "תכולת בית - כללי"] },
  { mainCategory: "אחזקת בית", namedRange: "CategorySubList_5", subCategories: ["חשמל", "מים וביוב", "גז", "ניקיון", "תיקונים בבית / במכשירים", "גינה", "אחזקת בית - כללי"] },
  { mainCategory: "טיפוח", namedRange: "CategorySubList_6", subCategories: ["מספרה", "קוסמטיקה", "טיפוח - כללי"] },
  { mainCategory: "חינוך", namedRange: "CategorySubList_7", subCategories: ["בית ספר", "מסגרות צהריים", "מסגרות יום", "צהרון / מטפלת", "הסעות", "שיעור פרטי", "מסגרות קיץ", "חוגים ותנועת נוער", "לימודים והשתלמות לבוגרים", "חינוך - כללי"] },
  { mainCategory: "אירועים, תרומות, צרכי דת", namedRange: "CategorySubList_8", subCategories: ["חגים וצרכי דת", "אירוע בעבודה / לחברים", "תרומות"] },
  { mainCategory: "בריאות", namedRange: "CategorySubList_9", subCategories: ["קופ\"ח תשלום קבוע", "ביטוח רפואי נוסף", "טיפולים פרטיים", "תרופות", "טיפולי שיניים / אורטודנט", "אופטיקה", "בריאות - כללי"] },
  { mainCategory: "תחבורה", namedRange: "CategorySubList_10", subCategories: ["דלק", "חניה", "כבישי אגרה", "ביטוח רכב", "תחזוקת רכב", "תחבורה ציבורית", "רישוי רכב", "תחבורה שיתופית", "ליסינג", "תחבורה - כללי"] },
  { mainCategory: "משפחה", namedRange: "CategorySubList_11", subCategories: ["ארועי שמחות במשפחה", "דמי כיס", "עזרה למשפחה", "תשלום מזונות", "משפחה - כללי"] },
  { mainCategory: "תקשורת", namedRange: "CategorySubList_12", subCategories: ["טלפון נייד ונייח", "טלויזיה ואינטרנט (ספק ותשתית)", "שירותי תוכן", "תקשורת - כללי"] },
  { mainCategory: "דיור", namedRange: "CategorySubList_13", subCategories: ["משכנתה", "שכר דירה", "מיסי ישוב / ועד בית", "ארנונה", "ביטוח נכס ותכולה", "דיור - כללי"] },
  { mainCategory: "התחייבויות", namedRange: "CategorySubList_14", subCategories: ["החזר חובות חודשי (למעט משכנתה) - כללי", "ריביות משיכת יתר"] },
  { mainCategory: "נכסים", namedRange: "CategorySubList_15", subCategories: ["הפקדות לחסכונות - כללי"] },
  { mainCategory: "פיננסים", namedRange: "CategorySubList_16", subCategories: ["עמלות", "ביטוח חיים", "ביטוח לאומי (למי שלא עובד)", "פיננסים - כללי"] },
  { mainCategory: "שכר", namedRange: "CategorySubList_17", subCategories: ["שכר עבודה 1", "שכר עבודה 2", "שכר עבודה 3", "שכר עבודה 4", "שכר - כללי"] },
  { mainCategory: "קצבאות", namedRange: "CategorySubList_18", subCategories: ["קצבת ילדים", "קצבת נכות", "סיוע בשכר דירה", "קצבת זיקנה", "קצבאות - כללי"] },
  { mainCategory: "הכנסות שונות", namedRange: "CategorySubList_19", subCategories: ["קבלת מזונות", "הכנסה מנכס", "עזרה מההורים", "הכנסות שונות - כללי"] },
];

const fallbackRules: MappingRule[] = [
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

  for (const [index, file] of files.entries()) {
    await writeFile(
      path.join(uploadsDir, `${index + 1}-${file.kind}-${sanitizeFileName(file.fileName)}`),
      file.buffer,
    );
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
  const fileName = `financial-classification-${jobId}.xlsx`;
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
        // A negative credit-card amount is a refund (income); otherwise trust an
        // explicit mapping direction, falling back to expense.
        direction: amount < 0 ? "הכנסה" : (classification.direction ?? "הוצאה"),
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
      mainCategory: "",
      recurrence: "",
      subCategory: "",
    };
  }

  const searchable = `${description} ${note}`.toLocaleLowerCase("he-IL");
  const match = mappings
    .slice()
    .sort((left, right) => right.keyword.length - left.keyword.length)
    .find((mapping) => searchable.includes(mapping.keyword.toLocaleLowerCase("he-IL")));

  if (match) {
    return {
      // Only an explicit mapping direction is a definitive signal. When the rule
      // has none, leave it undefined so the caller can use its own signal
      // (the bank debit/credit column, or the credit-card amount sign).
      direction: match.direction,
      mainCategory: match.mainCategory,
      recurrence: match.recurrence ?? "",
      subCategory: match.subCategory,
    };
  }

  const categoryFallback = classifyBySourceCategory(sourceCategory);
  return {
    direction: undefined as "הוצאה" | "הכנסה" | undefined,
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
    return null;
  }

  return null;
}

async function buildReportWorkbook(transactions: NormalizedTransaction[], summaries: ParsedSheetSummary[]) {
  const sortedTransactions = transactions.slice().sort((left, right) => compareDateText(left.date, right.date));
  const monthCount = Math.max(1, countDistinctMonths(sortedTransactions));
  const workbook = new ExcelJS.Workbook();
  workbook.creator = "Finance Classifier";
  workbook.created = new Date();
  workbook.calcProperties = { fullCalcOnLoad: true };
  workbook.views = [{ activeTab: 0, firstSheet: 0, height: 12000, visibility: "visible", width: 20000, x: 0, y: 0 }];

  appendExcelClassificationSheet(workbook, sortedTransactions, monthCount);
  appendExcelResultSheet(workbook, sortedTransactions, monthCount);
  appendExcelSummarySheet(workbook, summaries);
  appendExcelCategorySheet(workbook);
  appendExcelImportSheet(workbook, sortedTransactions);
  appendExcelChoicesSheet(workbook);
  sanitizeWorkbookForExcel(workbook);

  const buffer = await workbook.xlsx.writeBuffer();
  return Buffer.from(buffer);
}

function getRecurrenceMonthDivisor(recurrence: string, monthCount: number) {
  const normalized = recurrence.trim();

  if (!normalized || normalized === "חודשי/מזדמן") {
    return monthCount;
  }

  if (normalized === "שנתי") {
    return 12;
  }

  if (normalized === "דו-חודשי") {
    return 2;
  }

  if (normalized === "רבעוני") {
    return 3;
  }

  const numericValue = Number.parseFloat(normalized);
  return Number.isFinite(numericValue) && numericValue > 0 ? numericValue : monthCount;
}

// Mirrors the reference template's column-N formula:
//   IF(OR(F=monthly,F="",F=0), D/months,
//      IF(OR(F=yearly,F=bi-monthly,F=quarterly), D/(divisor*occurrences), D/divisor))
// Recurring (yearly/bi-monthly/quarterly) items are divided by the number of
// occurrences as well, so the same recurring obligation appearing on several
// rows is counted once rather than multiplied.
function transactionMonthlyAverage(
  transaction: NormalizedTransaction,
  monthCount: number,
  occurrences: number,
) {
  const recurrence = transaction.recurrence.trim();

  if (!recurrence || recurrence === "חודשי/מזדמן") {
    return monthCount > 0 ? transaction.amount / monthCount : 0;
  }

  if (recurrence === "שנתי") {
    return occurrences > 0 ? transaction.amount / (12 * occurrences) : 0;
  }

  if (recurrence === "דו-חודשי") {
    return occurrences > 0 ? transaction.amount / (2 * occurrences) : 0;
  }

  if (recurrence === "רבעוני") {
    return occurrences > 0 ? transaction.amount / (3 * occurrences) : 0;
  }

  const customDivisor = Number.parseFloat(recurrence);
  return Number.isFinite(customDivisor) && customDivisor > 0 ? transaction.amount / customDivisor : 0;
}

function computeSubCategoryAverages(transactions: NormalizedTransaction[], monthCount: number) {
  const averages = new Map<string, number>();

  for (const transaction of transactions) {
    if (!transaction.subCategory) {
      continue;
    }

    const occurrences = countDuplicateTransaction(transactions, transaction);
    averages.set(
      transaction.subCategory,
      (averages.get(transaction.subCategory) ?? 0) +
        transactionMonthlyAverage(transaction, monthCount, occurrences),
    );
  }

  return averages;
}

function sumCategoryAverages(categories: string[], averages: Map<string, number>) {
  return categories.reduce((sum, category) => sum + (averages.get(category) ?? 0), 0);
}

function countDuplicateTransaction(transactions: NormalizedTransaction[], target: NormalizedTransaction) {
  const duplicateCount = transactions.filter(
    (transaction) => transaction.description === target.description && transaction.subCategory === target.subCategory,
  ).length;
  return Math.max(1, duplicateCount);
}

function appendExcelSummarySheet(workbook: ExcelJS.Workbook, summaries: ParsedSheetSummary[]) {
  const sheet = workbook.addWorksheet("סיכום קבצים", {
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

function appendExcelClassificationSheet(
  workbook: ExcelJS.Workbook,
  transactions: NormalizedTransaction[],
  monthCount: number,
) {
  const sheet = workbook.addWorksheet("סיווג תנועות", {
    views: [{ activeCell: "G9", rightToLeft: true, state: "frozen", ySplit: 8 }],
  });
  const lastRow = transactions.length + 8;
  const averages = computeSubCategoryAverages(transactions, monthCount);
  const totalExpense = sumCategoryAverages(getExpenseSubCategories(), averages);
  const totalIncome = sumCategoryAverages(getIncomeSubCategories(), averages);

  sheet.addRows([
    ["שם המשפחה", "ישראלי"],
    ["תקופת השיקוף", monthCount],
    ["ממוצע הוצאות בחודש", { formula: "'תוצאות השיקוף'!B1", result: totalExpense }],
    ["ממוצע הכנסות בחודש", { formula: "'תוצאות השיקוף'!E1", result: totalIncome }],
    ["מאזן חודשי", { formula: "B4-B3", result: totalIncome - totalExpense }],
    ["שימו לב! ניתן לבחור ולתקן סעיף ראשי, שם סעיף והוצאה/הכנסה בעזרת הרשימות הנפתחות."],
    ["תנועות מתוך דפי בנק וכרטיסי אשראי"],
    classificationHeaders,
    ...transactions.map((transaction, index) => {
      const rowNumber = index + 9;
      const occurrences = countDuplicateTransaction(transactions, transaction);
      return [
      transaction.sourceName,
      transaction.date,
      transaction.description,
      transaction.amount,
      transaction.direction,
      transaction.recurrence,
      transaction.mainCategory,
      transaction.subCategory,
      transaction.note,
      transaction.originalAmount,
      transaction.chargeCurrency,
      transaction.cardOrAccount,
      "",
      {
        formula: `IFERROR(IF(OR(F${rowNumber}=$P$1,F${rowNumber}="",F${rowNumber}=0),D${rowNumber}/$B$2,IF(OR(F${rowNumber}=$P$2,F${rowNumber}=$P$3,F${rowNumber}=$P$4),D${rowNumber}/(O${rowNumber}*P${rowNumber}),D${rowNumber}/O${rowNumber})),0)`,
        result: transactionMonthlyAverage(transaction, monthCount, occurrences),
      },
      {
        formula: `IFERROR(VLOOKUP(F${rowNumber},$P$1:$Q$4,2,0),F${rowNumber})`,
        result: getRecurrenceMonthDivisor(transaction.recurrence, monthCount),
      },
      {
        formula: `IF(C${rowNumber}="",1,COUNTIFS($C$9:$C$1499,C${rowNumber},$H$9:$H$1499,H${rowNumber}))`,
        result: occurrences,
      },
      ];
    }),
  ]);

  // Recurrence -> month-divisor lookup table used by columns N and O (matches the
  // reference template's $P$1:$Q$4). Lives above the data area in the unused P/Q columns.
  ([
    ["P1", "חודשי/מזדמן"],
    ["P2", "שנתי"],
    ["P3", "דו-חודשי"],
    ["P4", "רבעוני"],
  ] as const).forEach(([address, label]) => {
    sheet.getCell(address).value = label;
  });
  sheet.getCell("Q1").value = { formula: "B2", result: monthCount };
  sheet.getCell("Q2").value = 12;
  sheet.getCell("Q3").value = 2;
  sheet.getCell("Q4").value = 3;

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
      formulae: ['"הוצאה,הכנסה"'],
      showErrorMessage: true,
      type: "list",
    };
    sheet.getCell(`F${rowNumber}`).dataValidation = {
      allowBlank: true,
      formulae: ["RecurrenceList"],
      type: "list",
    };
    sheet.getCell(`G${rowNumber}`).dataValidation = {
      allowBlank: true,
      formulae: [`IF($E${rowNumber}="הכנסה",IncomeMainCategoryList,ExpenseMainCategoryList)`],
      prompt: "בחרו סעיף ראשי כדי לצמצם את רשימת שמות הסעיף בעמודה הבאה.",
      promptTitle: "בחירת סעיף ראשי",
      showInputMessage: true,
      showErrorMessage: true,
      type: "list",
    };
    sheet.getCell(`H${rowNumber}`).dataValidation = {
      allowBlank: true,
      formulae: [
        `OFFSET('רשימת קטגוריות'!$A$2,0,MATCH($G${rowNumber},'רשימת קטגוריות'!$A$1:$S$1,0)-1,COUNTA(OFFSET('רשימת קטגוריות'!$A$2,0,MATCH($G${rowNumber},'רשימת קטגוריות'!$A$1:$S$1,0)-1,100,1)),1)`,
      ],
      prompt: "הרשימה כאן משתנה לפי הסעיף הראשי שנבחר בעמודה הקודמת.",
      promptTitle: "בחירת שם סעיף",
      showInputMessage: true,
      showErrorMessage: true,
      type: "list",
    };
    sheet.getCell(`G${rowNumber}`).fill = { fgColor: { argb: "FFFFF2CC" }, pattern: "solid", type: "pattern" };
    sheet.getCell(`H${rowNumber}`).fill = { fgColor: { argb: "FFFFF2CC" }, pattern: "solid", type: "pattern" };
  }
}

function appendExcelResultSheet(
  workbook: ExcelJS.Workbook,
  transactions: NormalizedTransaction[],
  monthCount: number,
) {
  const sheet = workbook.addWorksheet("תוצאות השיקוף", {
    views: [{ rightToLeft: true }],
  });
  const expenseSubCategories = getExpenseSubCategories();
  const incomeSubCategories = getIncomeSubCategories();
  const maxRows = Math.max(expenseSubCategories.length, incomeSubCategories.length);
  const averages = computeSubCategoryAverages(transactions, monthCount);
  const totalExpense = sumCategoryAverages(expenseSubCategories, averages);
  const totalIncome = sumCategoryAverages(incomeSubCategories, averages);

  // The three rows below are written first, so the per-category data starts on row 4.
  const firstDataRow = 4;
  const lastDataRow = maxRows + firstDataRow - 1;

  sheet.addRows([
    ["סה\"כ הוצאות", { formula: `SUM(B${firstDataRow}:B${lastDataRow})`, result: totalExpense }, "", "סה\"כ הכנסות", { formula: `SUM(E${firstDataRow}:E${lastDataRow})`, result: totalIncome }, "", "הפרש", { formula: "E1-B1", result: totalIncome - totalExpense }],
    [],
    ["קטגוריות הוצאה", "ממוצע חודשי", "", "קטגוריות הכנסה", "ממוצע חודשי"],
  ]);

  for (let index = 0; index < maxRows; index += 1) {
    const rowNumber = index + firstDataRow;
    const expenseCategory = expenseSubCategories[index] ?? "";
    const incomeCategory = incomeSubCategories[index] ?? "";
    sheet.addRow([
      expenseCategory,
      expenseCategory
        ? { formula: `SUMIF('סיווג תנועות'!$H$9:$H$1499,A${rowNumber},'סיווג תנועות'!$N$9:$N$1499)`, result: averages.get(expenseCategory) ?? 0 }
        : "",
      "",
      incomeCategory,
      incomeCategory
        ? { formula: `SUMIF('סיווג תנועות'!$H$9:$H$1499,D${rowNumber},'סיווג תנועות'!$N$9:$N$1499)`, result: averages.get(incomeCategory) ?? 0 }
        : "",
    ]);
  }

  sheet.columns = [{ width: 34 }, { width: 16 }, { width: 4 }, { width: 34 }, { width: 16 }, { width: 4 }, { width: 12 }, { width: 16 }];
  [1, 3].forEach((rowNumber) => styleHeaderRow(sheet.getRow(rowNumber)));
  ["B1", "E1", "H1"].forEach((cellAddress) => {
    sheet.getCell(cellAddress).numFmt = '[$₪-40D]#,##0;[Red]-[$₪-40D]#,##0;[$₪-40D]-';
    sheet.getCell(cellAddress).font = { bold: true };
  });
  for (let rowNumber = firstDataRow; rowNumber <= lastDataRow; rowNumber += 1) {
    sheet.getCell(`B${rowNumber}`).numFmt = '[$₪-40D]#,##0.00;[Red]-[$₪-40D]#,##0.00;[$₪-40D]-';
    sheet.getCell(`E${rowNumber}`).numFmt = '[$₪-40D]#,##0.00;[Red]-[$₪-40D]#,##0.00;[$₪-40D]-';
  }
}

function appendExcelCategorySheet(workbook: ExcelJS.Workbook) {
  const sheet = workbook.addWorksheet("רשימת קטגוריות", {
    views: [{ rightToLeft: true }],
  });
  const rows = buildCategorySheetRows();
  sheet.addRows(rows);
  sheet.columns = rows[0].map(() => ({ width: 22 }));
  styleHeaderRow(sheet.getRow(1));
  sheet.getRow(1).fill = { fgColor: { argb: "FFF28C28" }, pattern: "solid", type: "pattern" };
  sheet.getRow(1).font = { bold: true, color: { argb: "FFFFFFFF" } };
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
  const expenseMainCategories = getExpenseMainCategories();
  const incomeMainCategories = getIncomeMainCategories();
  const directions = ["הוצאה", "הכנסה"];
  const recurrences = ["", "חודשי/מזדמן", "שנתי", "דו-חודשי", "רבעוני", "חלוקת הסכום ב X"];
  const maxRows = Math.max(
    directions.length,
    recurrences.length,
    expenseMainCategories.length,
    incomeMainCategories.length,
  );

  sheet.addRow([
    "הוצאה/הכנסה",
    "מחזוריות",
    "סעיפי הוצאה",
    "סעיפי הכנסה",
  ]);

  for (let index = 0; index < maxRows; index += 1) {
    sheet.addRow([
      directions[index] ?? "",
      recurrences[index] ?? "",
      expenseMainCategories[index] ?? "",
      incomeMainCategories[index] ?? "",
    ]);
  }

  workbook.definedNames.add("'בחירות'!$C$2:$C$" + (expenseMainCategories.length + 1), "ExpenseMainCategoryList");
  workbook.definedNames.add("'בחירות'!$D$2:$D$" + (incomeMainCategories.length + 1), "IncomeMainCategoryList");
  workbook.definedNames.add("'בחירות'!$B$2:$B$" + (recurrences.length + 1), "RecurrenceList");
}

function getExpenseMainCategories() {
  return getCategoryGroups(0, 16).map((group) => group.mainCategory);
}

function getIncomeMainCategories() {
  return getCategoryGroups(16, 19).map((group) => group.mainCategory);
}

function getExpenseSubCategories() {
  return getCategoryGroups(0, 16).flatMap((group) => group.subCategories);
}

function getIncomeSubCategories() {
  return getCategoryGroups(16, 19).flatMap((group) => group.subCategories);
}

function getCategoryGroups(startColumn: number, endColumn: number): CategoryGroup[] {
  return categoryCatalog.slice(startColumn, endColumn);
}

function buildCategorySheetRows() {
  const maxSubCategoryCount = Math.max(...categoryCatalog.map((group) => group.subCategories.length));
  return [
    categoryCatalog.map((group) => group.mainCategory),
    ...Array.from({ length: maxSubCategoryCount }, (_, index) =>
      categoryCatalog.map((group) => group.subCategories[index] ?? ""),
    ),
  ];
}

function sanitizeWorkbookForExcel(workbook: ExcelJS.Workbook) {
  workbook.worksheets.forEach((sheet) => {
    sheet.pageSetup = {
      fitToHeight: 0,
      fitToWidth: 0,
      horizontalDpi: 600,
      orientation: "portrait",
      verticalDpi: 600,
    };
  });
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
