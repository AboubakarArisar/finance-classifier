import { randomUUID } from "crypto";
import { existsSync, readFileSync } from "fs";
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

// Special "Not for classification" label. A row tagged with this is shown to the
// user but kept out of every expense/income total — see the catalog entry below.
const notForClassificationLabel = "לא לסיווג";

// The client asked that NO category be filled automatically: every row is imported
// with empty סעיף ראשי / שם סעיף, and the end user classifies each one by hand
// using the dropdowns. Only the direction (הוצאה/הכנסה) is still derived
// automatically. Set this back to true to re-enable the rule-based engine.
const autoClassifyCategories = false;

// === Generated-report visual theme ==========================================
// One self-contained palette drives the look of every sheet. It is deliberately
// a cool teal/slate scheme with a coral accent so the output reads as its own
// product rather than a copy of the reference template's warm beige/orange
// styling. IMPORTANT: nothing here is referenced by a formula, data validation,
// named range or total — these values only set fills, fonts and borders, so
// restyling can never change a number, a dropdown or a column in the report.
const reportTheme = {
  headerFill: "FF124559", // deep teal — column / section header rows
  headerText: "FFFFFFFF", // white header lettering
  subHeaderFill: "FF3A7CA5", // lighter teal — the category-list helper header
  titleText: "FF124559", // teal for the big labels in the KPI block
  bandFill: "FFEFF4F7", // pale slate — zebra banding on data rows
  editableFill: "FFDCEFEA", // soft mint — the user-editable dropdown cells (G/H)
  kpiFill: "FFFBEAE3", // pale coral wash behind the KPI / total numbers
  kpiText: "FFB5532A", // burnt-coral lettering for KPI / total numbers
  noteText: "FF6B7B83", // muted slate for the helper note line
  border: "FFB7C5CC", // soft slate cell borders
};

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
  // "Not for classification" — a special main category the user can pick for any
  // row they want excluded from the expense/income totals (e.g. de-duplicating a
  // credit-card purchase that also appears as the monthly card-settlement charge
  // on the bank statement). It has NO sub-categories on purpose: once a row's
  // main category is "לא לסיווג" the שם סעיף dropdown is empty and nothing can be
  // chosen there, matching the reference system. It is also kept out of the
  // expense/income sub-category lists, so the result-sheet SUMIFs never sum it.
  { mainCategory: notForClassificationLabel, namedRange: "CategorySubList_20", subCategories: [] },
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
    // Try each known statement layout in turn. The first one that yields rows
    // wins, so the order only matters when a sheet could plausibly match more
    // than one (it can't: each layout is keyed on headers unique to it).
    let parsed = parseCreditSheet(rows, sheetName, file, mappings, () => getNextSourceName("credit"));

    if (parsed.transactions.length === 0) {
      parsed = parseIsracardSheet(rows, sheetName, file, mappings, () => getNextSourceName("credit"));
    }

    if (parsed.transactions.length === 0) {
      parsed = parseBankSheet(rows, sheetName, file, mappings, () => getNextSourceName("bank"));
    }

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
      // A negative credit-card amount is a refund (income); otherwise it's an expense.
      const isExpense = amount >= 0;
      const classification = classifyTransaction(description, note, isExpense, mappings, sourceCategory);

      return {
        amount: Math.abs(amount),
        cardOrAccount: readRowText(row, indexes.card),
        chargeCurrency: readRowText(row, indexes.chargeCurrency) || "₪",
        date: formatDateText(readRowText(row, indexes.date)),
        description,
        direction: isExpense ? "הוצאה" : "הכנסה",
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
      // The debit/credit column is the authoritative direction signal for a bank line.
      const isExpense = debit !== null;
      // A bank line that is the aggregate credit-card settlement (e.g. "כאל",
      // "מקס איט פיננסי") is tagged "לא לסיווג": its individual transactions are
      // already counted from the card statement, so counting it here too would
      // double-count. Tagging it (rather than leaving it blank) makes the
      // exclusion visible to the user and mirrors the reference tool.
      const classification =
        autoClassifyCategories && isCreditCardSettlement(description)
          ? { mainCategory: notForClassificationLabel, recurrence: "", subCategory: notForClassificationLabel }
          : classifyTransaction(description, details, isExpense, mappings);

      return {
        amount: Math.abs(amount),
        cardOrAccount: accountNumber,
        chargeCurrency: "",
        date: formatDateText(readRowText(row, indexes.date)),
        description,
        direction: isExpense ? "הוצאה" : "הכנסה",
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

// Isracard / "פירוט עסקאות" credit-card layout (e.g. Mastercard Gold). Unlike the
// Max layout handled by parseCreditSheet, its columns are "תאריך רכישה" / "שם בית עסק"
// / "סכום עסקה" / "סכום חיוב" / "פירוט נוסף", and the data block ends with a
// "סה״כ לחיוב החודש" total row that must not be counted as a transaction.
function parseIsracardSheet(
  rows: unknown[][],
  sheetName: string,
  file: UploadedWorkbook,
  mappings: MappingRule[],
  getSourceName: () => string,
): ParsedSheet {
  const headerIndex = findHeaderRowIndex(rows, ["שם בית עסק", "סכום חיוב"]);

  if (headerIndex < 0) {
    return emptyParsedSheet(file, sheetName, "");
  }

  const cardNumber = extractCardNumber(rows);
  const headers = rows[headerIndex].map((cell) => normalizeHeader(String(cell ?? "")));
  const indexes = {
    amount: findHeaderIndex(headers, ["סכום חיוב"]),
    chargeCurrency: findHeaderIndex(headers, ["מטבע חיוב"]),
    date: findHeaderIndex(headers, ["תאריך רכישה", "תאריך עסקה"]),
    description: findHeaderIndex(headers, ["שם בית עסק", "שם בית העסק"]),
    note: findHeaderIndex(headers, ["פירוט נוסף", "הערות"]),
    originalAmount: findHeaderIndex(headers, ["סכום עסקה", "סכום עסקה מקורי"]),
  };

  const transactions = rows
    .slice(headerIndex + 1)
    .map<NormalizedTransaction | null>((row) => {
      const description = readRowText(row, indexes.description);
      // Some lines carry only the original amount (e.g. a fully-discounted card
      // fee charges 0); fall back to it so the row is still counted.
      const amount = readRowNumber(row, indexes.amount) ?? readRowNumber(row, indexes.originalAmount);

      if (!description || amount === null || isSummaryLine(description)) {
        return null;
      }

      const note = readRowText(row, indexes.note);
      const isExpense = amount >= 0;
      const classification = classifyTransaction(description, note, isExpense, mappings);

      return {
        amount: Math.abs(amount),
        cardOrAccount: cardNumber,
        chargeCurrency: readRowText(row, indexes.chargeCurrency) || "₪",
        date: formatDateText(readRowText(row, indexes.date)),
        description,
        direction: isExpense ? "הוצאה" : "הכנסה",
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

type PaamonimTuple = [string, string, string, number, number];

type PaamonimEntry = { category: string; expenseFlag: number; subCategory: string };

type PaamonimIndex = {
  exact: Map<string, PaamonimEntry>;
  prefixByLength: Map<number, Map<string, PaamonimEntry>>;
  prefixLengthsDesc: number[];
};

const paamonimMappingPath = path.join(/*turbopackIgnore: true*/ process.cwd(), "data", "paamonim-mappings.json");
let paamonimIndexCache: PaamonimIndex | null = null;

// The reference database stores a handful of sub-categories with spellings that
// differ from this template's canonical list (categoryCatalog). Left unmapped,
// the report's SUMIF averages would silently miss those rows, so we reconcile to
// the catalog spelling — which is also what the reference's own output displays.
const subCategoryAliases: Record<string, string> = {
  "טיפולי שיניים / אורתודנט": "טיפולי שיניים / אורטודנט",
  "מיסי יישוב / ועד בית": "מיסי ישוב / ועד בית",
  "טלוויזיה ואינטרנט (ספק ותשתית)": "טלויזיה ואינטרנט (ספק ותשתית)",
  "עזרה ממשפחה": "עזרה למשפחה",
  "אירועי שמחות במשפחה": "ארועי שמחות במשפחה",
  "הכנסה מנכס או פיננסי": "הכנסה מנכס",
};

// Same normalization the reference classifier uses: trim, lower-case, collapse
// internal whitespace. Both the database keys and the lookup query go through it.
function normalizeClassifierKey(value: string) {
  return value ? value.trim().toLowerCase().replace(/\s+/g, " ") : "";
}

// Rebuilds the reference tool's RuleBasedClassifier from data/paamonim-mappings.json:
// an exact-key map plus a longest-prefix map (bucketed by key length). Built once
// and cached for the lifetime of the process.
function getPaamonimIndex(): PaamonimIndex {
  if (paamonimIndexCache) {
    return paamonimIndexCache;
  }

  const exact = new Map<string, PaamonimEntry>();
  const prefixByLength = new Map<number, Map<string, PaamonimEntry>>();

  if (existsSync(paamonimMappingPath)) {
    const tuples = JSON.parse(readFileSync(paamonimMappingPath, "utf8")) as PaamonimTuple[];

    for (const [cbValue, category, subCategory, expenseFlag, prefixFlag] of tuples) {
      const key = normalizeClassifierKey(cbValue);

      if (!key) {
        continue;
      }

      const entry: PaamonimEntry = {
        category,
        expenseFlag,
        subCategory: subCategoryAliases[subCategory] ?? subCategory,
      };

      if (prefixFlag === 1) {
        // The reference drops single-character prefixes to avoid matching everything.
        if (key.length <= 1) {
          continue;
        }

        let bucket = prefixByLength.get(key.length);
        if (!bucket) {
          bucket = new Map();
          prefixByLength.set(key.length, bucket);
        }
        if (!bucket.has(key)) {
          bucket.set(key, entry);
        }
      } else if (!exact.has(key)) {
        exact.set(key, entry);
      }
    }
  }

  paamonimIndexCache = {
    exact,
    prefixByLength,
    prefixLengthsDesc: Array.from(prefixByLength.keys()).sort((left, right) => right - left),
  };
  return paamonimIndexCache;
}

// A rule applies only when its direction matches the transaction's. expenseFlag 2
// means the rule is direction-agnostic (the reference's "isExpense === undefined").
function matchesDirection(expenseFlag: number, isExpense: boolean) {
  return expenseFlag === 2 || expenseFlag === (isExpense ? 1 : 0);
}

function classifyByPaamonim(description: string, isExpense: boolean) {
  const key = normalizeClassifierKey(description);

  if (!key) {
    return null;
  }

  const index = getPaamonimIndex();
  const exact = index.exact.get(key);

  if (exact && matchesDirection(exact.expenseFlag, isExpense)) {
    return { mainCategory: exact.category, subCategory: exact.subCategory };
  }

  // Longest prefix wins, exactly like findPrefixMapping in the reference.
  for (const length of index.prefixLengthsDesc) {
    if (key.length < length) {
      continue;
    }

    const hit = index.prefixByLength.get(length)?.get(key.slice(0, length));

    if (hit && matchesDirection(hit.expenseFlag, isExpense)) {
      return { mainCategory: hit.category, subCategory: hit.subCategory };
    }
  }

  return null;
}

// Credit-card companies as they appear as a single settlement line on a bank
// statement. Anchored to the start so it won't catch unrelated merchants.
const creditCardSettlementPattern =
  /^(כא"?ל|מקס איט|מקס פיננס|ויזה כא"?ל|ישראכרט|לאומי קארד|לאומי-קארד|אמריקן אקספרס|דיינרס|דירקט|כרטיסי אשראי)/;

function isCreditCardSettlement(description: string) {
  return creditCardSettlementPattern.test(normalizeClassifierKey(description));
}

function classifyTransaction(
  description: string,
  note: string,
  isExpense: boolean,
  mappings: MappingRule[],
  sourceCategory = "",
) {
  // Auto-classification disabled by client request: leave the category columns
  // empty so the user fills them in manually. Direction is set by the caller.
  if (!autoClassifyCategories) {
    return { mainCategory: "", recurrence: "", subCategory: "" };
  }

  // Multi-payment installments are left unclassified, mirroring the reference tool,
  // which sets these aside ("שקול רישום כחוב") rather than counting them as a
  // recurring monthly expense.
  if (/תשלום\s+\d+\s+מתוך\s+\d+/.test(note)) {
    return { mainCategory: "", recurrence: "", subCategory: "" };
  }

  // 1) Paamonim rule database — the reference engine, by far the most comprehensive.
  const fromDatabase = classifyByPaamonim(description, isExpense);
  if (fromDatabase) {
    return { mainCategory: fromDatabase.mainCategory, recurrence: "", subCategory: fromDatabase.subCategory };
  }

  // 2) User-supplied keyword overrides (data/category-mapping.xlsx + fallbackRules).
  const searchable = `${description} ${note}`.toLocaleLowerCase("he-IL");
  const match = mappings
    .slice()
    .sort((left, right) => right.keyword.length - left.keyword.length)
    .find((mapping) => searchable.includes(mapping.keyword.toLocaleLowerCase("he-IL")));

  if (match) {
    return { mainCategory: match.mainCategory, recurrence: match.recurrence ?? "", subCategory: match.subCategory };
  }

  // 3) Last resort: the statement's own category column.
  const categoryFallback = classifyBySourceCategory(sourceCategory);
  return {
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

  // Columns F, M, N, O and P are hidden (not deleted — the monthly-average
  // formula and the result-sheet SUMIFs still read them). F (מחזוריות) is the
  // recurrence selector, M is an empty spacer, and N/O/P are the three
  // "לשימוש פנימי" helper columns. Hiding keeps the user-facing sheet clean while
  // every formula keeps working; the user can unhide any of them in Excel.
  sheet.columns = [
    { width: 14 }, // A מקור
    { width: 12 }, // B תאריך
    { width: 34 }, // C תיאור / שם בית העסק
    { width: 14 }, // D סכום
    { width: 16 }, // E הוצאה/הכנסה
    { width: 18, hidden: true }, // F מחזוריות
    { width: 24 }, // G סעיף ראשי
    { width: 28 }, // H שם סעיף
    { width: 34 }, // I הערות
    { width: 14 }, // J סכום עסקה
    { width: 12 }, // K מטבע לחיוב
    { width: 16 }, // L מספר כרטיס/בנק
    { width: 4, hidden: true }, // M spacer
    { width: 20, hidden: true }, // N לממוצע חודשי (לשימוש פנימי)
    { width: 20, hidden: true }, // O מס' חודשים (לשימוש פנימי)
    { width: 20, hidden: true }, // P מס' מופעים (לשימוש פנימי)
  ];
  sheet.autoFilter = `A8:P${lastRow}`;
  styleHeaderRow(sheet.getRow(8));
  // KPI block (rows 1-5): teal labels on the left, coral "value chips" on the
  // right so the family's monthly summary reads as a little dashboard card.
  ["B3", "B4", "B5"].forEach((cellAddress) => {
    sheet.getCell(cellAddress).numFmt = '[$₪-40D]#,##0;[Red]-[$₪-40D]#,##0;[$₪-40D]-';
    sheet.getCell(cellAddress).font = { bold: true, color: { argb: reportTheme.kpiText } };
    sheet.getCell(cellAddress).fill = { fgColor: { argb: reportTheme.kpiFill }, pattern: "solid", type: "pattern" };
  });
  ["A1", "A2", "A3", "A4", "A5"].forEach((cellAddress) => {
    sheet.getCell(cellAddress).font = { bold: true, color: { argb: reportTheme.titleText } };
  });
  // Row 6 is the "you can edit the dropdowns" note; row 7 is the table banner.
  sheet.getCell("A6").font = { italic: true, color: { argb: reportTheme.noteText } };
  sheet.getCell("A7").font = { bold: true, size: 12, color: { argb: reportTheme.titleText } };

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
        `OFFSET('רשימת קטגוריות'!$A$2,0,MATCH($G${rowNumber},'רשימת קטגוריות'!$A$1:$T$1,0)-1,COUNTA(OFFSET('רשימת קטגוריות'!$A$2,0,MATCH($G${rowNumber},'רשימת קטגוריות'!$A$1:$T$1,0)-1,100,1)),1)`,
      ],
      prompt: "הרשימה כאן משתנה לפי הסעיף הראשי שנבחר בעמודה הקודמת.",
      promptTitle: "בחירת שם סעיף",
      showInputMessage: true,
      showErrorMessage: true,
      type: "list",
    };
    // Zebra-band the visible columns (A-L) on alternating rows for readability,
    // then paint the two editable dropdown cells (G/H) mint on top so they stay
    // obvious. Hidden helper columns (M-P) are left untouched.
    if (rowNumber % 2 === 1) {
      for (let col = 1; col <= 12; col += 1) {
        sheet.getCell(rowNumber, col).fill = {
          fgColor: { argb: reportTheme.bandFill },
          pattern: "solid",
          type: "pattern",
        };
      }
    }
    sheet.getCell(`G${rowNumber}`).fill = { fgColor: { argb: reportTheme.editableFill }, pattern: "solid", type: "pattern" };
    sheet.getCell(`H${rowNumber}`).fill = { fgColor: { argb: reportTheme.editableFill }, pattern: "solid", type: "pattern" };
  }

  // When a row's main category is "לא לסיווג" it has no sub-category and is kept
  // out of the totals, so both G and H for that row are shown white instead of the
  // mint editable tint — a visual cue that the row carries no classification. This
  // is live conditional formatting that follows the dropdown; it changes no value,
  // formula or total. ($G locks the column, the row is relative so each line tests
  // its own G cell.)
  sheet.addConditionalFormatting({
    ref: `G9:H${lastRow}`,
    rules: [
      {
        type: "expression",
        formulae: [`$G9="${notForClassificationLabel}"`],
        priority: 1,
        style: {
          fill: { type: "pattern", pattern: "solid", bgColor: { argb: "FFFFFFFF" } },
        },
      },
    ],
  });
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
  // The three totals on row 1 get the same coral "value chip" look as the
  // classification KPI block, so the two sheets feel like one product.
  ["B1", "E1", "H1"].forEach((cellAddress) => {
    sheet.getCell(cellAddress).numFmt = '[$₪-40D]#,##0;[Red]-[$₪-40D]#,##0;[$₪-40D]-';
    sheet.getCell(cellAddress).font = { bold: true, color: { argb: reportTheme.kpiText } };
    sheet.getCell(cellAddress).fill = { fgColor: { argb: reportTheme.kpiFill }, pattern: "solid", type: "pattern" };
  });
  for (let rowNumber = firstDataRow; rowNumber <= lastDataRow; rowNumber += 1) {
    sheet.getCell(`B${rowNumber}`).numFmt = '[$₪-40D]#,##0.00;[Red]-[$₪-40D]#,##0.00;[$₪-40D]-';
    sheet.getCell(`E${rowNumber}`).numFmt = '[$₪-40D]#,##0.00;[Red]-[$₪-40D]#,##0.00;[$₪-40D]-';
    // Zebra-band the expense (A/B) and income (D/E) columns on alternating rows.
    if (rowNumber % 2 === 0) {
      [1, 2, 4, 5].forEach((col) => {
        sheet.getCell(rowNumber, col).fill = {
          fgColor: { argb: reportTheme.bandFill },
          pattern: "solid",
          type: "pattern",
        };
      });
    }
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
  // Distinguish this behind-the-scenes category list with the lighter teal shade.
  sheet.getRow(1).fill = { fgColor: { argb: reportTheme.subHeaderFill }, pattern: "solid", type: "pattern" };
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

// "לא לסיווג" is offered in both directions so the user can exclude either an
// expense or an income row, regardless of the הוצאה/הכנסה value in column E.
function getExpenseMainCategories() {
  return [...getCategoryGroups(0, 16).map((group) => group.mainCategory), notForClassificationLabel];
}

function getIncomeMainCategories() {
  return [...getCategoryGroups(16, 19).map((group) => group.mainCategory), notForClassificationLabel];
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
  row.font = { bold: true, color: { argb: reportTheme.headerText } };
  row.fill = { fgColor: { argb: reportTheme.headerFill }, pattern: "solid", type: "pattern" };
  row.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
  row.eachCell((cell) => {
    cell.border = {
      bottom: { color: { argb: reportTheme.border }, style: "thin" },
      left: { color: { argb: reportTheme.border }, style: "thin" },
      right: { color: { argb: reportTheme.border }, style: "thin" },
      top: { color: { argb: reportTheme.border }, style: "thin" },
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
  const fourDigitYear = value.match(/^(\d{1,2})[./-](\d{1,2})[./-](\d{4})$/);

  if (fourDigitYear) {
    const [, day, month, year] = fourDigitYear;
    return `${day.padStart(2, "0")}/${month.padStart(2, "0")}/${year}`;
  }

  // Two-digit years show up in two flavours: Isracard prints day-first dotted
  // dates (22.03.26) while bank exports render real dates as US m/d/yy (1/30/26).
  // Disambiguate by the obvious out-of-range field first, then by separator.
  const twoDigitYear = value.match(/^(\d{1,2})([./-])(\d{1,2})[./-](\d{2})$/);

  if (twoDigitYear) {
    const [, first, separator, second, shortYear] = twoDigitYear;
    const a = Number(first);
    const b = Number(second);
    let day: number;
    let month: number;

    if (a > 12) {
      day = a;
      month = b;
    } else if (b > 12) {
      month = a;
      day = b;
    } else if (separator === "/") {
      month = a;
      day = b;
    } else {
      day = a;
      month = b;
    }

    return `${String(day).padStart(2, "0")}/${String(month).padStart(2, "0")}/${2000 + Number(shortYear)}`;
  }

  return value;
}

function isSummaryLine(description: string) {
  const compact = description.replace(/\s+/g, "");
  return compact.startsWith('סה"כ') || compact.startsWith("סהכ") || compact.startsWith("סךהכל");
}

function extractCardNumber(rows: unknown[][]) {
  // The card label lives in its own cell (e.g. "גולד - מסטרקארד - 1988"); read it
  // per-cell so a neighbouring amount column can't swallow the trailing digits.
  for (const row of rows.slice(0, 6)) {
    for (const cell of row) {
      const text = String(cell ?? "");

      if (/מסטרקארד|ויזה|ישראכרט|אמריקן|דיינרס/.test(text)) {
        const match = text.match(/(\d{3,4})(?!.*\d)/);

        if (match) {
          return match[1];
        }
      }
    }
  }

  return "";
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
