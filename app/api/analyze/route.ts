import { analyzeFinancialStatements, isExcelFileName, type UploadKind } from "@/lib/finance-analyzer";

export const runtime = "nodejs";

const maxFiles = 120;

// The analyzer auto-detects credit vs. bank from each sheet's headers, so the
// kind we pass here only labels the saved upload. Guess it from the file name
// for nicer diagnostics, defaulting to credit.
function guessKind(fileName: string): UploadKind {
  return /בנק|עו"?ש|עובר ושב|חשבון/.test(fileName) ? "bank" : "credit";
}

export async function POST(request: Request) {
  try {
    const formData = await request.formData();
    // Primary field is a single combined "files" list; fall back to the legacy
    // creditFile/bankFile fields so older clients keep working.
    const values = [...formData.getAll("files"), ...formData.getAll("creditFile"), ...formData.getAll("bankFile")].filter(
      (value): value is File => value instanceof File && value.size > 0,
    );

    if (values.length === 0) {
      throw new Error("יש להעלות לפחות קובץ Excel אחד.");
    }

    if (values.length > maxFiles) {
      throw new Error(`ניתן להעלות עד ${maxFiles} קבצי Excel.`);
    }

    const uploadedFiles = await Promise.all(
      values.map(async (value) => {
        if (!isExcelFileName(value.name)) {
          throw new Error(`הקובץ "${value.name}" אינו קובץ Excel תקין (XLS, XLSX או XLSM).`);
        }

        return {
          buffer: Buffer.from(await value.arrayBuffer()),
          fileName: value.name,
          kind: guessKind(value.name),
        };
      }),
    );

    return Response.json(await analyzeFinancialStatements(uploadedFiles));
  } catch (error) {
    const message = error instanceof Error ? error.message : "אירעה שגיאה בעיבוד הקבצים.";
    return Response.json({ error: message }, { status: 400 });
  }
}
