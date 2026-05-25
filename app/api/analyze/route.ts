import { analyzeFinancialStatements, isExcelFileName, type UploadKind } from "@/lib/finance-analyzer";

export const runtime = "nodejs";

const requiredFiles: { field: string; kind: UploadKind; label: string }[] = [
  { field: "creditFile", kind: "credit", label: "קובץ כרטיסי אשראי" },
  { field: "bankFile", kind: "bank", label: "קובץ חשבון בנק" },
];
const maxFilesPerGroup = 6;

export async function POST(request: Request) {
  try {
    const formData = await request.formData();
    const uploadedFiles = await Promise.all(
      requiredFiles.map(async (requiredFile) => {
        const values = formData
          .getAll(requiredFile.field)
          .filter((value): value is File => value instanceof File && value.size > 0);

        if (values.length === 0) {
          throw new Error(`חסר ${requiredFile.label}. יש להעלות בין 1 ל-6 קבצי Excel.`);
        }

        if (values.length > maxFilesPerGroup) {
          throw new Error(`${requiredFile.label} מוגבל ל-6 קבצי Excel לכל היותר.`);
        }

        return Promise.all(
          values.map(async (value) => {
            if (!isExcelFileName(value.name)) {
              throw new Error(`${requiredFile.label} חייב להיות קובץ Excel מסוג XLS, XLSX או XLSM.`);
            }

            return {
              buffer: Buffer.from(await value.arrayBuffer()),
              fileName: value.name,
              kind: requiredFile.kind,
            };
          }),
        );
      }),
    );

    return Response.json(await analyzeFinancialStatements(uploadedFiles.flat()));
  } catch (error) {
    const message = error instanceof Error ? error.message : "אירעה שגיאה בעיבוד הקבצים.";
    return Response.json({ error: message }, { status: 400 });
  }
}
