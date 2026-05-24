import { analyzeFinancialStatements, isExcelFileName, type UploadKind } from "@/lib/finance-analyzer";

export const runtime = "nodejs";

const requiredFiles: { field: string; kind: UploadKind; label: string }[] = [
  { field: "creditFile", kind: "credit", label: "קובץ כרטיסי אשראי" },
  { field: "bankFile", kind: "bank", label: "קובץ חשבון בנק" },
];

export async function POST(request: Request) {
  try {
    const formData = await request.formData();
    const uploadedFiles = await Promise.all(
      requiredFiles.map(async (requiredFile) => {
        const value = formData.get(requiredFile.field);

        if (!(value instanceof File) || value.size === 0) {
          throw new Error(`חסר ${requiredFile.label}.`);
        }

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

    return Response.json(await analyzeFinancialStatements(uploadedFiles));
  } catch (error) {
    const message = error instanceof Error ? error.message : "אירעה שגיאה בעיבוד הקבצים.";
    return Response.json({ error: message }, { status: 400 });
  }
}
