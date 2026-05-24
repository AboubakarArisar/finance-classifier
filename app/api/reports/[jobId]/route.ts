import { readReport } from "@/lib/finance-analyzer";

export const runtime = "nodejs";

export async function GET(_request: Request, context: { params: Promise<{ jobId: string }> }) {
  const { jobId } = await context.params;
  const report = await readReport(jobId);

  if (!report) {
    return Response.json({ error: "הדוח לא נמצא או שפג תוקפו." }, { status: 404 });
  }

  return new Response(report, {
    headers: {
      "Content-Disposition": `attachment; filename="classified-transactions-${jobId}.xlsx"`,
      "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    },
  });
}
