"use client";

import { ChangeEvent, FormEvent, useMemo, useState } from "react";

type UploadKey = "creditFile" | "bankFile";

type AnalyzeResponse = {
  downloadUrl: string;
  jobId: string;
  rowCount: number;
};

const uploadCards: Record<UploadKey, { description: string; label: string; step: string }> = {
  creditFile: {
    description: "קובץ Excel בלבד, עבור ששת החודשים המלאים האחרונים וללא פרטי כרטיס מלאים.",
    label: "פירוט כרטיסי אשראי",
    step: "שלב 1",
  },
  bankFile: {
    description: "קובץ Excel של תנועות חשבון הבנק עבור ששת החודשים האחרונים.",
    label: "פירוט חשבון בנק",
    step: "שלב 2",
  },
};

function isExcelFile(file: File) {
  const name = file.name.toLowerCase();
  return name.endsWith(".xls") || name.endsWith(".xlsx") || name.endsWith(".xlsm");
}

export default function Home() {
  const [files, setFiles] = useState<Record<UploadKey, File | null>>({
    bankFile: null,
    creditFile: null,
  });
  const [error, setError] = useState("");
  const [result, setResult] = useState<AnalyzeResponse | null>(null);
  const [isSubmitting, setIsSubmitting] = useState(false);

  const canSubmit = useMemo(() => Boolean(files.creditFile && files.bankFile), [files]);

  function handleFileChange(key: UploadKey, event: ChangeEvent<HTMLInputElement>) {
    const nextFile = event.target.files?.[0] ?? null;
    setResult(null);

    if (!nextFile) {
      setFiles((current) => ({ ...current, [key]: null }));
      return;
    }

    if (!isExcelFile(nextFile)) {
      event.target.value = "";
      setError("ניתן להעלות קבצי Excel בלבד: XLS, XLSX או XLSM.");
      setFiles((current) => ({ ...current, [key]: null }));
      return;
    }

    setError("");
    setFiles((current) => ({ ...current, [key]: nextFile }));
  }

  async function submitFiles(event: FormEvent<HTMLFormElement>) {
    event.preventDefault();

    if (!files.creditFile || !files.bankFile) {
      setError("יש להעלות גם קובץ כרטיסי אשראי וגם קובץ חשבון בנק.");
      return;
    }

    const formData = new FormData();
    formData.append("creditFile", files.creditFile);
    formData.append("bankFile", files.bankFile);

    setError("");
    setResult(null);
    setIsSubmitting(true);

    try {
      const response = await fetch("/api/analyze", {
        body: formData,
        method: "POST",
      });
      const payload = await response.json();

      if (!response.ok) {
        throw new Error(payload.error ?? "אירעה שגיאה בעיבוד הקבצים.");
      }

      setResult(payload);
    } catch (caughtError) {
      setError(caughtError instanceof Error ? caughtError.message : "אירעה שגיאה בעיבוד הקבצים.");
    } finally {
      setIsSubmitting(false);
    }
  }

  return (
    <main className="min-h-screen bg-[#f6f4ef] text-[#1d2521]" dir="rtl">
      <section className="mx-auto flex min-h-screen w-full max-w-6xl flex-col px-5 py-6 sm:px-8 lg:px-10">
        <header className="flex flex-col gap-4 border-b border-[#d7d6cc] pb-5 sm:flex-row sm:items-end sm:justify-between">
          <div className="max-w-3xl">
            <p className="text-sm font-semibold text-[#477061]">Benny Finance Classifier</p>
            <h1 className="mt-2 text-3xl font-semibold leading-tight text-[#17211d] sm:text-5xl">
              סיווג פעולות פיננסיות מקבצי Excel
            </h1>
            <p className="mt-4 max-w-2xl text-base leading-7 text-[#59645e]">
              העלו דפי כרטיס אשראי וחשבון בנק, והמערכת תחזיר קובץ Excel מסווג להורדה.
            </p>
          </div>
          <span className="inline-flex h-11 items-center border border-[#aeb7ae] bg-white px-4 text-sm font-semibold text-[#33443d]">
            שמירת קבצים ל-30 יום
          </span>
        </header>

        <div className="grid flex-1 gap-6 py-8 lg:grid-cols-[1fr_0.72fr]">
          <form className="flex flex-col gap-5" onSubmit={submitFiles}>
            <section className="border border-[#d7d6cc] bg-white p-5 shadow-sm">
              <div className="flex items-start justify-between gap-4">
                <div>
                  <p className="text-sm font-semibold text-[#477061]">העלאת קבצים</p>
                  <h2 className="mt-1 text-2xl font-semibold">בחרו את שני קבצי ה-Excel</h2>
                </div>
                <span className="border border-[#d7d6cc] px-3 py-1 text-xs font-semibold text-[#477061]">
                  XLS / XLSX / XLSM
                </span>
              </div>

              <div className="mt-5 grid gap-4 sm:grid-cols-2">
                {(Object.keys(uploadCards) as UploadKey[]).map((key) => (
                  <label
                    className="flex min-h-52 cursor-pointer flex-col justify-between border border-dashed border-[#b6beb4] bg-[#fbfaf6] p-5 transition hover:border-[#477061] hover:bg-white"
                    key={key}
                  >
                    <span>
                      <span className="text-sm font-semibold text-[#477061]">{uploadCards[key].step}</span>
                      <span className="mt-2 block text-xl font-semibold">{uploadCards[key].label}</span>
                      <span className="mt-2 block text-sm leading-6 text-[#626d66]">
                        {uploadCards[key].description}
                      </span>
                    </span>
                    <span className="mt-6 block">
                      <input
                        accept=".xls,.xlsx,.xlsm,application/vnd.ms-excel,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        className="sr-only"
                        name={key}
                        onChange={(event) => handleFileChange(key, event)}
                        type="file"
                      />
                      <span className="inline-flex min-h-11 w-full items-center justify-center border border-[#1d2521] px-4 text-center text-sm font-semibold">
                        {files[key]?.name ?? "בחירת קובץ"}
                      </span>
                    </span>
                  </label>
                ))}
              </div>

              {error ? (
                <p className="mt-4 border border-[#e7c7c7] bg-[#fff7f7] px-4 py-3 text-sm font-semibold text-[#9b2f2f]">
                  {error}
                </p>
              ) : null}

              <button
                className="mt-5 h-12 w-full bg-[#1d2521] px-6 text-sm font-semibold text-white transition enabled:hover:bg-[#34423b] disabled:cursor-not-allowed disabled:bg-[#aab3aa] sm:w-auto"
                disabled={!canSubmit || isSubmitting}
                type="submit"
              >
                {isSubmitting ? "מעבד קבצים..." : "התחל סיווג"}
              </button>
            </section>

            <section className="border border-[#d7d6cc] bg-white p-5 shadow-sm">
              <p className="text-sm font-semibold text-[#477061]">תוצאה</p>
              <h2 className="mt-1 text-2xl font-semibold">קובץ מסווג להורדה</h2>

              <div className="mt-5 bg-[#eef0ea]">
                <div
                  className="h-3 bg-[#477061] transition-all duration-500"
                  style={{ width: result ? "100%" : isSubmitting ? "70%" : "0%" }}
                />
              </div>

              <p className="mt-3 text-sm leading-6 text-[#626d66]">
                {result
                  ? `הסיווג הסתיים. נמצאו ${result.rowCount} תנועות בקבצים שהועלו.`
                  : isSubmitting
                    ? "המערכת קוראת את קבצי ה-Excel, מסווגת פעולות ומייצרת דוח."
                    : "לאחר ההעלאה יופיע כאן קישור להורדת קובץ ה-Excel המסווג."}
              </p>

              {result ? (
                <a
                  className="mt-5 inline-flex h-11 items-center justify-center bg-[#477061] px-5 text-sm font-semibold text-white transition hover:bg-[#36584b]"
                  href={result.downloadUrl}
                >
                  הורד דוח Excel
                </a>
              ) : null}
            </section>
          </form>

          <aside className="border border-[#d7d6cc] bg-white p-5 shadow-sm">
            <p className="text-sm font-semibold text-[#477061]">איך זה עובד</p>
            <h2 className="mt-1 text-2xl font-semibold">תהליך עיבוד מאובטח</h2>
            <div className="mt-5 grid gap-4 text-sm leading-7 text-[#59645e]">
              <p>
                המערכת שומרת את קבצי המקור והדוח שנוצר למשך 30 יום, ולא מציגה רשימת דוחות או קבצים קודמים למבקרים
                באתר.
              </p>
              <p>
                הסיווג מבוסס על קובץ מיפוי פנימי בשם category-mapping.xlsx, עם עמודות keyword, category ו-status.
              </p>
              <p>פעולות שלא נמצאה עבורן מילת מפתח מסומנות אוטומטית כלא מסווגות ונשלחות לבדיקה.</p>
            </div>

            <div className="mt-6 border border-[#d7d6cc]">
              <table className="w-full border-collapse text-right text-sm">
                <thead className="bg-[#eef0ea] text-[#33443d]">
                  <tr>
                    <th className="border-b border-[#d7d6cc] p-3 font-semibold">עמודה</th>
                    <th className="border-b border-[#d7d6cc] p-3 font-semibold">פלט</th>
                  </tr>
                </thead>
                <tbody>
                  {["תאריך", "מקור", "תיאור פעולה", "סיווג", "סכום", "סטטוס"].map((column) => (
                    <tr className="odd:bg-white even:bg-[#fbfaf6]" key={column}>
                      <td className="border-b border-[#ecece4] p-3">{column}</td>
                      <td className="border-b border-[#ecece4] p-3 text-[#59645e]">Excel</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </aside>
        </div>
      </section>
    </main>
  );
}
