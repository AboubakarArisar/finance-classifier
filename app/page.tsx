"use client";

import { ChangeEvent, DragEvent, FormEvent, useMemo, useState } from "react";

type AnalyzeResponse = {
  downloadUrl: string;
  fileName: string;
  jobId: string;
  reportBase64: string;
  rowCount: number;
};

const maxFiles = 120;

function isExcelFile(file: File) {
  const name = file.name.toLowerCase();
  return name.endsWith(".xls") || name.endsWith(".xlsx") || name.endsWith(".xlsm");
}

function fileKey(file: File) {
  return `${file.name}-${file.size}-${file.lastModified}`;
}

// Best-effort label from the file name; the server still auto-detects the real type.
function paymentMethod(file: File) {
  const name = file.name;
  if (/אשראי|סקיי|כרטיס|מסטרקארד|ויזה|כאל|מקס|ישראכרט|אמריקן|דיינרס|פירוט עסקאות/.test(name)) {
    return "כרטיס אשראי";
  }
  if (/בנק|עו"?ש|עובר ושב|חשבון|תנועות בחשבון/.test(name)) {
    return "חשבון בנק";
  }
  return "—";
}

export default function Home() {
  const [files, setFiles] = useState<File[]>([]);
  const [error, setError] = useState("");
  const [result, setResult] = useState<AnalyzeResponse | null>(null);
  const [isSubmitting, setIsSubmitting] = useState(false);
  const [isDragging, setIsDragging] = useState(false);
  const [inputVersion, setInputVersion] = useState(0);

  const canSubmit = useMemo(() => files.length > 0 && !isSubmitting, [files, isSubmitting]);

  function addFiles(incoming: File[]) {
    setResult(null);

    if (incoming.length === 0) {
      return;
    }

    if (incoming.some((file) => !isExcelFile(file))) {
      setError("ניתן להעלות קבצי Excel בלבד: XLS, XLSX או XLSM.");
      return;
    }

    setFiles((current) => {
      const existing = new Set(current.map(fileKey));
      const merged = [...current, ...incoming.filter((file) => !existing.has(fileKey(file)))];

      if (merged.length > maxFiles) {
        setError(`ניתן להעלות עד ${maxFiles} קבצי Excel.`);
        return current;
      }

      setError("");
      return merged;
    });
  }

  function handleInputChange(event: ChangeEvent<HTMLInputElement>) {
    const selected = Array.from(event.target.files ?? []);
    event.target.value = "";
    addFiles(selected);
  }

  function handleDrop(event: DragEvent<HTMLLabelElement>) {
    event.preventDefault();
    setIsDragging(false);
    addFiles(Array.from(event.dataTransfer.files ?? []));
  }

  function removeFile(target: File) {
    setResult(null);
    setError("");
    setFiles((current) => current.filter((file) => fileKey(file) !== fileKey(target)));
  }

  function resetForm() {
    setFiles([]);
    setError("");
    setResult(null);
    setIsSubmitting(false);
    setInputVersion((current) => current + 1);
  }

  async function submitFiles(event: FormEvent<HTMLFormElement>) {
    event.preventDefault();

    if (files.length === 0) {
      setError("יש להעלות לפחות קובץ Excel אחד.");
      return;
    }

    const formData = new FormData();
    files.forEach((file) => formData.append("files", file));

    setError("");
    setResult(null);
    setIsSubmitting(true);

    try {
      const response = await fetch("/api/analyze", { body: formData, method: "POST" });
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

  function downloadReport() {
    if (!result) {
      return;
    }

    const binary = window.atob(result.reportBase64);
    const bytes = Uint8Array.from(binary, (character) => character.charCodeAt(0));
    const blob = new Blob([bytes], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = result.fileName;
    link.click();
    URL.revokeObjectURL(url);
  }

  return (
    <main className="flex min-h-screen w-full justify-center px-4 py-10">
      <section className="w-full max-w-3xl rounded-card bg-surface p-6 shadow-card sm:p-8">
        {/* Header */}
        <header className="flex items-center justify-between border-b border-border/60 pb-5">
          <h1 className="text-2xl font-semibold text-text-strong sm:text-3xl">סיווג תנועות</h1>
          <p className="text-sm font-medium text-accent">דפי בנק וכרטיסי אשראי</p>
        </header>

        <form className="mt-6 flex flex-col gap-5" onSubmit={submitFiles}>
          {/* Upload instruction */}
          <p className="text-center text-sm text-text">
            ניתן להעלות יותר מקובץ אחד
            <span className="mx-1 font-semibold text-accent">·</span>
            <span className="font-semibold text-accent">קבצי אקסל XLS, XLSX בלבד</span>
          </p>

          {/* Dropzone */}
          <label
            className={`flex cursor-pointer flex-col items-center justify-center rounded-card border-2 border-dashed px-5 py-10 text-center transition-colors ${
              isDragging ? "border-accent bg-row-alt" : "border-border bg-bg hover:border-accent hover:bg-surface"
            }`}
            onDragOver={(event) => {
              event.preventDefault();
              setIsDragging(true);
            }}
            onDragLeave={() => setIsDragging(false)}
            onDrop={handleDrop}
          >
            <input
              accept=".xls,.xlsx,.xlsm,application/vnd.ms-excel,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
              className="sr-only"
              key={inputVersion}
              multiple
              onChange={handleInputChange}
              type="file"
            />
            <span className="text-base text-text">
              ניתן לגרור או לבחור <span className="font-semibold text-accent">קבצים כאן</span>
            </span>
            <span className="mt-2 text-xs text-text-muted">לכל כרטיס ולכל חודש קובץ נפרד · הקבצים מצטברים</span>
          </label>

          {/* Files table */}
          <div>
            <h2 className="mb-2 text-lg font-semibold text-text-strong">קבצים</h2>
            <div className="overflow-hidden rounded-card border border-border/70">
              <table className="w-full border-collapse text-right text-sm">
                <thead className="bg-table-head text-text-strong">
                  <tr>
                    <th className="px-4 py-3 font-semibold">קובץ</th>
                    <th className="px-4 py-3 font-semibold">אמצעי תשלום</th>
                    <th className="w-12 px-4 py-3 font-semibold" aria-label="הסרה" />
                  </tr>
                </thead>
                <tbody>
                  {files.length === 0 ? (
                    <tr>
                      <td className="px-4 py-8 text-center text-text-muted" colSpan={3}>
                        עדיין לא נבחרו קבצים
                      </td>
                    </tr>
                  ) : (
                    files.map((file) => (
                      <tr className="border-t border-border/40 odd:bg-surface even:bg-row-alt" key={fileKey(file)}>
                        <td className="px-4 py-3">
                          <span className="block truncate text-text-strong">{file.name}</span>
                        </td>
                        <td className="px-4 py-3 text-text">{paymentMethod(file)}</td>
                        <td className="px-4 py-3 text-center">
                          <button
                            aria-label={`הסר ${file.name}`}
                            className="font-bold text-text-muted transition-colors hover:text-danger"
                            onClick={() => removeFile(file)}
                            type="button"
                          >
                            ✕
                          </button>
                        </td>
                      </tr>
                    ))
                  )}
                </tbody>
              </table>
            </div>
          </div>

          {error ? (
            <p className="rounded-card border border-danger/40 bg-danger/5 px-4 py-3 text-sm font-medium text-danger">
              {error}
            </p>
          ) : null}

          {/* Progress + result */}
          {(isSubmitting || result) && !error ? (
            <div className="rounded-card border border-border/60 bg-bg px-4 py-4">
              <div className="h-2 overflow-hidden rounded-pill bg-table-head">
                <div
                  className="h-full bg-accent transition-all duration-500"
                  style={{ width: result ? "100%" : "70%" }}
                />
              </div>
              <p className="mt-3 text-sm text-text">
                {result
                  ? `הסיווג הסתיים — נמצאו ${result.rowCount} תנועות.`
                  : "המערכת קוראת את הקבצים, מסווגת פעולות ומייצרת דוח..."}
              </p>
            </div>
          ) : null}

          {/* Footer actions */}
          <div className="flex flex-wrap items-center justify-end gap-3 border-t border-border/60 pt-5">
            {result ? (
              <button
                className="inline-flex h-12 items-center justify-center rounded-pill bg-primary px-8 text-base font-semibold text-surface transition-colors hover:bg-primary-hover"
                onClick={downloadReport}
                type="button"
              >
                הורדת דוח Excel
              </button>
            ) : (
              <button
                className="inline-flex h-12 items-center justify-center rounded-pill bg-primary px-8 text-base font-semibold text-surface transition-colors hover:bg-primary-hover disabled:cursor-not-allowed disabled:opacity-50"
                disabled={!canSubmit}
                type="submit"
              >
                {isSubmitting ? "מעבד..." : "אישור"}
              </button>
            )}
            <button
              className="inline-flex h-12 items-center justify-center rounded-pill border border-border bg-surface px-8 text-base font-medium text-text transition-colors hover:border-primary hover:text-text-strong"
              onClick={resetForm}
              type="button"
            >
              {result ? "סיווג חדש" : "ביטול"}
            </button>
          </div>
        </form>
      </section>
    </main>
  );
}
