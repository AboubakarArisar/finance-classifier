"use client";

import { ChangeEvent, DragEvent, FormEvent, useEffect, useMemo, useRef, useState } from "react";

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

type FileKind = "bank" | "credit";

// Instant, best-effort guess from the file name (shown while the content is being
// read). Returns null when the name gives no hint.
function guessKindFromName(file: File): FileKind | null {
  const name = file.name;
  if (/אשראי|סקיי|כרטיס|מסטרקארד|ויזה|כאל|מקס|ישראכרט|אמריקן|דיינרס|פירוט עסקאות/.test(name)) {
    return "credit";
  }
  if (/בנק|עו"?ש|עובר ושב|חשבון|תנועות בחשבון/.test(name)) {
    return "bank";
  }
  return null;
}

function kindLabel(kind: FileKind) {
  return kind === "bank" ? "חשבון בנק" : "כרטיס אשראי";
}

// Mirrors the server's header normalisation so client-side detection matches
// exactly what parseWorkbook does.
function normalizeHeaderCell(value: unknown) {
  return String(value ?? "").replace(/\s+/g, " ").trim().toLocaleLowerCase("he-IL");
}

function rowMatchesHeaders(cells: string[], required: string[]) {
  return required.every((header) => cells.includes(normalizeHeaderCell(header)));
}

// Same header signatures the server uses to pick a statement layout
// (parseCreditSheet / parseBankSheet / parseIsracardSheet).
function detectKindFromRows(rows: unknown[][]): FileKind | null {
  for (const row of rows.slice(0, 20)) {
    const cells = row.map(normalizeHeaderCell);
    if (rowMatchesHeaders(cells, ["תאריך עסקה", "שם בית העסק", "סכום חיוב"])) return "credit";
    if (rowMatchesHeaders(cells, ["תאריך", "הפעולה", "חובה", "זכות"])) return "bank";
    if (rowMatchesHeaders(cells, ["שם בית עסק", "סכום חיוב"])) return "credit";
  }
  return null;
}

// Reads the workbook in the browser (SheetJS, loaded on demand) and returns the
// real statement type, exactly as the server would classify it. Used only for
// the upload preview label — the server remains the source of truth.
async function detectFileKind(file: File): Promise<FileKind | null> {
  try {
    const XLSX = await import("xlsx");
    const buffer = await file.arrayBuffer();
    const workbook = XLSX.read(buffer, { type: "array" });
    for (const sheetName of workbook.SheetNames) {
      const rows = XLSX.utils.sheet_to_json<unknown[]>(workbook.Sheets[sheetName], {
        blankrows: false,
        defval: "",
        header: 1,
        raw: false,
      });
      const kind = detectKindFromRows(rows);
      if (kind) return kind;
    }
  } catch {
    return null;
  }
  return null;
}

// A simple inline loading spinner used while the report is being generated or downloaded.
function Spinner({ className = "" }: { className?: string }) {
  return (
    <svg className={`animate-spin ${className}`} viewBox="0 0 24 24" fill="none" aria-hidden="true">
      <circle className="opacity-25" cx="12" cy="12" r="10" stroke="currentColor" strokeWidth="4" />
      <path
        className="opacity-90"
        d="M12 2a10 10 0 0 1 10 10"
        stroke="currentColor"
        strokeWidth="4"
        strokeLinecap="round"
      />
    </svg>
  );
}

const steps = [
  { title: "העלאת קבצים", text: "גררו או בחרו דפי בנק וכרטיסי אשראי — קובץ לכל חודש וכרטיס." },
  { title: "עיבוד אוטומטי", text: "המערכת קוראת, מאחדת ומכינה את כל התנועות לסיווג." },
  { title: "הורדת דוח", text: "מקבלים קובץ Excel מסודר, מוכן לסיווג ידני ולניתוח." },
];

export default function Home() {
  const [files, setFiles] = useState<File[]>([]);
  const [error, setError] = useState("");
  const [result, setResult] = useState<AnalyzeResponse | null>(null);
  const [isSubmitting, setIsSubmitting] = useState(false);
  const [isDownloading, setIsDownloading] = useState(false);
  const [isDragging, setIsDragging] = useState(false);
  const [inputVersion, setInputVersion] = useState(0);
  // Detected statement type per file (keyed by fileKey). "unknown" = detection ran but found nothing.
  const [kinds, setKinds] = useState<Record<string, FileKind | "unknown">>({});
  const detectingRef = useRef<Set<string>>(new Set());

  const canSubmit = useMemo(() => files.length > 0 && !isSubmitting, [files, isSubmitting]);

  // Detect the real statement type for any newly added file, in the background.
  useEffect(() => {
    files.forEach((file) => {
      const key = fileKey(file);
      if (kinds[key] !== undefined || detectingRef.current.has(key)) return;
      detectingRef.current.add(key);
      detectFileKind(file)
        .then((kind) => setKinds((prev) => ({ ...prev, [key]: kind ?? "unknown" })))
        .finally(() => detectingRef.current.delete(key));
    });
  }, [files, kinds]);

  // The label to show in the file table: detected type wins, then the name-based
  // guess, then a transient "detecting" note.
  function fileTypeLabel(file: File) {
    const detected = kinds[fileKey(file)];
    if (detected === "bank" || detected === "credit") return kindLabel(detected);
    const guessed = guessKindFromName(file);
    if (guessed) return kindLabel(guessed);
    return detected === "unknown" ? "לא זוהה" : "מזהה…";
  }

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
    const key = fileKey(target);
    setFiles((current) => current.filter((file) => fileKey(file) !== key));
    setKinds((current) => {
      if (!(key in current)) return current;
      const next = { ...current };
      delete next[key];
      return next;
    });
  }

  function resetForm() {
    setFiles([]);
    setError("");
    setResult(null);
    setIsSubmitting(false);
    setInputVersion((current) => current + 1);
    setKinds({});
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

  async function downloadReport() {
    if (!result || isDownloading) {
      return;
    }

    setIsDownloading(true);
    // Yield once so the spinner can paint before the synchronous blob work runs.
    await new Promise((resolve) => setTimeout(resolve, 50));

    try {
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
    } finally {
      setIsDownloading(false);
    }
  }

  return (
    <main className="flex min-h-screen w-full items-center justify-center px-4 py-10">
      <div className="grid w-full max-w-5xl gap-6 lg:grid-cols-[0.85fr_1.15fr]">
        {/* Brand / steps panel */}
        <aside className="flex flex-col justify-between gap-8 rounded-card bg-primary p-8 text-surface">
          <div>
            <span className="inline-flex items-center gap-2 rounded-pill bg-surface/10 px-4 py-1.5 text-sm font-medium text-surface/90">
              <span className="h-2 w-2 rounded-full bg-accent" />
              כלי סיווג תנועות
            </span>
            <h1 className="mt-6 text-2xl font-semibold leading-snug sm:text-3xl">
              דפי בנק וכרטיסי אשראי,
              <br />
              מסודרים לסיווג בלחיצה.
            </h1>
            <p className="mt-4 text-sm leading-relaxed text-surface/70">
              העלו את הקבצים, והמערכת תאחד את כל התנועות לדוח Excel נקי ומוכן לעבודה.
            </p>

            <div className="mt-6 overflow-hidden rounded-card border border-surface/10 shadow-card">
              <iframe
                className="aspect-video w-full"
                src="https://www.youtube.com/embed/DxrmYrGmhks"
                title="איך המערכת עובדת"
                allow="accelerometer; autoplay; clipboard-write; encrypted-media; gyroscope; picture-in-picture; web-share"
                referrerPolicy="strict-origin-when-cross-origin"
                allowFullScreen
              />
            </div>
          </div>

          <ol className="flex flex-col gap-4">
            {steps.map((step, index) => (
              <li key={step.title} className="flex gap-3">
                <span className="flex h-8 w-8 shrink-0 items-center justify-center rounded-pill bg-accent text-sm font-bold text-surface">
                  {index + 1}
                </span>
                <div>
                  <p className="font-semibold text-surface">{step.title}</p>
                  <p className="text-sm text-surface/60">{step.text}</p>
                </div>
              </li>
            ))}
          </ol>
        </aside>

        {/* Upload card */}
        <section className="rounded-card bg-surface p-6 shadow-card sm:p-8">
          <header className="flex items-center justify-between border-b border-border/60 pb-5">
            <h2 className="text-2xl font-semibold text-text-strong">העלאת קבצים</h2>
            <p className="text-sm font-medium text-accent">XLS · XLSX · XLSM</p>
          </header>

          <form className="mt-6 flex flex-col gap-5" onSubmit={submitFiles}>
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
              <div className="mb-2 flex items-center justify-between">
                <h3 className="text-lg font-semibold text-text-strong">קבצים</h3>
                {files.length > 0 ? (
                  <span className="rounded-pill bg-table-head px-3 py-1 text-xs font-medium text-text">
                    {files.length} קבצים
                  </span>
                ) : null}
              </div>
              <div className="overflow-hidden rounded-card border border-border/70">
                <div className="nice-scroll max-h-72 overflow-y-auto">
                  <table className="w-full border-collapse text-right text-sm">
                  <thead className="sticky top-0 z-10 bg-table-head text-text-strong">
                    <tr>
                      <th className="px-4 py-3 font-semibold">קובץ</th>
                      <th className="px-4 py-3 font-semibold">סוג הקובץ</th>
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
                          <td className="px-4 py-3 text-text">{fileTypeLabel(file)}</td>
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
            </div>

            {error ? (
              <p className="rounded-card border border-danger/40 bg-danger/5 px-4 py-3 text-sm font-medium text-danger">
                {error}
              </p>
            ) : null}

            {/* Progress + result */}
            {(isSubmitting || result) && !error ? (
              <div className="flex items-center gap-3 rounded-card border border-border/60 bg-bg px-4 py-4">
                {isSubmitting ? <Spinner className="h-5 w-5 shrink-0 text-accent" /> : null}
                <div className="min-w-0 flex-1">
                  <div className="h-2 overflow-hidden rounded-pill bg-table-head">
                    <div
                      className={`h-full bg-accent transition-all duration-500 ${isSubmitting ? "animate-pulse" : ""}`}
                      style={{ width: result ? "100%" : "70%" }}
                    />
                  </div>
                  <p className="mt-3 text-sm text-text">
                    {result
                      ? `הסיווג הסתיים — נמצאו ${result.rowCount} תנועות.`
                      : "המערכת קוראת את הקבצים, מסווגת פעולות ומייצרת דוח..."}
                  </p>
                </div>
              </div>
            ) : null}

            {/* Footer actions */}
            <div className="flex flex-wrap items-center justify-end gap-3 border-t border-border/60 pt-5">
              {result ? (
                <button
                  className="inline-flex h-12 items-center justify-center gap-2 rounded-pill bg-primary px-8 text-base font-semibold text-surface transition-colors hover:bg-primary-hover disabled:cursor-not-allowed disabled:opacity-60"
                  disabled={isDownloading}
                  onClick={downloadReport}
                  type="button"
                >
                  {isDownloading ? <Spinner className="h-5 w-5" /> : null}
                  <span>{isDownloading ? "מכין להורדה..." : "הורדת דוח Excel"}</span>
                </button>
              ) : (
                <button
                  className="inline-flex h-12 items-center justify-center gap-2 rounded-pill bg-primary px-8 text-base font-semibold text-surface transition-colors hover:bg-primary-hover disabled:cursor-not-allowed disabled:opacity-50"
                  disabled={!canSubmit}
                  type="submit"
                >
                  {isSubmitting ? <Spinner className="h-5 w-5" /> : null}
                  <span>{isSubmitting ? "מעבד..." : "אישור"}</span>
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
      </div>
    </main>
  );
}
