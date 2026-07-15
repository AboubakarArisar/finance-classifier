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
  { title: "העלאת קבצים", text: "גררו או בחרו את תדפיסי האקסל מהבנק ומחברת האשראי" },
  { title: "עיבוד אוטומטי", text: "לחצו אישור (לאחר שכל הקבצים מופיעים בחלון הקבצים)" },
  { title: "הורדת דוח", text: "המערכת מכינה ומורידה למחשב את קובץ השיקוף מוכן לסיווג" },
];

export default function Home() {
  const [files, setFiles] = useState<File[]>([]);
  const [error, setError] = useState("");
  const [result, setResult] = useState<AnalyzeResponse | null>(null);
  const [isSubmitting, setIsSubmitting] = useState(false);
  const [isDownloading, setIsDownloading] = useState(false);
  const [isDragging, setIsDragging] = useState(false);
  const [inputVersion, setInputVersion] = useState(0);
  const [acceptedTerms, setAcceptedTerms] = useState(false);
  // Detected statement type per file (keyed by fileKey). "unknown" = detection ran but found nothing.
  const [kinds, setKinds] = useState<Record<string, FileKind | "unknown">>({});
  const detectingRef = useRef<Set<string>>(new Set());

  const canSubmit = useMemo(
    () => files.length > 0 && !isSubmitting && acceptedTerms,
    [files, isSubmitting, acceptedTerms],
  );

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
        <aside className="flex flex-col md:order-first order-second justify-between gap-8 rounded-card bg-primary p-8 text-surface">
          <div>
            <span className="inline-flex items-center gap-2 rounded-pill bg-surface/10 px-4 py-1.5 text-sm font-medium text-surface/90">
              <span className="h-2 w-2 rounded-full bg-accent" />
              כלי לבניית שיקוף ותקציב
            </span>
            <h1 className="mt-6 text-2xl font-semibold leading-snug sm:text-3xl">
              דפי בנק וכרטיסי אשראי,
              <br />
              מסודרים לסיווג בלחיצה.
            </h1>
            <p className="mt-4 text-sm leading-relaxed text-surface/70">
              העלו את הקבצים והמערכת תבנה עבורכם את השיקוף החודשי הנדרש
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
        <section className="rounded-card md:order-second order-first bg-surface p-6 shadow-card sm:p-8">
          <header className="flex items-center justify-between border-b border-border/60 pb-5">
            <h2 className="text-2xl font-semibold text-text-strong">העלאת קבצים</h2>
            {/* eslint-disable-next-line @next/next/no-img-element */}
            <img
              src="https://res.cloudinary.com/dnpxugbk9/image/upload/v1783013365/logo_lrsktq.png"
              alt="Benny Vazana"
              className="h-14 w-44 object-contain"
            />
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
                <div className="nice-scroll h-72 overflow-y-auto">
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

            {/* Terms acceptance — must be checked before processing */}
            <label className="flex items-center justify-start gap-2 text-sm text-text-muted">
              <input
                type="checkbox"
                checked={acceptedTerms}
                onChange={(event) => setAcceptedTerms(event.target.checked)}
                className="h-4 w-4 accent-primary"
              />
              <span>
                אני מאשר את{" "}
                <a
                  href="https://bennyvazana.com/wp-content/uploads/2026/07/tazrim-plus.pdf"
                  target="_blank"
                  rel="noopener noreferrer"
                  className="font-semibold text-primary underline hover:text-primary-hover"
                >
                  הסכם שימוש ופרטיות
                </a>
              </span>
            </label>

            {/* Footer actions */}
            <div className="flex flex-wrap items-center justify-between gap-2 border-t border-border/60 pt-5 md:flex-nowrap">
              <a
                href="https://wa.me/972507817032?text=%D7%90%D7%A0%D7%99%20%D7%A0%D7%9E%D7%A6%D7%90%20%D7%91%D7%9E%D7%A2%D7%A8%D7%9B%D7%AA%20%D7%94%D7%A9%D7%99%D7%A7%D7%95%D7%A3%20%D7%A9%D7%9C%D7%9A%20%D7%95%D7%90%D7%A0%D7%99%20%D7%A6%D7%A8%D7%99%D7%9A%20%D7%A2%D7%96%D7%A8%D7%94%20%D7%98%D7%9B%D7%A0%D7%99%D7%AA"
                target="_blank"
                rel="noopener noreferrer"
                aria-label="צרו קשר בוואטסאפ"
                className="inline-flex h-12 shrink-0 items-center justify-center gap-2 rounded-pill bg-[#25D366] px-4 text-sm font-semibold text-white transition-colors hover:bg-[#1ebe5d]"
              >
                <svg viewBox="0 0 24 24" className="h-4 w-4" fill="currentColor" aria-hidden="true">
                  <path d="M17.472 14.382c-.297-.149-1.758-.867-2.03-.967-.273-.099-.471-.148-.67.15-.197.297-.767.966-.94 1.164-.173.199-.347.223-.644.075-.297-.15-1.255-.463-2.39-1.475-.883-.788-1.48-1.761-1.653-2.059-.173-.297-.018-.458.13-.606.134-.133.298-.347.446-.52.149-.174.198-.298.298-.497.099-.198.05-.371-.025-.52-.075-.149-.669-1.612-.916-2.207-.242-.579-.487-.5-.669-.51-.173-.008-.371-.01-.57-.01-.198 0-.52.074-.792.372-.272.297-1.04 1.016-1.04 2.479 0 1.462 1.065 2.875 1.213 3.074.149.198 2.096 3.2 5.077 4.487.709.306 1.262.489 1.694.625.712.227 1.36.195 1.871.118.571-.085 1.758-.719 2.006-1.413.248-.694.248-1.289.173-1.413-.074-.124-.272-.198-.57-.347m-5.421 7.403h-.004a9.87 9.87 0 01-5.031-1.378l-.361-.214-3.741.982.998-3.648-.235-.374a9.86 9.86 0 01-1.51-5.26c.001-5.45 4.436-9.884 9.888-9.884 2.64 0 5.122 1.03 6.988 2.898a9.825 9.825 0 012.893 6.994c-.003 5.45-4.437 9.885-9.885 9.885M20.52 3.449C18.24 1.245 15.24 0 12.045 0 5.463 0 .104 5.359.101 11.892c0 2.096.549 4.142 1.595 5.945L0 24l6.335-1.652a11.899 11.899 0 005.71 1.454h.006c6.585 0 11.945-5.359 11.949-11.893a11.821 11.821 0 00-3.495-8.46z" />
                </svg>
                <span>צרו קשר בוואטסאפ</span>
              </a>
              <div className="flex shrink-0 items-center gap-2">
              {result ? (
                <button
                  className="inline-flex h-12 items-center justify-center gap-2 rounded-pill bg-primary px-6 text-base font-semibold text-surface transition-colors hover:bg-primary-hover disabled:cursor-not-allowed disabled:opacity-60"
                  disabled={isDownloading}
                  onClick={downloadReport}
                  type="button"
                >
                  {isDownloading ? <Spinner className="h-5 w-5" /> : null}
                  <span>{isDownloading ? "מכין להורדה..." : "הורדת דוח Excel"}</span>
                </button>
              ) : (
                <button
                  className="inline-flex h-12 items-center justify-center gap-2 rounded-pill bg-primary px-6 text-base font-semibold text-surface transition-colors hover:bg-primary-hover disabled:cursor-not-allowed disabled:opacity-50"
                  disabled={!canSubmit}
                  type="submit"
                >
                  {isSubmitting ? <Spinner className="h-5 w-5" /> : null}
                  <span>{isSubmitting ? "מעבד..." : "אישור"}</span>
                </button>
              )}
              <button
                className="inline-flex h-12 items-center justify-center rounded-pill border border-border bg-surface px-6 text-base font-medium text-text transition-colors hover:border-primary hover:text-text-strong"
                onClick={resetForm}
                type="button"
              >
                {result ? "סיווג חדש" : "ביטול"}
              </button>
              </div>
            </div>
          </form>
        </section>
      </div>
    </main>
  );
}
