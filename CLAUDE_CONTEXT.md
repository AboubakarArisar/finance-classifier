# Project Overview
- Web app for **Benny Vazana** (finance coaching). A user uploads Israeli bank + credit-card Excel statements; the app parses them and generates a single Hebrew, right-to-left (RTL) Excel workbook ("שיקוף פיננסי") that the user downloads and classifies by hand.
- Main goals: (1) parse many statement formats reliably; (2) produce a polished, formula-driven RTL report with dropdowns, KPIs, and pie charts; (3) leave every transaction **unclassified on purpose** so the end user assigns categories manually.

# Tech Stack
- **Framework:** Next.js 16.2.6 (App Router, React 19.2.4). ⚠️ This Next.js has breaking changes vs. training data — read `node_modules/next/dist/docs/` before writing Next-specific code (see AGENTS.md).
- **Language:** TypeScript 5, Node.js runtime (`export const runtime = "nodejs"` on API routes).
- **Styling:** Tailwind CSS v4 (`@tailwindcss/postcss`).
- **Excel:** `exceljs` ^4.4.0 (report generation), `xlsx` ^0.18.5 (parsing uploads).
- **Database:** None. Jobs persist as files on disk (see below).
- **Hosting:** Vercel. Auto-deploys on push to `origin/main`.

# Current Architecture
Folder structure (source only; ignore `node_modules/`):
```
app/
  layout.tsx                  # <html lang="he" dir="rtl">
  page.tsx                    # single-page client UI (upload form, download, WhatsApp button)
  api/analyze/route.ts        # POST: accept files -> analyzeFinancialStatements()
  api/reports/[jobId]/route.ts# GET: stream a saved report.xlsx by jobId
lib/
  finance-analyzer.ts         # CORE (~1700 lines): parse, classify, build workbook
  excel-charts.ts             # injects native pie charts into the ExcelJS workbook
data/
  category-mapping.xlsx       # keyword->category overrides (built by scripts/)
  paamonim-mappings.json      # ~19k-rule reference classifier DB
  shamir-template.xlsx        # reference template
scripts/create-category-mapping.mjs
public/                       # logo.png/jpg + default svgs
```
- **Important services:** everything runs through `lib/finance-analyzer.ts`. Entry `analyzeFinancialStatements(files)` → writes uploads to a job dir, parses each workbook (auto-detects bank vs credit from headers), normalizes transactions, builds the report, returns `{ downloadUrl, fileName, jobId, reportBase64, rowCount }`. `readReport(jobId)` reads back `report.xlsx`.
- **Authentication:** None. Public endpoints.
- **APIs:** `POST /api/analyze` (multipart `files`; legacy `creditFile`/`bankFile` still accepted; max 120 files; Excel only). `GET /api/reports/[jobId]` (streams saved report). UI in `page.tsx` calls `/api/analyze` then downloads via `reportBase64`.
- **Data model (in-memory only):** `NormalizedTransaction` (amount, date, description, direction `הוצאה`/`הכנסה`, mainCategory, subCategory, recurrence, sourceName, …) and `CategoryGroup` (mainCategory, namedRange, subCategories[]). Category catalog = 19 groups (0–15 expense, 16–18 income).
- **Jobs on disk:** `getJobsDir()` → `FINANCE_DATA_DIR/jobs`, else serverless `os.tmpdir()/benny-finance-classifier/jobs`, else local `.data/jobs`. Each job: `<jobId>/uploads/*` + `<jobId>/report.xlsx`. 30-day retention (`cleanupExpiredJobs`).

# Features Completed
- **Multi-format parsing:** bank + credit statements; per-sheet auto-detection; date/total quirk handling.
- **Rule-based classification engine** (currently **disabled by client request** — all category cells left empty): layered over `paamonim-mappings.json` (~19k rules) + `data/category-mapping.xlsx` keyword overrides + statement's own category column. Code retained for reference.
- **Report workbook** — 6 sheets: `סיכום קבצים` (file summary), `סיווג תנועות` (main transaction table w/ dropdowns + validations + named ranges), `תוצאות השיקוף` (KPIs, category totals, pie charts), `רשימת קטגוריות`, `ליבוא מהתוכנה`, `בחירות`.
- **Design polish (recent):** filename `Benny-Vazana-Finance.xlsx`; whole transaction-table body right-aligned (headers untouched); conditional color-coding on sheet 1 columns E/G/H (income=blue, expense=burgundy/pink, uncategorized=yellow); KPI labels on one line (col A width 22); result-sheet G-column category names colored by direction (expense burgundy `FF9C0006`, income navy `FF1F4E79`, **text only, no fill**); WhatsApp contact button + single-line footer in the UI.

# Current State
- **Working:** upload → parse → generate → download end-to-end; all 6 sheets render; charts inject; RTL styling correct. Verified via openpyxl on real statements (711 transactions, 117 distinct dates).
- **Partially finished:** none outstanding.
- **Known bugs:** none open. Prior client concern about "only 01/01 dates" was diagnosed as not-a-bug (top rows coincidentally 01/01; dates stored as strings, unaffected by alignment).

# Environment Variables
- `FINANCE_DATA_DIR` — optional base dir for job storage; overrides default location.
- `VERCEL` / `AWS_LAMBDA_FUNCTION_NAME` / `WEBSITE_INSTANCE_ID` — platform-set flags; any truthy value makes the app store jobs in the OS temp dir (serverless mode). Not set manually.
- No secrets/API keys in this project.

# Important Decisions
- **DESIGN-ONLY constraint (standing, hard):** never change the report's internal working — no edits to formulas, validations, named ranges, column structure, or classification logic. All visual changes must be purely additive. This is the client's explicit, repeated instruction.
- **Auto-classification intentionally off:** every row imports with empty `סעיף ראשי` / `שם סעיף`; user classifies manually.
- **`תוצאות השיקוף` G/H blank-row gap is deliberate:** it separates the expense main-category block (top) from the income block (bottom). The income summary pie reads data from the row after the gap (`incomeStart = expenseEnd + 2`). Do **not** delete it — it would break the chart's data range.
- **ExcelJS quirks:** no `stopIfTrue` on conditional formatting — order by `priority` (lower number wins). CF fills use `bgColor`; direct cell fills use `fgColor`. Chart refs in `excel-charts.ts` are coupled to result-sheet row positions — keep them in sync.
- **Conventions:** match surrounding code style; comment intent (esp. "presentation only" on visual edits); Hebrew strings inline; git commits end with `Co-Authored-By: Claude Opus 4.8 <noreply@anthropic.com>`.

# Pending Tasks
- None active. Work is iterative design feedback from the client (relayed via the operator). Await next request.
- Priority when tasks arrive: honor the design-only constraint; verify every change by regenerating a real report and inspecting with openpyxl before committing.

# Commands
- **Dev:** `npm run dev`
- **Build:** `npm run build`
- **Typecheck:** `npx tsc --noEmit -p tsconfig.json`
- **Lint:** `npm run lint`
- **Rebuild keyword mapping:** `npm run create:mapping`
- **Test:** no test suite. Verification pattern = regenerate a report from real uploads, inspect with openpyxl (`PYTHONIOENCODING=utf-8` on Windows).
- **Deploy:** push to `origin/main` → Vercel auto-deploys.

### Verifying a generated report (throwaway harness)
Node's type-stripping can't resolve extensionless relative imports, so temporarily rewrite the import:
```
cp lib/finance-analyzer.ts lib/finance-analyzer.ts.bak \
  && sed -i 's#from "./excel-charts"#from "./excel-charts.ts"#' lib/finance-analyzer.ts \
  && node --experimental-strip-types lib/_gen_verify.ts; \
  mv lib/finance-analyzer.ts.bak lib/finance-analyzer.ts; rm -f lib/_gen_verify.ts
```
Harness calls `analyzeFinancialStatements` with `{ fileName, kind, buffer }` objects (`kind = name.includes("-bank-") ? "bank" : "credit"`), decodes `res.reportBase64`, writes `.data/final-verify.xlsx`.

# Things to Remember
- **Never edit** (data/reference assets): `data/paamonim-mappings.json`, `data/category-mapping.xlsx`, `data/shamir-template.xlsx`.
- **Do not touch the report's logic** in `lib/finance-analyzer.ts` — only styling. Read AGENTS.md (Next.js has breaking changes; consult `node_modules/next/dist/docs/`).
- **Windows/PowerShell env:** LF→CRLF git warnings are expected/harmless. Bash tool available for POSIX scripts.
- **Client workflow:** feedback arrives via the operator (often screenshots + Hebrew). Investigate in code before changing anything; confirm the client's hypothesis; keep all edits visual. The operator sometimes commits/pushes to `main` directly.
- Auto-memory index lives at `C:\Users\pinda\.claude\projects\D--client2\memory\MEMORY.md` (classification engine + file-format notes).
