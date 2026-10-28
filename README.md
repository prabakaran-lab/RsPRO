# RS PRO Datasheet Studio (XLSX → PDF platform)

Turns a product spreadsheet into branded **RS PRO datasheet PDFs** — one PDF per product row — and includes helper tools for product images, table extraction and supplier-workbook conversion.

| Tool (top menu) | What it does |
|---|---|
| **Create PDFs** | Reads a workbook (`.xlsx` / `.xls`), previews the first product, checks for duplicate specifications, then builds one PDF per row (single PDF or a ZIP). |
| **Image Library** | Upload / search / rename / delete product images. PDFs find pictures **by file name**. |
| **Table JSON** | Reads a table from a PDF or picture (manual crop, or bulk over many PDFs) and writes it as table JSON that can be pasted into a workbook cell. Uses PDF text where possible and OCR for pictures/scans. |
| **XLSX Formatter** | Converts a supplier workbook into the standard "uniform" workbook layout that **Create PDFs** reads. |

> **End-user documentation:** [`docs/RS_PRO_Datasheet_Studio_User_Manual.pdf`](docs/RS_PRO_Datasheet_Studio_User_Manual.pdf) — a screenshot-based manual for people who only *use* the website.

---

## 1. Tech stack

| Layer | Technology |
|---|---|
| Frontend | Next.js 14 (App Router), React 18, TypeScript, Tailwind CSS 3, a large custom stylesheet (`premium-ui.css`) |
| Backend | NestJS 10 (Express), TypeScript |
| PDF rendering | Puppeteer (headless Chrome) prints HTML templates to A4 PDF; `pdf-lib` / `@pdf-lib/fontkit` for PDF helpers; fonts: Oswald, Noto Sans |
| Table extraction | `pdfjs-dist` (PDF text layer), `tesseract.js` + `@tesseract.js-data/eng` (OCR), `@napi-rs/canvas` (image preparation) |
| Spreadsheets | SheetJS `xlsx` 0.20.2 |
| Database | SQLite via Prisma 5 (`backend/prisma/dev.db`) |
| ZIP | `archiver` |

---

## 2. Quick start (development)

### Requirements
- **Node.js 18.17 or newer** (developed and tested on Node 24, Windows 11) and npm
- Internet access on first install (npm packages, and Puppeteer downloads its own Chromium, ~170 MB)
- Python 3 is only needed for the optional helper script `setup_project.py`

### Install
```bash
# from the project root (the folder that contains package.json)
npm run install:all                 # installs root + backend + frontend dependencies

cd backend
npx prisma generate                 # create the Prisma client
npx prisma db push                  # create/update the SQLite schema (prisma/dev.db)
cd ..
```
`python setup_project.py` does the same installs + Prisma steps in one go.

### Run
```bash
npm run dev
```
This starts **both** servers together (via `concurrently`):

| Process | Command it runs | Address |
|---|---|---|
| Frontend (Next.js dev server) | `npm run dev --prefix frontend` → `next dev -p 3000` | http://localhost:3000 |
| Backend (NestJS, watch mode) | `npm run start:dev --prefix backend` → `nest start --watch` | http://localhost:8001 (API under `/api`) |

Open **http://localhost:3000**. The first start takes about a minute (the dev server compiles on demand). Stop with <kbd>Ctrl</kbd>+<kbd>C</kbd>.

On Windows, `run-localhost.cmd` runs the same `npm run dev` and writes its output to `codex-persistent.out.log` / `codex-persistent.err.log`.

### Run the pieces separately
```bash
npm run dev:backend      # backend only  → :8001
npm run dev:frontend     # frontend only → :3000
```

### Production-style run
```bash
npm run build:all                     # nest build + next build

# backend (note: compiled entry point is dist/main.js)
cd backend && node dist/main          # or: npm run start:prod

# frontend
cd frontend && npm start              # next start → :3000
```
`frontend/serve-built.js` is a minimal alternative: it serves the already-built `frontend/.next/server/app/index.html`, `/_next/*` assets and `public/` on `127.0.0.1:3000` without running Next.js (`node frontend/serve-built.js`).

---

## 3. Configuration (environment variables)

No `.env` file is required for local development. For any other host, set:

| Variable | Used by | Default | Purpose |
|---|---|---|---|
| `NEXT_PUBLIC_BACKEND_URL` | Frontend (**build time**) | `http://localhost:8001` | Base URL of the backend. The website calls `${NEXT_PUBLIC_BACKEND_URL}/api/...` and loads images from `${NEXT_PUBLIC_BACKEND_URL}/static/...`. Set it **before** `next build` / `next dev`. |
| `PORT` | Backend | `8001` | Port the NestJS server listens on. |
| `HOST_URL` | Backend | `http://localhost:8001` | Public base URL written into image/PDF links returned by the API (`${HOST_URL}/static/images/...`). Must match the address users reach the backend on. |

CORS is currently open to all origins (`origin: '*'`) and there is **no authentication** — put the app behind a login, VPN or reverse proxy before exposing it publicly.

---

## 4. How it runs (what happens at runtime)

1. **Create PDFs** — the browser uploads the workbook to `POST /api/job/upload`. The backend parses the first worksheet (two header rows: *roles* then *labels*; data from row 3), validates duplicate specifications, creates a `PdfJob` + one `PdfDetail` per row in SQLite and starts a background worker.
2. For each row the worker composes HTML (`pdf-compiler.service.ts`), resolves image names against `static/images`, lets Puppeteer print it to PDF and saves it as `static/pdfs/<stock_no>.pdf` (`/` and spaces in the stock number become `_`).
3. The browser polls `GET /api/job/:id` every 1.5 s for progress, then downloads the single PDF or `GET /api/job/:id/zip`.
4. **Preview** (`POST /api/pdf/preview`) renders only the first row; identical previews are cached in memory for 10 minutes.
5. **Table JSON** — `POST /api/extract/table` (manual crop) and `/api/extract/bulk` read PDF text positions with `pdfjs`, or fall back to OCR for pictures/scans (image is enlarged, a coloured header band is detected and read separately, each row is re-read one line at a time, then common glyphs such as Ø, ≤, Ω, mm² are repaired).
6. **Images** — uploaded files are saved to `static/images` (names are cleaned to `[A-Za-z0-9._-]`, same name = overwrite) and indexed in the `ImageAsset` table; files copied into the folder by hand are indexed automatically on the next list call.

Background resources created on demand: a headless Chrome instance (Puppeteer) and a Tesseract OCR worker (language data cached in `backend/.ocr-cache`).

---

## 5. Workbook format (what "Create PDFs" reads)

The **first worksheet** is used. Row 1 = *role* (technical column name), row 2 = *label* (friendly name), rows 3+ = one product per row.

| Role (row 1) | Content | Appears in the PDF as |
|---|---|---|
| `common_name` | Product family | Heading on every page |
| `cover_title` | Full product title | Large title, page 1 |
| `stock_no` | Unique product code | Part-number line + PDF file name |
| `cover_tagline` | Optional sentence | Bottom of page 1 (blank = standard text) |
| `cover_image` | Image **file name** | Framed photo, page 1 |
| `feature_list` | `<li><b>Title:</b> text</li>…` | Red FEATURES panel |
| `description_title`, `description` | Heading / `<p>…</p>` text | Description block |
| `variant_json` | One table-JSON per cell | Red tab + table |
| `image` | Image file names | Picture frames after the tables |

Columns become sections in the same left-to-right order as in the workbook; empty cells are skipped.

Table JSON examples:
```json
{"section":"Approvals","headers":[],"rows":[{"label":"Standards/Approvals","value":"EN61984;CE"}]}
{"section":"Solder contact spec","headers":["Number of contact","2","3"],"rows":[["Rated current (A)","13","13"],["Male contact face view","/SA12-face-view-2.png","/SA12-face-view-3.png"]]}
```
Image file names inside table cells are replaced by the picture.

---

## 6. HTTP API (backend, base `http://localhost:8001`)

| Method & path | Purpose |
|---|---|
| `POST /api/job/validate` | Duplicate-specification check for a workbook (`file`) |
| `POST /api/job/upload` | Create a PDF job (`file`, `headerMode` = `supplier`\|`rs`\|`both`, `cleanupDuplicates`) |
| `GET /api/job/list` · `GET /api/job/:id` · `DELETE /api/job/:id` | List / inspect / delete jobs |
| `GET /api/job/:id/zip` | Download all PDFs of a job as ZIP |
| `POST /api/pdf/preview` | Preview PDF of the first row (`file`, `headerMode`) |
| `GET /api/pdf/list?search=` · `GET /api/pdf/download/:filename` | List / download generated PDFs |
| `POST /api/template/upload` · `GET /api/template/exists` · `GET /api/template/view` | Reference-PDF template helpers |
| `POST /api/image/upload` (`files[]`) · `GET /api/image/list?search&type&page&limit` | Upload (up to 500/request) / list images |
| `PATCH /api/image/:id/rename` · `DELETE /api/image/:id` · `GET /api/image/download/:filename` | Rename / delete / download an image |
| `POST /api/extract/table` · `POST /api/extract/preview` | Manual-crop extraction / page preview (25 MB) |
| `POST /api/extract/bulk` · `POST /api/extract/bulk/xlsx` | Bulk PDF extraction (≤30 PDFs, 80 MB each) / build the result workbook |
| `POST /api/extract/table/xlsx` | Build a one-table workbook from table JSON |
| `POST /api/format-xlsx/analyze` · `POST /api/format-xlsx/convert` | Analyse / convert a supplier workbook (60 MB) |
| `GET /static/images/*` · `GET /static/pdfs/*` | Stored images and generated PDFs |

---

## 7. Database (Prisma / SQLite)

Schema: `backend/prisma/schema.prisma`, file: `backend/prisma/dev.db`.

| Model | Holds |
|---|---|
| `ImageAsset` | One row per library image (`filename`, `path`, `mimeType`, `sizeBytes`, `url`) |
| `PdfJob` | One row per generation run (`status`, `progress`, counts, `excelName`, `errorLog`) |
| `PdfDetail` | One row per product in a job (`stockNo`, `productName`, `filename`, `status`, `errorLog`) |

---

## 8. Project layout and file details

```
.
├── package.json                 root scripts (install:all, dev, build:all …)
├── backend/                     NestJS API  (port 8001)
├── frontend/                    Next.js website (port 3000)
├── static/                      runtime storage served at /static
│   ├── images/                  uploaded product images (the Image Library)
│   └── pdfs/                    generated datasheet PDFs
├── docs/                        end-user manual (PDF)
├── XSLX/                        sample supplier/uniform workbooks
├── PDF/                         sample supplier PDFs (used for table-extraction tests)
├── outputs/                     sample generated workbook output
├── spreadsheet-analysis/        table-extraction test inputs/results and small inspection scripts
├── scratch/                     throw-away helper script
└── (helper scripts and launch logs in the root — see below)
```

### `backend/`
| Path | What it is |
|---|---|
| `package.json`, `tsconfig.json`, `nest-cli.json` | Dependencies, scripts (`start:dev`, `build`, `start:prod`, `prisma:generate`, `prisma:migrate` = `prisma db push`) and Nest/TypeScript config |
| `prisma/schema.prisma`, `prisma/dev.db` | Database schema and the SQLite file |
| `src/main.ts` | Bootstraps Nest: CORS, 15 MB JSON limit, serves `/static`, listens on `PORT` (8001) |
| `src/app.module.ts` | Root module wiring all feature modules |
| `src/job/job.controller.ts` | Routes for jobs, preview, PDF list/download, template |
| `src/job/job.service.ts` | Workbook parsing (header rows → roles), duplicate validation/cleanup, job creation and the background PDF worker, ZIP, preview cache |
| `src/job/pdf-compiler.service.ts` | Builds the datasheet HTML/CSS: cover page, FEATURES panel, tables (key-value and multi-column), image frames, table-cell images, header layouts |
| `src/extract/table-extraction.controller.ts` | Routes for manual and bulk table extraction and result workbooks |
| `src/extract/table-extraction.service.ts` | PDF-text table reader, OCR pipeline (colour-band detection, per-row re-read, glyph repair), bulk grouping, workbook builders |
| `src/formatter/xlsx-formatter.*` | Supplier workbook → uniform workbook (attribute triples grouped into JSON columns; *Uniform Data*, *Source Backup*, *Mapping Guide* sheets) |
| `src/image/image.*` | Image upload/list/rename/delete/download and indexing of `static/images` |
| `src/storage/storage.*` | Creates and resolves `static/images` and `static/pdfs`, builds public URLs from `HOST_URL` |
| `src/prisma/prisma.*` | Prisma client provider |

### `frontend/`
| Path | What it is |
|---|---|
| `package.json`, `tsconfig.json`, `next-env.d.ts`, `postcss.config.js`, `tailwind.config.js` | Next.js / TypeScript / Tailwind configuration (brand colour `#CC0000`) |
| `src/app/layout.tsx` | Root layout: fonts, icon fonts, global styles, page title |
| `src/app/page.tsx` | The whole single-page UI: the four tools (Create PDFs, Image Library, Table JSON, XLSX Formatter), all API calls, crop selector, previews and inline styles |
| `src/app/globals.css`, `src/app/premium-ui.css` | Base styles and the layered visual theme (cards, drop areas, buttons, crop panel, tables) |
| `serve-built.js` | Tiny static server for the pre-built frontend (see section 2) |

### Root files
| File | What it is |
|---|---|
| `package.json`, `package-lock.json` | Root scripts + `concurrently` |
| `run-localhost.cmd` | Windows launcher: runs `npm run dev`, logging to `codex-persistent.*.log` |
| `setup_project.py` | One-shot installer: npm installs for root/backend/frontend + `prisma generate` + `prisma db push` |
| `logo.jpg`, `logo_b64.txt` | RS PRO logo and its base64 form (the PDF header logo is read from `logo.jpg`) |
| `compiled_debug.html` | Saved sample of compiled datasheet HTML, for debugging layouts |
| `build_page.py`, `get_logo.py` | Early helper scripts that generated the first version of the frontend page / the logo base64 string |
| `test_pdf_build.py`, `test_pdf_build_batteries.py`, `test_pdf_build_speciality.py` | Small scripts that upload a sample workbook to `/api/job/upload` and poll the job (smoke tests) |
| `check_xlsx_headers.py`, `inspect_pdf.py`, `scratch/read_row.py` | One-off inspection helpers (print workbook headers / PDF metadata / a row) |
| `spreadsheet-analysis/*.mjs`, `*.json`, `*.xlsx` | Inputs and outputs from table-extraction experiments (flat-head, bulk, manual crop) |
| `codex-job-survival.txt`, `shell-execute-test.txt` | Leftovers of process-launch tests; safe to delete |

> The helper scripts (`test_pdf_build*.py`, `check_xlsx_headers.py`, `inspect_pdf.py`, `scratch/read_row.py`, `spreadsheet-analysis/*.mjs`) contain **hard-coded Windows paths** from the original author's machine; edit the paths before running them. They are not needed to run the application.

### Not in the repository (ignored)
`node_modules/`, `.next/`, `dist/`, `*.log`, `*.tsbuildinfo`, `tmp/`, `.ocr-cache/`, `.env*` — see `.gitignore`. They are recreated by installing and building.

---

## 9. Deploying (checklist)

1. Install Node 18.17+ on the server; on Linux also install the system libraries Chromium needs for Puppeteer.
2. Set `NEXT_PUBLIC_BACKEND_URL`, `HOST_URL` and (optionally) `PORT` to the public addresses, then `npm run install:all`, Prisma generate/db push, `npm run build:all`.
3. Run backend (`node dist/main`) and frontend (`npm start`) under a process manager (PM2, systemd, …) behind a reverse proxy with HTTPS.
4. **Persist** `static/` (images and PDFs) and `backend/prisma/dev.db`; back them up regularly.
5. Add authentication / network restrictions — the app has none.
6. Be aware that the image library and `static/pdfs` are shared by all users: uploading an existing image name or generating a PDF for an existing stock number **overwrites** the earlier file.

---

## 10. Troubleshooting (developer)

| Symptom | Check |
|---|---|
| Browser shows empty lists / "Failed to fetch" | Backend not running on `NEXT_PUBLIC_BACKEND_URL` (default `:8001`); look for errors in the backend terminal |
| Image links point to `localhost` after deploying | `HOST_URL` (backend) and `NEXT_PUBLIC_BACKEND_URL` (frontend build) not set to the public address |
| PDF generation hangs or fails | Puppeteer could not start Chrome (missing system libraries) — see the backend log |
| OCR slow on first run | Tesseract language data is cached in `backend/.ocr-cache` after the first use |
| Frontend shows old backend URL | `NEXT_PUBLIC_*` values are baked in at build time — rebuild the frontend |

---

## License
Apache-2.0 — see [`LICENSE`](LICENSE).
