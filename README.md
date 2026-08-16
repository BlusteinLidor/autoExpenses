# AutoExpenses

Personal expense automation: downloads card/bank exports, categorizes transactions with OpenAI, writes results to Excel, and provides a dashboard to review runs and view summaries.

The UI lives in **`frontend-next`** (Next.js). The **FastAPI** backend in `api.py` serves the API and, after a static build, hosts the exported dashboard at the same origin.

## Prerequisites

- **Python 3.10+**
- **Node.js 20+** (for `frontend-next`)
- **Google Chrome** (used by Playwright for Max/Leumi downloads)
- Accounts/API keys as needed (see [Environment variables](#environment-variables))

## 1. Backend setup

From the project root:

```bash
python -m venv .venv
```

**Windows (PowerShell):**

```powershell
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
playwright install chromium
```

**macOS / Linux:**

```bash
source .venv/bin/activate
pip install -r requirements.txt
playwright install chromium
```

Create a `.env` file in the project root (see below). The app loads it automatically via `config.load_env()`.

## 2. Frontend setup

```bash
cd frontend-next
npm install
```

## Environment variables

Create **`.env`** in the project root (not inside `frontend-next`):

| Variable | Required | Purpose |
|----------|----------|---------|
| `OPENAI_API_KEY` | Yes | Expense categorization |
| `MAX_USERNAME` | Yes (for Max download) | Max card login |
| `MAX_PASSWORD` | Yes (for Max download) | Max card login |
| `ID` | Yes (for Max download) | Max user ID |
| `LEUMI_USERNAME` | When using Leumi | Bank login |
| `LEUMI_PASSWORD` | When using Leumi | Bank login |
| `GOOGLE_CLIENT_ID` | Optional | Google Drive upload |
| `GOOGLE_CLIENT_SECRET` | Optional | Google Drive upload |
| `GOOGLE_OAUTH_REDIRECT_URI` | Optional | e.g. `http://localhost:8000/drive/connect/callback` |
| `DASHBOARD_BASE_URL` | Optional | Where OAuth redirects after Drive connect (default `http://localhost:3000`) |

Example:

```env
OPENAI_API_KEY=sk-...
MAX_USERNAME=your_max_user
MAX_PASSWORD=your_max_password
ID=your_max_id
```

For Google Drive, see [GOOGLE_DRIVE_SETUP.md](GOOGLE_DRIVE_SETUP.md).

## How to run

### Easy start (Windows)

Double-click **`Start AutoExpenses.bat`** in the project folder. It starts the server and opens the dashboard at http://127.0.0.1:8000. Close the window (or press Ctrl+C) to stop.

If you change the Next.js UI, run **`Rebuild Dashboard.bat`** once, then start again.

### Manual setups

There are two common setups.

#### Option A — Single server (recommended)

Build the Next app as a static export, then let FastAPI serve both the API and the UI on one port. API calls use relative URLs (no extra frontend env file).

**Terminal 1 — build the UI (once, or after frontend changes):**

```bash
cd frontend-next
npm run build
```

This writes output to `frontend-next/out/`.

**Terminal 2 — start the API:**

From the project root (with the virtualenv activated):

```bash
uvicorn api:app --reload --host 127.0.0.1 --port 8000
```

Open the dashboard at **http://127.0.0.1:8000**.

After you change React/TypeScript code, run `npm run build` again in `frontend-next` and refresh the browser.

#### Option B — Split dev servers (UI hot reload)

Use this when you are actively editing the Next.js app.

**Terminal 1 — API:**

```bash
# from project root, venv activated
uvicorn api:app --reload --host 127.0.0.1 --port 8000
```

**Terminal 2 — Next dev server:**

Create `frontend-next/.env.local`:

```env
NEXT_PUBLIC_API_BASE_URL=http://127.0.0.1:8000
```

Then:

```bash
cd frontend-next
npm run dev
```

Open **http://localhost:3000**. Set `DASHBOARD_BASE_URL=http://localhost:3000` in the root `.env` if you use Google Drive OAuth during local dev.

> **Note:** The API and UI run on different origins in this mode. If the browser blocks API requests (CORS), use **Option A** or add CORS middleware to `api.py` for `http://localhost:3000`.

## Project layout

| Path | Role |
|------|------|
| `api.py` | FastAPI app: REST API + serves `frontend-next/out` when built |
| `frontend-next/` | Next.js dashboard (`output: "export"`) |
| `frontend/` | Legacy HTML UI (fallback if `frontend-next/out` is missing) |
| `data/` | Workbooks, exports, `state.json`, assets (created at runtime) |
| `pipeline.py` | Monthly expense pipeline |

## Production-style run (no reload)

```bash
cd frontend-next && npm run build
cd ..
uvicorn api:app --host 0.0.0.0 --port 8000
```

Or serve the static export with any static host and point the frontend at your API with `NEXT_PUBLIC_API_BASE_URL` at build time.

## Troubleshooting

- **Blank page at `/`** — Run `npm run build` in `frontend-next` so `frontend-next/out/index.html` exists.
- **Missing env errors when running a month** — Check root `.env` for `OPENAI_API_KEY` and Max credentials; add Leumi vars if that step is enabled.
- **Playwright / browser errors** — Run `playwright install chromium` inside the activated venv.
- **Excel file locked** — Close `data/expenses_output.xlsx` (or the monthly file) before finalize writes.

## Further reading

- [GOOGLE_DRIVE_SETUP.md](GOOGLE_DRIVE_SETUP.md) — OAuth and Drive upload paths
