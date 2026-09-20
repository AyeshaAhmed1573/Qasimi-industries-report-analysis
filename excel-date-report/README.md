# Daily Sheet Date Report

A MERN-stack tool: upload a multi-sheet Excel workbook (like a "Daily Balance
Sheet for Yarn" file), pick a date, and download an Excel file listing every
matching entry across **all** sheets.

Output columns: **Date | Sheet Name | Receiver/Issuer Name | Type (Receive/Issue) | Quantity | Net Weight (kg)**

## Why "Type" was added

Your sheets track both a Receive quantity and an Issue quantity on the same
row. Without knowing which one actually has a value on a given row,
"Quantity" and "Net Weight" are ambiguous — so each matched entry is tagged
Receive or Issue.

## How it handles messy/inconsistent sheets

This workbook has 175+ sheets and the column headers aren't 100% consistent
(some call it "RECEIVER NAME", some "RECEIVER/ISSUE NAME", some use
"ISSUE / RECEIVE" for the party name, some reuse "TOTAL NET WEIGHT / KG" for
both directions). The backend re-reads each sheet's own header row and
fuzzy-matches column names rather than assuming a fixed layout — see
`backend/lib/parser.js` for the matching rules. Sheets it can't parse at all
(no recognizable DATE column) are skipped and listed in a "warnings" panel
in the UI, so nothing silently vanishes.

## Requirements

- Node.js 18+
- npm

## Setup

Open two terminals.

**Terminal 1 — backend:**
```bash
cd backend
npm install
npm start
# -> Backend listening on http://localhost:4000
```

**Terminal 2 — frontend:**
```bash
cd frontend
npm install
npm run dev
# -> Local: http://localhost:5173
```

Open `http://localhost:5173` in your browser. The frontend dev server proxies
`/api/*` requests to the backend automatically (see `vite.config.js`), so no
extra config is needed.

## Usage

1. Drag your `.xlsx` workbook onto the upload zone (or click to browse).
2. Pick a date.
3. Click **Generate Report** to preview matches in the table.
4. Click **Download Excel** to save the filtered report.

## Notes / limitations

- The uploaded workbook is held in the backend's memory for up to 2 hours
  per session, then discarded — nothing is persisted to disk or a database.
  This is fine for local/single-user use; if you need multi-user history,
  add MongoDB to save past reports (not included, since it wasn't needed for
  the core feature).
- If a sheet doesn't have a recognizable "Receiver/Issuer" name column, the
  entry still shows up with `(not found)` in that field rather than being
  dropped — check the "warnings" panel to see which sheets need a closer
  look.
- Max upload size is 60MB (adjustable in `backend/server.js`).

## Deploying (backend → Render, frontend → Vercel)

### Backend on Render

1. Push this project to a GitHub repo.
2. In Render, **New → Blueprint**, point it at the repo — it will pick up
   `render.yaml` at the project root automatically (root dir `backend`,
   build `npm install`, start `npm start`, health check `/api/health`).
   If you'd rather create the service manually instead of via Blueprint:
   root directory `backend`, build command `npm install`, start command
   `npm start`, health check path `/api/health`.
3. Once deployed, note the backend's URL, e.g.
   `https://excel-date-report-backend.onrender.com`.
4. In the Render service's **Environment** tab, set `CORS_ORIGIN` to your
   Vercel frontend URL once you have it (step below) — comma-separate if you
   need more than one (e.g. a preview URL and the production URL).

### Frontend on Vercel

1. In Vercel, **New Project**, import the same repo, and set **Root
   Directory** to `frontend` (Vercel auto-detects the Vite build).
2. Add an environment variable `VITE_API_URL` set to your Render backend URL
   from above, with **no trailing slash**, e.g.
   `https://excel-date-report-backend.onrender.com`.
3. Deploy. Once it's live, copy the Vercel URL back into the backend's
   `CORS_ORIGIN` env var on Render (step 4 above) and redeploy the backend
   so it accepts requests from the frontend's actual domain.

### About the backend "sleeping" on Render's free plan

Render's free web services spin down after ~15 minutes of no traffic, and
the *next* request wakes it back up with a 20–50 second cold start. A
`/api/health` endpoint is included and wired into `render.yaml` as the
health check path — but that's for Render's own deploy checks, **it does
not by itself keep the service warm**.

To actually stop the cold starts, have something ping the health endpoint
on a schedule shorter than the idle timeout:

- **UptimeRobot** (free) — add an HTTP(s) monitor for
  `https://your-backend.onrender.com/api/health`, checking every 5 minutes.
- **cron-job.org** (free) — same idea, schedule a GET request every
  5–10 minutes.

Either keeps the free instance from sleeping. (This only helps for the free
plan — a paid Render instance doesn't spin down at all.)

