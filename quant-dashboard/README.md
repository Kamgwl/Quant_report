# Quant Strategy Dashboard

FY 2026–27 Q1 performance report (April – June 2026), built with **Vite + React + Recharts**.

## Run it

```bash
npm install        # first time only
npm run dev        # dev server with hot reload  → http://localhost:8445
```

For a production-style serve:

```bash
npm run build      # outputs to dist/
npm run preview    # serves the built app        → http://localhost:8445
```

Both servers bind `0.0.0.0:8445`, so the dashboard is reachable from other
machines on the network at `http://<this-machine-ip>:8445` (find the IP with
`ipconfig`). Open port 8445 in Windows Firewall if remote machines can't reach it.

## Project structure

```
quant-dashboard/
├─ index.html          # entry HTML (fonts, favicon, mount point)
├─ vite.config.js      # dev/preview on 0.0.0.0:8445
├─ public/favicon.svg
└─ src/
   ├─ main.jsx         # React root
   ├─ App.jsx          # light/dark theme provider + toggle
   ├─ Dashboard.jsx    # all views, charts and tables
   ├─ data.js          # account data extracted from the Excel sheet
   └─ theme.css        # CSS-variable design tokens (dark + light)
```

## Theming

A floating button (top-right) toggles **light / dark**. The choice is saved to
`localStorage`. Every color is a CSS variable defined in `src/theme.css`, so the
charts re-theme along with the rest of the UI.

## ROI logic

- Monthly ROI values come **straight from the Excel sheet** (no formula applied).
- **Q1 ROI** = April + May + June ROI (e.g. Ramakar Jha = 32.73%).
- The **Yearly (FY)** period sums whole quarters. Only Q1 exists today, so
  Yearly = Q1. When Q2–Q4 data arrives, extend `year`/`year_roi` in the
  `enrichRow` function in `src/Dashboard.jsx` (e.g. `q1 + q2 + q3 + q4`) and the
  Yearly view updates automatically. There is **no** ×4 annualized projection.
