import { useState, useMemo, useEffect, useRef } from 'react'
import {
  BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer,
  LineChart, Line, Cell, ReferenceLine, Legend,
} from 'recharts'
import ExcelJS from 'exceljs'
import { saveAs } from 'file-saver'
import { jsPDF } from 'jspdf'
import autoTable from 'jspdf-autotable'
import html2canvas from 'html2canvas'
import { RAW } from './data'


// ── CASH STRATEGY DETECTION ────────────────────────────────────────────────────
const CASH_STRATEGIES = ["CASH", "ETF", "NA"];
function isCashStrategy(r) {
  const seg = (r.segment || "").trim().toUpperCase();
  const s = (r.strategy || "").trim().toUpperCase();
  return seg === "CASH" || seg === "ETF" || 
         CASH_STRATEGIES.includes(s) ||
         r.name.toLowerCase().includes("cash") ||
         r.name.toLowerCase().includes("etf");
}

// ── COMPUTED DATA ─────────────────────────────────────────────────────────────
// Annualized ROI = Q1 ROI x 4 (3 months elapsed → x4 to project full year)
const MONTHS_ELAPSED = 3;
const ANN_FACTOR = 12 / MONTHS_ELAPSED; // 4

function enrichRow(r) {
const apr_roi = r.apr_roi / 100;
const may_roi = r.may_roi / 100;
const jun_roi = r.jun_roi / 100;
const jul_roi = (r.jul_roi || 0) / 100;
const aug_roi = (r.aug_roi || 0) / 100;
const sep_roi = (r.sep_roi || 0) / 100;
const jul = r.jul || 0, aug = r.aug || 0, sep = r.sep || 0;
const q1 = r.apr + r.may + r.jun;
const q1_roi = apr_roi + may_roi + jun_roi;
const q2 = jul + aug + sep;
const q2_roi = jul_roi + aug_roi + sep_roi;
// Yearly (FY) = sum of all quarters available so far (Q1 + Q2). Add q3/q4 here
// when those sheets arrive.
const year = q1 + q2;
const year_roi = q1_roi + q2_roi;
const ann_roi = q1_roi * ANN_FACTOR;   // annualized off the completed Q1 run-rate
const ann_pnl = q1 * ANN_FACTOR;
const strategy = (r.strategy || "").trim().toUpperCase();
return { ...r, strategy, apr_roi, may_roi, jun_roi, jul, aug, sep, jul_roi, aug_roi, sep_roi,
  q1, q1_roi, q2, q2_roi, year, year_roi, ann_roi, ann_pnl, isCash: isCashStrategy({ ...r, strategy }) };
}

const DATA_ALL = RAW.map(enrichRow);
const DATA_CASH = DATA_ALL.filter(r => r.isCash);
const DATA = DATA_ALL.filter(r => !r.isCash);

// Month configs
const MONTHS = [
{ key: "apr", label: "April 2026", roi_key: "apr_roi", color: "#4ade80" },
{ key: "may", label: "May 2026", roi_key: "may_roi", color: "#facc15" },
{ key: "jun", label: "June 2026", roi_key: "jun_roi", color: "#38bdf8" },
];

// Every FY month we hold data for, in calendar order. `cal` is the JS month
// index (0=Jan) used to pick the "current month" KPI tile. Extend with Oct–Mar
// when those sheets arrive.
const FY_MONTHS = [
{ key: "apr", label: "April",     short: "Apr", cal: 3, q: "q1" },
{ key: "may", label: "May",       short: "May", cal: 4, q: "q1" },
{ key: "jun", label: "June",      short: "Jun", cal: 5, q: "q1" },
{ key: "jul", label: "July",      short: "Jul", cal: 6, q: "q2" },
{ key: "aug", label: "August",    short: "Aug", cal: 7, q: "q2" },
{ key: "sep", label: "September", short: "Sep", cal: 8, q: "q2" },
];

// Quarter tiles in the KPI row. Clicking one opens the monthly-split popup.
// "fy" is the roll-up and spans every month above.
const QUARTERS = [
{ id: "q1", label: "Q1", span: "Apr–Jun", months: ["apr", "may", "jun"] },
{ id: "q2", label: "Q2", span: "Jul–Sep", months: ["jul", "aug", "sep"] },
];
const QUARTER_BY_ID = { ...Object.fromEntries(QUARTERS.map(q => [q.id, q])),
fy: { id: "fy", label: "FY (YTD)", span: "Apr–Sep", months: FY_MONTHS.map(m => m.key) } };

const fmt = (n) => {
if (Math.abs(n) >= 1e7) return `₹${(n/1e7).toFixed(2)}Cr`;
if (Math.abs(n) >= 1e5) return `₹${(n/1e5).toFixed(2)}L`;
if (Math.abs(n) >= 1e3) return `₹${(n/1e3).toFixed(1)}K`;
return `₹${n.toLocaleString("en-IN")}`;
};
// Fund is in Crores: display as ₹X.XXCr or ₹XL (if < 1Cr)
const fmtFund = (cr) => cr >= 1 ? `₹${cr.toFixed(2)}Cr` : `₹${(cr*100).toFixed(0)}L`;
const fmtROI = (n) => `${(n*100).toFixed(2)}%`;
const fmtSign = (n) => n >= 0 ? `+${fmt(n)}` : fmt(n);

// ── COLORS ─────────────────────────────────────────────────────────────────────
const ACCENT = "var(--accent)";
const POS = "var(--pos)";
const NEG = "var(--neg)";
const BG = "var(--bg)";
const CARD = "var(--card)";
const BORDER = "var(--border)";

// ── COMPONENTS ────────────────────────────────────────────────────────────────

// `onClick` turns the card into a toggle: it grows a hover lift and a caret that
// flips when `open`. `dropdown` is the panel that unfolds beneath the card — the
// quarter tiles use it to show their monthly split in place.
const StatCard = ({ label, value, sub, color, onClick, hint, open, dropdown }) => {
const ddRef = useRef(null);
// The panel hangs off the card's left edge, which overflows the viewport when
// the card sits in the last grid column. Right-align it in that case.
useEffect(() => {
  const el = ddRef.current;
  if (!open || !el) return;
  const fit = () => {
    el.style.left = "0px"; el.style.right = "auto";
    if (el.getBoundingClientRect().right > window.innerWidth - 12) {
      el.style.left = "auto"; el.style.right = "0px";
    }
  };
  fit();
  window.addEventListener("resize", fit);
  return () => window.removeEventListener("resize", fit);
}, [open]);
return (
<div onClick={onClick} role={onClick ? "button" : undefined} tabIndex={onClick ? 0 : undefined}
    aria-expanded={onClick ? !!open : undefined} data-kpi-card={onClick ? "1" : undefined}
    onKeyDown={onClick ? (e) => { if (e.key === "Enter" || e.key === " ") { e.preventDefault(); onClick(); } } : undefined}
    onMouseEnter={onClick ? (e) => { e.currentTarget.style.transform = "translateY(-2px)";
        e.currentTarget.style.boxShadow = "0 8px 20px rgba(0,0,0,0.18)"; } : undefined}
    onMouseLeave={onClick ? (e) => { e.currentTarget.style.transform = "none";
        e.currentTarget.style.boxShadow = "none"; } : undefined}
    style={{ background: CARD, border: `1px solid ${open ? (color || ACCENT) : BORDER}`, borderRadius: 12,
    padding: "18px 22px" , borderTop: `3px solid ${color || ACCENT}`,
    cursor: onClick ? "pointer" : "default", position: "relative",
    transition: "transform 0.15s ease, box-shadow 0.15s ease, border-color 0.15s ease", outline: "none",
    zIndex: open ? 60 : "auto" }}>
    <div style={{ color: "var(--muted)" , fontSize: 11, fontWeight: 700, letterSpacing: 2, textTransform: "uppercase" ,
        marginBottom: 6, paddingRight: onClick ? 16 : 0 }}>{label}</div>
    <div style={{ color: color || "var(--text)" , fontSize: 24, fontWeight: 800, fontFamily: "'DM Mono', monospace" }}>{value}
    </div>
    {sub && <div style={{ color: "var(--muted2)" , fontSize: 12, marginTop: 4 }}>{sub}</div>}
    {onClick && (
    <div className="no-print" style={{ position: "absolute", top: 13, right: 14,
        color: open ? (color || ACCENT) : "var(--muted2)", fontSize: 11, lineHeight: 1,
        transform: open ? "rotate(180deg)" : "none", transition: "transform 0.18s ease, color 0.15s ease" }}
        title={hint || "View monthly split"}>▼</div>
    )}
    {onClick && hint && (
    <div className="no-print" style={{ color: color || ACCENT, fontSize: 10, fontWeight: 700, letterSpacing: 0.6,
        marginTop: 8, opacity: 0.85, textTransform: "uppercase" }}>{open ? "Hide months" : hint}</div>
    )}
    {open && dropdown && (
    <div className="no-print" ref={ddRef} onClick={(e) => e.stopPropagation()}
        style={{ position: "absolute", top: "calc(100% + 8px)", left: 0, minWidth: 320, width: "max-content",
        maxWidth: "min(420px, 78vw)", background: CARD, border: `1px solid ${color || ACCENT}`, borderRadius: 12,
        boxShadow: "0 18px 38px rgba(0,0,0,0.35)", padding: "14px 16px", zIndex: 70, cursor: "default",
        transformOrigin: "top left", animation: "kpiDropOpen 0.16s ease-out" }}>
        <style>{`
          @keyframes kpiDropOpen {
            from { opacity: 0; transform: translateY(-6px); }
            to { opacity: 1; transform: translateY(0); }
          }
        `}</style>
        {dropdown}
    </div>
    )}
</div>
);
};

// Monthly split shown inside a quarter card's dropdown. `detail` comes from the
// `quarterDetail` memo; month ROI uses total AUM, matching the KPI cards.
const QuarterMonths = ({ detail, totalFund }) => (
<div>
    <div style={{ display: "flex", justifyContent: "space-between", alignItems: "baseline", gap: 10,
        paddingBottom: 10, marginBottom: 10, borderBottom: `1px solid ${BORDER}` }}>
        <div style={{ color: "var(--muted)", fontSize: 10, fontWeight: 700, letterSpacing: 1.2,
            textTransform: "uppercase" }}>{detail.q.label} · {detail.q.span}</div>
        <div style={{ color: detail.pnl >= 0 ? POS : NEG, fontSize: 13, fontWeight: 800,
            fontFamily: "'DM Mono', monospace" }}>{fmtSign(detail.pnl)}
            <span style={{ color: "var(--muted2)", fontSize: 10, fontWeight: 600, marginLeft: 6 }}>
                {fmtROI(detail.roi)}</span></div>
    </div>

    <div style={{ display: "flex", flexDirection: "column", gap: 9 }}>
        {detail.rows.map(m => {
        const c = m.pnl >= 0 ? POS : NEG;
        const pct = Math.round((Math.abs(m.pnl) / detail.peak) * 100);
        return (
        <div key={m.key} style={{ opacity: m.hasData ? 1 : 0.55 }}>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "baseline", gap: 10 }}>
                <div style={{ display: "flex", alignItems: "baseline", gap: 6 }}>
                    <span style={{ color: "var(--text)", fontSize: 12, fontWeight: 700 }}>{m.label}</span>
                    {m.hasData && detail.best && m.key === detail.best.key && (
                        <span style={{ fontSize: 8, fontWeight: 800, letterSpacing: 0.6, color: POS,
                            border: `1px solid ${POS}`, borderRadius: 3, padding: "0 4px" }}>BEST</span>)}
                    {m.hasData && detail.worst && m.key === detail.worst.key
                        && detail.best && detail.worst.key !== detail.best.key && (
                        <span style={{ fontSize: 8, fontWeight: 800, letterSpacing: 0.6, color: NEG,
                            border: `1px solid ${NEG}`, borderRadius: 3, padding: "0 4px" }}>WORST</span>)}
                    {!m.hasData && (
                        <span style={{ fontSize: 8, fontWeight: 800, letterSpacing: 0.6, color: "var(--muted2)",
                            border: `1px solid ${BORDER}`, borderRadius: 3, padding: "0 4px" }}>PENDING</span>)}
                </div>
                <div style={{ display: "flex", alignItems: "baseline", gap: 8, whiteSpace: "nowrap" }}>
                    <span style={{ color: m.hasData ? c : "var(--muted2)", fontSize: 13, fontWeight: 800,
                        fontFamily: "'DM Mono', monospace" }}>{m.hasData ? fmtSign(m.pnl) : "—"}</span>
                    {m.hasData && (
                    <span style={{ color: "var(--muted2)", fontSize: 10 }}>{fmtROI(m.roi)}</span>
                    )}
                </div>
            </div>
            <div style={{ height: 5, background: BORDER, borderRadius: 3, marginTop: 6, overflow: "hidden" }}>
                <div style={{ width: `${pct}%`, height: "100%", background: c, borderRadius: 3 }} />
            </div>
            {m.hasData && (
            <div style={{ color: "var(--muted2)", fontSize: 10, marginTop: 4 }}>
                {m.winners} winners · {m.losers} losers</div>
            )}
        </div>
        );
        })}
    </div>

    <div style={{ color: "var(--muted2)", fontSize: 10, marginTop: 11, paddingTop: 9,
        borderTop: `1px solid ${BORDER}` }}>
        ROI on {fmtFund(totalFund)} AUM
        {detail.pending.length > 0 && ` · ${detail.pending.map(m => m.short).join(", ")} not booked yet`}
    </div>
</div>
);

// Small per-sheet "Export to Excel" button (used at the top of each tab)
const ExportBtn = ({ onClick, label }) => (
  <button onClick={onClick} title="Download this sheet as a styled Excel file"
    style={{ display: "inline-flex", alignItems: "center", gap: 6,
      background: "rgba(0, 229, 255, 0.1)", border: `1px solid ${ACCENT}`, color: ACCENT,
      borderRadius: 8, padding: "7px 14px", cursor: "pointer", fontSize: 12, fontWeight: 700,
      whiteSpace: "nowrap", outline: "none" }}
    onMouseEnter={(e) => { e.currentTarget.style.background = ACCENT; e.currentTarget.style.color = "#000"; }}
    onMouseLeave={(e) => { e.currentTarget.style.background = "rgba(0, 229, 255, 0.1)"; e.currentTarget.style.color = ACCENT; }}>
    <span style={{ fontSize: 13 }}>📥</span> {label || "Export to Excel"}
  </button>
);

const CustomTooltip = ({ active, payload, label }) => {
if (!active || !payload?.length) return null;
return (
<div style={{ background: "var(--tooltip-bg)" , border: `1px solid ${BORDER}`, borderRadius: 8, padding: "10px 14px" , fontSize:
    12, color: "var(--axis)" }}>
    <div style={{ color: ACCENT, fontWeight: 700, marginBottom: 4 }}>{label}</div>
    {payload.map((p, i) => (
    <div key={i} style={{ color: p.value>= 0 ? POS : NEG }}>
        {p.name}: {p.name.includes("ROI") ? fmtROI(p.value) : fmt(p.value)}
    </div>
    ))}
</div>
);
};

// ── MAIN ──────────────────────────────────────────────────────────────────────
const EXPORT_COLUMNS = [
  { key: "code", label: "Account Code", align: "left" },
  { key: "name", label: "Account Name", align: "left" },
  { key: "strategy", label: "Strategy", align: "left" },
  { key: "fund", label: "Fund (Cr)", align: "right", numFmt: "₹#,##0.00 \"Cr\"" },
  { key: "apr", label: "April P&L", align: "right", numFmt: "+₹#,##,##0;-₹#,##,##0;\"—\"", colorCode: true },
  { key: "apr_roi", label: "April ROI", align: "right", numFmt: "+0.00%;-0.00%;0.00%", colorCode: true },
  { key: "may", label: "May P&L", align: "right", numFmt: "+₹#,##,##0;-₹#,##,##0;\"—\"", colorCode: true },
  { key: "may_roi", label: "May ROI", align: "right", numFmt: "+0.00%;-0.00%;0.00%", colorCode: true },
  { key: "jun", label: "June P&L", align: "right", numFmt: "+₹#,##,##0;-₹#,##,##0;\"—\"", colorCode: true },
  { key: "jun_roi", label: "June ROI", align: "right", numFmt: "+0.00%;-0.00%;0.00%", colorCode: true },
  { key: "q1", label: "Q1 Net P&L", align: "right", numFmt: "+₹#,##,##0;-₹#,##,##0;\"—\"", colorCode: true },
  { key: "q1_roi", label: "Q1 Net ROI", align: "right", numFmt: "+0.00%;-0.00%;0.00%", colorCode: true },
  { key: "jul", label: "July P&L", align: "right", numFmt: "+₹#,##,##0;-₹#,##,##0;\"—\"", colorCode: true },
  { key: "jul_roi", label: "July ROI", align: "right", numFmt: "+0.00%;-0.00%;0.00%", colorCode: true },
  { key: "q2", label: "Q2 Net P&L", align: "right", numFmt: "+₹#,##,##0;-₹#,##,##0;\"—\"", colorCode: true },
  { key: "q2_roi", label: "Q2 Net ROI", align: "right", numFmt: "+0.00%;-0.00%;0.00%", colorCode: true },
  { key: "year", label: "FY P&L (YTD)", align: "right", numFmt: "+₹#,##,##0;-₹#,##,##0;\"—\"", colorCode: true },
  { key: "year_roi", label: "FY ROI (YTD)", align: "right", numFmt: "+0.00%;-0.00%;0.00%", colorCode: true },
  { key: "ann_pnl", label: "Ann. P&L (Projected)", align: "right", numFmt: "+₹#,##,##0;-₹#,##,##0;\"—\"", colorCode: true },
  { key: "ann_roi", label: "Ann. ROI", align: "right", numFmt: "+0.00%;-0.00%;0.00%", colorCode: true },
];

// ── PER-SHEET EXPORT COLUMN SETS ────────────────────────────────────────────
const PNL_FMT  = "+₹#,##,##0;-₹#,##,##0;\"—\"";
const ROI_FMT  = "+0.00%;-0.00%;0.00%";
const FUND_FMT = "₹#,##0.00 \"Cr\"";

// Full per-account breakdown (used by User Performance, Quarter Summary, Cash)
const COLS_ACCOUNT_FULL = [
  { key: "code", label: "Code", align: "left" },
  { key: "name", label: "Name", align: "left" },
  { key: "strategy", label: "Strategy", align: "left" },
  { key: "fund", label: "Fund (Cr)", align: "right", numFmt: FUND_FMT },
  { key: "apr", label: "Apr P&L", align: "right", numFmt: PNL_FMT, colorCode: true },
  { key: "apr_roi", label: "Apr ROI", align: "right", numFmt: ROI_FMT, colorCode: true },
  { key: "may", label: "May P&L", align: "right", numFmt: PNL_FMT, colorCode: true },
  { key: "may_roi", label: "May ROI", align: "right", numFmt: ROI_FMT, colorCode: true },
  { key: "jun", label: "Jun P&L", align: "right", numFmt: PNL_FMT, colorCode: true },
  { key: "jun_roi", label: "Jun ROI", align: "right", numFmt: ROI_FMT, colorCode: true },
  { key: "q1", label: "Q1 P&L", align: "right", numFmt: PNL_FMT, colorCode: true },
  { key: "q1_roi", label: "Q1 ROI", align: "right", numFmt: ROI_FMT, colorCode: true },
  { key: "jul", label: "Jul P&L", align: "right", numFmt: PNL_FMT, colorCode: true },
  { key: "jul_roi", label: "Jul ROI", align: "right", numFmt: ROI_FMT, colorCode: true },
  { key: "aug", label: "Aug P&L", align: "right", numFmt: PNL_FMT, colorCode: true },
  { key: "aug_roi", label: "Aug ROI", align: "right", numFmt: ROI_FMT, colorCode: true },
  { key: "sep", label: "Sep P&L", align: "right", numFmt: PNL_FMT, colorCode: true },
  { key: "sep_roi", label: "Sep ROI", align: "right", numFmt: ROI_FMT, colorCode: true },
  { key: "q2", label: "Q2 P&L", align: "right", numFmt: PNL_FMT, colorCode: true },
  { key: "q2_roi", label: "Q2 ROI", align: "right", numFmt: ROI_FMT, colorCode: true },
  { key: "year", label: "FY P&L", align: "right", numFmt: PNL_FMT, colorCode: true },
  { key: "year_roi", label: "FY ROI", align: "right", numFmt: ROI_FMT, colorCode: true },
];

// Strategy aggregate table (roi here is already a ×100 percentage number)
const COLS_STRATEGY = [
  { key: "strategy", label: "Strategy", align: "left" },
  { key: "count", label: "Accounts", align: "right", numFmt: "0" },
  { key: "fund", label: "Fund (Cr)", align: "right", numFmt: "#,##0.00" },
  { key: "pnl", label: "Q1 Net P&L", align: "right", numFmt: PNL_FMT, colorCode: true },
  { key: "roi", label: "Q1 ROI %", align: "right", numFmt: "0.00\"%\"", colorCode: true },
];

const MONTH_LABEL = { apr:"April", may:"May", jun:"June", q1:"Q1", jul:"July", aug:"Aug", sep:"Sep", q2:"Q2", year:"FY" };

function Dashboard() {
const [theme, setTheme] = useState(() => localStorage.getItem("qd-theme") || "dark");

useEffect(() => {
  document.documentElement.setAttribute("data-theme", theme);
  localStorage.setItem("qd-theme", theme);
}, [theme]);

const MUTED_C  = theme === 'dark' ? '#6b8cbb' : '#56688a';
const AXIS_C   = theme === 'dark' ? '#b0c4de' : '#3a4a63';
const MUTED2_C = theme === 'dark' ? '#4a6490' : '#8493ad';
const TIP_BG   = theme === 'dark' ? '#111d33' : '#ffffff';
const TIP_BORDER = theme === 'dark' ? '#1e2d4a' : '#d5deec';

const [view, setView] = useState("monthly"); // monthly | quarterly | user | strategy
const [selectedMonth, setSelectedMonth] = useState("q1");
const [sortBy, setSortBy] = useState("roi");

const [showExportModal, setShowExportModal] = useState(false);
const [exportScope, setExportScope] = useState("all"); // 'all' | 'fo' | 'cash'

// Quarter drill-down dropdown: 'q1' | 'q2' | 'fy' | null. Clicking the open
// card toggles it shut; Esc or a click anywhere off the KPI cards also closes.
const [openQuarter, setOpenQuarter] = useState(null);
const toggleQuarter = (id) => setOpenQuarter(cur => (cur === id ? null : id));
useEffect(() => {
  if (!openQuarter) return;
  const onKey = (e) => { if (e.key === "Escape") setOpenQuarter(null); };
  const onDown = (e) => {
    if (e.target && e.target.closest && !e.target.closest("[data-kpi-card]")) setOpenQuarter(null);
  };
  window.addEventListener("keydown", onKey);
  document.addEventListener("mousedown", onDown);
  return () => { window.removeEventListener("keydown", onKey); document.removeEventListener("mousedown", onDown); };
}, [openQuarter]);

// ── Report / Print state ──
const [showPrintModal, setShowPrintModal] = useState(false);
const [printMode, setPrintMode] = useState(false);   // browser print path
const [reportMode, setReportMode] = useState(false); // off-screen render for PDF capture
const [reportBusy, setReportBusy] = useState(false);
const themeBeforeReport = useRef(null);
const [printSel, setPrintSel] = useState({
  summary: true, monthly: true, quarterly: true, user: true, strategy: true, cash: false,
});
// Browser-print path: render selected sections, fire print dialog, reset.
useEffect(() => {
  if (!printMode) return;
  const t = setTimeout(() => { window.print(); setPrintMode(false); }, 350);
  const after = () => setPrintMode(false);
  window.addEventListener("afterprint", after);
  return () => { clearTimeout(t); window.removeEventListener("afterprint", after); };
}, [printMode]);
// PDF-report path: once sections are rendered, capture charts + build the PDF.
useEffect(() => {
  if (!reportMode) return;
  let cancelled = false;
  const t = setTimeout(async () => {
    try { await buildPdfReport(); }
    catch (e) { console.error("Report failed", e); alert("Report generation failed: " + e.message); }
    finally {
      if (!cancelled) {
        setReportMode(false); setReportBusy(false);
        if (themeBeforeReport.current) setTheme(themeBeforeReport.current); // restore theme
      }
    }
  }, 650); // give Recharts + the light-theme switch time to lay out
  return () => { cancelled = true; clearTimeout(t); };
}, [reportMode]);
const startPrint = () => { setShowPrintModal(false); setPrintMode(true); };
const startReport = () => {
  themeBeforeReport.current = theme;
  setTheme("light");           // captured charts read cleanly on a light page
  setShowPrintModal(false); setReportBusy(true); setReportMode(true);
};
// A section renders when its tab is active, OR during print/report capture when ticked.
const showSection = (id) => (printMode || reportMode) ? printSel[id] : view === id;
const [selectedCols, setSelectedCols] = useState({
  code: true,
  name: true,
  strategy: true,
  fund: true,
  apr: true,
  apr_roi: true,
  may: true,
  may_roi: true,
  jun: true,
  jun_roi: true,
  q1: true,
  q1_roi: true,
  jul: true,
  jul_roi: true,
  q2: true,
  q2_roi: true,
  year: true,
  year_roi: true,
  ann_pnl: true,
  ann_roi: true,
});

// Generic styled-xlsx exporter — reused by the global modal AND every per-tab button.
const exportRowsToExcel = async (rows, columns, sheetLabel, fileName) => {
  try {
    if (!ExcelJS) { alert("ExcelJS library not loaded. Please check your internet connection."); return; }
    if (!columns || columns.length === 0) { alert("No columns selected to export."); return; }
    if (!rows || rows.length === 0) { alert("No data to export on this sheet."); return; }

    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet(String(sheetLabel).slice(0, 31));

    worksheet.columns = columns.map(col => ({
      header: col.label, key: col.key, width: 15,
      style: { font: { name: 'Segoe UI', size: 10 }, alignment: { vertical: 'middle', horizontal: col.align || 'left' } }
    }));

    const headerRow = worksheet.getRow(1);
    headerRow.height = 28;
    const isDark = theme === "dark";
    const headerBg = isDark ? "FF0D1526" : "FFF3F6FB";
    const headerFg = isDark ? "FF00E5FF" : "FF0091B3";
    const headerBorderColor = isDark ? "FF1E2D4A" : "FFD5DEEC";
    columns.forEach((colDef, idx) => {
      const cell = headerRow.getCell(idx + 1);
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: headerBg } };
      cell.font = { name: 'Segoe UI', size: 11, bold: true, color: { argb: headerFg } };
      cell.border = { bottom: { style: 'medium', color: { argb: headerBorderColor } } };
      cell.alignment = { vertical: 'middle', horizontal: colDef.align || 'left' };
    });

    rows.forEach((row) => {
      const rowData = {};
      columns.forEach(col => { rowData[col.key] = row[col.key] ?? null; });
      const addedRow = worksheet.addRow(rowData);
      addedRow.height = 20;
      columns.forEach((colDef, idx) => {
        const cell = addedRow.getCell(idx + 1);
        const val = row[colDef.key];
        cell.border = {
          top: { style: 'thin', color: { argb: 'FFE0E0E0' } },
          left: { style: 'thin', color: { argb: 'FFE0E0E0' } },
          bottom: { style: 'thin', color: { argb: 'FFE0E0E0' } },
          right: { style: 'thin', color: { argb: 'FFE0E0E0' } }
        };
        if (colDef.numFmt) cell.numFmt = colDef.numFmt;
        if (colDef.colorCode && typeof val === 'number') {
          if (val > 0) {
            cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE6F4EA' } };
            cell.font = { name: 'Segoe UI', size: 10, color: { argb: 'FF137333' }, bold: true };
          } else if (val < 0) {
            cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFCE8E6' } };
            cell.font = { name: 'Segoe UI', size: 10, color: { argb: 'FFC5221F' }, bold: true };
          }
        }
      });
    });

    columns.forEach((colDef, idx) => {
      const column = worksheet.getColumn(idx + 1);
      let maxLen = Math.max(12, String(colDef.label).length + 2);
      column.eachCell({ includeEmpty: true }, cell => {
        const valStr = (cell.value === null || cell.value === undefined) ? ''
          : (typeof cell.value === 'object' && cell.value.text ? cell.value.text.toString() : cell.value.toString());
        if (valStr.length + 6 > maxLen) maxLen = valStr.length + 6;
      });
      column.width = Math.min(maxLen, 30);
    });

    // Freeze header row for easier scrolling
    worksheet.views = [{ state: 'frozen', ySplit: 1 }];

    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
    saveAs(blob, fileName);
  } catch (error) {
    console.error("Export failed", error);
    alert("An error occurred during export: " + error.message);
  }
};

// Global modal export (scope + column picker)
const handleExportExcel = async () => {
  const exportData = exportScope === "all" ? DATA_ALL : exportScope === "cash" ? DATA_CASH : DATA;
  const activeCols = EXPORT_COLUMNS.filter(c => selectedCols[c.key]);
  await exportRowsToExcel(exportData, activeCols, "Quant FY2026-27", `Quant_FY2026-27_${exportScope.toUpperCase()}.xlsx`);
  setShowExportModal(false);
};

// ── Per-tab (per-sheet) exports ──
const exportUserSheet = () =>
  exportRowsToExcel(ranked, COLS_ACCOUNT_FULL, "User Performance", "Quant_UserPerformance_FY2026-27.xlsx");
const exportCashSheet = () =>
  exportRowsToExcel(cashRanked, COLS_ACCOUNT_FULL, "Cash ETF ATS", "Quant_Cash-ETF-ATS_FY2026-27.xlsx");
const exportStrategySheet = () =>
  exportRowsToExcel(strategyData, COLS_STRATEGY, "Strategy Analysis", "Quant_StrategyAnalysis_FY2026-27.xlsx");
const exportQuarterSheet = () =>
  exportRowsToExcel(ranked, COLS_ACCOUNT_FULL, "Quarter Summary", "Quant_QuarterSummary_FY2026-27.xlsx");
const exportMonthlySheet = () => {
  const key = selectedMonth, roiKey = key + "_roi";
  const rows = [...DATA].filter(r => Math.abs(r[key]) > 0)
    .sort((a, b) => (sortBy === "roi" ? b[roiKey] - a[roiKey] : b[key] - a[key]));
  const lbl = MONTH_LABEL[key] || key;
  const cols = [
    { key: "code", label: "Code", align: "left" },
    { key: "name", label: "Name", align: "left" },
    { key: "strategy", label: "Strategy", align: "left" },
    { key: "fund", label: "Fund (Cr)", align: "right", numFmt: FUND_FMT },
    { key, label: `${lbl} P&L`, align: "right", numFmt: PNL_FMT, colorCode: true },
    { key: roiKey, label: `${lbl} ROI`, align: "right", numFmt: ROI_FMT, colorCode: true },
  ];
  exportRowsToExcel(rows, cols, `Monthly ${lbl}`, `Quant_Monthly_${lbl}_FY2026-27.xlsx`);
};

// ── PDF REPORT BUILDER ──────────────────────────────────────────────────────
const buildPdfReport = async () => {
  const JsPDF = jsPDF;
  if (!JsPDF) { alert("PDF library not loaded — check your internet connection."); return; }

  const doc = new JsPDF({ orientation: "portrait", unit: "pt", format: "a4" });
  const W = doc.internal.pageSize.getWidth();
  const H = doc.internal.pageSize.getHeight();
  const M = 40, CONTENT_W = W - M * 2;
  const DARK = [13, 21, 38], CYAN = [0, 229, 255], GREEN = [22, 163, 74], RED = [225, 29, 72];
  let cursorY = 120;

  // ASCII-only formatters — jsPDF's standard fonts are Latin-1, so the ₹ glyph
  // (and any non-Latin char) corrupts the cell. Use "Rs" instead.
  const pdfNum = (v) => {
    const a = Math.abs(v);
    if (a >= 1e7) return (v / 1e7).toFixed(2) + " Cr";
    if (a >= 1e5) return (v / 1e5).toFixed(2) + " L";
    if (a >= 1e3) return (v / 1e3).toFixed(1) + " K";
    return String(Math.round(v));
  };
  const pnlStr = (v) => (v >= 0 ? "+Rs " : "-Rs ") + pdfNum(Math.abs(v));
  const roiStr = (v) => (v * 100).toFixed(2) + "%";
  const fundStr = (cr) => "Rs " + (cr >= 1 ? cr.toFixed(2) + " Cr" : (cr * 100).toFixed(0) + " L");
  const ensureSpace = (needed) => { if (cursorY + needed > H - M) { doc.addPage(); cursorY = M; } };
  let firstSection = true;
  const sectionTitle = (title) => {
    if (!firstSection) { doc.addPage(); cursorY = M; }  // each section starts on a new page
    firstSection = false;
    ensureSpace(46);
    doc.setDrawColor(0, 150, 179); doc.setLineWidth(3);
    doc.line(M, cursorY, M + 24, cursorY);
    doc.setFont("helvetica", "bold"); doc.setFontSize(13); doc.setTextColor(20, 30, 50);
    doc.text(title.toUpperCase(), M + 32, cursorY + 4);
    cursorY += 20;
  };
  const addTable = (head, body, colStyles) => {
    autoTable(doc, {
      startY: cursorY, margin: { left: M, right: M }, theme: "striped",
      head: [head], body,
      styles: { fontSize: 8, cellPadding: 3, overflow: "linebreak" },
      headStyles: { fillColor: DARK, textColor: CYAN, fontSize: 8 },
      alternateRowStyles: { fillColor: [244, 246, 251] },
      columnStyles: colStyles || {},
      didParseCell: (data) => {
        if (data.section === "body" && typeof data.cell.raw === "string") {
          const t = data.cell.raw;
          if (/^[+-]/.test(t)) data.cell.styles.textColor = t[0] === "-" ? RED : GREEN;
        }
      },
    });
    cursorY = doc.lastAutoTable.finalY + 22;
  };
  const addChart = async (id) => {
    const el = document.getElementById(id);
    if (!el || !html2canvas) return;
    const canvas = await html2canvas(el, { backgroundColor: "#ffffff", scale: 2, logging: false });
    const imgW = CONTENT_W, imgH = canvas.height * imgW / canvas.width;
    ensureSpace(imgH + 8);
    doc.addImage(canvas.toDataURL("image/png"), "PNG", M, cursorY, imgW, imgH);
    cursorY += imgH + 14;
  };
  // Rounded KPI card with an accent bar, label, big value and sub-line.
  const kpiCard = (x, y, w, h, label, value, sub, accent, valColor) => {
    doc.setFillColor(247, 249, 252); doc.setDrawColor(226, 232, 242); doc.setLineWidth(0.6);
    doc.roundedRect(x, y, w, h, 6, 6, "FD");
    doc.setFillColor(accent[0], accent[1], accent[2]); doc.roundedRect(x, y + 6, 4, h - 12, 2, 2, "F");
    doc.setFont("helvetica", "bold"); doc.setFontSize(7.5); doc.setTextColor(120, 133, 152);
    doc.text(label.toUpperCase(), x + 16, y + 17);
    doc.setFont("helvetica", "bold"); doc.setFontSize(15); doc.setTextColor(valColor[0], valColor[1], valColor[2]);
    doc.text(value, x + 16, y + 37);
    if (sub) { doc.setFont("helvetica", "normal"); doc.setFontSize(7.5); doc.setTextColor(130, 142, 160); doc.text(sub, x + 16, y + 51); }
  };
  // Leaderboard row: rank badge + name/strategy on the left, ROI + P&L on the right.
  const leaderRow = (x, y, w, rank, name, strat, roi, pnl, positive, badge) => {
    const h = 30, col = positive ? GREEN : RED;
    doc.setFillColor(247, 249, 252); doc.setDrawColor(226, 232, 242); doc.setLineWidth(0.5);
    doc.roundedRect(x, y, w, h, 5, 5, "FD");
    doc.setFillColor(badge[0], badge[1], badge[2]); doc.circle(x + 16, y + h / 2, 9, "F");
    doc.setFont("helvetica", "bold"); doc.setFontSize(9); doc.setTextColor(255, 255, 255);
    doc.text(String(rank), x + 16, y + h / 2 + 3, { align: "center" });
    doc.setFontSize(9.5); doc.setTextColor(20, 30, 50);
    doc.text(doc.splitTextToSize(name, w - 130)[0], x + 30, y + 13);
    doc.setFont("helvetica", "normal"); doc.setFontSize(7.5); doc.setTextColor(120, 133, 152);
    doc.text((strat || "-").slice(0, 22), x + 30, y + 23);
    doc.setFont("helvetica", "bold"); doc.setFontSize(10); doc.setTextColor(col[0], col[1], col[2]);
    doc.text(roi, x + w - 12, y + 13, { align: "right" });
    doc.setFont("helvetica", "normal"); doc.setFontSize(8); doc.setTextColor(90, 105, 125);
    doc.text(pnl, x + w - 12, y + 24, { align: "right" });
  };

  try {
    // Cover band
    doc.setFillColor(DARK[0], DARK[1], DARK[2]); doc.rect(0, 0, W, 96, "F");
    doc.setTextColor(0, 229, 255); doc.setFont("helvetica", "bold"); doc.setFontSize(20);
    doc.text("Quant Strategy Dashboard", M, 42);
    doc.setFont("helvetica", "normal"); doc.setFontSize(11); doc.setTextColor(210, 220, 235);
    doc.text("FY 2026-27   |   Q1 + Q2 Performance Report   |   April - September 2026", M, 64);
    doc.setFontSize(9); doc.setTextColor(150, 165, 185);
    doc.text("Generated " + new Date().toLocaleString("en-IN"), M, 82);

    if (printSel.summary) {
      sectionTitle("Executive Summary");
      const NAVY = [20, 30, 50], VIOLET = [124, 92, 255];
      const gr = (v) => (v >= 0 ? GREEN : RED);

      // ── Hero banner: headline FY Net P&L ──
      const heroH = 62, hero = gr(totals.fyPnL);
      doc.setFillColor(totals.fyPnL >= 0 ? 236 : 253, totals.fyPnL >= 0 ? 248 : 236, totals.fyPnL >= 0 ? 240 : 236);
      doc.roundedRect(M, cursorY, CONTENT_W, heroH, 8, 8, "F");
      doc.setFillColor(hero[0], hero[1], hero[2]); doc.roundedRect(M, cursorY, 6, heroH, 3, 3, "F");
      doc.setFont("helvetica", "bold"); doc.setFontSize(8); doc.setTextColor(90, 105, 125);
      doc.text("FY NET P&L  (YTD, Q1 + Q2)", M + 22, cursorY + 22);
      doc.setFontSize(24); doc.setTextColor(hero[0], hero[1], hero[2]);
      doc.text(pnlStr(totals.fyPnL), M + 22, cursorY + 48);
      doc.setFont("helvetica", "normal"); doc.setFontSize(10); doc.setTextColor(70, 84, 104);
      doc.text("ROI " + roiStr(totals.fyROI), W - M - 22, cursorY + 26, { align: "right" });
      doc.setFontSize(9); doc.setTextColor(110, 124, 144);
      doc.text(totals.count + " accounts  |  " + totals.winners + " winners / " + totals.losers + " losers",
        W - M - 22, cursorY + 44, { align: "right" });
      cursorY += heroH + 16;

      // ── KPI cards (2 x 2) ──
      const gapX = 16, gapY = 14, cardH = 62, cardW = (CONTENT_W - gapX) / 2;
      const cards = [
        ["Total AUM (Utilized)", fundStr(totals.totalFund), totals.count + " accounts", CYAN, NAVY],
        ["FY ROI (YTD)", roiStr(totals.fyROI), "Q1 + Q2 combined", CYAN, gr(totals.fyROI)],
        ["Q1 Net P&L (Apr-Jun)", pnlStr(totals.q1PnL), "ROI " + roiStr(totals.q1ROI), gr(totals.q1PnL), gr(totals.q1PnL)],
        ["Q2 Net P&L (Jul-Sep)", pnlStr(totals.q2PnL), "ROI " + roiStr(totals.q2ROI), gr(totals.q2PnL), gr(totals.q2PnL)],
      ];
      cards.forEach((c, i) => kpiCard(
        M + (i % 2) * (cardW + gapX), cursorY + Math.floor(i / 2) * (cardH + gapY),
        cardW, cardH, c[0], c[1], c[2], c[3], c[4]));
      cursorY += 2 * (cardH + gapY) + 8;
    }

    if (printSel.monthly) {
      sectionTitle("Monthly Performance");
      await addChart("rpt-monthly-pnl");
      await addChart("rpt-monthly-roi");
      const fund = totals.totalFund, mroi = (p) => fund > 0 ? p / (fund * 1e7) : 0;
      const grM = (v) => (v >= 0 ? GREEN : RED);
      const cH = 56, rowGap = 12;
      ensureSpace(50 + rowGap + 2 * (cH + rowGap) + 6);
      // Yearly P&L on top (full-width)
      kpiCard(M, cursorY, CONTENT_W, 50, "FY (YTD) Net P&L  -  Q1 + Q2",
        pnlStr(totals.fyPnL), "ROI " + roiStr(mroi(totals.fyPnL)), CYAN, grM(totals.fyPnL));
      cursorY += 50 + rowGap;
      // Q1 row, then Q2 row (July starts a new row)
      const gx = 10, cW = (CONTENT_W - 3 * gx) / 4;
      const drawRow = (arr, ry) => arr.forEach((m, i) => kpiCard(
        M + i * (cW + gx), ry, cW, cH, m[0], pnlStr(m[1]), "ROI " + roiStr(mroi(m[1])),
        m[2] ? CYAN : grM(m[1]), grM(m[1])));
      drawRow([["April", totals.aprPnL], ["May", totals.mayPnL], ["June", totals.junPnL], ["Q1 Total", totals.q1PnL, true]], cursorY);
      drawRow([["July", totals.julPnL], ["August", totals.augPnL], ["September", totals.sepPnL], ["Q2 Total", totals.q2PnL, true]], cursorY + cH + rowGap);
      cursorY += 2 * (cH + rowGap) + 4;
    }

    if (printSel.quarterly) {
      sectionTitle("Quarter Summary - Top / Bottom (Q1 ROI)");
      const colGap = 20, colW = (CONTENT_W - colGap) / 2;
      const leftX = M, rightX = M + colW + colGap;
      doc.setFont("helvetica", "bold"); doc.setFontSize(9);
      doc.setTextColor(GREEN[0], GREEN[1], GREEN[2]); doc.text("TOP 5 PERFORMERS", leftX, cursorY);
      doc.setTextColor(RED[0], RED[1], RED[2]); doc.text("BOTTOM 5 PERFORMERS", rightX, cursorY);
      cursorY += 8;
      const rowH = 30, rowGap = 8, gold = [212, 160, 23], slate = [100, 116, 139];
      top5.forEach((r, i) => leaderRow(leftX, cursorY + i * (rowH + rowGap), colW, i + 1,
        r.name, r.strategy || "-", roiStr(r.q1_roi), pnlStr(r.q1), true, i === 0 ? gold : slate));
      bot5.forEach((r, i) => leaderRow(rightX, cursorY + i * (rowH + rowGap), colW, i + 1,
        r.name, r.strategy || "-", roiStr(r.q1_roi), pnlStr(r.q1), false, RED));
      cursorY += 5 * (rowH + rowGap) + 6;
    }

    if (printSel.user) {
      sectionTitle("User Performance (Q1 / Q2 / FY)");
      addTable(["#", "Code", "Name", "Strategy", "Fund", "Q1 P&L", "Q1 ROI", "Q2 P&L", "Q2 ROI", "FY P&L", "FY ROI"],
        ranked.map((r, i) => [String(i + 1), r.code, r.name, r.strategy || "-", r.fund.toFixed(2),
          pnlStr(r.q1), roiStr(r.q1_roi), pnlStr(r.q2), roiStr(r.q2_roi), pnlStr(r.year), roiStr(r.year_roi)]),
        { 2: { cellWidth: 78 } });
    }

    if (printSel.strategy) {
      sectionTitle("Strategy Analysis");
      await addChart("rpt-strategy-pnl");
      doc.addPage(); cursorY = M;   // keep the full strategy table together on its own page
      addTable(["Strategy", "Accounts", "Fund (Cr)", "Q1 Net P&L", "Q1 ROI %"],
        strategyData.map(s => [s.strategy, String(s.count), s.fund.toFixed(2), pnlStr(s.pnl), s.roi.toFixed(2) + "%"]),
        { 0: { cellWidth: 130, fontStyle: "bold", textColor: [20, 30, 50] } });
    }

    if (printSel.cash) {
      sectionTitle("Cash / ETF / ATS");
      addTable(["#", "Code", "Name", "Strategy", "Fund", "Apr", "May", "Jun", "Q1 P&L", "Q1 ROI"],
        cashRanked.map((r, i) => [String(i + 1), r.code, r.name, r.strategy || "-", r.fund.toFixed(2),
          pnlStr(r.apr), pnlStr(r.may), pnlStr(r.jun), pnlStr(r.q1), roiStr(r.q1_roi)]),
        { 2: { cellWidth: 90 } });
    }

    // Footer page numbers
    const pages = doc.internal.getNumberOfPages();
    for (let p = 1; p <= pages; p++) {
      doc.setPage(p);
      doc.setFontSize(8); doc.setTextColor(150, 160, 175);
      doc.text("Quant Strategy FY2026-27  -  Confidential", M, H - 18);
      doc.text("Page " + p + " / " + pages, W - M - 55, H - 18);
    }
    doc.save("Quant_Dashboard_Report_FY2026-27.pdf");
  } finally { /* theme is restored by the report effect */ }
};

// ── Summary stats ──
const totals = useMemo(() => {
  const normalSection = DATA_ALL.filter(r => r.group === 1);
  const totalFund = normalSection.reduce((s, r) => s + r.fund, 0);
  const aprPnL = normalSection.reduce((s, r) => s + r.apr, 0);
  const mayPnL = normalSection.reduce((s, r) => s + r.may, 0);
  const junPnL = normalSection.reduce((s, r) => s + r.jun, 0);
  const q1PnL = aprPnL + mayPnL + junPnL;
  const q1ROI = totalFund > 0 ? q1PnL / (totalFund * 1e7) : 0;
  const annROI = q1ROI * ANN_FACTOR;
  const annPnL = q1PnL * ANN_FACTOR;
  // Q2 (Jul–Sep) and full-year-to-date (Q1 + Q2)
  const julPnL = normalSection.reduce((s, r) => s + r.jul, 0);
  const augPnL = normalSection.reduce((s, r) => s + r.aug, 0);
  const sepPnL = normalSection.reduce((s, r) => s + r.sep, 0);
  const q2PnL = julPnL + augPnL + sepPnL;
  const q2ROI = totalFund > 0 ? q2PnL / (totalFund * 1e7) : 0;
  const fyPnL = q1PnL + q2PnL;
  const fyROI = totalFund > 0 ? fyPnL / (totalFund * 1e7) : 0;
  const winners = normalSection.filter(r => r.q1 > 0).length;
  const losers = normalSection.filter(r => r.q1 < 0).length;
  const count = normalSection.length;
  // Per-month roll-up keyed by month key — drives the quarter popup and the
  // current-month tile without hardcoding another `xxxPnL` per month.
  const byMonth = {};
  FY_MONTHS.forEach(m => {
    const pnl = normalSection.reduce((s, r) => s + (r[m.key] || 0), 0);
    byMonth[m.key] = {
      pnl,
      roi: totalFund > 0 ? pnl / (totalFund * 1e7) : 0,
      winners: normalSection.filter(r => (r[m.key] || 0) > 0).length,
      losers: normalSection.filter(r => (r[m.key] || 0) < 0).length,
      hasData: normalSection.some(r => (r[m.key] || 0) !== 0),
    };
  });
  return { totalFund, aprPnL, mayPnL, junPnL, q1PnL, q1ROI, annROI, annPnL,
    julPnL, augPnL, sepPnL, q2PnL, q2ROI, fyPnL, fyROI, winners, losers, count, byMonth };
}, []);

// ── Current-month KPI tile ──
// Prefer the latest FY month that actually has numbers; only fall back to the
// calendar month when nothing has been booked yet, so a freshly-rolled month
// never blanks the headline tile at ₹0.
const currentMonth = useMemo(() => {
  const latest = [...FY_MONTHS].reverse().find(m => totals.byMonth[m.key].hasData);
  if (latest) return latest;
  const cal = new Date().getMonth();
  return FY_MONTHS.find(m => m.cal === cal) || FY_MONTHS[FY_MONTHS.length - 1];
}, [totals]);

// ── Quarter popup contents ──
// Month ROI uses total AUM as the denominator, matching the KPI cards.
const quarterDetail = useMemo(() => {
  if (!openQuarter) return null;
  const q = QUARTER_BY_ID[openQuarter];
  const rows = q.months.map(k => ({ ...FY_MONTHS.find(m => m.key === k), ...totals.byMonth[k] }));
  const pnl = rows.reduce((s, r) => s + r.pnl, 0);
  const roi = totals.totalFund > 0 ? pnl / (totals.totalFund * 1e7) : 0;
  const peak = Math.max(1, ...rows.map(r => Math.abs(r.pnl)));
  const live = rows.filter(r => r.hasData);
  const best  = live.length ? live.reduce((a, b) => (b.pnl > a.pnl ? b : a)) : null;
  const worst = live.length ? live.reduce((a, b) => (b.pnl < a.pnl ? b : a)) : null;
  return { q, rows, pnl, roi, peak, best, worst, pending: rows.filter(r => !r.hasData) };
}, [openQuarter, totals]);

// Only the open card renders a panel, so one shared node is enough.
const quarterDropdown = quarterDetail
  ? <QuarterMonths detail={quarterDetail} totalFund={totals.totalFund} />
  : null; // ── Month bar data (top 20 by absolute PnL) ──
    const monthKey=selectedMonth; const monthROIKey=selectedMonth + "_roi"; const barData=useMemo(()=> {
    const sorted = [...DATA]
    .filter(r => Math.abs(r[monthKey]) > 0)
    .sort((a, b) => {
    if (sortBy === "roi") return b[monthROIKey] - a[monthROIKey];
    return b[monthKey] - a[monthKey];
    })
    .slice(0, 25);
    return sorted.map(r => ({
    name: r.name,
    fullName: r.name,
    pnl: r[monthKey],
    roi: +(r[monthROIKey] * 100).toFixed(2), // actual ROI (monthly or Q1), no projection
    fund: r.fund,
    }));
    }, [selectedMonth, sortBy, monthKey, monthROIKey]);

    // ── Best/Worst performers Q1 ──
    const ranked = useMemo(() => [...DATA]
    .filter(r => !r.isCash && r.q1 !== 0)
    .sort((a, b) => b.q1_roi - a.q1_roi), []);

    const top5 = useMemo(() => ranked
        .filter(r => {
            const n = r.name.toLowerCase();
            return !n.includes("jinesh jain") && !n.includes("ramakar jha");
        })
        .slice(0, 5), [ranked]);

    const bot5 = useMemo(() => ranked
        .filter(r => {
            const n = r.name.toLowerCase();
            return !n.includes("jinesh jain") && !n.includes("ramakar jha");
        })
        .slice(-5).reverse(), [ranked]);

    // ── Best/Worst performers Q2 (F&O) ──
    const q2Ranked = useMemo(() => [...DATA]
    .filter(r => !r.isCash && r.q2 !== 0)
    .sort((a, b) => b.q2_roi - a.q2_roi), []);

    const q2Top5 = useMemo(() => q2Ranked
        .filter(r => {
            const n = r.name.toLowerCase();
            return !n.includes("jinesh jain") && !n.includes("ramakar jha");
        })
        .slice(0, 5), [q2Ranked]);

    const q2Bot5 = useMemo(() => q2Ranked
        .filter(r => {
            const n = r.name.toLowerCase();
            return !n.includes("jinesh jain") && !n.includes("ramakar jha");
        })
        .slice(-5).reverse(), [q2Ranked]);

    // ── Best/Worst performers Q1 (Cash) ──
    const cashRanked = useMemo(() => [...DATA_CASH]
    .filter(r => r.q1 !== 0)
    .sort((a, b) => b.q1_roi - a.q1_roi), []);

    const cashTop5 = cashRanked.slice(0, 5);
    const cashBot5 = cashRanked.slice(-5).reverse();

    // ── Strategy summary ──
    const strategyData = useMemo(() => {
    const map = {};
    DATA.forEach(r => {
    const s = r.strategy || "Unclassified";
    if (!map[s]) map[s] = { strategy: s, pnl: 0, fund: 0, count: 0 };
    map[s].pnl += r.q1;
    map[s].fund += r.fund;
    map[s].count++;
    });
    return Object.values(map)
    .sort((a, b) => b.pnl - a.pnl)
    .map(s => ({ ...s, roi: s.fund > 0 ? (s.pnl / (s.fund * 1e7)) * 100 : 0 }));
    }, []);

    // ── Monthly trend (Apr–Sep; Q2 months appear once data lands) ──
    const roiOf = (p) => totals.totalFund > 0 ? p / (totals.totalFund * 1e7) : 0;
    const trendData = [
    { month: "Apr", pnl: totals.aprPnL, roi: roiOf(totals.aprPnL) },
    { month: "May", pnl: totals.mayPnL, roi: roiOf(totals.mayPnL) },
    { month: "Jun", pnl: totals.junPnL, roi: roiOf(totals.junPnL) },
    { month: "Jul", pnl: totals.julPnL, roi: roiOf(totals.julPnL) },
    { month: "Aug", pnl: totals.augPnL, roi: roiOf(totals.augPnL) },
    { month: "Sep", pnl: totals.sepPnL, roi: roiOf(totals.sepPnL) },
    ];

    const tabs = [
    { id: "monthly", label: "Monthly Report" },
    { id: "quarterly", label: "Quarter Summary" },
    { id: "user", label: "User Performance" },
    { id: "strategy", label: "Strategy Analysis" },
    { id: "cash", label: "💰 Cash / ETF / ATS" },
    ];

    const tabStyle = (id) => ({
    padding: "9px 20px", borderRadius: 8, cursor: "pointer", fontWeight: 700,
    fontSize: 13, border: "none", letterSpacing: 0.5,
    background: view === id ? ACCENT : "transparent",
    color: view === id ? "#000" : "var(--muted)",
    transition: "all 0.2s",
    });

    const sectionHead = (title) => (
    <div style={{ color: ACCENT, fontSize: 13, fontWeight: 800, letterSpacing: 3, textTransform: "uppercase" ,
        marginBottom: 16, borderBottom: `1px solid ${BORDER}`, paddingBottom: 8 }}>
        {title}
    </div>
    );

    return (
    <div style={{ background: BG, minHeight: "100vh" , fontFamily: "'DM Sans', sans-serif" , color: "var(--text)" ,
        padding: "0 0 60px" }}>
        <link
            href="https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;700;800&family=DM+Mono:wght@400;500;700&display=swap"
            rel="stylesheet" />


        {/* Header */}
        <div style={{ background: "var(--header-bg)", borderBottom: `1px solid ${BORDER}`, padding: "24px 32px 20px" }}>
            <div style={{ display: "flex" , alignItems: "center" , justifyContent: "space-between" , flexWrap: "wrap" ,
                gap: 12 }}>
                <div>
                    <div style={{ display: "flex" , alignItems: "center" , gap: 12, marginBottom: 4 }}>
                        <div style={{ width: 8, height: 36, background: ACCENT, borderRadius: 4 }} />
                        <div>
                            <h1 style={{ margin: 0, fontSize: 26, fontWeight: 800, letterSpacing: -0.5 }}>Quant Strategy
                                Dashboard</h1>
                            <div style={{ color: "var(--muted2)" , fontSize: 12, marginTop: 2 }}>FY 2026–27 · Q1 + Q2
                                Performance Report · April – September 2026</div>
                        </div>
                    </div>
                </div>
                <div className="no-print" style={{ display: "flex" , gap: 10, alignItems: "center" }}>
                    <button
                        onClick={() => setShowPrintModal(true)}
                        title="Generate a PDF report of selected sections"
                        style={{
                            display: "flex", alignItems: "center", gap: 8,
                            background: "rgba(167, 139, 250, 0.12)",
                            border: "1px solid var(--violet)", color: "var(--violet)",
                            borderRadius: 8, padding: "9px 16px", cursor: "pointer",
                            fontSize: 12, fontWeight: 700, transition: "all 0.2s ease-in-out", outline: "none"
                        }}
                        onMouseEnter={(e) => { e.currentTarget.style.background = "var(--violet)"; e.currentTarget.style.color = "#000"; }}
                        onMouseLeave={(e) => { e.currentTarget.style.background = "rgba(167, 139, 250, 0.12)"; e.currentTarget.style.color = "var(--violet)"; }}
                    >
                        <span style={{ fontSize: 14 }}>📄</span> Report
                    </button>
                    <button
                        onClick={() => setShowExportModal(true)}
                        style={{
                            display: "flex",
                            alignItems: "center",
                            gap: 8,
                            background: "rgba(0, 229, 255, 0.1)",
                            border: `1px solid ${ACCENT}`,
                            color: ACCENT,
                            borderRadius: 8,
                            padding: "9px 16px",
                            cursor: "pointer",
                            fontSize: 12,
                            fontWeight: 700,
                            transition: "all 0.2s ease-in-out",
                            boxShadow: "0 0 10px rgba(0, 229, 255, 0.05)",
                            outline: "none"
                        }}
                        onMouseEnter={(e) => {
                            e.currentTarget.style.background = ACCENT;
                            e.currentTarget.style.color = "#000";
                            e.currentTarget.style.boxShadow = `0 0 15px ${ACCENT}66`;
                        }}
                        onMouseLeave={(e) => {
                            e.currentTarget.style.background = "rgba(0, 229, 255, 0.1)";
                            e.currentTarget.style.color = ACCENT;
                            e.currentTarget.style.boxShadow = "0 0 10px rgba(0, 229, 255, 0.05)";
                        }}
                    >
                        <span style={{ fontSize: 14 }}>📥</span> Export Excel
                    </button>

                    <button
                        className="theme-toggle"
                        onClick={() => setTheme((t) => (t === 'dark' ? 'light' : 'dark'))}
                        aria-label="Toggle light / dark theme"
                        title="Toggle light / dark theme"
                    >
                        {theme === 'dark' ? '☀️ Light' : '🌙 Dark'}
                    </button>

                    <div style={{ background: totals.fyPnL>= 0 ? "rgba(34,197,94,0.15)" : "rgba(244,63,94,0.15)",
                        border: `1px solid ${totals.fyPnL >= 0 ? POS : NEG}`, borderRadius: 8, padding: "8px 16px",
                        textAlign: "center" }}>
                        <div style={{ fontSize: 10, color: "var(--muted)" , fontWeight: 700, letterSpacing: 2 }}>FY NET P&L
                            (YTD)</div>
                        <div style={{ fontSize: 20, fontWeight: 800, color: totals.fyPnL>= 0 ? POS : NEG, fontFamily:
                            "'DM Mono', monospace" }}>{fmtSign(totals.fyPnL)}</div>
                    </div>
                </div>
            </div>

            {/* KPI Row */}
            {(!printMode || printSel.summary) && (
            <div className="print-section" style={{ display: "grid" , gridTemplateColumns: "repeat(auto-fit, minmax(160px, 1fr))" , gap: 12,
                marginTop: 20 }}>
                <StatCard label="Total AUM" value={fmtFund(totals.totalFund)}
                    sub="Total Utilized Fund" color={ACCENT} />
                <StatCard label="Q1 Net P&L" value={fmtSign(totals.q1PnL)} sub={`ROI ${fmtROI(totals.q1ROI)} · Apr–Jun`}
                    color={totals.q1PnL>= 0 ? POS : NEG}
                    onClick={printMode ? undefined : () => toggleQuarter("q1")} hint="Monthly split"
                    open={openQuarter === "q1"} dropdown={quarterDropdown} />
                <StatCard label="Q2 Net P&L" value={fmtSign(totals.q2PnL)} sub={`ROI ${fmtROI(totals.q2ROI)} · Jul–Sep`}
                    color={totals.q2PnL>= 0 ? POS : NEG}
                    onClick={printMode ? undefined : () => toggleQuarter("q2")} hint="Monthly split"
                    open={openQuarter === "q2"} dropdown={quarterDropdown} />
                <StatCard label="FY Net P&L (YTD)" value={fmtSign(totals.fyPnL)} sub={`ROI ${fmtROI(totals.fyROI)} · Q1 + Q2`}
                    color={totals.fyPnL>= 0 ? POS : NEG}
                    onClick={printMode ? undefined : () => toggleQuarter("fy")} hint="All months"
                    open={openQuarter === "fy"} dropdown={quarterDropdown} />
                <StatCard label={`${currentMonth.label} P&L`} value={fmtSign(totals.byMonth[currentMonth.key].pnl)}
                    sub={`ROI ${fmtROI(totals.byMonth[currentMonth.key].roi)} · Current month`}
                    color={totals.byMonth[currentMonth.key].pnl >= 0 ? POS : NEG} />
                <StatCard label="Winners / Losers" value={`${totals.winners} / ${totals.losers}`}
                    sub={`${totals.count} total accounts`} color="#a78bfa" />
                {/* Browser print can't open popups — lay every month out flat so the
                    printed summary stays complete. */}
                {printMode && FY_MONTHS.map(m => (
                <StatCard key={m.key} label={`${m.label} P&L`} value={fmtSign(totals.byMonth[m.key].pnl)}
                    sub={fmtROI(totals.byMonth[m.key].roi)}
                    color={totals.byMonth[m.key].pnl >= 0 ? POS : NEG} />
                ))}
            </div>
            )}
        </div>

        {/* Tabs */}
        <div className="no-print" style={{ padding: "16px 32px 0" , display: "flex" , gap: 6, borderBottom: `1px solid ${BORDER}`,
            background: "var(--panel)" }}>
            {tabs.map(t => (
            <button key={t.id} style={tabStyle(t.id)} onClick={()=> setView(t.id)}>{t.label}</button>
            ))}
        </div>

        <div style={{ padding: "28px 32px" }}>

            {/* ── MONTHLY ── */}
            {showSection("monthly") && (
            <div className="print-section">
                <div className="no-print" style={{ display: "flex", justifyContent: "flex-end", marginBottom: 12 }}>
                    <ExportBtn onClick={exportMonthlySheet} label="Export Monthly Sheet" />
                </div>
                {sectionHead("Monthly P&L Trend")}
                <div style={{ display: "grid" , gridTemplateColumns: "1fr 1fr" , gap: 24, marginBottom: 32 }}>
                    <div id="rpt-monthly-pnl" style={{ background: CARD, borderRadius: 12, padding: 20, border: `1px solid ${BORDER}` }}>
                        <div style={{ color: MUTED_C , fontSize: 12, fontWeight: 700, marginBottom: 14 }}>MONTHLY NET
                            P&L (₹)</div>
                        <ResponsiveContainer width="100%" height={220}>
                            <BarChart data={trendData} barSize={50}>
                                <CartesianGrid strokeDasharray="3 3" stroke={BORDER} />
                                <XAxis dataKey="month" tick={{ fill: MUTED_C , fontSize: 12 }} axisLine={false}
                                    tickLine={false} />
                                <YAxis tick={{ fill: MUTED_C , fontSize: 11 }} tickFormatter={v=> fmt(v)}
                                    axisLine={false} tickLine={false} />
                                    <Tooltip content={<CustomTooltip />} />
                                    <ReferenceLine y={0} stroke={BORDER} />
                                    <Bar dataKey="pnl" name="Net P&L" radius={[6, 6, 0, 0]}>
                                        {trendData.map((d, i) => <Cell key={i} fill={d.pnl>= 0 ? POS : NEG} />)}
                                    </Bar>
                            </BarChart>
                        </ResponsiveContainer>
                    </div>
                    <div id="rpt-monthly-roi" style={{ background: CARD, borderRadius: 12, padding: 20, border: `1px solid ${BORDER}` }}>
                        <div style={{ color: MUTED_C , fontSize: 12, fontWeight: 700, marginBottom: 14 }}>MONTHLY ROI
                            (%)</div>
                        <ResponsiveContainer width="100%" height={220}>
                            <LineChart data={trendData}>
                                <CartesianGrid strokeDasharray="3 3" stroke={BORDER} />
                                <XAxis dataKey="month" tick={{ fill: MUTED_C , fontSize: 12 }} axisLine={false}
                                    tickLine={false} />
                                <YAxis tick={{ fill: MUTED_C , fontSize: 11 }} tickFormatter={v=> `${(v*100).toFixed(2)}%`}
                                    axisLine={false} tickLine={false} />
                                    <Tooltip content={<CustomTooltip />} />
                                    <ReferenceLine y={0} stroke={BORDER} />
                                    <Line dataKey="roi" name="ROI %" stroke={ACCENT} strokeWidth={3} dot={{ fill:
                                        ACCENT, r: 6 }} />
                            </LineChart>
                        </ResponsiveContainer>
                    </div>
                </div>

                {/* Month selector + bar chart */}
                {sectionHead("Per-User Performance by Period")}
                <div style={{ display: "flex" , gap: 10, marginBottom: 16, alignItems: "center" , flexWrap: "wrap" }}>
                    {[{ id: "apr", label: "April" }, { id: "may", label: "May" }, { id: "jun", label: "June" },
                    { id: "q1", label: "Q1 Total" }, { id: "jul", label: "July" }, { id: "aug", label: "Aug" },
                    { id: "sep", label: "Sep" }, { id: "q2", label: "Q2 Total" },
                    { id: "year", label: "Yearly (FY)" }].map(m => (
                    <button key={m.id} onClick={()=> setSelectedMonth(m.id)} style={{
                        padding: "6px 16px", borderRadius: 6, border: `1px solid ${BORDER}`, cursor: "pointer",
                        fontSize: 12, fontWeight: 700,
                        background: selectedMonth === m.id ? ACCENT : CARD, color: selectedMonth === m.id ? "#000" :
                        MUTED_C,
                        }}>{m.label}</button>
                    ))}
                    <div style={{ marginLeft: "auto" , display: "flex" , gap: 8 }}>
                        <span style={{ color: MUTED_C , fontSize: 12, alignSelf: "center" }}>Sort:</span>
                        {["roi", "pnl"].map(s => (
                        <button key={s} onClick={()=> setSortBy(s)} style={{
                            padding: "6px 14px", borderRadius: 6, border: `1px solid ${BORDER}`, cursor: "pointer",
                            fontSize: 12,
                            background: sortBy === s ? BORDER : "transparent", color: sortBy === s ? ACCENT :
                            MUTED_C,
                            }}>{s.toUpperCase()}</button>
                        ))}
                    </div>
                </div>

                <div style={{ background: CARD, borderRadius: 12, padding: 20, border: `1px solid ${BORDER}` }}>
                    <ResponsiveContainer width="100%" height={420}>
                        <BarChart data={barData} layout="vertical" barSize={16} margin={{ left: 160 }}>
                            <CartesianGrid strokeDasharray="3 3" stroke={BORDER} horizontal={false} />
                            <XAxis type="number" tick={{ fill: MUTED_C , fontSize: 11 }} tickFormatter={v=> fmt(v)}
                                axisLine={false} tickLine={false} />
                                <YAxis type="category" dataKey="name" tick={{ fill: AXIS_C , fontSize: 12 }}
                                    interval={0} width={160} axisLine={false} tickLine={false} />
                                <Tooltip content={({ active, payload, label })=> {
                                    if (!active || !payload?.length) return null;
                                    const d = payload[0]?.payload;
                                    return (
                                    <div style={{ background: TIP_BG , border: `1px solid ${TIP_BORDER}`, borderRadius:
                                        8, padding: "10px 14px" , fontSize: 12 }}>
                                        <div style={{ color: ACCENT, fontWeight: 700, marginBottom: 4 }}>{d?.fullName}
                                        </div>
                                        <div style={{ color: d?.pnl>= 0 ? POS : NEG }}>P&L: {fmt(d?.pnl)}</div>
                                        <div style={{ color: d?.roi>= 0 ? POS : NEG }}>ROI: {d?.roi?.toFixed(2)}%
                                        </div>
                                        <div style={{ color: MUTED_C }}>Fund: ₹{d?.fund} Cr</div>
                                    </div>
                                    );
                                    }} />
                                    <ReferenceLine x={0} stroke={BORDER} />
                                    <Bar dataKey="pnl" name="P&L" radius={[0, 4, 4, 0]}>
                                        {barData.map((d, i) => <Cell key={i} fill={d.pnl>= 0 ? POS : NEG} />)}
                                    </Bar>
                        </BarChart>
                    </ResponsiveContainer>
                </div>
            </div>
            )}

            {/* ── QUARTERLY ── */}
            {showSection("quarterly") && (
            <div className="print-section">
                <div className="no-print" style={{ display: "flex", justifyContent: "flex-end", marginBottom: 12 }}>
                    <ExportBtn onClick={exportQuarterSheet} label="Export Quarter Sheet" />
                </div>
                {sectionHead("Q1 2026–27 Consolidated Report (Apr + May + Jun)")}

                {/* Best vs Worst */}
                <div style={{ display: "grid" , gridTemplateColumns: "1fr 1fr" , gap: 24, marginBottom: 28 }}>
                    {/* Top 5 */}
                    <div style={{ background: CARD, borderRadius: 12, padding: 20, border: `1px solid ${BORDER}` }}>
                        <div style={{ color: POS, fontSize: 12, fontWeight: 800, letterSpacing: 2, marginBottom: 14 }}>
                            🏆 TOP 5 PERFORMERS — Q1 ROI</div>
                        {top5.map((r, i) => (
                        <div key={r.code} style={{ display: "flex" , justifyContent: "space-between" ,
                            alignItems: "center" , padding: "10px 0" , borderBottom: i < top5.length - 1 ? `1px solid ${BORDER}`
                            : "none" }}>
                            <div style={{ display: "flex" , gap: 10, alignItems: "center" }}>
                                <div style={{ width: 26, height: 26, borderRadius: "50%" , background: POS,
                                    color: "#000" , fontWeight: 800, fontSize: 12, display: "flex" ,
                                    alignItems: "center" , justifyContent: "center" }}>{i + 1}</div>
                                <div>
                                    <div style={{ fontWeight: 700, fontSize: 15, color: "var(--text)" }}>{r.name}</div>
                                    <div style={{ color: MUTED2_C , fontSize: 11 }}>{r.strategy || "—"} · {fmtFund(r.fund)}
                                    </div>
                                </div>
                            </div>
                            <div style={{ textAlign: "right" }}>
                                <div style={{ color: POS, fontWeight: 800, fontFamily: "'DM Mono', monospace" ,
                                    fontSize: 14 }}>{fmtROI(r.q1_roi)}</div>
                                <div style={{ color: MUTED_C , fontSize: 11 }}>{fmtSign(r.q1)}</div>
                            </div>
                        </div>
                        ))}
                    </div>

                    {/* Bot 5 */}
                    <div style={{ background: CARD, borderRadius: 12, padding: 20, border: `1px solid ${BORDER}` }}>
                        <div style={{ color: NEG, fontSize: 12, fontWeight: 800, letterSpacing: 2, marginBottom: 14 }}>
                            ⚠️ BOTTOM 5 PERFORMERS — Q1 ROI</div>
                        {bot5.map((r, i) => (
                        <div key={r.code} style={{ display: "flex" , justifyContent: "space-between" ,
                            alignItems: "center" , padding: "10px 0" , borderBottom: i < bot5.length - 1 ? `1px solid ${BORDER}`
                            : "none" }}>
                            <div style={{ display: "flex" , gap: 10, alignItems: "center" }}>
                                <div style={{ width: 26, height: 26, borderRadius: "50%" , background: NEG,
                                    color: "#fff" , fontWeight: 800, fontSize: 12, display: "flex" ,
                                    alignItems: "center" , justifyContent: "center" }}>{i + 1}</div>
                                <div>
                                    <div style={{ fontWeight: 700, fontSize: 15, color: "var(--text)" }}>{r.name}</div>
                                    <div style={{ color: MUTED2_C , fontSize: 11 }}>{r.strategy || "—"} · {fmtFund(r.fund)}
                                    </div>
                                </div>
                            </div>
                            <div style={{ textAlign: "right" }}>
                                <div style={{ color: NEG, fontWeight: 800, fontFamily: "'DM Mono', monospace" ,
                                    fontSize: 14 }}>{fmtROI(r.q1_roi)}</div>
                                <div style={{ color: MUTED_C , fontSize: 11 }}>{fmt(r.q1)}</div>
                            </div>
                        </div>
                        ))}
                    </div>
                </div>

                {/* Best vs Worst — Q2 */}
                {sectionHead("Q2 Best / Worst Performers (Jul + Aug + Sep)")}
                {q2Ranked.length === 0 ? (
                    <div style={{ background: CARD, borderRadius: 12, padding: 20, border: `1px solid ${BORDER}`,
                        color: MUTED_C, fontSize: 13, marginBottom: 28 }}>
                        No Q2 trades booked yet — this will populate as July / August / September data is entered.
                    </div>
                ) : (
                <div style={{ display: "grid" , gridTemplateColumns: "1fr 1fr" , gap: 24, marginBottom: 28 }}>
                    {/* Q2 Top 5 */}
                    <div style={{ background: CARD, borderRadius: 12, padding: 20, border: `1px solid ${BORDER}` }}>
                        <div style={{ color: POS, fontSize: 12, fontWeight: 800, letterSpacing: 2, marginBottom: 14 }}>
                            🏆 TOP 5 PERFORMERS — Q2 ROI</div>
                        {q2Top5.map((r, i) => (
                        <div key={r.code} style={{ display: "flex" , justifyContent: "space-between" ,
                            alignItems: "center" , padding: "10px 0" , borderBottom: i < q2Top5.length - 1 ? `1px solid ${BORDER}`
                            : "none" }}>
                            <div style={{ display: "flex" , gap: 10, alignItems: "center" }}>
                                <div style={{ width: 26, height: 26, borderRadius: "50%" , background: POS,
                                    color: "#000" , fontWeight: 800, fontSize: 12, display: "flex" ,
                                    alignItems: "center" , justifyContent: "center" }}>{i + 1}</div>
                                <div>
                                    <div style={{ fontWeight: 700, fontSize: 15, color: "var(--text)" }}>{r.name}</div>
                                    <div style={{ color: MUTED2_C , fontSize: 11 }}>{r.strategy || "—"} · {fmtFund(r.fund)}
                                    </div>
                                </div>
                            </div>
                            <div style={{ textAlign: "right" }}>
                                <div style={{ color: POS, fontWeight: 800, fontFamily: "'DM Mono', monospace" ,
                                    fontSize: 14 }}>{fmtROI(r.q2_roi)}</div>
                                <div style={{ color: MUTED_C , fontSize: 11 }}>{fmtSign(r.q2)}</div>
                            </div>
                        </div>
                        ))}
                    </div>

                    {/* Q2 Bot 5 */}
                    <div style={{ background: CARD, borderRadius: 12, padding: 20, border: `1px solid ${BORDER}` }}>
                        <div style={{ color: NEG, fontSize: 12, fontWeight: 800, letterSpacing: 2, marginBottom: 14 }}>
                            ⚠️ BOTTOM 5 PERFORMERS — Q2 ROI</div>
                        {q2Bot5.map((r, i) => (
                        <div key={r.code} style={{ display: "flex" , justifyContent: "space-between" ,
                            alignItems: "center" , padding: "10px 0" , borderBottom: i < q2Bot5.length - 1 ? `1px solid ${BORDER}`
                            : "none" }}>
                            <div style={{ display: "flex" , gap: 10, alignItems: "center" }}>
                                <div style={{ width: 26, height: 26, borderRadius: "50%" , background: NEG,
                                    color: "#fff" , fontWeight: 800, fontSize: 12, display: "flex" ,
                                    alignItems: "center" , justifyContent: "center" }}>{i + 1}</div>
                                <div>
                                    <div style={{ fontWeight: 700, fontSize: 15, color: "var(--text)" }}>{r.name}</div>
                                    <div style={{ color: MUTED2_C , fontSize: 11 }}>{r.strategy || "—"} · {fmtFund(r.fund)}
                                    </div>
                                </div>
                            </div>
                            <div style={{ textAlign: "right" }}>
                                <div style={{ color: NEG, fontWeight: 800, fontFamily: "'DM Mono', monospace" ,
                                    fontSize: 14 }}>{fmtROI(r.q2_roi)}</div>
                                <div style={{ color: MUTED_C , fontSize: 11 }}>{fmt(r.q2)}</div>
                            </div>
                        </div>
                        ))}
                    </div>
                </div>
                )}

                {/* Full Q1 bar */}
                {sectionHead("Full Q1 P&L & ROI — Best to Worst (F&O Only)")}
                <div style={{ background: CARD, borderRadius: 12, padding: 20, border: `1px solid ${BORDER}` }}>
                    <ResponsiveContainer width="100%" height={Math.max(500, ranked.length * 22)}>
                        <BarChart data={ranked.map(r=> ({ name: r.name,
                            pnl: r.q1, roi: +(r.q1_roi * 100).toFixed(2), fund: r.fund, fullName: r.name }))}
                            layout="vertical" barSize={14} margin={{ left: 165 }}>
                            <CartesianGrid strokeDasharray="3 3" stroke={BORDER} horizontal={false} />
                            <XAxis type="number" tick={{ fill: MUTED_C , fontSize: 11 }} tickFormatter={v=> fmt(v)}
                                axisLine={false} tickLine={false} />
                                <YAxis type="category" dataKey="name" tick={{ fill: AXIS_C , fontSize: 12 }}
                                    interval={0} width={165} axisLine={false} tickLine={false} />
                                <Tooltip content={({ active, payload })=> {
                                    if (!active || !payload?.length) return null;
                                    const d = payload[0]?.payload;
                                    return (
                                    <div style={{ background: TIP_BG , border: `1px solid ${TIP_BORDER}`, borderRadius:
                                        8, padding: "10px 14px" , fontSize: 12 }}>
                                        <div style={{ color: ACCENT, fontWeight: 700 }}>{d?.fullName}</div>
                                        <div style={{ color: d?.pnl>= 0 ? POS : NEG }}>Q1 P&L: {fmt(d?.pnl)}</div>
                                        <div style={{ color: d?.roi>= 0 ? POS : NEG }}>Q1 ROI: {d?.roi?.toFixed(2)}%
                                        </div>
                                    </div>
                                    );
                                    }} />
                                    <ReferenceLine x={0} stroke={BORDER} />
                                    <Bar dataKey="pnl" name="Q1 P&L" radius={[0, 4, 4, 0]}>
                                        {ranked.map((d, i) => <Cell key={i} fill={d.q1>= 0 ? POS : NEG} />)}
                                    </Bar>
                        </BarChart>
                    </ResponsiveContainer>
                </div>

                <div style={{ height: 40 }} />

                {sectionHead("Q1 2026–27 Consolidated Cash & ETF Report")}

                {/* Full Cash Q1 Bar Chart */}
                {sectionHead("Full Q1 P&L & ROI — Best to Worst (Cash & ETF Only)")}
                <div style={{ background: CARD, borderRadius: 12, padding: 20, border: `1px solid ${BORDER}` }}>
                    <ResponsiveContainer width="100%" height={Math.max(250, cashRanked.length * 35)}>
                        <BarChart data={cashRanked.map(r=> ({ name: r.name,
                            pnl: r.q1, roi: +(r.q1_roi * 100).toFixed(2), fund: r.fund, fullName: r.name, strategy: r.strategy }))}
                            layout="vertical" barSize={16} margin={{ left: 165 }}>
                            <CartesianGrid strokeDasharray="3 3" stroke={BORDER} horizontal={false} />
                            <XAxis type="number" tick={{ fill: MUTED_C , fontSize: 11 }} tickFormatter={v=> fmt(v)}
                                axisLine={false} tickLine={false} />
                            <YAxis type="category" dataKey="name" tick={{ fill: AXIS_C , fontSize: 12 }}
                                interval={0} width={165} axisLine={false} tickLine={false} />
                            <Tooltip content={({ active, payload })=> {
                                if (!active || !payload?.length) return null;
                                const d = payload[0]?.payload;
                                return (
                                <div style={{ background: TIP_BG , border: `1px solid ${TIP_BORDER}`, borderRadius:
                                    8, padding: "10px 14px" , fontSize: 12 }}>
                                    <div style={{ color: "var(--gold)" , fontWeight: 700 }}>{d?.fullName}</div>
                                    <div style={{ color: MUTED_C , fontSize: 11 }}>{d?.strategy || "—"} ·
                                        ₹{d?.fund}Cr</div>
                                    <div style={{ color: d?.pnl>= 0 ? POS : NEG }}>Q1 P&L: {fmt(d?.pnl)}</div>
                                    <div style={{ color: d?.roi>= 0 ? POS : NEG }}>Q1 ROI: {d?.roi?.toFixed(2)}%
                                    </div>
                                </div>
                                );
                                }} />
                            <ReferenceLine x={0} stroke={BORDER} />
                            <Bar dataKey="pnl" name="Q1 P&L" radius={[0, 4, 4, 0]}>
                                {cashRanked.map((d, i) => <Cell key={i} fill={d.q1>= 0 ? "var(--gold)" : NEG} />)}
                            </Bar>
                        </BarChart>
                    </ResponsiveContainer>
                </div>
            </div>
            )}

            {/* ── USER PERFORMANCE ── */}
            {showSection("user") && (
            <div className="print-section">
                <div className="no-print" style={{ display: "flex", justifyContent: "flex-end", marginBottom: 12 }}>
                    <ExportBtn onClick={exportUserSheet} label="Export User Sheet" />
                </div>
                {sectionHead("User-Wise Performance Table")}
            <div style={{ background: CARD, borderRadius: 12, border: `1px solid ${BORDER}`, overflow: "auto" }}>
                    <table style={{ width: "100%" , borderCollapse: "collapse" , fontSize: 11 }}>
                        <thead>
                            <tr style={{ background: "var(--panel)" , borderBottom: `1px solid ${BORDER}` }}>
                                {["#", "Code", "Name", "Strategy", "Fund (Cr)", "Apr P&L", "Apr ROI", "May P&L", "May ROI", "Jun P&L", "Jun ROI", "Q1 P&L", "Q1 ROI", "Jul P&L", "Jul ROI", "Q2 P&L", "Q2 ROI", "FY P&L", "FY ROI"].map(h => (
                                <th key={h} style={{ padding: "11px 10px" , textAlign: "left" , color: "var(--muted)" ,
                                    fontWeight: 700, letterSpacing: 1, whiteSpace: "nowrap" }}>{h}</th>
                                ))}
                            </tr>
                        </thead>
                        <tbody>
                            {ranked.map((r, i) => {
                            const rowBg = i % 2 === 0 ? CARD : "var(--row-alt)";
                            return (
                            <tr key={r.code} style={{ background: rowBg, borderBottom: `1px solid ${BORDER}22` }}>
                                <td style={{ padding: "9px 10px" , color: "var(--muted2)" }}>{i + 1}</td>
                                <td style={{ padding: "9px 10px" , color: ACCENT, fontFamily: "'DM Mono', monospace" ,
                                    fontSize: 11 }}>{r.code}</td>
                                <td style={{ padding: "9px 10px" , fontWeight: 700, color: "var(--text)", whiteSpace: "nowrap" }}>{r.name}</td>
                                <td style={{ padding: "9px 10px" , color: "var(--muted)" , fontSize: 11 }}>{r.strategy || "—"}
                                </td>
                                <td style={{ padding: "9px 10px" , fontFamily: "'DM Mono', monospace" , color: "var(--axis)" }}>
                                    {r.fund}</td>
                                {[
                                [r.apr, r.apr_roi],
                                [r.may, r.may_roi],
                                [r.jun, r.jun_roi],
                                [r.q1, r.q1_roi],
                                [r.jul, r.jul_roi],
                                [r.q2, r.q2_roi],
                                [r.year, r.year_roi],
                                ].map(([pnl, roi], j) => (
                                <>
                                    <td key={`pnl${j}`} style={{ padding: "9px 10px" , fontFamily: "'DM Mono', monospace" ,
                                        color: pnl>= 0 ? POS : NEG, whiteSpace: "nowrap", fontWeight: pnl !== 0 ? 700 : 400 }}>{pnl === 0 ? "—" :
                                        fmtSign(pnl)}</td>
                                    <td key={`roi${j}`} style={{ padding: "9px 10px" , fontFamily: "'DM Mono', monospace" ,
                                        color: roi>= 0 ? POS : NEG, fontWeight: roi !== 0 ? 700 : 400 }}>{roi === 0 ? "—" : fmtROI(roi)}</td>
                                </>
                                ))}
                            </tr>
                            );
                            })}
                        </tbody>
                    </table>
                </div>
            </div>
            )}

            {/* ── STRATEGY ── */}
            {showSection("strategy") && (
            <div className="print-section">
                <div className="no-print" style={{ display: "flex", justifyContent: "flex-end", marginBottom: 12 }}>
                    <ExportBtn onClick={exportStrategySheet} label="Export Strategy Sheet" />
                </div>
                {sectionHead("Strategy-Wise Q1 Performance")}
                <div style={{ display: "grid" , gridTemplateColumns: "1fr 1fr" , gap: 24, marginBottom: 28 }}>
                    <div id="rpt-strategy-pnl" style={{ background: CARD, borderRadius: 12, padding: 20, border: `1px solid ${BORDER}` }}>
                        <div style={{ color: MUTED_C , fontSize: 12, fontWeight: 700, marginBottom: 14 }}>STRATEGY NET
                            P&L — Q1</div>
                        <ResponsiveContainer width="100%" height={320}>
                            <BarChart data={strategyData.slice(0, 14)} layout="vertical" barSize={18} margin={{ left: 120
                                }}>
                                <CartesianGrid strokeDasharray="3 3" stroke={BORDER} horizontal={false} />
                                <XAxis type="number" tick={{ fill: MUTED_C , fontSize: 11 }} tickFormatter={v=>
                                    fmt(v)} axisLine={false} tickLine={false} />
                                    <YAxis type="category" dataKey="strategy" tick={{ fill: AXIS_C , fontSize: 12 }}
                                        interval={0} width={120} axisLine={false} tickLine={false} />
                                    <Tooltip content={<CustomTooltip />} />
                                    <ReferenceLine x={0} stroke={BORDER} />
                                    <Bar dataKey="pnl" name="Q1 P&L" radius={[0, 4, 4, 0]}>
                                        {strategyData.slice(0, 14).map((d, i) => <Cell key={i} fill={d.pnl>= 0 ? POS :
                                            NEG} />)}
                                    </Bar>
                            </BarChart>
                        </ResponsiveContainer>
                    </div>

                    <div style={{ background: CARD, borderRadius: 12, padding: 20, border: `1px solid ${BORDER}` }}>
                        <div style={{ color: MUTED_C , fontSize: 12, fontWeight: 700, marginBottom: 14 }}>STRATEGY ROI
                            % — Q1</div>
                        <ResponsiveContainer width="100%" height={320}>
                            <BarChart data={strategyData.filter(d=> d.fund > 0).slice(0, 14)} layout="vertical"
                                barSize={18} margin={{ left: 120 }}>
                                <CartesianGrid strokeDasharray="3 3" stroke={BORDER} horizontal={false} />
                                <XAxis type="number" tick={{ fill: MUTED_C , fontSize: 11 }} tickFormatter={v=>
                                    `${v}%`} axisLine={false} tickLine={false} />
                                    <YAxis type="category" dataKey="strategy" tick={{ fill: AXIS_C , fontSize: 12 }}
                                        interval={0} width={120} axisLine={false} tickLine={false} />
                                    <Tooltip content={<CustomTooltip />} />
                                    <ReferenceLine x={0} stroke={BORDER} />
                                    <Bar dataKey="roi" name="ROI %" radius={[0, 4, 4, 0]}>
                                        {strategyData.filter(d => d.fund > 0).slice(0, 14).map((d, i) => <Cell key={i}
                                            fill={d.roi>= 0 ? POS : NEG} />)}
                                    </Bar>
                            </BarChart>
                        </ResponsiveContainer>
                    </div>
                </div>

                {/* Strategy table */}
                <div style={{ background: CARD, borderRadius: 12, border: `1px solid ${BORDER}`, overflow: "auto" }}>
                    <table style={{ width: "100%" , borderCollapse: "collapse" , fontSize: 11 }}>
                        <thead>
                            <tr style={{ background: "var(--panel)" , borderBottom: `1px solid ${BORDER}` }}>
                                {["Strategy", "Accounts", "Fund (Cr)", "Q1 Net P&L", "Q1 ROI", "Status"].map(h => (
                                <th key={h} style={{ padding: "11px 14px" , textAlign: "left" , color: "var(--muted)" ,
                                    fontWeight: 700, letterSpacing: 1 }}>{h}</th>
                                ))}
                            </tr>
                        </thead>
                        <tbody>
                            {strategyData.map((s, i) => (
                            <tr key={s.strategy} style={{ background: i % 2===0 ? CARD : "var(--row-alt)" , borderBottom: `1px
                                solid ${BORDER}22` }}>
                                <td style={{ padding: "9px 14px" , fontWeight: 700, color: "var(--text)" }}>{s.strategy}</td>
                                <td style={{ padding: "9px 14px" , color: "var(--muted)" }}>{s.count}</td>
                                <td style={{ padding: "9px 14px" , fontFamily: "'DM Mono', monospace" ,
                                    color: "var(--axis)" }}>₹{s.fund.toFixed(1)}</td>
                                <td style={{ padding: "9px 14px" , fontFamily: "'DM Mono', monospace" , color: s.pnl>=
                                    0 ? POS : NEG, fontWeight: 700 }}>{fmtSign(s.pnl)}</td>
                                <td style={{ padding: "9px 14px" , fontFamily: "'DM Mono', monospace" , color: s.roi>=
                                    0 ? POS : NEG, fontWeight: 700 }}>{s.fund > 0 ? `${s.roi.toFixed(2)}%` : "—"}</td>
                                <td style={{ padding: "9px 14px" }}>
                                    <span style={{ background: s.pnl>= 0 ? "rgba(34,197,94,0.15)" :
                                        "rgba(244,63,94,0.15)", color: s.pnl >= 0 ? POS : NEG, padding: "3px 10px",
                                        borderRadius: 20, fontSize: 11, fontWeight: 700 }}>
                                        {s.pnl >= 0 ? "PROFIT" : "LOSS"}
                                    </span>
                                </td>
                            </tr>
                            ))}
                        </tbody>
                    </table>
                </div>
            </div>
            )}


            {/* ── CASH / ETF / ATS ── */}
            {showSection("cash") && (() => {
            const cashSection = DATA_ALL.filter(r => r.group === 2 && r.code !== "P3361");
            const cashTotalFund = cashSection.reduce((s, r) => s + r.fund, 0);
            const cashApr = cashSection.reduce((s, r) => s + r.apr, 0);
            const cashMay = cashSection.reduce((s, r) => s + r.may, 0);
            const cashJun = cashSection.reduce((s, r) => s + r.jun, 0);
            const cashQ1 = cashApr + cashMay + cashJun;
            const cashQ1ROI = cashTotalFund > 0 ? cashQ1 / (cashTotalFund * 1e7) : 0;
            const cashTrend = [
            { month: "April", pnl: cashApr, roi: cashTotalFund > 0 ? cashApr / (cashTotalFund * 1e7) : 0 },
            { month: "May", pnl: cashMay, roi: cashTotalFund > 0 ? cashMay / (cashTotalFund * 1e7) : 0 },
            { month: "June", pnl: cashJun, roi: cashTotalFund > 0 ? cashJun / (cashTotalFund * 1e7) : 0 },
            ];
            return (
            <div className="print-section">
                <div className="no-print" style={{ display: "flex", justifyContent: "flex-end", marginBottom: 12 }}>
                    <ExportBtn onClick={exportCashSheet} label="Export Cash / ETF Sheet" />
                </div>
                {/* Notice banner */}
                <div style={{ background: "rgba(234,179,8,0.08)" , border: "1px solid rgba(234,179,8,0.3)" ,
                    borderRadius: 10, padding: "12px 18px" , marginBottom: 24, display: "flex" , alignItems: "center" ,
                    gap: 10 }}>
                    <span style={{ fontSize: 18 }}>💰</span>
                    <span style={{ color: "#fbbf24" , fontSize: 13 }}>
                        Cash / ETF / ATS accounts are shown <strong>separately</strong> here and excluded from all other
                        reports to avoid skewing F&O strategy numbers.
                    </span>
                </div>

                {/* KPIs */}
                <div style={{ display: "grid" , gridTemplateColumns: "repeat(auto-fit, minmax(150px, 1fr))" , gap: 12,
                    marginBottom: 24 }}>
                    <StatCard label="Cash AUM" value={fmtFund(cashTotalFund)}
                        sub="Total Utilized Fund" color="#fbbf24" />
                    <StatCard label="Q1 Net P&L" value={fmtSign(cashQ1)} color={cashQ1>= 0 ? POS : NEG} />
                        <StatCard label="Q1 ROI" value={fmtROI(cashQ1ROI)} color={cashQ1ROI>= 0 ? POS : NEG} />
                            <StatCard label="April P&L" value={fmtSign(cashApr)} color={cashApr>= 0 ? POS : NEG} />
                                <StatCard label="May P&L" value={fmtSign(cashMay)} color={cashMay>= 0 ? POS : NEG} />
                                    <StatCard label="June P&L" value={fmtSign(cashJun)} color={cashJun>= 0 ? POS : NEG}
                                        />
                </div>

                {/* Trend charts */}
                {sectionHead("Cash / ETF Monthly Trend")}
                <div style={{ display: "grid" , gridTemplateColumns: "1fr 1fr" , gap: 24, marginBottom: 28 }}>
                    <div style={{ background: CARD, borderRadius: 12, padding: 20, border: `1px solid ${BORDER}` }}>
                        <div style={{ color: MUTED_C , fontSize: 12, fontWeight: 700, marginBottom: 14 }}>MONTHLY NET
                            P&L (₹)</div>
                        <ResponsiveContainer width="100%" height={200}>
                            <BarChart data={cashTrend} barSize={50}>
                                <CartesianGrid strokeDasharray="3 3" stroke={BORDER} />
                                <XAxis dataKey="month" tick={{ fill: MUTED_C , fontSize: 12 }} axisLine={false}
                                    tickLine={false} />
                                <YAxis tick={{ fill: MUTED_C , fontSize: 11 }} tickFormatter={v=> fmt(v)}
                                    axisLine={false} tickLine={false} />
                                    <Tooltip content={<CustomTooltip />} />
                                    <ReferenceLine y={0} stroke={BORDER} />
                                    <Bar dataKey="pnl" name="Net P&L" radius={[6,6,0,0]}>
                                        {cashTrend.map((d, i) => <Cell key={i} fill={d.pnl>= 0 ? "#fbbf24" : NEG} />)}
                                    </Bar>
                            </BarChart>
                        </ResponsiveContainer>
                    </div>
                    <div style={{ background: CARD, borderRadius: 12, padding: 20, border: `1px solid ${BORDER}` }}>
                        <div style={{ color: MUTED_C , fontSize: 12, fontWeight: 700, marginBottom: 14 }}>MONTHLY ROI
                            (%)</div>
                        <ResponsiveContainer width="100%" height={200}>
                            <LineChart data={cashTrend}>
                                <CartesianGrid strokeDasharray="3 3" stroke={BORDER} />
                                <XAxis dataKey="month" tick={{ fill: MUTED_C , fontSize: 12 }} axisLine={false}
                                    tickLine={false} />
                                <YAxis tick={{ fill: MUTED_C , fontSize: 11 }} tickFormatter={v=> `${(v*100).toFixed(2)}%`}
                                    axisLine={false} tickLine={false} />
                                    <Tooltip content={<CustomTooltip />} />
                                    <ReferenceLine y={0} stroke={BORDER} />
                                    <Line dataKey="roi" name="ROI %" stroke="#fbbf24" strokeWidth={3} dot={{
                                        fill: "#fbbf24" , r: 6 }} />
                            </LineChart>
                        </ResponsiveContainer>
                    </div>
                </div>

                {/* Q1 bar chart */}
                {sectionHead("Cash / ETF — Best to Worst (Q1)")}
                <div style={{ background: CARD, borderRadius: 12, padding: 20, border: `1px solid ${BORDER}`,
                    marginBottom: 28 }}>
                    <ResponsiveContainer width="100%" height={Math.max(200, cashRanked.length * 38)}>
                        <BarChart data={cashRanked.map(r=> ({ name: r.name, pnl: r.q1, roi: +(r.q1_roi*100).toFixed(2), fund: r.fund, fullName: r.name,
                            strategy: r.strategy }))} layout="vertical" barSize={20} margin={{ left: 165 }}>
                            <CartesianGrid strokeDasharray="3 3" stroke={BORDER} horizontal={false} />
                            <XAxis type="number" tick={{ fill: MUTED_C , fontSize: 11 }} tickFormatter={v=> fmt(v)}
                                axisLine={false} tickLine={false} />
                                <YAxis type="category" dataKey="name" tick={{ fill: AXIS_C , fontSize: 12 }}
                                    interval={0} width={165} axisLine={false} tickLine={false} />
                                <Tooltip content={({ active, payload })=> {
                                    if (!active || !payload?.length) return null;
                                    const d = payload[0]?.payload;
                                    return (
                                    <div style={{ background: TIP_BG , border: `1px solid ${TIP_BORDER}`, borderRadius:
                                        8, padding: "10px 14px" , fontSize: 12 }}>
                                        <div style={{ color: "var(--gold)" , fontWeight: 700 }}>{d?.fullName}</div>
                                        <div style={{ color: MUTED_C , fontSize: 11 }}>{d?.strategy || "—"} ·
                                            ₹{d?.fund}Cr</div>
                                        <div style={{ color: d?.pnl>= 0 ? POS : NEG }}>Q1 P&L: {fmt(d?.pnl)}</div>
                                        <div style={{ color: d?.roi>= 0 ? POS : NEG }}>Q1 ROI: {d?.roi?.toFixed(2)}%
                                        </div>
                                    </div>
                                    );
                                    }} />
                                    <ReferenceLine x={0} stroke={BORDER} />
                                    <Bar dataKey="pnl" name="Q1 P&L" radius={[0,4,4,0]}>
                                        {cashRanked.map((d, i) => <Cell key={i} fill={d.q1>= 0 ? "#fbbf24" : NEG} />)}
                                    </Bar>
                        </BarChart>
                    </ResponsiveContainer>
                </div>

                {/* Detail table */}
                {sectionHead("Cash / ETF Account Detail")}
                <div style={{ background: CARD, borderRadius: 12, border: `1px solid ${BORDER}`, overflow: "auto" }}>
                    <table style={{ width: "100%" , borderCollapse: "collapse" , fontSize: 11 }}>
                        <thead>
                            <tr style={{ background: "var(--panel)" , borderBottom: `1px solid ${BORDER}` }}>
                                {["#","Code","Name","Strategy","Fund (Cr)","Apr P&L","May P&L","Jun P&L","Q1 P&L","Q1 ROI"].map(h => (
                                <th key={h} style={{ padding:"11px 10px", textAlign:"left", color:"var(--muted)",
                                    fontWeight:700, letterSpacing:1, whiteSpace:"nowrap" }}>{h}</th>
                                ))}
                            </tr>
                        </thead>
                        <tbody>
                            {cashRanked.map((r, i) => (
                            <tr key={r.code} style={{ background: i % 2===0 ? CARD : "var(--row-alt)" , borderBottom: `1px
                                solid ${BORDER}22` }}>
                                <td style={{ padding:"9px 10px", color:"var(--muted2)" }}>{i+1}</td>
                                <td style={{ padding:"9px 10px", color:"var(--gold)", fontFamily:"'DM Mono',monospace",
                                    fontSize:11 }}>{r.code}</td>
                                <td style={{ padding:"9px 10px", fontWeight:700, color:"var(--text)" }}>{r.name}</td>
                                <td style={{ padding:"9px 10px", color:"var(--muted)", fontSize:11 }}>{r.strategy || "—"}</td>
                                <td style={{ padding:"9px 10px", fontFamily:"'DM Mono',monospace", color:"var(--axis)" }}>
                                    {r.fund}</td>
                                {[r.apr, r.may, r.jun, r.q1].map((v, j) => (
                                <td key={j} style={{ padding:"9px 10px", fontFamily:"'DM Mono',monospace", color: v>= 0 ?
                                    POS : NEG, whiteSpace:"nowrap", fontWeight: v !== 0 ? 700 : 400 }}>{v === 0 ? "—" : fmtSign(v)}</td>
                                ))}
                                <td style={{ padding:"9px 10px", fontFamily:"'DM Mono',monospace", color: r.q1_roi>= 0 ? POS
                                    : NEG, fontWeight: 700 }}>{fmtROI(r.q1_roi)}</td>
                            </tr>
                            ))}
                        </tbody>
                    </table>
                </div>
            </div>
            );
            })()}

        </div>

        {/* Report generating overlay */}
        {reportBusy && (
          <div className="no-print" style={{ position: "fixed", top: 0, left: 0, width: "100%", height: "100%",
            background: "rgba(4,7,13,0.92)", backdropFilter: "blur(4px)", display: "flex", flexDirection: "column",
            justifyContent: "center", alignItems: "center", zIndex: 2000, gap: 16 }}>
            <div style={{ width: 46, height: 46, border: "4px solid var(--border)", borderTopColor: "var(--violet)",
              borderRadius: "50%", animation: "qspin 0.8s linear infinite" }} />
            <div style={{ color: "var(--text)", fontSize: 15, fontWeight: 700 }}>Generating PDF report…</div>
            <div style={{ color: "var(--muted)", fontSize: 12 }}>Rendering charts &amp; tables — this can take a few seconds.</div>
            <style>{`@keyframes qspin { to { transform: rotate(360deg); } }`}</style>
          </div>
        )}

        {/* Print Modal overlay */}
        {showPrintModal && (
          <div className="no-print" style={{ position: "fixed", top: 0, left: 0, width: "100%", height: "100%",
            background: "rgba(4,7,13,0.85)", backdropFilter: "blur(6px)", display: "flex", justifyContent: "center",
            alignItems: "center", zIndex: 1000 }}>
            <div style={{ background: CARD, border: `1px solid ${BORDER}`, borderRadius: 16, padding: 28, width: 470,
              maxWidth: "92%", boxShadow: "0 20px 45px rgba(0,0,0,0.6)" }}>
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 6 }}>
                <h3 style={{ fontSize: 18, fontWeight: 800, color: "var(--text)", letterSpacing: -0.5 }}>📄 Generate Dashboard Report</h3>
                <button onClick={() => setShowPrintModal(false)} style={{ background: "transparent", border: "none",
                  color: "var(--muted)", fontSize: 24, cursor: "pointer", lineHeight: 1, outline: "none" }}>&times;</button>
              </div>
              <p style={{ color: "var(--muted)", fontSize: 12, marginBottom: 18 }}>
                Pick the sections to include, then download a multi-page PDF report (cover, summary, charts &amp; tables).</p>

              <div style={{ display: "flex", flexDirection: "column", gap: 3, marginBottom: 20 }}>
                {[
                  { id: "summary",   label: "Executive Summary (KPI cards)" },
                  { id: "monthly",   label: "Monthly Report (trend + selected month)" },
                  { id: "quarterly", label: "Quarter Summary (Top/Bottom + full Q1)" },
                  { id: "user",      label: "User Performance table" },
                  { id: "strategy",  label: "Strategy Analysis" },
                  { id: "cash",      label: "Cash / ETF / ATS" },
                ].map(s => (
                  <label key={s.id} style={{ display: "flex", alignItems: "center", gap: 10, padding: "9px 10px",
                    borderRadius: 8, cursor: "pointer", background: printSel[s.id] ? "rgba(167,139,250,0.10)" : "transparent",
                    fontSize: 13, fontWeight: printSel[s.id] ? 700 : 500, color: "var(--text)" }}>
                    <input type="checkbox" checked={printSel[s.id]}
                      onChange={(e) => setPrintSel({ ...printSel, [s.id]: e.target.checked })}
                      style={{ accentColor: "var(--violet)", width: 15, height: 15, cursor: "pointer" }} />
                    {s.label}
                  </label>
                ))}
              </div>

              <div style={{ display: "flex", justifyContent: "space-between", gap: 12, alignItems: "center" }}>
                <div style={{ display: "flex", gap: 8, alignItems: "center" }}>
                  <button onClick={() => setPrintSel({ summary: true, monthly: true, quarterly: true, user: true, strategy: true, cash: true })}
                    style={{ background: "transparent", border: "none", color: "var(--violet)", fontSize: 11, fontWeight: 700, cursor: "pointer", outline: "none" }}>Select all</button>
                  <span style={{ color: BORDER }}>|</span>
                  <button onClick={() => setPrintSel({ summary: false, monthly: false, quarterly: false, user: false, strategy: false, cash: false })}
                    style={{ background: "transparent", border: "none", color: "var(--muted)", fontSize: 11, fontWeight: 700, cursor: "pointer", outline: "none" }}>Clear</button>
                </div>
                <div style={{ display: "flex", gap: 10 }}>
                  <button onClick={startPrint} disabled={!Object.values(printSel).some(Boolean)}
                    title="Use the browser print dialog instead"
                    style={{ background: "transparent", border: `1px solid ${BORDER}`,
                    color: "var(--muted)", borderRadius: 8, padding: "10px 16px", fontSize: 12, fontWeight: 700,
                    cursor: Object.values(printSel).some(Boolean) ? "pointer" : "not-allowed", outline: "none" }}>🖨️ Browser Print</button>
                  <button onClick={startReport} disabled={!Object.values(printSel).some(Boolean)}
                    style={{ background: "var(--violet)", border: "none", color: "#000", borderRadius: 8, padding: "10px 22px",
                      fontSize: 13, fontWeight: 800, cursor: Object.values(printSel).some(Boolean) ? "pointer" : "not-allowed",
                      outline: "none", opacity: Object.values(printSel).some(Boolean) ? 1 : 0.5 }}>
                    📄 Generate PDF
                  </button>
                </div>
              </div>
            </div>
          </div>
        )}

        {/* Export Modal overlay */}
        {showExportModal && (
          <div style={{
            position: "fixed",
            top: 0,
            left: 0,
            width: "100%",
            height: "100%",
            background: "rgba(4, 7, 13, 0.85)",
            backdropFilter: "blur(6px)",
            display: "flex",
            justifyContent: "center",
            alignItems: "center",
            zIndex: 1000,
            transition: "all 0.3s ease",
          }}>
            <div style={{
              background: CARD,
              border: `1px solid ${BORDER}`,
              borderRadius: 16,
              padding: 28,
              width: 540,
              maxWidth: "92%",
              boxShadow: "0 20px 45px rgba(0, 0, 0, 0.6), 0 0 30px rgba(0, 229, 255, 0.05)",
              animation: "modalFadeIn 0.2s ease-out",
            }}>
              <style>{`
                @keyframes modalFadeIn {
                  from { transform: scale(0.95); opacity: 0; }
                  to { transform: scale(1); opacity: 1; }
                }
              `}</style>
              
              {/* Modal Title */}
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 20 }}>
                <div>
                  <h3 style={{ fontSize: 18, fontWeight: 800, color: "var(--text)", letterSpacing: -0.5 }}>📥 Export Performance to Excel</h3>
                  <p style={{ color: "var(--muted)", fontSize: 11, marginTop: 4 }}>Select the account scope and specific columns to include in the spreadsheet.</p>
                </div>
                <button 
                  onClick={() => setShowExportModal(false)}
                  style={{
                    background: "transparent",
                    border: "none",
                    color: "#6b8cbb",
                    fontSize: 24,
                    cursor: "pointer",
                    padding: 4,
                    lineHeight: 1,
                    outline: "none"
                  }}
                  onMouseEnter={(e) => e.currentTarget.style.color = ACCENT}
                  onMouseLeave={(e) => e.currentTarget.style.color = "#6b8cbb"}
                >
                  &times;
                </button>
              </div>

              <div style={{ borderBottom: `1px solid ${BORDER}`, marginBottom: 20 }} />

              {/* Scope Section */}
              <div style={{ marginBottom: 22 }}>
                <label style={{ display: "block", color: ACCENT, fontSize: 10, fontWeight: 800, letterSpacing: 1.5, textTransform: "uppercase", marginBottom: 10 }}>1. Select Account Scope</label>
                <div style={{ display: "grid", gridTemplateColumns: "repeat(3, 1fr)", gap: 10 }}>
                  {[
                    { id: "all", label: "All Accounts", count: DATA_ALL.length },
                    { id: "fo", label: "F&O Only", count: DATA.length },
                    { id: "cash", label: "Cash & ETF Only", count: DATA_CASH.length },
                  ].map(opt => (
                    <button
                      key={opt.id}
                      onClick={() => setExportScope(opt.id)}
                      style={{
                        background: exportScope === opt.id ? "rgba(0, 229, 255, 0.12)" : "var(--bg)",
                        border: `1px solid ${exportScope === opt.id ? ACCENT : BORDER}`,
                        color: exportScope === opt.id ? ACCENT : "var(--muted)",
                        borderRadius: 8,
                        padding: "10px 8px",
                        cursor: "pointer",
                        fontWeight: 700,
                        fontSize: 12,
                        transition: "all 0.15s ease",
                        textAlign: "center",
                        outline: "none"
                      }}
                      onMouseEnter={(e) => {
                        if (exportScope !== opt.id) e.currentTarget.style.borderColor = "#6b8cbb";
                      }}
                      onMouseLeave={(e) => {
                        if (exportScope !== opt.id) e.currentTarget.style.borderColor = BORDER;
                      }}
                    >
                      <div style={{ fontSize: 11 }}>{opt.label}</div>
                      <div style={{ fontSize: 9, opacity: 0.7, marginTop: 2 }}>({opt.count} rows)</div>
                    </button>
                  ))}
                </div>
              </div>

              {/* Columns Selection */}
              <div style={{ marginBottom: 24 }}>
                <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 10 }}>
                  <label style={{ color: ACCENT, fontSize: 10, fontWeight: 800, letterSpacing: 1.5, textTransform: "uppercase" }}>2. Select Attributes</label>
                  <div style={{ display: "flex", gap: 8 }}>
                    <button
                      onClick={() => {
                        const updated = {};
                        EXPORT_COLUMNS.forEach(c => updated[c.key] = true);
                        setSelectedCols(updated);
                      }}
                      style={{ background: "transparent", border: "none", color: ACCENT, fontSize: 11, cursor: "pointer", fontWeight: 700, outline: "none" }}
                    >
                      Select All
                    </button>
                    <span style={{ color: "#1e2d4a" }}>|</span>
                    <button
                      onClick={() => {
                        const updated = {};
                        EXPORT_COLUMNS.forEach(c => updated[c.key] = c.key === "code");
                        setSelectedCols(updated);
                      }}
                      style={{ background: "transparent", border: "none", color: "#6b8cbb", fontSize: 11, cursor: "pointer", fontWeight: 700, outline: "none" }}
                    >
                      Deselect All
                    </button>
                  </div>
                </div>

                <div style={{
                  background: "var(--bg)",
                  border: `1px solid ${BORDER}`,
                  borderRadius: 8,
                  padding: 16,
                  display: "grid",
                  gridTemplateColumns: "1fr 1fr",
                  gap: "10px 16px",
                  maxHeight: 220,
                  overflowY: "auto",
                }}>
                  {EXPORT_COLUMNS.map(col => (
                    <label
                      key={col.key}
                      style={{
                        display: "flex",
                        alignItems: "center",
                        gap: 8,
                        fontSize: 12,
                        cursor: col.key === "code" ? "not-allowed" : "pointer",
                        color: col.key === "code" ? "var(--muted2)" : "var(--text)",
                        fontWeight: selectedCols[col.key] ? 600 : 400,
                        opacity: col.key === "code" ? 0.7 : 1,
                      }}
                    >
                      <input
                        type="checkbox"
                        checked={selectedCols[col.key]}
                        disabled={col.key === "code"}
                        onChange={(e) => {
                          if (col.key !== "code") {
                            setSelectedCols({
                              ...selectedCols,
                              [col.key]: e.target.checked
                            });
                          }
                        }}
                        style={{
                          accentColor: ACCENT,
                          cursor: col.key === "code" ? "not-allowed" : "pointer",
                          width: 14,
                          height: 14,
                        }}
                      />
                      {col.label}
                    </label>
                  ))}
                </div>
              </div>

              <div style={{ borderBottom: `1px solid ${BORDER}`, marginBottom: 22 }} />

              {/* Footer / Actions */}
              <div style={{ display: "flex", justifyContent: "flex-end", gap: 12 }}>
                <button
                  onClick={() => setShowExportModal(false)}
                  style={{
                    background: "transparent",
                    border: `1px solid ${BORDER}`,
                    color: "var(--muted)",
                    borderRadius: 8,
                    padding: "10px 20px",
                    fontSize: 13,
                    fontWeight: 700,
                    cursor: "pointer",
                    transition: "all 0.15s ease",
                    outline: "none"
                  }}
                  onMouseEnter={(e) => {
                    e.currentTarget.style.borderColor = "var(--muted)";
                    e.currentTarget.style.color = "var(--text)";
                  }}
                  onMouseLeave={(e) => {
                    e.currentTarget.style.borderColor = BORDER;
                    e.currentTarget.style.color = "var(--muted)";
                  }}
                >
                  Cancel
                </button>
                <button
                  onClick={handleExportExcel}
                  style={{
                    background: ACCENT,
                    border: "none",
                    color: "#000",
                    borderRadius: 8,
                    padding: "10px 24px",
                    fontSize: 13,
                    fontWeight: 800,
                    cursor: "pointer",
                    transition: "all 0.15s ease",
                    boxShadow: `0 4px 14px ${ACCENT}33`,
                    outline: "none"
                  }}
                  onMouseEnter={(e) => {
                    e.currentTarget.style.boxShadow = `0 6px 20px ${ACCENT}55`;
                    e.currentTarget.style.transform = "translateY(-1px)";
                  }}
                  onMouseLeave={(e) => {
                    e.currentTarget.style.boxShadow = `0 4px 14px ${ACCENT}33`;
                    e.currentTarget.style.transform = "translateY(0)";
                  }}
                >
                  Confirm & Export
                </button>
              </div>
            </div>
          </div>
        )}

        <div style={{ textAlign: "center" , color: "var(--border)" , fontSize: 11, marginTop: 20 }}>
            Quant Strategy FY 2026–27 · Data sourced from uploaded Excel · Q1 + Q2 (Apr–Sep) · Q3–Q4 pending
        </div>
    </div>
    );
    }


export default Dashboard
