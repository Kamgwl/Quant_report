import { useState, useMemo } from 'react'
import {
  BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer,
  LineChart, Line, Cell, ReferenceLine, Legend,
} from 'recharts'
import { RAW } from './data'
import ExcelJS from 'exceljs'
import { saveAs } from 'file-saver'

// ── RAW DATA ──────────────────────────────────────────────────────────────────


// ── CASH STRATEGY DETECTION ────────────────────────────────────────────────────
const CASH_STRATEGIES = ["CASH", "ETF", "NA"];
function isCashStrategy(r) {
const s = (r.strategy || "").trim().toUpperCase();
return CASH_STRATEGIES.includes(s) ||
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
const q1 = r.apr + r.may + r.jun;
const q1_roi = apr_roi + may_roi + jun_roi;
// Yearly (FY) = sum of all quarters available so far. Only Q1 exists now,
// so year == Q1. When Q2/Q3/Q4 P&L is added, include them here (e.g. q1 + q2 + q3 + q4).
const year = q1;
const year_roi = q1_roi;
const ann_roi = q1_roi * ANN_FACTOR;
const ann_pnl = q1 * ANN_FACTOR;
const strategy = (r.strategy || "").trim().toUpperCase();
return { ...r, strategy, apr_roi, may_roi, jun_roi, q1, q1_roi, year, year_roi, ann_roi, ann_pnl, isCash: isCashStrategy({ ...r, strategy }) };
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

const StatCard = ({ label, value, sub, color }) => (
<div style={{ background: CARD, border: `1px solid ${BORDER}`, borderRadius: 12, padding: "18px 22px" , borderTop: `3px
    solid ${color || ACCENT}`, }}>
    <div style={{ color: "var(--muted)" , fontSize: 11, fontWeight: 700, letterSpacing: 2, textTransform: "uppercase" ,
        marginBottom: 6 }}>{label}</div>
    <div style={{ color: color || "#fff" , fontSize: 24, fontWeight: 800, fontFamily: "'DM Mono', monospace" }}>{value}
    </div>
    {sub && <div style={{ color: "var(--muted2)" , fontSize: 12, marginTop: 4 }}>{sub}</div>}
</div>
);

const CustomTooltip = ({ active, payload, label }) => {
if (!active || !payload?.length) return null;
return (
<div style={{ background: "var(--tooltip-bg)" , border: `1px solid ${BORDER}`, borderRadius: 8, padding: "10px 14px" , fontSize:
    12 }}>
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
  { key: "ann_pnl", label: "Ann. P&L (Projected)", align: "right", numFmt: "+₹#,##,##0;-₹#,##,##0;\"—\"", colorCode: true },
  { key: "ann_roi", label: "Ann. ROI", align: "right", numFmt: "+0.00%;-0.00%;0.00%", colorCode: true },
];

function Dashboard({ theme = 'dark' }) {
const [view, setView] = useState("monthly"); // monthly | quarterly | user | strategy
const [selectedMonth, setSelectedMonth] = useState("q1");
const [sortBy, setSortBy] = useState("roi");

const [showExportModal, setShowExportModal] = useState(false);
const [exportScope, setExportScope] = useState("all"); // 'all' | 'fo' | 'cash'
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
  ann_pnl: true,
  ann_roi: true,
});

const handleExportExcel = async () => {
  try {
    let exportData = [];
    if (exportScope === "all") {
      exportData = DATA_ALL;
    } else if (exportScope === "fo") {
      exportData = DATA;
    } else if (exportScope === "cash") {
      exportData = DATA_CASH;
    }

    if (!ExcelJS) {
      alert("ExcelJS library not loaded.");
      return;
    }

    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet("Quant Q1 FY2026-27");

    const activeCols = EXPORT_COLUMNS.filter(c => selectedCols[c.key]);

    if (activeCols.length === 0) {
      alert("Please select at least one column to export.");
      return;
    }

    worksheet.columns = activeCols.map(col => ({
      header: col.label,
      key: col.key,
      width: 15,
      style: {
        font: { name: 'Segoe UI', size: 10 },
        alignment: { vertical: 'middle', horizontal: col.align }
      }
    }));

    const headerRow = worksheet.getRow(1);
    headerRow.height = 28;
    const isDark = theme === "dark";
    const headerBg = isDark ? "FF0D1526" : "FFF3F6FB";
    const headerFg = isDark ? "FF00E5FF" : "FF0091B3";
    const headerBorderColor = isDark ? "FF1E2D4A" : "FFD5DEEC";

    activeCols.forEach((colDef, idx) => {
      const cell = headerRow.getCell(idx + 1);
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: headerBg }
      };
      cell.font = {
        name: 'Segoe UI',
        size: 11,
        bold: true,
        color: { argb: headerFg }
      };
      cell.border = {
        bottom: { style: 'medium', color: { argb: headerBorderColor } }
      };
      cell.alignment = {
        vertical: 'middle',
        horizontal: colDef.align
      };
    });

    exportData.forEach((row) => {
      const rowData = {};
      activeCols.forEach(col => {
        rowData[col.key] = row[col.key] ?? null;
      });

      const addedRow = worksheet.addRow(rowData);
      addedRow.height = 20;

      activeCols.forEach((colDef, idx) => {
        const cell = addedRow.getCell(idx + 1);
        const val = row[colDef.key];

        cell.border = {
          top: { style: 'thin', color: { argb: 'FFE0E0E0' } },
          left: { style: 'thin', color: { argb: 'FFE0E0E0' } },
          bottom: { style: 'thin', color: { argb: 'FFE0E0E0' } },
          right: { style: 'thin', color: { argb: 'FFE0E0E0' } }
        };

        if (colDef.numFmt) {
          cell.numFmt = colDef.numFmt;
        }

        if (colDef.colorCode && typeof val === 'number') {
          if (val > 0) {
            cell.fill = {
              type: 'pattern',
              pattern: 'solid',
              fgColor: { argb: 'FFE6F4EA' }
            };
            cell.font = {
              name: 'Segoe UI',
              size: 10,
              color: { argb: 'FF137333' },
              bold: true
            };
          } else if (val < 0) {
            cell.fill = {
              type: 'pattern',
              pattern: 'solid',
              fgColor: { argb: 'FFFCE8E6' }
            };
            cell.font = {
              name: 'Segoe UI',
              size: 10,
              color: { argb: 'FFC5221F' },
              bold: true
            };
          }
        }
      });
    });

    activeCols.forEach((colDef, idx) => {
      const column = worksheet.getColumn(idx + 1);
      let maxLen = 12;
      column.eachCell({ includeEmpty: true }, cell => {
        let valStr = '';
        if (cell.value !== null && cell.value !== undefined) {
          if (typeof cell.value === 'object' && cell.value.text) {
            valStr = cell.value.text.toString();
          } else {
            valStr = cell.value.toString();
          }
        }
        if (colDef.key && colDef.key.endsWith('_roi')) {
          valStr += '% +';
        } else if (colDef.key === 'fund') {
          valStr += ' Cr ₹';
        } else if (['apr', 'may', 'jun', 'q1', 'ann_pnl'].includes(colDef.key)) {
          valStr = '₹+' + valStr + ',000';
        }
        
        if (valStr.length > maxLen) {
          maxLen = valStr.length;
        }
      });
      column.width = Math.min(maxLen + 4, 30);
    });

    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
    saveAs(blob, `Quant_Q1_FY2026-27_${exportScope.toUpperCase()}.xlsx`);

    setShowExportModal(false);
  } catch (error) {
    console.error("Export failed", error);
    alert("An error occurred during export: " + error.message);
  }
};

// ── Summary stats ──
const totals = useMemo(() => {
const totalFund = DATA.reduce((s, r) => s + r.fund, 0);
const aprPnL = DATA.reduce((s, r) => s + r.apr, 0);
const mayPnL = DATA.reduce((s, r) => s + r.may, 0);
const junPnL = DATA.reduce((s, r) => s + r.jun, 0);
const q1PnL = aprPnL + mayPnL + junPnL;
const q1ROI = totalFund > 0 ? q1PnL / (totalFund * 1e7) : 0;
const annROI = q1ROI * ANN_FACTOR;
const annPnL = q1PnL * ANN_FACTOR;
const winners = DATA.filter(r => r.q1 > 0).length;
const losers = DATA.filter(r => r.q1 < 0).length; return { totalFund, aprPnL, mayPnL, junPnL, q1PnL, q1ROI, annROI,
    annPnL, winners, losers }; }, []); // ── Month bar data (top 20 by absolute PnL) ──
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
    .filter(r => r.q1 !== 0)
    .sort((a, b) => b.q1_roi - a.q1_roi), []);

    const top5 = ranked.slice(0, 5);
    const bot5 = ranked.slice(-5).reverse();

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

    // ── Monthly trend ──
    const trendData = [
    { month: "April", pnl: totals.aprPnL, roi: totals.totalFund > 0 ? totals.aprPnL / (totals.totalFund * 1e7) : 0 },
    { month: "May", pnl: totals.mayPnL, roi: totals.totalFund > 0 ? totals.mayPnL / (totals.totalFund * 1e7) : 0 },
    { month: "June", pnl: totals.junPnL, roi: totals.totalFund > 0 ? totals.junPnL / (totals.totalFund * 1e7) : 0 },
    ];

    const tabs = [
    { id: "monthly", label: "Monthly Report" },
    { id: "quarterly", label: "Q1 Summary" },
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

        {/* Header */}
        <div style={{ background: "linear-gradient(135deg, #0d1526 0%, #0a1020 100%)" , borderBottom: `1px solid
            ${BORDER}`, padding: "24px 32px 20px" }}>
            <div style={{ display: "flex" , alignItems: "center" , justifyContent: "space-between" , flexWrap: "wrap" ,
                gap: 12 }}>
                <div>
                    <div style={{ display: "flex" , alignItems: "center" , gap: 12, marginBottom: 4 }}>
                        <div style={{ width: 8, height: 36, background: ACCENT, borderRadius: 4 }} />
                        <div>
                            <h1 style={{ margin: 0, fontSize: 26, fontWeight: 800, letterSpacing: -0.5 }}>Quant Strategy
                                Dashboard</h1>
                            <div style={{ color: "var(--muted2)" , fontSize: 12, marginTop: 2 }}>FY 2026–27 · Q1 Performance
                                Report · April – June 2026</div>
                        </div>
                    </div>
                </div>
                <div style={{ display: "flex" , gap: 12, alignItems: "center" }}>
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

                    <div style={{ background: totals.q1PnL>= 0 ? "rgba(34,197,94,0.15)" : "rgba(244,63,94,0.15)",
                        border: `1px solid ${totals.q1PnL >= 0 ? POS : NEG}`, borderRadius: 8, padding: "8px 16px",
                        textAlign: "center" }}>
                        <div style={{ fontSize: 10, color: "var(--muted)" , fontWeight: 700, letterSpacing: 2 }}>Q1 NET P&L
                        </div>
                        <div style={{ fontSize: 20, fontWeight: 800, color: totals.q1PnL>= 0 ? POS : NEG, fontFamily:
                            "'DM Mono', monospace" }}>{fmtSign(totals.q1PnL)}</div>
                    </div>
                </div>
            </div>

            {/* KPI Row */}
            <div style={{ display: "grid" , gridTemplateColumns: "repeat(auto-fit, minmax(160px, 1fr))" , gap: 12,
                marginTop: 20 }}>
                <StatCard label="Total AUM" value={fmtFund(totals.totalFund)}
                    sub="Total Utilized Fund" color={ACCENT} />
                <StatCard label="Q1 Net ROI" value={fmtROI(totals.q1ROI)} sub="Apr + May + Jun (3 months)"
                    color={totals.q1ROI>= 0 ? POS : NEG} />
                    <StatCard label="Q1 Net P&L" value={fmtSign(totals.q1PnL)} sub="April – June 2026"
                        color={totals.q1PnL>= 0 ? POS : NEG} />
                            <StatCard label="April P&L" value={fmtSign(totals.aprPnL)} sub={fmtROI(totals.totalFund > 0 ? totals.aprPnL /
                                (totals.totalFund * 1e7) : 0)} color={totals.aprPnL>= 0 ? POS : NEG} />
                                <StatCard label="May P&L" value={fmtSign(totals.mayPnL)} sub={fmtROI(totals.totalFund > 0 ? totals.mayPnL /
                                    (totals.totalFund * 1e7) : 0)} color={totals.mayPnL>= 0 ? POS : NEG} />
                                    <StatCard label="June P&L" value={fmtSign(totals.junPnL)} sub={fmtROI(totals.totalFund > 0 ? totals.junPnL
                                        / (totals.totalFund * 1e7) : 0)} color={totals.junPnL>= 0 ? POS : NEG} />
                                        <StatCard label="Winners / Losers" value={`${totals.winners} /
                                            ${totals.losers}`} sub={`${DATA.length} total accounts`} color="var(--violet)" />
            </div>
        </div>

        {/* Tabs */}
        <div style={{ padding: "16px 32px 0" , display: "flex" , gap: 6, borderBottom: `1px solid ${BORDER}`,
            background: "var(--panel)" }}>
            {tabs.map(t => (
            <button key={t.id} style={tabStyle(t.id)} onClick={()=> setView(t.id)}>{t.label}</button>
            ))}
        </div>

        <div style={{ padding: "28px 32px" }}>

            {/* ── MONTHLY ── */}
            {view === "monthly" && (
            <div>
                {sectionHead("Monthly P&L Trend")}
                <div style={{ display: "grid" , gridTemplateColumns: "1fr 1fr" , gap: 24, marginBottom: 32 }}>
                    <div style={{ background: CARD, borderRadius: 12, padding: 20, border: `1px solid ${BORDER}` }}>
                        <div style={{ color: "var(--muted)" , fontSize: 12, fontWeight: 700, marginBottom: 14 }}>MONTHLY NET
                            P&L (₹)</div>
                        <ResponsiveContainer width="100%" height={220}>
                            <BarChart data={trendData} barSize={50}>
                                <CartesianGrid strokeDasharray="3 3" stroke={BORDER} />
                                <XAxis dataKey="month" tick={{ fill: "var(--muted)" , fontSize: 12 }} axisLine={false}
                                    tickLine={false} />
                                <YAxis tick={{ fill: "var(--muted)" , fontSize: 11 }} tickFormatter={v=> fmt(v)}
                                    axisLine={false} tickLine={false} />
                                    <Tooltip content={<CustomTooltip />} />
                                    <ReferenceLine y={0} stroke={BORDER} />
                                    <Bar dataKey="pnl" name="Net P&L" radius={[6, 6, 0, 0]}>
                                        {trendData.map((d, i) => <Cell key={i} fill={d.pnl>= 0 ? POS : NEG} />)}
                                    </Bar>
                            </BarChart>
                        </ResponsiveContainer>
                    </div>
                    <div style={{ background: CARD, borderRadius: 12, padding: 20, border: `1px solid ${BORDER}` }}>
                        <div style={{ color: "var(--muted)" , fontSize: 12, fontWeight: 700, marginBottom: 14 }}>MONTHLY ROI
                            (%)</div>
                        <ResponsiveContainer width="100%" height={220}>
                            <LineChart data={trendData}>
                                <CartesianGrid strokeDasharray="3 3" stroke={BORDER} />
                                <XAxis dataKey="month" tick={{ fill: "var(--muted)" , fontSize: 12 }} axisLine={false}
                                    tickLine={false} />
                                <YAxis tick={{ fill: "var(--muted)" , fontSize: 11 }} tickFormatter={v=> `${(v*100).toFixed(2)}%`}
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
                    {[{ id: "apr", label: "April" }, { id: "may", label: "May" }, { id: "jun", label: "June" }, { id:
                    "q1", label: "Q1 Total" }, { id: "year", label: "Yearly (FY)" }].map(m => (
                    <button key={m.id} onClick={()=> setSelectedMonth(m.id)} style={{
                        padding: "6px 16px", borderRadius: 6, border: `1px solid ${BORDER}`, cursor: "pointer",
                        fontSize: 12, fontWeight: 700,
                        background: selectedMonth === m.id ? ACCENT : CARD, color: selectedMonth === m.id ? "#000" :
                        "var(--muted)",
                        }}>{m.label}</button>
                    ))}
                    <div style={{ marginLeft: "auto" , display: "flex" , gap: 8 }}>
                        <span style={{ color: "var(--muted)" , fontSize: 12, alignSelf: "center" }}>Sort:</span>
                        {["roi", "pnl"].map(s => (
                        <button key={s} onClick={()=> setSortBy(s)} style={{
                            padding: "6px 14px", borderRadius: 6, border: `1px solid ${BORDER}`, cursor: "pointer",
                            fontSize: 12,
                            background: sortBy === s ? "var(--border)" : "transparent", color: sortBy === s ? ACCENT :
                            "var(--muted)",
                            }}>{s.toUpperCase()}</button>
                        ))}
                    </div>
                </div>

                <div style={{ background: CARD, borderRadius: 12, padding: 20, border: `1px solid ${BORDER}` }}>
                    <ResponsiveContainer width="100%" height={420}>
                        <BarChart data={barData} layout="vertical" barSize={16} margin={{ left: 160 }}>
                            <CartesianGrid strokeDasharray="3 3" stroke={BORDER} horizontal={false} />
                            <XAxis type="number" tick={{ fill: "var(--muted)" , fontSize: 11 }} tickFormatter={v=> fmt(v)}
                                axisLine={false} tickLine={false} />
                                <YAxis type="category" dataKey="name" tick={{ fill: "var(--axis)" , fontSize: 10 }}
                                    interval={0} width={160} axisLine={false} tickLine={false} />
                                <Tooltip content={({ active, payload, label })=> {
                                    if (!active || !payload?.length) return null;
                                    const d = payload[0]?.payload;
                                    return (
                                    <div style={{ background: "var(--tooltip-bg)" , border: `1px solid ${BORDER}`, borderRadius:
                                        8, padding: "10px 14px" , fontSize: 12 }}>
                                        <div style={{ color: ACCENT, fontWeight: 700, marginBottom: 4 }}>{d?.fullName}
                                        </div>
                                        <div style={{ color: d?.pnl>= 0 ? POS : NEG }}>P&L: {fmt(d?.pnl)}</div>
                                        <div style={{ color: d?.roi>= 0 ? POS : NEG }}>ROI: {d?.roi?.toFixed(2)}%
                                        </div>
                                        <div style={{ color: "var(--muted)" }}>Fund: ₹{d?.fund} Cr</div>
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
            {view === "quarterly" && (
            <div>
                {sectionHead("Q1 2026–27 Consolidated Report (Apr + May + Jun)")}

                {/* Best vs Worst */}
                <div style={{ display: "grid" , gridTemplateColumns: "1fr 1fr" , gap: 24, marginBottom: 28 }}>
                    {/* Top 5 */}
                    <div style={{ background: CARD, borderRadius: 12, padding: 20, border: `1px solid ${BORDER}` }}>
                        <div style={{ color: POS, fontSize: 12, fontWeight: 800, letterSpacing: 2, marginBottom: 14 }}>
                            🏆 TOP 5 PERFORMERS — Q1 ROI</div>
                        {top5.map((r, i) => (
                        <div key={r.code} style={{ display: "flex" , justifyContent: "space-between" ,
                            alignItems: "center" , padding: "10px 0" , borderBottom: i < 4 ? `1px solid ${BORDER}`
                            : "none" }}>
                            <div style={{ display: "flex" , gap: 10, alignItems: "center" }}>
                                <div style={{ width: 26, height: 26, borderRadius: "50%" , background: POS,
                                    color: "#000" , fontWeight: 800, fontSize: 12, display: "flex" ,
                                    alignItems: "center" , justifyContent: "center" }}>{i + 1}</div>
                                <div>
                                    <div style={{ fontWeight: 700, fontSize: 13 }}>{r.name}</div>
                                    <div style={{ color: "var(--muted2)" , fontSize: 11 }}>{r.strategy || "—"} · ₹{r.fund}L
                                    </div>
                                </div>
                            </div>
                            <div style={{ textAlign: "right" }}>
                                <div style={{ color: POS, fontWeight: 800, fontFamily: "'DM Mono', monospace" ,
                                    fontSize: 14 }}>{fmtROI(r.q1_roi)}</div>
                                <div style={{ color: "var(--muted)" , fontSize: 11 }}>{fmtSign(r.q1)}</div>
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
                            alignItems: "center" , padding: "10px 0" , borderBottom: i < 4 ? `1px solid ${BORDER}`
                            : "none" }}>
                            <div style={{ display: "flex" , gap: 10, alignItems: "center" }}>
                                <div style={{ width: 26, height: 26, borderRadius: "50%" , background: NEG,
                                    color: "#fff" , fontWeight: 800, fontSize: 12, display: "flex" ,
                                    alignItems: "center" , justifyContent: "center" }}>{i + 1}</div>
                                <div>
                                    <div style={{ fontWeight: 700, fontSize: 13 }}>{r.name}</div>
                                    <div style={{ color: "var(--muted2)" , fontSize: 11 }}>{r.strategy || "—"} · ₹{r.fund}L
                                    </div>
                                </div>
                            </div>
                            <div style={{ textAlign: "right" }}>
                                <div style={{ color: NEG, fontWeight: 800, fontFamily: "'DM Mono', monospace" ,
                                    fontSize: 14 }}>{fmtROI(r.q1_roi)}</div>
                                <div style={{ color: "var(--muted)" , fontSize: 11 }}>{fmt(r.q1)}</div>
                            </div>
                        </div>
                        ))}
                    </div>
                </div>

                {/* Full Q1 bar */}
                {sectionHead("Full Q1 P&L & ROI — Best to Worst (All Users)")}
                <div style={{ background: CARD, borderRadius: 12, padding: 20, border: `1px solid ${BORDER}` }}>
                    <ResponsiveContainer width="100%" height={Math.max(500, ranked.length * 22)}>
                        <BarChart data={ranked.map(r=> ({ name: r.name,
                            pnl: r.q1, roi: +(r.q1_roi * 100).toFixed(2), fund: r.fund, fullName: r.name }))}
                            layout="vertical" barSize={14} margin={{ left: 165 }}>
                            <CartesianGrid strokeDasharray="3 3" stroke={BORDER} horizontal={false} />
                            <XAxis type="number" tick={{ fill: "var(--muted)" , fontSize: 11 }} tickFormatter={v=> fmt(v)}
                                axisLine={false} tickLine={false} />
                                <YAxis type="category" dataKey="name" tick={{ fill: "var(--axis)" , fontSize: 10 }}
                                    interval={0} width={165} axisLine={false} tickLine={false} />
                                <Tooltip content={({ active, payload })=> {
                                    if (!active || !payload?.length) return null;
                                    const d = payload[0]?.payload;
                                    return (
                                    <div style={{ background: "var(--tooltip-bg)" , border: `1px solid ${BORDER}`, borderRadius:
                                        8, padding: "10px 14px" , fontSize: 12 }}>
                                        <div style={{ color: ACCENT, fontWeight: 700 }}>{d?.fullName}</div>
                                        <div style={{ color: d?.pnl>= 0 ? POS : NEG }}>Q1 P&L: {fmt(d?.pnl)}</div>
                                        <div style={{ color: d?.roi>= 0 ? POS : NEG }}>Q1 ROI: {d?.roi?.toFixed(2)}%
                                        </div>
                                    </div>
                                    );
                                    }} />
                                    <ReferenceLine x={0} stroke="var(--border)" />
                                    <Bar dataKey="pnl" name="Q1 P&L" radius={[0, 4, 4, 0]}>
                                        {ranked.map((d, i) => <Cell key={i} fill={d.q1>= 0 ? POS : NEG} />)}
                                    </Bar>
                        </BarChart>
                    </ResponsiveContainer>
                </div>
            </div>
            )}

            {/* ── USER PERFORMANCE ── */}
            {view === "user" && (
            <div>
                {sectionHead("User-Wise Performance Table")}
                <div style={{ background: CARD, borderRadius: 12, border: `1px solid ${BORDER}`, overflow: "auto" }}>
                    <table style={{ width: "100%" , borderCollapse: "collapse" , fontSize: 12 }}>
                        <thead>
                            <tr style={{ background: "var(--panel)" , borderBottom: `1px solid ${BORDER}` }}>
                                {["#", "Code", "Name", "Strategy", "Fund (Cr)", "Apr P&L", "Apr ROI", "May P&L", "May ROI", "Jun P&L", "Jun ROI", "Q1 P&L", "Q1 ROI"].map(h => (
                                <th key={h} style={{ padding: "12px 10px" , textAlign: "left" , color: "var(--muted)" ,
                                    fontWeight: 700, letterSpacing: 1, whiteSpace: "nowrap" }}>{h}</th>
                                ))}
                            </tr>
                        </thead>
                        <tbody>
                            {ranked.map((r, i) => {
                            const rowBg = i % 2 === 0 ? CARD : "var(--row-alt)";
                            return (
                            <tr key={r.code} style={{ background: rowBg, borderBottom: `1px solid ${BORDER}22` }}>
                                <td style={{ padding: "10px" , color: "var(--muted2)" }}>{i + 1}</td>
                                <td style={{ padding: "10px" , color: ACCENT, fontFamily: "'DM Mono', monospace" ,
                                    fontSize: 11 }}>{r.code}</td>
                                <td style={{ padding: "10px" , fontWeight: 600, whiteSpace: "nowrap" }}>{r.name}</td>
                                <td style={{ padding: "10px" , color: "var(--muted)" , fontSize: 11 }}>{r.strategy || "—"}
                                </td>
                                <td style={{ padding: "10px" , fontFamily: "'DM Mono', monospace" , color: "var(--muted2)" }}>
                                    {r.fund}</td>
                                {[
                                [r.apr, r.apr_roi],
                                [r.may, r.may_roi],
                                [r.jun, r.jun_roi],
                                [r.q1, r.q1_roi],
                                ].map(([pnl, roi], j) => (
                                <>
                                    <td key={`pnl${j}`} style={{ padding: "10px" , fontFamily: "'DM Mono', monospace" ,
                                        color: pnl>= 0 ? POS : NEG, whiteSpace: "nowrap" }}>{pnl === 0 ? "—" :
                                        fmtSign(pnl)}</td>
                                    <td key={`roi${j}`} style={{ padding: "10px" , fontFamily: "'DM Mono', monospace" ,
                                        color: roi>= 0 ? POS : NEG }}>{roi === 0 ? "—" : fmtROI(roi)}</td>
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
            {view === "strategy" && (
            <div>
                {sectionHead("Strategy-Wise Q1 Performance")}
                <div style={{ display: "grid" , gridTemplateColumns: "1fr 1fr" , gap: 24, marginBottom: 28 }}>
                    <div style={{ background: CARD, borderRadius: 12, padding: 20, border: `1px solid ${BORDER}` }}>
                        <div style={{ color: "var(--muted)" , fontSize: 12, fontWeight: 700, marginBottom: 14 }}>STRATEGY NET
                            P&L — Q1</div>
                        <ResponsiveContainer width="100%" height={320}>
                            <BarChart data={strategyData.slice(0, 14)} layout="vertical" barSize={18} margin={{ left: 120
                                }}>
                                <CartesianGrid strokeDasharray="3 3" stroke={BORDER} horizontal={false} />
                                <XAxis type="number" tick={{ fill: "var(--muted)" , fontSize: 10 }} tickFormatter={v=>
                                    fmt(v)} axisLine={false} tickLine={false} />
                                    <YAxis type="category" dataKey="strategy" tick={{ fill: "var(--axis)" , fontSize: 10 }}
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
                        <div style={{ color: "var(--muted)" , fontSize: 12, fontWeight: 700, marginBottom: 14 }}>STRATEGY ROI
                            % — Q1</div>
                        <ResponsiveContainer width="100%" height={320}>
                            <BarChart data={strategyData.filter(d=> d.fund > 0).slice(0, 14)} layout="vertical"
                                barSize={18} margin={{ left: 120 }}>
                                <CartesianGrid strokeDasharray="3 3" stroke={BORDER} horizontal={false} />
                                <XAxis type="number" tick={{ fill: "var(--muted)" , fontSize: 10 }} tickFormatter={v=>
                                    `${v}%`} axisLine={false} tickLine={false} />
                                    <YAxis type="category" dataKey="strategy" tick={{ fill: "var(--axis)" , fontSize: 10 }}
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
                    <table style={{ width: "100%" , borderCollapse: "collapse" , fontSize: 12 }}>
                        <thead>
                            <tr style={{ background: "var(--panel)" , borderBottom: `1px solid ${BORDER}` }}>
                                {["Strategy", "Accounts", "Fund (Cr)", "Q1 Net P&L", "Q1 ROI", "Status"].map(h => (
                                <th key={h} style={{ padding: "12px 14px" , textAlign: "left" , color: "var(--muted)" ,
                                    fontWeight: 700, letterSpacing: 1 }}>{h}</th>
                                ))}
                            </tr>
                        </thead>
                        <tbody>
                            {strategyData.map((s, i) => (
                            <tr key={s.strategy} style={{ background: i % 2===0 ? CARD : "var(--row-alt)" , borderBottom: `1px
                                solid ${BORDER}22` }}>
                                <td style={{ padding: "10px 14px" , fontWeight: 700 }}>{s.strategy}</td>
                                <td style={{ padding: "10px 14px" , color: "var(--muted2)" }}>{s.count}</td>
                                <td style={{ padding: "10px 14px" , fontFamily: "'DM Mono', monospace" ,
                                    color: "var(--muted2)" }}>₹{s.fund.toFixed(1)}</td>
                                <td style={{ padding: "10px 14px" , fontFamily: "'DM Mono', monospace" , color: s.pnl>=
                                    0 ? POS : NEG, fontWeight: 700 }}>{fmtSign(s.pnl)}</td>
                                <td style={{ padding: "10px 14px" , fontFamily: "'DM Mono', monospace" , color: s.roi>=
                                    0 ? POS : NEG }}>{s.fund > 0 ? `${s.roi.toFixed(2)}%` : "—"}</td>
                                <td style={{ padding: "10px 14px" }}>
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
            {view === "cash" && (() => {
            const cashRanked = [...DATA_CASH].filter(r => r.q1 !== 0).sort((a, b) => b.q1_roi - a.q1_roi);
            const cashTotalFund = DATA_CASH.reduce((s, r) => s + r.fund, 0);
            const cashApr = DATA_CASH.reduce((s, r) => s + r.apr, 0);
            const cashMay = DATA_CASH.reduce((s, r) => s + r.may, 0);
            const cashJun = DATA_CASH.reduce((s, r) => s + r.jun, 0);
            const cashQ1 = cashApr + cashMay + cashJun;
            const cashQ1ROI = cashTotalFund > 0 ? cashQ1 / (cashTotalFund * 1e7) : 0;
            const cashTrend = [
            { month: "April", pnl: cashApr, roi: cashTotalFund > 0 ? cashApr / (cashTotalFund * 1e7) : 0 },
            { month: "May", pnl: cashMay, roi: cashTotalFund > 0 ? cashMay / (cashTotalFund * 1e7) : 0 },
            { month: "June", pnl: cashJun, roi: cashTotalFund > 0 ? cashJun / (cashTotalFund * 1e7) : 0 },
            ];
            return (
            <div>
                {/* Notice banner */}
                <div style={{ background: "rgba(234,179,8,0.08)" , border: "1px solid rgba(234,179,8,0.3)" ,
                    borderRadius: 10, padding: "12px 18px" , marginBottom: 24, display: "flex" , alignItems: "center" ,
                    gap: 10 }}>
                    <span style={{ fontSize: 18 }}>💰</span>
                    <span style={{ color: "var(--gold)" , fontSize: 13 }}>
                        Cash / ETF / ATS accounts are shown <strong>separately</strong> here and excluded from all other
                        reports to avoid skewing F&O strategy numbers.
                    </span>
                </div>

                {/* KPIs */}
                <div style={{ display: "grid" , gridTemplateColumns: "repeat(auto-fit, minmax(150px, 1fr))" , gap: 12,
                    marginBottom: 24 }}>
                    <StatCard label="Cash AUM" value={fmtFund(cashTotalFund)}
                        sub="Total Utilized Fund" color="var(--gold)" />
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
                        <div style={{ color: "var(--muted)" , fontSize: 12, fontWeight: 700, marginBottom: 14 }}>MONTHLY NET
                            P&L (₹)</div>
                        <ResponsiveContainer width="100%" height={200}>
                            <BarChart data={cashTrend} barSize={50}>
                                <CartesianGrid strokeDasharray="3 3" stroke={BORDER} />
                                <XAxis dataKey="month" tick={{ fill: "var(--muted)" , fontSize: 12 }} axisLine={false}
                                    tickLine={false} />
                                <YAxis tick={{ fill: "var(--muted)" , fontSize: 11 }} tickFormatter={v=> fmt(v)}
                                    axisLine={false} tickLine={false} />
                                    <Tooltip content={<CustomTooltip />} />
                                    <ReferenceLine y={0} stroke={BORDER} />
                                    <Bar dataKey="pnl" name="Net P&L" radius={[6,6,0,0]}>
                                        {cashTrend.map((d, i) => <Cell key={i} fill={d.pnl>= 0 ? "var(--gold)" : NEG} />)}
                                    </Bar>
                            </BarChart>
                        </ResponsiveContainer>
                    </div>
                    <div style={{ background: CARD, borderRadius: 12, padding: 20, border: `1px solid ${BORDER}` }}>
                        <div style={{ color: "var(--muted)" , fontSize: 12, fontWeight: 700, marginBottom: 14 }}>MONTHLY ROI
                            (%)</div>
                        <ResponsiveContainer width="100%" height={200}>
                            <LineChart data={cashTrend}>
                                <CartesianGrid strokeDasharray="3 3" stroke={BORDER} />
                                <XAxis dataKey="month" tick={{ fill: "var(--muted)" , fontSize: 12 }} axisLine={false}
                                    tickLine={false} />
                                <YAxis tick={{ fill: "var(--muted)" , fontSize: 11 }} tickFormatter={v=> `${(v*100).toFixed(2)}%`}
                                    axisLine={false} tickLine={false} />
                                    <Tooltip content={<CustomTooltip />} />
                                    <ReferenceLine y={0} stroke={BORDER} />
                                    <Line dataKey="roi" name="ROI %" stroke="var(--gold)" strokeWidth={3} dot={{
                                        fill: "var(--gold)" , r: 6 }} />
                            </LineChart>
                        </ResponsiveContainer>
                    </div>
                </div>

                {/* Q1 bar chart */}
                {sectionHead("Cash / ETF — Best to Worst (Q1)")}
                <div style={{ background: CARD, borderRadius: 12, padding: 20, border: `1px solid ${BORDER}`,
                    marginBottom: 28 }}>
                    <ResponsiveContainer width="100%" height={Math.max(200, cashRanked.length * 38)}>
                        <BarChart data={cashRanked.map(r=> ({ name: r.name,
                            pnl: r.q1, roi: +(r.q1_roi*100).toFixed(2), fund: r.fund, fullName: r.name,
                            strategy: r.strategy }))} layout="vertical" barSize={20} margin={{ left: 165 }}>
                            <CartesianGrid strokeDasharray="3 3" stroke={BORDER} horizontal={false} />
                            <XAxis type="number" tick={{ fill: "var(--muted)" , fontSize: 11 }} tickFormatter={v=> fmt(v)}
                                axisLine={false} tickLine={false} />
                                <YAxis type="category" dataKey="name" tick={{ fill: "var(--axis)" , fontSize: 10 }}
                                    interval={0} width={165} axisLine={false} tickLine={false} />
                                <Tooltip content={({ active, payload })=> {
                                    if (!active || !payload?.length) return null;
                                    const d = payload[0]?.payload;
                                    return (
                                    <div style={{ background: "var(--tooltip-bg)" , border: `1px solid ${BORDER}`, borderRadius:
                                        8, padding: "10px 14px" , fontSize: 12 }}>
                                        <div style={{ color: "var(--gold)" , fontWeight: 700 }}>{d?.fullName}</div>
                                        <div style={{ color: "var(--muted)" , fontSize: 11 }}>{d?.strategy || "—"} ·
                                            ₹{d?.fund}Cr</div>
                                        <div style={{ color: d?.pnl>= 0 ? POS : NEG }}>Q1 P&L: {fmt(d?.pnl)}</div>
                                        <div style={{ color: d?.roi>= 0 ? POS : NEG }}>Q1 ROI: {d?.roi?.toFixed(2)}%
                                        </div>
                                    </div>
                                    );
                                    }} />
                                    <ReferenceLine x={0} stroke={BORDER} />
                                    <Bar dataKey="pnl" name="Q1 P&L" radius={[0,4,4,0]}>
                                        {cashRanked.map((d, i) => <Cell key={i} fill={d.q1>= 0 ? "var(--gold)" : NEG} />)}
                                    </Bar>
                        </BarChart>
                    </ResponsiveContainer>
                </div>

                {/* Detail table */}
                {sectionHead("Cash / ETF Account Detail")}
                <div style={{ background: CARD, borderRadius: 12, border: `1px solid ${BORDER}`, overflow: "auto" }}>
                    <table style={{ width: "100%" , borderCollapse: "collapse" , fontSize: 12 }}>
                        <thead>
                            <tr style={{ background: "var(--panel)" , borderBottom: `1px solid ${BORDER}` }}>
                                {["#","Code","Name","Strategy","Fund (Cr)","Apr P&L","May P&L","Jun P&L","Q1 P&L","Q1 ROI"].map(h => (
                                <th key={h} style={{ padding:"12px 10px", textAlign:"left", color:"var(--muted)",
                                    fontWeight:700, letterSpacing:1, whiteSpace:"nowrap" }}>{h}</th>
                                ))}
                            </tr>
                        </thead>
                        <tbody>
                            {cashRanked.map((r, i) => (
                            <tr key={r.code} style={{ background: i % 2===0 ? CARD : "var(--row-alt)" , borderBottom: `1px
                                solid ${BORDER}22` }}>
                                <td style={{ padding:"10px", color:"var(--muted2)" }}>{i+1}</td>
                                <td style={{ padding:"10px", color:"var(--gold)", fontFamily:"'DM Mono',monospace",
                                    fontSize:11 }}>{r.code}</td>
                                <td style={{ padding:"10px", fontWeight:600 }}>{r.name}</td>
                                <td style={{ padding:"10px", color:"var(--muted)", fontSize:11 }}>{r.strategy || "—"}</td>
                                <td style={{ padding:"10px", fontFamily:"'DM Mono',monospace", color:"var(--muted2)" }}>
                                    {r.fund}</td>
                                {[r.apr, r.may, r.jun, r.q1].map((v, j) => (
                                <td key={j} style={{ padding:"10px", fontFamily:"'DM Mono',monospace", color: v>= 0 ?
                                    POS : NEG, whiteSpace:"nowrap" }}>{v === 0 ? "—" : fmtSign(v)}</td>
                                ))}
                                <td style={{ padding:"10px", fontFamily:"'DM Mono',monospace", color: r.q1_roi>= 0 ? POS
                                    : NEG }}>{fmtROI(r.q1_roi)}</td>
                            </tr>
                            ))}
                        </tbody>
                    </table>
                </div>
            </div>
            );
            })()}

        </div>

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
                  <h3 style={{ fontSize: 18, fontWeight: 800, color: "#fff", letterSpacing: -0.5 }}>📥 Export Performance to Excel</h3>
                  <p style={{ color: "var(--muted)", fontSize: 11, marginTop: 4 }}>Select the account scope and specific columns to include in the spreadsheet.</p>
                </div>
                <button 
                  onClick={() => setShowExportModal(false)}
                  style={{
                    background: "transparent",
                    border: "none",
                    color: "var(--muted)",
                    fontSize: 24,
                    cursor: "pointer",
                    padding: 4,
                    lineHeight: 1,
                    outline: "none"
                  }}
                  onMouseEnter={(e) => e.currentTarget.style.color = ACCENT}
                  onMouseLeave={(e) => e.currentTarget.style.color = "var(--muted)"}
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
                        background: exportScope === opt.id ? "rgba(0, 229, 255, 0.12)" : "#070b14",
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
                        if (exportScope !== opt.id) e.currentTarget.style.borderColor = "var(--muted)";
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
                    <span style={{ color: "var(--border)" }}>|</span>
                    <button
                      onClick={() => {
                        const updated = {};
                        EXPORT_COLUMNS.forEach(c => updated[c.key] = c.key === "code");
                        setSelectedCols(updated);
                      }}
                      style={{ background: "transparent", border: "none", color: "var(--muted)", fontSize: 11, cursor: "pointer", fontWeight: 700, outline: "none" }}
                    >
                      Deselect All
                    </button>
                  </div>
                </div>

                <div style={{
                  background: "#070b14",
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
                        color: col.key === "code" ? "var(--muted2)" : "#e2eaf7",
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
                    e.currentTarget.style.color = "#fff";
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
            Quant Strategy FY 2026–27 · Data sourced from uploaded Excel · Q2–Q4 pending
        </div>
    </div>
    );
    }

export default Dashboard
