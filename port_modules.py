"""
port_modules.py
---------------
Regenerates the two dormant dashboard copies from index.html.

index.html is the ONLY file GitHub Pages serves and the only one meant to be
hand-edited. Its component is written for the browser: React/Recharts arrive as
globals and the PDF/Excel libraries hang off `window`. The other two copies are
ES modules and import those instead:

    indx.html                      React ES-module copy, keeps RAW inline
    quant-dashboard/src/Dashboard.jsx   Vite component, imports RAW from ./data

Hand-editing either of those is how they drift. Run this after every change to
index.html, then commit all three together.

    python port_modules.py
"""

import io
import re
import sys

SRC = "index.html"
OUT_INDX = "indx.html"
OUT_JSX = "quant-dashboard/src/Dashboard.jsx"

IMPORTS = """import { useState, useMemo, useEffect, useRef } from 'react'
import {
  BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer,
  LineChart, Line, Cell, ReferenceLine, Legend, AreaChart, Area,
} from 'recharts'
import ExcelJS from 'exceljs'
import { saveAs } from 'file-saver'
import { jsPDF } from 'jspdf'
import autoTable from 'jspdf-autotable'
import html2canvas from 'html2canvas'
"""

# Browser-global -> ES-module import. Order matters: the two `const` lines are
# dropped outright because the name is already bound by an import above.
DROP_LINES = (
    "const ExcelJS = window.ExcelJS;",
    "const saveAs = window.saveAs || (window.FileSaver && window.FileSaver.saveAs);",
)
SUBS = (
    ("const JsPDF = window.jspdf && window.jspdf.jsPDF;", "const JsPDF = jsPDF;"),
    ("window.html2canvas", "html2canvas"),
    ("doc.autoTable(", "autoTable(doc, "),
)


def extract_component(src):
    """The component body: everything between the babel <script> and the mount."""
    start = src.index('<script type="text/babel"')
    start = src.index("\n", start) + 1
    end = src.index("// Mount the Dashboard")
    body = src[start:end]

    # Drop the two global destructure blocks — the imports replace them.
    body = re.sub(r"// Destructure hooks from React global\r?\n"
                  r"const \{[^}]*\} = React;\r?\n", "", body)
    body = re.sub(r"// Destructure Recharts components from Recharts global\r?\n"
                  r"const \{[^}]*\} = Recharts;\r?\n", "", body)
    return body.strip("\r\n")


def to_module(body):
    for line in DROP_LINES:
        body = re.sub(r"[ \t]*" + re.escape(line) + r"\r?\n", "", body)
    for old, new in SUBS:
        body = body.replace(old, new)
    return body


def split_raw(body):
    """(body_without_RAW, raw_block) — Dashboard.jsx imports RAW from ./data."""
    m = re.search(r"// ── RAW DATA ─+\r?\n(?:export )?const RAW = \[.*?\n\];\r?\n",
                  body, re.DOTALL)
    if not m:
        sys.exit("ERROR: could not locate the RAW block in index.html")
    return body[:m.start()] + body[m.end():], m.group(0)


def check(name, text):
    """Cheap structural sanity checks — a bad transform must not reach git."""
    problems = []
    if text.count("function Dashboard") != 1:
        problems.append("expected exactly 1 `function Dashboard`, got %d"
                        % text.count("function Dashboard"))
    if text.count("export default Dashboard") != 1:
        problems.append("missing/duplicated `export default Dashboard`")
    if text.count("{") != text.count("}"):
        problems.append("unbalanced braces (%d { vs %d })"
                        % (text.count("{"), text.count("}")))
    stray = set(re.findall(r"window\.[A-Za-z0-9_]+", text))
    allowed = {"window.addEventListener", "window.removeEventListener",
               "window.print", "window.innerWidth"}
    if stray - allowed:
        problems.append("unported browser globals: %s" % ", ".join(sorted(stray - allowed)))
    if problems:
        sys.exit("ERROR in %s:\n  - %s" % (name, "\n  - ".join(problems)))
    print("    %-34s OK  (%d lines)" % (name, text.count("\n") + 1))


def write(path, text):
    io.open(path, "w", encoding="utf-8", newline="\r\n").write(text)


def main():
    src = io.open(SRC, encoding="utf-8", newline="").read()
    body = to_module(extract_component(src))

    indx = IMPORTS + "\n" + body + "\n\n\nexport default Dashboard\n"
    stripped, _raw = split_raw(body)
    jsx = (IMPORTS + "import { RAW } from './data'\n\n"
           + stripped.strip("\r\n") + "\n\n\nexport default Dashboard\n")

    indx = indx.replace("\r\n", "\n")
    jsx = jsx.replace("\r\n", "\n")

    print("[1/2] Checking generated modules ...")
    check(OUT_INDX, indx)
    check(OUT_JSX, jsx)

    print("[2/2] Writing ...")
    write(OUT_INDX, indx)
    write(OUT_JSX, jsx)
    print("    Done. index.html -> indx.html + Dashboard.jsx")


if __name__ == "__main__":
    main()
