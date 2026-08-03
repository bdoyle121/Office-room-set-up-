/*
 * McShane Daily Setup Report -> Excel
 *
 * Runs entirely in the browser: pdf.js reads the PDF's text with its x/y
 * position on the page, this file re-derives the report's table columns
 * from those positions (the same trick used in the original Python/
 * pdfplumber version of this tool), walks the rows as a small state
 * machine per event, and xlsx-js-style writes the result to a downloadable
 * .xlsx file. No server, no upload - the PDF never leaves the browser.
 *
 * If Fordham changes the Daily Setup Report layout, the COL_* ranges below
 * are the first thing to check (open the PDF, note roughly where each
 * column starts in points, update the ranges).
 */

const COL_READY = [0, 100];   // "Setup Ready By" time
const COL_B = [100, 280];     // labels/values: Setup Starts / Pre-Event / Event / Location name / Qty+Resource
const COL_C = [280, 425];     // Event title / Headcount / Layout / Resource instructions
const COL_D = [425, 495];     // Location instructions column
const COL_E = [495, 700];     // Requestor / Scheduler / Organization
// anything >= 700 is the Reference column

const TARGET_BUILDINGS = ["JMCC", "BEPLER", "CAMPBELL"]; // substring matches
const KE_FLOORS_ALLOWED = ["1ST", "3RD"];                 // only these KE floors

function whichCol(x) {
  if (x >= COL_READY[0] && x < COL_READY[1]) return "A";
  if (x >= COL_B[0] && x < COL_B[1]) return "B";
  if (x >= COL_C[0] && x < COL_C[1]) return "C";
  if (x >= COL_D[0] && x < COL_D[1]) return "D";
  if (x >= COL_E[0] && x < COL_E[1]) return "E";
  return "F";
}

function clusterRows(items, tol = 3) {
  // pdf.js's y increases upward; sort top-to-bottom (descending y), then left-to-right.
  const sorted = [...items].sort((a, b) => b.y - a.y || a.x - b.x);
  const rows = [];
  let cur = [];
  let curY = null;
  for (const it of sorted) {
    if (cur.length && Math.abs(it.y - curY) > tol) {
      rows.push(cur);
      cur = [];
    }
    cur.push(it);
    curY = cur.reduce((s, w) => s + w.y, 0) / cur.length;
  }
  if (cur.length) rows.push(cur);
  return rows;
}

function bucketRow(row) {
  const b = { A: [], B: [], C: [], D: [], E: [], F: [] };
  for (const w of [...row].sort((a, c) => a.x - c.x)) {
    b[whichCol(w.x)].push(w.str);
  }
  return b;
}

function isBuildingKept(location) {
  const loc = location.toUpperCase();
  if (loc.startsWith("KE")) {
    return KE_FLOORS_ALLOWED.some((f) => new RegExp(`\\bKE\\s*${f}\\b`).test(loc));
  }
  return TARGET_BUILDINGS.some((bld) => loc.includes(bld));
}

/**
 * Extract all text items (with x/y position) from every page of the PDF,
 * grouped by page in reading order isn't guaranteed by pdf.js, so we sort
 * within each page ourselves.
 */
async function extractPages(pdfDoc) {
  const pages = [];
  for (let pageNum = 1; pageNum <= pdfDoc.numPages; pageNum++) {
    const page = await pdfDoc.getPage(pageNum);
    const content = await page.getTextContent();
    const items = content.items
      .map((it) => ({ str: it.str, x: it.transform[4], y: it.transform[5] }))
      .filter((it) => it.str.trim() !== "");
    pages.push(items);
  }
  return pages;
}

function parseReport(pages) {
  const events = {}; // reference -> event object
  const order = [];
  let reportDate = null;
  const dateRe = /[A-Za-z]+,\s*[A-Za-z]+\s+\d{1,2}\s+\d{4}/;

  for (const items of pages) {
    const rows = clusterRows(items).map(bucketRow);

    if (!reportDate) {
      for (const r of rows) {
        const text = [...r.A, ...r.B, ...r.C].join(" ");
        const m = text.match(dateRe);
        if (m) {
          reportDate = m[0];
          break;
        }
      }
    }

    let currentRef = null;
    let mode = null; // null | "location" | "resource"

    for (const r of rows) {
      // Page header row: "Setup Ready By | Times | Event | Contact | Reference"
      if (r.B[0] === "Times" && r.C[0] === "Event") continue;
      // Footer rows
      if (r.C.some((t) => t.startsWith("Report Printed"))) continue;
      if (r.C.some((t) => t.startsWith("Events:"))) continue;

      // New event row
      if (r.A.length && r.B[0] === "Setup Starts:") {
        const readyBy = r.A[0];
        const setupVal = r.B.length > 1 ? r.B.slice(1).join(" ") : "";
        const title = r.C.join(" ");
        const ref = r.F[0];
        if (!ref) continue;
        currentRef = ref;
        mode = null;
        if (!events[ref]) {
          events[ref] = {
            readyBy,
            setupStarts: setupVal,
            title,
            eventStart: null,
            eventEnd: null,
            locations: [],
            resources: [],
          };
          order.push(ref);
        } else if (!events[ref].title && title) {
          events[ref].title = title;
        }
        continue;
      }

      if (!currentRef) continue;
      const ev = events[currentRef];

      if (r.B[0] === "Pre-Event:") continue;

      if (r.B[0] === "Event:") {
        const times = r.B.slice(1).join(" ");
        const m = times.match(/^(.+?)\s*-\s*(.+)$/);
        if (m) {
          ev.eventStart = m[1].trim();
          ev.eventEnd = m[2].trim();
        }
        mode = null;
        continue;
      }

      if (r.B[0] === "Location" && r.C[0] === "Layout") {
        mode = "location";
        continue;
      }

      if (r.B[0] === "Qty" && r.B[1] === "Resource") {
        mode = "resource";
        continue;
      }

      if (mode === "location") {
        if (r.B.length) ev.locations.push(r.B.join(" "));
        continue;
      }

      if (mode === "resource") {
        if (!r.B.length) continue;
        if (r.B[0].startsWith("No resources")) continue;
        if (/^\d+$/.test(r.B[0])) {
          const qty = r.B[0];
          const name = r.B.slice(1).join(" ");
          const instr = [...r.C, ...r.D, ...r.E].join(" ").trim();
          ev.resources.push([qty, name, instr]);
        } else if (ev.resources.length) {
          // wrapped continuation of the previous resource's name
          ev.resources[ev.resources.length - 1][1] += " " + r.B.join(" ");
        }
        continue;
      }
    }
  }

  return { reportDate, events: order.map((ref) => events[ref]) };
}

function buildRows(reportDate, events) {
  const out = [];
  for (const ev of events) {
    const uniqueLocations = [...new Set(ev.locations)];
    const kept = uniqueLocations.filter(isBuildingKept);
    if (!kept.length) continue;

    let setupTime;
    if (ev.setupStarts && ev.setupStarts !== "no setup time defined") {
      setupTime = ev.setupStarts;
    } else {
      setupTime = `${ev.readyBy} (ready-by; no explicit setup lead time given)`;
    }

    const resources =
      ev.resources
        .map(([q, name, instr]) => `${q}x ${name}${instr ? ` (${instr})` : ""}`)
        .join("; ") || "None";

    out.push({
      Date: reportDate,
      "Building / Room": kept.join("; "),
      Event: ev.title,
      "Setup Initials": "",
      "Setup Time": setupTime,
      "Resources Needed": resources,
      "Take Down Time (Event End)": ev.eventEnd,
      "Take Down Initials": "",
    });
  }
  return out;
}

function rowsToWorkbook(rows) {
  const headers = [
    "Date",
    "Building / Room",
    "Event",
    "Setup Initials",
    "Setup Time",
    "Resources Needed",
    "Take Down Time (Event End)",
    "Take Down Initials",
  ];
  const aoa = [headers, ...rows.map((r) => headers.map((h) => r[h] ?? ""))];
  const ws = XLSX.utils.aoa_to_sheet(aoa);

  const headerStyle = {
    font: { bold: true, color: { rgb: "FFFFFF" } },
    fill: { fgColor: { rgb: "8B1E3F" } },
    alignment: { horizontal: "center", vertical: "center" },
  };
  headers.forEach((_, i) => {
    const cellRef = XLSX.utils.encode_cell({ r: 0, c: i });
    if (ws[cellRef]) ws[cellRef].s = headerStyle;
  });

  ws["!cols"] = headers.map((h, i) => {
    const maxLen = Math.max(h.length, ...rows.map((r) => String(r[h] ?? "").length));
    return { wch: Math.min(maxLen + 2, 70) };
  });
  ws["!freeze"] = { xSplit: 0, ySplit: 1 };
  ws["!autofilter"] = { ref: XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: 0, c: headers.length - 1 } }) };

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Daily Setup");
  return wb;
}

const PDFJS_VERSION = "3.11.174";
const PDFJS_CDN_BASE = `https://cdn.jsdelivr.net/npm/pdfjs-dist@${PDFJS_VERSION}/`;

async function convertPdfToWorkbook(arrayBuffer) {
  const pdf = await pdfjsLib.getDocument({
    data: arrayBuffer,
    // Needed so pdf.js can correctly decode fonts/special characters
    // (curly quotes, etc.) in the report - without these some text can
    // silently drop out of the extracted content.
    cMapUrl: `${PDFJS_CDN_BASE}cmaps/`,
    cMapPacked: true,
    standardFontDataUrl: `${PDFJS_CDN_BASE}standard_fonts/`,
  }).promise;
  const pages = await extractPages(pdf);
  const { reportDate, events } = parseReport(pages);
  const rows = buildRows(reportDate, events);
  if (!rows.length) {
    throw new Error(
      "No matching events found (JMCC / KE 1st / KE 3rd / Bepler / Campbell). Check that this is a McShane Daily Setup Report PDF."
    );
  }
  return { reportDate, rows, workbook: rowsToWorkbook(rows) };
}

// Expose for the UI wiring in index.html and for the Node test harness.
if (typeof module !== "undefined") {
  module.exports = { convertPdfToWorkbook, parseReport, buildRows, extractPages };
}