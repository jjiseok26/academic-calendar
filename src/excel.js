import XLSXStyle from "xlsx-js-style";
import { classifyTitle } from "./calendar.js";
import { toKey } from "./dates.js";
import { FORMATS, renderSheet } from "./render.js";

const XLSX = XLSXStyle.utils ? XLSXStyle : XLSXStyle.default;

const DATE_KEYS = ["날짜", "일자", "date", "일시", "년월일"];
const TITLE_KEYS = ["내용", "일정", "행사", "학사활동", "제목", "title", "event"];
const TYPE_KEYS = ["구분", "유형", "종류", "type"];

const EXAM_FILL = "FEFCBF";
const ACTIVITY_FILL = "FEFCBF";
const CLOSURE_FILL = "FED7D7";
const HOLIDAY_FILL = "FFF5F5";
const SAT_FILL = "EBF8FF";
const SUM_FILL = "FAF089";
const MONTH_SUM_FILL = "C6F6D5";
const MUTED_FILL = "F7FAFC";
const HEADER_FILL = "1B365D";

function pick(row, keys) {
  const entries = Object.entries(row);
  for (const key of keys) {
    const found = entries.find(([k]) => k.replace(/\s/g, "").toLowerCase() === key.toLowerCase());
    if (found) return found[1];
  }
  return undefined;
}

function excelSerialToDate(n) {
  const utc = Date.UTC(1899, 11, 30) + Math.round(n * 86400000);
  const d = new Date(utc);
  return new Date(d.getFullYear(), d.getMonth(), d.getDate());
}

function parseDate(value, year) {
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return new Date(value.getFullYear(), value.getMonth(), value.getDate());
  }
  if (typeof value === "number" && value > 20000 && value < 80000) {
    return excelSerialToDate(value);
  }
  const text = String(value).trim();
  if (!text) return null;
  const iso = text.match(/(\d{4})[./-](\d{1,2})[./-](\d{1,2})/);
  if (iso) return new Date(Number(iso[1]), Number(iso[2]) - 1, Number(iso[3]));
  const md = text.match(/^(\d{1,2})[./-](\d{1,2})$/);
  if (md) {
    const month = Number(md[1]);
    const day = Number(md[2]);
    const y = month >= 3 ? year : year + 1;
    return new Date(y, month - 1, day);
  }
  const korean = text.match(/(\d{1,2})\s*월\s*(\d{1,2})\s*일/);
  if (korean) {
    const month = Number(korean[1]);
    const day = Number(korean[2]);
    const y = month >= 3 ? year : year + 1;
    return new Date(y, month - 1, day);
  }
  return null;
}

function typeFrom(value, title) {
  const t = String(value || "").trim();
  if (/휴업|방학/.test(t)) return "closure";
  if (/고사|시험|지필/.test(t)) return "exam";
  if (/공휴|휴일/.test(t)) return "holiday";
  if (/학사|행사|활동/.test(t)) return "activity";
  return classifyTitle(title);
}

function edge() {
  return { style: "thin", color: { rgb: "4A5568" } };
}

function boxBorder() {
  return { top: edge(), bottom: edge(), left: edge(), right: edge() };
}

function styleOf({ header = false, fill, center = false, bold = false, color, size = 9 }) {
  return {
    font: {
      name: "맑은 고딕",
      sz: header ? 10 : size,
      bold: header || bold,
      color: { rgb: color || (header ? "FFFFFF" : "1A202C") },
    },
    fill: { patternType: "solid", fgColor: { rgb: fill || "FFFFFF" } },
    alignment: {
      horizontal: header || center ? "center" : "left",
      vertical: header || center ? "center" : "top",
      wrapText: true,
    },
    border: boxBorder(),
  };
}

function paintMerges(ws) {
  for (const m of ws["!merges"] || []) {
    const origin = ws[XLSX.utils.encode_cell({ r: m.s.r, c: m.s.c })];
    const base = origin?.s || styleOf({});
    for (let r = m.s.r; r <= m.e.r; r++) {
      for (let c = m.s.c; c <= m.e.c; c++) {
        const addr = XLSX.utils.encode_cell({ r, c });
        if (!ws[addr]) ws[addr] = { t: "s", v: "" };
        ws[addr].s = {
          font: ws[addr].s?.font || base.font,
          fill: ws[addr].s?.fill || base.fill,
          alignment: ws[addr].s?.alignment || base.alignment,
          border: boxBorder(),
        };
      }
    }
  }
}

function paintSheet(ws, fillForCell) {
  const ref = ws["!ref"];
  if (!ref) return;
  const range = XLSX.utils.decode_range(ref);
  for (let r = range.s.r; r <= range.e.r; r++) {
    for (let c = range.s.c; c <= range.e.c; c++) {
      const addr = XLSX.utils.encode_cell({ r, c });
      if (!ws[addr]) ws[addr] = { t: "s", v: "" };
      ws[addr].s = styleOf(fillForCell(ws[addr], r, c));
    }
  }
}

function fillForType(type) {
  if (type === "고사" || type === "학사활동") return ACTIVITY_FILL;
  if (type === "휴업일") return CLOSURE_FILL;
  if (type === "공휴일") return HOLIDAY_FILL;
  return undefined;
}

function styleListSheet(ws) {
  paintSheet(ws, (cell, r, c) => {
    if (r === 0) return { header: true, fill: HEADER_FILL, center: true };
    const typeCell = ws[XLSX.utils.encode_cell({ r, c: 2 })];
    return {
      fill: fillForType(String(typeCell?.v || "")),
      center: c !== 1,
    };
  });
}

function htmlCellFill(el) {
  const cls = el.classList;
  if (cls.contains("exam") || el.querySelector(".exam, .act.exam")) return EXAM_FILL;
  if (cls.contains("events") || el.querySelector(".act")) return ACTIVITY_FILL;
  if (cls.contains("off")) return CLOSURE_FILL;
  if (cls.contains("holidays") || cls.contains("sun") || el.querySelector(".hol")) return HOLIDAY_FILL;
  if (cls.contains("sat")) return SAT_FILL;
  if (cls.contains("month-sum")) return MONTH_SUM_FILL;
  if (cls.contains("vac") || cls.contains("muted")) return MUTED_FILL;
  if (el.closest("tr")?.classList.contains("sum")) return SUM_FILL;
  if (el.tagName === "TH") return HEADER_FILL;
  return undefined;
}

function clsHas(el, names) {
  return names.some((n) => el.classList.contains(n));
}

function cleanLine(text) {
  return String(text || "").replace(/\s+/g, " ").trim();
}

function walkText(node) {
  const BLOCK = new Set(["DIV", "P", "LI", "TR", "H1", "H2", "H3", "UL", "OL", "ARTICLE", "HEADER", "FOOTER", "SECTION"]);
  if (node.nodeType === 3) return node.textContent.replace(/[\t\r\n]+/g, " ");
  if (node.nodeName === "BR") return "\n";
  if (node.nodeName === "STYLE" || node.nodeName === "SCRIPT") return "";
  let out = "";
  for (const child of node.childNodes) out += walkText(child);
  if (BLOCK.has(node.nodeName)) out = `${out.replace(/[ \t]+$/g, "")}\n`;
  return out.replace(/[ \t]{2,}/g, " ").replace(/\n{3,}/g, "\n\n");
}

function cellText(el) {
  const day = el.querySelector(":scope > .n, :scope > .md")?.textContent.trim() || "";
  const host = el.querySelector(":scope > .ev, :scope > .cell-ev") || (el.matches("td, th") ? el : null);
  if (host) {
    const bits = [...host.querySelectorAll(".act, .hol")].map((n) => cleanLine(n.textContent)).filter(Boolean);
    if (bits.length || day) return [day, ...bits].filter(Boolean).join("\n");
  }
  if (el.matches("ul, ol, .month-events") || el.querySelector(":scope > li")) {
    return [...el.querySelectorAll("li")].map((n) => cleanLine(n.textContent)).filter(Boolean).join("\n");
  }
  return walkText(el)
    .replace(/[ \t]+\n/g, "\n")
    .replace(/\n[ \t]+/g, "\n")
    .trim();
}

function putCell(ws, r, c, value, opts = {}) {
  ws[XLSX.utils.encode_cell({ r, c })] = {
    t: "s",
    v: String(value ?? "").replace(/\n/g, "\r\n"),
    s: styleOf({ size: 8, ...opts }),
  };
}

function placeTable(ws, table, r0, c0) {
  if (!ws["!merges"]) ws["!merges"] = [];
  const taken = new Set();
  const key = (r, c) => `${r},${c}`;
  let maxR = r0 - 1;
  let maxC = c0 - 1;
  [...table.rows].forEach((tr, ri) => {
    let c = c0;
    [...tr.cells].forEach((el) => {
      while (taken.has(key(r0 + ri, c))) c += 1;
      const rs = el.rowSpan || 1;
      const cs = el.colSpan || 1;
      const header = el.tagName === "TH";
      putCell(ws, r0 + ri, c, cellText(el), {
        header,
        fill: htmlCellFill(el),
        center: header || clsHas(el, ["num", "month", "week", "n", "d"]) || !!el.closest(".mini-cal"),
        color: el.classList.contains("sun") ? "C53030" : el.classList.contains("sat") ? "2B6CB0" : undefined,
        size: header ? 9 : 8,
      });
      if (rs > 1 || cs > 1) {
        ws["!merges"].push({
          s: { r: r0 + ri, c },
          e: { r: r0 + ri + rs - 1, c: c + cs - 1 },
        });
      }
      for (let i = 0; i < rs; i++) {
        for (let j = 0; j < cs; j++) taken.add(key(r0 + ri + i, c + j));
      }
      maxR = Math.max(maxR, r0 + ri + rs - 1);
      maxC = Math.max(maxC, c + cs - 1);
      c += cs;
    });
  });
  return { rows: Math.max(0, maxR - r0 + 1), cols: Math.max(0, maxC - c0 + 1) };
}

function columnWidths(table, landscape) {
  const row = table.tHead?.rows[0] || table.rows[0];
  if (!row) return null;
  const widths = [];
  [...row.cells].forEach((cell) => {
    const span = cell.colSpan || 1;
    const w = (cell.getBoundingClientRect().width || cell.offsetWidth || 1) / span;
    for (let i = 0; i < span; i++) widths.push(w);
  });
  const sum = widths.reduce((a, b) => a + b, 0);
  if (!sum) return null;
  const pageWch = ((landscape ? 281 : 194) * 96) / 25.4 / 7;
  return widths.map((w) => ({ wch: Math.max(2.5, (w / sum) * pageWch) }));
}

function fitSheet(ws, landscape, paper) {
  let maxR = 0;
  let maxC = 0;
  for (const addr of Object.keys(ws)) {
    if (addr.startsWith("!")) continue;
    const { r, c } = XLSX.utils.decode_cell(addr);
    maxR = Math.max(maxR, r);
    maxC = Math.max(maxC, c);
  }
  ws["!ref"] = XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: maxR, c: maxC } });
  const cols = [];
  const rows = [];
  for (let c = 0; c <= maxC; c++) {
    let w = 5;
    for (let r = 0; r <= maxR; r++) {
      const v = String(ws[XLSX.utils.encode_cell({ r, c })]?.v || "");
      const longest = v.split(/\r?\n/).reduce((m, line) => Math.max(m, [...line].length), 0);
      w = Math.max(w, Math.min(longest + 1, 22));
    }
    cols[c] = { wch: w };
  }
  const table = paper?.querySelector(".year-week, .sem-week, .month-big");
  const live = table ? columnWidths(table, landscape) : null;
  if (live?.length) {
    for (let c = 0; c < live.length && c <= maxC; c++) cols[c] = live[c];
  }
  const vMerge = new Set();
  for (const m of ws["!merges"] || []) {
    if (m.e.r === m.s.r) continue;
    for (let r = m.s.r; r <= m.e.r; r++) vMerge.add(r);
  }
  for (let r = 0; r <= maxR; r++) {
    let lines = 1;
    if (!vMerge.has(r)) {
      for (let c = 0; c <= maxC; c++) {
        const v = String(ws[XLSX.utils.encode_cell({ r, c })]?.v || "");
        lines = Math.max(lines, v.split(/\r?\n/).length);
      }
    }
    rows[r] = { hpt: Math.min(16 + 13 * lines, 36) };
  }
  for (const m of ws["!merges"] || []) {
    const v = String(ws[XLSX.utils.encode_cell({ r: m.s.r, c: m.s.c })]?.v || "");
    const lines = Math.max(1, v.split(/\r?\n/).length);
    const span = m.e.r - m.s.r + 1;
    if (span > 1) {
      const per = Math.max(16, Math.ceil((15 * lines) / span));
      for (let r = m.s.r; r <= m.e.r; r++) {
        rows[r] = { hpt: Math.max(rows[r]?.hpt || 0, per) };
      }
    } else {
      rows[m.s.r] = { hpt: Math.max(rows[m.s.r]?.hpt || 0, Math.min(16 + 13 * lines, 220)) };
    }
  }
  ws["!cols"] = cols;
  ws["!rows"] = Array.from({ length: maxR + 1 }, (_, r) => {
    const hpt = rows[r]?.hpt || 18;
    return { hpt, hpx: Math.round((hpt * 96) / 72) };
  });
  if (ws["!merges"]?.length) {
    const title = ws.A1?.v;
    if (title && maxC > 0) {
      const covered = ws["!merges"].some((m) => m.s.r === 0 && m.s.c === 0);
      if (!covered) ws["!merges"].unshift({ s: { r: 0, c: 0 }, e: { r: 0, c: maxC } });
    }
  } else if (maxC > 0 && ws.A1) {
    ws["!merges"] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: maxC } }];
  }
  paintMerges(ws);
  ws["!pageSetup"] = {
    paperSize: 9,
    orientation: landscape ? "landscape" : "portrait",
    fitToWidth: 1,
    fitToHeight: 1,
    fitToPage: true,
  };
  ws["!printOptions"] = { horizontalCentered: false, verticalCentered: false };
  ws["!margins"] = { left: 0.2, right: 0.2, top: 0.25, bottom: 0.25, header: 0, footer: 0 };
}

function paperEl(source) {
  if (source && source.querySelector) return source;
  const box = document.createElement("div");
  box.innerHTML = typeof source === "string" ? source : "";
  return box.querySelector(".paper") || box;
}

function paperToSheet(source, landscape) {
  const paper = paperEl(source);
  const ws = { "!merges": [] };
  let row = 0;
  const title = paper.querySelector("h1, .ym-box")?.innerText.trim() || "";
  const school = paper.querySelector(".school, .sheet-foot")?.innerText.trim() || "";
  if (title) {
    putCell(ws, 0, 0, school ? `${title}  ${school}` : title, {
      header: true,
      fill: HEADER_FILL,
      center: true,
      size: 12,
    });
    row = 2;
  }

  const cards = [...paper.querySelectorAll(".month-card, .sem-col")];
  if (cards.length) {
    const cols = paper.querySelector(".sem-cals") ? cards.length : 4;
    const bandH = [];
    cards.forEach((card, i) => {
      const col = i % cols;
      const band = Math.floor(i / cols);
      const c0 = col * 8;
      const r0 = row + bandH.slice(0, band).reduce((a, b) => a + b, 0);
      const heading = card.querySelector("h3")?.innerText.trim() || "";
      const top = card.querySelector(".month-top")?.innerText.trim() || "";
      putCell(ws, r0, c0, [heading, top].filter(Boolean).join(" · "), {
        header: true,
        fill: HEADER_FILL,
        center: true,
      });
      ws["!merges"].push({ s: { r: r0, c: c0 }, e: { r: r0, c: c0 + 6 } });
      let used = 1;
      const table = card.querySelector("table");
      if (table) {
        const placed = placeTable(ws, table, r0 + 1, c0);
        used += placed.rows;
      }
      const extra = card.querySelector(".month-events, .daylist");
      if (extra) {
        const text = cellText(extra);
        const lines = Math.max(2, text.split("\n").length);
        putCell(ws, r0 + used, c0, text);
        ws["!merges"].push({
          s: { r: r0 + used, c: c0 },
          e: { r: r0 + used + lines - 1, c: c0 + 6 },
        });
        used += lines;
      }
      bandH[band] = Math.max(bandH[band] || 0, used + 1);
    });
  } else {
    for (const table of paper.querySelectorAll("table")) {
      const placed = placeTable(ws, table, row, 0);
      row += placed.rows + 1;
    }
    const note = paper.querySelector(".note")?.innerText.trim();
    if (note) putCell(ws, row, 0, note, { size: 8 });
  }

  fitSheet(ws, landscape, paper);
  return ws;
}

function sheetName(name) {
  return name.replace(/[\\/?*[\]]/g, "").slice(0, 31);
}

export function parseEventWorkbook(buffer, year) {
  const wb = XLSX.read(buffer, { type: "array", cellDates: true });
  const sheet = wb.Sheets[wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(sheet, { defval: "", raw: true });
  const events = [];
  for (const row of rows) {
    const dateVal = pick(row, DATE_KEYS) ?? Object.values(row)[0];
    const titleVal = pick(row, TITLE_KEYS) ?? Object.values(row)[1];
    const typeVal = pick(row, TYPE_KEYS);
    const title = String(titleVal || "").trim();
    const date = parseDate(dateVal, year);
    if (!date || !title) continue;
    events.push({ date: toKey(date), title, type: typeFrom(typeVal, title) });
  }
  return events;
}

const TYPE_LABEL = {
  activity: "학사활동",
  exam: "고사",
  closure: "휴업일",
  holiday: "공휴일",
};

export function eventTemplateWorkbook(events = []) {
  const body = (Array.isArray(events) ? events : [])
    .filter((e) => e && e.date && String(e.title || "").trim())
    .sort((a, b) => String(a.date).localeCompare(String(b.date)))
    .map((e) => [String(e.date), String(e.title).trim(), TYPE_LABEL[e.type] || e.type || ""]);
  const rows = [["날짜", "내용", "구분"], ...body];
  const ws = XLSX.utils.aoa_to_sheet(rows, { cellDates: false });
  ws["!cols"] = [{ wch: 14 }, { wch: 28 }, { wch: 12 }];
  styleListSheet(ws);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "학사일정");
  return wb;
}

export function exportWorkbook(state, model, paper) {
  const wb = XLSX.utils.book_new();
  const fmt = FORMATS.find((f) => f.id === state.format) || FORMATS[0];
  const source = paper || renderSheet(state, model);
  const landscape = paper?.classList?.contains("landscape") ?? fmt.landscape;
  XLSX.utils.book_append_sheet(wb, paperToSheet(source, landscape), sheetName(fmt.label));
  return wb;
}

export function downloadWorkbook(wb, filename) {
  const out = XLSX.write(wb, { bookType: "xlsx", type: "array", cellStyles: true });
  const blob = new Blob([out], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 1500);
}
