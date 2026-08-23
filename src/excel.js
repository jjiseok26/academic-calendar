import XLSXStyle from "xlsx-js-style";
import { classifyTitle, monthBlocks } from "./calendar.js";
import { toKey } from "./dates.js";
import { renderSheet } from "./render.js";

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

function styleOf({ header = false, fill, center = false, bold = false, color, size = 9 }) {
  const border = {
    style: "thin",
    color: { rgb: "4A5568" },
  };
  return {
    font: {
      name: "맑은 고딕",
      sz: header ? 10 : size,
      bold: header || bold,
      color: { rgb: color || (header ? "FFFFFF" : "1A202C") },
    },
    fill: fill ? { patternType: "solid", fgColor: { rgb: fill } } : undefined,
    alignment: {
      horizontal: header || center ? "center" : "left",
      vertical: header || center ? "center" : "top",
      wrapText: true,
    },
    border: { top: border, bottom: border, left: border, right: border },
  };
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

function styleHolidaySheet(ws) {
  paintSheet(ws, (cell, r, c) => {
    if (r === 0) return { header: true, fill: HEADER_FILL, center: true };
    return { fill: HOLIDAY_FILL, center: c === 0 };
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

function cellText(el) {
  const ev = el.querySelector(":scope > .ev");
  if (ev) {
    const day = el.querySelector(":scope > .n")?.textContent.trim() || "";
    const parts = [...ev.querySelectorAll(".act, .hol")].map((n) => n.textContent.trim()).filter(Boolean);
    if (parts.length) return [day, ...parts].filter(Boolean).join("\n");
  }
  return (el.innerText || "")
    .replace(/\r\n?/g, "\n")
    .replace(/[ \t]+\n/g, "\n")
    .replace(/\n[ \t]+/g, "\n")
    .replace(/\n{3,}/g, "\n\n")
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
        center: header || clsHas(el, ["num", "month", "week", "n"]),
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

function fitSheet(ws, landscape) {
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
      const longest = v.split("\n").reduce((m, line) => Math.max(m, [...line].length), 0);
      w = Math.max(w, Math.min(longest + 1, 22));
    }
    cols[c] = { wch: w };
  }
  for (let r = 0; r <= maxR; r++) {
    let lines = 1;
    for (let c = 0; c <= maxC; c++) {
      const v = String(ws[XLSX.utils.encode_cell({ r, c })]?.v || "");
      lines = Math.max(lines, v.split("\n").length);
    }
    rows[r] = { hpt: Math.min(12 + 11 * lines, 78) };
  }
  ws["!cols"] = cols;
  ws["!rows"] = rows;
  if (ws["!merges"]?.length) {
    const title = ws.A1?.v;
    if (title && maxC > 0) {
      const covered = ws["!merges"].some((m) => m.s.r === 0 && m.s.c === 0);
      if (!covered) ws["!merges"].unshift({ s: { r: 0, c: 0 }, e: { r: 0, c: maxC } });
    }
  } else if (maxC > 0 && ws.A1) {
    ws["!merges"] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: maxC } }];
  }
  ws["!pageSetup"] = {
    paperSize: 9,
    orientation: landscape ? "landscape" : "portrait",
    fitToWidth: 1,
    fitToHeight: 1,
    fitToPage: true,
  };
  ws["!printOptions"] = { horizontalCentered: true };
  ws["!margins"] = { left: 0.25, right: 0.25, top: 0.4, bottom: 0.4, header: 0.15, footer: 0.15 };
}

function paperToSheet(html, landscape) {
  const box = document.createElement("div");
  box.innerHTML = html;
  const paper = box.querySelector(".paper") || box;
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
        putCell(ws, r0 + used, c0, text);
        ws["!merges"].push({
          s: { r: r0 + used, c: c0 },
          e: { r: r0 + used, c: c0 + 6 },
        });
        used += Math.max(2, text.split("\n").length);
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

  fitSheet(ws, landscape);
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

export function eventTemplateWorkbook() {
  const rows = [
    ["날짜", "내용", "구분"],
    ["2026-03-03", "입학식/개학식", "학사활동"],
    ["2026-04-10", "스포츠 데이", "학사활동"],
    ["2026-04-27", "1차 지필평가", "고사"],
    ["2026-05-04", "재량휴업일", "휴업일"],
    ["2026-07-21", "여름방학식", "학사활동"],
    ["2026-08-11", "개학일", "학사활동"],
    ["2026-12-31", "종업식, 졸업식", "학사활동"],
  ];
  const ws = XLSX.utils.aoa_to_sheet(rows);
  ws["!cols"] = [{ wch: 14 }, { wch: 28 }, { wch: 12 }];
  styleListSheet(ws);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "학사일정");
  return wb;
}

export function exportWorkbook(state, model) {
  const wb = XLSX.utils.book_new();
  const list = [
    ["날짜", "내용", "구분"],
    ...state.events.map((e) => [
      e.date,
      e.title,
      { activity: "학사활동", exam: "고사", closure: "휴업일", holiday: "공휴일" }[e.type] || e.type,
    ]),
  ];
  const listSheet = XLSX.utils.aoa_to_sheet(list);
  listSheet["!cols"] = [{ wch: 14 }, { wch: 28 }, { wch: 12 }];
  styleListSheet(listSheet);
  XLSX.utils.book_append_sheet(wb, listSheet, "학사일정");

  const holidayRows = [["날짜", "공휴일"]];
  for (const [key, name] of model.holidays) holidayRows.push([key, name]);
  const holidaySheet = XLSX.utils.aoa_to_sheet(holidayRows);
  holidaySheet["!cols"] = [{ wch: 14 }, { wch: 22 }];
  styleHolidaySheet(holidaySheet);
  XLSX.utils.book_append_sheet(wb, holidaySheet, "공휴일");

  const jobs = [
    { name: "연간주간", format: "yearly-week", landscape: false },
    { name: "연간달력", format: "yearly-grid", landscape: false },
    { name: "1학기달력", format: "semester-cal", landscape: true, semester: 1, month: 2, endMonth: 7 },
    { name: "2학기달력", format: "semester-cal", landscape: true, semester: 2, month: 8, endMonth: 1 },
    { name: "1학기주간", format: "semester-week", landscape: false, semester: 1 },
    { name: "2학기주간", format: "semester-week", landscape: false, semester: 2 },
    ...monthBlocks(model).map((m) => ({
      name: `월중_${m.label}`,
      format: "monthly",
      landscape: true,
      month: m.monthIndex,
    })),
  ];
  for (const job of jobs) {
    const html = renderSheet({ ...state, ...job }, model);
    XLSX.utils.book_append_sheet(wb, paperToSheet(html, job.landscape), sheetName(job.name));
  }
  return wb;
}

export function downloadWorkbook(wb, filename) {
  XLSX.writeFile(wb, filename);
}
