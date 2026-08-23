import * as XLSX from "xlsx";
import { classifyTitle } from "./calendar.js";
import { toKey } from "./dates.js";

const DATE_KEYS = ["날짜", "일자", "date", "일시", "년월일"];
const TITLE_KEYS = ["내용", "일정", "행사", "학사활동", "제목", "title", "event"];
const TYPE_KEYS = ["구분", "유형", "종류", "type"];

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
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "학사일정");
  return wb;
}

export function exportWorkbook(state, model, htmlTable) {
  const wb = XLSX.utils.book_new();
  const list = [
    ["날짜", "내용", "구분"],
    ...state.events.map((e) => [
      e.date,
      e.title,
      { activity: "학사활동", exam: "고사", closure: "휴업일", holiday: "공휴일" }[e.type] || e.type,
    ]),
  ];
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(list), "학사일정");

  const holidayRows = [["날짜", "공휴일"]];
  for (const [key, name] of model.holidays) holidayRows.push([key, name]);
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(holidayRows), "공휴일");

  if (htmlTable) {
    const ws = XLSX.utils.table_to_sheet(htmlTable);
    XLSX.utils.book_append_sheet(wb, ws, "인쇄표");
  }
  return wb;
}

export function downloadWorkbook(wb, filename) {
  XLSX.writeFile(wb, filename);
}
