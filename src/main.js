import { SAMPLE_EVENTS } from "./sample.js";
import { selfCheck as dateCheck } from "./dates.js";
import { selfCheck as holidayCheck } from "./holidays.js";
import { buildCalendar, inferTerms, selfCheck as calendarCheck } from "./calendar.js";
import { FORMATS, renderSheet } from "./render.js";
import {
  downloadWorkbook,
  eventTemplateWorkbook,
  exportWorkbook,
  parseEventWorkbook,
} from "./excel.js";

const STORAGE_KEY = "academic-calendar-state";
const $ = (id) => document.getElementById(id);

function defaultState() {
  const year = 2026;
  return {
    schoolName: "금구중학교",
    year,
    anniversary: "05-01",
    includeLaborDay: true,
    includeSuneung: true,
    ...inferTerms(year, SAMPLE_EVENTS),
    events: SAMPLE_EVENTS.map((e) => ({ ...e })),
    format: "yearly-week",
    semester: 1,
    month: 2,
  };
}

function loadState() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) return defaultState();
    return { ...defaultState(), ...JSON.parse(raw) };
  } catch {
    return defaultState();
  }
}

let state = loadState();
{
  const q = new URLSearchParams(location.search);
  const format = q.get("format");
  if (format && FORMATS.some((f) => f.id === format)) state.format = format;
  const semesterRaw = q.get("semester");
  if (semesterRaw === "1" || semesterRaw === "2") state.semester = Number(semesterRaw);
  const monthRaw = q.get("month");
  if (monthRaw !== null && monthRaw !== "") {
    const month = Number(monthRaw);
    if (Number.isInteger(month) && month >= 0 && month <= 11) state.month = month;
  }
}

function saveState() {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
}

function fillSelects() {
  $("format").innerHTML = FORMATS.map(
    (f) => `<option value="${f.id}">[${f.group}] ${f.label}</option>`,
  ).join("");
  $("month").innerHTML = [2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 0, 1]
    .map((m) => `<option value="${m}">${m + 1}월</option>`)
    .join("");
}

function bindForm() {
  $("schoolName").value = state.schoolName;
  $("year").value = state.year;
  $("anniversary").value = state.anniversary;
  $("includeLaborDay").checked = state.includeLaborDay;
  $("includeSuneung").checked = state.includeSuneung;
  $("sem1Start").value = state.sem1Start;
  $("sem1End").value = state.sem1End;
  $("sem2Start").value = state.sem2Start;
  $("sem2End").value = state.sem2End;
  $("format").value = state.format;
  $("semester").value = String(state.semester);
  $("month").value = String(state.month);
}

function readForm() {
  state.schoolName = $("schoolName").value.trim() || "학교명";
  state.year = Number($("year").value) || 2026;
  state.anniversary = $("anniversary").value.trim();
  state.includeLaborDay = $("includeLaborDay").checked;
  state.includeSuneung = $("includeSuneung").checked;
  state.sem1Start = $("sem1Start").value;
  state.sem1End = $("sem1End").value;
  state.sem2Start = $("sem2Start").value;
  state.sem2End = $("sem2End").value;
  state.format = $("format").value;
  state.semester = Number($("semester").value);
  state.month = Number($("month").value);
}

function renderEvents() {
  const sorted = [...state.events].sort((a, b) => a.date.localeCompare(b.date));
  $("eventList").innerHTML = sorted
    .map(
      (e) => `<div class="ev">
        <span class="d">${e.date}</span>
        <span class="t ${e.type}">${e.title}</span>
        <button type="button" data-del="${e.date}|${e.title}" aria-label="삭제">×</button>
      </div>`,
    )
    .join("");
}

function currentFormat() {
  return FORMATS.find((f) => f.id === state.format) || FORMATS[0];
}

function setPageSize(landscape) {
  let style = document.getElementById("page-size");
  if (!style) {
    style = document.createElement("style");
    style.id = "page-size";
    document.head.appendChild(style);
  }
  style.textContent = `@page { size: A4 ${landscape ? "landscape" : "portrait"}; margin: 0; }`;
}

function render() {
  readForm();
  saveState();
  const model = buildCalendar(state);
  const fmt = currentFormat();
  $("preview").innerHTML = renderSheet(state, model);
  renderEvents();
  document.documentElement.classList.toggle("landscape-format", fmt.landscape);
  const app = document.querySelector(".app");
  app.classList.toggle("needs-semester", fmt.id === "semester-cal" || fmt.id === "semester-week");
  app.classList.toggle("needs-month", fmt.id === "monthly" || fmt.id === "semester-cal");
  $("monthLabel").textContent = fmt.id === "semester-cal" ? "시작 월" : "월";
  setPageSize(fmt.landscape);
}

function shiftKey(key, delta) {
  const [y, m, d] = key.split("-").map(Number);
  const date = new Date(y + delta, m - 1, d);
  const mm = String(date.getMonth() + 1).padStart(2, "0");
  const dd = String(date.getDate()).padStart(2, "0");
  return `${date.getFullYear()}-${mm}-${dd}`;
}

function shiftYear(nextYear) {
  const delta = nextYear - Number(state.year);
  if (!delta) return;
  state.year = nextYear;
  state.sem1Start = shiftKey(state.sem1Start, delta);
  state.sem1End = shiftKey(state.sem1End, delta);
  state.sem2Start = shiftKey(state.sem2Start, delta);
  state.sem2End = shiftKey(state.sem2End, delta);
  state.events = state.events.map((e) => ({ ...e, date: shiftKey(e.date, delta) }));
}

function setup() {
  dateCheck();
  holidayCheck();
  calendarCheck();
  fillSelects();
  bindForm();

  for (const id of [
    "schoolName",
    "anniversary",
    "includeLaborDay",
    "includeSuneung",
    "sem1Start",
    "sem1End",
    "sem2Start",
    "sem2End",
    "format",
    "month",
  ]) {
    $(id).addEventListener("change", render);
    $(id).addEventListener("input", render);
  }

  $("semester").addEventListener("change", () => {
    if ($("format").value === "semester-cal") {
      $("month").value = $("semester").value === "2" ? "8" : "2";
    }
    render();
  });

  $("year").addEventListener("change", () => {
    shiftYear(Number($("year").value));
    bindForm();
    render();
  });

  $("excelFile").addEventListener("change", async (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    const events = parseEventWorkbook(await file.arrayBuffer(), state.year);
    if (!events.length) {
      alert("날짜와 내용 열을 찾지 못했습니다. 입력 양식을 내려받아 형식을 확인해 주세요.");
      return;
    }
    state.events = events;
    Object.assign(state, inferTerms(state.year, events));
    bindForm();
    render();
    e.target.value = "";
  });

  $("templateBtn").addEventListener("click", () => {
    downloadWorkbook(eventTemplateWorkbook(), "학사일정_입력양식.xlsx");
  });
  $("sampleBtn").addEventListener("click", () => {
    state = defaultState();
    bindForm();
    render();
  });
  $("clearBtn").addEventListener("click", () => {
    state.events = [];
    render();
  });
  $("printBtn").addEventListener("click", () => window.print());
  $("xlsxBtn").addEventListener("click", () => {
    const model = buildCalendar(state);
    downloadWorkbook(
      exportWorkbook(state, model, document.querySelector(".paper table")),
      `${state.year}학년도_학사일정.xlsx`,
    );
  });
  $("addBtn").addEventListener("click", () => {
    const date = $("newDate").value;
    const title = $("newTitle").value.trim();
    if (!date || !title) return;
    state.events.push({ date, title, type: $("newType").value });
    $("newTitle").value = "";
    render();
  });
  $("eventList").addEventListener("click", (e) => {
    const btn = e.target.closest("button[data-del]");
    if (!btn) return;
    const [date, title] = btn.dataset.del.split("|");
    state.events = state.events.filter((ev) => !(ev.date === date && ev.title === title));
    render();
  });

  render();
}

setup();
