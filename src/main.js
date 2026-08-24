import { SAMPLE_EVENTS, isSampleEvents } from "./sample.js";
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
import { downloadPdf } from "./pdf.js";
import { downloadHwpx, selfCheck as hwpxCheck } from "./hwpx.js";

const STORAGE_KEY = "academic-calendar-state";
const $ = (id) => document.getElementById(id);

function defaultState() {
  const year = 2026;
  return {
    schoolName: "금구중학교",
    year,
    includeLaborDay: true,
    includeSuneung: true,
    ...inferTerms(year, SAMPLE_EVENTS),
    events: SAMPLE_EVENTS.map((e) => ({ ...e })),
    format: "yearly-week",
    semester: 1,
    month: 2,
    endMonth: 7,
  };
}

function loadState() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) return defaultState();
    const parsed = JSON.parse(raw);
    const loaded = { ...defaultState(), ...parsed };
    if (!Object.hasOwn(parsed, "endMonth")) {
      loaded.endMonth = loaded.semester === 2 || loaded.month === 8 ? 1 : 7;
    }
    if (isSampleEvents(loaded.events)) {
      loaded.events = SAMPLE_EVENTS.map((e) => ({ ...e }));
    }
    return loaded;
  } catch {
    return defaultState();
  }
}

let state = loadState();
let editingIndex = null;
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
  const endMonthRaw = q.get("endMonth");
  if (endMonthRaw !== null && endMonthRaw !== "") {
    const endMonth = Number(endMonthRaw);
    if (Number.isInteger(endMonth) && endMonth >= 0 && endMonth <= 11) state.endMonth = endMonth;
  }
}

function saveState() {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
}

function persistAll() {
  readForm();
  saveState();
}

function fillSelects() {
  $("format").innerHTML = FORMATS.map(
    (f) => `<option value="${f.id}">[${f.group}] ${f.label}</option>`,
  ).join("");
  const monthOpts = [2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 0, 1]
    .map((m) => `<option value="${m}">${m + 1}월</option>`)
    .join("");
  $("month").innerHTML = monthOpts;
  $("endMonth").innerHTML = monthOpts;
}

function bindForm() {
  $("schoolName").value = state.schoolName;
  $("year").value = state.year;
  $("includeLaborDay").checked = state.includeLaborDay;
  $("includeSuneung").checked = state.includeSuneung;
  $("sem1Start").value = state.sem1Start;
  $("sem1End").value = state.sem1End;
  $("sem2Start").value = state.sem2Start;
  $("sem2End").value = state.sem2End;
  $("format").value = state.format;
  $("semester").value = String(state.semester);
  $("month").value = String(state.month);
  $("endMonth").value = String(state.endMonth);
}

function readForm() {
  state.schoolName = $("schoolName").value.trim() || "학교명";
  state.year = Number($("year").value) || 2026;
  state.includeLaborDay = $("includeLaborDay").checked;
  state.includeSuneung = $("includeSuneung").checked;
  state.sem1Start = $("sem1Start").value;
  state.sem1End = $("sem1End").value;
  state.sem2Start = $("sem2Start").value;
  state.sem2End = $("sem2End").value;
  state.format = $("format").value;
  state.semester = Number($("semester").value);
  state.month = Number($("month").value);
  state.endMonth = Number($("endMonth").value);
}

function inAcademicYear(date, year) {
  return date >= `${year}-03-01` && date <= `${year + 1}-02-29`;
}

function stopEditing() {
  editingIndex = null;
  $("addBtn").textContent = "추가";
}

function typeLabel(type) {
  return { activity: "학사활동", exam: "고사", closure: "휴업일", holiday: "공휴일" }[type] || type;
}

function renderEvents() {
  const year = Number(state.year);
  const sorted = state.events
    .map((e, i) => ({ ...e, i }))
    .filter((e) => inAcademicYear(e.date, year))
    .sort((a, b) => a.date.localeCompare(b.date));
  $("eventList").innerHTML = sorted
    .map(
      (e) => `<div class="ev${e.i === editingIndex ? " editing" : ""}">
        <span class="d">${e.date}</span>
        <span class="t ${e.type}">${e.title}</span>
        <span class="k">${typeLabel(e.type)}</span>
        <button type="button" data-edit="${e.i}">수정</button>
        <button type="button" data-del="${e.i}" aria-label="삭제">×</button>
      </div>`,
    )
    .join("");
  $("addBtn").textContent = editingIndex === null ? "추가" : "저장";
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
  style.textContent = landscape
    ? `@page { size: 297mm 210mm; margin: 0; }`
    : `@page { size: 210mm 297mm; margin: 0; }`;
}

function fitSemesterWeek() {
  const paper = document.querySelector(".paper.fmt-semester-week");
  if (!paper) return;
  const table = paper.querySelector(".sem-week");
  const prevOverflow = paper.style.overflow;
  paper.style.overflow = "visible";
  let fs = 12;
  let pad = 2;
  let lh = 1.25;
  const apply = () => {
    paper.style.setProperty("--sw-fs", `${fs}px`);
    paper.style.setProperty("--sw-lh", String(lh));
    paper.style.setProperty("--sw-pad", `${pad}px`);
  };
  apply();
  if (table) {
    table.style.height = "auto";
    table.style.flex = "0 0 auto";
  }
  for (let i = 0; i < 24 && paper.scrollHeight > paper.clientHeight + 1; i++) {
    fs = Math.max(6, Math.round((fs - 0.35) * 100) / 100);
    pad = Math.max(0, Math.round((pad - 0.1) * 100) / 100);
    lh = Math.max(1.05, Math.round((lh - 0.015) * 1000) / 1000);
    apply();
  }
  if (table && paper.scrollHeight <= paper.clientHeight + 1) {
    table.style.height = "";
    table.style.flex = "";
  }
  paper.style.overflow = prevOverflow;
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
  app.classList.toggle("needs-end-month", fmt.id === "semester-cal");
  $("monthLabel").textContent = fmt.id === "semester-cal" ? "시작 월" : "월";
  setPageSize(fmt.landscape);
  fitSemesterWeek();
}

function shiftYear(nextYear) {
  const year = Number(nextYear) || 2026;
  if (year === Number(state.year)) return;
  state.year = year;
  if (isSampleEvents(state.events)) {
    state.events = SAMPLE_EVENTS.map((e) => ({ ...e }));
  }
  Object.assign(state, inferTerms(year, state.events));
}

function setup() {
  dateCheck();
  holidayCheck();
  calendarCheck();
  hwpxCheck();
  fillSelects();
  bindForm();

  for (const id of [
    "schoolName",
    "includeLaborDay",
    "includeSuneung",
    "sem1Start",
    "sem1End",
    "sem2Start",
    "sem2End",
    "format",
    "month",
    "endMonth",
  ]) {
    $(id).addEventListener("change", render);
    $(id).addEventListener("input", render);
  }

  $("semester").addEventListener("change", () => {
    if ($("format").value === "semester-cal") {
      if ($("semester").value === "2") {
        $("month").value = "8";
        $("endMonth").value = "1";
      } else {
        $("month").value = "2";
        $("endMonth").value = "7";
      }
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
    stopEditing();
    Object.assign(state, inferTerms(state.year, events));
    bindForm();
    render();
    e.target.value = "";
  });

  $("templateBtn").addEventListener("click", () => {
    persistAll();
    const events = Array.isArray(state.events) ? state.events.map((e) => ({ ...e })) : [];
    downloadWorkbook(
      eventTemplateWorkbook(events),
      `${state.year}학년도_학사일정_입력양식.xlsx`,
    );
  });
  $("sampleBtn").addEventListener("click", () => {
    state = defaultState();
    stopEditing();
    bindForm();
    render();
  });
  $("clearBtn").addEventListener("click", () => {
    state.events = [];
    stopEditing();
    render();
  });
  $("printBtn").addEventListener("click", () => window.print());
  $("pdfBtn").addEventListener("click", async () => {
    persistAll();
    const paper = document.querySelector(".paper");
    if (!paper) return;
    $("pdfBtn").disabled = true;
    try {
      await downloadPdf(paper, `${state.year}학년도_학사일정.pdf`);
    } finally {
      $("pdfBtn").disabled = false;
    }
  });
  $("xlsxBtn").addEventListener("click", () => {
    persistAll();
    const model = buildCalendar(state);
    downloadWorkbook(
      exportWorkbook(state, model, document.querySelector(".paper")),
      `${state.year}학년도_학사일정.xlsx`,
    );
  });
  $("hwpxBtn").addEventListener("click", () => {
    persistAll();
    const paper = document.querySelector(".paper");
    if (!paper) return;
    downloadHwpx(paper, `${state.year}학년도_학사일정.hwpx`);
  });
  $("addBtn").addEventListener("click", () => {
    const date = $("newDate").value;
    const title = $("newTitle").value.trim();
    if (!date || !title) return;
    const item = { date, title, type: $("newType").value };
    if (editingIndex !== null && state.events[editingIndex]) {
      state.events[editingIndex] = item;
    } else {
      state.events.push(item);
    }
    $("newTitle").value = "";
    stopEditing();
    render();
  });
  $("eventList").addEventListener("click", (e) => {
    const editBtn = e.target.closest("button[data-edit]");
    if (editBtn) {
      const i = Number(editBtn.dataset.edit);
      const ev = state.events[i];
      if (!ev) return;
      editingIndex = i;
      $("newDate").value = ev.date;
      $("newTitle").value = ev.title;
      $("newType").value = ev.type;
      renderEvents();
      return;
    }
    const delBtn = e.target.closest("button[data-del]");
    if (!delBtn) return;
    const i = Number(delBtn.dataset.del);
    if (editingIndex === i) stopEditing();
    else if (editingIndex !== null && editingIndex > i) editingIndex -= 1;
    state.events.splice(i, 1);
    render();
  });

  window.addEventListener("pagehide", persistAll);
  render();
}

setup();
