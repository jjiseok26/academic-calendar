import { WEEKDAYS, formatMD, formatMDW, toKey, weekdayLabel } from "./dates.js";
import { countInstructional, monthBlocks, monthGrid, weekRows } from "./calendar.js";

function esc(s) {
  return String(s ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;");
}

function infoOf(model, date) {
  return model.byKey.get(toKey(date));
}

function schoolEvents(day) {
  return day?.events?.filter((e) => e.type !== "holiday") || [];
}

function dayClass(model, date, monthIndex) {
  const info = infoOf(model, date);
  const holiday = info?.holiday;
  const cls = ["d"];
  if (date.getMonth() !== monthIndex) cls.push("muted");
  if (date.getDay() === 0 || holiday) cls.push("sun");
  else if (date.getDay() === 6) cls.push("sat");
  if (info?.closure) cls.push("off");
  if (info?.events?.some((e) => e.type === "exam")) cls.push("exam");
  if (info && !info.inTerm) cls.push("vac");
  return cls.join(" ");
}

function monthEventLists(model, year, monthIndex) {
  const acts = [];
  const hols = [];
  for (const day of model.days) {
    if (day.date.getFullYear() !== year || day.date.getMonth() !== monthIndex) continue;
    if (day.holiday) hols.push({ date: day.date, text: day.holiday });
    for (const ev of schoolEvents(day)) acts.push({ date: day.date, text: ev.title, type: ev.type });
  }
  return { acts, hols };
}

function instructionalMonth(model, year, monthIndex) {
  return countInstructional(
    model,
    (d) => d.date.getFullYear() === year && d.date.getMonth() === monthIndex,
  );
}

function paper(format, landscape, inner) {
  return `<section class="paper ${landscape ? "landscape" : "portrait"} fmt-${format}">${inner}<p class="copy">© jiseok</p></section>`;
}

function miniMonth(model, year, monthIndex) {
  const { cells } = monthGrid(year, monthIndex);
  const heads = WEEKDAYS.map(
    (w, i) => `<th class="${i === 0 ? "sun" : i === 6 ? "sat" : ""}">${w}</th>`,
  ).join("");
  const weeks = [];
  for (let r = 0; r < 6; r++) {
    const slice = cells.slice(r * 7, r * 7 + 7);
    if (r === 5 && slice.every((d) => d.getMonth() !== monthIndex)) continue;
    weeks.push(
      `<tr>${slice
        .map((d) => `<td class="${dayClass(model, d, monthIndex)}">${d.getDate()}</td>`)
        .join("")}</tr>`,
    );
  }
  return `<table class="mini-cal"><thead><tr>${heads}</tr></thead><tbody>${weeks.join("")}</tbody></table>`;
}

function rowSemester(model, row) {
  const hits = row.days
    .map((d) => infoOf(model, d))
    .filter((d) => d && d.date.getMonth() === row.monthIndex && d.date.getFullYear() === row.year);
  if (hits.some((d) => d.semester === 1)) return 1;
  if (hits.some((d) => d.semester === 2)) return 2;
  if (row.monthIndex >= 2 && row.monthIndex <= 6) return 1;
  return 2;
}

function weekdayCounts(model, pred) {
  return [0, 1, 2, 3, 4, 5, 6].map((wd) =>
    countInstructional(model, (d) => pred(d) && d.weekday === wd),
  );
}

function summaryRow(model, semester, label) {
  const pred = (d) => (semester === 0 ? true : d.semester === semester);
  const counts = weekdayCounts(model, pred);
  const total = countInstructional(model, pred);
  return `<tr class="sum">
    <td colspan="2">${label}</td>
    ${counts.map((n) => `<td>${n || ""}</td>`).join("")}
    <td>${total}</td><td>${total}</td><td>${total}</td>
    <td colspan="2"></td>
  </tr>`;
}

export function renderYearlyWeek(state, model) {
  const rows = weekRows(model);
  const monthMeta = new Map();
  for (const row of rows) {
    const id = `${row.year}-${row.monthIndex}`;
    if (!monthMeta.has(id)) monthMeta.set(id, { count: 0 });
    monthMeta.get(id).count += 1;
  }

  let cumulative = 0;
  const seen = new Set();
  const body = [];

  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];
    const id = `${row.year}-${row.monthIndex}`;
    const first = !seen.has(id);
    if (first) seen.add(id);
    const meta = monthMeta.get(id);
    const monthCount = instructionalMonth(model, row.year, row.monthIndex);
    if (first) cumulative += monthCount;

    const weekDays = row.days
      .map((d) => `<td class="${dayClass(model, d, row.monthIndex)}">${d.getDate()}</td>`)
      .join("");
    const weekCount = row.days.filter((d) => {
      const info = infoOf(model, d);
      return info?.instructional && d.getMonth() === row.monthIndex && d.getFullYear() === row.year;
    }).length;

    let actCell = "";
    let holCell = "";
    if (first) {
      const { acts, hols } = monthEventLists(model, row.year, row.monthIndex);
      actCell = acts
        .map((e) => `<div class="act ${e.type}">${esc(formatMDW(e.date))} ${esc(e.text)}</div>`)
        .join("");
      holCell = hols
        .map((e) => `<div class="hol">${esc(formatMDW(e.date))} ${esc(e.text)}</div>`)
        .join("");
    }

    body.push(`<tr>
      ${first ? `<td class="month" rowspan="${meta.count}">${row.monthIndex + 1}</td>` : ""}
      <td class="week">${row.weekNo}</td>
      ${weekDays}
      <td class="num">${weekCount}</td>
      ${first ? `<td class="num month-sum" rowspan="${meta.count}">${monthCount}</td>` : ""}
      ${first ? `<td class="num month-sum" rowspan="${meta.count}">${cumulative}</td>` : ""}
      ${first ? `<td class="events" rowspan="${meta.count}">${actCell}</td>` : ""}
      ${first ? `<td class="holidays" rowspan="${meta.count}">${holCell}</td>` : ""}
    </tr>`);

    const next = rows[i + 1];
    if (rowSemester(model, row) === 1 && (!next || rowSemester(model, next) !== 1)) {
      body.push(summaryRow(model, 1, "1학기 계"));
    }
  }
  body.push(summaryRow(model, 2, "2학기 계"));
  body.push(summaryRow(model, 0, "합계"));

  const inner = `
    <header class="sheet-head">
      <h1>${state.year}학년도 학사일정</h1>
      <div class="school">${esc(state.schoolName)}</div>
    </header>
    <table class="year-week">
      <thead>
        <tr>
          <th>월</th><th>주</th>
          <th class="sun">일</th><th>월</th><th>화</th><th>수</th><th>목</th><th>금</th><th class="sat">토</th>
          <th>수업일수(주)</th><th>수업일수(월)</th><th>수업일수 누계</th>
          <th>학사활동</th><th>공휴일</th>
        </tr>
      </thead>
      <tbody>${body.join("")}</tbody>
    </table>
    <p class="note">※ ${state.year}학년도 학사일정은 교육부 등 상위기관 또는 학교 사정에 의해 변경될 수 있음.</p>
  `;
  return paper("yearly-week", false, inner);
}

export function renderYearlyGrid(state, model) {
  const blocks = monthBlocks(model)
    .map((m) => {
      const { acts, hols } = monthEventLists(model, m.year, m.monthIndex);
      const days = instructionalMonth(model, m.year, m.monthIndex);
      const list = [
        ...hols.map((e) => `<li class="hol">${esc(formatMDW(e.date))} ${esc(e.text)}</li>`),
        ...acts.map((e) => `<li class="act ${e.type}">${esc(formatMDW(e.date))} ${esc(e.text)}</li>`),
      ].join("");
      return `<article class="month-card">
        <div class="month-top">수업일수: ${days}일</div>
        <h3>${m.label}</h3>
        ${miniMonth(model, m.year, m.monthIndex)}
        <ul class="month-events">${list}</ul>
      </article>`;
    })
    .join("");
  const inner = `
    <header class="sheet-head center">
      <h1>${state.year}학년도 학사달력</h1>
    </header>
    <div class="month-grid-12">${blocks}</div>
    <footer class="sheet-foot">${esc(state.schoolName)}</footer>
  `;
  return paper("yearly-grid", false, inner);
}

function monthRange(all, startMonth, endMonth) {
  const startIdx = all.findIndex((m) => m.monthIndex === Number(startMonth));
  const endIdx = all.findIndex((m) => m.monthIndex === Number(endMonth));
  const s = startIdx < 0 ? 0 : startIdx;
  const e = endIdx < 0 ? (s + 5) % 12 : endIdx;
  const months = [];
  let i = s;
  while (months.length < 12) {
    months.push(all[i]);
    if (i === e) break;
    i = (i + 1) % 12;
  }
  return months;
}

export function renderSemesterCal(state, model) {
  const all = monthBlocks(model);
  const months = monthRange(all, state.month, state.endMonth);
  const cols = months
    .map((m) => {
      const last = new Date(m.year, m.monthIndex + 1, 0).getDate();
      const lines = [];
      for (let d = 1; d <= last; d++) {
        const date = new Date(m.year, m.monthIndex, d);
        const info = infoOf(model, date);
        const texts = [];
        if (info?.holiday) texts.push(`<span class="hol">${esc(info.holiday)}</span>`);
        for (const ev of schoolEvents(info)) {
          texts.push(`<span class="act ${ev.type}">${esc(ev.title)}</span>`);
        }
        const wcls = date.getDay() === 0 || info?.holiday ? "sun" : date.getDay() === 6 ? "sat" : "";
        lines.push(
          `<div class="dayline ${wcls}"><span class="n">${d}</span><span class="w">${weekdayLabel(date)}</span><span class="t">${texts.join(" ")}</span></div>`,
        );
      }
      return `<article class="sem-col">
        <h3>${m.label}</h3>
        ${miniMonth(model, m.year, m.monthIndex)}
        <div class="daylist">${lines.join("")}</div>
      </article>`;
    })
    .join("");
  const inner = `
    <header class="sheet-head">
      <div class="school">${esc(state.schoolName)}</div>
      <h1>${state.year}년 ${months[0].label}~${months[months.length - 1].label} 학사달력</h1>
    </header>
    <div class="sem-cals" style="grid-template-columns: repeat(${months.length}, 1fr)">${cols}</div>
  `;
  return paper("semester-cal", true, inner);
}

export function renderMonthly(state, model) {
  const monthIndex = Number(state.month);
  const year = monthIndex < 2 ? model.year + 1 : model.year;
  const { cells } = monthGrid(year, monthIndex);
  const weeks = [];
  for (let r = 0; r < 6; r++) {
    const slice = cells.slice(r * 7, r * 7 + 7);
    if (r === 5 && slice.every((d) => d.getMonth() !== monthIndex)) continue;
    weeks.push(
      `<tr>${slice
        .map((d) => {
          const info = infoOf(model, d);
          const bits = [];
          if (info?.holiday) bits.push(`<div class="hol">${esc(info.holiday)}</div>`);
          for (const ev of schoolEvents(info)) {
            bits.push(`<div class="act ${ev.type}">${esc(ev.title)}</div>`);
          }
          return `<td class="${dayClass(model, d, monthIndex)}">
            <div class="md">${formatMD(d)}</div>
            <div class="cell-ev">${bits.join("")}</div>
          </td>`;
        })
        .join("")}</tr>`,
    );
  }
  const heads = WEEKDAYS.map(
    (w, i) => `<th class="${i === 0 ? "sun" : i === 6 ? "sat" : ""}">${w}</th>`,
  ).join("");
  const inner = `
    <header class="month-head">
      <div class="ym-box">${year}년도 ${monthIndex + 1}월</div>
      <div class="school">${esc(state.schoolName)}</div>
    </header>
    <table class="month-big">
      <thead><tr>${heads}</tr></thead>
      <tbody>${weeks.join("")}</tbody>
    </table>
  `;
  return paper("monthly", true, inner);
}

export function renderSemesterWeek(state, model) {
  const sem = Number(state.semester) === 2 ? 2 : 1;
  const rows = weekRows(model).filter((row) => rowSemester(model, row) === sem);
  const monthMeta = new Map();
  for (const row of rows) {
    const id = `${row.year}-${row.monthIndex}`;
    if (!monthMeta.has(id)) monthMeta.set(id, { count: 0 });
    monthMeta.get(id).count += 1;
  }
  const seen = new Set();
  const body = [];
  for (const row of rows) {
    const id = `${row.year}-${row.monthIndex}`;
    const first = !seen.has(id);
    if (first) seen.add(id);
    const weekdays = row.days.slice(1, 6);
    const cells = weekdays
      .map((d) => {
        const inMonth = d.getMonth() === row.monthIndex && d.getFullYear() === row.year;
        const info = infoOf(model, d);
        const texts = [];
        if (inMonth && info?.holiday) texts.push(`<span class="hol">${esc(info.holiday)}</span>`);
        if (inMonth) {
          for (const ev of schoolEvents(info)) texts.push(`<span class="act ${ev.type}">${esc(ev.title)}</span>`);
        }
        const vac = info && !info.inTerm;
        return `<td class="${dayClass(model, d, row.monthIndex)}${vac ? " vac" : ""}">
          <span class="n">${d.getDate()}</span>
          <span class="ev">${texts.join("")}</span>
        </td>`;
      })
      .join("");
    const weekCount = weekdays.filter((d) => {
      const info = infoOf(model, d);
      return info?.instructional && d.getMonth() === row.monthIndex && d.getFullYear() === row.year;
    }).length;
    const monthCount = instructionalMonth(model, row.year, row.monthIndex);
    body.push(`<tr>
      ${first ? `<td class="month" rowspan="${monthMeta.get(id).count}">${row.monthIndex + 1}</td>` : ""}
      <td class="week">${row.weekNo}</td>
      ${cells}
      <td class="num">${weekCount}</td>
      ${first ? `<td class="num month-sum" rowspan="${monthMeta.get(id).count}">${monthCount}</td>` : ""}
    </tr>`);
  }

  const byWd = [1, 2, 3, 4, 5].map((wd) => ({
    off: model.days.filter((d) => d.semester === sem && d.weekday === wd && (d.holiday || d.closure)).length,
    exam: model.days.filter(
      (d) => d.semester === sem && d.weekday === wd && d.events.some((e) => e.type === "exam"),
    ).length,
    cls: countInstructional(model, (d) => d.semester === sem && d.weekday === wd),
  }));
  const sum = (k) => byWd.reduce((a, x) => a + x[k], 0);
  const s1 = countInstructional(model, (d) => d.semester === 1);
  const s2 = countInstructional(model, (d) => d.semester === 2);

  const inner = `
    <header class="sheet-head">
      <h1>${state.year}학년도 학사 일정표${sem === 2 ? " (2학기)" : ""}</h1>
      <div class="school">${esc(state.schoolName)}</div>
    </header>
    <table class="sem-week">
      <colgroup>
        <col class="sw-month" />
        <col class="sw-week" />
        <col class="sw-day" />
        <col class="sw-day" />
        <col class="sw-day" />
        <col class="sw-day" />
        <col class="sw-day" />
        <col class="sw-sum" />
        <col class="sw-sum" />
      </colgroup>
      <thead>
        <tr>
          <th class="month" rowspan="2">월</th><th class="week" rowspan="2">주</th>
          <th colspan="5">교육행사</th>
          <th class="num" rowspan="2">주계</th><th class="num" rowspan="2">월계</th>
        </tr>
        <tr><th>월</th><th>화</th><th>수</th><th>목</th><th>금</th></tr>
      </thead>
      <tbody>${body.join("")}</tbody>
    </table>
    <div class="stats">
      <table>
        <thead><tr><th></th><th>월</th><th>화</th><th>수</th><th>목</th><th>금</th><th>계</th></tr></thead>
        <tbody>
          <tr><th>${sem}학기 휴업일/공휴일</th>${byWd.map((x) => `<td>${x.off}</td>`).join("")}<td>${sum("off")}</td></tr>
          <tr><th>고사일</th>${byWd.map((x) => `<td>${x.exam}</td>`).join("")}<td>${sum("exam")}</td></tr>
          <tr><th>수업일수</th>${byWd.map((x) => `<td>${x.cls}</td>`).join("")}<td class="total">${sum("cls")}</td></tr>
        </tbody>
      </table>
      <table class="year-sum">
        <thead><tr><th>구분</th><th>기간</th><th>수업일수</th><th>연간 수업일수</th></tr></thead>
        <tbody>
          <tr><td>1학기</td><td>${esc(state.sem1Start)} ~ ${esc(state.sem1End)}</td><td>${s1}</td><td rowspan="2" class="total">${s1 + s2}</td></tr>
          <tr><td>2학기</td><td>${esc(state.sem2Start)} ~ ${esc(state.sem2End)}</td><td>${s2}</td></tr>
        </tbody>
      </table>
    </div>
    <p class="note">※ 학사일정은 학교사정에 의해 추후 변동될 수 있음.</p>
  `;
  return paper("semester-week", false, inner);
}

export const FORMATS = [
  { id: "yearly-week", label: "연간 학사일정표 (주간)", group: "연간", landscape: false },
  { id: "yearly-grid", label: "연간 학사달력 (12개월)", group: "연간", landscape: false },
  { id: "semester-cal", label: "학기 달력 (6개월)", group: "학기", landscape: true },
  { id: "semester-week", label: "학기 학사일정표 (월~금)", group: "학기", landscape: false },
  { id: "monthly", label: "월중 달력", group: "월중", landscape: true },
];

export function renderSheet(state, model) {
  switch (state.format) {
    case "yearly-grid":
      return renderYearlyGrid(state, model);
    case "semester-cal":
      return renderSemesterCal(state, model);
    case "monthly":
      return renderMonthly(state, model);
    case "semester-week":
      return renderSemesterWeek(state, model);
    default:
      return renderYearlyWeek(state, model);
  }
}
