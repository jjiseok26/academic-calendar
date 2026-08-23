import { academicHolidays } from "./holidays.js";
import {
  academicRange,
  addDays,
  eachDate,
  inRange,
  parseKey,
  startOfWeek,
  toKey,
} from "./dates.js";

export function classifyTitle(title) {
  const t = title.replace(/\s/g, "");
  if (/재량휴업|휴업일/.test(t)) return "closure";
  if (/지필|고사|정기시험|시험/.test(t)) return "exam";
  if (/공휴|대체공휴|연휴/.test(t)) return "holiday";
  return "activity";
}

export function inferTerms(year, events) {
  const terms = {
    sem1Start: `${year}-03-01`,
    sem1End: `${year}-07-20`,
    sem2Start: `${year}-08-16`,
    sem2End: `${year}-12-31`,
  };
  const { start, end } = academicRange(year);
  const from = toKey(start);
  const to = toKey(end);
  const inYear = events.filter((e) => e.date >= from && e.date <= to);
  const summer = inYear.find((e) => /방학식/.test(e.title) && e.date.startsWith(String(year)));
  const reopen = inYear.find(
    (e) => /개학/.test(e.title) && e.date >= `${year}-07-01` && e.date <= `${year}-09-15`,
  );
  const close = inYear.find((e) => /종업식|졸업식/.test(e.title));
  if (summer) terms.sem1End = summer.date;
  if (reopen) terms.sem2Start = reopen.date;
  if (close) terms.sem2End = close.date;
  return terms;
}

export function buildCalendar(state) {
  const year = Number(state.year);
  const { start, end } = academicRange(year);
  const holidays = academicHolidays(year, {
    includeLaborDay: state.includeLaborDay,
    includeSuneung: state.includeSuneung,
  });

  const eventsByDate = new Map();
  for (const ev of state.events) {
    if (!ev.date || !ev.title) continue;
    const list = eventsByDate.get(ev.date) || [];
    list.push(ev);
    eventsByDate.set(ev.date, list);
  }

  const sem1Start = parseKey(state.sem1Start);
  const sem1End = parseKey(state.sem1End);
  const sem2Start = parseKey(state.sem2Start);
  const sem2End = parseKey(state.sem2End);

  const days = eachDate(start, end).map((date) => {
    const key = toKey(date);
    const holiday = holidays.get(key) || "";
    const events = eventsByDate.get(key) || [];
    const closure = events.some((e) => e.type === "closure");
    const inTerm =
      inRange(date, sem1Start, sem1End) || inRange(date, sem2Start, sem2End);
    const weekday = date.getDay();
    const instructional =
      inTerm && weekday >= 1 && weekday <= 5 && !holiday && !closure;
    const semester = inRange(date, sem1Start, sem1End)
      ? 1
      : inRange(date, sem2Start, sem2End)
        ? 2
        : 0;
    return {
      date,
      key,
      weekday,
      holiday,
      events,
      instructional,
      inTerm,
      semester,
      closure,
    };
  });

  const byKey = new Map(days.map((d) => [d.key, d]));
  return { year, start, end, days, byKey, holidays };
}

export function monthBlocks(model) {
  const months = [];
  for (let i = 0; i < 12; i++) {
    const monthIndex = (2 + i) % 12;
    const year = monthIndex < 2 ? model.year + 1 : model.year;
    months.push({ year, monthIndex, label: `${monthIndex + 1}월` });
  }
  return months;
}

export function monthGrid(year, monthIndex) {
  const first = new Date(year, monthIndex, 1);
  const gridStart = startOfWeek(first);
  const cells = [];
  for (let i = 0; i < 42; i++) cells.push(addDays(gridStart, i));
  return { first, cells };
}

export function weekRows(model) {
  const rows = [];
  let weekNo = 0;
  let cursor = startOfWeek(model.start);
  const last = startOfWeek(model.end);

  while (cursor <= last) {
    weekNo += 1;
    const days = Array.from({ length: 7 }, (_, i) => addDays(cursor, i));
    const parts = [];
    for (const d of days) {
      if (d < model.start || d > model.end) continue;
      const id = `${d.getFullYear()}-${d.getMonth()}`;
      if (!parts.includes(id)) parts.push(id);
    }
    for (const id of parts) {
      const [y, m] = id.split("-").map(Number);
      rows.push({
        weekNo,
        year: y,
        monthIndex: m,
        days,
      });
    }
    cursor = addDays(cursor, 7);
  }
  return rows;
}

export function countInstructional(model, pred) {
  return model.days.filter((d) => d.instructional && pred(d)).length;
}

export function selfCheck() {
  const model = buildCalendar({
    year: 2026,
    includeLaborDay: true,
    includeSuneung: true,
    sem1Start: "2026-03-01",
    sem1End: "2026-07-21",
    sem2Start: "2026-08-11",
    sem2End: "2026-12-31",
    events: [{ date: "2026-05-04", title: "재량휴업일", type: "closure" }],
  });
  const mar1 = model.byKey.get("2026-03-01");
  console.assert(mar1 && !mar1.instructional && mar1.holiday.includes("삼일절"));
  const mar3 = model.byKey.get("2026-03-03");
  console.assert(mar3.instructional, "3/3은 화요일 수업일");
  const may4 = model.byKey.get("2026-05-04");
  console.assert(!may4.instructional, "재량휴업일은 수업일 아님");
  const weeks = weekRows(model);
  const week5 = weeks.filter((w) => w.weekNo === 5);
  console.assert(week5.length === 2, "월을 걸치는 주는 두 행");
  const terms2027 = inferTerms(2027, [
    { date: "2026-07-21", title: "여름방학식" },
    { date: "2026-08-11", title: "개학일" },
    { date: "2026-12-31", title: "종업식, 졸업식" },
  ]);
  console.assert(terms2027.sem1End === "2027-07-20", "다른 학년도 방학식을 쓰지 않음");
  console.assert(terms2027.sem2Start === "2027-08-16", "다른 학년도 개학일을 쓰지 않음");
  console.assert(terms2027.sem2End === "2027-12-31", "다른 학년도 종업식을 쓰지 않음");
}
