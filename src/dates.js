export const WEEKDAYS = ["일", "월", "화", "수", "목", "금", "토"];

export function toKey(date) {
  const y = date.getFullYear();
  const m = String(date.getMonth() + 1).padStart(2, "0");
  const d = String(date.getDate()).padStart(2, "0");
  return `${y}-${m}-${d}`;
}

export function parseKey(key) {
  const [y, m, d] = key.split("-").map(Number);
  return new Date(y, m - 1, d);
}

export function addDays(date, n) {
  const next = new Date(date.getFullYear(), date.getMonth(), date.getDate() + n);
  return next;
}

export function academicRange(year) {
  return {
    start: new Date(year, 2, 1),
    end: new Date(year + 1, 2, 0),
  };
}

export function eachDate(start, end) {
  const dates = [];
  for (let d = new Date(start); d <= end; d = addDays(d, 1)) dates.push(new Date(d));
  return dates;
}

export function startOfWeek(date) {
  return addDays(date, -date.getDay());
}

export function weekdayLabel(date) {
  return WEEKDAYS[date.getDay()];
}

export function monthLabel(monthIndex) {
  return `${monthIndex + 1}월`;
}

export function formatMD(date) {
  return `${date.getMonth() + 1}/${date.getDate()}`;
}

export function formatMDW(date) {
  return `${date.getMonth() + 1}/${date.getDate()}(${weekdayLabel(date)})`;
}

export function isWeekend(date) {
  return date.getDay() === 0 || date.getDay() === 6;
}

export function inRange(date, start, end) {
  const t = date.getTime();
  return t >= start.getTime() && t <= end.getTime();
}

export function selfCheck() {
  const { start, end } = academicRange(2026);
  console.assert(toKey(start) === "2026-03-01", "학년도 시작은 3월 1일");
  console.assert(toKey(end) === "2027-02-28", "2026학년도 종료는 2027-02-28");
  const leap = academicRange(2027);
  console.assert(toKey(leap.end) === "2028-02-29", "윤년 2월 말일");
}
