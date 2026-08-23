import { addDays, parseKey, toKey } from "./dates.js";

// 설날·추석·부처님오신날 당일의 양력 (한국천문연구원/위키 기준)
const LUNAR = {
  2020: { seol: [1, 25], chuseok: [10, 1], buddha: [4, 30] },
  2021: { seol: [2, 12], chuseok: [9, 21], buddha: [5, 19] },
  2022: { seol: [2, 1], chuseok: [9, 10], buddha: [5, 8] },
  2023: { seol: [1, 22], chuseok: [9, 29], buddha: [5, 27] },
  2024: { seol: [2, 10], chuseok: [9, 17], buddha: [5, 15] },
  2025: { seol: [1, 29], chuseok: [10, 6], buddha: [5, 5] },
  2026: { seol: [2, 17], chuseok: [9, 25], buddha: [5, 24] },
  2027: { seol: [2, 7], chuseok: [9, 15], buddha: [5, 13] },
  2028: { seol: [1, 27], chuseok: [10, 3], buddha: [5, 2] },
  2029: { seol: [2, 13], chuseok: [9, 22], buddha: [5, 20] },
  2030: { seol: [2, 3], chuseok: [9, 12], buddha: [5, 9] },
  2031: { seol: [1, 23], chuseok: [10, 1], buddha: [5, 28] },
  2032: { seol: [2, 11], chuseok: [9, 19], buddha: [5, 16] },
  2033: { seol: [1, 31], chuseok: [9, 8], buddha: [5, 6] },
  2034: { seol: [2, 19], chuseok: [9, 28], buddha: [5, 25] },
  2035: { seol: [2, 8], chuseok: [9, 17], buddha: [5, 15] },
  2036: { seol: [1, 28], chuseok: [10, 4], buddha: [5, 3] },
  2037: { seol: [2, 15], chuseok: [9, 24], buddha: [5, 22] },
  2038: { seol: [2, 4], chuseok: [9, 13], buddha: [5, 11] },
  2039: { seol: [1, 24], chuseok: [10, 2], buddha: [4, 30] },
  2040: { seol: [2, 12], chuseok: [9, 21], buddha: [5, 18] },
};

const ELECTIONS = {
  "2022-03-09": "대통령선거",
  "2022-06-01": "지방선거",
  "2024-04-10": "국회의원선거",
  "2026-06-03": "지방선거",
  "2027-03-03": "대통령선거",
  "2028-04-12": "국회의원선거",
  "2030-06-12": "지방선거",
  "2032-03-03": "대통령선거",
  "2032-04-14": "국회의원선거",
};

const SUNEUNG = {
  2024: [11, 14],
  2025: [11, 13],
  2026: [11, 12],
  2027: [11, 18],
  2028: [11, 16],
  2029: [11, 15],
  2030: [11, 14],
  2031: [11, 13],
  2032: [11, 18],
};

const SUBSTITUTE_SINGLE = new Set([
  "삼일절",
  "어린이날",
  "부처님오신날",
  "광복절",
  "개천절",
  "한글날",
  "성탄절",
]);

function ymd(year, month, day) {
  return new Date(year, month - 1, day);
}

function put(map, date, name) {
  const key = toKey(date);
  const prev = map.get(key);
  map.set(key, prev ? `${prev}, ${name}` : name);
}

function nextOpenWeekday(map, from) {
  let d = addDays(from, 1);
  while (d.getDay() === 0 || d.getDay() === 6 || map.has(toKey(d))) {
    d = addDays(d, 1);
  }
  return d;
}

function addThreeDay(map, year, month, day, names) {
  const center = ymd(year, month, day);
  const days = [addDays(center, -1), center, addDays(center, 1)];
  days.forEach((d, i) => put(map, d, names[i]));
  return days;
}

function addSubstitutes(map) {
  const singles = [];
  const spans = [];
  for (const [key, name] of map) {
    if (name.includes("설날") && !name.includes("연휴") && !name.includes("대체")) {
      spans.push({ key, kind: "설날" });
    } else if (name.includes("추석") && !name.includes("연휴") && !name.includes("대체")) {
      spans.push({ key, kind: "추석" });
    }
    const first = name.split(",")[0].trim();
    if (SUBSTITUTE_SINGLE.has(first)) singles.push({ key, name: first });
  }

  for (const { key } of spans) {
    const center = parseKey(key);
    const days = [addDays(center, -1), center, addDays(center, 1)];
    const needs =
      days.some((d) => d.getDay() === 0) ||
      days.some((d) => {
        const n = map.get(toKey(d)) || "";
        return n.includes(",") && !n.includes("연휴");
      });
    if (needs) put(map, nextOpenWeekday(map, days[2]), "대체공휴일");
  }

  for (const { key } of singles) {
    const d = parseKey(key);
    if (d.getDay() === 0 || d.getDay() === 6) {
      put(map, nextOpenWeekday(map, d), "대체공휴일");
    }
  }
}

export function holidaysForCalendarYear(year, options = {}) {
  const map = new Map();
  const lunar = LUNAR[year];
  if (!lunar) return map;

  put(map, ymd(year, 1, 1), "신정");
  addThreeDay(map, year, lunar.seol[0], lunar.seol[1], ["설날연휴", "설날", "설날연휴"]);
  put(map, ymd(year, 3, 1), "삼일절");
  put(map, ymd(year, lunar.buddha[0], lunar.buddha[1]), "부처님오신날");
  put(map, ymd(year, 5, 5), "어린이날");
  put(map, ymd(year, 6, 6), "현충일");
  put(map, ymd(year, 8, 15), "광복절");
  addThreeDay(map, year, lunar.chuseok[0], lunar.chuseok[1], ["추석연휴", "추석", "추석연휴"]);
  put(map, ymd(year, 10, 3), "개천절");
  put(map, ymd(year, 10, 9), "한글날");
  put(map, ymd(year, 12, 25), "성탄절");

  if (options.includeLaborDay) put(map, ymd(year, 5, 1), "노동절");
  if (options.includeSuneung && SUNEUNG[year]) {
    put(map, ymd(year, SUNEUNG[year][0], SUNEUNG[year][1]), "수능일");
  }
  for (const [key, name] of Object.entries(ELECTIONS)) {
    if (key.startsWith(`${year}-`)) put(map, parseKey(key), name);
  }

  addSubstitutes(map);
  return map;
}

export function academicHolidays(year, options = {}) {
  const a = holidaysForCalendarYear(year, options);
  const b = holidaysForCalendarYear(year + 1, options);
  const out = new Map();
  for (const [key, name] of a) {
    if (key >= `${year}-03-01`) out.set(key, name);
  }
  for (const [key, name] of b) {
    if (key.startsWith(`${year + 1}-01-`) || key.startsWith(`${year + 1}-02-`)) {
      out.set(key, name);
    }
  }
  return out;
}

export function selfCheck() {
  const h = academicHolidays(2026, { includeLaborDay: true, includeSuneung: true });
  console.assert(h.get("2026-03-01")?.includes("삼일절"), "삼일절");
  console.assert(h.get("2026-03-02")?.includes("대체공휴일"), "삼일절 대체공휴일");
  console.assert(h.get("2026-09-25")?.includes("추석"), "2026 추석");
  console.assert(h.get("2026-05-24")?.includes("부처님오신날"), "부처님오신날");
  console.assert(h.get("2027-02-07")?.includes("설날"), "2027 설날");
  console.assert(h.get("2026-11-12")?.includes("수능일"), "2026 수능");
}
