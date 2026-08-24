export function isSampleEvents(events) {
  if (!events || events.length !== SAMPLE_EVENTS.length) return false;
  const fp = (list) =>
    [...list]
      .map((e) => `${String(e.date).slice(5)}\t${e.title}\t${e.type}`)
      .sort()
      .join("\n");
  return fp(events) === fp(SAMPLE_EVENTS);
}

export const SAMPLE_EVENTS = [
  { date: "2026-03-03", title: "입학식/개학식", type: "activity" },
  { date: "2026-03-09", title: "교직원회의", type: "activity" },
  { date: "2026-03-25", title: "교육과정설명회", type: "activity" },
  { date: "2026-04-06", title: "교직원회의", type: "activity" },
  { date: "2026-04-10", title: "스포츠 데이", type: "activity" },
  { date: "2026-04-27", title: "1차 정기시험", type: "exam" },
  { date: "2026-04-28", title: "1차 정기시험", type: "exam" },
  { date: "2026-05-04", title: "재량휴업일", type: "closure" },
  { date: "2026-05-11", title: "교직원회의", type: "activity" },
  { date: "2026-05-14", title: "수학여행/수련활동", type: "activity" },
  { date: "2026-05-15", title: "수학여행/수련활동", type: "activity" },
  { date: "2026-06-01", title: "교직원회의", type: "activity" },
  { date: "2026-07-02", title: "2차 정기시험", type: "exam" },
  { date: "2026-07-03", title: "2차 정기시험", type: "exam" },
  { date: "2026-07-15", title: "교육과정평가회", type: "activity" },
  { date: "2026-07-21", title: "여름방학식", type: "activity" },
  { date: "2026-08-11", title: "개학일", type: "activity" },
  { date: "2026-09-07", title: "교직원회의", type: "activity" },
  { date: "2026-10-02", title: "교육과정 설명회", type: "activity" },
  { date: "2026-10-06", title: "1차 정기시험", type: "exam" },
  { date: "2026-10-07", title: "1차 정기시험", type: "exam" },
  { date: "2026-10-12", title: "교직원회의", type: "activity" },
  { date: "2026-10-15", title: "현장체험학습", type: "activity" },
  { date: "2026-11-02", title: "교직원회의", type: "activity" },
  { date: "2026-11-19", title: "대학수학능력시험일", type: "closure" },
  { date: "2026-11-24", title: "2차 정기시험(3)", type: "exam" },
  { date: "2026-11-25", title: "2차 정기시험(3)", type: "exam" },
  { date: "2026-12-08", title: "2차 정기시험(1,2)", type: "exam" },
  { date: "2026-12-09", title: "2차 정기시험(1,2)", type: "exam" },
  { date: "2026-12-24", title: "축제", type: "activity" },
  { date: "2026-12-28", title: "교육과정 평가회", type: "activity" },
  { date: "2026-12-31", title: "종업식, 졸업식", type: "activity" },
  { date: "2027-02-17", title: "새학년 준비연수", type: "activity" },
  { date: "2027-02-18", title: "새학년 준비연수", type: "activity" },
  { date: "2027-02-26", title: "신학기 준비", type: "activity" },
];
