/**
 * ============================================================================
 *  집합공정 생산계획 자동화 — Publisher.gs (결과 스프레드시트 렌더링)
 * ----------------------------------------------------------------------------
 *  역할 : 작업목록 배열 → 결과 파일(CFG.RESULT_SS_ID, 고정 ID) 탭 18장
 *         통합 · 집합01호기~집합10호기 · 재고흐름 · 요약 · 작업목록 · 일별생산 · 오류 · 이력
 *
 *  원칙
 *   - [작업목록]이 단일 진실 원천: 먼저 쓰고 → 다시 읽어서 → [일별생산]을 계산하고 → 모든 뷰는 그 둘만으로 그린다.
 *     시뮬레이션 결과 객체를 뷰 코드가 직접 참조하지 않는다. (CAPA·휴무·출하계획은 기준 데이터라 참조한다)
 *   - 파일은 새로 만들거나 지우지 않는다. 탭이 없으면 그 탭만 만들고, 실행마다 탭 내용만 덮어쓴다.
 *     [오류]·[이력]은 덮어쓰지 않고 append 누적.
 *   - 시트 I/O는 배치: 탭당 setValues 1회 + setBackgrounds 1회 + setNumberFormats 1회가 기본. 셀 단위 루프 금지.
 *   - 병합 셀 금지, 헤더는 1행에만(호기 탭은 3행), 일시는 Date 값.
 *   - 뷰 추가 = 렌더러 함수 추가 + CFG.PUBLISH.VIEWS에 이름 추가. 코어(Planner/Scheduler) 수정 금지.
 * ============================================================================
 */

/* 색상 (연한 톤) */
const COLOR_ = {
  OFF: '#f4cccc',      // 주말·휴무 열, 음수 재고
  CHANGE: '#fff2cc',   // 교체=Y 행
  NEW: '#fce5cd',      // 상태 신규/변경
  DONE: '#efefef',     // 상태 확정/완료
  HEAD: '#d9d9d9',     // 헤더
  ACTUAL: '#b7b7b7',   // 실적 구간 열 헤더 (규칙 11 — 계획이 아니라 실적값이라는 표시)
};
const JOB_COLS_ = ['계획ID', '호기', '순번', '품목코드', '규격', '고객사', '출하일', '시작일시', '종료일시', '소요(h)', '생산량(kg)', '교체', '청크', '상태'];
const DAILY_COLS_ = ['계획ID', '호기', '순번', '품목코드', '날짜', '가동시간(h)', '생산량(kg)'];
const HISTORY_COLS_ = ['계획ID', '시각', '트리거 사유', '작업 건수', '총 kg', '교체 횟수', '납기위험 수', '소요 초', '사본URL'];
const ERROR_COLS_ = ['시각', '단계', '메시지', '구분'];
const KO_DOW_ = ['일', '월', '화', '수', '목', '금', '토'];

/** 결과 파일 탭 순서 (왼쪽부터) */
function resultTabOrder_() {
  return [SHEET.RESULT.INTEGRATED].concat(CFG.MACHINES).concat([SHEET.RESULT.INVENTORY, SHEET.RESULT.SUMMARY, SHEET.RESULT.JOBS, SHEET.RESULT.DAILY, SHEET.RESULT.ERRORS, SHEET.RESULT.HISTORY]);
}

/* ────────────────────────────────────────────────────────────────────────────
 *  진입점
 * ──────────────────────────────────────────────────────────────────────────── */

/**
 * @param {string} planId
 * @param {Array}  jobs   Planner 작업목록 배열
 * @param {{data:object, plan?:object, reason?:string, views?:string[], updatedAt?:Date}} opts
 *        data = loadData_() 결과 (capa·holidays·ship·planMonth·machineList 참조)
 * @returns {{timing:object, jobRows:Array, daily:Array, month:object, negatives:Array}}
 */
function publishPlan_(planId, jobs, opts) {
  opts = opts || {};
  const views = opts.views || CFG.PUBLISH.VIEWS;
  const data = opts.data;
  if (!data) throw new Error('publishPlan_: opts.data(loadData_ 결과)가 필요합니다');
  const rs = openResult_();
  const timing = {};
  const tick = (name, fn) => { const t = Date.now(); const r = fn(); timing[name] = Math.round((Date.now() - t) / 100) / 10; return r; };

  tick('tabs', () => ensureResultTabs_(rs));

  /* 1) 작업목록 — 쓰고 다시 읽는다 (단일 진실 원천) */
  tick('jobs', () => writeJobsTab_(rs, jobs));
  const jobRows = readJobsTab_(rs);

  /* 2) 일별생산 — 작업목록에서 결정적으로 계산 → 검증 → 쓰기 */
  const daily = tick('daily', () => {
    const d = buildDaily_(jobRows, data);
    verifyDaily_(d, jobRows, data);
    writeDailyTab_(rs, d);
    return d;
  });
  const month = resolveMonth_(data, jobRows);

  /* 3) 뷰 */
  let negatives = [];
  if (views.indexOf('integrated') >= 0) tick('integrated', () => renderIntegrated_(rs, jobRows, daily, data, month));
  if (views.indexOf('machineTabs') >= 0) tick('machineTabs', () => renderMachineTabs_(rs, jobRows, data, planId, opts));
  if (views.indexOf('inventory') >= 0) negatives = tick('inventory', () => renderInventory_(rs, jobRows, daily, data, month, opts));
  if (views.indexOf('summary') >= 0) tick('summary', () => renderSummary_(rs, jobRows, daily, data, month, planId, opts, negatives));
  if (views.indexOf('workOrderFiles') >= 0) tick('workOrders', () => renderWorkOrderFiles_(jobRows, data, planId, opts));

  tick('order', () => orderResultTabs_(rs));
  SpreadsheetApp.flush();
  const total = Object.keys(timing).reduce((s, k) => s + timing[k], 0);
  Logger.log(`[렌더링] ${total.toFixed(1)}초 — ` + Object.keys(timing).map(k => `${k} ${timing[k]}s`).join(' · '));
  if (total > 60) warn_('렌더링', `publishPlan_ ${total.toFixed(1)}초 — 느린 렌더러: ` + Object.keys(timing).sort((a, b) => timing[b] - timing[a]).slice(0, 3).map(k => `${k} ${timing[k]}s`).join(', '));
  return { timing, jobRows, daily, month, negatives };
}

/** 탭 18장이 없으면 만든다 (있으면 그대로) */
function ensureResultTabs_(rs) {
  resultTabOrder_().forEach(name => ensureSheet_(rs, name));
}

/** 탭 순서를 명세대로 (이미 맞으면 이동하지 않음) */
function orderResultTabs_(rs) {
  const want = resultTabOrder_();
  const current = rs.getSheets().map(s => s.getName());
  const ok = want.every((n, i) => current[i] === n);
  if (ok) return;
  want.forEach((name, i) => {
    const sh = rs.getSheetByName(name);
    if (!sh) return;
    rs.setActiveSheet(sh);
    rs.moveActiveSheet(i + 1);
  });
}

/* ────────────────────────────────────────────────────────────────────────────
 *  [작업목록] — 데이터 모델 컬럼 그대로. 호기 → 순번 정렬. 필터·헤더 고정
 * ──────────────────────────────────────────────────────────────────────────── */
function writeJobsTab_(rs, jobs) {
  const sh = ensureSheet_(rs, SHEET.RESULT.JOBS);
  const order = (m) => CFG.MACHINES.indexOf(m);
  const rows = jobs.slice().sort((a, b) => order(a.호기) - order(b.호기) || a.순번 - b.순번)
    .map(j => JOB_COLS_.map(c => (j[c] === undefined || j[c] === null) ? '' : j[c]));
  const f = sh.getFilter(); if (f) f.remove();
  sh.clear();
  const n = rows.length;
  if (n) {
    sh.getRange(2, 4, n, 1).setNumberFormat('@');
    sh.getRange(2, 7, n, 1).setNumberFormat('yyyy-mm-dd');
    sh.getRange(2, 8, n, 2).setNumberFormat('yyyy-mm-dd hh:mm');
    sh.getRange(2, 10, n, 1).setNumberFormat('0.00');
    sh.getRange(2, 11, n, 1).setNumberFormat('#,##0');
  }
  sh.getRange(1, 1, n + 1, JOB_COLS_.length).setValues([JOB_COLS_].concat(rows));
  sh.getRange(1, 1, 1, JOB_COLS_.length).setFontWeight('bold').setBackground(COLOR_.HEAD);
  sh.setFrozenRows(1);
  if (n) sh.getRange(1, 1, n + 1, JOB_COLS_.length).createFilter();
}

/** [작업목록] 탭 → 행 객체 배열 (Date 유지, 품목코드 문자열). 탭이 없거나 비어 있으면 [] */
function readJobsTab_(rs) {
  const t = readTable_(rs, SHEET.RESULT.JOBS, false);
  return t.rows.map(r => ({
    계획ID: String(r['계획ID'] || ''),
    호기: normalizeMachine_(r['호기']),
    순번: toNumber_(r['순번']),
    품목코드: normalizeItemCode_(r['품목코드']),
    규격: String(r['규격'] || ''),
    고객사: String(r['고객사'] || ''),
    출하일: r['출하일'] instanceof Date ? r['출하일'] : (r['출하일'] ? toDate_(r['출하일'], '[작업목록] 출하일') : null),
    시작일시: r['시작일시'] instanceof Date ? r['시작일시'] : toDate_(r['시작일시'], '[작업목록] 시작일시'),
    종료일시: r['종료일시'] instanceof Date ? r['종료일시'] : toDate_(r['종료일시'], '[작업목록] 종료일시'),
    '소요(h)': toNumber_(r['소요(h)']),
    '생산량(kg)': toNumber_(r['생산량(kg)']),
    교체: String(r['교체'] || 'N'),
    청크: String(r['청크'] || ''),
    상태: String(r['상태'] || '계획'),
  })).filter(r => r.호기 && r.품목코드 && r.시작일시 && r.종료일시);
}

/** 이전 실행의 [작업목록] 스냅샷 (재계획 diff용, #6). 결과 파일이 없거나 탭이 비면 [] */
function readJobsTabSafe_() {
  try { return readJobsTab_(openResult_()); } catch (e) { return []; }
}

/* ────────────────────────────────────────────────────────────────────────────
 *  [일별생산] — 작업(여러 날) → 날짜별 가동시간·생산량. 일 단위 뷰의 유일한 원천
 *  가동시간 = 그 날짜에 실제 생산한 시간 (교체·첫날 준비 제외, 주말·휴무 0)
 *  생산 구간 = 종료일시에서 (생산량 ÷ kg/hr)만큼 근무시간을 거슬러 간 시각 ~ 종료일시
 *  생산량은 가동시간 비례 배분, 반올림 오차는 마지막 날에 몰아 작업별 합 = 작업 kg 보장
 * ──────────────────────────────────────────────────────────────────────────── */
function splitJobByDay_(job, capa, holidays) {
  const kg = job['생산량(kg)'];
  const runHr = kg / capa.kgPerHr;
  const end = job.종료일시;
  const prodStart = subWorkHours_(end, runHr, holidays);
  const out = [];
  let cur = alignToWorkTime_(prodStart, holidays);
  let guard = 0;
  while (cur.getTime() < end.getTime() - 1 && guard++ < 400) {
    const ds = dayStartOf_(cur);
    const de = nextDayStart_(ds);
    const segEnd = de.getTime() < end.getTime() ? de : end;
    const hr = (segEnd.getTime() - cur.getTime()) / 3600000;
    if (hr > 1e-9) out.push({ date: new Date(ds.getFullYear(), ds.getMonth(), ds.getDate()), hr, kg: 0 });
    cur = nextWorkingDayStart_(de, holidays);
  }
  if (!out.length) out.push({ date: new Date(end.getFullYear(), end.getMonth(), end.getDate()), hr: 0, kg: 0 });
  let acc = 0;
  out.forEach((d, i) => {
    if (i === out.length - 1) d.kg = Math.round((kg - acc) * 10) / 10;
    else { d.kg = Math.round(kg * d.hr / runHr); acc += d.kg; }
  });
  return out;
}

function buildDaily_(jobRows, data) {
  const rows = [];
  jobRows.forEach(j => {
    const capa = data.capa[j.품목코드];
    if (!capa) throw new Error(`[일별생산] ${j.품목코드}의 CAPA가 없어 일별 분해를 할 수 없습니다`);
    /* 상태 `완료`는 실적 기록이므로 시간 모델로 쪼개지 않고 실적일자 하루에 전액을 올린다 (규칙 11).
     * 보빈 여러 개를 병렬로 만든 실적이라 kg ÷ kg/hr이 24h를 넘을 수 있어 가동시간은 24h로 표시만 자른다. */
    if (j.상태 === '완료') {
      const ds = dayStartOf_(j.시작일시);
      rows.push({
        계획ID: j.계획ID, 호기: j.호기, 순번: j.순번, 품목코드: j.품목코드,
        날짜: new Date(ds.getFullYear(), ds.getMonth(), ds.getDate()),
        '가동시간(h)': Math.round(Math.min(j['생산량(kg)'] / capa.kgPerHr, 24) * 100) / 100,
        '생산량(kg)': j['생산량(kg)'],
      });
      return;
    }
    splitJobByDay_(j, capa, data.holidays).forEach(d => rows.push({
      계획ID: j.계획ID, 호기: j.호기, 순번: j.순번, 품목코드: j.품목코드, 날짜: d.date, '가동시간(h)': Math.round(d.hr * 100) / 100, '생산량(kg)': d.kg,
    }));
  });
  return rows;
}

/** 검증: Σ일별 = Σ작업(품목별·호기별·작업별), 호기별 하루 가동시간 ≤ 24h(첫날 21h). 어긋나면 에러 */
function verifyDaily_(daily, jobRows, data) {
  const eps = 0.51;
  const byJob = {}, byItemD = {}, byMachD = {}, hrByMachDay = {};
  daily.forEach(d => {
    const k = d.호기 + '|' + d.순번;
    byJob[k] = (byJob[k] || 0) + d['생산량(kg)'];
    byItemD[d.품목코드] = (byItemD[d.품목코드] || 0) + d['생산량(kg)'];
    byMachD[d.호기] = (byMachD[d.호기] || 0) + d['생산량(kg)'];
    const hk = d.호기 + '|' + dateKey_(d.날짜);
    hrByMachDay[hk] = (hrByMachDay[hk] || 0) + d['가동시간(h)'];
  });
  /* 실적 구간(상태 `완료`)은 보빈 병렬 생산이라 하루 24h 한도를 적용하지 않는다 */
  const actualKeys = {};
  jobRows.forEach(j => { if (j.상태 === '완료') actualKeys[j.호기 + '|' + dateKey_(dayStartOf_(j.시작일시))] = true; });
  const byItemJ = {}, byMachJ = {};
  let firstDay = null;
  jobRows.forEach(j => {
    const k = j.호기 + '|' + j.순번;
    if (Math.abs((byJob[k] || 0) - j['생산량(kg)']) > eps) throw new Error(`[일별생산] 검증 실패: ${j.호기} 순번 ${j.순번} 일별 합 ${byJob[k]} ≠ 작업 ${j['생산량(kg)']}`);
    byItemJ[j.품목코드] = (byItemJ[j.품목코드] || 0) + j['생산량(kg)'];
    byMachJ[j.호기] = (byMachJ[j.호기] || 0) + j['생산량(kg)'];
    const ds = dayStartOf_(j.시작일시);
    if (!firstDay || ds < firstDay) firstDay = ds;
  });
  Object.keys(byItemJ).forEach(c => { if (Math.abs(byItemJ[c] - (byItemD[c] || 0)) > eps * 50) throw new Error(`[일별생산] 검증 실패: 품목 ${c} Σ일별 ${byItemD[c]} ≠ Σ작업 ${byItemJ[c]}`); });
  Object.keys(byMachJ).forEach(m => { if (Math.abs(byMachJ[m] - (byMachD[m] || 0)) > eps * 50) throw new Error(`[일별생산] 검증 실패: ${m} Σ일별 ${byMachD[m]} ≠ Σ작업 ${byMachJ[m]}`); });
  Object.keys(hrByMachDay).forEach(k => {
    if (actualKeys[k]) return;
    const limit = (firstDay && k.split('|')[1] === dateKey_(firstDay)) ? 24 - CFG.PLAN.INITIAL_READY_HOURS : 24;
    if (hrByMachDay[k] > limit + 0.02) throw new Error(`[일별생산] 검증 실패: ${k} 하루 가동 ${hrByMachDay[k].toFixed(2)}h > ${limit}h`);
  });
}

function writeDailyTab_(rs, daily) {
  const sh = ensureSheet_(rs, SHEET.RESULT.DAILY);
  const f = sh.getFilter(); if (f) f.remove();
  sh.clear();
  const rows = daily.map(d => DAILY_COLS_.map(c => d[c]));
  const n = rows.length;
  if (n) {
    sh.getRange(2, 4, n, 1).setNumberFormat('@');
    sh.getRange(2, 5, n, 1).setNumberFormat('yyyy-mm-dd');
    sh.getRange(2, 6, n, 1).setNumberFormat('0.00');
    sh.getRange(2, 7, n, 1).setNumberFormat('#,##0.#');
  }
  sh.getRange(1, 1, n + 1, DAILY_COLS_.length).setValues([DAILY_COLS_].concat(rows));
  sh.getRange(1, 1, 1, DAILY_COLS_.length).setFontWeight('bold').setBackground(COLOR_.HEAD);
  sh.setFrozenRows(1);
  if (n) sh.getRange(1, 1, n + 1, DAILY_COLS_.length).createFilter();
}

/* ────────────────────────────────────────────────────────────────────────────
 *  계획월 — PLAN_MONTH, 없으면 작업 시작이 가장 빠른 달. 1일~말일 Date 배열과 헤더 문자열(9/1(화))
 * ──────────────────────────────────────────────────────────────────────────── */
function resolveMonth_(data, jobRows) {
  let y, m;
  if (data.planMonth) { y = data.planMonth.year; m = data.planMonth.month - 1; }
  else if (jobRows.length) { const first = jobRows.reduce((a, j) => (!a || j.시작일시 < a) ? j.시작일시 : a, null); y = first.getFullYear(); m = first.getMonth(); }
  else if (data.ship && data.ship.length) { const first = data.ship.reduce((a, r) => (!a || r.출하일 < a) ? r.출하일 : a, null); y = first.getFullYear(); m = first.getMonth(); }
  else { const t = new Date(); y = t.getFullYear(); m = t.getMonth(); }
  const last = new Date(y, m + 1, 0).getDate();
  const days = [];
  for (let d = 1; d <= last; d++) days.push(new Date(y, m, d));
  const keys = days.map(dateKey_);
  const labels = days.map(d => `${d.getMonth() + 1}/${d.getDate()}(${KO_DOW_[d.getDay()]})`);
  const off = days.map(d => !isWorkingDay_(new Date(d.getFullYear(), d.getMonth(), d.getDate(), CFG.PLAN.START_HOUR), data.holidays));
  const workDays = off.filter(x => !x).length;
  return { year: y, month: m + 1, label: `${y}-${String(m + 1).padStart(2, '0')}`, days, keys, labels, off, workDays };
}

/* ────────────────────────────────────────────────────────────────────────────
 *  [통합] — 월 통합 생산계획 양식. 행 = (호기×품목) 블록 3행(생산량(KG)·조장(m)·생산시간(Hr)), 열 = 고정 6열 + 날짜
 * ──────────────────────────────────────────────────────────────────────────── */
function renderIntegrated_(rs, jobRows, daily, data, month) {
  const sh = ensureSheet_(rs, SHEET.RESULT.INTEGRATED);
  sh.clear();
  const FIXED = ['호기', '품목코드', '규격', '만권조장', '구분', '합계'];
  const nd = month.days.length;
  const ncol = FIXED.length + nd;
  const colOf = {}; month.keys.forEach((k, i) => { colOf[k] = FIXED.length + i; });

  // 블록 목록: 호기 순 → 그 호기에서 처음 등장하는 품목 순
  const blocks = [];
  const seen = {};
  const order = (m) => CFG.MACHINES.indexOf(m);
  jobRows.slice().sort((a, b) => order(a.호기) - order(b.호기) || a.순번 - b.순번).forEach(j => {
    const k = j.호기 + '|' + j.품목코드;
    if (!seen[k]) { seen[k] = true; blocks.push({ machine: j.호기, item: j.품목코드, key: k }); }
  });
  // 일별 합산: 호기|품목|날짜 → {kg, hr}
  const agg = {};
  daily.forEach(d => {
    const k = d.호기 + '|' + d.품목코드 + '|' + dateKey_(d.날짜);
    const a = agg[k] = agg[k] || { kg: 0, hr: 0 };
    a.kg += d['생산량(kg)']; a.hr += d['가동시간(h)'];
  });
  // 신규/변경 하이라이트 대상 (상태 컬럼 하나만 근거)
  const hiJobs = {};
  jobRows.forEach(j => { if (j.상태 === '신규' || j.상태 === '변경') hiJobs[j.호기 + '|' + j.순번] = true; });
  const hiCells = {};   // 블록키|날짜키
  daily.forEach(d => { if (hiJobs[d.호기 + '|' + d.순번]) hiCells[d.호기 + '|' + d.품목코드 + '|' + dateKey_(d.날짜)] = true; });

  const values = [FIXED.concat(month.labels)];
  const formats = [new Array(ncol).fill('@')];
  const bg = [month.off.map(o => o ? COLOR_.OFF : null)];
  bg[0] = new Array(FIXED.length).fill(COLOR_.HEAD).concat(bg[0].map(c => c || COLOR_.HEAD));
  const rowTypes = [['생산량(KG)', '#,##0'], ['조장(m)', '#,##0'], ['생산시간(Hr)', '0.0']];

  blocks.forEach(b => {
    const capa = data.capa[b.item] || { 규격: '', lengthM: 0, bobbinKg: 1 };
    rowTypes.forEach((rt, ri) => {
      const row = new Array(ncol).fill(0);
      row[0] = ri === 0 ? b.machine : '';
      row[1] = ri === 0 ? b.item : '';
      row[2] = ri === 0 ? capa.규격 : '';
      row[3] = ri === 0 ? capa.lengthM : '';
      row[4] = rt[0];
      let sum = 0;
      month.keys.forEach((k, i) => {
        const a = agg[b.key + '|' + k];
        let v = 0;
        if (a) v = ri === 0 ? a.kg : (ri === 1 ? (capa.bobbinKg ? a.kg * capa.lengthM / capa.bobbinKg : 0) : a.hr);
        v = ri === 2 ? Math.round(v * 10) / 10 : Math.round(v);
        row[FIXED.length + i] = v;
        sum += v;
      });
      row[5] = ri === 2 ? Math.round(sum * 10) / 10 : Math.round(sum);
      values.push(row);
      formats.push(['@', '@', '@', '#,##0', '@', rt[1]].concat(new Array(nd).fill(rt[1])));
      bg.push([null, null, null, null, null, null].concat(month.keys.map(k => month.off[colOf[k] - FIXED.length] ? COLOR_.OFF : (hiCells[b.key + '|' + k] ? COLOR_.NEW : null))));
    });
  });

  const nrow = values.length;
  const rng = sh.getRange(1, 1, nrow, ncol);
  rng.setNumberFormats(formats);
  rng.setValues(values);
  rng.setBackgrounds(bg);
  rng.setHorizontalAlignment('center');
  sh.getRange(1, 1, 1, ncol).setFontWeight('bold').setBorder(true, true, true, true, true, true);
  // 블록 박스 · 고정열 세로선 · 합계열 굵은 오른선
  blocks.forEach((b, i) => sh.getRange(2 + i * 3, 1, 3, ncol).setBorder(true, true, true, true, null, null));
  if (nrow > 1) {
    sh.getRange(1, 1, nrow, FIXED.length).setBorder(null, null, null, null, true, null);
    sh.getRange(1, FIXED.length, nrow, 1).setBorder(null, null, null, true, null, null, '#000000', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  }
  sh.setFrozenRows(1);
  sh.setFrozenColumns(FIXED.length);
  sh.setColumnWidths(FIXED.length + 1, nd, 62);
  sh.setColumnWidth(1, 90); sh.setColumnWidth(2, 80); sh.setColumnWidth(3, 80); sh.setColumnWidth(4, 80); sh.setColumnWidth(5, 95); sh.setColumnWidth(6, 80);
}

/* ────────────────────────────────────────────────────────────────────────────
 *  [집합01호기]~[집합10호기] — 현장 작업지시서. 렌더러 하나로 10개
 *  1행 요약 4셀 · 3행 헤더 · 4행부터 작업 1건 = 1행. 시작일시는 싣지 않는다(순번대로 이어서 작업 → 예상 종료시간만)
 *  opts.horizonDays 가 있으면 오늘~D+(n-1) 작업만 (현장 배포용, 대회 중엔 미사용)
 * ──────────────────────────────────────────────────────────────────────────── */
/* 현장 작업지시서 컬럼 — 고객사·청크는 현장에서 쓰지 않아 뺐다(사용자 지시).
 * 교체는 앞 작업과 규격이 달라 교체가 필요한지를 현장이 바로 봐야 해서 열로 둔다(행 노란색과 함께).
 * 청크 분할은 같은 품목이 연달아 나오는 것으로 알 수 있고, 고객사·청크 값은 [작업목록] 탭에 남아 있다. */
const MACHINE_TAB_COLS_ = ['순번', '품목코드', '규격', '생산량(kg)', '조장(m)', '보빈수', '예상 종료시간', '소요시간(h)', '출하일', '교체', '상태'];

/**
 * 호기 작업지시 한 장을 시트에 그린다 — 결과 파일의 호기 탭과 호기별 개별 파일이 이 함수 하나를 공유한다.
 * 1행 요약 4셀 · 2행 공백 · 3행 헤더 · 4행부터 작업 1건 = 1행. 시트 I/O는 setValues 1회 + setBackgrounds 1회.
 */
function writeMachineOrder_(sh, machine, mj, data, planId, updated) {
  const ncol = MACHINE_TAB_COLS_.length;
  sh.clear();
  const kg = mj.reduce((s, j) => s + j['생산량(kg)'], 0);
  const chg = mj.filter(j => j.교체 === 'Y').length;
  const head = [[`${machine} 작업지시${mj.length ? '' : ' — 이번 달 배정 없음'}`, `계획ID ${planId}`, `갱신 ${updated}`,
    `이번 달 작업 ${mj.length}건 · 총 ${Math.round(kg).toLocaleString()}kg · 교체 ${chg}회`]];
  sh.getRange(1, 1, 1, 4).setValues(head).setFontWeight('bold');
  const rows = mj.map(j => {
    const capa = data.capa[j.품목코드] || { lengthM: 0, bobbinKg: 0 };
    const q = j['생산량(kg)'];
    return [j.순번, j.품목코드, j.규격, q, capa.bobbinKg ? Math.round(q * capa.lengthM / capa.bobbinKg) : 0, capa.bobbinKg ? Math.round(q / capa.bobbinKg * 100) / 100 : 0,
      j.종료일시, j['소요(h)'], j.출하일 || '', j.교체, j.상태];
  });
  const n = rows.length;
  const body = sh.getRange(3, 1, n + 1, ncol);
  if (n) {
    sh.getRange(4, 2, n, 1).setNumberFormat('@');
    sh.getRange(4, 4, n, 2).setNumberFormat('#,##0');
    sh.getRange(4, 6, n, 1).setNumberFormat('0.##');
    sh.getRange(4, 7, n, 1).setNumberFormat('yyyy-mm-dd hh:mm');
    sh.getRange(4, 8, n, 1).setNumberFormat('0.0');
    sh.getRange(4, 9, n, 1).setNumberFormat('yyyy-mm-dd');
  }
  body.setValues([MACHINE_TAB_COLS_].concat(rows));
  const bg = [new Array(ncol).fill(COLOR_.HEAD)].concat(mj.map(j => {
    const c = j.교체 === 'Y' ? COLOR_.CHANGE : ((j.상태 === '신규' || j.상태 === '변경') ? COLOR_.NEW : ((j.상태 === '확정' || j.상태 === '완료') ? COLOR_.DONE : null));
    return new Array(ncol).fill(c);
  }));
  body.setBackgrounds(bg);
  body.setHorizontalAlignment('center').setBorder(true, true, true, true, true, true);
  sh.getRange(3, 1, 1, ncol).setFontWeight('bold');
  /* 현장 오독 방지(제조현장 검토 반영) */
  sh.getRange(3, 8).setNote(`소요시간(h) = 생산량 ÷ kg/hr (+규격교체 ${CFG.PLAN.CHANGE_HOURS}h). 계획 첫날 첫 작업의 가동준비 ${CFG.PLAN.INITIAL_READY_HOURS}h는 여기 포함되지 않고 예상 종료시간에 반영됩니다. 순번대로 이어서 작업하세요.`);
  sh.getRange(3, 7).setNote('계획 작업의 예상 종료 시각입니다. 상태가 `완료`인 회색 행은 이미 만든 실적 기록이라 이 시각이 실제 종료 시각과 다를 수 있습니다.');
  sh.getRange(3, 10).setNote(`Y면 앞 작업과 규격이 달라 규격교체가 필요합니다. 교체 시간 ${CFG.PLAN.CHANGE_HOURS}h는 소요시간과 예상 종료시간에 이미 들어 있습니다. 이 행은 노란색으로 표시됩니다.`);
  sh.getRange(3, 11).setNote('계획: 이번에 세운 계획 · 신규: 이전 계획에 없던 작업 · 변경: 시각이나 수량이 달라진 작업 · 확정: 기준일 전 시작해 그대로 두는 작업 · 완료: 실적으로 대체된 작업\n\n행 색: 노란색 = 규격교체 필요 · 주황색 = 신규·변경 · 회색 = 확정·완료');
  sh.setFrozenRows(3);
  sh.setColumnWidth(7, 130); sh.setColumnWidth(10, 50); sh.setColumnWidth(11, 70);
  return { count: mj.length, kg, changeovers: chg };
}

/** 그 호기의 작업만 순번 순으로 (opts.horizonDays가 있으면 오늘~D+n-1 구간만 — 현장 배포용) */
function machineJobs_(jobRows, machine, opts) {
  opts = opts || {};
  let mj = jobRows.filter(j => j.호기 === machine).sort((a, b) => a.순번 - b.순번);
  if (opts.horizonDays) {
    const t = opts.today || new Date();
    const from = new Date(t.getFullYear(), t.getMonth(), t.getDate());
    const to = new Date(from.getTime()); to.setDate(to.getDate() + opts.horizonDays);
    mj = mj.filter(j => j.종료일시 >= from && j.시작일시 < to);
  }
  return mj;
}

function renderMachineTabs_(rs, jobRows, data, planId, opts) {
  opts = opts || {};
  const updated = fmtDate_(opts.updatedAt || new Date(), 'yyyy-MM-dd HH:mm');
  (opts.machines || CFG.MACHINES).forEach(m => {
    writeMachineOrder_(ensureSheet_(rs, m), m, machineJobs_(jobRows, m, opts), data, planId, updated);
  });
}

/* ────────────────────────────────────────────────────────────────────────────
 *  작업지시서/ 폴더 — 호기마다 파일 하나 (현장 배포용)
 *  각 호기 반장은 자기 파일 링크(또는 QR) 하나만 본다. 다른 호기 일정·재고·기준정보는 보이지 않는다.
 *  파일은 최초 1회만 만들고 ID를 스크립트 속성에 저장해 계속 재사용한다 — 링크·QR이 바뀌면 안 된다.
 * ──────────────────────────────────────────────────────────────────────────── */
function renderWorkOrderFiles_(jobRows, data, planId, opts) {
  opts = opts || {};
  const updated = fmtDate_(opts.updatedAt || new Date(), 'yyyy-MM-dd HH:mm');
  const out = [];
  const failed = [];
  (opts.machines || CFG.MACHINES).forEach(m => {
    /* 한 호기가 실패해도(권한·Drive 오류) 나머지 호기와 전체 실행을 죽이지 않는다.
     * 결과 파일은 이미 다 썼으므로, 여기서 예외가 나가면 백업 사본·이력·알림까지 잃는다. */
    try {
      const ss = openWorkOrderSs_(m);
      const sh = ensureSheet_(ss, SHEET.WORKORDER);
      const r = writeMachineOrder_(sh, m, machineJobs_(jobRows, m, opts), data, planId, updated);
      /* 첫 생성 시 남아 있는 기본 시트(시트1 등)를 지운다 — 현장이 빈 탭을 보지 않게.
       * getSheets()는 스냅샷 배열이라 순회 중 삭제해도 안전하고, 작업지시 탭은 건너뛰므로 마지막 시트가 남는다. */
      ss.getSheets().forEach(s => { if (s.getName() !== SHEET.WORKORDER) { try { ss.deleteSheet(s); } catch (e) { /* 마지막 시트는 삭제 불가 */ } } });
      out.push({ machine: m, url: ss.getUrl(), count: r.count, kg: r.kg });
    } catch (e) {
      failed.push(m);
      warn_('작업지시서', `${m} 파일을 갱신하지 못했습니다 — 이 호기 링크는 이전 내용 그대로입니다: ${e.message}`);
    }
  });
  Logger.log('[작업지시서] 호기별 파일 ' + out.length + '개 갱신 — ' + out.map(o => `${o.machine} ${o.count}건`).join(' · ')
    + (failed.length ? ` · 실패 ${failed.length}개: ${failed.join(', ')}` : ''));
  return out;
}

/* ────────────────────────────────────────────────────────────────────────────
 *  [재고흐름] — 품목당 5행(생산계획 · 생산량 · 출하계획 · 출하량 · 재고), 열 = 품목코드 · 기초재고 · 구분 · 합계 + 계획월 1일~말일
 *  음수 = 납기위험(연빨강). 맨 아래 일일 생산 합계·일일 출하 합계. 반환: 최초 음수 발생 [{item, date, value}]
 * ──────────────────────────────────────────────────────────────────────────── */
function renderInventory_(rs, jobRows, daily, data, month, opts) {
  opts = opts || {};
  const plan = opts.plan || {};
  const act = (plan.actuals && plan.actuals.applied) ? plan.actuals : null;
  const asOfKey = act ? dateKey_(act.asOf) : null;
  const sh = ensureSheet_(rs, SHEET.RESULT.INVENTORY);
  sh.clear();
  const items = CFG.ITEMS.filter(c => data.capa[c] || data.stock[c] !== undefined)
    .concat(Object.keys(data.capa).filter(c => CFG.ITEMS.indexOf(c) < 0).sort());
  const nd = month.days.length;

  /* 4개 맵 — 계획과 실적을 따로 모은다.
   *  생산계획: [일별생산]에서 상태가 `완료`가 아닌 작업분
   *  생산량  : [생산실적] (= [일별생산]의 `완료` 작업분과 같은 값)
   *  출하계획: [출하계획] 탭 + 이월분(이월된 날짜에 더한다)
   *  출하량  : [출하실적] */
  const completedSeq = {};
  (jobRows || []).forEach(j => { if (j.상태 === '완료') completedSeq[j.호기 + '|' + j.순번] = true; });
  const prodPlan = {}, prodAct = {}, shipPlan = {}, shipAct = {};
  daily.forEach(d => {
    const k = d.품목코드 + '|' + dateKey_(d.날짜);
    const target = completedSeq[d.호기 + '|' + d.순번] ? prodAct : prodPlan;
    target[k] = (target[k] || 0) + d['생산량(kg)'];
  });
  data.ship.forEach(r => { const k = r.품목코드 + '|' + dateKey_(r.출하일); shipPlan[k] = (shipPlan[k] || 0) + r.출하량; });
  (plan.carryOver || []).forEach(x => { const k = x.품목코드 + '|' + dateKey_(x.출하일); shipPlan[k] = (shipPlan[k] || 0) + x.출하량; });
  if (act) Object.keys(act.raw.shipByItemDate).forEach(k => { shipAct[k] = (shipAct[k] || 0) + act.raw.shipByItemDate[k]; });

  const FIXED = ['품목코드', '기초재고', '구분', '합계'];
  const SUB = [
    { label: '생산계획', map: prodPlan },
    { label: '생산량', map: prodAct },
    { label: '출하계획', map: shipPlan },
    { label: '출하량', map: shipAct },
  ];
  const BLOCK = SUB.length + 1;                        // 품목당 5행 (+ 재고)
  const ncol = FIXED.length + nd;

  const headBg = [COLOR_.HEAD, COLOR_.HEAD, COLOR_.HEAD, COLOR_.HEAD]
    .concat(month.keys.map((k, i) => (act && k < asOfKey) ? COLOR_.ACTUAL : (month.off[i] ? COLOR_.OFF : COLOR_.HEAD)));
  const values = [FIXED.concat(month.labels)];
  const bg = [headBg];
  const negatives = [];
  const dailyProd = new Array(nd).fill(0), dailyShip = new Array(nd).fill(0);
  const offBg = month.off.map(o => o ? COLOR_.OFF : null);

  items.forEach(item => {
    /* 재고 흐름 — 기준일 전은 실적, 기준일 당일은 실적(오전) + 계획(잔여), 기준일 이후는 계획 (규칙 11) */
    let stock = data.stock[item] || 0;
    const stockCells = [];
    let firstNeg = null;
    month.keys.forEach((k, i) => {
      const pa = prodAct[item + '|' + k] || 0, pp = prodPlan[item + '|' + k] || 0;
      const sa = shipAct[item + '|' + k] || 0, sp = shipPlan[item + '|' + k] || 0;
      let p, s;
      if (act && k < asOfKey) { p = pa; s = sa; }
      else if (act && k === asOfKey) { p = pa + pp; s = sa > 0 ? sa : sp; }
      else { p = pp; s = sp; }
      dailyProd[i] += p; dailyShip[i] += s;
      stock = stock + p - s;
      const v = Math.round(stock);
      stockCells.push(v);
      if (v < 0 && !firstNeg) firstNeg = { item, date: month.days[i], value: v };
    });
    if (firstNeg) negatives.push(firstNeg);

    SUB.forEach((sub, ri) => {
      const cells = month.keys.map(k => Math.round(sub.map[item + '|' + k] || 0));
      values.push([ri === 0 ? item : '', ri === 0 ? Math.round(data.stock[item] || 0) : '', sub.label,
        Math.round(cells.reduce((s, v) => s + v, 0))].concat(cells));
      bg.push([null, null, null, null].concat(offBg));
    });
    /* 재고 행 — 합계 열에는 월말 재고를 쓰고, 음수(납기위험)는 연빨강 */
    values.push(['', '', '재고', stockCells[stockCells.length - 1]].concat(stockCells));
    bg.push([null, null, null, stockCells[stockCells.length - 1] < 0 ? COLOR_.OFF : null]
      .concat(stockCells.map((v, i) => v < 0 ? COLOR_.OFF : offBg[i])));
  });

  values.push(['', '', '일일 생산 합계', Math.round(dailyProd.reduce((s, v) => s + v, 0))].concat(dailyProd.map(v => Math.round(v))));
  values.push(['', '', '일일 출하 합계', Math.round(dailyShip.reduce((s, v) => s + v, 0))].concat(dailyShip.map(v => Math.round(v))));
  bg.push([null, null, null, null].concat(offBg));
  bg.push([null, null, null, null].concat(offBg));

  const nrow = values.length;
  const rng = sh.getRange(1, 1, nrow, ncol);
  sh.getRange(2, 1, nrow - 1, 1).setNumberFormat('@');
  sh.getRange(2, 2, nrow - 1, 1).setNumberFormat('#,##0');
  sh.getRange(2, 4, nrow - 1, ncol - 3).setNumberFormat('#,##0');
  rng.setValues(values);
  rng.setBackgrounds(bg);
  rng.setHorizontalAlignment('center').setBorder(true, true, true, true, true, true);
  sh.getRange(1, 1, 1, ncol).setFontWeight('bold');
  /* 품목 블록(5행)을 굵은 테두리로 감싼다 — 품목 경계가 한눈에 보이게 (사용자 요청)
   * 안쪽 격자는 위에서 이미 얇은 선으로 그렸고, 여기서는 블록 바깥 4면만 굵게 덮어쓴다 */
  items.forEach((it, i) => sh.getRange(2 + i * BLOCK, 1, BLOCK, ncol)
    .setBorder(true, true, true, true, null, null, '#000000', SpreadsheetApp.BorderStyle.SOLID_MEDIUM));
  /* 헤더 행 아래도 굵게 — 헤더와 첫 품목 블록을 분리 */
  sh.getRange(1, 1, 1, ncol).setBorder(null, null, true, null, null, null, '#000000', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  sh.getRange(1, FIXED.length, nrow, 1).setBorder(null, null, null, true, null, null, '#000000', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  sh.getRange(nrow - 1, 1, 2, ncol).setFontWeight('bold');
  sh.setFrozenRows(1); sh.setFrozenColumns(FIXED.length);
  sh.setColumnWidths(FIXED.length + 1, nd, 62);
  sh.setColumnWidth(1, 90); sh.setColumnWidth(2, 80); sh.setColumnWidth(3, 85); sh.setColumnWidth(4, 80);
  return negatives;
}

/* ────────────────────────────────────────────────────────────────────────────
 *  [요약] — 대상월 · 품목별 총 생산/출하/월말 재고 · 일 평균 생산 · 일 최대 출하 · 교체 횟수 · 납기위험 · 이월 · 적정재고 미달 · 호기별 가동률
 * ──────────────────────────────────────────────────────────────────────────── */
function renderSummary_(rs, jobRows, daily, data, month, planId, opts, negatives) {
  const sh = ensureSheet_(rs, SHEET.RESULT.SUMMARY);
  sh.clear();
  const plan = opts.plan || {};
  const act = (plan.actuals && plan.actuals.applied) ? plan.actuals : null;
  const asOfKey = act ? dateKey_(act.asOf) : null;
  /* 실적으로 대체된 `완료` 작업은 "이번 계획으로 만들 양"이 아니므로 계획 집계에서 뺀다 (규칙 11) */
  const plannedRows = jobRows.filter(j => j.상태 !== '완료');
  const completedRows = jobRows.length - plannedRows.length;
  const totalKg = plannedRows.reduce((s, j) => s + j['생산량(kg)'], 0);
  const changeovers = plannedRows.filter(j => j.교체 === 'Y').length;
  const inMonth = (d) => d && d.getFullYear() === month.year && d.getMonth() + 1 === month.month;

  // 품목별 생산 = [일별생산] (실적 구간은 `완료` 작업 = 실적값)
  const prodByItem = {}, shipByItem = {}, shipByDay = {};
  daily.forEach(d => { if (inMonth(d.날짜)) prodByItem[d.품목코드] = (prodByItem[d.품목코드] || 0) + d['생산량(kg)']; });
  /* 출하 = 재고흐름과 같은 기준: 실적 구간은 [출하실적], 기준일부터는 [출하계획] + 이월분 */
  const addShip = (code, d, qty) => {
    if (!inMonth(d)) return;
    shipByItem[code] = (shipByItem[code] || 0) + qty;
    const k = dateKey_(d);
    shipByDay[k] = (shipByDay[k] || 0) + qty;
  };
  data.ship.forEach(r => { if (!(act && dateKey_(r.출하일) < asOfKey)) addShip(r.품목코드, r.출하일, r.출하량); });
  if (act) {
    Object.keys(act.raw.shipByItemDate).forEach(k => {
      const p = k.split('|');
      if (p[1] >= asOfKey) return;
      addShip(p[0], toDate_(p[1], '[출하실적] 출고일자'), act.raw.shipByItemDate[k]);
    });
    (plan.carryOver || []).forEach(x => addShip(x.품목코드, x.출하일, x.출하량));
  }
  const items = CFG.ITEMS.concat(Object.keys(data.capa).filter(c => CFG.ITEMS.indexOf(c) < 0).sort());
  let maxShipDay = '', maxShip = 0;
  Object.keys(shipByDay).forEach(k => { if (shipByDay[k] > maxShip) { maxShip = shipByDay[k]; maxShipDay = k; } });
  const monthProd = Object.keys(prodByItem).reduce((s, k) => s + prodByItem[k], 0);
  // 적정재고 미달
  const shortfalls = (plan.replenish || []).filter(r => r.after < r.target - 1e-9).map(r => `${r.item}(${Math.round(r.target - r.after).toLocaleString()}kg 부족)`);
  const carryList = plan.carryOver || [];   // 미출하 이월 (규칙 11)

  const kv = [
    ['대상월', month.label],
    ['계획ID', planId],
    ['생성 시각', fmtDate_(opts.updatedAt || new Date(), 'yyyy-MM-dd HH:mm')],
    ['트리거 사유', opts.reason || '수동'],
    ['기준일(AS_OF_DATE)', act ? `${asOfKey} — 이 날짜부터가 계획, 이전은 실적` : '실적 파일 없음 (전 구간 계획)'],
    ['실적 반영', act ? `생산실적 ${act.raw.prodRows}행 · 출하실적 ${act.raw.shipRows}행 → 완료 작업 ${completedRows}건` : '없음'],
    ['총 작업 건수', `${plannedRows.length}건` + (completedRows ? ` (+ 완료 ${completedRows}건)` : '')],
    ['총 생산량(kg)', Math.round(totalKg)],
    ['규격교체 총 횟수', changeovers],
    ['근무일수', month.workDays],
    ['일 평균 생산량(kg)', month.workDays ? Math.round(monthProd / month.workDays) : 0],
    ['일 최대 출하량(kg)', maxShipDay ? `${Math.round(maxShip).toLocaleString()} (${maxShipDay})` : '-'],
    ['납기위험 품목(최초 음수일)', negatives.length ? negatives.map(n => `${n.item} ${dateKey_(n.date)} (${n.value.toLocaleString()})`).join(', ') : '없음'],
    ['이월', carryList.length
      ? `${carryList.length}건 — ` + carryList.slice(0, 3).map(x => `${x.품목코드} ${x.고객사} ${dateKey_(x.원출하일)}→${dateKey_(x.출하일)} ${Math.round(x.출하량).toLocaleString()}kg`).join(', ') + (carryList.length > 3 ? ` 외 ${carryList.length - 3}건` : '')
      : '0건'],
    ['적정재고 미달 품목', shortfalls.length ? shortfalls.join(', ') : '없음'],
    ['미배정 수요', (plan.unassigned || []).length ? (plan.unassigned || []).map(d => `${d.item} ${dateKey_(d.dueDate)}`).join(', ') : '없음'],
  ];
  const values = [['항목', '값']].concat(kv);
  values.push(['', '']);
  values.push(['품목코드', '기초재고(kg)', '계획시작 재고(kg)', '총 생산(kg)', '총 출하(kg)', '월말 재고(kg)']);
  items.forEach(c => {
    const p = prodByItem[c] || 0, s = shipByItem[c] || 0, base = data.stock[c] || 0;
    const start = act && act.stockStart[c] !== undefined ? act.stockStart[c] : base;
    values.push([c, Math.round(base), Math.round(start), Math.round(p), Math.round(s), Math.round(base + p - s)]);
  });
  // 호기별 가동률 = Σ점유시간(준비·교체·가동 포함) ÷ (근무일수 × 24h). 95% 이상은 고부하 경고(제조현장 검토 반영)
  const capHr = month.workDays * 24;
  const util = CFG.MACHINES.map(m => {
    const mj = plannedRows.filter(j => j.호기 === m);   // 계획 부하 기준 (실적 구간 제외)
    const occ = mj.reduce((s, j) => s + workHoursBetween_(j.시작일시, j.종료일시, data.holidays), 0);
    return { m, n: mj.length, occ, pct: capHr ? occ / capHr * 100 : 0 };
  });
  const HIGH_LOAD_PCT = 95;
  const highLoad = util.filter(u => u.pct >= HIGH_LOAD_PCT);
  if (highLoad.length) warn_('요약', `고부하 호기(가동률 ${HIGH_LOAD_PCT}% 이상): ` + highLoad.map(u => `${u.m} ${u.pct.toFixed(1)}%(여유 ${Math.max(0, capHr - u.occ).toFixed(0)}h)`).join(', ') + ' — 추가 수요·설비 장애 시 납기위험');
  values.splice(kv.length + 1, 0, ['⚠ 고부하 호기(≥95%)', highLoad.length ? highLoad.map(u => `${u.m} ${u.pct.toFixed(1)}% (여유 ${Math.max(0, capHr - u.occ).toFixed(0)}h)`).join(', ') : '없음']);
  values.push(['', '']);
  values.push(['호기', '작업 건수', '점유시간(h)', '가동률(%)']);
  util.forEach(u => values.push([u.m, u.n, Math.round(u.occ * 10) / 10, Math.round(u.pct * 10) / 10]));
  const ncol = 6;
  const rows = values.map(r => { const c = r.slice(); while (c.length < ncol) c.push(''); return c; });
  sh.getRange(1, 1, rows.length, ncol).setValues(rows);
  sh.getRange(1, 1, 1, 2).setFontWeight('bold').setBackground(COLOR_.HEAD);
  const itemHeadRow = kv.length + 4, machHeadRow = itemHeadRow + items.length + 2;   // +1: 고부하 행 삽입
  sh.getRange(itemHeadRow, 1, 1, ncol).setFontWeight('bold').setBackground(COLOR_.HEAD);
  sh.getRange(machHeadRow, 1, 1, 4).setFontWeight('bold').setBackground(COLOR_.HEAD);
  sh.getRange(itemHeadRow + 1, 1, items.length, 1).setNumberFormat('@');
  sh.getRange(itemHeadRow + 1, 2, items.length, 5).setNumberFormat('#,##0');
  sh.getRange(machHeadRow + 1, 3, CFG.MACHINES.length, 1).setNumberFormat('#,##0.0');
  sh.getRange(machHeadRow + 1, 4, CFG.MACHINES.length, 1).setNumberFormat('0.0');
  sh.setColumnWidth(1, 200); sh.setColumnWidth(2, 260);
}

/* ────────────────────────────────────────────────────────────────────────────
 *  [오류] · [이력] — 결과 파일에만 append (기준정보·출하계획 파일에 쓰면 트리거 무한 루프)
 * ──────────────────────────────────────────────────────────────────────────── */
function appendError_(stage, message, kind) {
  try {
    const rs = openResult_();
    const sh = ensureSheet_(rs, SHEET.RESULT.ERRORS);
    if (sh.getLastRow() === 0) { sh.appendRow(ERROR_COLS_); sh.getRange(1, 1, 1, ERROR_COLS_.length).setFontWeight('bold').setBackground(COLOR_.HEAD); sh.setFrozenRows(1); }
    sh.appendRow([new Date(), String(stage), String(message), kind || '오류']);
  } catch (e) {
    Logger.log(`[오류 탭 기록 실패] ${stage}: ${message} — ${e.message}`);
  }
}

function appendHistory_(h) {
  try {
    const rs = openResult_();
    const sh = ensureSheet_(rs, SHEET.RESULT.HISTORY);
    if (sh.getLastRow() === 0) { sh.appendRow(HISTORY_COLS_); sh.getRange(1, 1, 1, HISTORY_COLS_.length).setFontWeight('bold').setBackground(COLOR_.HEAD); sh.setFrozenRows(1); }
    sh.appendRow([h.planId || '-', new Date(), h.reason || '', h.jobs === undefined ? '' : h.jobs, h.kg === undefined ? '' : Math.round(h.kg), h.changeovers === undefined ? '' : h.changeovers, h.late === undefined ? '' : h.late, h.sec === undefined ? '' : h.sec, h.backupUrl || (h.note || '')]);
  } catch (e) {
    Logger.log(`[이력 탭 기록 실패] ${e.message}`);
  }
}

/** 결과 파일 사본을 일별생산계획/ 폴더에 "YYMMDD-HHmm_생산계획"으로 저장하고 URL 반환 (= 기존 운영의 일일생산계획 파일) */
function backupResult_(planId) {
  const rs = openResult_();
  const folder = openBackupFolder_();
  const copy = DriveApp.getFileById(rs.getId()).makeCopy(`${planId}_${CFG.FILE_NAMES.RESULT}`, folder);
  return copy.getUrl();
}
