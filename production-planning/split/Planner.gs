/**
 * ============================================================================
 *  집합공정 생산계획 자동화 — Planner.gs (호기 배정)
 * ----------------------------------------------------------------------------
 *  역할 : 출하계획(수요) + 기준정보 → 호기 배정 → 작업목록 배열 반환.
 *   ① 수요 산출  buildDemands_      : 출하일 순 정렬 → 재고 차감 → 부족분을 보빈(CAPA 중량(KG)) 배수로 올림
 *   ② 호기 배정  chooseMachine_     : 절대차단 → Payoff → 규격교체 회피(8-1) → 전담 예약(8-2) → 우선설비(2) → 동일규격(3) → 최속(4)
 *   ③ 적정재고   replenishTargetStock_: 납기 작업 뒤 여유 시간에만 (규칙 12)
 *  시간 계산은 Scheduler.gs(createSimulation_)에 맡긴다.
 *  ⚠ 이 파일에는 시트 서식·I/O 코드를 넣지 않는다. 결과는 배열로만 반환한다 (뷰는 Publisher.gs).
 * ============================================================================
 */

/**
 * 생산계획 생성 — 진입점
 * @param {object} data  loadData_() 결과
 * @param {{planId?:string, planStart?:Date, horizonEnd?:Date}} [opts]
 *        planStart 기본 = 계획월(PLAN_MONTH) 1일 08:00을 첫 근무일로 정렬 (전 구간 계획). 실적 반영(#6)은 기준일을 넘긴다
 *        horizonEnd 기본 = 계획월 말일의 근무일 끝(말일+1 08:00)
 * @returns {{planId, planStart, horizonEnd, jobs, demands, unassigned, lateDemands, replenish, stockEnd}}
 */
function generateProductionPlan_(data, opts) {
  opts = opts || {};
  const planId = opts.planId || fmtDate_(new Date(), 'yyMMdd-HHmm');
  /* 실적 반영(규칙 11): applyActuals_ 결과가 오면 기준일부터만 계획하고 시작 재고·호기 상태를 실적에서 이어받는다.
   * 실적 파일이 없으면 act = null → 기존 동작(계획월 1일부터 전 구간 계획). */
  const act = (opts.actuals && opts.actuals.applied) ? opts.actuals : null;
  const planStart = alignToWorkTime_(opts.planStart || (act ? act.planStart : defaultPlanStart_(data)), data.holidays);
  const horizonEnd = opts.horizonEnd || defaultHorizonEnd_(data, planStart);
  const sim = createSimulation_(data, planStart, act ? { machineAvailable: act.machineAvailable, currentItem: act.currentItem, warm: act.warm } : null);
  const stock = Object.assign({}, act ? act.stockStart : data.stock);   // 작업용 재고 사본 (수요 차감·올림분 반영)

  /* ① 수요 산출 — 실적 반영 시 기준일 이후 출하 + 이월분만 대상 */
  const demands = buildDemands_(data, stock, act ? act.demandShipRows : data.ship);

  /* ② 호기 배정 — 출하일 빠른 순(동일 출하일은 품목코드 순) */
  const remaining = {};                                  // 규칙 8-2용: 품목별 아직 배정되지 않은 수요 건수
  demands.forEach(d => { remaining[d.item] = (remaining[d.item] || 0) + 1; });
  const result = { planId, planStart, horizonEnd, jobs: [], demands, unassigned: [], lateDemands: [], replenish: [], stockEnd: null,
    asOf: act ? act.asOf : (data.asOf || null), carryOver: act ? act.carryOver : [], actuals: act || null,
    stockStart: Object.assign({}, act ? act.stockStart : data.stock) };

  demands.forEach(d => {
    remaining[d.item]--;
    if (!data.capa[d.item]) { d.status = '배정불가(CAPA 없음)'; result.unassigned.push(d); return; }
    const pick = chooseMachine_(sim, data, d, remaining);
    if (!pick) {
      d.status = '배정불가(가용 호기 없음 — 절대차단/Payoff 확인)';
      result.unassigned.push(d);
      warn_('배정', `${d.item} ${dateKey_(d.dueDate)} ${d.qty}kg: 배정 가능한 호기가 없습니다 (절대차단·Payoff 제약)`);
      return;
    }
    const jobs = sim.commit(pick, { kind: '납기', demandId: d.id, dueDate: d.dueDate, deadline: d.deadline, customers: d.customers });
    d.machine = pick.machine; d.start = pick.start; d.end = pick.end; d.changeover = pick.changeover;
    d.late = pick.end.getTime() > d.deadline.getTime();
    d.status = d.late ? '납기위험' : '배정';
    if (d.late) result.lateDemands.push(d);
    jobs.forEach(j => { j._late = d.late; result.jobs.push(j); });
  });

  /* ③ 적정재고 보충 (규칙 12) */
  replenishTargetStock_(sim, data, stock, horizonEnd, result);
  result.stockEnd = stock;

  /* 실적 구간 작업(상태 `완료`)을 앞에 붙인다 — 호기 탭·[통합]에서 "어제까지 만든 것"이 보이게 */
  if (act) act.completedJobs.forEach(j => result.jobs.push(j));

  finalizeJobs_(result.jobs, data, planId);
  return result;
}

/** 계획 시작 기본값 — PLAN_MONTH 1일 08:00 (없으면 가장 빠른 출하일의 월 1일, 그것도 없으면 오늘) */
function defaultPlanStart_(data) {
  if (data.planMonth) return new Date(data.planMonth.year, data.planMonth.month - 1, 1, CFG.PLAN.START_HOUR, 0, 0, 0);
  if (data.ship && data.ship.length) {
    const first = data.ship.reduce((a, r) => (!a || r.출하일 < a) ? r.출하일 : a, null);
    return new Date(first.getFullYear(), first.getMonth(), 1, CFG.PLAN.START_HOUR, 0, 0, 0);
  }
  const t = new Date();
  return new Date(t.getFullYear(), t.getMonth(), t.getDate(), CFG.PLAN.START_HOUR, 0, 0, 0);
}

/** 계획 구간 끝 기본값 — 계획월 말일 근무일의 끝(= 다음 달 1일 08:00). 계획월이 없으면 planStart 기준 월 */
function defaultHorizonEnd_(data, planStart) {
  const y = data.planMonth ? data.planMonth.year : planStart.getFullYear();
  const m = data.planMonth ? data.planMonth.month - 1 : planStart.getMonth();
  return new Date(y, m + 1, 1, CFG.PLAN.START_HOUR, 0, 0, 0);
}

/** 출하일의 납기 시각 — 그 근무일의 끝(출하일+1 08:00). 당일 생산분은 당일 출하에 쓸 수 있다(일 마감 기준) */
function deadlineOf_(dueDate) {
  return nextDayStart_(new Date(dueDate.getFullYear(), dueDate.getMonth(), dueDate.getDate(), CFG.PLAN.START_HOUR, 0, 0, 0));
}

/**
 * ① 수요 산출 (규칙 7)
 *  - 출하 건을 출하일 오름차순 → [출하우선순위] 고객사 순 → 품목코드 순으로 정렬
 *  - **같은 품목·같은 출하일은 고객사와 무관하게 하나로 묶어** 재고를 차감하고, 부족분만 생산
 *  - 생산량은 CAPA 중량(KG)(보빈 1개) 배수로 올림. 올림분은 재고로 남아 다음 출하 건에 쓰인다
 *  stock 객체는 호출 측 사본이며 여기서 갱신된다(끝나면 계획 구간 말 예상 재고)
 * @returns {Array<{id, item, dueDate, deadline, qty, bobbins, shortage, fromStock, customers, kind}>}
 */
function buildDemands_(data, stock, shipRows) {
  const prio = (c) => (data.shipPrio[c] !== undefined ? data.shipPrio[c] : 999);
  const rows = (shipRows || data.ship).slice().sort((a, b) =>
    a.출하일.getTime() - b.출하일.getTime() || prio(a.고객사) - prio(b.고객사) || (a.품목코드 < b.품목코드 ? -1 : a.품목코드 > b.품목코드 ? 1 : 0));

  const groups = {};
  const order = [];
  rows.forEach(r => {
    const k = r.품목코드 + '|' + dateKey_(r.출하일);
    if (!groups[k]) { groups[k] = { id: k, item: r.품목코드, dueDate: r.출하일, qty: 0, customers: [] }; order.push(k); }
    const g = groups[k];
    g.qty += r.출하량;
    if (g.customers.indexOf(r.고객사) < 0) g.customers.push(r.고객사);
  });
  order.sort((a, b) => groups[a].dueDate.getTime() - groups[b].dueDate.getTime() || (groups[a].item < groups[b].item ? -1 : groups[a].item > groups[b].item ? 1 : 0));

  const demands = [];
  order.forEach(k => {
    const g = groups[k];
    g.customers.sort((a, b) => prio(a) - prio(b) || a.localeCompare(b, 'ko'));
    const have = stock[g.item] || 0;
    if (have >= g.qty - 1e-9) { stock[g.item] = have - g.qty; return; }        // 재고로 충당 — 생산 불필요
    const shortage = g.qty - have;
    const capa = data.capa[g.item];
    if (!capa) {
      stock[g.item] = 0;
      demands.push({ id: k, item: g.item, dueDate: g.dueDate, deadline: deadlineOf_(g.dueDate), qty: shortage, bobbins: 0, shortage, fromStock: have, customers: g.customers, kind: '납기' });
      return;
    }
    const bobbins = Math.ceil(shortage / capa.bobbinKg - 1e-9);
    const produce = bobbins * capa.bobbinKg;
    stock[g.item] = produce - shortage;                                        // 올림분은 재고로
    demands.push({ id: k, item: g.item, dueDate: g.dueDate, deadline: deadlineOf_(g.dueDate), qty: produce, bobbins, shortage, fromStock: have, customers: g.customers, kind: '납기' });
  });
  return demands;
}

/**
 * ② 호기 선택 (규칙 1 · 9 · 8-1 · 8-2 · 2 · 3 · 4 · 5)
 *  후보 필터: 절대차단 제외 → 가닥수 ≥ 37이면 Payoff수 ≥ 2 호기만 → 기종(적용설비) 일치
 *  각 후보의 미리보기(규칙 5 반영)를 낸 뒤 아래 순서로 고른다:
 *   8-1 교체 없이 납기 안에 끝나는 호기가 있으면 그 중에서 (우선설비 → 종료 빠른 순)  ※ 납기 없는 보충 작업엔 미적용
 *   8-2 다른 품목이 걸려 있고 그 품목 수요가 남은 호기는 후순위(전담 예약). 남는 후보가 없으면 예약 해제
 *    2  우선설비 — 가장 빨리 비는 우선 호기의 대기가 PREF_WAIT_LIMIT_HR(48h) 이내면 선택
 *    3  동일규격 — 같은 품목이 걸린 호기의 대기가 SAME_WAIT_LIMIT_HR(168h) 이내면 선택
 *    4  그 외 — 가장 빨리 비는 호기 (동률: 종료 빠른 순 → 호기 번호 순)
 *  대기(h) = 그 호기의 시작 가능 시각 − 후보 중 가장 빠른 시작 가능 시각
 *  12-1 (보충 전용, opts.horizonEnd) 납기가 없는 보충 작업은 급하지 않으므로, 교체 없이 계획 구간 끝 안에
 *       1보빈 이상 만들 수 있는 호기가 있으면 대기 한도와 무관하게 그 호기를 쓴다(들어가는 보빈 수 많은 순 → 종료 빠른 순).
 *       규칙 8-1의 "교체는 손실" 취지를 보충에 확장한 것 — 전문가 교차검증(#7) 대상.
 * @returns {object|null} preview 결과 (+waitHr·pref·meets) 또는 후보가 없으면 null
 */
function chooseMachine_(sim, data, d, remaining, opts) {
  opts = opts || {};
  const capa = data.capa[d.item];
  const needPayoff2 = capa.가닥수 >= CFG.PLAN.PAYOFF2_MIN_STRANDS;
  const cands = data.machineList.filter(m => {
    if (data.blocked[d.item + '|' + m]) return false;                       // 규칙 1 절대차단
    const mc = data.machines[m];
    if (needPayoff2 && mc.payoff < 2) return false;                          // 규칙 9 Payoff
    if (capa.적용설비 && mc.적용설비 && capa.적용설비 !== mc.적용설비) return false;   // 기종 매칭
    return true;
  });
  if (!cands.length) return null;

  const prefList = data.pref[d.item] || [];
  const previews = cands.map(m => sim.preview(m, d.item, d.qty, { maxRun: data.maxRun[d.item] }));
  const earliest = Math.min.apply(null, previews.map(p => p.start.getTime()));
  previews.forEach(p => {
    p.waitHr = (p.start.getTime() - earliest) / 3600000;
    p.pref = prefList.indexOf(p.machine) >= 0;
    p.prefRank = p.pref ? prefList.indexOf(p.machine) : 999;
    p.meets = d.deadline ? p.end.getTime() <= d.deadline.getTime() : true;
    p.order = data.machineList.indexOf(p.machine);
  });
  const byStart = (a, b) => a.start.getTime() - b.start.getTime() || a.end.getTime() - b.end.getTime() || a.order - b.order;
  const byEnd = (a, b) => a.end.getTime() - b.end.getTime() || a.order - b.order;

  /* 8-1 규격교체 회피 */
  if (d.deadline) {
    const noChg = previews.filter(p => !p.changeover && p.meets);
    if (noChg.length) {
      const prefFirst = noChg.filter(p => p.pref);
      return (prefFirst.length ? prefFirst : noChg).sort(byEnd)[0];
    }
  }

  /* 12-1 보충 작업: 교체 없이 구간 끝 안에 만들 수 있는 호기 우선 */
  if (!d.deadline && opts.horizonEnd) {
    previews.forEach(p => {
      const availHr = workHoursBetween_(p.start, opts.horizonEnd, data.holidays) - p.setupHr;
      p.fitBobbins = Math.max(0, Math.floor(availHr * capa.kgPerHr / capa.bobbinKg + 1e-9));
    });
    const noChgFit = previews.filter(p => !p.changeover && p.fitBobbins > 0);
    if (noChgFit.length) return noChgFit.sort((a, b) => b.fitBobbins - a.fitBobbins || byEnd(a, b))[0];
  }

  /* 8-1-b 납기 우선 (규칙 8-1의 취지 확장 — 전문가 교차검증 반영)
   * 무교체로 납기를 맞출 수 있는 호기가 없을 때는, 교체를 감수하더라도 납기를 지키는 후보만 남긴다.
   * 규칙 8-1이 "교체 없이 **납기 안에** 끝낼 수 있으면 그 호기를 쓴다"이므로, 납기를 못 지키는 무교체 호기보다
   * 납기를 지키는 교체 호기가 낫다. 납기를 지킬 후보가 하나도 없으면(능력 부족) 전체 후보를 그대로 쓴다. */
  let scope = previews;
  if (d.deadline) {
    const meets = previews.filter(p => p.meets);
    if (meets.length) scope = meets;
  }

  /* 8-2 전담 호기 예약 */
  let pool = scope.filter(p => !(p.changeover && (remaining[sim.currentItemOf(p.machine)] || 0) > 0));
  if (!pool.length) pool = scope;                                             // 쓸 호기가 그것뿐이면 쓴다
  else if (pool.length < scope.length) {
    const minWait = Math.min.apply(null, pool.map(p => p.waitHr));
    if (minWait > CFG.PLAN.DEDICATED_WAIT_LIMIT_HR) pool = scope;             // 기다림이 한도를 넘으면 예약 해제
  }

  /* 2 우선설비 */
  const prefs = pool.filter(p => p.pref).sort((a, b) => a.prefRank - b.prefRank || byStart(a, b));
  if (prefs.length && prefs[0].waitHr <= CFG.PLAN.PREF_WAIT_LIMIT_HR) return prefs[0];

  /* 3 동일규격 */
  const same = pool.filter(p => !p.changeover).sort(byStart);
  if (same.length && same[0].waitHr <= CFG.PLAN.SAME_WAIT_LIMIT_HR) return same[0];

  /* 4 최속 가용 */
  return pool.slice().sort(byStart)[0];
}

/**
 * ③ 적정재고 보충 (규칙 12)
 *  납기 수요 배정이 모두 끝난 뒤, 계획 구간 말 예상 재고 < [적정재고]인 품목만 부족분을 보빈 배수로 올림해 배정한다.
 *  기존 작업 뒤에 붙이므로 납기 작업을 밀지 않는다. 계획 구간 끝(horizonEnd) 안에 들어가는 보빈 수만 배정하고,
 *  못 채운 양은 result.replenish에 미달로 기록만 한다. [적정재고]에 없는 품목은 대상이 아니다.
 */
function replenishTargetStock_(sim, data, stock, horizonEnd, result) {
  Object.keys(data.targetStock).sort().forEach(item => {
    const target = data.targetStock[item];
    const have = stock[item] || 0;
    if (have >= target - 1e-9) return;
    const capa = data.capa[item];
    const rec = { item, target, before: have, need: target - have, produced: 0, after: have, machine: '', status: '' };
    result.replenish.push(rec);
    if (!capa) { rec.status = 'CAPA 없음'; return; }

    let bobbins = Math.ceil((target - have) / capa.bobbinKg - 1e-9);
    const d = { item, qty: bobbins * capa.bobbinKg, deadline: null, kind: '보충' };
    let pick = chooseMachine_(sim, data, d, {}, { horizonEnd });
    if (!pick) { rec.status = '배정 불가(가용 호기 없음)'; return; }

    if (pick.end.getTime() > horizonEnd.getTime()) {
      // 여유 시간 안에 들어가는 보빈 수로 줄인다
      const availHr = workHoursBetween_(pick.start, horizonEnd, data.holidays) - pick.setupHr;
      const fit = Math.floor(availHr * capa.kgPerHr / capa.bobbinKg + 1e-9);
      if (fit <= 0) { rec.status = `여유 시간 없음 — 미달 ${Math.round(target - have).toLocaleString()}kg`; return; }
      bobbins = Math.min(bobbins, fit);
      pick = sim.preview(pick.machine, item, bobbins * capa.bobbinKg, { maxRun: data.maxRun[item] });
    }
    const jobs = sim.commit(pick, { kind: '보충', demandId: 'replenish|' + item, customers: [] });
    jobs.forEach(j => { j._late = false; result.jobs.push(j); });
    rec.produced = bobbins * capa.bobbinKg;
    rec.after = have + rec.produced;
    rec.machine = pick.machine;
    stock[item] = rec.after;
    rec.status = rec.after >= target - 1e-9 ? '충족' : `부분 충족 — 미달 ${Math.round(target - rec.after).toLocaleString()}kg`;
  });
}

/**
 * 작업 배열 마무리 — 계획ID·규격 조인·호기 내 순번(시작일시 순)·정렬(호기 → 순번)
 */
function finalizeJobs_(jobs, data, planId) {
  const order = (m) => data.machineList.indexOf(m);
  jobs.sort((a, b) => order(a.호기) - order(b.호기) || a.시작일시.getTime() - b.시작일시.getTime());
  let seq = 0, lastMachine = null;
  jobs.forEach(j => {
    if (j.호기 !== lastMachine) { seq = 0; lastMachine = j.호기; }
    j.계획ID = planId;
    j.순번 = ++seq;
    j.규격 = data.capa[j.품목코드] ? data.capa[j.품목코드].규격 : '';
  });
  return jobs;
}
