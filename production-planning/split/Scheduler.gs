/**
 * ============================================================================
 *  집합공정 생산계획 자동화 — Scheduler.gs (시간 시뮬레이션)
 * ----------------------------------------------------------------------------
 *  역할 : 호기별 점유 시간 시뮬레이션. Planner가 "이 호기에 이 작업을 넣으면 언제 끝나나"를 물으면(preview)
 *         답하고, 확정(commit)하면 호기 상태를 갱신한다. 시간 규칙은 Utils.gs 근무시간 모델(addWorkHours_ 등)만 쓴다.
 *
 *  시간 규칙 (CLAUDE.md 계획 로직 규칙 6 · 8 · 8-3)
 *   - 계획 첫날(planStart가 속한 근무일)에 시작하는 호기의 첫 작업 앞에 가동준비 INITIAL_READY_HOURS(3h)
 *   - 직전에 걸린 품목과 다르면 작업 앞에 규격교체 CHANGE_HOURS(1.5h). 같은 품목 연속이면 0
 *     (호기의 시작 품목은 [설비규격교체현황])
 *   - 가동시간 = 생산량 ÷ kg/hr (호기 kg/hr = 그 호기 적용설비 기준 CAPA — 현 구성은 전 호기 동일 기종)
 *   - 연속 점유가 MAX_CHUNK_HOURS(72h)를 넘으면 **보빈 단위**로 청크 분할: 72h 안에 들어가는 보빈 수만 담고
 *     나머지는 다음 청크로. 청크는 같은 호기에서 연속 진행(같은 품목이라 청크 사이 교체 없음)
 *   - 주말·[휴무]는 건너뛰고 다음 근무일 08:00부터 이어서 (addWorkHours_)
 *   - 같은 호기의 작업은 시간이 겹치지 않는다 (availableAt 순차 진행)
 *   - 규칙 5: 같은 품목이 동시에 가동되는 호기 수가 [최대설비가동수]를 넘지 않게 시작시각을 뒤로 민다
 *  ⚠ 이 파일에는 시트 I/O·서식 코드가 없다.
 * ============================================================================
 */

/**
 * 시뮬레이션 객체 생성
 * @param {object} data   loadData_() 결과 (capa · machines · machineList · currentItem · holidays)
 * @param {Date}   planStart 계획 시작 시각 (근무일 08:00으로 정렬해서 넘길 것)
 */
function createSimulation_(data, planStart, opts) {
  opts = opts || {};
  const holidays = data.holidays || {};
  const firstDayKey = dateKey_(dayStartOf_(planStart));
  /* 실적 반영(규칙 11) 시 호기별 시작 조건을 덮어쓴다:
   *  machineAvailable[호기] = 기준일 당일 실적 이후 잔여 시간의 시작 시각
   *  currentItem[호기]      = 실적의 마지막 품목 (교체시간 판단용)
   *  warm[호기]             = true면 이미 가동 중이므로 가동준비 3h를 붙이지 않는다 */
  const mAvail = opts.machineAvailable || {};
  const mItem = opts.currentItem || {};
  const warm = opts.warm || {};
  const machines = {};
  data.machineList.forEach(m => {
    const start = mAvail[m] ? new Date(mAvail[m].getTime()) : new Date(planStart.getTime());
    const item = (mItem[m] !== undefined ? mItem[m] : data.currentItem[m]) || null;
    machines[m] = { name: m, availableAt: start, currentItem: item, initialItem: item, jobs: [], warm: !!warm[m] };
  });
  const allJobs = [];   // 규칙 5(동시 가동 수) 검사용 — {item, machine, start, end}

  /** 구간 [s, e) 안에서 다른 호기가 같은 품목을 동시에 돌리는 최대 개수 */
  function maxConcurrentOthers_(item, machineName, s, e) {
    const others = allJobs.filter(j => j.item === item && j.machine !== machineName && j.start.getTime() < e.getTime() && j.end.getTime() > s.getTime());
    if (!others.length) return { count: 0, others };
    const events = [];
    others.forEach(j => { events.push([Math.max(j.start.getTime(), s.getTime()), 1]); events.push([Math.min(j.end.getTime(), e.getTime()), -1]); });
    events.sort((a, b) => a[0] - b[0] || a[1] - b[1]);   // 같은 시각이면 종료(-1) 먼저
    let cur = 0, peak = 0;
    events.forEach(ev => { cur += ev[1]; if (cur > peak) peak = cur; });
    return { count: peak, others };
  }

  /** 규칙 5 — 동시 가동 수가 한도에 걸리면 겹치는 작업이 끝나는 시각으로 시작을 미룬다 */
  function adjustForMaxRun_(item, machineName, start, occupyHr, maxRun) {
    let s = start;
    let guard = 0;
    while (guard++ < 300) {
      const e = addWorkHours_(s, occupyHr, holidays);
      const r = maxConcurrentOthers_(item, machineName, s, e);
      if (r.count < maxRun) return s;
      const nextFree = Math.min.apply(null, r.others.filter(j => j.end.getTime() > s.getTime()).map(j => j.end.getTime()));
      if (!isFinite(nextFree)) return s;
      s = alignToWorkTime_(new Date(nextFree), holidays);
    }
    return s;
  }

  /**
   * 미리보기 — 호기 상태를 바꾸지 않고 "이 작업을 넣으면 어떻게 되나"를 계산한다
   * @returns {{machine, item, qty, start, end, readyHr, changeHr, runHr, setupHr, changeover, timeline:[{idx, of, start, prodStart, end, qty, bobbins, readyHr, changeHr, runHr}]}}
   */
  function preview(machineName, item, qty, opt) {
    opt = opt || {};
    const ms = machines[machineName];
    const capa = data.capa[item];
    if (!ms) throw new Error(`시뮬레이션: 호기 ${machineName} 없음`);
    if (!capa) throw new Error(`시뮬레이션: ${item} CAPA 없음`);
    const kgPerHr = capa.kgPerHr;

    const readyHr = (!ms.warm && ms.jobs.length === 0 && dateKey_(dayStartOf_(alignToWorkTime_(ms.availableAt, holidays))) === firstDayKey) ? CFG.PLAN.INITIAL_READY_HOURS : 0;
    const changeHr = (ms.currentItem && ms.currentItem !== item) ? CFG.PLAN.CHANGE_HOURS : 0;

    // 청크 분할 (규칙 8-3): 72h 안에 들어가는 보빈 수 (최소 1보빈)
    const bobbinHr = capa.bobbinKg / kgPerHr;
    const perChunk = Math.max(1, Math.floor(CFG.PLAN.MAX_CHUNK_HOURS / bobbinHr + 1e-9));
    const totalBobbins = Math.max(1, Math.round(qty / capa.bobbinKg));
    const chunks = [];
    for (let left = totalBobbins; left > 0; left -= perChunk) chunks.push(Math.min(perChunk, left));

    const totalRunHr = qty / kgPerHr;
    let start = alignToWorkTime_(ms.availableAt, holidays);
    if (opt.maxRun) start = adjustForMaxRun_(item, machineName, start, readyHr + changeHr + totalRunHr, opt.maxRun);

    const timeline = [];
    let cur = start;
    chunks.forEach((b, i) => {
      const runHr = b * capa.bobbinKg / kgPerHr;
      const setup = i === 0 ? readyHr + changeHr : 0;
      const prodStart = addWorkHours_(cur, setup, holidays);
      const end = addWorkHours_(prodStart, runHr, holidays);
      timeline.push({ idx: i + 1, of: chunks.length, start: cur, prodStart, end, qty: b * capa.bobbinKg, bobbins: b, readyHr: i === 0 ? readyHr : 0, changeHr: i === 0 ? changeHr : 0, runHr });
      cur = end;
    });
    return { machine: machineName, item, qty, start, end: cur, readyHr, changeHr, runHr: totalRunHr, setupHr: readyHr + changeHr, changeover: changeHr > 0, timeline };
  }

  /**
   * 확정 — 미리보기 결과를 호기에 배정하고 청크별 작업 객체를 돌려준다
   * @param {object} pv preview() 결과
   * @param {{kind:string, demandId?:string, dueDate?:Date, deadline?:Date, customers?:string[]}} meta
   */
  function commit(pv, meta) {
    meta = meta || {};
    const ms = machines[pv.machine];
    const jobs = pv.timeline.map(t => ({
      호기: pv.machine,
      품목코드: pv.item,
      고객사: (meta.customers || []).join('·'),
      출하일: meta.dueDate || null,
      시작일시: t.start,
      종료일시: t.end,
      '소요(h)': Math.round((t.runHr + t.changeHr) * 100) / 100,     // 데이터 모델: 생산량 ÷ kg/hr (+교체 1.5h)
      '생산량(kg)': t.qty,
      교체: t.changeHr > 0 ? 'Y' : 'N',
      청크: t.of > 1 ? `${t.idx}/${t.of}` : '',
      상태: '계획',
      _kind: meta.kind || '납기',
      _demandId: meta.demandId || '',
      _deadline: meta.deadline || null,
      _prodStart: t.prodStart,
      _readyHr: t.readyHr,
      _changeHr: t.changeHr,
      _runHr: t.runHr,
      _bobbins: t.bobbins,
    }));
    jobs.forEach(j => { ms.jobs.push(j); allJobs.push({ item: pv.item, machine: pv.machine, start: j.시작일시, end: j.종료일시 }); });
    ms.availableAt = new Date(pv.end.getTime());
    ms.currentItem = pv.item;
    return jobs;
  }

  return {
    machines,
    holidays,
    planStart,
    preview,
    commit,
    currentItemOf: (m) => machines[m] ? machines[m].currentItem : null,
    availableAtOf: (m) => machines[m] ? machines[m].availableAt : null,
    jobsOf: (m) => machines[m] ? machines[m].jobs : [],
    allJobs: () => allJobs.slice(),
  };
}

/**
 * 시간 시뮬레이션 단독 실행 (명세에 정한 함수 이름). Planner가 만든 수요 목록을 순서대로 넣어 결과만 계산할 때 사용.
 * @param {{data:object, planStart:Date, assignments:Array<{machine:string,item:string,qty:number,meta?:object}>}} params
 * @returns {{jobs:Array, sim:object}}
 */
function detailedSchedulingSimulation_(params) {
  const sim = createSimulation_(params.data, params.planStart);
  const jobs = [];
  (params.assignments || []).forEach(a => {
    const pv = sim.preview(a.machine, a.item, a.qty, { maxRun: params.data.maxRun[a.item] });
    sim.commit(pv, a.meta || {}).forEach(j => jobs.push(j));
  });
  return { jobs, sim };
}
