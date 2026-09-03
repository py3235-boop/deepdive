/**
 * 미니 생산스케줄링 — 호기(설비)별로 어느 날 어떤 품목을 만드는지 시뮬레이션해서,
 * 품목별 "그 날짜에 실제로 생산된 kg"을 계산한다. ProductionCapa.gs는 이 결과를 누적해서
 * "그 날짜까지 생산 가능한 누적량" 상한을 만드는 데 쓴다.
 *
 * 기준정보.xlsx에서 이만큼을 반영한다:
 *  - 설비규격교체현황: 이번 달 1일 시작 시점에 각 호기가 이미 어떤 품목으로 세팅돼 있는지
 *    (세팅이 그대로면 교체 없이 바로 생산)
 *  - 안되는품목: 특정 호기가 절대 못 만드는 품목(그 호기는 배정 후보에서 제외)
 *  - 우선순위: 같은 호기를 여러 품목이 동시에 원할 때, 이 표에 먼저 나온 품목이 그 호기를 가져감
 *  - 최대설비가동수: 품목별로 동시에 쓸 수 있는 호기 대수 상한
 *  - CAPA: 호기 1대가 그 품목을 하루 종일 돌렸을 때 나오는 kg(kg/일/24hr)
 *
 * ⚠ 단순화한 부분(실제 데이터/운영 규칙이 없어서):
 *  - 교체시간 실측값이 없어서, 호기가 다른 품목으로 바뀌는 그날 하루는 생산량을 0으로 처리한다
 *    (실제 교체에 몇 시간 걸리는지에 따라 다를 텐데 일단 "교체일 = 하루 손실"로 고정).
 *  - 우선순위 탭은 순위 숫자가 없고 (품목코드,호기) 쌍만 나열돼 있어서, "그 호기 기준으로 먼저
 *    나온 행일수록 우선순위가 높다"고 해석했다.
 *  - 호기가 이미 다른(우선순위 낮은) 품목으로 차 있어도 필요하면 뺏어오는데(preempt), 이때도
 *    교체일 하루 손실이 적용된다.
 */

/**
 * year/month 기준으로 호기 배정을 시뮬레이션한다.
 * demandCodes: 이번 실행에서 실제로 계획을 짜야 하는 품목코드 Set(주문에 있는 코드만 — 관련 없는
 *   품목이 괜히 호기를 차지하지 않도록 범위를 좁힌다).
 * 반환: { code: { 'yyyy-MM-dd': 그날 생산된 kg } }
 */
function buildMachineSchedule_(ss, year, month, holidaySet, demandCodes) {
  const machines = _loadMachineRoster_(ss);
  const capaByCode = _loadCapaByCode_(ss);
  const maxUnitsByCode = _loadMaxUnitsByCode_(ss);
  const excluded = _loadExcludedPairs_(ss);
  const priorityByMachine = _loadPriorityByMachine_(ss);
  const currentAssignment = _loadInitialSetup_(ss); // machine -> code (그대로 두면 변형됨)

  const producedByCodeDate = {};
  const totalDays = daysInMonth_(year, month);
  const codes = Array.from(demandCodes).filter(code => capaByCode[code]); // capa 정보 없는 품목은 시뮬레이션 대상에서 제외

  for (let d = 1; d <= totalDays; d++) {
    const date = new Date(year, month - 1, d);
    const day = date.getDay();
    if (day === 0 || day === 6 || holidaySet.has(dateKey_(date))) continue; // 휴일엔 생산 안 함(단순화)
    const key = dateKey_(date);

    const changedToday = {}; // machineId -> true (오늘 교체됨 → 오늘 생산 0)

    const assignedCountByCode = {};
    machines.forEach(m => {
      const code = currentAssignment[m.id];
      if (code) assignedCountByCode[code] = (assignedCountByCode[code] || 0) + 1;
    });

    // 아직 최대가동수를 못 채운 품목은 빈 호기 → 뺏을 수 있는 호기 순으로 확보 시도
    codes.forEach(code => {
      const maxUnits = maxUnitsByCode[code] || 0;
      let have = assignedCountByCode[code] || 0;
      if (have >= maxUnits) return;

      const candidates = machines
        .filter(m => currentAssignment[m.id] !== code && !excluded.has(code + '|' + m.id))
        .sort((a, b) => _machineScore_(a, code, currentAssignment, priorityByMachine) -
                        _machineScore_(b, code, currentAssignment, priorityByMachine));

      for (let i = 0; i < candidates.length && have < maxUnits; i++) {
        const m = candidates[i];
        const occupant = currentAssignment[m.id];
        if (occupant && !_canPreempt_(code, occupant, m.id, priorityByMachine)) continue;
        currentAssignment[m.id] = code;
        changedToday[m.id] = true; // 교체일 — 오늘 생산 0
        have++;
      }
      assignedCountByCode[code] = have;
    });

    machines.forEach(m => {
      if (changedToday[m.id]) return;
      const code = currentAssignment[m.id];
      if (!code || !capaByCode[code]) return;
      producedByCodeDate[code] = producedByCodeDate[code] || {};
      producedByCodeDate[code][key] = (producedByCodeDate[code][key] || 0) + capaByCode[code].kgPerDay;
    });
  }

  return producedByCodeDate;
}

/** 낮을수록 그 호기를 이 품목에 먼저 배정한다: 빈 호기(0) > 우선순위 있는 자리 뺏기(10+순위) > 그 외(99). */
function _machineScore_(machine, code, currentAssignment, priorityByMachine) {
  const occupant = currentAssignment[machine.id];
  if (!occupant) return 0;
  const list = priorityByMachine[machine.id] || [];
  const myRank = list.indexOf(code);
  return myRank === -1 ? 99 : 10 + myRank;
}

/** code가 machineId를 지금 쓰고 있는 occupant한테서 뺏어올 수 있는지(우선순위표 기준). */
function _canPreempt_(code, occupant, machineId, priorityByMachine) {
  const list = priorityByMachine[machineId] || [];
  const myRank = list.indexOf(code);
  if (myRank === -1) return false; // 이 호기 우선순위표에 없으면 남의 자리 못 뺏음
  const occupantRank = list.indexOf(occupant);
  return occupantRank === -1 || myRank < occupantRank;
}

function _loadMachineRoster_(ss) {
  const sheet = ss.getSheetByName('집합설비현황');
  if (!sheet) return [];
  const values = sheet.getDataRange().getValues();
  const header = values[0].map(h => String(h).trim());
  const idIdx = header.indexOf('호기');
  const equipIdx = header.indexOf('적용설비');
  return values.slice(1).filter(r => r[idIdx]).map(r => ({
    id: String(r[idIdx]).trim(),
    equipType: equipIdx !== -1 ? String(r[equipIdx]).trim() : '',
  }));
}

function _loadCapaByCode_(ss) {
  const sheet = ss.getSheetByName('CAPA');
  if (!sheet) return {};
  const values = sheet.getDataRange().getValues();
  const header = values[0].map(h => String(h).trim());
  const codeIdx = header.indexOf('품목코드');
  const equipIdx = header.indexOf('적용설비');
  const kgDayIdx = header.indexOf('kg/일/24hr');
  const map = {};
  values.slice(1).forEach(r => {
    if (!r[codeIdx]) return;
    map[_normCode_(r[codeIdx])] = {
      equipType: equipIdx !== -1 ? String(r[equipIdx]).trim() : '',
      kgPerDay: Number(r[kgDayIdx]) || 0,
    };
  });
  return map;
}

function _loadMaxUnitsByCode_(ss) {
  const sheet = ss.getSheetByName('최대설비가동수');
  if (!sheet) return {};
  const values = sheet.getDataRange().getValues();
  const header = values[0].map(h => String(h).trim());
  const codeIdx = header.indexOf('품목코드');
  const unitsIdx = header.indexOf('최대 가동수');
  const map = {};
  values.slice(1).forEach(r => {
    if (!r[codeIdx]) return;
    map[_normCode_(r[codeIdx])] = Number(r[unitsIdx]) || 0;
  });
  return map;
}

/** "품목코드|호기" 조합 Set — 이 조합이면 그 호기는 그 품목을 못 만든다. */
function _loadExcludedPairs_(ss) {
  const sheet = ss.getSheetByName('안되는품목');
  const set = new Set();
  if (!sheet) return set;
  const values = sheet.getDataRange().getValues();
  values.slice(1).forEach(r => {
    if (!r[0] || !r[1]) return;
    set.add(_normCode_(r[0]) + '|' + String(r[1]).trim());
  });
  return set;
}

/** machine -> [품목코드...] 순서대로(먼저 나올수록 그 호기에서 우선순위 높음). */
function _loadPriorityByMachine_(ss) {
  const sheet = ss.getSheetByName('우선순위');
  const map = {};
  if (!sheet) return map;
  const values = sheet.getDataRange().getValues();
  values.slice(1).forEach(r => {
    if (!r[0] || !r[1]) return;
    const machine = String(r[1]).trim();
    map[machine] = map[machine] || [];
    map[machine].push(_normCode_(r[0]));
  });
  return map;
}

/** machine -> 품목코드 (이번 달 시작 시점 세팅). */
function _loadInitialSetup_(ss) {
  const sheet = ss.getSheetByName('설비규격교체현황');
  const map = {};
  if (!sheet) return map;
  const values = sheet.getDataRange().getValues();
  values.slice(1).forEach(r => {
    if (!r[0] || !r[1]) return;
    map[String(r[0]).trim()] = _normCode_(r[1]);
  });
  return map;
}
