/**
 * 기준정보.xlsx 기반 품목코드별 생산capa 정보.
 *
 * 예전엔 "CAPA × 최대설비가동수"로 하루 상한을 계산하는 단순 공식이었는데, 지금은 MachineSchedule.gs가
 * 호기별로 실제 어느 날 어떤 품목을 만드는지 미니 시뮬레이션을 돌려서 그 결과(품목별 날짜별 생산량)를
 * 여기서 누적해서 쓴다 — 설비규격교체현황/안되는품목/우선순위까지 반영된 값이다.
 *
 * 기초재고 탭은 같은 품목코드가 여러 행(422, 422, 419, 419, 417...)으로 반복되는 걸 확인했는데,
 * 날짜 컬럼이 없어서 정확한 의미(이력인지/오탈자인지)를 알 수 없다 — 일단 첫 번째로 나오는 값만
 * 기초재고로 쓴다. 실제 의미가 확인되면 이 부분을 고쳐야 한다.
 */
function loadProductionCapaInfo_(year, month, holidaySet, demandCodes) {
  const info = {}; // code -> { baseStock, targetStock, producedByDate }
  if (!CONFIG.REFERENCE_FOLDER_ID) return info;

  try {
    const ss = getReferenceSpreadsheet_('기준정보');

    const baseStockSheet = ss.getSheetByName('기초재고');
    if (baseStockSheet) {
      const values = baseStockSheet.getDataRange().getValues();
      const header = values[0].map(h => String(h).trim());
      const codeIdx = header.indexOf('품목코드');
      const stockIdx = header.indexOf('기초재고');
      values.slice(1).forEach(r => {
        if (!r[codeIdx]) return;
        const code = _normCode_(r[codeIdx]);
        if (info[code] && info[code].baseStock !== undefined) return; // 첫 값만 사용
        info[code] = info[code] || {};
        info[code].baseStock = Number(r[stockIdx]) || 0;
      });
    }

    const targetStockSheet = ss.getSheetByName('적정재고');
    if (targetStockSheet) {
      const values = targetStockSheet.getDataRange().getValues();
      const header = values[0].map(h => String(h).trim());
      const codeIdx = header.indexOf('품목코드');
      const stockIdx = header.indexOf('적정 재고량');
      values.slice(1).forEach(r => {
        if (!r[codeIdx]) return;
        const code = _normCode_(r[codeIdx]);
        info[code] = info[code] || {};
        info[code].targetStock = Number(r[stockIdx]) || 0;
      });
    }

    // 호기별 미니 생산스케줄링 — year/month/holidaySet/demandCodes가 없으면(옛 호출부 호환) 스킵
    if (year && month && demandCodes) {
      const produced = buildMachineSchedule_(ss, year, month, holidaySet || new Set(), demandCodes);
      Object.keys(produced).forEach(code => {
        info[code] = info[code] || {};
        info[code].producedByDate = produced[code];
      });
    }
  } catch (e) {
    Logger.log('기준정보.xlsx 생산capa 로드 실패(검증 없이 진행): ' + e.message);
    return {};
  }

  return info;
}

function _normCode_(v) {
  return String(v).trim();
}

/**
 * 품목코드의 "그 날짜까지 누적 생산 가능량" 계산 함수를 만든다.
 *   cum = 기초재고 + (계획월 1일부터 그 날짜까지 MachineSchedule.gs가 시뮬레이션한 실제 생산량 합) - 적정재고
 *   0 밑으로는 안 내려감.
 * 기초재고/적정재고 정보 자체가 없는 품목(기준정보에 아예 없음)만 검증을 스킵(무제한 취급)한다.
 * producedByDate가 없는 품목(=설비 경쟁에서 밀려 이번 달 호기를 하나도 못 받음)은 무제한이 아니라
 * **생산량 0으로 취급**한다 — "생산을 못 받았으니 검증 안 함"이 아니라 "생산을 못 받았으니 기초재고
 * 넘는 만큼만(대개 0)"이 맞는 방향이다. 이걸 반대로 하면 설비를 하나도 못 받은, 오히려 가장 위험한
 * 품목이 검증에서 완전히 빠지는 결과가 된다.
 * ⚠ AS_OF_DATE는 쓰지 않는다 — 실적 반영이 없는 순수 계획 단계라서 계획월 1일부터 근무일마다 그대로
 *   쌓인다고 본다(실적 반영 단계에서만 AS_OF_DATE 개념이 의미 있어짐).
 */
function buildCapaCumulativeFn_(capaInfo, code, year, month, holidaySet) {
  const entry = capaInfo[code];
  // 기초재고/적정재고 둘 다 아예 없으면(기준정보에 이 품목 자체가 없음) 검증 불가 — 스킵.
  if (!entry || (entry.baseStock === undefined && entry.targetStock === undefined)) return null;

  const baseStock = entry.baseStock || 0;
  const targetStock = entry.targetStock || 0;
  const produced = entry.producedByDate || {}; // 없으면 생산량 0으로 취급(위 설명 참고)
  const monthStart = new Date(year, month - 1, 1);

  return function (date) {
    let cumProduced = 0;
    for (let d = new Date(monthStart); d <= date; d.setDate(d.getDate() + 1)) {
      cumProduced += produced[dateKey_(d)] || 0;
    }
    const cum = baseStock + cumProduced - targetStock;
    return Math.max(0, cum);
  };
}
