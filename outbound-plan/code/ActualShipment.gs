/**
 * 실적 반영(2단계) — 74340 실적 파일(REFERENCE_FOLDER_ID 폴더 안)을 읽어서 마지막 계획(LAST_PLAN)의
 * 과거 날짜 칸을 실적값으로 덮어쓰고, 남은 물량(발주량 - 실적)만 "오늘(=실적 파일의 최신 출고일) 이후"
 * 날짜에 다시 배분한다. 74340 파일만은 헤더 텍스트가 아니라 고정 컬럼 위치로 읽는다(더미 생성기와
 * 컬럼 순서를 맞춰뒀다는 전제).
 *
 * ⚠ AS_OF_DATE는 고정값이 아니라 "이 실적 파일에 실제로 적힌 가장 최근 출고일자"로 판단한다 —
 * 유효한 실적 행이 하나도 없으면 반영할 게 없다는 뜻이므로 아무 것도 안 하고 끝낸다.
 */

// 74340 파일 고정 컬럼 인덱스(0-based, getValues() 배열 기준) — 헤더 자동인식 아님
const ACTUAL_COL_ = {
  code: 1, // B열: 품번(scCode)
  qty: 6, // G열: 수량(kg)
  bobbin: 7, // H열: 보빈(kg) — 0/빈칸이면 정상 출고 기록이 아니라고 보고 제외
  shipDate: 8, // I열: 출고일자
  customer: 10, // K열: 납품처 ("고객사A" 형태 → 마지막 알파벳만 추출해 "고객사A"로 정규화)
};

function applyActualShipment() {
  try {
    _applyActualShipment_();
  } catch (e) {
    appendExecutionLog_('applyActualShipment', '실패', e.message);
    notifyChat_('🚨 실적 반영 실패: ' + e.message);
    throw e;
  }
}

function _applyActualShipment_() {
  const raw = PropertiesService.getDocumentProperties().getProperty('LAST_PLAN');
  if (!raw) {
    throw new Error('저장된 계획이 없습니다. 먼저 "② 계획 생성"을 실행하세요.');
  }
  const saved = JSON.parse(raw);
  const year = saved.year;
  const month = saved.month;
  const holidaySet = new Set(saved.holidayDates || []);

  const actual = _loadActualShipments_();
  if (!actual) {
    appendExecutionLog_('applyActualShipment', '스킵', '실적(74340) 파일을 못 찾았거나 유효한 실적 행이 없음');
    notifyChat_('ℹ️ 실적 반영: 실적(74340) 파일이 없거나 비어있어서 아무 것도 반영하지 않았습니다.');
    SpreadsheetApp.getActiveSpreadsheet().toast('실적(74340) 파일이 없거나 비어있어서 아무 것도 반영하지 않았습니다.', '안내', 8);
    return;
  }

  const asOfDate = actual.asOfDate;
  const asOfKey = dateKey_(asOfDate);
  const minDate = new Date(asOfDate);
  minDate.setDate(minDate.getDate() + 1); // "오늘(asOfDate) 이후"만 재배분 대상

  const previousSnapshot = loadPreviousSnapshot_(); // vendor|code|spec -> 발주량(기준값)
  const demandCodes = new Set(saved.rows.map(r => r.code)); // 재배분 대상이 될 수 있는 품목만 시뮬레이션
  const capaInfo = loadProductionCapaInfo_(year, month, holidaySet, demandCodes);
  const weights = loadUnitWeights_();
  const codeDateLoad = {};
  const issues = [];

  const updatedRows = [];
  const remainingByVendor = {}; // vendor -> [{code, spec, qty, __row}]

  saved.rows.forEach(row => {
    const vendorCodeKey = row.vendor + '|' + row.code;
    const snapshotKey = row.vendor + '|' + row.code + '|' + row.spec;
    const plannedQty = previousSnapshot[snapshotKey];
    const actualEntry = actual.byVendorCode[vendorCodeKey];

    // 과거(asOfDate 포함) 날짜는 전부 실적값으로 덮어씀(실적 파일에 없으면 그 날짜는 출고 안 된 것 = 0/빈칸)
    const newDateMap = {};
    const totalDays = daysInMonth_(year, month);
    for (let d = 1; d <= totalDays; d++) {
      const date = new Date(year, month - 1, d);
      const key = dateKey_(date);
      if (key > asOfKey) continue;
      const qty = actualEntry ? actualEntry.byDate[key] : undefined;
      if (qty) newDateMap[key] = qty;
    }

    const actualTotal = actualEntry ? actualEntry.total : 0;
    let remaining = 0;
    if (plannedQty === undefined) {
      issues.push(row.vendor + ' ' + row.code + ': 발주스냅샷에서 원래 발주량을 못 찾아 잔여물량 계산 불가(0으로 처리)');
    } else {
      remaining = Math.max(0, plannedQty - actualTotal);
    }

    const orderedQty = plannedQty !== undefined ? plannedQty : row.orderedQty;
    updatedRows.push({ code: row.code, spec: row.spec, vendor: row.vendor, orderedQty: orderedQty, dateMap: newDateMap, changed: false });

    if (remaining > 0.0001) {
      remainingByVendor[row.vendor] = remainingByVendor[row.vendor] || [];
      remainingByVendor[row.vendor].push({ code: row.code, spec: row.spec, qty: remaining, __row: updatedRows.length - 1 });
    }
  });

  // 남은 물량을 업체 타입에 맞는 알고리즘으로 "오늘 이후" 날짜에만 재배분
  const vendorDateLoad = {}; // 전사 날짜별 부하 — 공휴일 시프트용, 업체 간 공유(generatePlan과 같은 패턴)

  // 요일 고정형(균등분배)을 먼저 처리하고 트럭버킷(여유일을 유동적으로 고름)을 나중에 처리한다 —
  // generatePlan과 같은 이유(Main.gs 주석 참고): 순서를 반대로 하면 트럭버킷 쪽이 고정형의 미래
  // 부하를 모른 채 같은 날에 겹쳐 배정할 수 있다.
  const vendorOrder = Object.keys(VENDOR_TYPE_MAP).sort((a, b) => {
    const aFixed = BUCKET_WEEKDAYS[VENDOR_TYPE_MAP[a]] === undefined;
    const bFixed = BUCKET_WEEKDAYS[VENDOR_TYPE_MAP[b]] === undefined;
    return aFixed === bFixed ? 0 : (aFixed ? -1 : 1);
  });

  vendorOrder.forEach(vendor => {
    const items = remainingByVendor[vendor];
    if (!items || items.length === 0) return;

    const type = VENDOR_TYPE_MAP[vendor];
    const bucketWeekday = BUCKET_WEEKDAYS[type];
    let planByItem = {};

    if (bucketWeekday !== undefined) {
      // 트럭버킷 타입은 buildTruckBucketPlan_ 내부에서 이미 vendorDateLoad를 트럭대수로 갱신하므로
      // 아래서 또 합산하면 중복 반영됨 — 안 함.
      const result = buildTruckBucketPlan_(items, year, month, bucketWeekday, holidaySet, capaInfo, codeDateLoad, weights, minDate, vendorDateLoad);
      planByItem = result.plan;
      result.issues.forEach(i => issues.push(vendor + ' ' + i.code + ': ' + i.message));
    } else {
      // 트럭버킷이 아닌 타입은 이 업체의 그 날짜 kg 총합을 먼저 구한 뒤 한 번만 트럭대수로 올려서
      // vendorDateLoad에 반영한다(품목별로 각각 올려서 더하면 실제보다 부풀려짐 — Main.gs와 동일 이유).
      const dateTotals = {};
      items.forEach(it => {
        const dateMap = allocateEvenOnWeekday_([{ qty: it.qty }], year, month, 5, holidaySet, minDate, vendorDateLoad);
        planByItem[it.code + '|' + it.spec] = dateMap;
        Object.keys(dateMap).forEach(key => {
          dateTotals[key] = (dateTotals[key] || 0) + dateMap[key];
        });
      });
      Object.keys(dateTotals).forEach(key => {
        const trucks = Math.ceil(dateTotals[key] / CONFIG.TRUCK_KG);
        vendorDateLoad[key] = (vendorDateLoad[key] || 0) + trucks;
      });
    }

    items.forEach(it => {
      const targetRow = updatedRows[it.__row];
      const futureDateMap = planByItem[it.code + '|' + it.spec] || {};
      Object.keys(futureDateMap).forEach(key => {
        targetRow.dateMap[key] = (targetRow.dateMap[key] || 0) + futureDateMap[key];
      });
      if (Object.keys(futureDateMap).length === 0) {
        issues.push(vendor + ' ' + it.code + ': 재배분 결과가 비어있음 — ' + Math.round(it.qty) + 'kg 미배정 가능성');
      }
    });
  });

  writePlanSheet(year, month, updatedRows, {
    holidaySet: holidaySet,
    issueCount: issues.length,
    viewMode: 'item',
    actualCutoffDateKey: asOfKey,
  });

  PropertiesService.getDocumentProperties().setProperty('LAST_PLAN', JSON.stringify({
    year: year,
    month: month,
    rows: updatedRows,
    issueCount: issues.length,
    holidayDates: Array.from(holidaySet),
    actualCutoffDateKey: asOfKey,
  }));

  appendExecutionLog_(
    'applyActualShipment',
    issues.length > 0 ? '경고' : '완료',
    (issues.length > 0 ? issues.join('\n') : '이슈 없음') + '\n(기준일: ' + asOfKey + ')'
  );

  notifyChat_(
    (issues.length > 0 ? '⚠️ ' : '📦 ') +
    '실적 반영 완료(기준일: ' + asOfKey + ')\n' +
    (issues.length > 0 ? '검증 경고 ' + issues.length + '건 (실행이력 탭 참고)' : '이슈 없음') +
    '\n' + SpreadsheetApp.getActiveSpreadsheet().getUrl()
  );

  SpreadsheetApp.getActiveSpreadsheet().toast('실적 반영 완료 (기준일: ' + asOfKey + ')', '완료', 8);
}

/**
 * 74340 실적 파일을 읽어서 { asOfDate, byVendorCode: { 'vendor|code': {byDate:{}, total} } } 반환.
 * 파일이 없거나 유효한 실적 행이 하나도 없으면 null.
 */
function _loadActualShipments_() {
  let sheet;
  try {
    sheet = getReferenceSpreadsheet_('74340').getSheets()[0];
  } catch (e) {
    return null;
  }

  const values = sheet.getDataRange().getValues();
  const byVendorCode = {};
  let latestDate = null;

  values.slice(1).forEach(row => {
    const bobbin = row[ACTUAL_COL_.bobbin];
    if (!bobbin) return; // 보빈 0/빈칸이면 정상 출고 기록이 아님

    const codeCell = row[ACTUAL_COL_.code];
    const qty = Number(row[ACTUAL_COL_.qty]) || 0;
    if (!codeCell || !qty) return;

    const dateCell = row[ACTUAL_COL_.shipDate];
    const shipDate = dateCell instanceof Date ? dateCell : new Date(dateCell);
    if (isNaN(shipDate.getTime())) return;

    const customerCell = String(row[ACTUAL_COL_.customer] || '');
    const m = customerCell.match(/([A-Za-z])\s*$/);
    if (!m) return;
    const vendor = '고객사' + m[1].toUpperCase();

    const key = vendor + '|' + String(codeCell).trim();
    byVendorCode[key] = byVendorCode[key] || { byDate: {}, total: 0 };
    const dKey = dateKey_(shipDate);
    byVendorCode[key].byDate[dKey] = (byVendorCode[key].byDate[dKey] || 0) + qty;
    byVendorCode[key].total += qty;

    if (!latestDate || shipDate > latestDate) latestDate = shipDate;
  });

  if (!latestDate) return null;
  return { asOfDate: latestDate, byVendorCode: byVendorCode };
}
