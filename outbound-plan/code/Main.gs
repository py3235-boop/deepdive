/**
 * 메인 실행. year/month를 안 주면 오늘 기준 월로 계획을 만든다.
 */
function generatePlan(year, month) {
  try {
    _generatePlan_(year, month);
  } catch (e) {
    appendExecutionLog_('generatePlan', '실패', e.message);
    notifyChat_('🚨 출고계획 생성 실패: ' + e.message);
    throw e; // 편집기/시트에 원래 에러가 그대로 보이게 다시 던짐
  }
}

function _generatePlan_(year, month) {
  const today = new Date();
  year = year || today.getFullYear();
  month = month || today.getMonth() + 1;

  const orders = loadOrderStatus();
  const byVendor = _groupByVendor_(orders);
  const originalByVendor = Object.assign({}, byVendor); // 아래서 byVendor를 지워가며 순회하므로 스냅샷용 원본을 따로 보관

  // 발주지문 비교 — 지난 실행과 발주 내역이 완전히 같으면 재계산 자체를 스킵한다.
  const currentFingerprints = buildOrderFingerprints_(orders);
  const previousFingerprints = loadPreviousFingerprints_();
  if (fingerprintsUnchanged_(previousFingerprints, currentFingerprints)) {
    appendExecutionLog_('generatePlan', '스킵', '발주 데이터가 지난 실행과 동일 — 재계산 건너뜀');
    notifyChat_('ℹ️ 출고계획: 발주서가 지난 실행과 동일해서 재계산을 건너뛰었습니다.');
    SpreadsheetApp.getActiveSpreadsheet().toast(
      '발주서가 지난 실행과 똑같아서 계획을 다시 만들지 않았습니다. 강제로 다시 만들려면 forceRegeneratePlan()을 실행하세요.',
      '스킵',
      8
    );
    return;
  }

  // 품목별 변경 여부 확인용 — 지난 스냅샷과 이번 발주량을 비교해서 새로 생기거나 바뀐 품목을 표시한다.
  // 스냅샷 자체가 하나도 없으면(첫 실행) 비교 대상이 없다는 뜻이므로 전부 "안 바뀜"(검정)으로 둔다.
  const previousSnapshot = loadPreviousSnapshot_();
  const isFirstSnapshot = Object.keys(previousSnapshot).length === 0;
  const itemQtyMap = {};
  Object.keys(originalByVendor).forEach(vendor => {
    originalByVendor[vendor].forEach(it => {
      itemQtyMap[vendor + '|' + it.code + '|' + it.spec] = it.qty;
    });
  });

  // 참고파일(REFERENCE_FOLDER_ID) 연동 — 안 돼있으면 빈 값으로 폴백해서 capa/공휴일 검증 없이 진행
  const holidaySet = loadHolidays_(year);
  const demandCodes = new Set(orders.map(o => o.code)); // 이번 발주에 실제로 나온 품목만 호기 시뮬레이션 대상으로
  const capaInfo = loadProductionCapaInfo_(year, month, holidaySet, demandCodes);
  const weights = loadUnitWeights_();
  const codeDateLoad = {}; // 품목코드별 날짜별 이미 배정된 kg — 여러 업체가 같은 품목코드를 쓸 때 생산capa 공유

  const rows = [];
  const unknownVendors = new Set();
  const allIssues = [];
  const vendorDateLoad = {}; // 전사 날짜별 부하(kg) — 공휴일 시프트 때 "가장 한가한 날"을 고르는 기준으로 업체 간 공유

  // 요일 고정형(균등분배)을 먼저 처리하고 트럭버킷(여유일을 유동적으로 고름)을 나중에 처리한다 —
  // 순서를 반대로 하면 트럭버킷 쪽이 아직 반영 안 된 고정형의 미래 부하를 모른 채 "빈 날"로 착각해서
  // 같은 날에 겹쳐 배정해버릴 수 있다. 같은 그룹 안에서는 VENDOR_TYPE_MAP 순서를 그대로 유지해서
  // 실행마다 "가장 가벼운 날" 판단이 안 바뀌게 한다(정렬은 안정 정렬이라 그룹 내 순서 보존됨).
  const vendorOrder = Object.keys(VENDOR_TYPE_MAP).sort((a, b) => {
    const aFixed = BUCKET_WEEKDAYS[VENDOR_TYPE_MAP[a]] === undefined;
    const bFixed = BUCKET_WEEKDAYS[VENDOR_TYPE_MAP[b]] === undefined;
    return aFixed === bFixed ? 0 : (aFixed ? -1 : 1);
  });
  vendorOrder.forEach(vendor => {
    const items = byVendor[vendor];
    if (!items) return; // 이번 발주현황엔 이 업체 발주가 없음
    delete byVendor[vendor];

    const type = VENDOR_TYPE_MAP[vendor];
    const bucketWeekday = BUCKET_WEEKDAYS[type];

    if (bucketWeekday !== undefined) {
      // 트럭버킷 타입은 buildTruckBucketPlan_ 내부에서 이미 vendorDateLoad를 트럭대수로 갱신하므로
      // (companyBuckets.load를 CONFIG.TRUCK_KG로 나눠 올림), 여기서 또 합산하면 중복 반영됨 — 안 함.
      const result = buildTruckBucketPlan_(items, year, month, bucketWeekday, holidaySet, capaInfo, codeDateLoad, weights, null, vendorDateLoad);
      items.forEach(it => {
        const dateMap = result.plan[it.code + '|' + it.spec] || {};
        rows.push({ code: it.code, spec: it.spec, vendor: vendor, orderedQty: it.qty, dateMap: dateMap });
      });
      result.issues.forEach(i => allIssues.push(vendor + ' ' + i.code + ': ' + i.message));
    } else {
      // 트럭버킷이 아닌 타입(friday_even 등)은 여기서 직접 vendorDateLoad를 갱신해야 함 — 품목별로
      // 각각 올림해서 더하면(예: 2,789→1대 + 2,338→1대 + 1,913→1대 = 3대) 실제보다 부풀려진다.
      // 이 업체가 그 날짜에 실은 kg 총합을 먼저 구한 뒤 그걸 한 번만 트럭 1대(TRUCK_KG)로 나눠
      // 올려야 진짜 트럭 대수(예: 합계 7,040kg → 2대)가 된다.
      const dateTotals = {};
      items.forEach(it => {
        const dateMap = allocateByType(type, it.orders, year, month, holidaySet, null, vendorDateLoad);
        rows.push({ code: it.code, spec: it.spec, vendor: vendor, orderedQty: it.qty, dateMap: dateMap });
        Object.keys(dateMap).forEach(key => {
          dateTotals[key] = (dateTotals[key] || 0) + dateMap[key];
        });
      });
      Object.keys(dateTotals).forEach(key => {
        const trucks = Math.ceil(dateTotals[key] / CONFIG.TRUCK_KG);
        vendorDateLoad[key] = (vendorDateLoad[key] || 0) + trucks;
      });
    }
  });

  // VENDOR_TYPE_MAP에 없는 업체들(byVendor에 남은 것들)은 전부 제외
  Object.keys(byVendor).forEach(vendor => unknownVendors.add(vendor));

  rows.sort((a, b) => (a.vendor + a.code).localeCompare(b.vendor + b.code));

  // 신규 품목(스냅샷에 없음) / 변경된 품목(발주량이 지난번과 다름)에 changed 표시 — WritePlan.gs가 강조해서 그림
  rows.forEach(row => {
    const key = row.vendor + '|' + row.code + '|' + row.spec;
    const prevQty = previousSnapshot[key];
    const currentQty = itemQtyMap[key];
    row.changed = !isFirstSnapshot && (prevQty === undefined || Math.abs(prevQty - currentQty) > 0.01);
  });

  writePlanSheet(year, month, rows, { holidaySet: holidaySet, issueCount: allIssues.length, viewMode: 'item' });
  cleanupReferenceFiles_();
  cleanupOrderFile_();

  saveOrderSnapshot_(originalByVendor);
  saveFingerprints_(currentFingerprints);

  // 보기 방식 전환(showByVendor/showByItem)이 재계산 없이 다시 그릴 수 있도록 마지막 계획을 저장해둔다.
  PropertiesService.getDocumentProperties().setProperty('LAST_PLAN', JSON.stringify({
    year: year,
    month: month,
    rows: rows,
    issueCount: allIssues.length,
    holidayDates: Array.from(holidaySet),
    actualCutoffDateKey: null, // 새로 계획을 짠 것이므로 실적반영 기준일은 초기화
  }));

  const messages = [];
  if (unknownVendors.size > 0) messages.push('타입 매핑 없는 업체 제외: ' + Array.from(unknownVendors).join(', '));
  if (allIssues.length > 0) messages.push('검증 경고 ' + allIssues.length + '건 (실행이력 탭 참고)');

  // 챗 알림에는 건수만이 아니라 몇 품목 중 몇 개가 신규/변경인지, 결과 시트 링크까지 넣어서
  // 알림만 보고도 바로 클릭해서 확인할 수 있게 한다.
  const changedCount = rows.filter(r => r.changed).length;
  notifyChat_(
    (messages.length ? '⚠️ ' : '📦 ') +
    year + '년 ' + month + '월 출고계획 생성 완료\n' +
    '품목 ' + rows.length + '개(신규/변경 ' + changedCount + '개)' +
    (messages.length ? '\n' + messages.join('\n') : '') +
    '\n' + SpreadsheetApp.getActiveSpreadsheet().getUrl()
  );

  // 실행이력 탭에 남길 메시지 — 건수만이 아니라 실제 경고 내용을 그대로 적는다.
  const logLines = [];
  if (unknownVendors.size > 0) logLines.push('타입 매핑 없는 업체 제외: ' + Array.from(unknownVendors).join(', '));
  if (allIssues.length > 0) logLines.push.apply(logLines, allIssues);
  appendExecutionLog_(
    'generatePlan',
    logLines.length > 0 ? '경고' : '완료',
    logLines.length > 0 ? logLines.join('\n') : (year + '년 ' + month + '월 계획 생성, 이슈 없음')
  );

  Logger.log([
    '=== generatePlan ' + year + '-' + month + ' ===',
    'REFERENCE_FOLDER_ID 연동: ' + (CONFIG.REFERENCE_FOLDER_ID ? '됨' : '안 됨(capa/공휴일 검증 스킵)'),
  ].concat(allIssues).join('\n'));

  SpreadsheetApp.getActiveSpreadsheet().toast(
    messages.length ? messages.join(' / ') : '출고계획 생성 완료: ' + year + '년 ' + month + '월',
    messages.length ? '경고' : '완료',
    10
  );
}

/** 업체+품목코드+규격 기준으로 발주 건을 묶는다. */
function _groupOrders(orders) {
  const groups = {};
  orders.forEach(o => {
    const key = o.vendor + '|' + o.code + '|' + o.spec;
    if (!groups[key]) groups[key] = [];
    groups[key].push(o);
  });
  return groups;
}

/** 업체별로 [{code, spec, qty(합계), orders}] 목록을 만든다. */
function _groupByVendor_(orders) {
  const byKey = _groupOrders(orders);
  const byVendor = {};
  Object.keys(byKey).forEach(key => {
    const group = byKey[key];
    const vendor = group[0].vendor;
    const qty = group.reduce((s, o) => s + o.qty, 0);
    byVendor[vendor] = byVendor[vendor] || [];
    byVendor[vendor].push({ code: group[0].code, spec: group[0].spec, qty: qty, orders: group });
  });
  return byVendor;
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('출고계획 실행')
    .addItem('① 테스트 발주서 생성', 'createTestOrderStatus')
    .addItem('② 계획 생성', 'generatePlan')
    .addItem('③ 실적 반영', 'applyActualShipment')
    .addSeparator()
    .addItem('④ 회사별로 보기', 'showByVendor')
    .addItem('⑤ 품목별로 보기 (기본)', 'showByItem')
    .addToUi();
}
