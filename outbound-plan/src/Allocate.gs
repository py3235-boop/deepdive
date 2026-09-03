/**
 * 업체 타입별 날짜 배분 로직 — date_as_is / friday_even 전용.
 * 각 함수는 같은 업체+품목코드+규격으로 묶인 발주 건 배열(orders)을 받아
 * { 'yyyy-MM-dd': 배정수량, ... } 형태의 맵을 반환한다.
 *
 * wednesday_bucket(트럭버킷 타입)은 한 업체의 여러 품목을 같이(트럭 1대 상한을 품목끼리 나눠 써야
 * 해서) 처리해야 하므로 이 함수로 안 부르고, Main.gs가 TruckBucket.gs의 buildTruckBucketPlan_()을
 * 업체 단위로 직접 호출한다.
 */

function allocateByType(type, orders, year, month, holidaySet, minDate, vendorDateLoad) {
  switch (type) {
    case 'date_as_is':
      return allocateDateAsIs(orders);
    case 'friday_even':
      return allocateEvenOnWeekday_(orders, year, month, 5, holidaySet, minDate, vendorDateLoad); // 5 = 금요일
    default:
      throw new Error('allocateByType()에서 처리할 수 없는 타입: ' + type);
  }
}

/** 납기일에 적힌 그대로 사용. 같은 날짜에 여러 건이면 합산한다. 납기일 없는 행은 스킵(orphaned_weight_no_date). */
function allocateDateAsIs(orders) {
  const map = {};
  orders.forEach(o => {
    if (!o.dueDate) return;
    const key = dateKey_(o.dueDate);
    map[key] = (map[key] || 0) + o.qty;
  });
  return map;
}

/**
 * 그 달의 특정 요일 전체에 총 발주량을 균등분배한다(마지막 날짜가 반올림 잔여를 흡수).
 * 기준일이 공휴일과 겹치면 TruckBucket.gs의 _shiftToLightestWorkday_로 다음 근무일로 옮긴다
 * (트럭버킷 타입과 같은 방식 — friday_even도 공휴일엔 배정되면 안 되므로).
 *
 * minDate를 주면(실적반영에서 "오늘 이후"만 재배분할 때) 그 날짜보다 이른 요일은 후보에서 뺀다.
 * minDate 때문에 후보가 하나도 안 남으면(이번 달엔 더 배정 못함) 에러 대신 빈 맵을 반환한다.
 */
function allocateEvenOnWeekday_(orders, year, month, weekday, holidaySet, minDate, vendorDateLoad) {
  holidaySet = holidaySet || new Set();
  const total = orders.reduce((sum, o) => sum + o.qty, 0);

  let dates = _generateBaseDates_(year, month, weekday, holidaySet, vendorDateLoad);
  if (minDate) dates = dates.filter(d => d > minDate);
  dates.sort((a, b) => a - b);

  if (dates.length === 0) {
    if (minDate) return {}; // 재배분할 날짜가 이번 달에 더 없음 — 호출부가 issue로 처리
    throw new Error(year + '-' + month + '에는 해당 요일이 없습니다(weekday=' + weekday + ')');
  }

  const map = {};
  const base = Math.round(total / dates.length);
  let assigned = 0;
  dates.forEach((date, i) => {
    const isLast = i === dates.length - 1;
    const qty = isLast ? total - assigned : base;
    map[dateKey_(date)] = qty;
    assigned += qty;
  });
  return map;
}
