/**
 * 트럭버킷 배분 — 업체 타입별 요일 규칙 중 트럭 상한이 있는 타입(monday_bucket/wednesday_bucket 등) 전용.
 *
 * 핵심 차이(예전 버전 대비): 예전엔 "연속값으로 대충 나눈 뒤 나중에 단위중량 배수로 반올림"하는
 * 2단계 방식이라, 반올림 오차가 다른 날짜로 넘어가면서 하루 상한을 다시 넘기는 버그가 반복됐다.
 * 지금은 **처음부터 끝까지 "한 번에 단위중량 1개씩만" 배정**하는 라운드로빈이라, 반올림이라는
 * 단계 자체가 없다 — 그래서 애초에 상한을 넘는 값이 나올 수 없다.
 *
 * 같은 업체(요일 규칙이 같은 업체)의 품목들을 한꺼번에 받아서, 주간목표 재계산 + 하루 상한(업체별
 * 독립, 정확히 트럭 1대=TRUCK_KG) + 공휴일 시프트 + 여유일 확보 + 생산capa 검증까지 한 번에 처리한다.
 *
 * items: [{ code, spec, qty }] (qty = 이번 배분에 필요한 총중량, 같은 업체 내 code+spec 기준 합산됨)
 * vendorDateLoad: { 'yyyy-MM-dd': 그 날짜에 이미 확정된 업체별 트럭 대수(정수, 업체마다 올림해서 합산) } —
 *   공휴일 시프트/여유일 선택 때 "전사 기준 가장 한가한 날"을 고르는 용도로만 쓰고, 이 업체 자신의
 *   하루 상한에는 영향을 안 준다(업체별 독립 원칙).
 * codeDateLoad: { code: { 'yyyy-MM-dd': 이미 배정된 kg } } — 여러 업체가 같은 품목코드를 쓸 때
 *   생산capa를 전사 공통으로 공유하기 위한 레지스트리(호출부에서 만들어서 여러 번 넘겨줌)
 *
 * 반환: { plan: { 'code|spec': {날짜: kg} }, issues: [{code, message}] }
 */
function buildTruckBucketPlan_(items, year, month, weekday, holidaySet, capaInfo, codeDateLoad, weights, minDate, vendorDateLoad) {
  const truckKg = CONFIG.TRUCK_KG; // 업체별 하루 상한 = 정확히 트럭 1대
  vendorDateLoad = vendorDateLoad || {};
  codeDateLoad = codeDateLoad || {};
  weights = weights || { byCode: {}, bySpec: {} };
  const issues = [];

  let baseDates = _generateBaseDates_(year, month, weekday, holidaySet, vendorDateLoad);
  if (minDate) baseDates = baseDates.filter(d => d > minDate);

  const itemState = items.map(it => ({
    key: it.code + '|' + it.spec,
    code: it.code,
    spec: it.spec,
    remaining: it.qty,
    unitWeight: lookupUnitWeight_(weights, it.code, it.spec),
    dateMap: {},
  }));
  // 단위중량 큰 품목부터(같으면 물량 큰 품목부터) 처리 
  itemState.sort((a, b) => (b.unitWeight || 0) - (a.unitWeight || 0) || b.remaining - a.remaining);

  const totalWeight = itemState.reduce((s, it) => s + it.remaining, 0);
  if (baseDates.length === 0 || totalWeight <= 0) {
    const plan = {};
    itemState.forEach(it => { plan[it.key] = {}; });
    if (baseDates.length === 0 && totalWeight > 0) {
      itemState.forEach(it => {
        if (it.remaining > 0.0001) {
          issues.push({ code: it.code, message: '이번 달엔 기준요일(또는 재배분 가능한 날짜)이 없음 — ' + Math.round(it.remaining) + 'kg 미배정' });
        }
      });
    }
    return { plan: plan, issues: issues };
  }

  // 이번 실행 중 이 업체가 실제로 쓴 모든 날짜의 버킷(cap=truckKg 그대로) — 자투리 마무리 때 새 날짜를
  // 여는 대신 이미 확보해둔 날짜 중 여유 있는 곳부터 채우기 위함.
  const companyBuckets = {};
  // 이 업체가 각 날짜에 대해 vendorDateLoad에 마지막으로 반영한 트럭대수 — 소량 자투리 병합 때
  // "얼마를 더하고 뺐는지"를 라이브 버킷 값이 아니라 이 값 기준으로 계산해야, 다른 업체 몫을
  // 잘못 갉아먹지 않는다(라이브 버킷은 vendorDateLoad에 아직 안 올라간 초과배정을 포함할 수 있음).
  const vendorTrucksReported = {};

  function getBucket(date) {
    const key = dateKey_(date);
    let b = companyBuckets[key];
    if (!b) { b = { date: date, load: 0, cap: truckKg }; companyBuckets[key] = b; }
    return b;
  }

  // 품목코드별 생산capa 제약이 걸리는 날짜만 걸러내는 필터. capa 정보 없는 품목은 필터링 없이 그대로 통과.
  // 이번 주 후보 중 capa가 열린 날짜가 하나도 없으면(예전엔 여기서 capa 무시하고 아무 날짜나 썼음)
  // 빈 배열을 그대로 반환한다 — 호출부가 "이번 주는 건너뛰고 다음 주(그만큼 더 생산된 뒤)로 미룬다".
  function productionFiltered(it, weekBuckets) {
    const curveFn = buildCapaCumulativeFn_(capaInfo, it.code, year, month, holidaySet);
    if (!curveFn) return weekBuckets;
    const codeLoad = codeDateLoad[it.code] || (codeDateLoad[it.code] = {});
    return weekBuckets.filter(b => {
      const cap = curveFn(b.date);
      const used = codeLoad[dateKey_(b.date)] || 0;
      return used + it.unitWeight <= cap;
    });
  }
  function recordCodeLoad(it, key, kg) {
    const curveFn = buildCapaCumulativeFn_(capaInfo, it.code, year, month, holidaySet);
    if (!curveFn) return;
    const codeLoad = codeDateLoad[it.code] || (codeDateLoad[it.code] = {});
    codeLoad[key] = (codeLoad[key] || 0) + kg;
    const cap = curveFn(new Date(key));
    if (codeLoad[key] > cap) {
      issues.push({ code: it.code, message: key + ' 생산capa(' + Math.round(cap) + 'kg) 초과해서 배정됨(production_capa_short)' });
    }
  }

  // 여러 후보(candidates) 중에서 채울 버킷과 이번에 담을 최대량을 정한다. 그 품목 "자기 자신"이 이미
  // 그 날짜에 얼마나 실렸는지(ownLoadByDate)를 최우선으로 봐서, 아직 안 실린 날부터 채운다 — 이게
  // 없으면 한 품목의 물량 상당수가 맨 첫 날짜에만 몰리는 문제가 생긴다.
  function pickBucket(candidates, unitWeight, ownLoadByDate) {
    const fittable = candidates.filter(b => b.load < b.cap && _floorToUnit_(Math.min(b.cap, b.cap - b.load), unitWeight) > 0);
    if (fittable.length > 0) {
      const ownLoad = b => (ownLoadByDate && ownLoadByDate[dateKey_(b.date)]) || 0;
      let best = fittable[0];
      let bestOwn = ownLoad(best);
      for (let i = 1; i < fittable.length; i++) {
        const b = fittable[i];
        const bOwn = ownLoad(b);
        if (bOwn !== bestOwn) {
          if (bOwn < bestOwn) { best = b; bestOwn = bOwn; }
          continue;
        }
        if (b.load > best.load) { best = b; bestOwn = bOwn; continue; }
        if (b.load === best.load && b.date < best.date) { best = b; bestOwn = bOwn; }
      }
      return { bucket: best, maxTake: best.cap - best.load };
    }
    // 자리가 전혀 없으면(다 꽉 참) 그나마 가장 덜 실린 날에 단위중량 1개만큼 강제로 배정(최후수단)
    let best2 = candidates[0];
    for (let j = 1; j < candidates.length; j++) {
      if (candidates[j].load < best2.load) best2 = candidates[j];
    }
    return { bucket: best2, maxTake: Math.max(best2.cap - best2.load, unitWeight || 0) };
  }

  // ① 기준일이 도는 순서대로(=한 주씩) 처리
  baseDates.forEach((baseDate, weekIdx) => {
    const remainingTotal = itemState.reduce((s, it) => s + it.remaining, 0);
    if (remainingTotal <= 0) return;

    const weeksLeft = baseDates.length - weekIdx;
    const thisWeekTarget = remainingTotal / weeksLeft; // 매주 다시 계산 — 부족분이 마지막 주로만 안 몰리게
    const daysNeeded = Math.max(1, Math.ceil(thisWeekTarget / truckKg));

    let weekDates = [baseDate];
    if (daysNeeded > 1) {
      const extra = _weekScopedExtraDates_(baseDate, daysNeeded - 1, weekDates, vendorDateLoad, holidaySet, minDate);
      weekDates = weekDates.concat(extra).sort((a, b) => a - b);
    }
    let weekBuckets = weekDates.map(getBucket);

    function ensureRoom(it) {
      const hasRoom = weekBuckets.some(b => b.load < b.cap && _floorToUnit_(Math.min(b.cap, b.cap - b.load), it.unitWeight) > 0);
      if (hasRoom) return;
      // "새 날짜"는 이번 주만이 아니라 이 업체가 지금까지 쓴 모든 날짜(companyBuckets)를 제외해야
      // 진짜 새 날짜가 나온다 — 안 그러면 dateWeightMap 기준 "가벼운 날"로 보이는 옛 날짜로 계속 돌아감.
      const usedDates = Object.keys(companyBuckets).map(k => companyBuckets[k].date);
      const more = _weekScopedExtraDates_(baseDate, 1, usedDates, vendorDateLoad, holidaySet, minDate);
      const newDate = more.length > 0 ? more[0] : _monthWideExtraDate_(year, month, usedDates, vendorDateLoad, holidaySet, minDate);
      if (!newDate) return; // 이번 달 전체를 뒤져도 더 쓸 평일이 없음 — 최후수단(초과)으로 갈 수밖에 없음
      weekDates.push(newDate);
      weekDates.sort((a, b) => a - b);
      weekBuckets.push(getBucket(newDate));
    }

    // 라운드로빈: 모든 품목이 한 차례씩 딱 단위중량 1개만 배정받고 다음 품목으로 넘어감
    let weekRemain = thisWeekTarget;
    let pending = itemState.filter(it => it.remaining > 0 && it.unitWeight > 0);
    let guard = 0;
    while (pending.length > 0 && weekRemain > 0 && guard < 50000) {
      guard++;
      let progressed = false;
      pending.forEach(it => {
        if (it.remaining <= 0 || weekRemain <= 0) return;
        ensureRoom(it);
        const capaCandidates = productionFiltered(it, weekBuckets);
        if (capaCandidates.length === 0) return; // 이번 주는 생산capa가 전혀 안 열림 — 억지로 배정하지 않고 다음 주로 미룸
        const picked = pickBucket(capaCandidates, it.unitWeight, it.dateMap);
        const best = picked.bucket;

        // picked가 진짜 여유 있는 자리가 아니라 최후수단 fallback이면(트럭 자리가 다 찼을 수 있음),
        // 트럭 자리 기준(raw weekBuckets)으로는 진짜 여유가 있는지 다시 확인해서,
        // 있으면 이번 차례는 그냥 건너뛴다(억지로 트럭 1대를 넘기지 않음).
        const trulyFits = _floorToUnit_(Math.min(best.cap, best.cap - best.load), it.unitWeight) > 0;
        if (!trulyFits) {
          const rawHasRoom = weekBuckets.some(b => b.load < b.cap && _floorToUnit_(Math.min(b.cap, b.cap - b.load), it.unitWeight) > 0);
          if (rawHasRoom) return;
        }

        const take = _floorToUnit_(Math.min(it.remaining, picked.maxTake, it.unitWeight, weekRemain), it.unitWeight);
        if (take <= 0) return;

        const key = dateKey_(best.date);
        it.dateMap[key] = (it.dateMap[key] || 0) + take;
        best.load += take;
        it.remaining -= take;
        weekRemain -= take;
        recordCodeLoad(it, key, take);
        progressed = true;
      });
      pending = pending.filter(it => it.remaining > 0);
      if (!progressed) break;
    }

    // 이 업체가 이 날짜에 실은 자기 몫(kg)을 트럭 대수로 올림해서 vendorDateLoad에 반영
    weekBuckets.forEach(b => {
      if (b.load <= 0) return;
      const key = dateKey_(b.date);
      const trucks = Math.ceil(b.load / CONFIG.TRUCK_KG);
      vendorDateLoad[key] = (vendorDateLoad[key] || 0) + trucks;
      vendorTrucksReported[key] = trucks;
    });
  });

  // 단위중량을 모르는 품목은 기준일 전체를 대상으로 한 번에 몰아서 배정
  const noUnitWeightItems = itemState.filter(it => it.remaining > 0 && !(it.unitWeight > 0));
  if (noUnitWeightItems.length > 0) {
    const allBuckets = baseDates.map(d => ({ date: d, load: 0, cap: truckKg }));
    noUnitWeightItems.forEach(it => {
      while (it.remaining > 0) {
        const picked = pickBucket(allBuckets, it.unitWeight, it.dateMap);
        const take = Math.min(it.remaining, picked.maxTake);
        if (take <= 0) break;
        const key = dateKey_(picked.bucket.date);
        it.dateMap[key] = (it.dateMap[key] || 0) + take;
        picked.bucket.load += take;
        it.remaining -= take;
      }
    });
    allBuckets.forEach(b => {
      if (b.load <= 0) return;
      const key = dateKey_(b.date);
      const trucks = Math.ceil(b.load / CONFIG.TRUCK_KG);
      vendorDateLoad[key] = (vendorDateLoad[key] || 0) + trucks;
      vendorTrucksReported[key] = trucks;
    });
  }

  // ② 자투리 마무리: 이 업체가 이미 확보해둔 날짜들 중 여유 있는 곳부터 채우고, 그래도 안 되면 새
  // 날짜를 하나 더 찾고, 그마저 없을 때만 마지막 기준일에 최소한만 초과해서 마무리한다.
  itemState.filter(it => it.remaining > 0 && it.unitWeight > 0).forEach(it => {
    let candidates = Object.values(companyBuckets);
    const hasRoom = candidates.some(b => b.load < b.cap && _floorToUnit_(Math.min(b.cap, b.cap - b.load), it.unitWeight) > 0);
    if (!hasRoom) {
      const newDate = _monthWideExtraDate_(year, month, candidates.map(b => b.date), vendorDateLoad, holidaySet, minDate);
      if (newDate) {
        getBucket(newDate);
        candidates = Object.values(companyBuckets);
      }
    }
    if (candidates.length === 0) candidates = [{ date: baseDates[baseDates.length - 1], load: 0, cap: truckKg }];

    // 여기서도 capa가 열려있는 날짜가 있으면 그쪽을 우선한다 — 정말 하나도 없을 때만(월말까지 가도
    // 생산이 못 따라온 경우) capa를 무시하고 배정해서 최소한 이번 달에 나가긴 나가게 한다.
    const curveFn = buildCapaCumulativeFn_(capaInfo, it.code, year, month, holidaySet);
    let pickCandidates = candidates;
    if (curveFn) {
      const codeLoad = codeDateLoad[it.code] || (codeDateLoad[it.code] = {});
      const capaOk = candidates.filter(b => (codeLoad[dateKey_(b.date)] || 0) + it.unitWeight <= curveFn(b.date));
      if (capaOk.length > 0) pickCandidates = capaOk;
    }

    const roundedRemain = _roundToUnit_(it.remaining, it.unitWeight);
    if (roundedRemain > 0) {
      const picked = pickBucket(pickCandidates, it.unitWeight, it.dateMap);
      const key = dateKey_(picked.bucket.date);
      it.dateMap[key] = (it.dateMap[key] || 0) + roundedRemain;
      picked.bucket.load += roundedRemain;
      recordCodeLoad(it, key, roundedRemain);
    }
    it.remaining = 0;
  });

  // ③ 소량 자투리 병합: 배정이 다 끝난 뒤, 어떤 날짜의 배정량이 트럭 절반(MIN_SHIPMENT_KG)보다
  // 작으면 그 품목의 다음(더 나중) 배정일로 합친다. capa는 시간이 지날수록만 커지므로 "앞으로만"
  // 옮기는 건 항상 안전하다 — 뒤로 옮기면 그 이전 시점 capa를 넘길 수 있어서 하지 않는다. 트럭
  // 상한이나 생산capa를 넘기게 되면 합치지 않고 그대로 둔다(억지로 옮기지 않음).
  // unitWeight가 없는 품목은 companyBuckets가 아니라 별도의 allBuckets를 썼으므로(위 "단위중량을
  // 모르는 품목" 블록) companyBuckets에는 이 품목의 진짜 부하가 반영돼 있지 않다 — 그런 품목을
  // 여기서 건드리면 companyBuckets를 실제와 다르게 오염시키게 되므로 대상에서 제외한다.
  itemState.filter(it => it.unitWeight > 0).forEach(it => {
    const curveFn = buildCapaCumulativeFn_(capaInfo, it.code, year, month, holidaySet);
    const codeLoad = curveFn ? (codeDateLoad[it.code] || (codeDateLoad[it.code] = {})) : null;

    const keys = Object.keys(it.dateMap).sort();
    for (let i = 0; i < keys.length - 1; i++) {
      const sourceKey = keys[i];
      const amount = it.dateMap[sourceKey];
      if (!amount || amount >= CONFIG.MIN_SHIPMENT_KG) continue;

      const targetKey = keys[i + 1];
      const targetBucket = companyBuckets[targetKey];
      if (!targetBucket) continue; // 이론상 항상 있어야 하지만 방어적으로
      if (targetBucket.load + amount > targetBucket.cap) continue; // 트럭 상한 초과 — 합치지 않음

      if (curveFn) {
        const targetCap = curveFn(targetBucket.date);
        const targetUsed = codeLoad[targetKey] || 0;
        if (targetUsed + amount > targetCap) continue; // 생산capa 초과 — 합치지 않음
      }

      const sourceBucket = companyBuckets[sourceKey];

      it.dateMap[targetKey] = (it.dateMap[targetKey] || 0) + amount;
      delete it.dateMap[sourceKey];
      if (sourceBucket) sourceBucket.load -= amount;
      targetBucket.load += amount;
      if (codeLoad) {
        codeLoad[targetKey] = (codeLoad[targetKey] || 0) + amount;
        if (codeLoad[sourceKey] !== undefined) {
          codeLoad[sourceKey] -= amount;
          if (codeLoad[sourceKey] <= 0.0001) delete codeLoad[sourceKey];
        }
      }

      // vendorDateLoad는 이 업체가 "마지막으로 보고한" 트럭대수(vendorTrucksReported) 기준으로만
      // 델타를 적용한다 — 라이브 버킷값을 기준으로 하면, 이 버킷에 최후수단으로 트럭 상한을 넘겨서
      // 강제 배정된(그러나 아직 vendorDateLoad에 반영 안 된) 몫까지 함께 계산에 끼어들어 다른
      // 업체 몫을 잘못 갉아먹을 수 있다.
      const afterSourceTrucks = sourceBucket ? Math.ceil(Math.max(0, sourceBucket.load) / truckKg) : 0;
      const afterTargetTrucks = Math.ceil(targetBucket.load / truckKg);
      if (sourceBucket) {
        const beforeSourceTrucks = vendorTrucksReported[sourceKey] || 0;
        vendorDateLoad[sourceKey] = Math.max(0, (vendorDateLoad[sourceKey] || 0) - beforeSourceTrucks + afterSourceTrucks);
        vendorTrucksReported[sourceKey] = afterSourceTrucks;
      }
      const beforeTargetTrucks = vendorTrucksReported[targetKey] || 0;
      vendorDateLoad[targetKey] = Math.max(0, (vendorDateLoad[targetKey] || 0) - beforeTargetTrucks + afterTargetTrucks);
      vendorTrucksReported[targetKey] = afterTargetTrucks;
    }
  });

  const plan = {};
  itemState.forEach(it => { plan[it.key] = it.dateMap; });
  return { plan: plan, issues: issues };
}

/** 값을 unitWeight의 배수로 반올림(모르면 정수 반올림). */
function _roundToUnit_(value, unitWeight) {
  if (!unitWeight || unitWeight <= 0) return Math.round(value);
  return Math.round(value / unitWeight) * unitWeight;
}

/** 값을 unitWeight의 배수로 내림(모르면 그대로). */
function _floorToUnit_(value, unitWeight) {
  if (!unitWeight || unitWeight <= 0) return value;
  return Math.floor(value / unitWeight) * unitWeight;
}

/** 그 달의 기준요일 날짜 전체를 만들고, 공휴일과 겹치면 시프트한 뒤, 같은 날짜로 밀린 중복은 제거한다. */
function _generateBaseDates_(year, month, weekday, holidaySet, vendorDateLoad) {
  const raw = [];
  const d = new Date(year, month - 1, 1);
  while (d.getMonth() === month - 1) {
    if (d.getDay() === weekday) raw.push(new Date(d));
    d.setDate(d.getDate() + 1);
  }

  const seen = new Set();
  const dates = [];
  raw.forEach(rd => {
    const shifted = _shiftToLightestWorkday_(rd, holidaySet, vendorDateLoad);
    const key = dateKey_(shifted);
    if (seen.has(key)) return; // 여러 기준일이 같은 날로 밀리면 하나만 남김
    seen.add(key);
    dates.push(shifted);
  });
  return dates;
}

/**
 * 기준요일이 주말/공휴일과 겹치면 다음 근무일 후보를 최대 3개 모아서, vendorDateLoad(전사 트럭
 * 대수) 기준으로 가장 한가한 날을 고른다. 원래 평일이면 손대지 않는다.
 */
function _shiftToLightestWorkday_(date, holidaySet, vendorDateLoad) {
  if (!(date.getDay() === 0 || date.getDay() === 6 || holidaySet.has(dateKey_(date)))) return new Date(date);

  const candidates = [];
  let d = new Date(date);
  let guard = 0;
  while (candidates.length < 3 && guard < 30) {
    d = new Date(d.getFullYear(), d.getMonth(), d.getDate() + 1);
    guard++;
    if (d.getDay() === 0 || d.getDay() === 6 || holidaySet.has(dateKey_(d))) continue;
    candidates.push(new Date(d));
  }
  if (candidates.length === 0) return _shiftToWorkday_(date, holidaySet);

  let best = candidates[0];
  let bestLoad = (vendorDateLoad && vendorDateLoad[dateKey_(best)]) || 0;
  for (let i = 1; i < candidates.length; i++) {
    const load = (vendorDateLoad && vendorDateLoad[dateKey_(candidates[i])]) || 0;
    if (load < bestLoad) { best = candidates[i]; bestLoad = load; }
  }
  return best;
}

/** 주말이거나 공휴일이면 평일이 나올 때까지 하루씩 미룬다(그 달을 넘어가도 계속 밀림). */
function _shiftToWorkday_(date, holidaySet) {
  const d = new Date(date);
  let guard = 0;
  while ((d.getDay() === 0 || d.getDay() === 6 || holidaySet.has(dateKey_(d))) && guard < 30) {
    d.setDate(d.getDate() + 1);
    guard++;
  }
  return d;
}

/**
 * baseDate가 속한 주(월~금, 같은 달 안에서만) 중 existingDates와 안 겹치는 평일을 count개 고른다.
 * "전사 기준 가벼운 날"(LIGHT_DAY_TRUCK_THRESHOLD 이하) 우선 + 이미 고른 날과 하루 이상 떨어진 날 우선.
 */
function _weekScopedExtraDates_(baseDate, count, existingDates, vendorDateLoad, holidaySet, minDate) {
  vendorDateLoad = vendorDateLoad || {};
  const day = baseDate.getDay();
  const mondayOffset = day === 0 ? -6 : 1 - day;
  const monday = new Date(baseDate.getFullYear(), baseDate.getMonth(), baseDate.getDate() + mondayOffset);

  let candidates = [];
  for (let i = 0; i < 5; i++) {
    const d = new Date(monday.getFullYear(), monday.getMonth(), monday.getDate() + i);
    if (d.getMonth() !== baseDate.getMonth()) continue; // 그 주가 달 경계를 넘으면 다음/이전 달 날짜는 제외
    if (holidaySet.has(dateKey_(d))) continue;
    candidates.push(d);
  }
  if (minDate) candidates = candidates.filter(d => d > minDate);

  const existingKeys = new Set(existingDates.map(d => dateKey_(d)));
  candidates = candidates.filter(d => !existingKeys.has(dateKey_(d)));

  const scored = candidates.map(d => {
    const trucksUsed = vendorDateLoad[dateKey_(d)] || 0;
    return { date: d, isLight: trucksUsed < CONFIG.LIGHT_DAY_TRUCK_THRESHOLD, load: trucksUsed };
  });
  scored.sort((a, b) => {
    if (a.isLight !== b.isLight) return a.isLight ? -1 : 1;
    if (a.load !== b.load) return a.load - b.load;
    return a.date - b.date;
  });

  const ONE_DAY_MS = 24 * 60 * 60 * 1000;
  const chosenTimes = existingDates.map(d => d.getTime());
  function isTooClose(t) {
    return chosenTimes.some(ct => Math.abs(ct - t) <= ONE_DAY_MS);
  }

  // "가벼운 날"이 최우선 기준이다 — 이미 다른 업체로 꽉 찬 날은, 기준일과 하루밖에 안 떨어진 날보다도
  // 뒤로 밀려야 한다. 그래서 가벼운 후보군 안에서 먼저 "하루 이상 떨어진 날"을 고르고, 그래도 count를
  // 못 채우면 가벼운 후보군의 "가까운 날"로 채우고, 가벼운 후보가 아예 없을 때만 무거운 후보로 넘어간다.
  const result = [];
  function fillFrom(pool) {
    const far = pool.filter(s => !isTooClose(s.date.getTime()));
    const close = pool.filter(s => isTooClose(s.date.getTime()));
    far.concat(close).forEach(s => {
      if (result.length >= count) return;
      result.push(s.date);
      chosenTimes.push(s.date.getTime());
    });
  }
  fillFrom(scored.filter(s => s.isLight));
  if (result.length < count) fillFrom(scored.filter(s => !s.isLight));

  result.sort((a, b) => a - b);
  return result;
}

/** 그 주 안에서 정말 자리가 없을 때(월말/월초+공휴일 겹침 등), 이번 달 전체에서 가장 한가한 새 평일을 찾는다. */
function _monthWideExtraDate_(year, month, existingDates, vendorDateLoad, holidaySet, minDate) {
  vendorDateLoad = vendorDateLoad || {};
  const totalDays = daysInMonth_(year, month);
  const existingKeys = new Set(existingDates.map(d => dateKey_(d)));

  let candidates = [];
  for (let d = 1; d <= totalDays; d++) {
    const date = new Date(year, month - 1, d);
    const day = date.getDay();
    if (day === 0 || day === 6) continue;
    if (holidaySet.has(dateKey_(date))) continue;
    if (existingKeys.has(dateKey_(date))) continue;
    candidates.push(date);
  }
  if (minDate) candidates = candidates.filter(d => d > minDate);
  if (candidates.length === 0) return null;

  let best = candidates[0];
  let bestLoad = vendorDateLoad[dateKey_(best)] || 0;
  for (let i = 1; i < candidates.length; i++) {
    const load = vendorDateLoad[dateKey_(candidates[i])] || 0;
    if (load < bestLoad) { best = candidates[i]; bestLoad = load; }
  }
  return best;
}
