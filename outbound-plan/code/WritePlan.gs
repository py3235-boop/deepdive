/**
 * 출고계획 결과를 시트에 쓴다.
 * rows: [{ code, spec, vendor, orderedQty(발주량), dateMap: {'yyyy-MM-dd': qty} }, ...]
 * meta: { holidaySet(Set), issueCount(검증사항 건수), viewMode('item'|'vendor'), sheetName(선택) }
 *
 * viewMode:
 *  - 'item'(기본): 배너 없이 규격순 정렬
 *  - 'vendor': 업체 순서(VENDOR_TYPE_MAP 고정 순서)대로 "◆ 업체명" 배너 + 그 업체 소계 행을 넣어서 보여줌
 *
 * 색상/배경 등 디자인은 일단 다 뺐고, 굵게/정렬/숫자서식/틀고정 같은 구조적인 것만 남겼다.
 *
 * 실제 데이터가 담긴 탭은 "N월 출고계획"(그 계획의 대상 월) 이름으로 쓰고, 달마다 별도 탭으로
 * 계속 남는다(다음 달 발주서를 미리 넣어놔도 이번 달 탭이 안 지워짐). 이 중 **오늘 날짜 기준
 * 이번 달** 탭만 1번째로 옮기고, 그 외(미리 준비해둔 다음 달 등)는 맨 뒤로 보낸다.
 * `meta.sheetName`을 넘기면(resetPlan()의 "기본 시트") 그 고정 이름 탭에 쓰고, 이번 달 탭이
 * 이미 있으면 그 뒤(2번째)로, 없으면 1번째로 온다.
 *
 * 파일 이름 자체는 항상 "오늘 기준 이번 달_출고계획(통합)"으로 맞춘다(쓰는 대상 월이 아니라
 * 오늘 날짜 기준) — 다른 파트가 폴더 위치가 아니라 이 파일명 패턴으로 "이번 달 출고계획" 파일을
 * 찾기 때문에, 다음 달 탭을 미리 만들어놔도 파일 이름 자체는 항상 이번 달 기준이어야 한다.
 */

function writePlanSheet(year, month, rows, meta) {
  meta = meta || {};
  const issueCount = meta.issueCount || 0;
  const viewMode = meta.viewMode === 'vendor' ? 'vendor' : 'item';
  const holidaySet = meta.holidaySet || new Set();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const today = new Date();
  const todayYear = today.getFullYear();
  const todayMonth = today.getMonth() + 1;
  const isCurrentMonth = year === todayYear && month === todayMonth;
  const isBaseSheet = !!meta.sheetName;
  const sheetName = meta.sheetName || (month + '월 출고계획');

  const fileName = todayMonth + '월_출고계획(통합)';
  if (ss.getName() !== fileName) ss.rename(fileName);

  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) sheet = ss.insertSheet(sheetName);

  sheet.clear();
  sheet.getRange(1, 1).breakApart(); // 이전 실행에서 남아있을 수 있는 병합 셀 해제
  sheet.setFrozenRows(0);
  sheet.setFrozenColumns(0); // 고정 상태로는 열 병합이 막히므로 병합 전에 미리 풀어둔다

  const totalDays = daysInMonth_(year, month);
  const dateKeys = [];
  const header = ['품목코드', '규격', '구분(고객사)', '발주량', '합계'];
  for (let d = 1; d <= totalDays; d++) {
    const date = new Date(year, month - 1, d);
    dateKeys.push(dateKey_(date));
    header.push(formatDayLabel_(d));
  }

  const orderedQtyCol = 4; // D열: 발주량
  const totalCol = 5; // E열: 합계
  const firstDateCol = 6; // F열부터 날짜
  const totalCols = header.length;

  // --- 행 배치: 1행 제목, 2행 헤더, 3행부터 데이터 ---
  const titleRow = 1;
  const headerRow = 2;
  const dataStartRow = 3;

  const titleSuffix = viewMode === 'vendor' ? ' — 회사별 보기' : ' — 품목별 보기';
  sheet.getRange(titleRow, 1).setValue(month + '월 출고계획(통합)' + titleSuffix)
    .setFontWeight('bold').setHorizontalAlignment('left');

  // --- 헤더 ---
  sheet.getRange(headerRow, 1, 1, totalCols).setValues([header])
    .setFontWeight('bold').setHorizontalAlignment('center');

  // 주말/공휴일만 색상 표시(일요일·공휴일=빨강, 토요일=파랑) — 셀마다 따로 부르지 않고 헤더 행
  // 전체를 배경색/글자색 배열로 만들어서 한 번에 적용한다(아래 데이터 영역과 같은 이유).
  const headerBg = new Array(totalCols).fill(null);
  const headerFg = new Array(totalCols).fill(null);
  for (let d = 1; d <= totalDays; d++) {
    const date = new Date(year, month - 1, d);
    const day = date.getDay();
    const idx = firstDateCol - 1 + d - 1;
    if (day === 0 || holidaySet.has(dateKey_(date))) {
      headerBg[idx] = '#C0392B'; headerFg[idx] = '#FFFFFF';
    } else if (day === 6) {
      headerBg[idx] = '#2980B9'; headerFg[idx] = '#FFFFFF';
    }
  }
  sheet.getRange(headerRow, 1, 1, totalCols).setBackgrounds([headerBg]).setFontColors([headerFg]);

  // --- 표시 순서 구성: 'item'은 규격순 한 덩어리, 'vendor'는 업체별 배너+소계로 분리 ---
  const blocks = _buildDisplayBlocks_(rows, viewMode); // [{ vendor|null, items: [row,...] }, ...]

  // 데이터 영역(3행부터 총계 행까지)을 품목 수만큼 Range API를 반복 호출하는 대신, 값/서식을 배열로
  // 전부 모았다가 setValues/setNumberFormats/setFontWeights/setFontColors/setHorizontalAlignments를
  // 딱 한 번씩만 부른다 — 품목이 많을 때(자동 트리거 감지 후 실제로 계획이 채워지기까지 몇 분씩
  // 걸리던 원인) 이 방식이 압도적으로 빠르다. 합계/소계 열도 SUM 수식 대신, 이미 메모리에 있는
  // dateMap 값을 그대로 더해 값으로 채운다(수식 설정 자체가 느린 연산이라 값으로 대체).
  // 신규/변경 품목 강조색 — 글자색만으로는 발표 화면에서 눈에 잘 안 띄어서, 배경색까지 같이 입혀
  // "어디가 왜 바뀌었는지"가 표를 훑어보기만 해도 바로 보이게 한다(아래 범례 문구와 짝을 이룸).
  const CHANGE_STYLE = {
    new: { bg: '#D9EAD3', fg: '#38761D' }, // 연두 배경 + 진한 초록 글자: 이번에 새로 생긴 품목
    changed: { bg: '#FCE5CD', fg: '#B45F06' }, // 연주황 배경 + 진한 주황 글자: 발주량이 지난번과 달라진 품목
  };

  const numFmt = '#,##0';
  const valuesGrid = [];
  const numberFormatsGrid = [];
  const fontWeightsGrid = [];
  const fontColorsGrid = [];
  const backgroundsGrid = [];
  const alignGrid = [];

  blocks.forEach(block => {
    if (block.vendor !== null) {
      const bannerLine = new Array(totalCols).fill('');
      bannerLine[0] = '◆ ' + block.vendor;
      valuesGrid.push(bannerLine);
      numberFormatsGrid.push(new Array(totalCols).fill('General'));
      fontWeightsGrid.push(new Array(totalCols).fill('bold'));
      fontColorsGrid.push(new Array(totalCols).fill(null));
      backgroundsGrid.push(new Array(totalCols).fill(null));
      alignGrid.push(new Array(totalCols).fill(null)); // 원래도 정렬 지정 없이 기본값(왼쪽) 그대로였음
    }

    const blockStartIdx = valuesGrid.length;
    block.items.forEach(row => {
      const line = new Array(totalCols).fill('');
      const fmts = new Array(totalCols).fill('General');
      line[0] = row.code;
      line[1] = row.spec;
      line[2] = row.vendor;
      line[orderedQtyCol - 1] = row.orderedQty != null ? row.orderedQty : '';
      fmts[orderedQtyCol - 1] = numFmt;

      let sum = 0;
      dateKeys.forEach((key, i) => {
        const qty = row.dateMap[key];
        line[firstDateCol - 1 + i] = qty ? qty : '';
        fmts[firstDateCol - 1 + i] = numFmt;
        sum += qty || 0;
      });
      line[totalCol - 1] = sum || '';
      fmts[totalCol - 1] = numFmt;

      valuesGrid.push(line);
      numberFormatsGrid.push(fmts);
      // 신규/변경 품목(발주스냅샷 대비) — 합계·날짜별 배정값까지 포함해서 그 품목 행 전체를 강조
      const style = CHANGE_STYLE[row.changeType];
      fontWeightsGrid.push(new Array(totalCols).fill(style ? 'bold' : 'normal'));
      fontColorsGrid.push(new Array(totalCols).fill(style ? style.fg : null));
      backgroundsGrid.push(new Array(totalCols).fill(style ? style.bg : null));
      alignGrid.push(new Array(totalCols).fill('center'));
    });
    const blockEndIdx = valuesGrid.length - 1;

    if (block.vendor !== null && blockEndIdx >= blockStartIdx) {
      const subtotalLine = new Array(totalCols).fill('');
      const fmts = new Array(totalCols).fill('General');
      subtotalLine[2] = block.vendor + ' 소계';

      let subOrderedQty = 0;
      for (let i = blockStartIdx; i <= blockEndIdx; i++) subOrderedQty += Number(valuesGrid[i][orderedQtyCol - 1]) || 0;
      subtotalLine[orderedQtyCol - 1] = subOrderedQty;
      fmts[orderedQtyCol - 1] = numFmt;

      let subTotal = 0;
      for (let c = firstDateCol; c < firstDateCol + totalDays; c++) {
        let colSum = 0;
        for (let i = blockStartIdx; i <= blockEndIdx; i++) colSum += Number(valuesGrid[i][c - 1]) || 0;
        subtotalLine[c - 1] = colSum || '';
        fmts[c - 1] = numFmt;
        subTotal += colSum;
      }
      subtotalLine[totalCol - 1] = subTotal || '';
      fmts[totalCol - 1] = numFmt;

      valuesGrid.push(subtotalLine);
      numberFormatsGrid.push(fmts);
      fontWeightsGrid.push(new Array(totalCols).fill('bold'));
      fontColorsGrid.push(new Array(totalCols).fill(null));
      backgroundsGrid.push(new Array(totalCols).fill(null));
      alignGrid.push(new Array(totalCols).fill('center'));
    }
  });

  const hasData = rows.length > 0;

  // --- 총계 행 (전체 총합 — 원본 품목 데이터 행들만 더해야 하므로 rows 기준으로 직접 계산해서 값으로 씀) ---
  const totalLine = new Array(totalCols).fill('');
  const totalFmts = new Array(totalCols).fill('General');
  totalLine[2] = '총계';
  if (hasData) {
    dateKeys.forEach((key, i) => {
      const sum = rows.reduce((s, row) => s + (row.dateMap[key] || 0), 0);
      totalLine[firstDateCol - 1 + i] = sum || '';
      totalFmts[firstDateCol - 1 + i] = numFmt;
    });
    totalLine[orderedQtyCol - 1] = rows.reduce((s, row) => s + (row.orderedQty || 0), 0);
    totalFmts[orderedQtyCol - 1] = numFmt;
    totalLine[totalCol - 1] = rows.reduce((s, row) => {
      return s + dateKeys.reduce((s2, key) => s2 + (row.dateMap[key] || 0), 0);
    }, 0);
    totalFmts[totalCol - 1] = numFmt;
  }
  valuesGrid.push(totalLine);
  numberFormatsGrid.push(totalFmts);
  fontWeightsGrid.push(new Array(totalCols).fill('bold'));
  fontColorsGrid.push(new Array(totalCols).fill(null));
  backgroundsGrid.push(new Array(totalCols).fill(null));
  alignGrid.push(new Array(totalCols).fill('center'));

  const totalRow = dataStartRow + valuesGrid.length - 1;

  // --- 실적 반영 기준일(actualCutoffDateKey) 이전=확정(파랑), 이후=재배분(검정, 기본값 그대로) ---
  if (meta.actualCutoffDateKey) {
    for (let d = 1; d <= totalDays; d++) {
      const date = new Date(year, month - 1, d);
      const key = dateKey_(date);
      if (key > meta.actualCutoffDateKey) continue;
      const idx = firstDateCol - 1 + d - 1;
      for (let r = 0; r < fontColorsGrid.length; r++) fontColorsGrid[r][idx] = '#0070C0';
    }
  }

  // --- 한 번에 기록 ---
  if (valuesGrid.length > 0) {
    const dataRange = sheet.getRange(dataStartRow, 1, valuesGrid.length, totalCols);
    dataRange.setValues(valuesGrid);
    dataRange.setNumberFormats(numberFormatsGrid);
    dataRange.setFontWeights(fontWeightsGrid);
    dataRange.setFontColors(fontColorsGrid);
    dataRange.setBackgrounds(backgroundsGrid);
    dataRange.setHorizontalAlignments(alignGrid);
  }

  // --- 테두리 (기본 검정, 색상 없음) ---
  sheet.getRange(headerRow, 1, totalRow - headerRow + 1, totalCols)
    .setBorder(true, true, true, true, true, true);

  // --- 트럭 안내 문구 / 변경 색상 범례 / 검증사항 ---
  const truckNoteRow = totalRow + 2;
  sheet.getRange(truckNoteRow, 1).setValue(
    '🚚 업체별 하루 상한: 트럭 1대 ' + CONFIG.TRUCK_KG.toLocaleString() + 'kg (물량이 많으면 배송일을 늘림)'
  ).setFontWeight('bold');

  // 신규/변경 품목이 하나라도 있을 때만 범례를 보여준다 — 위 CHANGE_STYLE 색과 그대로 맞춰서, 표만
  // 훑어봐도 "이 색이 무슨 뜻인지" 바로 알 수 있게 한다(발표 데모에서 강조 포인트).
  const hasChangeHighlight = rows.some(r => r.changeType === 'new' || r.changeType === 'changed');
  let legendRowCount = 0;
  if (hasChangeHighlight) {
    legendRowCount = 1;
    const legendRow = truckNoteRow + 1;
    const legendCell = sheet.getRange(legendRow, 1);
    legendCell.setValue('🟩 신규 발주 품목      🟧 발주량 변경 품목').setFontWeight('bold');
  }

  const warningRow = truckNoteRow + legendRowCount + 1;
  if (issueCount > 0) {
    sheet.getRange(warningRow, 1).setValue('📋 검증사항 ' + issueCount + '건 (실행 로그 참고)').setFontWeight('bold');
  } else {
    sheet.getRange(warningRow, 1).setValue('✓ 검증사항 없음');
  }

  // 열 수가 많아서(메타 5열 + 날짜 최대 31열) 자동맞춤 대신 고정 폭으로 압축해 한 화면에 들어오게 함
  sheet.getRange(headerRow, 1, totalRow - headerRow + 1, totalCols).setFontSize(9);
  sheet.setColumnWidth(1, 85); // 품목코드
  sheet.setColumnWidth(2, 75); // 규격
  sheet.setColumnWidth(3, 65); // 구분(고객사)
  sheet.setColumnWidth(orderedQtyCol, 70); // 발주량
  sheet.setColumnWidth(totalCol, 65); // 합계
  sheet.setColumnWidths(firstDateCol, totalDays, 40); // 날짜 칸은 전부 좁게 통일
  sheet.setRowHeights(dataStartRow, totalRow - dataStartRow + 1, 20);

  sheet.setFrozenRows(headerRow);
  sheet.setFrozenColumns(totalCol);

  // 탭 순서 정리: "오늘 기준 이번 달" 계획 탭만 항상 1번째. 다른 달(미리 준비해둔 다음 달 등)은
  // 맨 뒤로 보낸다. "기본 시트"(초기화)는 이번 달 계획 탭이 있으면 그 바로 뒤(2번째)로, 없으면 1번째로.
  ss.setActiveSheet(sheet);
  if (isBaseSheet) {
    const currentMonthSheet = ss.getSheetByName(todayMonth + '월 출고계획');
    const hasCurrentMonthSheet = currentMonthSheet && currentMonthSheet.getSheetId() !== sheet.getSheetId();
    ss.moveActiveSheet(hasCurrentMonthSheet ? 2 : 1);
  } else if (isCurrentMonth) {
    ss.moveActiveSheet(1);
  } else {
    ss.moveActiveSheet(ss.getSheets().length); // 이번 달이 아니면 맨 뒤로
  }
}

/**
 * viewMode에 따라 [{ vendor: string|null, items: [row,...] }, ...] 형태로 표시 순서를 만든다.
 * - 'item': 배너 없이(vendor: null인 블록 하나) 규격순 정렬
 * - 'vendor': VENDOR_TYPE_MAP 고정 순서대로 업체별 블록(각자 품목코드순 정렬)
 */
function _buildDisplayBlocks_(rows, viewMode) {
  if (viewMode === 'vendor') {
    const byVendor = {};
    rows.forEach(row => {
      byVendor[row.vendor] = byVendor[row.vendor] || [];
      byVendor[row.vendor].push(row);
    });
    const vendorOrder = Object.keys(VENDOR_TYPE_MAP).filter(v => byVendor[v]);
    // VENDOR_TYPE_MAP에 없는 업체가 섞여 있으면(이론상 안 나와야 하지만 방어적으로) 뒤에 붙임
    Object.keys(byVendor).forEach(v => { if (vendorOrder.indexOf(v) === -1) vendorOrder.push(v); });

    return vendorOrder.map(vendor => ({
      vendor: vendor,
      items: byVendor[vendor].slice().sort((a, b) => a.code.localeCompare(b.code)),
    }));
  }

  // 'item' 보기: 배너 없이 규격순 정렬
  const sorted = rows.slice().sort((a, b) => (a.spec + a.code).localeCompare(b.spec + b.code));
  return [{ vendor: null, items: sorted }];
}

/**
 * 데모/시연용 초기화 — "오늘 기준 이번 달" 출고계획 탭만 지우고, 헤더·서식만 있는 "기본 시트"를
 * 1번째 탭으로 남긴다(발주스냅샷/발주지문은 안 건드림). 미리 준비해둔 다른 달(예: 다음 달) 탭은
 * 안 건드린다. 이 상태에서 이번 달 발주서 파일을 새로 만들거나 고치면, 자동 감지 트리거
 * (AutoTrigger.gs)나 "① 계획 생성"이 "이번 달 출고계획" 탭을 다시 만들어서 1번째로 올리고,
 * 기본 시트는 2번째로 밀려난다.
 */
function resetPlan() {
  const today = new Date();
  const year = today.getFullYear();
  const month = today.getMonth() + 1;
  const holidaySet = loadHolidays_(year);

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const currentMonthSheet = ss.getSheetByName(month + '월 출고계획');
  if (currentMonthSheet) ss.deleteSheet(currentMonthSheet);

  writePlanSheet(year, month, [], {
    holidaySet: holidaySet,
    issueCount: 0,
    viewMode: 'item',
    sheetName: CONFIG.PLAN_SHEET_NAME,
  });

  PropertiesService.getDocumentProperties().setProperty('LAST_PLAN', JSON.stringify({
    year: year,
    month: month,
    rows: [],
    issueCount: 0,
    holidayDates: Array.from(holidaySet),
    actualCutoffDateKey: null,
  }));

  appendExecutionLog_('resetPlan', '완료', '데모용 초기화 — 헤더만 남기고 데이터 지움');
  SpreadsheetApp.getActiveSpreadsheet().toast(
    '출고계획을 초기화했습니다(헤더만 남음). 발주서가 갱신되면 자동으로 다시 채워집니다.',
    '초기화 완료',
    8
  );
}

/**
 * "출고계획 보기" 메뉴 — 재계산 없이 마지막에 생성한 계획(LAST_PLAN)을 다른 배치로 다시 그린다.
 * generatePlan()이 매번 문서 속성(DocumentProperties)에 저장해둔 값을 읽어서 쓰므로,
 * 발주서를 다시 읽거나 트럭버킷 계산을 다시 돌리지 않는다.
 */
function showByVendor() {
  _rebuildPlanLayout_('vendor');
}

function showByItem() {
  _rebuildPlanLayout_('item');
}

function _rebuildPlanLayout_(viewMode) {
  const raw = PropertiesService.getDocumentProperties().getProperty('LAST_PLAN');
  if (!raw) {
    SpreadsheetApp.getActiveSpreadsheet().toast(
      '저장된 계획이 없습니다. 먼저 "② 계획 생성"을 한 번 실행하세요.', '안내', 6
    );
    return;
  }

  const saved = JSON.parse(raw);
  const holidaySet = new Set(saved.holidayDates || []);

  writePlanSheet(saved.year, saved.month, saved.rows, {
    holidaySet: holidaySet,
    issueCount: saved.issueCount || 0,
    viewMode: viewMode,
    actualCutoffDateKey: saved.actualCutoffDateKey || null,
  });

  SpreadsheetApp.getActiveSpreadsheet().toast(
    (viewMode === 'vendor' ? '회사별 보기로 전환' : '품목별 보기로 전환') + ' 완료', '완료', 5
  );
}
