/**
 * 출고계획 결과를 시트에 쓴다.
 * rows: [{ code, spec, vendor, orderedQty(발주량), dateMap: {'yyyy-MM-dd': qty} }, ...]
 * meta: { holidaySet(Set), issueCount(검증 경고 건수), viewMode('item'|'vendor') }
 *
 * viewMode:
 *  - 'item'(기본): 배너 없이 규격순 정렬
 *  - 'vendor': 업체 순서(VENDOR_TYPE_MAP 고정 순서)대로 "◆ 업체명" 배너 + 그 업체 소계 행을 넣어서 보여줌
 *
 * 색상/배경 등 디자인은 일단 다 뺐고, 굵게/정렬/숫자서식/틀고정 같은 구조적인 것만 남겼다.
 *
 * CONFIG.PLAN_SHEET_NAME("출고계획")에 그대로 덮어쓴다 — 유일하게 남은 탭일 수 있어서 삭제 후
 * 재생성이 안 되므로(구글시트는 마지막 탭 삭제 불가), 삭제/재생성 대신 clear()로 내용만 지우고
 * 같은 탭에 다시 쓴다. 이름이 그새 바뀌어서 못 찾으면 첫 번째 탭을 그대로 쓴다.
 *
 * 파일 이름 자체도 매번 "N월_출고계획(통합)"으로 맞춘다 — 뒤 파트가 폴더명이 아니라 이 파일명
 * 패턴으로 출고계획 파일을 찾기 때문에, 시트 안 제목 셀만으로는 안 되고 파일명이 실제로 이래야 한다.
 */

function writePlanSheet(year, month, rows, meta) {
  meta = meta || {};
  const issueCount = meta.issueCount || 0;
  const viewMode = meta.viewMode === 'vendor' ? 'vendor' : 'item';
  const holidaySet = meta.holidaySet || new Set();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = CONFIG.PLAN_SHEET_NAME;

  const fileName = month + '월_출고계획(통합)';
  if (ss.getName() !== fileName) ss.rename(fileName);

  const sheet = ss.getSheetByName(sheetName) || ss.getSheets()[0];
  sheet.clear();
  sheet.getRange(1, 1).breakApart(); // 이전 실행에서 남아있을 수 있는 병합 셀 해제
  sheet.setFrozenRows(0);
  sheet.setFrozenColumns(0); // 고정 상태로는 열 병합이 막히므로 병합 전에 미리 풀어둔다
  if (sheet.getName() !== sheetName) sheet.setName(sheetName);

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

  // 주말/공휴일만 색상 표시(일요일·공휴일=빨강, 토요일=파랑) — 그 외 디자인은 색 없이 유지
  for (let d = 1; d <= totalDays; d++) {
    const date = new Date(year, month - 1, d);
    const day = date.getDay();
    const col = firstDateCol + d - 1;
    if (day === 0 || holidaySet.has(dateKey_(date))) {
      sheet.getRange(headerRow, col).setBackground('#C0392B').setFontColor('#FFFFFF');
    } else if (day === 6) {
      sheet.getRange(headerRow, col).setBackground('#2980B9').setFontColor('#FFFFFF');
    }
  }

  // --- 표시 순서 구성: 'item'은 규격순 한 덩어리, 'vendor'는 업체별 배너+소계로 분리 ---
  const blocks = _buildDisplayBlocks_(rows, viewMode); // [{ vendor|null, items: [row,...] }, ...]

  let cursor = dataStartRow;
  blocks.forEach(block => {
    if (block.vendor !== null) {
      const bannerLine = new Array(totalCols).fill('');
      bannerLine[0] = '◆ ' + block.vendor;
      sheet.getRange(cursor, 1, 1, totalCols).setValues([bannerLine]).setFontWeight('bold');
      cursor++;
    }

    const blockStartRow = cursor;
    block.items.forEach(row => {
      const line = new Array(totalCols).fill('');
      line[0] = row.code;
      line[1] = row.spec;
      line[2] = row.vendor;
      line[orderedQtyCol - 1] = row.orderedQty != null ? row.orderedQty : '';
      dateKeys.forEach((key, i) => {
        const qty = row.dateMap[key];
        line[firstDateCol - 1 + i] = qty ? qty : '';
      });

      sheet.getRange(cursor, 1, 1, totalCols).setValues([line]).setHorizontalAlignment('center');
      sheet.getRange(cursor, orderedQtyCol).setNumberFormat('#,##0');
      sheet.getRange(cursor, firstDateCol, 1, totalDays).setNumberFormat('#,##0');
      sheet.getRange(cursor, totalCol).setNumberFormat('#,##0').setFormula(
        '=SUM(' + sheet.getRange(cursor, firstDateCol, 1, totalDays).getA1Notation() + ')'
      );
      // 신규/변경 발주량(발주스냅샷 대비) — 합계·날짜별 배정값까지 포함해서 그 품목 행 전체를 강조
      if (row.changed) {
        sheet.getRange(cursor, 1, 1, totalCols).setFontColor('#C0392B').setFontWeight('bold');
      }
      cursor++;
    });
    const blockEndRow = cursor - 1;

    if (block.vendor !== null && blockEndRow >= blockStartRow) {
      const subtotalLine = new Array(totalCols).fill('');
      subtotalLine[2] = block.vendor + ' 소계';
      sheet.getRange(cursor, 1, 1, totalCols).setValues([subtotalLine]);
      sheet.getRange(cursor, orderedQtyCol).setFormula(
        '=SUM(' + sheet.getRange(blockStartRow, orderedQtyCol, blockEndRow - blockStartRow + 1, 1).getA1Notation() + ')'
      );
      sheet.getRange(cursor, totalCol).setFormula(
        '=SUM(' + sheet.getRange(blockStartRow, totalCol, blockEndRow - blockStartRow + 1, 1).getA1Notation() + ')'
      );
      for (let c = firstDateCol; c < firstDateCol + totalDays; c++) {
        sheet.getRange(cursor, c).setFormula(
          '=SUM(' + sheet.getRange(blockStartRow, c, blockEndRow - blockStartRow + 1, 1).getA1Notation() + ')'
        );
      }
      sheet.getRange(cursor, firstDateCol, 1, totalDays).setNumberFormat('#,##0');
      sheet.getRange(cursor, orderedQtyCol).setNumberFormat('#,##0');
      sheet.getRange(cursor, totalCol).setNumberFormat('#,##0');
      sheet.getRange(cursor, 1, 1, totalCols).setFontWeight('bold').setHorizontalAlignment('center');
      cursor++;
    }
  });

  const hasData = rows.length > 0;

  // --- 총계 행 (전체 총합 — 원본 품목 데이터 행들만 더해야 하므로 rows 기준으로 직접 계산해서 값으로 씀) ---
  const totalRow = cursor;
  const totalLine = new Array(totalCols).fill('');
  totalLine[2] = '총계';
  if (hasData) {
    dateKeys.forEach((key, i) => {
      const sum = rows.reduce((s, row) => s + (row.dateMap[key] || 0), 0);
      totalLine[firstDateCol - 1 + i] = sum || '';
    });
    totalLine[orderedQtyCol - 1] = rows.reduce((s, row) => s + (row.orderedQty || 0), 0);
    totalLine[totalCol - 1] = rows.reduce((s, row) => {
      return s + dateKeys.reduce((s2, key) => s2 + (row.dateMap[key] || 0), 0);
    }, 0);
  }
  sheet.getRange(totalRow, 1, 1, totalCols).setValues([totalLine]);
  if (hasData) {
    sheet.getRange(totalRow, firstDateCol, 1, totalDays).setNumberFormat('#,##0');
    sheet.getRange(totalRow, orderedQtyCol).setNumberFormat('#,##0');
    sheet.getRange(totalRow, totalCol).setNumberFormat('#,##0');
  }
  sheet.getRange(totalRow, 1, 1, totalCols).setFontWeight('bold').setHorizontalAlignment('center');

  // --- 실적 반영 기준일(actualCutoffDateKey) 이전=확정(파랑), 이후=재배분(검정, 기본값 그대로) ---
  if (meta.actualCutoffDateKey) {
    for (let d = 1; d <= totalDays; d++) {
      const date = new Date(year, month - 1, d);
      const key = dateKey_(date);
      if (key > meta.actualCutoffDateKey) continue;
      const col = firstDateCol + d - 1;
      sheet.getRange(dataStartRow, col, totalRow - dataStartRow + 1, 1).setFontColor('#0070C0');
    }
  }

  // --- 테두리 (기본 검정, 색상 없음) ---
  sheet.getRange(headerRow, 1, totalRow - headerRow + 1, totalCols)
    .setBorder(true, true, true, true, true, true);

  // --- 트럭 안내 문구 / 검증 경고 ---
  const truckNoteRow = totalRow + 2;
  sheet.getRange(truckNoteRow, 1).setValue(
    '🚚 업체별 하루 상한: 트럭 1대 ' + CONFIG.TRUCK_KG.toLocaleString() + 'kg (물량이 많으면 배송일을 늘림)'
  ).setFontWeight('bold');

  const warningRow = totalRow + 3;
  if (issueCount > 0) {
    sheet.getRange(warningRow, 1).setValue('⚠ 검증 경고 ' + issueCount + '건 (실행 로그 참고)').setFontWeight('bold');
  } else {
    sheet.getRange(warningRow, 1).setValue('✓ 검증 경고 없음');
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
