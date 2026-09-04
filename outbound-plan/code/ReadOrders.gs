/**
 * 발주서를 읽어서 { orders: [{업체, 품목코드, 규격, 수량, 납기일}], issues: [{vendor, message}] }를
 * 반환한다.
 *
 * CONFIG.ORDER_FOLDER_ID가 있으면 그 드라이브 폴더에서 'YYYY-MM 발주서' 파일을 찾아 읽고,
 * 없으면(테스트 모드) 이 스프레드시트 자체 안의 "고객사 A/B/C" 탭을 발주서로 간주한다.
 *
 * 한 발주서 파일 안에 업체별로 시트가 나뉘어 있고, 시트명이 '고객사 X' 패턴에 매치되면 그 알파벳을
 * "고객사X"로 매핑해서 업체를 판별한다. 각 시트는 1행에서 CONFIG.DEEP_DIVE_MARKER('딥다이브') 텍스트를
 * 찾아 그 열부터 표가 시작한다고 보고, 2행 헤더 텍스트로 품목코드/규격/중량/납기일 컬럼을 인식한다.
 *
 * "고객사 X" 패턴은 맞는데 마커나 필수 컬럼을 못 찾으면, 그 업체 발주 전체가 조용히 빠지는 대신
 * issues에 남겨서 경고로 노출한다(sheet_columns_missing) — 그래야 시트 서식이 깨져서 업체 하나가
 * 통째로 누락돼도 아무 표시 없이 지나가는 일이 없다.
 */
function loadOrderStatus() {
  const ss = _resolveOrderSpreadsheet_();
  const orders = [];
  const issues = [];

  ss.getSheets().forEach(sheet => {
    const vendor = _detectVendorFromSheetName_(sheet.getName());
    if (!vendor) return; // '고객사 X' 패턴이 아니면 발주서 탭이 아니라고 보고 스킵

    const result = _readOrderSheet_(sheet, vendor);
    orders.push.apply(orders, result.orders);
    issues.push.apply(issues, result.issues);
  });

  return { orders: orders, issues: issues };
}

function _detectVendorFromSheetName_(sheetName) {
  // "고객사 A 원본"처럼 원본/참고용으로 같이 남겨둔 탭은 변환된 표가 아니므로 발주서 탭 대상에서
  // 제외한다 — 안 그러면 "고객사 X" 패턴에 걸린다는 이유만으로 이 탭까지 읽으려다 딥다이브 마커를
  // 못 찾아 sheet_columns_missing 경고가 잘못 뜬다.
  if (sheetName.indexOf('원본') !== -1) return null;
  const m = sheetName.match(/고객사\s*([A-Za-z])/);
  return m ? '고객사' + m[1].toUpperCase() : null;
}

/** ORDER_FOLDER_ID 폴더에서 'YYYY-MM 발주서'로 시작하는 파일 중 가장 최근에 수정된 것을 찾는다.
 *  테스트 모드(ORDER_FOLDER_ID 없음)면 null. AutoTrigger.gs가 파일 변경 감지용으로도 재사용한다. */
function _findLatestOrderFile_() {
  if (!CONFIG.ORDER_FOLDER_ID) return null;

  const folder = DriveApp.getFolderById(CONFIG.ORDER_FOLDER_ID);
  const it = folder.getFiles();
  let best = null;
  while (it.hasNext()) {
    const f = it.next();
    if (!/^\d{4}-\d{2}\s*발주서/.test(f.getName())) continue;
    if (!best || f.getLastUpdated() > best.getLastUpdated()) best = f;
  }
  return best;
}

/**
 * 발주서 파일명("2026-10 발주서...")에서 연/월을 뽑는다 — 그 발주서가 몇 월치인지는 오늘 날짜가
 * 아니라 파일명에 적힌 값을 기준으로 삼아야, 다음 달 발주서를 미리 올려놔도 그 달 계획으로 잡힌다.
 * 테스트 모드거나 파일을 못 찾으면(파일명에서 못 뽑으면) 오늘 날짜로 폴백한다.
 */
function resolveOrderFileYearMonth_() {
  const file = _findLatestOrderFile_();
  if (file) {
    const m = file.getName().match(/^(\d{4})-(\d{2})\s*발주서/);
    if (m) return { year: Number(m[1]), month: Number(m[2]) };
  }
  const today = new Date();
  return { year: today.getFullYear(), month: today.getMonth() + 1 };
}

function _resolveOrderSpreadsheet_() {
  if (!CONFIG.ORDER_FOLDER_ID) {
    // 테스트 모드: 이 스프레드시트 자체 안의 "고객사 A/B/C" 탭을 발주서로 취급
    return SpreadsheetApp.getActiveSpreadsheet();
  }

  const best = _findLatestOrderFile_();
  if (!best) {
    throw new Error("ORDER_FOLDER_ID 폴더에서 'YYYY-MM 발주서' 파일을 못 찾았습니다.");
  }

  if (best.getMimeType() === MimeType.GOOGLE_SHEETS) {
    return SpreadsheetApp.open(best);
  }

  // xlsx는 구글시트 변환 사본을 만들어서 읽는다(고급 Drive 서비스 필요) — 이 변환 자체가 몇 초씩
  // 걸리는데, 자동 감지 트리거는 발주서가 바뀔 때만 generatePlan()을 부르므로 그때마다 매번 새로
  // 변환해도 원래 큰 낭비는 아니지만, 수동으로 여러 번 다시 실행할 때도 매번 재변환하던 걸 참고파일과
  // 같은 방식으로 캐시해서 없앤다 — 발주서 파일이 실제로 안 바뀌었으면(같은 파일ID·같은 수정시각)
  // 재변환 없이 지난번 변환 사본을 그대로 재사용한다. 부모 폴더를 안 정해주면 내 드라이브 루트에
  // 생기므로 ORDER_FOLDER_ID 폴더 안에 만들어지도록 명시한다. Drive API 고급 서비스는 v3가 붙으면
  // Files.insert가 없고 Files.create로 바뀌므로 둘 다 지원한다.
  const props = PropertiesService.getScriptProperties();
  const cacheProp = 'ORDER_FILE_CACHE';
  const sourceUpdatedAt = String(best.getLastUpdated().getTime());
  const cached = _readJsonProp_(props, cacheProp);

  if (cached && cached.sourceId === best.getId() && cached.sourceUpdatedAt === sourceUpdatedAt) {
    try {
      return SpreadsheetApp.openById(cached.tempFileId);
    } catch (e) {
      // 캐시된 사본이 지워졌거나 접근 불가 — 아래에서 새로 변환
    }
  }

  if (cached && cached.tempFileId) {
    try { DriveApp.getFileById(cached.tempFileId).setTrashed(true); } catch (e) { /* 이미 지워졌으면 무시 */ }
  }
  const tempFileId = Drive.Files.create
    ? Drive.Files.create({ name: '__cache_' + best.getName(), mimeType: MimeType.GOOGLE_SHEETS, parents: [CONFIG.ORDER_FOLDER_ID] }, best.getBlob()).id // v3
    : Drive.Files.insert({ title: '__cache_' + best.getName(), mimeType: MimeType.GOOGLE_SHEETS, parents: [{ id: CONFIG.ORDER_FOLDER_ID }] }, best.getBlob(), { convert: true }).id; // v2
  props.setProperty(cacheProp, JSON.stringify({ sourceId: best.getId(), sourceUpdatedAt: sourceUpdatedAt, tempFileId: tempFileId }));
  return SpreadsheetApp.openById(tempFileId);
}

/** generatePlan() 끝에서 호출됨 — 발주서 변환 사본은 이제 계속 캐시로 재사용하므로 더 지울 게 없다. */
function cleanupOrderFile_() {
  // no-op: getReferenceSpreadsheet_와 같은 이유로 캐시를 유지한다(ReferenceFiles.gs 참고)
}

function _readOrderSheet_(sheet, vendor) {
  const values = sheet.getDataRange().getValues();
  if (values.length < 3) {
    return { orders: [], issues: [{ vendor: vendor, message: '시트에 데이터가 3행 미만이라 발주 내역이 없는 것으로 처리됨(sheet_columns_missing)' }] };
  }

  const markerRow = values[0];
  // 마커 셀이 "딥다이브" 하나만 딱 들어있지 않고 다른 글자와 같이 있어도 인식되게, 완전 일치
  // 대신 "포함" 여부로 확인한다(예: "딥다이브(변환결과)"처럼 부가 텍스트가 붙어있는 경우 대비).
  const startCol = markerRow.findIndex(v => String(v).indexOf(CONFIG.DEEP_DIVE_MARKER) !== -1);
  if (startCol === -1) {
    // 마커 없으면 발주서 표가 아니라고 보고 스킵하지만, "고객사 X" 탭인데 마커가 없다는 건
    // 서식이 깨졌을 가능성이 높으므로 조용히 넘기지 않고 경고를 남긴다. 실제로 1행에 뭐가
    // 들어있었는지도 같이 남겨서, 원인(다른 파일이 읽혔는지/마커 표기가 다른지)을 바로 알 수 있게 한다.
    return { orders: [], issues: [{ vendor: vendor, message: '1행에서 "' + CONFIG.DEEP_DIVE_MARKER + '" 마커를 못 찾아 이 업체 발주 전체가 제외됨(sheet_columns_missing) — 실제 1행: ' + JSON.stringify(markerRow) }] };
  }

  const headerRow = values[1].slice(startCol);
  const col = {
    code: findColumnIndex_(headerRow, ORDER_COLUMN_KEYWORDS.code),
    spec: findColumnIndex_(headerRow, ORDER_COLUMN_KEYWORDS.spec),
    weight: findColumnIndex_(headerRow, ORDER_COLUMN_KEYWORDS.weight),
    date: findColumnIndex_(headerRow, ORDER_COLUMN_KEYWORDS.date),
  };
  if (col.code === -1 || col.weight === -1) {
    return { orders: [], issues: [{ vendor: vendor, message: '필수 컬럼(품목코드/중량)을 못 찾아 이 업체 발주 전체가 제외됨(sheet_columns_missing)' }] };
  }

  const orders = [];
  values.slice(2).forEach(row => {
    const cells = row.slice(startCol);
    const codeCell = cells[col.code];
    if (!codeCell) return;

    const qty = Number(cells[col.weight]) || 0;
    if (!qty) return;

    const dateCell = col.date !== -1 ? cells[col.date] : null;
    const dueDate = dateCell ? (dateCell instanceof Date ? dateCell : new Date(dateCell)) : null;

    orders.push({
      vendor: vendor,
      code: String(codeCell).trim(),
      spec: col.spec !== -1 && cells[col.spec] ? String(cells[col.spec]).trim() : '',
      qty: qty,
      dueDate: dueDate,
    });
  });

  return { orders: orders, issues: [] };
}
