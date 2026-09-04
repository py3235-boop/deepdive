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
  const m = sheetName.match(/고객사\s*([A-Za-z])/);
  return m ? '고객사' + m[1].toUpperCase() : null;
}

var _orderTempFileId_ = null;

function _resolveOrderSpreadsheet_() {
  if (!CONFIG.ORDER_FOLDER_ID) {
    // 테스트 모드: 이 스프레드시트 자체 안의 "고객사 A/B/C" 탭을 발주서로 취급
    return SpreadsheetApp.getActiveSpreadsheet();
  }

  const folder = DriveApp.getFolderById(CONFIG.ORDER_FOLDER_ID);
  const it = folder.getFiles();
  let best = null;
  while (it.hasNext()) {
    const f = it.next();
    if (!/^\d{4}-\d{2}\s*발주서/.test(f.getName())) continue;
    if (!best || f.getLastUpdated() > best.getLastUpdated()) best = f;
  }
  if (!best) {
    throw new Error("ORDER_FOLDER_ID 폴더에서 'YYYY-MM 발주서' 파일을 못 찾았습니다.");
  }

  if (best.getMimeType() === MimeType.GOOGLE_SHEETS) {
    return SpreadsheetApp.open(best);
  }

  // xlsx 등은 구글시트 변환 사본을 임시로 만들어서 읽는다(고급 Drive 서비스 필요).
  // Drive API 고급 서비스는 v3가 붙으면 Files.insert가 없고 Files.create로 바뀌므로 둘 다 지원한다.
  const tempFileId = Drive.Files.create
    ? Drive.Files.create({ name: '__tmp_' + best.getName(), mimeType: MimeType.GOOGLE_SHEETS }, best.getBlob()).id // v3
    : Drive.Files.insert({ title: '__tmp_' + best.getName(), mimeType: MimeType.GOOGLE_SHEETS }, best.getBlob(), { convert: true }).id; // v2
  _orderTempFileId_ = tempFileId;
  return SpreadsheetApp.openById(tempFileId);
}

/** generatePlan() 끝에서 호출 — 발주서 xlsx 변환 과정에서 생긴 임시 사본을 정리한다. */
function cleanupOrderFile_() {
  if (_orderTempFileId_) {
    try {
      DriveApp.getFileById(_orderTempFileId_).setTrashed(true);
    } catch (e) {
      // 이미 지워졌거나 접근 불가하면 무시
    }
    _orderTempFileId_ = null;
  }
}

function _readOrderSheet_(sheet, vendor) {
  const values = sheet.getDataRange().getValues();
  if (values.length < 3) {
    return { orders: [], issues: [{ vendor: vendor, message: '시트에 데이터가 3행 미만이라 발주 내역이 없는 것으로 처리됨(sheet_columns_missing)' }] };
  }

  const markerRow = values[0];
  const startCol = markerRow.findIndex(v => String(v).trim() === CONFIG.DEEP_DIVE_MARKER);
  if (startCol === -1) {
    // 마커 없으면 발주서 표가 아니라고 보고 스킵하지만, "고객사 X" 탭인데 마커가 없다는 건
    // 서식이 깨졌을 가능성이 높으므로 조용히 넘기지 않고 경고를 남긴다.
    return { orders: [], issues: [{ vendor: vendor, message: '1행에서 "' + CONFIG.DEEP_DIVE_MARKER + '" 마커를 못 찾아 이 업체 발주 전체가 제외됨(sheet_columns_missing)' }] };
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
