/**
 * 발주 스냅샷(발주중량 탭) + 발주지문(발주지문 탭) — 출고계획 파일과 별도인
 * "출고계획_발주스냅샷" 스프레드시트로 저장한다. 위치는 REFERENCE_FOLDER_ID가 아니라
 * **지금 이 출고계획 스프레드시트가 들어있는 폴더**와 같은 곳 — 그래서 REFERENCE_FOLDER_ID
 * 설정 여부와 상관없이 항상 동작한다.
 *
 * 발주중량: 업체+품목코드+규격별 발주량을 저장해뒀다가, 다음 실행 때 지난 값과 다른 품목을
 *   찾아서(WritePlan.gs가) 빨간 글씨로 강조하는 데 쓴다.
 * 발주지문: 업체별 발주 내역 전체(코드+납기일+중량)를 MD5로 해싱해서 저장. 다음 실행 때 지문이
 *   전부 똑같으면(신규 업체도 없으면) "발주서가 하나도 안 바뀌었다"는 뜻이므로 재계산 자체를 스킵한다.
 */

const SNAPSHOT_FILE_NAME = '출고계획_발주스냅샷';
const SNAPSHOT_SHEET_NAME = '발주중량';
const FINGERPRINT_SHEET_NAME = '발주지문';

var _snapshotSpreadsheetCache_ = null;

/** 지금 활성 스프레드시트(출고계획)가 들어있는 드라이브 폴더를 반환한다(못 찾으면 내 드라이브 루트). */
function _getActiveSpreadsheetFolder_() {
  const file = DriveApp.getFileById(SpreadsheetApp.getActiveSpreadsheet().getId());
  const parents = file.getParents();
  return parents.hasNext() ? parents.next() : DriveApp.getRootFolder();
}

/** "출고계획_발주스냅샷" 파일을 찾아서 열거나(출고계획 시트와 같은 폴더 안), 없으면 새로 만든다. */
function _getSnapshotSpreadsheet_() {
  if (_snapshotSpreadsheetCache_) return _snapshotSpreadsheetCache_;

  const folder = _getActiveSpreadsheetFolder_();
  const existing = folder.getFilesByName(SNAPSHOT_FILE_NAME);

  let ss;
  if (existing.hasNext()) {
    ss = SpreadsheetApp.open(existing.next());
  } else {
    ss = SpreadsheetApp.create(SNAPSHOT_FILE_NAME);
    const file = DriveApp.getFileById(ss.getId());
    folder.addFile(file);
    DriveApp.getRootFolder().removeFile(file); // create()가 기본으로 넣어준 내 드라이브 루트에서는 뺌
  }

  _snapshotSpreadsheetCache_ = ss;
  return ss;
}

/** create()가 만들어준 빈 기본 탭("시트1"/"Sheet1")이 남아있으면 정리한다(내용 있는 탭은 안 건드림). */
function _cleanupSnapshotDefaultSheet_(ss) {
  ss.getSheets().forEach(sheet => {
    const name = sheet.getName();
    if (name === SNAPSHOT_SHEET_NAME || name === FINGERPRINT_SHEET_NAME) return;
    if (sheet.getLastRow() === 0 && sheet.getLastColumn() === 0) {
      ss.deleteSheet(sheet);
    }
  });
}

/** 업체별로 "코드_납기일(또는 NoDate)_중량" 서명을 모아 정렬 후 MD5 해싱한다. */
function buildOrderFingerprints_(orders) {
  const byVendor = {};
  orders.forEach(o => {
    byVendor[o.vendor] = byVendor[o.vendor] || [];
    byVendor[o.vendor].push(o);
  });

  const result = {};
  Object.keys(byVendor).forEach(vendor => {
    const rows = byVendor[vendor];
    const signatures = rows
      .map(o => o.code + '_' + (o.dueDate ? dateKey_(o.dueDate) : 'NoDate') + '_' + o.qty)
      .sort();
    result[vendor] = {
      rowCount: rows.length,
      totalWeight: rows.reduce((s, o) => s + o.qty, 0),
      hash: _md5Hex_(signatures.join('|')),
    };
  });
  return result;
}

/** 지난 실행 지문과 이번 지문이 업체 구성·값 전부 완전히 같은지 확인한다. */
function fingerprintsUnchanged_(prev, current) {
  const prevVendors = Object.keys(prev);
  const currentVendors = Object.keys(current);
  if (prevVendors.length === 0) return false; // 첫 실행(지문 없음)이면 무조건 재계산
  if (prevVendors.length !== currentVendors.length) return false;

  return currentVendors.every(vendor => {
    const p = prev[vendor];
    const c = current[vendor];
    if (!p) return false; // 신규 업체 등장
    return p.hash === c.hash && p.rowCount === c.rowCount && p.totalWeight === c.totalWeight;
  });
}

function loadPreviousFingerprints_() {
  const sheet = _getSnapshotSpreadsheet_().getSheetByName(FINGERPRINT_SHEET_NAME);
  if (!sheet) return {};

  const values = sheet.getDataRange().getValues();
  const result = {};
  values.slice(1).forEach(r => {
    if (!r[0]) return;
    result[r[0]] = { rowCount: r[1], totalWeight: r[2], hash: r[3] };
  });
  return result;
}

function saveFingerprints_(fingerprints) {
  const ss = _getSnapshotSpreadsheet_();
  // 삭제 후 재생성 대신 clear()로 내용만 지운다 — 이 탭이 파일에 남은 마지막 탭이면 삭제 자체가 안 됨
  const sheet = ss.getSheetByName(FINGERPRINT_SHEET_NAME) || ss.insertSheet(FINGERPRINT_SHEET_NAME);
  sheet.clear();

  sheet.appendRow(['구분', '행수', '중량합계', '코드지문(MD5)']);
  sheet.getRange(1, 1, 1, 4).setFontWeight('bold');

  Object.keys(fingerprints).forEach(vendor => {
    const f = fingerprints[vendor];
    sheet.appendRow([vendor, f.rowCount, f.totalWeight, f.hash]);
  });
  sheet.autoResizeColumns(1, 4);

  _cleanupSnapshotDefaultSheet_(ss);
}

/** 업체+품목코드+규격별 발주량(원본, 반올림 전) 스냅샷을 저장한다. byVendor는 _groupByVendor_() 결과. */
function saveOrderSnapshot_(byVendor) {
  const ss = _getSnapshotSpreadsheet_();
  // 삭제 후 재생성 대신 clear()로 내용만 지운다 — 이 탭이 파일에 남은 마지막 탭이면 삭제 자체가 안 됨
  const sheet = ss.getSheetByName(SNAPSHOT_SHEET_NAME) || ss.insertSheet(SNAPSHOT_SHEET_NAME);
  sheet.clear();

  sheet.appendRow(['구분', '품목코드', '규격', '발주중량']);
  sheet.getRange(1, 1, 1, 4).setFontWeight('bold');

  Object.keys(byVendor).forEach(vendor => {
    byVendor[vendor].forEach(it => {
      sheet.appendRow([vendor, it.code, it.spec, it.qty]);
    });
  });
  sheet.autoResizeColumns(1, 4);
}

/** {vendor|code|spec: 발주중량} 맵으로 지난 스냅샷을 읽어온다. 파일/탭이 없으면(첫 실행) 빈 객체. */
function loadPreviousSnapshot_() {
  const sheet = _getSnapshotSpreadsheet_().getSheetByName(SNAPSHOT_SHEET_NAME);
  if (!sheet) return {};

  const values = sheet.getDataRange().getValues();
  const result = {};
  values.slice(1).forEach(r => {
    if (!r[0]) return;
    result[r[0] + '|' + r[1] + '|' + r[2]] = Number(r[3]) || 0;
  });
  return result;
}

/** 발주지문 탭을 지워서, 발주가 그대로여도 다음 generatePlan() 실행을 강제로 다시 돌게 만든다. */
function forceRegeneratePlan() {
  const ss = _getSnapshotSpreadsheet_();
  const sheet = ss.getSheetByName(FINGERPRINT_SHEET_NAME);
  if (sheet) ss.deleteSheet(sheet);
  generatePlan();
}

function _md5Hex_(text) {
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, text, Utilities.Charset.UTF_8);
  return bytes.map(b => ((b < 0 ? b + 256 : b).toString(16)).padStart(2, '0')).join('');
}
