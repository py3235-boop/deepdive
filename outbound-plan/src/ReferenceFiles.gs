/**
 * 참고파일(공휴일.xlsx, 기준정보.xlsx, 제품 중량.xlsx 등)을 CONFIG.REFERENCE_FOLDER_ID 폴더(구글
 * 워크스페이스 공유문서함)에서 찾아 읽기 위한 헬퍼.
 *
 * xlsx 파일은 SpreadsheetApp이 바로 못 여니까, 임시로 구글시트 사본으로 변환해서 연 뒤 실행이 끝나면
 * (cleanupReferenceFiles_) 지운다. 이미 구글시트로 올라가 있는 파일이면 변환 없이 그냥 연다.
 *
 * 사전 준비(한 번만):
 *  1) Apps Script 편집기 왼쪽 "서비스" + 버튼 → "Drive API"(고급 서비스) 추가
 *  2) 공휴일/기준정보/제품 중량 파일이 있는 공유문서함 폴더의 ID를 스크립트 속성
 *     REFERENCE_FOLDER_ID 에 등록 (폴더를 열었을 때 주소창 .../folders/{여기}가 폴더 ID)
 */

var _refCache_ = {};

function findReferenceFile_(keyword) {
  if (!CONFIG.REFERENCE_FOLDER_ID) {
    throw new Error('REFERENCE_FOLDER_ID가 설정되지 않았습니다. 스크립트 속성에 등록하세요.');
  }
  const folder = DriveApp.getFolderById(CONFIG.REFERENCE_FOLDER_ID);
  const it = folder.getFiles();
  let best = null;
  while (it.hasNext()) {
    const f = it.next();
    if (f.getName().indexOf(keyword) === -1) continue;
    if (!best || f.getLastUpdated() > best.getLastUpdated()) best = f;
  }
  if (!best) {
    throw new Error("'" + keyword + "'가 포함된 참고파일을 REFERENCE_FOLDER_ID 폴더에서 못 찾았습니다.");
  }
  return best;
}

/** keyword로 찾은 참고파일을 스프레드시트로 열어서 반환한다(캐시됨, 같은 실행 중 재사용). */
function getReferenceSpreadsheet_(keyword) {
  if (_refCache_[keyword]) return _refCache_[keyword].ss;

  const file = findReferenceFile_(keyword);
  let ss, tempFileId = null;

  if (file.getMimeType() === MimeType.GOOGLE_SHEETS) {
    ss = SpreadsheetApp.open(file);
  } else {
    // xlsx → 구글시트 변환 사본을 임시로 만든다(고급 Drive 서비스 필요).
    // Drive API 고급 서비스는 v3가 붙으면 Files.insert가 없고 Files.create로 바뀌므로 둘 다 지원한다.
    tempFileId = Drive.Files.create
      ? Drive.Files.create({ name: '__tmp_' + file.getName(), mimeType: MimeType.GOOGLE_SHEETS }, file.getBlob()).id // v3
      : Drive.Files.insert({ title: '__tmp_' + file.getName(), mimeType: MimeType.GOOGLE_SHEETS }, file.getBlob(), { convert: true }).id; // v2
    ss = SpreadsheetApp.openById(tempFileId);
  }

  _refCache_[keyword] = { ss: ss, tempFileId: tempFileId };
  return ss;
}

/** generatePlan() 끝에서 한 번 호출 — 변환 과정에서 생긴 임시 구글시트 사본을 휴지통으로 보낸다. */
function cleanupReferenceFiles_() {
  Object.keys(_refCache_).forEach(keyword => {
    const entry = _refCache_[keyword];
    if (entry.tempFileId) {
      try {
        DriveApp.getFileById(entry.tempFileId).setTrashed(true);
      } catch (e) {
        // 이미 지워졌거나 접근 불가하면 무시
      }
    }
  });
  _refCache_ = {};
}

/** 헤더 배열에서 keywords 중 하나라도 부분일치하는 첫 번째 열 인덱스를 찾는다. 못 찾으면 -1. */
function findColumnIndex_(header, keywords) {
  for (let i = 0; i < header.length; i++) {
    const h = String(header[i]).trim().toLowerCase();
    if (keywords.some(k => h.indexOf(k.toLowerCase()) !== -1)) return i;
  }
  return -1;
}
