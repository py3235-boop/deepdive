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

/**
 * 배송/근무일 계산에 쓰는 휴일 집합.
 *
 * 공휴일.xlsx(공개 공휴일, '월'/'일' 반복 표기 — 연도 없음) + 기준정보.xlsx의 '휴무' 탭(회사 자체
 * 휴무, 정확한 날짜 있음)을 합쳐서 하나의 휴일 Set으로 쓴다. 원래는 성격이 다른 값이지만, 이번
 * 버전에서는 "이 날은 배송/생산 근무일이 아니다"라는 같은 목적으로만 쓰기 때문에 하나로 합쳤다.
 * REFERENCE_FOLDER_ID가 없으면(참고파일 미연동) 빈 Set을 반환해서 공휴일 없이 그냥 진행한다.
 */
function loadHolidays_(year) {
  const set = new Set();
  if (!CONFIG.REFERENCE_FOLDER_ID) return set;

  try {
    const ss = getReferenceSpreadsheet_('공휴일');
    const sheet = ss.getSheets()[0];
    const values = sheet.getDataRange().getValues();
    const header = values[0].map(h => String(h).trim());
    const mIdx = header.indexOf('월');
    const dIdx = header.indexOf('일');
    if (mIdx !== -1 && dIdx !== -1) {
      values.slice(1).forEach(r => {
        if (r[mIdx] === '' || r[mIdx] == null) return;
        const date = new Date(year, Number(r[mIdx]) - 1, Number(r[dIdx]));
        set.add(dateKey_(date));
      });
    }
  } catch (e) {
    Logger.log('공휴일.xlsx 로드 실패(무시하고 진행): ' + e.message);
  }

  try {
    const ss = getReferenceSpreadsheet_('기준정보');
    const sheet = ss.getSheetByName('휴무');
    if (sheet) {
      const values = sheet.getDataRange().getValues();
      const header = values[0].map(h => String(h).trim());
      const dateIdx = header.indexOf('날짜');
      if (dateIdx !== -1) {
        values.slice(1).forEach(r => {
          const v = r[dateIdx];
          if (!v) return;
          const date = v instanceof Date ? v : new Date(v);
          if (date.getFullYear() === year) set.add(dateKey_(date));
        });
      }
    }
  } catch (e) {
    Logger.log("기준정보.xlsx '휴무' 탭 로드 실패(무시하고 진행): " + e.message);
  }

  return set;
}

/**
 * '제품 중량' 참고파일에서 품목코드/규격별 단위중량(보빈 1개 기준 등)을 읽어온다.
 * 코드로 못 찾으면 규격 문자열로 한 번 더 시도한다(`lookupUnitWeight_`).
 */
function loadUnitWeights_() {
  const weights = { byCode: {}, bySpec: {} };
  if (!CONFIG.REFERENCE_FOLDER_ID) return weights;

  try {
    const ss = getReferenceSpreadsheet_('제품 중량');
    const sheet = ss.getSheets()[0];
    const values = sheet.getDataRange().getValues();
    const header = values[0].map(h => String(h).trim());

    const codeIdx = findColumnIndex_(header, ['품목코드', '제품코드', 'sc코드']);
    const specIdx = findColumnIndex_(header, ['규격', '사양']);
    const weightIdx = findColumnIndex_(header, ['평균중량', '단위중량', '중량']);
    if (weightIdx === -1) return weights;

    values.slice(1).forEach(r => {
      const w = Number(r[weightIdx]);
      if (!w) return;
      if (codeIdx !== -1 && r[codeIdx]) {
        const code = String(r[codeIdx]).trim();
        weights.byCode[code] = w;
        const stripped = code.replace(/^[A-Za-z]+/, '');
        if (stripped && stripped !== code) weights.byCode[stripped] = w;
      }
      if (specIdx !== -1 && r[specIdx]) {
        weights.bySpec[String(r[specIdx]).trim()] = w;
      }
    });
  } catch (e) {
    Logger.log('제품 중량 참고파일 로드 실패(단위중량 보정 없이 진행): ' + e.message);
  }

  return weights;
}

function lookupUnitWeight_(weights, code, spec) {
  if (weights.byCode[code] != null) return weights.byCode[code];
  const stripped = String(code).replace(/^[A-Za-z]+/, '');
  if (weights.byCode[stripped] != null) return weights.byCode[stripped];
  if (spec && weights.bySpec[spec] != null) return weights.bySpec[spec];
  return null;
}
