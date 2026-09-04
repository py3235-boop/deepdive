/**
 * 참고파일(공휴일.xlsx, 기준정보.xlsx, 제품 중량.xlsx 등)을 CONFIG.REFERENCE_FOLDER_ID 폴더(구글
 * 워크스페이스 공유문서함)에서 찾아 읽기 위한 헬퍼.
 *
 * xlsx 파일은 SpreadsheetApp이 바로 못 여니까, 임시로 구글시트 사본으로 변환해서 연다. 이 변환
 * 자체가 몇 초~십수 초씩 걸려서, 예전처럼 실행마다(매번 트리거가 돌 때마다) 새로 변환하고 끝에
 * 지우기를 반복하면 참고파일이 안 바뀌었어도 매번 그 비용을 그대로 문다 — 실행이 느려지는 가장 큰
 * 원인이었다. 그래서 지금은 변환한 사본을 스크립트 속성에 캐시해두고, 원본 파일이 실제로 바뀌었을
 * 때(파일 ID나 수정시각이 달라졌을 때)만 다시 변환한다. 이미 구글시트로 올라가 있는 파일이면 원래도
 * 변환이 필요 없다.
 *
 * 사전 준비(한 번만):
 *  1) Apps Script 편집기 왼쪽 "서비스" + 버튼 → "Drive API"(고급 서비스) 추가
 *  2) 공휴일/기준정보/제품 중량 파일이 있는 공유문서함 폴더의 ID를 스크립트 속성
 *     REFERENCE_FOLDER_ID 에 등록 (폴더를 열었을 때 주소창 .../folders/{여기}가 폴더 ID)
 */

var _refCache_ = {}; // 이번 실행 안에서만 재사용(같은 실행 중 같은 keyword 중복 변환 방지)

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

/** keyword로 찾은 참고파일을 스프레드시트로 열어서 반환한다(원본이 안 바뀌었으면 재변환 없이 캐시 재사용). */
function getReferenceSpreadsheet_(keyword) {
  if (_refCache_[keyword]) return _refCache_[keyword].ss;

  const file = findReferenceFile_(keyword);

  if (file.getMimeType() === MimeType.GOOGLE_SHEETS) {
    const ss = SpreadsheetApp.open(file);
    _refCache_[keyword] = { ss: ss };
    return ss;
  }

  const props = PropertiesService.getScriptProperties();
  const cacheProp = 'REF_CACHE_' + keyword;
  const sourceUpdatedAt = String(file.getLastUpdated().getTime());
  const cached = _readJsonProp_(props, cacheProp);

  if (cached && cached.sourceId === file.getId() && cached.sourceUpdatedAt === sourceUpdatedAt) {
    // 원본이 지난번 변환 이후 안 바뀜 — 예전에 만들어둔 구글시트 사본을 재변환 없이 그대로 재사용
    try {
      const ss = SpreadsheetApp.openById(cached.tempFileId);
      _refCache_[keyword] = { ss: ss };
      return ss;
    } catch (e) {
      // 캐시된 사본이 지워졌거나 접근 불가 — 아래에서 새로 변환
    }
  }

  // 원본이 바뀌었거나 캐시가 없음 — 새로 변환(고급 Drive 서비스 필요)하고 예전 캐시 사본은 정리.
  // 부모 폴더를 안 정해주면 내 드라이브 루트에 생기므로 REFERENCE_FOLDER_ID 폴더 안에 만들어지도록 명시.
  // Drive API 고급 서비스는 v3가 붙으면 Files.insert가 없고 Files.create로 바뀌므로 둘 다 지원한다.
  if (cached && cached.tempFileId) {
    try { DriveApp.getFileById(cached.tempFileId).setTrashed(true); } catch (e) { /* 이미 지워졌으면 무시 */ }
  }
  const tempFileId = Drive.Files.create
    ? Drive.Files.create({ name: '__cache_' + file.getName(), mimeType: MimeType.GOOGLE_SHEETS, parents: [CONFIG.REFERENCE_FOLDER_ID] }, file.getBlob()).id // v3
    : Drive.Files.insert({ title: '__cache_' + file.getName(), mimeType: MimeType.GOOGLE_SHEETS, parents: [{ id: CONFIG.REFERENCE_FOLDER_ID }] }, file.getBlob(), { convert: true }).id; // v2
  props.setProperty(cacheProp, JSON.stringify({ sourceId: file.getId(), sourceUpdatedAt: sourceUpdatedAt, tempFileId: tempFileId }));

  const ss = SpreadsheetApp.openById(tempFileId);
  _refCache_[keyword] = { ss: ss };
  return ss;
}

function _readJsonProp_(props, key) {
  try {
    const raw = props.getProperty(key);
    return raw ? JSON.parse(raw) : null;
  } catch (e) {
    return null;
  }
}

/** generatePlan() 끝에서 호출 — 캐시 사본은 이제 계속 재사용하므로 지우지 않고, 이번 실행용 메모리 캐시만 비운다. */
function cleanupReferenceFiles_() {
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
