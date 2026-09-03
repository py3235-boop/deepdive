/**
 * ============================================================================
 *  집합공정 생산계획 자동화 — Utils.gs (공통 헬퍼)
 * ----------------------------------------------------------------------------
 *  이후 모든 단계(변환·계획·렌더링·트리거)는 여기 있는 헬퍼만 쓴다.
 *   1. 경고 수집            warn_ · getWarnings_ · clearWarnings_
 *   2. 값 정규화            normalizeItemCode_ · toDate_ · toNumber_ · dateKey_ · fmtDate_ · isBlankRow_
 *   3. 탭 읽기              readTable_ · readTab_  (헤더 이름 기준, 빈 행 제거, 헤더 별칭 폴백)
 *   4. 스크립트 속성        getProp_ · setProp_ · getJsonProp_ · setJsonProp_
 *   5. 파일 자동 해결       openMaster_ · getMasterFolder_ · openShip_ · openResult_ · openBackupFolder_
 *   6. Drive 탐색·외부 파일  getOrderRootFolder_ · findFilesByPattern_ · findOrderFiles_ · findActualFiles_ ·
 *                          withSpreadsheetFile_ · readExternalTab_ · parseDriveId_
 *   7. [설정] 탭            loadSettings_ · getSetting_
 *  시트 I/O는 전부 배치(getValues/setValues). 셀 단위 루프 금지.
 * ============================================================================
 */

/* ────────────────────────────────────────────────────────────────────────────
 *  1. 경고 수집 — 실행 중 발생한 경고를 모아두고, Publisher가 [오류] 탭에 기록한다 (#4)
 * ──────────────────────────────────────────────────────────────────────────── */
const WARNINGS_ = [];

/** 경고 1건 기록 (실행을 멈추지 않는 문제). stage = 어느 단계인지 (예: '변환', '배정', 'testLoadOnly') */
function warn_(stage, message) {
  WARNINGS_.push({ at: new Date(), stage: String(stage), message: String(message) });
  Logger.log(`⚠ [${stage}] ${message}`);
}
function getWarnings_() { return WARNINGS_.slice(); }
function clearWarnings_() { WARNINGS_.length = 0; }

/* ────────────────────────────────────────────────────────────────────────────
 *  2. 값 정규화
 * ──────────────────────────────────────────────────────────────────────────── */

/**
 * 품목코드를 문자열로 통일한다.
 *  - 기준 탭의 품목코드는 텍스트지만 CAPA 탭·실적 파일의 품목코드는 숫자로 들어온다.
 *  - 숫자 7000260 → '7000260', 문자열 '7000260.0' → '7000260', ' 7000260 ' → '7000260'
 *  - 빈 값(null/undefined/'')은 '' 반환 (호출 측에서 빈 코드를 걸러낸다)
 */
function normalizeItemCode_(v) {
  if (v === null || v === undefined) return '';
  if (typeof v === 'number') return isFinite(v) ? String(Math.round(v)) : '';
  let s = String(v).trim();
  if (/^\d+\.0+$/.test(s)) s = s.replace(/\.0+$/, '');   // 숫자가 문자열 '7000260.0'으로 온 경우
  return s;
}

/**
 * 날짜 값을 Date로 통일한다. 실적 파일의 날짜·시각 열은 전부 텍스트로 들어온다.
 *  - Date → 그대로 (Invalid Date는 에러)
 *  - '' / null → null (빈 셀. 호출 측이 판단)
 *  - 'YYYY-MM-DD' · 'YYYY-MM-DD HH:mm' · 'YYYY-MM-DD HH:mm:ss' (구분자 - . / 허용) → 스크립트 타임존(Asia/Seoul) Date
 *  - 'MM/DD' · 'M/D' 텍스트(판독 파트 원본 헤더) → opts.defaultYear가 있을 때만 그 연도로 해석
 *  - 숫자(스프레드시트 일련번호) → 1899-12-30 기준 환산
 *  - 그 외는 어느 탭 어느 값인지(ctx) 포함해 에러
 * @param {*} v            셀 값
 * @param {string} ctx     에러 메시지용 위치 설명 (예: '[생산실적] 12행 실적일자')
 * @param {{defaultYear?: number}} [opts]
 */
function toDate_(v, ctx, opts) {
  opts = opts || {};
  const where = ctx ? ` (${ctx})` : '';
  if (v instanceof Date) {
    if (isNaN(v.getTime())) throw new Error(`날짜 값이 올바르지 않습니다${where}`);
    return v;
  }
  if (v === null || v === undefined || v === '') return null;
  if (typeof v === 'number') {
    // 스프레드시트 일련번호 (getValues는 날짜 셀을 Date로 주지만, 텍스트→숫자 변환된 값 대비)
    const base = new Date(1899, 11, 30);
    base.setDate(base.getDate() + Math.floor(v));
    const frac = v - Math.floor(v);
    if (frac) base.setTime(base.getTime() + Math.round(frac * 86400000));
    return base;
  }
  const s = String(v).trim();
  let m = s.match(/^(\d{4})[-./](\d{1,2})[-./](\d{1,2})(?:[ T](\d{1,2}):(\d{2})(?::(\d{2}))?)?\.?$/);
  if (m) return new Date(+m[1], +m[2] - 1, +m[3], +(m[4] || 0), +(m[5] || 0), +(m[6] || 0));
  m = s.match(/^(\d{4})(\d{2})(\d{2})$/);                        // 'YYYYMMDD'
  if (m) return new Date(+m[1], +m[2] - 1, +m[3]);
  m = s.match(/^(\d{1,2})[/.-](\d{1,2})[/.-](\d{4})$/);           // 'M/D/YYYY' (미국식 표기로 내보낸 날짜 헤더)
  if (m) return new Date(+m[3], +m[1] - 1, +m[2]);
  m = s.match(/^(\d{1,2})[/.-](\d{1,2})$/);                       // 'MM/DD' — 연도는 opts.defaultYear
  if (m && opts.defaultYear) return new Date(opts.defaultYear, +m[1] - 1, +m[2]);
  throw new Error(`날짜 파싱 실패: "${s}"${where} — 허용 형식 YYYY-MM-DD[ HH:mm[:ss]]`);
}

/**
 * 고객사 표기를 `고객사A`·`고객사B`·`고객사C`로 통일한다 (익명화 절대 규칙 · [출하우선순위] 탭과 키 일치).
 *  판독 파트 산출물은 `A사`·`B사`·`C사`, 또는 `고객사 A`·`A`처럼 올 수 있다. 인식 못 하는 표기는 trim만 해서 그대로 둔다.
 */
function normalizeCustomer_(v) {
  const s = String(v === null || v === undefined ? '' : v).trim();
  if (!s) return '';
  let m = s.match(/^고객사\s*([A-Za-z])$/) || s.match(/^([A-Za-z])\s*사$/) || s.match(/^([A-Za-z])$/);
  if (m) return '고객사' + m[1].toUpperCase();
  return s;
}

/** 숫자로 통일 — 천단위 콤마·공백 제거. 빈 값·비숫자는 0 */
function toNumber_(v) {
  if (typeof v === 'number') return isFinite(v) ? v : 0;
  if (v === null || v === undefined || v === '') return 0;
  const n = parseFloat(String(v).replace(/[,\s]/g, ''));
  return isFinite(n) ? n : 0;
}

/** Date → 'yyyy-MM-dd' (스크립트 타임존). 재고흐름·휴무 Set 등 날짜 키로 쓴다 */
function dateKey_(d) { return Utilities.formatDate(d, CFG.TZ, 'yyyy-MM-dd'); }

/** Date → 임의 패턴 (스크립트 타임존). 예: fmtDate_(d, 'yyMMdd-HHmm') = 계획ID */
function fmtDate_(d, pattern) { return Utilities.formatDate(d, CFG.TZ, pattern); }

/** 모든 셀이 빈 행인지 (xlsx→Sheets 변환 탭의 빈 패딩 행 판별) */
function isBlankRow_(row) {
  for (let i = 0; i < row.length; i++) {
    const c = row[i];
    if (c !== '' && c !== null && c !== undefined) return false;
  }
  return true;
}

/* ────────────────────────────────────────────────────────────────────────────
 *  3. 탭 읽기 — 열 순서가 아니라 헤더 이름으로 읽는다
 *     실데이터 양식 탭(CAPA 16열, 실적 13열/11열)은 안 쓰는 열이 많다. 필요한 열만 헤더로 골라 읽고 나머지는 무시.
 * ──────────────────────────────────────────────────────────────────────────── */

/**
 * 헤더 별칭 — 실데이터 이식 대비. 정식 헤더가 없을 때만 별칭을 정식 이름으로도 매핑한다.
 *   '…품목코드'(접두어 변형, 예: `집합 품목코드`) → '품목코드'   (CAPA 폴백)
 *   '품번' → '품목코드', '품명' → '품목명'                        (출하실적 실데이터 폴백)
 */
const HEADER_ALIASES_ = { '품번': '품목코드', '품명': '품목명' };

/**
 * 탭 하나를 읽어 {headers, rows, sheet} 반환. rows = 헤더 이름을 키로 한 객체 배열.
 *  - 헤더 행 = 첫 번째 비어 있지 않은 행 (그 위의 빈 행은 무시)
 *  - 모든 셀이 빈 데이터 행은 건너뛴다 (xlsx 변환 탭의 빈 패딩 행)
 *  - 필수 탭이 없으면 탭 이름을 포함한 에러, 선택 탭이 없으면 빈 결과
 *  - 별칭 헤더는 원래 이름과 정식 이름 두 키로 모두 넣는다 (row['집합 품목코드'] === row['품목코드'])
 *  시트 I/O: getDataRange().getValues() 1회.
 */
function readTable_(ss, name, required) {
  const sh = ss.getSheetByName(name);
  if (!sh) {
    if (required) throw new Error(`필수 탭 [${name}]이(가) 없습니다 (파일: ${ss.getName()}) — 탭 이름이 명세와 정확히 같은지 확인하세요`);
    return { headers: [], rows: [], sheet: null, tabName: name };
  }
  const values = sh.getDataRange().getValues();
  let hi = 0;
  while (hi < values.length && isBlankRow_(values[hi])) hi++;
  if (hi >= values.length) return { headers: [], rows: [], sheet: sh, tabName: name };

  const headers = values[hi].map(h => String(h === null || h === undefined ? '' : h).trim());
  const canon = headers.map(h => canonHeader_(h, headers));   // 정식 이름 (없으면 원래 이름과 동일)

  const rows = [];
  for (let r = hi + 1; r < values.length; r++) {
    const row = values[r];
    if (isBlankRow_(row)) continue;
    const obj = {};
    for (let c = 0; c < headers.length; c++) {
      if (!headers[c]) continue;
      obj[headers[c]] = row[c];
      if (canon[c] !== headers[c]) obj[canon[c]] = row[c];
    }
    rows.push(obj);
  }
  return { headers, rows, sheet: sh, tabName: name };
}

/** readTable_의 rows만 (대부분의 호출은 이것으로 충분) */
function readTab_(ss, name, required) { return readTable_(ss, name, required).rows; }

/** 헤더 정식 이름 결정 — 정식 헤더가 이미 있으면 별칭을 적용하지 않는다(중복 방지) */
function canonHeader_(h, allHeaders) {
  if (!h) return h;
  if (HEADER_ALIASES_[h] && allHeaders.indexOf(HEADER_ALIASES_[h]) < 0) return HEADER_ALIASES_[h];
  if (h !== '품목코드' && /품목코드$/.test(h) && allHeaders.indexOf('품목코드') < 0) return '품목코드';
  return h;
}

/* ────────────────────────────────────────────────────────────────────────────
 *  4. 스크립트 속성 (자동 생성한 파일 ID·처리이력·해시 저장소)
 * ──────────────────────────────────────────────────────────────────────────── */
function getProp_(key) { return PropertiesService.getScriptProperties().getProperty(key); }
function setProp_(key, value) { PropertiesService.getScriptProperties().setProperty(key, String(value)); }
function getJsonProp_(key, dflt) {
  const raw = getProp_(key);
  if (!raw) return dflt;
  try { return JSON.parse(raw); } catch (e) { warn_('속성', `${key} JSON 파싱 실패 — 기본값 사용`); return dflt; }
}
function setJsonProp_(key, obj) { setProp_(key, JSON.stringify(obj)); }

/* ────────────────────────────────────────────────────────────────────────────
 *  5. 파일 자동 해결 — 사람이 입력할 ID 없음
 *     우선순위: CFG 값 → 스크립트 속성 → (첫 실행 때만) 기준정보와 같은 폴더에 생성 후 속성 저장
 * ──────────────────────────────────────────────────────────────────────────── */

/** 기준정보 스프레드시트 — CFG.MASTER_SS_ID가 비면 바인딩 컨테이너 자신 */
function openMaster_() {
  if (CFG.MASTER_SS_ID) return SpreadsheetApp.openById(CFG.MASTER_SS_ID);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('기준정보 스프레드시트를 열 수 없습니다 — CFG.MASTER_SS_ID가 비어 있고 스크립트가 바인딩된 파일도 없습니다 (clasp create --parentId 확인)');
  return ss;
}

/** 기준정보 파일이 들어 있는 폴더(=프로젝트 루트). 자동 생성 파일은 모두 여기에 만든다 */
function getMasterFolder_() {
  const file = DriveApp.getFileById(openMaster_().getId());
  const parents = file.getParents();
  return parents.hasNext() ? parents.next() : DriveApp.getRootFolder();
}

/**
 * 리소스 ID 해결 공통 로직.
 *  1) cfgValue가 있으면 그것 (항상 우선)
 *  2) 스크립트 속성 propKey에 저장된 ID가 있고 Drive에 살아 있으면 그것 (휴지통이면 복원해서 재사용 — 링크 보존)
 *  3) 없고 createFn이 있으면 만들어 속성에 저장 (첫 실행 때만 일어난다). createFn이 없으면 null
 * @param {'file'|'folder'} kind
 */
function resolveResourceId_(cfgValue, propKey, createFn, kind) {
  if (cfgValue) return cfgValue;
  const saved = getProp_(propKey);
  if (saved) {
    const status = checkDriveItem_(saved, kind);
    if (status === 'ok') return saved;
    if (status === 'restored') { warn_('자동해결', `${propKey}의 ${kind}(${saved})이(가) 휴지통에 있어 복원해 재사용합니다`); return saved; }
    warn_('자동해결', `${propKey} 속성의 ${kind}(${saved})을(를) Drive에서 찾을 수 없어 새로 만듭니다 — 기존 공유 링크는 더 이상 열리지 않습니다`);
  }
  if (!createFn) return null;
  const id = createFn();
  setProp_(propKey, id);
  Logger.log(`[자동해결] ${propKey} 신규 생성 → ${id}`);
  return id;
}

/** Drive 항목 상태: 'ok' | 'restored'(휴지통에서 복원) | 'missing' */
function checkDriveItem_(id, kind) {
  try {
    const item = kind === 'folder' ? DriveApp.getFolderById(id) : DriveApp.getFileById(id);
    if (item.isTrashed()) { item.setTrashed(false); return 'restored'; }
    return 'ok';
  } catch (e) {
    return 'missing';
  }
}

/** 기준정보와 같은 폴더에 새 스프레드시트 생성 (로케일 ko_KR · 타임존 Asia/Seoul) */
function createSpreadsheetInMasterFolder_(name) {
  const ss = SpreadsheetApp.create(name);
  ss.setSpreadsheetLocale('ko_KR');
  ss.setSpreadsheetTimeZone(CFG.TZ);
  DriveApp.getFileById(ss.getId()).moveTo(getMasterFolder_());
  return ss;
}

/**
 * 출하계획 스프레드시트 (판독 파트 인터페이스). 없을 때:
 *  - opts.create === true → 기준정보와 같은 폴더에 `출하계획` 생성, [출하계획] 4열 헤더, 품목코드 열 텍스트 서식, ID 속성 저장
 *  - 그 외 → null 반환 (runAll은 "출하계획 없음"으로 중단, testLoadOnly는 존재 여부만 표시)
 */
function openShip_(opts) {
  opts = opts || {};
  const id = resolveResourceId_(CFG.SHIP_SS_ID, CFG.PROP.SHIP_SS_ID, opts.create ? function () {
    const ss = createSpreadsheetInMasterFolder_(CFG.FILE_NAMES.SHIP);
    const sh = ss.getSheets()[0];
    sh.setName(SHEET.SHIP);
    sh.getRange(1, 1, 1, SHEET.SHIP_HEADERS.length).setValues([SHEET.SHIP_HEADERS]).setFontWeight('bold');
    sh.getRange('A:A').setNumberFormat('@');   // 품목코드 텍스트 (앞자리 0·숫자 변환 방지)
    sh.setFrozenRows(1);
    return ss.getId();
  } : null, 'file');
  return id ? SpreadsheetApp.openById(id) : null;
}

/** 결과 스프레드시트 `생산계획` (고정 ID). 없으면 첫 실행 때만 생성 — 탭 구성은 Publisher가 채운다 (#4) */
function openResult_() {
  const id = resolveResourceId_(CFG.RESULT_SS_ID, CFG.PROP.RESULT_SS_ID, function () {
    const ss = createSpreadsheetInMasterFolder_(CFG.FILE_NAMES.RESULT);
    ss.getSheets()[0].setName(SHEET.RESULT.INTEGRATED);
    return ss.getId();
  }, 'file');
  return SpreadsheetApp.openById(id);
}

/** 백업 폴더 `일별생산계획/` (실행마다 결과 사본 저장). 없으면 첫 실행 때만 생성 */
function openBackupFolder_() {
  const id = resolveResourceId_(CFG.BACKUP_FOLDER_ID, CFG.PROP.BACKUP_FOLDER_ID, function () {
    return getMasterFolder_().createFolder(CFG.FILE_NAMES.BACKUP_FOLDER).getId();
  }, 'folder');
  return DriveApp.getFolderById(id);
}

/** 작업지시서 폴더 `작업지시서/` — 호기별 파일 10개가 들어간다. 없으면 첫 실행 때만 생성 */
function openWorkOrderFolder_() {
  const id = resolveResourceId_(CFG.WORKORDER_FOLDER_ID, CFG.PROP.WORKORDER_FOLDER_ID, function () {
    return getMasterFolder_().createFolder(CFG.FILE_NAMES.WORKORDER_FOLDER).getId();
  }, 'folder');
  return DriveApp.getFolderById(id);
}

/**
 * 호기별 작업지시 스프레드시트 — `집합01호기 작업지시` 처럼 호기마다 파일 하나.
 * ID를 스크립트 속성(WORKORDER_SS_JSON)에 호기별로 저장해 계속 재사용한다 — 현장에 뿌린 링크·QR이 바뀌면 안 된다.
 * 파일이 휴지통에 있으면 복원해 쓰고, Drive에서 사라졌을 때만 새로 만든다.
 */
function openWorkOrderSs_(machine) {
  const map = getJsonProp_(CFG.PROP.WORKORDER_SS_JSON, {});
  const saved = map[machine];
  if (saved) {
    const status = checkDriveItem_(saved, 'file');
    if (status === 'ok' || status === 'restored') {
      if (status === 'restored') warn_('작업지시서', `${machine} 파일이 휴지통에 있어 복원했습니다`);
      return SpreadsheetApp.openById(saved);
    }
    warn_('작업지시서', `${machine} 파일(${saved})을 Drive에서 찾을 수 없어 새로 만듭니다 — 현장에 뿌린 이전 링크는 열리지 않습니다`);
  }
  const ss = SpreadsheetApp.create(machine + CFG.FILE_NAMES.WORKORDER_SUFFIX);
  ss.setSpreadsheetLocale('ko_KR');
  ss.setSpreadsheetTimeZone(CFG.TZ);
  DriveApp.getFileById(ss.getId()).moveTo(openWorkOrderFolder_());
  map[machine] = ss.getId();
  setJsonProp_(CFG.PROP.WORKORDER_SS_JSON, map);
  Logger.log(`[작업지시서] ${machine} 파일 신규 생성 → ${ss.getUrl()}`);
  return ss;
}

/** 호기별 작업지시 파일 링크 목록 (QR·공유용). 아직 안 만들어진 호기는 빈 URL */
function workOrderLinks_() {
  const map = getJsonProp_(CFG.PROP.WORKORDER_SS_JSON, {});
  return CFG.MACHINES.map(m => {
    let url = '';
    if (map[m] && checkDriveItem_(map[m], 'file') !== 'missing') {
      try { url = SpreadsheetApp.openById(map[m]).getUrl(); } catch (e) { url = ''; }
    }
    return { machine: m, url };
  });
}

/** 탭이 없으면 만들고 있으면 그대로 반환 (결과 파일 탭 보장용) */
function ensureSheet_(ss, name) {
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

/* ────────────────────────────────────────────────────────────────────────────
 *  6. Drive 탐색 · 외부 파일 읽기
 *     출고계획·실적 파일은 "폴더명·위치가 아니라 파일명 패턴"으로 찾는다 (판독 파트 폴더명 미정).
 * ──────────────────────────────────────────────────────────────────────────── */

/** 탐색 루트 — CFG.ORDER_ROOT_FOLDER_ID가 비면 기준정보의 부모 폴더(프로젝트 루트) */
function getOrderRootFolder_() {
  return CFG.ORDER_ROOT_FOLDER_ID ? DriveApp.getFolderById(parseDriveId_(CFG.ORDER_ROOT_FOLDER_ID)) : getMasterFolder_();
}

/* 실행 1회 안에서 같은 루트를 여러 패턴으로 찾을 때 Drive를 한 번만 훑도록 캐시 */
const FILE_LIST_CACHE_ = {};

/**
 * 폴더 아래 모든 파일을 재귀적으로 나열 (휴지통 제외, 깊이 한도 CFG.SEARCH_MAX_DEPTH).
 * 바로가기(shortcut)는 대상 파일로 풀어서 ID·MIME·수정시각을 대상 기준으로 기록한다.
 * @returns {Array<{id,name,mimeType,lastUpdated:Date,folderPath,url}>}
 */
function listFilesUnder_(rootFolder) {
  const rootId = rootFolder.getId();
  if (FILE_LIST_CACHE_[rootId]) return FILE_LIST_CACHE_[rootId];
  const out = [];
  const seen = {};
  (function walk(folder, path, depth) {
    if (depth > CFG.SEARCH_MAX_DEPTH) { warn_('탐색', `폴더 깊이 한도(${CFG.SEARCH_MAX_DEPTH}) 초과 — ${path} 아래는 건너뜁니다`); return; }
    const files = folder.getFiles();
    while (files.hasNext()) {
      const f = files.next();
      if (f.isTrashed()) continue;
      let id = f.getId(), mime = f.getMimeType(), lastUpdated = f.getLastUpdated(), url = f.getUrl();
      if (mime === MimeType.SHORTCUT) {
        const tid = f.getTargetId();
        if (!tid) continue;
        try {
          const tf = DriveApp.getFileById(tid);
          if (tf.isTrashed()) continue;
          id = tf.getId(); mime = tf.getMimeType(); lastUpdated = tf.getLastUpdated(); url = tf.getUrl();
        } catch (e) { continue; }
      }
      if (seen[id]) continue;
      seen[id] = true;
      out.push({ id, name: f.getName(), mimeType: mime, lastUpdated, folderPath: path, url });
    }
    const subs = folder.getFolders();
    while (subs.hasNext()) {
      const sf = subs.next();
      if (sf.isTrashed()) continue;
      walk(sf, path + '/' + sf.getName(), depth + 1);
    }
  })(rootFolder, rootFolder.getName(), 0);
  FILE_LIST_CACHE_[rootId] = out;
  return out;
}

/** 루트 아래에서 파일명이 pattern(정규식)에 맞는 파일 — 이름순 */
function findFilesByPattern_(rootFolder, pattern) {
  return listFilesUnder_(rootFolder)
    .filter(f => pattern.test(f.name))
    .sort((a, b) => a.name.localeCompare(b.name, 'ko'));
}

/** 출고계획 파일 전부 (`X월_출고계획(통합)` 패턴, 폴더 무관, 이름순) */
function findOrderFiles_() {
  return findFilesByPattern_(getOrderRootFolder_(), CFG.ORDER_FILE_PATTERN);
}

/**
 * 실적 파일 — 생산실적·출하실적 각각 패턴에 맞는 파일 중 가장 최근 수정본 1개 (없으면 null)
 * @returns {{prod: object|null, ship: object|null}}
 */
function findActualFiles_() {
  const root = getOrderRootFolder_();
  const latest = (pattern) => {
    const list = findFilesByPattern_(root, pattern);
    if (!list.length) return null;
    list.sort((a, b) => b.lastUpdated.getTime() - a.lastUpdated.getTime());
    if (list.length > 1) Logger.log(`[탐색] ${pattern} 파일 ${list.length}개 중 최신본 사용: ${list[0].name} (${fmtDate_(list[0].lastUpdated, 'yyyy-MM-dd HH:mm')})`);
    return list[0];
  };
  return { prod: latest(CFG.ACTUAL_PROD_PATTERN), ship: latest(CFG.ACTUAL_SHIP_PATTERN) };
}

/**
 * 파일 정보(listFilesUnder_ 항목)를 스프레드시트로 열어 fn(ss)를 실행한다.
 *  - Google 스프레드시트 → 바로 openById
 *  - xlsx 등 → Drive 고급 서비스(v3)로 임시 구글시트 변환 후 읽고, try/finally로 반드시 삭제
 *  - 고급 서비스가 꺼져 있으면 에러 (호출 측이 잡아서 그 파일만 건너뛰고 경고)
 *  원본 파일은 읽기만 한다 — 옮기거나 이름을 바꾸거나 쓰지 않는다.
 */
function withSpreadsheetFile_(fileInfo, fn) {
  if (fileInfo.mimeType === MimeType.GOOGLE_SHEETS) return fn(SpreadsheetApp.openById(fileInfo.id));
  if (typeof Drive === 'undefined' || !Drive.Files) {
    throw new Error(`"${fileInfo.name}"은(는) Google 스프레드시트가 아니라서(${fileInfo.mimeType}) 변환이 필요한데 Drive 고급 서비스가 꺼져 있습니다 — appsscript.json의 enabledAdvancedServices(Drive v3) 확인`);
  }
  let tempId = null;
  try {
    const copy = Drive.Files.copy({ name: '_tmp_변환_' + fileInfo.name, mimeType: MimeType.GOOGLE_SHEETS }, fileInfo.id);
    tempId = copy.id;
    return fn(SpreadsheetApp.openById(tempId));
  } finally {
    if (tempId) {
      try { Drive.Files.remove(tempId); }
      catch (e1) {
        try { DriveApp.getFileById(tempId).setTrashed(true); }
        catch (e2) { warn_('변환', `임시 변환 파일 삭제 실패(${tempId}) — 수동 삭제 필요`); }
      }
    }
  }
}

/**
 * 외부 파일(출고계획·실적)의 탭 하나를 읽는다. preferredTab 이름의 탭이 있으면 그것, 없으면 첫 탭
 * (실적 내보내기 시트명이 제각각일 수 있음). 반환은 readTable_와 같고 tabName에 실제 읽은 탭 이름.
 */
function readExternalTab_(fileInfo, preferredTab) {
  return withSpreadsheetFile_(fileInfo, function (ss) {
    const sh = (preferredTab && ss.getSheetByName(preferredTab)) || ss.getSheets()[0];
    if (!sh) throw new Error(`"${fileInfo.name}"에 읽을 탭이 없습니다`);
    return readTable_(ss, sh.getName(), true);
  });
}

/** Drive URL 또는 ID 문자열에서 ID만 추출 (폴더 URL·시트 URL·`?id=` 모두 허용). ID가 아니면 원문 trim */
function parseDriveId_(v) {
  const s = String(v || '').trim();
  if (!s) return '';
  let m = s.match(/\/folders\/([A-Za-z0-9_-]+)/) || s.match(/\/d\/([A-Za-z0-9_-]+)/) || s.match(/[?&]id=([A-Za-z0-9_-]+)/);
  return m ? m[1] : s;
}

/* ────────────────────────────────────────────────────────────────────────────
 *  7. [설정] 탭 — 키 · 값 · 설명. 없는 키는 CFG 기본값
 * ──────────────────────────────────────────────────────────────────────────── */

/** [설정] 탭을 {키: 값} 맵으로 (탭이 없으면 빈 객체). 값은 원형 그대로(Date·숫자·문자열) */
function loadSettings_(master) {
  const rows = readTab_(master, SHEET.MASTER.SETTINGS, false);
  const map = {};
  rows.forEach(r => {
    const k = String(r['키'] === undefined ? '' : r['키']).trim();
    if (k) map[k] = r['값'];
  });
  return map;
}

/** 설정값 조회 — 없거나 빈 값이면 dflt */
function getSetting_(settings, key, dflt) {
  const v = settings ? settings[key] : undefined;
  return (v === undefined || v === null || v === '') ? dflt : v;
}

/** [설정] PLAN_MONTH('YYYY-MM' 문자열 또는 Date) → {year, month}. 없으면 null */
function getPlanMonth_(settings) {
  const v = getSetting_(settings, SETTING_KEYS.PLAN_MONTH, null);
  if (!v) return null;
  if (v instanceof Date) return { year: v.getFullYear(), month: v.getMonth() + 1 };
  const m = String(v).trim().match(/^(\d{4})[-./](\d{1,2})/);
  return m ? { year: +m[1], month: +m[2] } : null;
}

/**
 * 'MM/DD'처럼 연도가 없는 날짜의 연도를 계획월 기준으로 정한다 (12월↔1월 걸침 보정).
 *  계획월 12월에 1월 날짜 → 다음 해, 계획월 1월에 12월 날짜 → 전 해, 그 외 → 계획월의 해. 계획월이 없으면 올해
 */
function yearForMonth_(planMonth, month) {
  if (!planMonth) return new Date().getFullYear();
  if (planMonth.month === 12 && month === 1) return planMonth.year + 1;
  if (planMonth.month === 1 && month === 12) return planMonth.year - 1;
  return planMonth.year;
}

/* ────────────────────────────────────────────────────────────────────────────
 *  8. 날짜 셀 정규화 · 내용 해시 · 알림 스텁 · 폴더 헬퍼
 * ──────────────────────────────────────────────────────────────────────────── */

/**
 * 셀의 날짜 값을 "그 스프레드시트에서 보이는 날짜" 기준으로 스크립트 타임존 00:00 Date로 통일한다.
 *  Sheets의 Date 셀은 그 파일의 타임존으로 해석되므로, 파일 타임존(tz)이 스크립트 타임존과 다르면
 *  그냥 쓰면 하루가 밀릴 수 있다 → tz로 yyyy-MM-dd를 뽑은 뒤 스크립트 타임존으로 다시 만든다.
 *  문자열은 toDate_로 파싱(opts.defaultYear 지원). 빈 값은 null.
 */
function normalizeDateCell_(v, tz, ctx, opts) {
  if (v instanceof Date) {
    if (isNaN(v.getTime())) throw new Error(`날짜 값이 올바르지 않습니다 (${ctx || '위치 미상'})`);
    const p = Utilities.formatDate(v, tz || CFG.TZ, 'yyyy-MM-dd').split('-');
    return new Date(+p[0], +p[1] - 1, +p[2]);
  }
  const d = toDate_(v, ctx, opts);
  return d ? new Date(d.getFullYear(), d.getMonth(), d.getDate()) : null;
}

/**
 * 탭 내용 해시 (MD5 hex) — 갱신 감지용. getDataRange().getValues()를 JSON으로 만들어 해시한다.
 * 탭이 없으면 빈 배열의 해시. 결과 파일은 해시 대상이 아니다(자기 자신을 다시 깨우지 않게).
 */
function hashValues_(ss, tabName) {
  const sh = ss.getSheetByName(tabName);
  const values = sh ? sh.getDataRange().getValues() : [];
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, JSON.stringify(values), Utilities.Charset.UTF_8);
  return bytes.map(b => ('0' + ((b + 256) % 256).toString(16)).slice(-2)).join('');
}

/* ────────────────────────────────────────────────────────────────────────────
 *  알림 — CFG.NOTIFY.channel: 'chat'(기본, Google Chat 웹훅 카드) | 'telegram' | 'mail'
 *  비밀값은 코드에 없다: CHAT_WEBHOOK_URL / TELEGRAM_TOKEN / TELEGRAM_CHAT_ID 스크립트 속성.
 *  CHAT_WEBHOOK_URL 속성이 없으면 기준정보 [설정] WEBHOOK_URL 폴백(리허설 편의 — public 전환 전 삭제).
 *  둘 다 없으면 에러 대신 "알림 생략" 로그. 전송 실패도 try/catch로 runAll을 죽이지 않는다.
 * ──────────────────────────────────────────────────────────────────────────── */

/**
 * 알림 단일 진입점
 * @param {string} message 본문 (\n 줄바꿈, URL 포함 가능 — chat에서는 "바로 열기" 앵커로 바뀐다)
 * @param {string} [label] 라벨 (예: '👤 [수동실행]' · '🤖 [자동실행]' · '👤 [수동실행 오류]')
 * @returns {boolean} 실제 전송 여부
 */
function notify_(message, label) {
  label = label || '';
  try {
    const ch = CFG.NOTIFY.channel;
    if (ch === 'chat') return notifyGoogleChat_(label, message);
    if (ch === 'telegram') return notifyTelegram_(label, message);
    if (ch === 'mail') return notifyMail_(label, message);
    Logger.log(`[알림 생략] 알 수 없는 채널 '${ch}'`);
    return false;
  } catch (e) {
    Logger.log(`[알림 실패] ${e.message}`);
    warn_('알림', `알림 전송 실패: ${e.message}`);
    return false;
  }
}

/** 웹훅 URL — 스크립트 속성 → [설정] 탭 폴백. 없으면 '' */
function getWebhookUrl_() {
  let url = getProp_(CFG.NOTIFY.PROP_CHAT_WEBHOOK) || '';
  if (!url) {
    try { url = String(getSetting_(loadSettings_(openMaster_()), CFG.NOTIFY.SETTING_WEBHOOK_KEY, '') || ''); } catch (e) { url = ''; }
  }
  return url.trim();
}

/** 라벨 문자열로 색 결정: '오류' 빨강 · '자동실행' 파랑 · '수동실행' 초록 · 그 외 검정 */
function labelColor_(label) {
  if (/오류/.test(label)) return '#FF0000';
  if (/자동실행/.test(label)) return '#0000FF';
  if (/수동실행/.test(label)) return '#008000';
  return '#000000';
}

/**
 * Google Chat cardsV2 페이로드 생성 (순수 함수 — 테스트 가능).
 *  본문: \n → <br>, URL은 <a href="URL">바로 열기</a> 앵커로 치환. 라벨은 <font color><b></b></font>
 */
function buildChatCard_(label, bodyText) {
  const esc = (s) => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  const body = esc(bodyText)
    .replace(/(https?:\/\/[^\s<]+)/g, '<a href="$1">바로 열기</a>')
    .replace(/\n/g, '<br>');
  const text = `<font color="${labelColor_(label)}"><b>${esc(label)}</b></font> ${body}`;
  return { cardsV2: [{ cardId: 'plan-' + Date.now(), card: { sections: [{ widgets: [{ textParagraph: { text } }] }] } }] };
}

/** Google Chat 웹훅 POST. URL 미설정이면 로그 후 false */
function notifyGoogleChat_(label, bodyText) {
  const url = getWebhookUrl_();
  if (!url) { Logger.log(`[알림 생략] ${CFG.NOTIFY.PROP_CHAT_WEBHOOK} 속성과 [설정] ${CFG.NOTIFY.SETTING_WEBHOOK_KEY} 모두 미설정 — ${label} ${String(bodyText).split('\n')[0]}`); return false; }
  const res = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json; charset=UTF-8',
    payload: JSON.stringify(buildChatCard_(label, bodyText)),
    muteHttpExceptions: true,
  });
  const code = res.getResponseCode();
  if (code < 200 || code >= 300) throw new Error(`Google Chat 웹훅 HTTP ${code}: ${String(res.getContentText()).slice(0, 200)}`);
  Logger.log(`[알림] Google Chat 전송 완료 (${label})`);
  return true;
}

/** 텔레그램 폴백 — 색·앵커 없이 아이콘+라벨+원문 URL */
function notifyTelegram_(label, bodyText) {
  const token = getProp_(CFG.NOTIFY.PROP_TELEGRAM_TOKEN), chatId = getProp_(CFG.NOTIFY.PROP_TELEGRAM_CHAT_ID);
  if (!token || !chatId) { Logger.log('[알림 생략] TELEGRAM_TOKEN / TELEGRAM_CHAT_ID 미설정'); return false; }
  const res = UrlFetchApp.fetch(`https://api.telegram.org/bot${token}/sendMessage`, {
    method: 'post', contentType: 'application/json; charset=UTF-8', muteHttpExceptions: true,
    payload: JSON.stringify({ chat_id: chatId, text: `${label}\n${bodyText}`, disable_web_page_preview: true }),
  });
  const code = res.getResponseCode();
  if (code < 200 || code >= 300) throw new Error(`텔레그램 HTTP ${code}: ${String(res.getContentText()).slice(0, 200)}`);
  return true;
}

/** 메일 폴백 — MailApp. 수신자는 CFG.NOTIFY.MAIL_TO, 비어 있으면 실행 계정 */
function notifyMail_(label, bodyText) {
  const to = CFG.NOTIFY.MAIL_TO || Session.getActiveUser().getEmail();
  if (!to) { Logger.log('[알림 생략] 메일 수신자 없음'); return false; }
  MailApp.sendEmail(to, `${label} 집합 생산계획`, `${label}\n${bodyText}`);
  return true;
}

/** 부모 폴더 아래 이름이 같은 하위 폴더를 찾고, 없으면 만든다 (휴지통 제외) */
function getOrCreateSubFolder_(parent, name) {
  const it = parent.getFoldersByName(name);
  while (it.hasNext()) {
    const f = it.next();
    if (!f.isTrashed()) return f;
  }
  return parent.createFolder(name);
}

/** 파일 목록 캐시 무효화 — 실행 중 Drive에 파일을 만들거나 복사한 뒤 호출 */
function invalidateFileListCache_() {
  Object.keys(FILE_LIST_CACHE_).forEach(k => { delete FILE_LIST_CACHE_[k]; });
}

/** 행 객체에서 후보 헤더 이름 중 처음 존재하는 값 (실데이터 헤더 변형 대비). 없으면 undefined */
function field_(row, names) {
  for (let i = 0; i < names.length; i++) {
    if (row[names[i]] !== undefined && row[names[i]] !== null && row[names[i]] !== '') return row[names[i]];
  }
  return undefined;
}

/**
 * 호기 표기를 `집합01호기`~`집합10호기`로 통일 (`집합1호기`·`집합 01 호기`·숫자 1 등 변형 허용). 인식 못 하면 trim만
 */
function normalizeMachine_(v) {
  const s = String(v === null || v === undefined ? '' : v).trim();
  if (!s) return '';
  let m = s.match(/^집합\s*(\d{1,2})\s*호기$/) || s.match(/^(\d{1,2})$/) || s.match(/^(\d{1,2})\s*호기$/);
  if (m) return '집합' + String(+m[1]).padStart(2, '0') + '호기';
  return s;
}

/* ────────────────────────────────────────────────────────────────────────────
 *  9. 근무시간 모델 — 계획 엔진·일별 분해가 공유하는 유일한 시간 규칙
 *     근무일 D = [D 08:00, D+1 08:00) (START_HOUR 8). 설비는 근무일 안에서 24시간 연속 가동.
 *     주말·[휴무] 날짜는 근무일이 아니며 통째로 건너뛴다 (그 날짜 08:00 ~ 다음날 08:00 구간 전체).
 *     holidays = { 'yyyy-MM-dd': 명칭 }
 * ──────────────────────────────────────────────────────────────────────────── */

/** 시각 t가 속한 근무일의 시작(08:00). 08:00 이전 시각은 전날 근무일에 속한다 */
function dayStartOf_(t) {
  const s = new Date(t.getFullYear(), t.getMonth(), t.getDate(), CFG.PLAN.START_HOUR, 0, 0, 0);
  if (t.getTime() < s.getTime()) s.setDate(s.getDate() - 1);
  return s;
}

/** 근무일 시작(08:00) → 다음 날 08:00 */
function nextDayStart_(dayStart) {
  const n = new Date(dayStart.getTime());
  n.setDate(n.getDate() + 1);
  n.setHours(CFG.PLAN.START_HOUR, 0, 0, 0);
  return n;
}

/** 그 근무일이 생산 가능한 날인지 (토·일·[휴무] 제외) */
function isWorkingDay_(dayStart, holidays) {
  const dow = dayStart.getDay();
  if (dow === 0 || dow === 6) return false;
  return !(holidays && holidays[dateKey_(dayStart)]);
}

/** dayStart부터 처음 만나는 근무일의 시작 (dayStart 자신이 근무일이면 그대로) */
function nextWorkingDayStart_(dayStart, holidays) {
  let d = dayStart;
  let guard = 0;
  while (!isWorkingDay_(d, holidays) && guard++ < 400) d = nextDayStart_(d);
  return d;
}

/** 기준 날짜 "다음" 근무일의 시작(08:00) — 이월 출하일 계산용 (기준일 당일은 제외) */
function firstWorkingDayAfter_(date, holidays) {
  const d = new Date(date.getFullYear(), date.getMonth(), date.getDate(), CFG.PLAN.START_HOUR, 0, 0, 0);
  return nextWorkingDayStart_(nextDayStart_(d), holidays);
}

/** 시각 t가 비근무일 안이면 다음 근무일 08:00으로, 근무일 안이면 그대로 */
function alignToWorkTime_(t, holidays) {
  const ds = dayStartOf_(t);
  return isWorkingDay_(ds, holidays) ? t : nextWorkingDayStart_(ds, holidays);
}

/** start에서 근무시간 hours만큼 지난 시각 (비근무일은 건너뜀). hours 0이면 start를 근무시간으로 정렬만 */
function addWorkHours_(start, hours, holidays) {
  let cur = alignToWorkTime_(start, holidays);
  let remain = hours;
  let guard = 0;
  while (remain > 1e-9 && guard++ < 2000) {
    const de = nextDayStart_(dayStartOf_(cur));
    const avail = (de.getTime() - cur.getTime()) / 3600000;
    if (remain <= avail + 1e-9) return new Date(cur.getTime() + Math.round(remain * 3600000));
    remain -= avail;
    cur = nextWorkingDayStart_(de, holidays);
  }
  return cur;
}

/** end에서 근무시간 hours만큼 거슬러 간 시각 (비근무일은 건너뜀) — 작업의 실제 생산 시작 시각 역산용 */
function subWorkHours_(end, hours, holidays) {
  let cur = end;
  let remain = hours;
  let guard = 0;
  // end가 비근무일 안이면 그 직전 근무일의 끝으로
  let ds = dayStartOf_(new Date(cur.getTime() - 1));
  if (!isWorkingDay_(ds, holidays)) { ds = prevWorkingDayStart_(ds, holidays); cur = nextDayStart_(ds); }
  while (remain > 1e-9 && guard++ < 2000) {
    ds = dayStartOf_(new Date(cur.getTime() - 1));
    const avail = (cur.getTime() - ds.getTime()) / 3600000;
    if (remain <= avail + 1e-9) return new Date(cur.getTime() - Math.round(remain * 3600000));
    remain -= avail;
    cur = nextDayStart_(prevWorkingDayStart_(new Date(ds.getTime() - 1), holidays));
  }
  return cur;
}

/** t가 속한 날부터 거슬러 처음 만나는 근무일의 시작 */
function prevWorkingDayStart_(t, holidays) {
  let d = dayStartOf_(t);
  let guard = 0;
  while (!isWorkingDay_(d, holidays) && guard++ < 400) { d = new Date(d.getTime()); d.setDate(d.getDate() - 1); d.setHours(CFG.PLAN.START_HOUR, 0, 0, 0); }
  return d;
}

/** a~b 사이의 근무시간(h) — 비근무일 제외. b ≤ a면 0 */
function workHoursBetween_(a, b, holidays) {
  if (b.getTime() <= a.getTime()) return 0;
  let cur = alignToWorkTime_(a, holidays);
  let total = 0;
  let guard = 0;
  while (cur.getTime() < b.getTime() && guard++ < 2000) {
    const de = nextDayStart_(dayStartOf_(cur));
    const segEnd = de.getTime() < b.getTime() ? de : b;
    total += (segEnd.getTime() - cur.getTime()) / 3600000;
    cur = nextWorkingDayStart_(de, holidays);
  }
  return total;
}
