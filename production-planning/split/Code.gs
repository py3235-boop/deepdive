/**
 * ============================================================================
 *  집합공정 생산계획 자동화 — Code.gs (진입점)
 * ----------------------------------------------------------------------------
 *  역할 : 설정(CFG·SHEET) · 데이터 로드 · 트리거 · [▶ 생산계획] 메뉴 · 테스트 진입점
 *
 *  이 스크립트는 "기준정보" 스프레드시트에 바인딩(container-bound)되어 있다.
 *   → 단순 onOpen() 함수만으로 기준정보 파일에 [▶ 생산계획] 메뉴가 뜬다.
 *   → 자동 재계획용 설치형 트리거 2개(onChange·1분)는 별개로 installTrigger()에서 등록한다 (#5).
 *
 *  파일 구성
 *   - Code.gs      : 이 파일. CFG/SHEET 상수, 데이터 로드(#3), runAll(#4), 트리거(#5), 메뉴, 테스트 진입점
 *   - Utils.gs     : 날짜·품목코드 정규화, 탭 읽기, 파일 자동 해결(출하계획·결과·백업), Drive 탐색
 *   - Planner.gs   : 호기 배정 → 작업목록 배열 (#3)
 *   - Scheduler.gs : 시간 시뮬레이션 (#3)
 *   - Publisher.gs : 결과 스프레드시트 렌더링 (#4)
 *
 *  ⚠ 코드에 사람이 넣어야 하는 ID는 하나도 없다. 모든 파일·폴더는 "ID 자동 해결" 규칙으로 찾거나 만든다.
 *  ⚠ 웹훅 URL·토큰 등 비밀값은 코드에 쓰지 않는다. 스크립트 속성에서 읽는다 (CFG.NOTIFY에는 키 이름만).
 *  ⚠ 고객사는 고객사A/B/C, 설비는 집합01호기~집합10호기 표기만 사용한다 (익명화 절대 규칙).
 * ============================================================================
 */

/* ────────────────────────────────────────────────────────────────────────────
 *  CFG — 시스템 설정 (설정값은 여기 한 곳에만 둔다)
 * ──────────────────────────────────────────────────────────────────────────── */
const CFG = {
  /* ── 파일·폴더 ID (전부 빈 문자열 기본 = 자동 해결) ─────────────────────────
   *  MASTER_SS_ID       비어 있으면 스크립트가 바인딩된 파일 자신(getActiveSpreadsheet)
   *  SHIP_SS_ID         비어 있으면 스크립트 속성 → 없으면 첫 변환 때 기준정보와 같은 폴더에 `출하계획` 생성
   *  RESULT_SS_ID       비어 있으면 스크립트 속성 → 없으면 첫 실행 때 기준정보와 같은 폴더에 `생산계획` 생성
   *  BACKUP_FOLDER_ID   비어 있으면 스크립트 속성 → 없으면 첫 실행 때 기준정보와 같은 폴더에 `일별생산계획/` 생성
   *  ORDER_ROOT_FOLDER_ID  출고계획·실적 파일 탐색 루트. 비어 있으면 기준정보 파일의 부모 폴더(=프로젝트 루트)
   *  ORDER_SOURCE_FOLDER_ID (선택) 판독 파트가 루트 밖 폴더를 쓸 때의 소스 폴더(URL/ID). 비면 [설정] 탭 값 → 없으면 동기화 생략
   *  값이 있으면 항상 그것을 우선한다. 두 번째 실행부터는 절대 새로 만들지 않는다(링크 불변).           */
  MASTER_SS_ID: '',
  SHIP_SS_ID: '',
  RESULT_SS_ID: '',
  BACKUP_FOLDER_ID: '',
  WORKORDER_FOLDER_ID: '',   // 작업지시서/ 폴더 — 호기마다 파일 하나(현장 배포용). 비면 첫 실행 때 기준정보와 같은 폴더에 생성
  ORDER_ROOT_FOLDER_ID: '',
  ORDER_SOURCE_FOLDER_ID: '',
  /* 외부 소스 폴더 동기화 스위치 — false면 [설정] ORDER_SOURCE_FOLDER_ID가 있어도 syncOrderFiles_를 실행하지 않는다.
   * 사용자 지시(2026-09-03): 판독 파트 출고계획 폴더 연동은 마지막 단계에서 사용자가 말할 때 켠다. 그 전까지 false 유지 */
  ORDER_SYNC_ENABLED: true,

  /* ── 파일명 패턴 (폴더명·확장자·MIME 무관 — 이름만 본다) ─────────────────────
   *  출고계획 파일명 두 가지를 받는다:
   *   ① `X월_출고계획(통합)…`  — 더미데이터·데모투입용(뒤에 `_추가발주` 등이 붙어도 걸린다)
   *   ② `출고계획`             — 판독 파트 산출물의 실제 파일명(정확히 일치할 때만.
   *                              `출고계획_발주스냅샷` 같은 부수 파일은 걸리지 않는다) */
  ORDER_FILE_PATTERN: /^(\d{1,2}월[ _]?출고계획\(통합\)|출고계획(\.xlsx?)?$)/,
  /* 정식 이름 — 후보가 여럿일 때 이 패턴에 맞는 파일을 먼저 고른다 (사용자 지시 2026-09-04).
   * 동기화 사본이 옛 이름(`출고계획`)으로 남아 있어도 정식 이름 파일이 있으면 그쪽을 쓴다. */
  ORDER_FILE_PATTERN_MAIN: /^\d{1,2}월[ _]?출고계획\(통합\)/,
  ACTUAL_PROD_PATTERN: /^생산실적/,                      // 실적 시스템이 매일 올리는 생산실적 (같은 패턴 여러 개면 최신 수정본)
  ACTUAL_SHIP_PATTERN: /^출하실적/,                      // 실적 시스템이 매일 올리는 출하실적

  /* ── 자동 생성 파일·폴더 이름 ──────────────────────────────────────────── */
  FILE_NAMES: {
    SHIP: '출하계획',            // 변환 산출물 = 판독 파트 인터페이스
    RESULT: '생산계획',          // 결과 파일 (고정 ID, 탭 내용만 덮어씀)
    BACKUP_FOLDER: '일별생산계획', // 실행마다 결과 사본 "YYMMDD-HHmm_생산계획"
    WORKORDER_FOLDER: '작업지시서', // 호기별 작업지시 파일 10개가 들어가는 폴더
    WORKORDER_SUFFIX: ' 작업지시서', // 파일명 = 호기 + 이 접미사 (예: 집합01호기 작업지시서)
    ORDER_COPY_FOLDER: '출고계획', // (선택) 외부 소스 폴더 동기화 복사본 위치
  },

  /* ── 스크립트 속성 키 이름 (값은 PropertiesService에 저장) ─────────────────── */
  PROP: {
    SHIP_SS_ID: 'SHIP_SS_ID',
    RESULT_SS_ID: 'RESULT_SS_ID',
    BACKUP_FOLDER_ID: 'BACKUP_FOLDER_ID',
    WORKORDER_FOLDER_ID: 'WORKORDER_FOLDER_ID',
    WORKORDER_SS_JSON: 'WORKORDER_SS_JSON',    // {호기: 스프레드시트ID} — 링크·QR이 바뀌지 않게 재사용
    ORDER_PROCESSED: 'ORDER_PROCESSED',        // {출고계획 파일ID: lastUpdated ISO}
    ORDER_SYNC: 'ORDER_SYNC_JSON',             // {소스ID: {copyId, stamp}}
    ACTUAL_PROCESSED: 'ACTUAL_PROCESSED_JSON', // {실적 파일ID: lastUpdated ISO}
    HASH_SHIP: 'HASH_SHIP',                    // 출하계획 [출하계획] 탭 내용 해시
    HASH_MASTER_PREFIX: 'HASH_M_',             // 기준정보 탭별 해시 (HASH_M_<탭이름>)
  },

  /* ── 범위 (고정 — 임의 확장 금지) ─────────────────────────────────────────── */
  ITEMS: ['7000260', '7000320', '1600190', '1900160', '1900190', '2400190', '3000173', '3700260'],
  MACHINES: Array.from({ length: 10 }, (_, i) => '집합' + String(i + 1).padStart(2, '0') + '호기'),
  CUSTOMERS: ['고객사A', '고객사B', '고객사C'],

  /* ── 계획 로직 상수 (현장 검증값 — 임의 변경 금지) ──────────────────────────── */
  PLAN: {
    START_HOUR: 8,              // 일과 08:00 시작
    INITIAL_READY_HOURS: 3,     // 계획 첫날 가동준비 시간
    CHANGE_HOURS: 1.5,          // 집합 규격교체 시간
    MAX_CHUNK_HOURS: 72,        // 한 작업 연속 배정 한도 (초과 시 보빈 단위로 분할)
    SAME_WAIT_LIMIT_HR: 168,    // 동일규격 호기 대기 한도 (7일)
    PREF_WAIT_LIMIT_HR: 48,     // 우선설비 대기 한도 (2일)
    PAYOFF2_MIN_STRANDS: 37,    // 이 가닥수 이상은 Payoff 2 호기 전용
    DEDICATED_WAIT_LIMIT_HR: 99999, // 전담 호기 예약 유지 대기 한도 (규칙 8-2, 조정 손잡이 — 현장 검증 상수 아님)
  },

  /* ── 알림 (채널 이름과 속성 키 이름만 — URL·토큰은 스크립트 속성) ──────────────── */
  NOTIFY: {
    channel: 'chat',                          // 'chat' | 'telegram' | 'mail'
    PROP_CHAT_WEBHOOK: 'CHAT_WEBHOOK_URL',    // 기본. 없으면 [설정] 탭 WEBHOOK_URL 폴백(리허설 편의)
    PROP_TELEGRAM_TOKEN: 'TELEGRAM_TOKEN',
    PROP_TELEGRAM_CHAT_ID: 'TELEGRAM_CHAT_ID',
    SETTING_WEBHOOK_KEY: 'WEBHOOK_URL',
    MAIL_TO: '',                              // channel:'mail' 폴백 수신자
  },

  /* ── 현장 배포 확장 자리 (대회 중엔 DIST_SPREADSHEET_ID 비워둠) ──────────────── */
  PUBLISH: {
    DIST_SPREADSHEET_ID: '',
    VIEWS: ['integrated', 'machineTabs', 'inventory', 'summary', 'workOrderFiles'],
    HORIZON_DAYS: 3,
    BRIEFING_HOUR: 7,
  },

  /* ── 기타 ─────────────────────────────────────────────────────────────── */
  TZ: 'Asia/Seoul',
  LOCK_WAIT_MS: 30000,     // LockService 대기 (중복 실행 방지)
  SEARCH_MAX_DEPTH: 10,    // Drive 재귀 탐색 깊이 한도
};

/* ────────────────────────────────────────────────────────────────────────────
 *  SHEET — 탭 이름 (getSheetByName은 반드시 이 상수로만 호출)
 * ──────────────────────────────────────────────────────────────────────────── */
const SHEET = {
  /* 기준정보 스프레드시트 11탭 — 출하계획·출고계획·실적은 여기 없다 */
  MASTER: {
    CAPA: 'CAPA',
    MACHINES: '집합설비현황',
    BLOCKED: '안되는품목',
    PREF: '우선순위',
    STOCK: '기초재고',
    MAX_RUN: '최대설비가동수',
    CURRENT: '설비규격교체현황',
    SHIP_PRIO: '출하우선순위',
    TARGET_STOCK: '적정재고',
    HOLIDAY: '휴무',
    SETTINGS: '설정',
  },
  /* 필수 탭 — 없으면 탭 이름을 명시한 에러로 중단. 나머지는 "있으면 적용" */
  MASTER_REQUIRED: ['CAPA', '집합설비현황', '안되는품목', '기초재고', '휴무'],

  /* 호기별 작업지시 파일의 탭 이름 (파일마다 이 탭 하나만 둔다) */
  WORKORDER: '작업지시서',

  /* 출하계획 스프레드시트 (자동 생성, 별개 파일) */
  SHIP: '출하계획',
  SHIP_HEADERS: ['품목코드', '고객사', '출하일', '출하량(kg)'],

  /* 실적 파일 (별개 파일, 매일 갱신) — 이 이름의 탭이 없으면 첫 탭을 읽는다 */
  ACTUAL_PROD: '생산실적',
  ACTUAL_SHIP: '출하실적',

  /* 결과 스프레드시트 — 통합 · 집합01~10호기 · 재고흐름 · 요약 · 작업목록 · 일별생산 · 오류 · 이력 */
  RESULT: {
    INTEGRATED: '통합',
    INVENTORY: '재고흐름',
    SUMMARY: '요약',
    JOBS: '작업목록',
    DAILY: '일별생산',
    ERRORS: '오류',
    HISTORY: '이력',
  },
};

/* [설정] 탭의 키 이름 (없는 키는 CFG 기본값) */
const SETTING_KEYS = {
  AS_OF_DATE: 'AS_OF_DATE',                       // 기준일 (비어 있으면 오늘)
  PLAN_MONTH: 'PLAN_MONTH',                       // 계획월 'YYYY-MM'
  HORIZON_DAYS: 'HORIZON_DAYS',
  BRIEFING_HOUR: 'BRIEFING_HOUR',
  ORDER_SOURCE_FOLDER_ID: 'ORDER_SOURCE_FOLDER_ID', // (선택) 출고계획 소스 폴더 URL/ID
  WEBHOOK_URL: 'WEBHOOK_URL',                     // 리허설 편의용 폴백 — public 전환 전 삭제
};

/* ────────────────────────────────────────────────────────────────────────────
 *  [▶ 생산계획] 메뉴 — 기준정보 파일을 열면 자동으로 뜬다 (바인딩 스크립트의 단순 트리거)
 *  지금 재계획 · 최신 계획 열기 · 이전 계획 보기 · 데이터 점검 · (설정) 자동 재계획 트리거 설치 · 알림 테스트
 *  메뉴 실행은 기준정보 편집 권한자용. 결과 파일·백업 사본에는 스크립트가 없다.
 * ──────────────────────────────────────────────────────────────────────────── */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('▶ 생산계획')
    .addItem('🔄 지금 재계획', 'menuRunAll')
    .addItem('📄 최신 계획 열기', 'menuOpenLatest')
    .addItem('🕘 이전 계획 보기', 'menuOpenHistory')
    .addItem('📋 호기별 작업지시서 링크', 'menuOpenWorkOrders')
    .addSeparator()
    .addItem('🔍 데이터 점검', 'menuCheckData')
    .addSubMenu(ui.createMenu('⚙️ 설정')
      .addItem('⏱️ 자동 재계획 트리거 설치', 'installTrigger')
      .addItem('🔔 알림 채널 테스트', 'testNotifyOnly')
      .addItem('🧹 출하계획 초기화(재변환 준비)', 'menuResetShipPlan'))
    .addToUi();
}

/** 메뉴 [설정 → 출하계획 초기화] — 확인 후 resetShipPlan() */
function menuResetShipPlan() {
  const ui = SpreadsheetApp.getUi();
  const answer = ui.alert('출하계획 초기화',
    '[출하계획] 탭 내용을 비우고 출고계획 처리이력을 지웁니다.\n다음 재계획(또는 testConvertOnly)에서 출고계획을 처음부터 다시 읽어 출하계획을 새로 만듭니다.\n파일·링크는 그대로 유지됩니다. 계속할까요?',
    ui.ButtonSet.YES_NO);
  if (answer !== ui.Button.YES) return;
  const r = resetShipPlan();
  ui.alert('출하계획 초기화', `완료 — 출하계획 ${r.cleared}행을 비웠습니다.\n이제 [지금 재계획] 또는 testConvertOnly를 실행하면 출고계획을 다시 읽어 변환합니다.`, ui.ButtonSet.OK);
}

/** 메뉴 [지금 재계획] — runAll('수동') 후 결과 요약과 링크를 다이얼로그로 */
function menuRunAll() {
  const ui = SpreadsheetApp.getUi();
  try {
    const r = runAll('수동');
    if (!r) { ui.alert('재계획', '다른 실행이 진행 중이라 건너뛰었습니다. 잠시 후 다시 시도하세요.', ui.ButtonSet.OK); return; }
    const html = `<div style="font-family:sans-serif;font-size:13px;line-height:1.7">
      <b>생산계획 생성 완료</b> (계획ID ${r.planId})<br>
      작업 ${r.jobs}건 · 총 ${Math.round(r.kg).toLocaleString()}kg · 교체 ${r.changeovers}회 · 납기위험 ${r.late}건 · ${r.sec}초<br><br>
      <a href="${r.resultUrl}" target="_blank">최신 계획 열기</a> &nbsp;|&nbsp; <a href="${r.backupUrl}" target="_blank">이번 실행 사본 열기</a></div>`;
    ui.showModalDialog(HtmlService.createHtmlOutput(html).setWidth(420).setHeight(140), '재계획 완료');
  } catch (e) {
    ui.alert('재계획 실패', String(e && e.message || e), ui.ButtonSet.OK);
  }
}

/** 메뉴 [최신 계획 열기] — 결과 파일 링크 */
function menuOpenLatest() {
  showLinkDialog_('최신 계획', openResult_().getUrl(), '최신 계획(생산계획) 열기');
}

/** 메뉴 [이전 계획 보기] — 결과 파일 [이력] 탭 링크 (사본URL 열을 클릭해 과거 계획 열람) */
function menuOpenHistory() {
  const rs = openResult_();
  const sh = ensureSheet_(rs, SHEET.RESULT.HISTORY);
  showLinkDialog_('이전 계획', `${rs.getUrl()}#gid=${sh.getSheetId()}`, '[이력] 탭 열기 — 사본URL 열에서 과거 계획을 엽니다');
}

/** 메뉴 [호기별 작업지시서 링크] — 현장에 나눠줄 10개 링크를 한 번에 보여준다 (QR로 뽑을 때도 여기서 복사) */
function menuOpenWorkOrders() {
  const links = workOrderLinks_();
  const made = links.filter(x => x.url);
  const rows = links.map(x => x.url
    ? `<tr><td style="padding:2px 8px">${x.machine}</td><td style="padding:2px 8px"><a href="${x.url}" target="_blank">열기</a></td></tr>`
    : `<tr><td style="padding:2px 8px">${x.machine}</td><td style="padding:2px 8px;color:#888">아직 없음</td></tr>`).join('');
  const html = `<div style="font-family:sans-serif;font-size:13px">
    <b>호기별 작업지시서</b> — ${made.length}/${links.length}개<br>
    ${made.length ? '각 호기 반장에게 해당 링크만 주면 됩니다.' : '아직 만들어지지 않았습니다. [🔄 지금 재계획]을 한 번 실행하세요.'}
    <table style="border-collapse:collapse;margin-top:6px">${rows}</table></div>`;
  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(html).setWidth(320).setHeight(360), '작업지시서 링크');
  links.forEach(x => Logger.log(`[작업지시서] ${x.machine}: ${x.url || '(없음)'}`));
}

/** 링크 앵커 다이얼로그 (getUi는 스프레드시트 밖에서 새 창을 직접 열 수 없어 앵커로 안내) */
function showLinkDialog_(title, url, text) {
  const html = `<div style="font-family:sans-serif;font-size:13px;line-height:1.8"><a href="${url}" target="_blank">${text}</a></div>`;
  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(html).setWidth(380).setHeight(80), title);
}

/**
 * 메뉴 [데이터 점검] — testLoadOnly()를 실행하고 요약을 다이얼로그로 보여준다.
 * 상세(탭별 행수·헤더)는 GAS 편집기 [실행 로그]에 남는다.
 */
function menuCheckData() {
  const ui = SpreadsheetApp.getUi();
  try {
    const r = testLoadOnly();
    const msg = [
      `기준정보 탭: ${r.tabFound}/${r.tabTotal}개` + (r.tabMissing.length ? `  (없음: ${r.tabMissing.join(', ')})` : ''),
      `출고계획 파일: ${r.orderFiles.length}개` + (r.orderFiles.length ? '  → ' + r.orderFiles.map(f => f.name).join(', ') : '  (찾지 못함)'),
      `출하계획 파일: ${r.shipExists ? '있음' : '없음 (정상 — #2 변환이 생성)'}`,
      `실적 파일: 생산실적 ${r.actual.prod ? r.actual.prod.rows + '행' : '없음'} · 출하실적 ${r.actual.ship ? r.actual.ship.rows + '행' : '없음'}`,
      r.warnings.length ? `경고 ${r.warnings.length}건 — 실행 로그 참조` : '경고 없음',
      `소요 ${r.elapsedSec}초`,
    ].join('\n');
    ui.alert('데이터 점검', msg, ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('데이터 점검 실패', String(e && e.message || e), ui.ButtonSet.OK);
  }
}

/* ────────────────────────────────────────────────────────────────────────────
 *  testLoadOnly — 단계 1 단독 테스트 (기준정보 11탭 · 출고계획 파일 탐색 · 출하계획 존재 · 실적 파일)
 *  더미데이터 기대값: CAPA 8행 · 집합설비현황 10행 · 기초재고 72행 · 설비규격교체현황 10행 ·
 *                   출하우선순위 3행 · 설정 6키 · (별도 파일) 생산실적 38행 · 출하실적 42행
 *  행수가 수백 행으로 나오면 빈 패딩 행을 못 거른 것(readTable_ 버그).
 * ──────────────────────────────────────────────────────────────────────────── */
function testLoadOnly() {
  const t0 = Date.now();
  clearWarnings_();
  const lines = [];
  const log = (s) => { lines.push(s); Logger.log(s); };

  /* 더미데이터 기준 기대 행수 (실데이터에서는 다를 수 있음 — 참고 표시만) */
  const EXPECTED = { 'CAPA': 8, '집합설비현황': 10, '기초재고': 72, '설비규격교체현황': 10, '출하우선순위': 3, '설정': 6 };
  const EXPECTED_ACTUAL = { prod: 38, ship: 42 };

  /* 1) 기준정보 */
  const master = openMaster_();
  const rootFolder = getOrderRootFolder_();
  log(`■ 기준정보: ${master.getName()}  |  탐색 루트 폴더: ${rootFolder.getName()}`);

  const tabNames = Object.keys(SHEET.MASTER).map(k => SHEET.MASTER[k]);
  const tabMissing = [];
  let tabFound = 0;
  tabNames.forEach(name => {
    const required = SHEET.MASTER_REQUIRED.indexOf(name) >= 0;
    if (!master.getSheetByName(name)) {
      tabMissing.push(name + (required ? '(필수!)' : ''));
      log(`  ✗ [${name}] 탭 없음 ${required ? '— 필수 탭입니다' : '(선택 — 없으면 미적용)'}`);
      return;
    }
    tabFound++;
    const t = readTable_(master, name, required);
    const exp = EXPECTED[name];
    const mark = exp === undefined ? '' : (t.rows.length === exp ? ' ✓' : ` ✗(기대 ${exp}행)`);
    log(`  [${name}] ${t.rows.length}행${mark}  | 헤더: ${t.headers.filter(String).join(' · ')}`);
  });

  /* 1-1) CAPA 품목코드 정규화 확인 — 숫자로 들어온 코드가 문자열 8종으로 통일되는지 */
  if (master.getSheetByName(SHEET.MASTER.CAPA)) {
    const capa = readTab_(master, SHEET.MASTER.CAPA, true);
    const codes = capa.map(r => normalizeItemCode_(r['품목코드']));
    const unknown = codes.filter(c => CFG.ITEMS.indexOf(c) < 0);
    log(`  CAPA 품목코드(정규화): ${codes.join(', ')}` + (unknown.length ? `  ✗ 범위 밖: ${unknown.join(', ')}` : '  ✓ 8종 범위 일치'));
    if (codes.some(c => !c)) warn_('testLoadOnly', 'CAPA에 품목코드가 비어 있는 행이 있습니다 — 헤더 `품목코드`(또는 …품목코드) 인식 실패 가능');
  }

  /* 1-2) [설정] 키 목록 */
  const settings = loadSettings_(master);
  const keys = Object.keys(settings);
  log(`  [설정] ${keys.length}키: ${keys.map(k => `${k}=${k === SETTING_KEYS.WEBHOOK_URL ? (settings[k] ? '(설정됨)' : '(빈값)') : fmtSettingValue_(settings[k])}`).join(' · ')}`);

  /* 2) 출고계획 파일 탐색 (파일명 패턴, 폴더 무관) */
  const orderFiles = findOrderFiles_();
  log(`■ 출고계획 파일: ${orderFiles.length}개 사용 (패턴 ${CFG.ORDER_FILE_PATTERN} — 여러 개면 최신본 1개만)`);
  orderFiles.forEach(f => log(`  - ${f.name}  | 폴더: ${f.folderPath}  | ${shortMime_(f.mimeType)}  | 수정: ${fmtDate_(f.lastUpdated, 'yyyy-MM-dd HH:mm')}`));
  if (!orderFiles.length) warn_('testLoadOnly', '출고계획 파일을 찾지 못했습니다 — 프로젝트 루트 아래에 `X월_출고계획(통합)` 스프레드시트가 있어야 합니다');

  /* 3) 출하계획 파일 (이 시점엔 없음이 정상 — #2 변환이 생성) */
  const ship = openShip_({ create: false });
  const shipExists = !!ship;
  log(`■ 출하계획 파일: ${shipExists ? '있음 → ' + ship.getUrl() : '없음 (정상 — #2 convertOrderToPlan이 생성)'}`);

  /* 4) 실적 파일 (별개 파일, 있으면 읽어 행수 확인) */
  const actualFiles = findActualFiles_();
  const actual = { prod: null, ship: null };
  [['prod', actualFiles.prod, SHEET.ACTUAL_PROD], ['ship', actualFiles.ship, SHEET.ACTUAL_SHIP]].forEach(([key, file, tab]) => {
    if (!file) { log(`■ ${tab} 파일: 없음 (규칙 11 건너뜀)`); return; }
    const t = readExternalTab_(file, tab);
    const exp = EXPECTED_ACTUAL[key];
    actual[key] = { name: file.name, rows: t.rows.length, tab: t.tabName };
    log(`■ ${tab} 파일: ${file.name} | 폴더: ${file.folderPath} | ${shortMime_(file.mimeType)} | 탭 [${t.tabName}] ${t.rows.length}행${t.rows.length === exp ? ' ✓' : ` ✗(기대 ${exp}행)`} | 헤더: ${t.headers.filter(String).join(' · ')}`);
  });

  const warnings = getWarnings_();
  const elapsedSec = Math.round((Date.now() - t0) / 100) / 10;
  log(`■ 완료: 기준정보 ${tabFound}/${tabNames.length}탭 · 출고계획 ${orderFiles.length}개 · 출하계획 ${shipExists ? '있음' : '없음'} · 경고 ${warnings.length}건 · ${elapsedSec}초`);
  if (warnings.length) warnings.forEach(w => Logger.log(`  ⚠ [${w.stage}] ${w.message}`));

  return { tabTotal: tabNames.length, tabFound, tabMissing, orderFiles, shipExists, actual, warnings, elapsedSec, lines };
}

/** 로그 표시용 — Date는 yyyy-MM-dd, 그 외는 문자열 그대로 */
function fmtSettingValue_(v) {
  if (v instanceof Date) return fmtDate_(v, 'yyyy-MM-dd');
  return String(v);
}

/** 로그 표시용 MIME 축약 */
function shortMime_(mime) {
  if (mime === MimeType.GOOGLE_SHEETS) return 'Google 스프레드시트';
  if (mime === MimeType.MICROSOFT_EXCEL || mime === MimeType.MICROSOFT_EXCEL_LEGACY) return 'Excel(xlsx — 임시 변환해 읽음)';
  return mime;
}

/* ════════════════════════════════════════════════════════════════════════════
 *  출고계획 스프레드시트 → 출하계획 스프레드시트 변환
 * ----------------------------------------------------------------------------
 *  입력  : 탐색 루트 아래 어느 폴더든 파일명이 `X월_출고계획(통합)…` 인 파일 (Google 스프레드시트 또는 xlsx).
 *          가로 날짜형 — 품목코드 · (규격) · 구분(고객사) · (발주량) · (합계) · 날짜열… (헤더 1행)
 *  출력  : 출하계획 스프레드시트 [출하계획] 탭 — 품목코드 · 고객사 · 출하일 · 출하량(kg) (세로형, 4열)
 *  규칙  : (품목코드, 고객사, 출하일) 키로 upsert(같은 키는 덮어쓰기, 새 키는 추가) → 출하일·품목코드·고객사 순 정렬 →
 *          탭 전체를 배치로 다시 씀. 같은 파일을 두 번 처리해도 결과가 같다(멱등).
 *          출고계획 파일은 옮기지도 이름을 바꾸지도 않는다(판독 파트 소유). 대신 처리이력 {파일ID: lastUpdated}로 재처리를 막는다.
 * ════════════════════════════════════════════════════════════════════════════ */

/* 출고계획 표의 열 이름
 *  고객사 열은 양식마다 이름이 다르다 — 더미데이터는 `구분`, 판독 파트 산출물은 `구분(업체)`,
 *  실데이터에서는 `고객사`·`납품처`·`업체`로 올 수 있다. 그래서 이름이 정확히 같은지 보지 않고
 *  '구분'·'고객'·'납품'·'업체' 중 하나를 포함하는 헤더를 고객사 열로 본다. */
const ORDER_COLS_ = {
  CUSTOMER_HINTS: ['구분', '고객', '납품', '업체'],
  IGNORE: ['NO', '규격', '발주량', '합계', '품목명', '품명', '비고', '단위'],   // 계산에 쓰지 않는 열 (검산용)
};

/** 헤더가 고객사 열인지 */
function isCustomerHeader_(h) {
  if (typeof h !== 'string' || !h) return false;
  for (let i = 0; i < ORDER_COLS_.CUSTOMER_HINTS.length; i++) {
    if (h.indexOf(ORDER_COLS_.CUSTOMER_HINTS[i]) >= 0) return true;
  }
  return false;
}

/**
 * 출고계획 → 출하계획 변환 (변환 단계의 진입점. checkAndRunOnUpdate·runAll도 이것을 호출)
 * @param {{force?: boolean, label?: string}} [opts]
 *   force=true면 처리이력을 무시하고 찾은 파일 전부를 다시 읽는다.
 *   label은 알림에 붙일 실행 주체 표시 — 트리거가 돌린 것이면 runAll이 '🤖 [자동실행]'을 넘긴다.
 *   비우면 사람이 직접 실행한 것으로 보고 '👤 [수동실행]'을 쓴다 (편집기·메뉴 실행).
 * @returns {{processed, skipped, added, updated, unchanged, meltedRows, shipRows, shipKg, shipUrl, sync, elapsedSec}}
 */
function convertOrderToPlan(opts) {
  opts = opts || {};
  const t0 = Date.now();
  const result = { processed: [], skipped: [], added: 0, updated: 0, removed: 0, removedKg: 0, unchanged: 0, meltedRows: 0, shipRows: 0, shipKg: 0, shipUrl: '', sync: null, elapsedSec: 0 };
  const finish = () => {
    result.elapsedSec = Math.round((Date.now() - t0) / 100) / 10;
    const shipPart = result.processed.length ? `출하계획 ${result.shipRows}행 ${Math.round(result.shipKg).toLocaleString()}kg` : '출하계획 변경 없음';
    Logger.log(`[변환] 처리 ${result.processed.length}개 · 건너뜀 ${result.skipped.length}개 · 추가 ${result.added} · 변경 ${result.updated} · 삭제 ${result.removed} · 동일 ${result.unchanged} · ${shipPart} · ${result.elapsedSec}초`);
    return result;
  };

  const master = openMaster_();
  const settings = loadSettings_(master);
  const planMonth = getPlanMonth_(settings);
  if (!planMonth) warn_('변환', '[설정] PLAN_MONTH가 없어 MM/DD 형식 날짜 헤더는 올해로 해석합니다');
  const capaCodes = {};
  readTab_(master, SHEET.MASTER.CAPA, true).forEach(r => { const c = normalizeItemCode_(r['품목코드']); if (c) capaCodes[c] = true; });

  /* 0) (선택) 외부 소스 폴더 동기화 — 판독 파트가 프로젝트 루트 밖 폴더를 쓸 때만 동작 */
  result.sync = syncOrderFiles_(settings);

  /* 1) 출고계획 파일 탐색 → 처리이력과 비교해 새 파일·수정된 파일만 고른다 */
  const files = findOrderFiles_();
  if (!files.length) {
    warn_('변환', `출고계획 파일을 찾지 못했습니다 — 탐색 루트 "${getOrderRootFolder_().getName()}" 아래에 파일명 ${CFG.ORDER_FILE_PATTERN} 파일이 있어야 합니다`);
    return finish();
  }
  const history = getJsonProp_(CFG.PROP.ORDER_PROCESSED, {});
  const toProcess = [];
  files.forEach(f => {
    const stamp = f.lastUpdated.toISOString();
    if (!opts.force && history[f.id] === stamp) {
      result.skipped.push({ name: f.name, folderPath: f.folderPath, reason: '변경 없음(처리이력과 lastUpdated 동일)' });
      Logger.log(`[변환] skip: ${f.folderPath}/${f.name} — 처리이력과 동일`);
      return;
    }
    toProcess.push({ file: f, stamp });
  });
  if (!toProcess.length) return finish();

  /* 2) melt — 가로 날짜형 → 세로 4열. 파일 하나가 실패해도 나머지는 계속(경고만) */
  const melted = [];
  const done = [];
  toProcess.forEach(({ file, stamp }) => {
    try {
      const rows = meltOrderFile_(file, planMonth, capaCodes);
      if (!rows.length) {
        warn_('변환', `${file.folderPath}/${file.name}: 변환 결과 0건 — 처리이력에 기록하지 않습니다`);
        result.skipped.push({ name: file.name, folderPath: file.folderPath, reason: '변환 0건' });
        return;
      }
      rows.forEach(r => melted.push(r));
      done.push({ file, stamp, rows: rows.length, kg: rows.reduce((s, r) => s + r.출하량, 0) });
      Logger.log(`[변환] ${file.folderPath}/${file.name} (${shortMime_(file.mimeType)}) → ${rows.length}건 ${Math.round(done[done.length - 1].kg).toLocaleString()}kg`);
    } catch (e) {
      warn_('변환', `${file.folderPath}/${file.name} 읽기 실패 — 이 파일만 건너뜁니다: ${e.message}`);
      result.skipped.push({ name: file.name, folderPath: file.folderPath, reason: '읽기 실패: ' + e.message });
    }
  });
  result.meltedRows = melted.length;
  if (!melted.length) {
    warn_('변환', '변환 결과가 0건입니다 — 출하계획 파일을 쓰지 않고 처리이력도 갱신하지 않습니다');
    return finish();
  }

  /* 3) 출하계획 파일 upsert (없으면 이때 생성) */
  const ship = openShip_({ create: true });
  /* 이번에 읽은 출고계획이 담당하는 달 — 그 달 안에서 파일에 없는 행은 취소된 출고로 보고 지운다 */
  const months = [];
  melted.forEach(m => { const mk = monthKey_(m.출하일); if (months.indexOf(mk) < 0) months.push(mk); });
  const up = upsertShipRows_(ship, melted, { authoritativeMonths: months });
  result.added = up.added; result.updated = up.updated; result.unchanged = up.unchanged;
  result.removed = up.removed; result.removedKg = up.removedKg;
  up.removedRows.forEach(r => Logger.log(`[변환] 삭제(출고계획에서 빠짐): ${r.품목코드} ${r.고객사} ${dateKey_(r.출하일)} ${Math.round(r.출하량).toLocaleString()}kg`));
  result.shipRows = up.total; result.shipKg = up.totalKg; result.shipUrl = ship.getUrl();

  /* 4) 처리이력·출하계획 해시 저장 — 같은 변경으로 트리거가 두 번 돌지 않게 */
  done.forEach(d => {
    history[d.file.id] = d.stamp;
    result.processed.push({ name: d.file.name, folderPath: d.file.folderPath, rows: d.rows, kg: d.kg });
  });
  setJsonProp_(CFG.PROP.ORDER_PROCESSED, history);
  setProp_(CFG.PROP.HASH_SHIP, hashValues_(ship, SHEET.SHIP));

  /* 5) 알림 */
  const msg = [
    /* 첫 줄 ✅ — 채팅방에서 생산계획 생성 알림과 한눈에 구분되게 (사용자 지시 2026-09-04) */
    `✅ 출고계획 변환 완료 (${fmtDate_(new Date(), 'yyyy-MM-dd HH:mm')})`,
    ...result.processed.map(p => `- ${p.folderPath}/${p.name}: ${p.rows}건 ${Math.round(p.kg).toLocaleString()}kg`),
    `출하계획: 추가 ${result.added}건 · 변경 ${result.updated}건 · 삭제 ${result.removed}건 · 동일 ${result.unchanged}건 → 총 ${result.shipRows}행 ${Math.round(result.shipKg).toLocaleString()}kg`,
    `출하계획 파일: ${result.shipUrl}`,
  ].join('\n');
  /* 실행 주체 표시 — 트리거로 돌아온 변환은 자동, 사람이 직접 실행한 것은 수동.
   * runAll이 감지 사유로 판단한 라벨을 넘겨준다 (사용자 지시 2026-09-04). */
  notify_(msg, opts.label || '👤 [수동실행]');
  return finish();
}

/**
 * 출고계획 파일 한 개를 세로형으로 melt한다.
 *  - 첫 탭을 읽는다. 헤더 행 = 위에서 10행 안에 `품목코드`(또는 …품목코드) 셀이 있는 첫 행
 *  - 날짜 열 판별: Date 셀(파일 타임존 기준으로 일자 추출) · 'yyyy-MM-dd' · 'M/D/yyyy' · 'MM/DD'(연도는 PLAN_MONTH, 12↔1월 보정)
 *  - 값이 0·빈 셀은 제외. 같은 (품목코드, 고객사, 출하일)이 한 파일에 두 번 나오면 합산 + 경고
 *  - 품목코드가 8종 밖이거나 CAPA에 없으면 행은 살리고 경고 (배수 검증은 하지 않는다 — 실데이터는 배수가 아닐 수 있음)
 * @returns {Array<{품목코드:string, 고객사:string, 출하일:Date, 출하량:number}>}
 */
function meltOrderFile_(file, planMonth, capaCodes) {
  return withSpreadsheetFile_(file, function (ss) {
    const tz = ss.getSpreadsheetTimeZone();
    const sheets = ss.getSheets();
    const sh = sheets[0];                                    // 첫 탭만 읽는다 (탭이 여럿이면 로그로 알림)
    const values = sh.getDataRange().getValues();
    const where = `"${file.name}"[${sh.getName()}]`;
    Logger.log(`[변환] ${where} 읽기 — 탭 ${sheets.length}개: ${sheets.map(s => s.getName()).join(' · ')}` + (sheets.length > 1 ? ' (첫 탭만 사용)' : ''));

    // 헤더 행 찾기
    let hi = -1;
    for (let r = 0; r < Math.min(values.length, 10); r++) {
      if (values[r].some(c => /품목코드$/.test(String(c).trim()))) { hi = r; break; }
    }
    if (hi < 0) throw new Error(`${where}: 위 10행 안에 '품목코드' 헤더가 없습니다`);
    const headers = values[hi];
    const headerStr = headers.map(h => (h instanceof Date) ? h : String(h === null || h === undefined ? '' : h).trim());

    const itemIdx = headerStr.findIndex(h => h === '품목코드');
    const itemIdx2 = itemIdx >= 0 ? itemIdx : headerStr.findIndex(h => typeof h === 'string' && /품목코드$/.test(h));
    const custIdx = headerStr.findIndex(isCustomerHeader_);
    if (custIdx < 0) throw new Error(`${where}: 고객사 열을 찾지 못했습니다 (헤더에 ${ORDER_COLS_.CUSTOMER_HINTS.join('/')} 중 하나가 있어야 합니다) — 헤더: ${headerStr.filter(h => typeof h === 'string' && h).join(' · ')}`);

    // 날짜 열 판별
    const dateCols = [];
    headerStr.forEach((h, c) => {
      if (c === itemIdx2 || c === custIdx) return;
      if (typeof h === 'string' && (!h || ORDER_COLS_.IGNORE.indexOf(h) >= 0)) return;
      const d = parseHeaderDate_(h, tz, planMonth);
      if (d) dateCols.push({ idx: c, date: d, key: dateKey_(d) });
      else Logger.log(`[변환] ${where}: 날짜가 아닌 열 "${h}" 무시`);
    });
    if (!dateCols.length) throw new Error(`${where}: 날짜 열을 하나도 인식하지 못했습니다 — 헤더: ${headerStr.map(String).join(' · ')}`);

    // melt
    const map = {};
    const warnedCode = {};
    let dup = 0, noCust = 0;
    for (let r = hi + 1; r < values.length; r++) {
      const row = values[r];
      if (isBlankRow_(row)) continue;
      const code = normalizeItemCode_(row[itemIdx2]);
      if (!code) continue;                                  // 합계행 등 품목코드 없는 행
      const cust = normalizeCustomer_(row[custIdx]);        // 'A사' → '고객사A'
      if (!cust) { noCust++; continue; }                    // 메모 행(품목코드 자리에 안내문, 고객사 없음) 등
      if (!warnedCode['cust:' + cust] && CFG.CUSTOMERS.indexOf(cust) < 0) {
        warnedCode['cust:' + cust] = true;
        warn_('변환', `${where}: 고객사 "${cust}"는 기준 표기(${CFG.CUSTOMERS.join('/')})가 아닙니다 — 행은 유지, [출하우선순위]와 매칭되지 않을 수 있음`);
      }
      if (!warnedCode[code]) {
        warnedCode[code] = true;
        if (CFG.ITEMS.indexOf(code) < 0) warn_('변환', `${where}: 품목코드 ${code}는 범위(8종) 밖입니다 — 행은 유지`);
        else if (!capaCodes[code]) warn_('변환', `${where}: 품목코드 ${code}가 [CAPA]에 없습니다 — 계획 시 kg/hr를 구할 수 없음`);
      }
      dateCols.forEach(dc => {
        const q = toNumber_(row[dc.idx]);
        if (!q) return;                                     // 0·빈 셀 제외
        const key = `${code}|${cust}|${dc.key}`;
        if (map[key]) { map[key].출하량 += q; dup++; }
        else map[key] = { 품목코드: code, 고객사: cust, 출하일: dc.date, 출하량: q };
      });
    }
    if (dup) warn_('변환', `${where}: 같은 (품목코드, 고객사, 출하일)이 ${dup}번 중복되어 합산했습니다`);
    if (noCust) warn_('변환', `${where}: 고객사가 빈 행 ${noCust}건을 건너뛰었습니다`);
    return Object.keys(map).map(k => map[k]);
  });
}

/**
 * 출고계획 날짜 헤더 → 스크립트 타임존 00:00 Date. 날짜가 아니면 null.
 *  Date 셀은 파일 타임존(tz)으로 일자를 뽑는다(UTC 경유로 하루 밀리는 사고 방지).
 *  'MM/DD'·'M/D'·'M월 D일'은 연도가 없으므로 PLAN_MONTH로 정한다(12월↔1월 걸침 보정).
 */
function parseHeaderDate_(h, tz, planMonth) {
  if (h instanceof Date) return normalizeDateCell_(h, tz, '출고계획 날짜 헤더');
  if (typeof h === 'number') {
    // 스프레드시트 날짜 일련번호로 보이는 범위(약 1982~2119년)만 날짜로 본다 — `34`처럼 단순 숫자 헤더(열 개수 등)는 무시
    return (h >= 30000 && h <= 80000) ? normalizeDateCell_(h, tz, '출고계획 날짜 헤더') : null;
  }
  const s = String(h).trim();
  if (!s) return null;
  let m = s.match(/^(\d{1,2})[/.-](\d{1,2})$/) || s.match(/^(\d{1,2})월\s*(\d{1,2})일$/);
  if (m) {
    const month = +m[1], day = +m[2];
    if (month < 1 || month > 12 || day < 1 || day > 31) return null;
    return new Date(yearForMonth_(planMonth, month), month - 1, day);
  }
  /* `1일` `2일` … — 판독 파트 산출물의 날짜 헤더. 월·연도가 없으므로 [설정] PLAN_MONTH로만 해석한다 */
  m = s.match(/^(\d{1,2})\s*일$/);
  if (m) {
    const day = +m[1];
    if (day < 1 || day > 31) return null;
    if (!planMonth) { warn_('변환', `날짜 헤더 "${s}"는 일자만 있어 [설정] PLAN_MONTH가 없으면 해석할 수 없습니다`); return null; }
    return new Date(planMonth.year, planMonth.month - 1, day);
  }
  if (/^\d{4}[-./]\d{1,2}[-./]\d{1,2}/.test(s) || /^\d{1,2}[/.-]\d{1,2}[/.-]\d{4}$/.test(s) || /^\d{8}$/.test(s)) {
    try { return normalizeDateCell_(s, tz, '출고계획 날짜 헤더'); } catch (e) { return null; }
  }
  return null;
}

/* ────────────────────────────────────────────────────────────────────────────
 *  출하계획 스프레드시트 [출하계획] 탭 읽기·쓰기 (계획 엔진(#3)도 readShipRows_로 수요를 읽는다)
 * ──────────────────────────────────────────────────────────────────────────── */

/** [출하계획] 탭 → 정규화된 행 배열 (품목코드 문자열 · 고객사 trim · 출하일 Date 00:00 · 출하량 숫자). 잘못된 행은 경고 후 제외 */
function readShipRows_(ship) {
  const tz = ship.getSpreadsheetTimeZone();
  const t = readTable_(ship, SHEET.SHIP, false);
  const out = [];
  t.rows.forEach((r, i) => {
    const code = normalizeItemCode_(r['품목코드']);
    const cust = normalizeCustomer_(r['고객사']);           // 판독 파트가 'A사'로 append해도 '고객사A'로 통일
    let d = null;
    try { d = normalizeDateCell_(r['출하일'], tz, `[출하계획] ${i + 2}행 출하일`); } catch (e) { warn_('출하계획', e.message); }
    const qty = toNumber_(r['출하량(kg)']);
    if (!code || !cust || !d) { warn_('출하계획', `[출하계획] ${i + 2}행: 품목코드/고객사/출하일이 비어 있어 제외 (${code || '-'} · ${cust || '-'})`); return; }
    out.push({ 품목코드: code, 고객사: cust, 출하일: d, 출하량: qty });
  });
  return out;
}

/** upsert 키 */
function shipKey_(r) { return `${r.품목코드}|${r.고객사}|${dateKey_(r.출하일)}`; }

/** 정렬: 출하일 → 품목코드 → 고객사 */
function sortShipRows_(a, b) {
  const d = a.출하일.getTime() - b.출하일.getTime();
  if (d) return d;
  if (a.품목코드 !== b.품목코드) return a.품목코드 < b.품목코드 ? -1 : 1;
  return a.고객사.localeCompare(b.고객사, 'ko');
}

/**
 * [출하계획] 탭에 (품목코드, 고객사, 출하일) 키로 upsert하고 정렬해 탭 전체를 다시 쓴다.
 *  같은 키 → 출하량 덮어쓰기(값이 같으면 '동일'), 새 키 → 추가.
 *
 * opts.authoritativeMonths: ['2026-09', …]
 *  출고계획 파일은 그 달 전체를 담은 계획표이므로, **그 달에 관해서는 파일이 유일한 근거**다.
 *  파일에서 빠진 (품목·고객사·출하일)은 출고가 취소·이동된 것이므로 출하계획에서도 지운다.
 *  지우지 않으면 취소된 출고가 계속 남아 생산계획이 그만큼 과잉으로 잡힌다.
 *  파일이 다루지 않는 달의 행은 건드리지 않는다 — 판독 파트가 직접 넣은 다른 달 발주를 지우면 안 된다.
 *  opts가 없으면(예: demoAddOrder) 지우지 않고 기존처럼 추가·변경만 한다.
 */
function upsertShipRows_(ship, melted, opts) {
  opts = opts || {};
  const map = {};
  readShipRows_(ship).forEach(r => { map[shipKey_(r)] = r; });
  let added = 0, updated = 0, unchanged = 0;
  const inFile = {};
  melted.forEach(m => {
    const k = shipKey_(m);
    inFile[k] = true;
    const cur = map[k];
    if (!cur) { map[k] = { 품목코드: m.품목코드, 고객사: m.고객사, 출하일: m.출하일, 출하량: m.출하량 }; added++; }
    else if (Math.abs(cur.출하량 - m.출하량) > 1e-9) { cur.출하량 = m.출하량; updated++; }
    else unchanged++;
  });

  /* 파일이 담당하는 달 안에서, 파일에 없는 행을 지운다 */
  const scope = opts.authoritativeMonths || null;
  const removedRows = [];
  if (scope && scope.length) {
    Object.keys(map).forEach(k => {
      if (inFile[k]) return;
      const r = map[k];
      if (scope.indexOf(monthKey_(r.출하일)) < 0) return;
      removedRows.push(r);
      delete map[k];
    });
  }

  const rows = Object.keys(map).map(k => map[k]).sort(sortShipRows_);
  writeShipRows_(ship, rows);
  const removedKg = removedRows.reduce((s, r) => s + r.출하량, 0);
  return {
    added, updated, unchanged,
    removed: removedRows.length, removedKg, removedRows,
    total: rows.length, totalKg: rows.reduce((s, r) => s + r.출하량, 0),
  };
}

/** [출하계획] 탭 전체 재작성 — clearContents 후 setValues 1회. 품목코드 텍스트 · 출하일 yyyy-mm-dd · 출하량 천단위 */
function writeShipRows_(ship, rows) {
  const sh = ensureSheet_(ship, SHEET.SHIP);
  sh.clearContents();
  const n = rows.length;
  if (n) {
    sh.getRange(2, 1, n, 1).setNumberFormat('@');           // 값을 쓰기 전에 텍스트 서식 → 앞자리 0·숫자 변환 방지
    sh.getRange(2, 3, n, 1).setNumberFormat('yyyy-mm-dd');
    sh.getRange(2, 4, n, 1).setNumberFormat('#,##0');
  }
  const values = [SHEET.SHIP_HEADERS].concat(rows.map(r => [r.품목코드, r.고객사, r.출하일, r.출하량]));
  sh.getRange(1, 1, values.length, SHEET.SHIP_HEADERS.length).setValues(values);
  sh.getRange(1, 1, 1, SHEET.SHIP_HEADERS.length).setFontWeight('bold');
  sh.setFrozenRows(1);
}

/* ────────────────────────────────────────────────────────────────────────────
 *  외부 소스 폴더 동기화 (선택) — 판독 파트가 프로젝트 루트 "밖" 폴더에서 출고계획을 관리할 때
 *  소스: CFG.ORDER_SOURCE_FOLDER_ID(우선) 또는 [설정] ORDER_SOURCE_FOLDER_ID (폴더 URL/ID 모두 허용)
 *  동작: 소스 아래를 재귀 탐색 → **`X월_출고계획(통합)` 정식 이름**이면서 lastUpdated가 바뀐 파일만
 *        프로젝트 루트의 `출고계획/`로 복사한다 (사용자 지시 2026-09-04 — 정식 이름만 가져온다).
 *        복사본이 이미 있으면 파일을 새로 만들지 않고 첫 탭 내용만 덮어써 복사본 URL을 유지하되,
 *        **이름은 원본 이름으로 맞춘다** — 원본이 이름을 바꿔도 사본이 옛 이름에 굳지 않게.
 *  원본 폴더·파일은 읽기만 한다 — 이동·이름변경·수정 절대 금지. 실패는 경고 후 계속(변환을 죽이지 않음).
 * ──────────────────────────────────────────────────────────────────────────── */
function syncOrderFiles_(settings) {
  const res = { enabled: false, copied: 0, updated: 0, skipped: 0, failed: 0, reason: '' };
  if (!CFG.ORDER_SYNC_ENABLED) { res.reason = '비활성(CFG.ORDER_SYNC_ENABLED=false — 판독 파트 폴더 연동은 사용자 지시 시 켠다)'; Logger.log('[동기화] ' + res.reason); return res; }
  const raw = CFG.ORDER_SOURCE_FOLDER_ID || getSetting_(settings, SETTING_KEYS.ORDER_SOURCE_FOLDER_ID, '');
  const sourceId = parseDriveId_(raw);
  if (!sourceId) { res.reason = '소스 폴더 미지정 — 동기화 생략'; return res; }
  const root = getOrderRootFolder_();
  if (sourceId === root.getId()) { res.reason = '소스 폴더가 탐색 루트와 같음 — 중복 복사 방지로 생략'; Logger.log('[동기화] ' + res.reason); return res; }

  let source;
  try { source = DriveApp.getFolderById(sourceId); }
  catch (e) { res.reason = '소스 폴더 접근 실패'; warn_('동기화', `소스 폴더(${sourceId})에 접근할 수 없습니다 — 공유 권한 확인. 동기화 생략`); return res; }
  res.enabled = true;

  /* 소스에서는 `X월_출고계획(통합)` 정식 이름만 가져온다 — 사본·임시본이 딸려 오지 않게 (사용자 지시 2026-09-04) */
  const srcFiles = findFilesByPattern_(source, CFG.ORDER_FILE_PATTERN_MAIN);
  if (!srcFiles.length) {
    res.reason = `소스 폴더 "${source.getName()}"에 정식 이름(X월_출고계획(통합)) 파일 없음`;
    Logger.log('[동기화] ' + res.reason);
    return res;
  }
  const copyFolder = getOrCreateSubFolder_(getMasterFolder_(), CFG.FILE_NAMES.ORDER_COPY_FOLDER);
  const sync = getJsonProp_(CFG.PROP.ORDER_SYNC, {});
  let changed = false;

  srcFiles.forEach(f => {
    const stamp = f.lastUpdated.toISOString();
    let entry = sync[f.id];
    /* 이력이 없거나 가리키는 사본이 사라졌으면, 폴더에서 같은 이름의 사본을 찾아 이어 쓴다.
     * 이력만 믿으면 속성이 지워졌을 때 사본을 새로 만들어 파일이 여러 개로 갈라진다. */
    if (!entry || !entry.copyId || checkDriveItem_(entry.copyId, 'file') === 'missing') {
      const found = findExistingCopy_(copyFolder, f.name);
      if (found) {
        entry = { copyId: found.getId(), stamp: '', name: f.name };
        Logger.log(`[동기화] 기존 사본을 이어서 씁니다 — 출고계획/${found.getName()}`);
      }
    }
    const copyAlive = entry && entry.copyId && checkDriveItem_(entry.copyId, 'file') !== 'missing';
    if (entry && copyAlive && entry.stamp === stamp) { res.skipped++; return; }
    try {
      if (copyAlive) {
        // 복사본 첫 탭 내용만 덮어쓰기 (URL 유지)
        const copySs = SpreadsheetApp.openById(entry.copyId);
        const values = withSpreadsheetFile_(f, ss => ss.getSheets()[0].getDataRange().getValues());
        const target = copySs.getSheets()[0];
        target.clearContents();
        if (values.length) target.getRange(1, 1, values.length, values[0].length).setValues(values);
        /* 사본 이름을 원본 이름으로 맞춘다 — 파일 ID는 그대로라 URL은 유지된다.
         * 이게 없으면 원본이 이름을 바꿔도 사본은 처음 복사할 때의 옛 이름에 굳는다. */
        const wantName = f.name.replace(/\.xlsx?$/i, '');
        if (copySs.getName() !== wantName) {
          const before = copySs.getName();
          copySs.rename(wantName);
          Logger.log(`[동기화] 사본 이름 정정: ${before} → ${wantName}`);
        }
        sync[f.id] = { copyId: entry.copyId, stamp, name: f.name };
        res.updated++;
        Logger.log(`[동기화] 갱신: ${f.folderPath}/${f.name} → 출고계획/${copySs.getName()}`);
      } else {
        // 새 복사본 — Google 스프레드시트로 (xlsx는 Drive 고급 서비스로 변환 복사, 없으면 그대로 복사)
        const name = f.name.replace(/\.xlsx?$/i, '');
        let copyId;
        if (f.mimeType === MimeType.GOOGLE_SHEETS || typeof Drive === 'undefined' || !Drive.Files) {
          copyId = DriveApp.getFileById(f.id).makeCopy(f.mimeType === MimeType.GOOGLE_SHEETS ? name : f.name, copyFolder).getId();
        } else {
          copyId = Drive.Files.copy({ name, mimeType: MimeType.GOOGLE_SHEETS, parents: [copyFolder.getId()] }, f.id).id;
        }
        sync[f.id] = { copyId, stamp, name: f.name };
        res.copied++;
        Logger.log(`[동기화] 복사: ${f.folderPath}/${f.name} → 출고계획/${name}`);
      }
      changed = true;
    } catch (e) {
      res.failed++;
      warn_('동기화', `${f.folderPath}/${f.name} 동기화 실패 — 건너뜀: ${e.message}`);
    }
  });
  if (changed) { setJsonProp_(CFG.PROP.ORDER_SYNC, sync); invalidateFileListCache_(); }
  Logger.log(`[동기화] 소스 "${source.getName()}": 복사 ${res.copied} · 갱신 ${res.updated} · 동일 ${res.skipped} · 실패 ${res.failed}`);
  return res;
}

/* ════════════════════════════════════════════════════════════════════════════
 *  데이터 로드 (기준정보 11탭 + 출하계획) → 계획 엔진 입력 객체
 * ----------------------------------------------------------------------------
 *  모든 탭은 readTab_(헤더 이름 기준)으로 읽고, 품목코드·호기·고객사는 normalize* 로 통일한다.
 *  필수 탭(CAPA·집합설비현황·안되는품목·기초재고·휴무)이 없으면 탭 이름을 명시한 에러로 중단.
 * ════════════════════════════════════════════════════════════════════════════ */

/**
 * @param {{skipShip?: boolean}} [opts] skipShip=true면 출하계획 파일을 읽지 않는다(기준정보만 점검할 때)
 * @returns {{master, settings, planMonth, asOf, capa, machines, machineList, blocked, pref, stock, maxRun, currentItem, shipPrio, targetStock, holidays, ship}}
 */
function loadData_(opts) {
  opts = opts || {};
  const master = openMaster_();
  const tz = master.getSpreadsheetTimeZone();
  const settings = loadSettings_(master);
  const planMonth = getPlanMonth_(settings);
  const M = SHEET.MASTER;

  /* CAPA — 품목 → {규격, 적용설비, 가닥수, kg/hr, 보빈 중량, 조장} (kg/hr 비면 kg/일/24hr ÷ 24) */
  const capa = {};
  readTab_(master, M.CAPA, true).forEach((r, i) => {
    const code = normalizeItemCode_(r['품목코드']);
    if (!code) return;
    let kgPerHr = toNumber_(field_(r, ['kg/hr', 'KG/HR', 'kg/h']));
    if (!kgPerHr) kgPerHr = toNumber_(field_(r, ['kg/일/24hr', 'kg/일', 'KG/일'])) / 24;
    const bobbinKg = toNumber_(field_(r, ['중량(KG)', '중량(kg)', '중량']));
    if (!kgPerHr || !bobbinKg) throw new Error(`[CAPA] ${i + 2}행 ${code}: kg/hr(${kgPerHr}) 또는 중량(KG)(${bobbinKg})이 비어 있어 계획을 세울 수 없습니다`);
    capa[code] = {
      code,
      규격: String(field_(r, ['규격']) || '').trim(),
      적용설비: String(field_(r, ['적용설비']) || '').trim(),
      가닥수: toNumber_(field_(r, ['가닥수'])),
      kgPerHr,
      bobbinKg,
      lengthM: toNumber_(field_(r, ['조장(m)', '조장'])),
    };
  });
  if (!Object.keys(capa).length) throw new Error('[CAPA]에 품목이 없습니다');

  /* 집합설비현황 — 호기 → {적용설비, Payoff수} */
  const machines = {};
  const machineList = [];
  readTab_(master, M.MACHINES, true).forEach(r => {
    const name = normalizeMachine_(r['호기']);
    if (!name || machines[name]) return;
    machines[name] = { name, 적용설비: String(field_(r, ['적용설비']) || '').trim(), payoff: toNumber_(field_(r, ['Payoff수', 'payoff수', 'Payoff'])) };
    machineList.push(name);
  });
  if (!machineList.length) throw new Error('[집합설비현황]에 호기가 없습니다');

  /* 안되는품목 — 절대차단 (필수) */
  const blocked = {};
  readTab_(master, M.BLOCKED, true).forEach(r => {
    const c = normalizeItemCode_(r['품목코드']), m = normalizeMachine_(r['호기']);
    if (c && m) blocked[c + '|' + m] = true;
  });

  /* 우선순위 — 품목 → [호기…] (행 순서 = 우선순위) */
  const pref = {};
  readTab_(master, M.PREF, false).forEach(r => {
    const c = normalizeItemCode_(r['품목코드']), m = normalizeMachine_(r['호기']);
    if (c && m) (pref[c] = pref[c] || []).push(m);
  });

  /* 기초재고 — 보빈 1행이 여러 개 → 품목별 합산 */
  const stock = {};
  readTab_(master, M.STOCK, true).forEach(r => {
    const c = normalizeItemCode_(r['품목코드']);
    if (c) stock[c] = (stock[c] || 0) + toNumber_(field_(r, ['기초재고', '재고', '수량']));
  });

  /* 최대설비가동수 */
  const maxRun = {};
  readTab_(master, M.MAX_RUN, false).forEach(r => {
    const c = normalizeItemCode_(r['품목코드']), n = toNumber_(field_(r, ['최대 가동수', '최대가동수', '최대설비가동수']));
    if (c && n > 0) maxRun[c] = n;
  });

  /* 설비규격교체현황 — 호기 → 현재 걸린 품목 */
  const currentItem = {};
  readTab_(master, M.CURRENT, false).forEach(r => {
    const m = normalizeMachine_(r['호기']), c = normalizeItemCode_(r['품목코드']);
    if (m && c) currentItem[m] = c;
  });

  /* 출하우선순위 — 고객사 → 순위 */
  const shipPrio = {};
  readTab_(master, M.SHIP_PRIO, false).forEach(r => {
    const cust = normalizeCustomer_(r['고객사']);
    if (cust) shipPrio[cust] = toNumber_(field_(r, ['순위'])) || 999;
  });

  /* 적정재고 — 일부 품목만 등재될 수 있음 */
  const targetStock = {};
  readTab_(master, M.TARGET_STOCK, false).forEach(r => {
    const c = normalizeItemCode_(r['품목코드']), n = toNumber_(field_(r, ['적정 재고량', '적정재고량', '적정재고']));
    if (c && n > 0) targetStock[c] = n;
  });

  /* 휴무 — { 'yyyy-MM-dd': 명칭 } (필수) */
  const holidays = {};
  readTab_(master, M.HOLIDAY, true).forEach((r, i) => {
    const d = normalizeDateCell_(r['날짜'], tz, `[휴무] ${i + 2}행 날짜`);
    if (d) holidays[dateKey_(d)] = String(r['명칭'] || '').trim();
  });

  /* 기준일 — [설정] AS_OF_DATE, 비어 있으면 오늘 */
  const asOfRaw = getSetting_(settings, SETTING_KEYS.AS_OF_DATE, null);
  let asOf;
  if (asOfRaw) asOf = normalizeDateCell_(asOfRaw, tz, '[설정] AS_OF_DATE');
  else { const t = new Date(); asOf = new Date(t.getFullYear(), t.getMonth(), t.getDate()); }

  /* 출하계획 — 수요 입력 (별개 파일). 없으면 계획을 세울 수 없다 */
  let ship = [];
  if (!opts.skipShip) {
    const shipSs = openShip_({ create: false });
    if (!shipSs) throw new Error('출하계획 없음 — `X월_출고계획(통합)` 파일을 찾지 못했거나 아직 변환되지 않았습니다 (convertOrderToPlan 먼저 실행)');
    ship = readShipRows_(shipSs);
    ship.forEach(r => { if (!capa[r.품목코드]) warn_('로드', `[출하계획] 품목 ${r.품목코드}는 CAPA에 없어 배정할 수 없습니다`); });
  }

  return { master, settings, planMonth, asOf, capa, machines, machineList, blocked, pref, stock, maxRun, currentItem, shipPrio, targetStock, holidays, ship };
}

/* ────────────────────────────────────────────────────────────────────────────
 *  계획 단독 테스트 — 계획만 생성해 검증 포인트 ①~⑨를 점검하고 결과 파일 [_test] 탭에 덤프
 * ──────────────────────────────────────────────────────────────────────────── */

/** 작업목록 배열 → 검증 결과 {ok, issues[], stats}. 시트에 의존하지 않아 로컬(node)에서도 돌릴 수 있다 */
function validatePlan_(plan, data) {
  const issues = [];
  /* 상태 `완료`는 실적을 그대로 옮긴 기록이라 시간 모델·보빈 배수 검증 대상이 아니다 (규칙 11) */
  const jobs = plan.jobs.filter(j => j.상태 !== '완료');
  const H = data.holidays;
  const hr = (a, b) => (b.getTime() - a.getTime()) / 3600000;

  jobs.forEach(j => {
    const capa = data.capa[j.품목코드];
    /* ① 절대차단 */
    if (data.blocked[j.품목코드 + '|' + j.호기]) issues.push(`① 절대차단 위반: ${j.호기} ${j.품목코드} (순번 ${j.순번})`);
    /* ② Payoff */
    if (capa && capa.가닥수 >= CFG.PLAN.PAYOFF2_MIN_STRANDS && data.machines[j.호기].payoff < 2) issues.push(`② Payoff 위반: ${j.호기}(Payoff ${data.machines[j.호기].payoff})에 ${j.품목코드} 배정`);
    /* ④ 휴무·주말 배정 */
    if (!isWorkingDay_(dayStartOf_(j.시작일시), H)) issues.push(`④ 비근무일 시작: ${j.호기} 순번 ${j.순번} ${fmtDate_(j.시작일시, 'yyyy-MM-dd HH:mm')}`);
    if (!isWorkingDay_(dayStartOf_(new Date(j.종료일시.getTime() - 1)), H)) issues.push(`④ 비근무일 종료: ${j.호기} 순번 ${j.순번} ${fmtDate_(j.종료일시, 'yyyy-MM-dd HH:mm')}`);
    /* ⑥ 보빈 배수 */
    if (capa && Math.abs(j['생산량(kg)'] / capa.bobbinKg - Math.round(j['생산량(kg)'] / capa.bobbinKg)) > 1e-6) issues.push(`⑥ 보빈 배수 아님: ${j.호기} 순번 ${j.순번} ${j.품목코드} ${j['생산량(kg)']}kg (보빈 ${capa.bobbinKg})`);
    /* ⑧ 소요(h) = 생산량 ÷ kg/hr (+교체) */
    if (capa) {
      const expect = j['생산량(kg)'] / capa.kgPerHr + (j.교체 === 'Y' ? CFG.PLAN.CHANGE_HOURS : 0);
      if (Math.abs(expect - j['소요(h)']) > 0.02) issues.push(`⑧ 소요 불일치: ${j.호기} 순번 ${j.순번} 소요 ${j['소요(h)']}h vs 기대 ${expect.toFixed(2)}h`);
      // 점유시간(종료−시작, 비근무 제외)이 준비+교체+가동과 같아야 한다
      const occ = workHoursBetween_(j.시작일시, j.종료일시, H);
      const expOcc = j._readyHr + j._changeHr + j._runHr;
      if (Math.abs(occ - expOcc) > 0.02) issues.push(`⑧ 점유시간 불일치: ${j.호기} 순번 ${j.순번} 근무시간 ${occ.toFixed(2)}h vs 준비+교체+가동 ${expOcc.toFixed(2)}h`);
    }
  });

  /* ③ 최대 가동수 — 품목별 동시 가동 호기 수 스윕 */
  Object.keys(data.maxRun).forEach(item => {
    const evs = [];
    jobs.filter(j => j.품목코드 === item).forEach(j => { evs.push([j.시작일시.getTime(), 1, j.호기]); evs.push([j.종료일시.getTime(), -1, j.호기]); });
    evs.sort((a, b) => a[0] - b[0] || a[1] - b[1]);
    const running = {}; let peak = 0, peakAt = null;
    evs.forEach(e => { if (e[1] > 0) running[e[2]] = (running[e[2]] || 0) + 1; else { running[e[2]]--; if (!running[e[2]]) delete running[e[2]]; } const n = Object.keys(running).length; if (n > peak) { peak = n; peakAt = e[0]; } });
    if (peak > data.maxRun[item]) issues.push(`③ 최대 가동수 초과: ${item} 동시 ${peak}대 > ${data.maxRun[item]} (${fmtDate_(new Date(peakAt), 'MM-dd HH:mm')})`);
  });

  /* ⑦ 같은 호기 시간 겹침 · 순번 = 시작일시 순  /  ⑨ 보충 작업이 납기 작업보다 뒤 */
  data.machineList.forEach(m => {
    const mj = jobs.filter(j => j.호기 === m).sort((a, b) => a.순번 - b.순번);
    for (let i = 1; i < mj.length; i++) {
      if (mj[i].시작일시.getTime() < mj[i - 1].종료일시.getTime() - 1) issues.push(`⑦ 시간 겹침: ${m} 순번 ${mj[i - 1].순번}→${mj[i].순번}`);
      if (mj[i].시작일시.getTime() < mj[i - 1].시작일시.getTime()) issues.push(`⑦ 순번이 시작일시 순이 아님: ${m} 순번 ${mj[i].순번}`);
    }
    const lastDue = mj.filter(j => j._kind === '납기').reduce((t, j) => Math.max(t, j.종료일시.getTime()), 0);
    mj.filter(j => j._kind === '보충').forEach(j => { if (j.시작일시.getTime() < lastDue) issues.push(`⑨ 보충 작업이 납기 작업 앞에 배정: ${m} 순번 ${j.순번}`); });
  });

  /* ⑤ 납기 */
  plan.lateDemands.forEach(d => issues.push(`⑤ 납기위험: ${d.item} 출하일 ${dateKey_(d.dueDate)} ${d.qty}kg → ${d.machine} 종료 ${fmtDate_(d.end, 'MM-dd HH:mm')} (${hr(d.deadline, d.end).toFixed(1)}h 초과)`));
  plan.unassigned.forEach(d => issues.push(`⑤ 미배정: ${d.item} 출하일 ${dateKey_(d.dueDate)} ${d.qty}kg — ${d.status}`));

  /* 통계 */
  const stats = { jobs: jobs.length, kg: 0, changeovers: 0, byMachine: {}, byItem: {}, late: plan.lateDemands.length, unassigned: plan.unassigned.length, replenishJobs: jobs.filter(j => j._kind === '보충').length };
  jobs.forEach(j => {
    stats.kg += j['생산량(kg)'];
    if (j.교체 === 'Y') stats.changeovers++;
    const bm = stats.byMachine[j.호기] = stats.byMachine[j.호기] || { n: 0, kg: 0, hr: 0, end: null };
    bm.n++; bm.kg += j['생산량(kg)']; bm.hr += j._readyHr + j._changeHr + j._runHr; if (!bm.end || j.종료일시 > bm.end) bm.end = j.종료일시;
    stats.byItem[j.품목코드] = (stats.byItem[j.품목코드] || 0) + j['생산량(kg)'];
  });
  return { ok: issues.length === 0, issues, stats };
}

/**
 * 계획 단독 테스트. 검증 ①~⑨ 결과와 호기별·품목별 요약을 로그로 남기고,
 * 작업목록을 결과 파일 [_test] 탭에 덤프한다 (검증이 끝나면 deleteTestTab()으로 삭제).
 */
function testPlanOnly() {
  clearWarnings_();
  const t0 = Date.now();
  const data = loadData_();
  const plan = generateProductionPlan_(data);
  const v = validatePlan_(plan, data);
  const s = v.stats;
  Logger.log(`■ 계획 ${plan.planId} · 구간 ${fmtDate_(plan.planStart, 'yyyy-MM-dd HH:mm')} ~ ${fmtDate_(plan.horizonEnd, 'yyyy-MM-dd HH:mm')} · 수요 ${plan.demands.length}건 → 작업 ${s.jobs}건(보충 ${s.replenishJobs}) · 총 ${Math.round(s.kg).toLocaleString()}kg · 교체 ${s.changeovers}회 · 납기위험 ${s.late} · 미배정 ${s.unassigned} · ${Math.round((Date.now() - t0) / 100) / 10}초`);
  Logger.log(`■ 검증: ${v.ok ? '✓ 위반 없음' : '✗ ' + v.issues.length + '건'}`);
  v.issues.forEach(x => Logger.log('  ' + x));
  Logger.log('■ 호기별: ' + data.machineList.map(m => { const b = s.byMachine[m]; return b ? `${m} ${b.n}건 ${Math.round(b.kg).toLocaleString()}kg ${b.hr.toFixed(0)}h` : `${m} 배정 없음`; }).join(' | '));
  Logger.log('■ 품목별 생산: ' + Object.keys(s.byItem).sort().map(c => `${c} ${Math.round(s.byItem[c]).toLocaleString()}`).join(' · '));
  Logger.log('■ 구간 말 예상 재고: ' + Object.keys(plan.stockEnd).sort().map(c => `${c} ${Math.round(plan.stockEnd[c]).toLocaleString()}`).join(' · '));
  plan.replenish.forEach(r => Logger.log(`■ 적정재고 ${r.item}: 목표 ${r.target.toLocaleString()} / 보충 전 ${Math.round(r.before).toLocaleString()} → 생산 ${r.produced.toLocaleString()} (${r.machine || '-'}) → ${r.status}`));
  const ws = getWarnings_();
  if (ws.length) { Logger.log(`■ 경고 ${ws.length}건`); ws.forEach(w => Logger.log(`  ⚠ [${w.stage}] ${w.message}`)); }

  /* [_test] 덤프 */
  const cols = ['계획ID', '호기', '순번', '품목코드', '규격', '고객사', '출하일', '시작일시', '종료일시', '소요(h)', '생산량(kg)', '교체', '청크', '상태', '구분', '납기'];
  const rows = plan.jobs.map(j => cols.slice(0, 14).map(c => j[c] === undefined || j[c] === null ? '' : j[c]).concat([j._kind, j._late ? '위험' : '']));
  const rs = openResult_();
  const sh = ensureSheet_(rs, '_test');
  sh.clearContents();
  sh.getRange(1, 1, rows.length + 1, cols.length).setValues([cols].concat(rows));
  if (rows.length) {
    sh.getRange(2, 4, rows.length, 1).setNumberFormat('@');
    sh.getRange(2, 7, rows.length, 1).setNumberFormat('yyyy-mm-dd');
    sh.getRange(2, 8, rows.length, 2).setNumberFormat('yyyy-mm-dd hh:mm');
  }
  sh.setFrozenRows(1);
  Logger.log(`■ [_test] 탭 ${rows.length}행 덤프 → ${rs.getUrl()}`);
  return { plan, validation: v };
}

/** [_test] 임시 탭 삭제 (검증이 끝난 뒤) */
function deleteTestTab() {
  const rs = openResult_();
  const sh = rs.getSheetByName('_test');
  if (sh) { rs.deleteSheet(sh); Logger.log('[_test] 탭 삭제'); } else Logger.log('[_test] 탭 없음');
}

/* ════════════════════════════════════════════════════════════════════════════
 *  전체 파이프라인 runAll(reason)
 *  출고계획 변환(신규/수정 파일만) → 기준정보·출하계획 로드 → 계획 → 검증 → 결과 탭 렌더링 → 백업 사본 → [이력] → 해시 저장 → 알림
 *  reason: '수동' | '출고계획 반영: 파일명' | '출하계획 갱신(판독 확정)' | '기준정보 변경' (트리거는 #5)
 * ════════════════════════════════════════════════════════════════════════════ */
function runAll(reason) {
  reason = reason || '수동';
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(CFG.LOCK_WAIT_MS)) {
    Logger.log('[runAll] skip — 다른 실행이 진행 중');
    appendHistory_({ planId: '-', reason, note: 'skip(실행 중)' });
    return null;
  }
  const t0 = Date.now();
  const startedAt = new Date();
  const planId = fmtDate_(startedAt, 'yyMMdd-HHmm');
  const isManual = reason === '수동';
  const isReplan = /추가발주/.test(reason);
  const label = isReplan ? '🔄 [추가발주 재계획]' : (isManual ? '👤 [수동실행]' : '🤖 [자동실행]');
  let stage = '시작';
  clearWarnings_();
  try {
    stage = '출고계획 변환';
    /* 변환 알림도 runAll과 같은 실행 주체로 보이게 라벨을 넘긴다 —
     * 트리거가 감지해 돌린 변환에 '수동실행'이 뜨면 안 된다. */
    const conv = convertOrderToPlan({ label });

    stage = '데이터 로드';
    const data = loadData_();

    stage = '실적 반영';
    const actuals = applyActuals_(data, loadActuals_(data));   // 규칙 11 — 실적 파일이 없으면 applied=false로 건너뜀

    stage = '계획 생성';
    const plan = generateProductionPlan_(data, { planId, actuals });

    stage = '계획 검증';
    const v = validatePlan_(plan, data);
    const hard = v.issues.filter(x => !/^⑤/.test(x));                 // ⑤ 납기위험·미배정은 계획 결과(경고), 나머지는 버그
    if (hard.length) throw new Error(`계획 검증 실패 ${hard.length}건: ${hard.slice(0, 5).join(' / ')}${hard.length > 5 ? ` 외 ${hard.length - 5}건` : ''}`);
    v.issues.filter(x => /^⑤/.test(x)).forEach(x => warn_('계획', x));
    if (!plan.jobs.length) warn_('계획', '작업이 0건입니다 — 수요가 전부 재고로 충당되었거나 출하계획이 비어 있습니다 (규칙 10: 완료로 넘기지 않음)');

    stage = '이전 계획 스냅샷';
    const prevJobs = readJobsTabSafe_();                                 // 결과를 덮어쓰기 전에 직전 [작업목록]을 읽어둔다
    const statusCounts = applyJobStatus_(plan.jobs, prevJobs, plan.asOf); // 완료·확정·신규·변경·계획 판정 (뷰는 이 컬럼만 본다)

    stage = '결과 렌더링';
    const pub = publishPlan_(planId, plan.jobs, { data, plan, reason, prevJobs, updatedAt: startedAt });

    stage = '백업 사본';
    const backupUrl = backupResult_(planId);

    stage = '이력 기록';
    /* 총 kg·교체 횟수는 "이번 계획으로 만들 양"이므로 실적으로 대체된 `완료` 작업은 빼고 센다 */
    const plannedJobs = plan.jobs.filter(j => j.상태 !== '완료');
    const totalKg = plannedJobs.reduce((s, j) => s + j['생산량(kg)'], 0);
    const changeovers = plannedJobs.filter(j => j.교체 === 'Y').length;
    const sec = Math.round((Date.now() - t0) / 100) / 10;
    appendHistory_({ planId, reason, jobs: plannedJobs.length, kg: totalKg, changeovers, late: plan.lateDemands.length, sec, backupUrl });

    stage = '해시 저장';
    saveContentHashes_();
    saveActualStamps_();

    stage = '경고 기록';
    getWarnings_().forEach(w => appendError_(w.stage, w.message, '경고'));

    stage = '알림';
    const resultUrl = openResult_().getUrl();
    const lateItems = plan.lateDemands.map(d => `${d.item}(${dateKey_(d.dueDate)})`);
    const lines = [
      `${plannedJobs.length ? (isReplan ? '추가발주를 반영해 재계획했습니다.' : '생산계획 생성이 완료되었습니다.') : '⚠ 생산계획 작업이 0건입니다 — 출하계획·재고를 확인하세요.'} (${fmtDate_(startedAt, 'yyyy-MM-dd HH:mm')})`,
    ];
    if (!isManual && !isReplan) lines.push(reasonSentence_(reason));
    /* 재계획이면 추가된 발주와 납기 판정을 먼저 보여준다 */
    if (isReplan) {
      const added = plan.jobs.filter(j => j.상태 === '신규');
      if (added.length) {
        const head = added.slice(0, 3).map(j => `${j.품목코드} ${j.출하일 ? fmtDate_(j.출하일, 'MM/dd') : '-'} ${Math.round(j['생산량(kg)']).toLocaleString()}kg(${j.호기})`).join(' · ');
        lines.push(`추가·변경된 발주: ${head}${added.length > 3 ? ` 외 ${added.length - 3}건` : ''}`);
      }
      if (lateItems.length) {
        const worst = plan.lateDemands[0];
        lines.push(`❗ 납기 충족 불가: ${lateItems.join(', ')} — 최속 가능일 ${fmtDate_(worst.end, 'MM/dd HH:mm')}`);
      } else {
        const lastEnd = plannedJobs.reduce((t, j) => (!t || j.종료일시 > t) ? j.종료일시 : t, null);
        lines.push(`✅ 납기 충족 가능${lastEnd ? ` — 마지막 완료예정 ${fmtDate_(lastEnd, 'MM/dd HH:mm')}` : ''}`);
      }
      lines.push(`변경 표시: 신규 ${statusCounts.신규}건 · 변경 ${statusCounts.변경}건 · 확정 ${statusCounts.확정}건 · 완료 ${statusCounts.완료}건`);
    }
    lines.push(`대상월 ${pub.month.label} / 총 ${plannedJobs.length}건 · 총 ${Math.round(totalKg).toLocaleString()}kg / 교체 ${changeovers}회 / 납기위험 ${lateItems.length ? lateItems.join(', ') : '없음'}` +
      (plan.carryOver && plan.carryOver.length ? ` / 이월 ${plan.carryOver.length}건` : '') +
      (conv && conv.processed.length ? ` / 출고계획 반영 ${conv.processed.map(p => p.name).join(', ')}` : ''));
    lines.push('', '📄 생성된 파일:', `${planId}_${CFG.FILE_NAMES.RESULT}: ${backupUrl}`, `최신 계획: ${resultUrl}`);
    notify_(lines.join('\n'), label);

    Logger.log(`[runAll] 완료 ${planId} · ${reason} · 작업 ${plan.jobs.length}건 · ${Math.round(totalKg).toLocaleString()}kg · 교체 ${changeovers} · 납기위험 ${plan.lateDemands.length} · ${sec}초`);
    return { planId, reason, jobs: plan.jobs.length, kg: totalKg, changeovers, late: plan.lateDemands.length, sec, resultUrl, backupUrl, timing: pub.timing };
  } catch (e) {
    const msg = String(e && e.message || e);
    Logger.log(`[runAll] 실패 @${stage}: ${msg}`);
    appendError_(stage, msg, '오류');
    try { notify_(`생산계획 실행에 실패했습니다. (${fmtDate_(new Date(), 'yyyy-MM-dd HH:mm')})\n${stage}\n${msg}`, isReplan ? '🔄 [추가발주 재계획 오류]' : (isManual ? '👤 [수동실행 오류]' : '🤖 [자동실행 오류]')); } catch (e2) { Logger.log('알림 실패: ' + e2.message); }
    throw e;
  } finally {
    lock.releaseLock();
  }
}

/** 트리거 사유 → 알림 둘째 줄 문장 */
function reasonSentence_(reason) {
  if (/^출고계획 반영/.test(reason)) return `출고계획 파일 변경이 감지되어 실행되었습니다. (${reason.replace(/^출고계획 반영:\s*/, '')})`;
  if (/^실적 반영/.test(reason)) return `실적 파일 갱신이 감지되어 실행되었습니다. (${reason.replace(/^실적 반영:\s*/, '')})`;
  if (/^출하계획 갱신/.test(reason)) return '출하계획 파일 변경이 감지되어 실행되었습니다.';
  if (/^기준정보 변경/.test(reason)) return `기준정보 변경이 감지되어 실행되었습니다.${reason.indexOf('(') >= 0 ? ' ' + reason.slice(reason.indexOf('(')) : ''}`;
  return `${reason} 사유로 실행되었습니다.`;
}

/**
 * 내용 해시 저장 — 출하계획 [출하계획] 탭(HASH_SHIP) + 기준정보 11탭(HASH_M_<탭>). 실행이 끝날 때 호출해
 * 같은 변경으로 트리거가 두 번 돌지 않게 한다. 결과 파일은 대상이 아니다. (감지 로직은 #5 checkAndRunOnUpdate)
 */
function saveContentHashes_() {
  const props = PropertiesService.getScriptProperties();
  const master = openMaster_();
  const upd = {};
  Object.keys(SHEET.MASTER).forEach(k => { const tab = SHEET.MASTER[k]; upd[CFG.PROP.HASH_MASTER_PREFIX + tab] = hashValues_(master, tab); });
  const ship = openShip_({ create: false });
  if (ship) upd[CFG.PROP.HASH_SHIP] = hashValues_(ship, SHEET.SHIP);
  props.setProperties(upd);
}

/** 실적 파일 반영이력 저장 — 찾은 생산실적·출하실적의 lastUpdated를 기록해 같은 파일로 두 번 돌지 않게 (규칙 11 반영은 #6) */
function saveActualStamps_() {
  const act = findActualFiles_();
  const stamps = getJsonProp_(CFG.PROP.ACTUAL_PROCESSED, {});
  [act.prod, act.ship].forEach(f => { if (f) stamps[f.id] = f.lastUpdated.toISOString(); });
  setJsonProp_(CFG.PROP.ACTUAL_PROCESSED, stamps);
}

/* ════════════════════════════════════════════════════════════════════════════
 *  갱신 감지·자동 재계획
 *  installTrigger()가 등록하는 트리거 2개(기준정보 onChange · 1분 시간)가 모두 checkAndRunOnUpdate()로 들어온다.
 *  실제 재계획 여부는 아래 검사가 결정한다 (onChange는 서식 변경에도 오므로 해시로 판단):
 *   1   출고계획 파일이 처리이력에 없거나 lastUpdated가 바뀜        → runAll('출고계획 반영: 파일명')  (runAll이 변환을 먼저 함)
 *   1-2 실적 파일(생산실적·출하실적)이 새로 오거나 lastUpdated가 바뀜 → runAll('실적 반영: 파일명')
 *   2   출하계획 [출하계획] 탭 해시 ≠ 저장값 (판독 확정 append·수동 편집) → runAll('출하계획 갱신(판독 확정)')
 *   3   기준정보 11탭 해시 ≠ 저장값                                   → runAll('기준정보 변경 (탭…)')
 *  결과 파일은 해시 대상이 아니므로 자기 자신을 다시 깨우지 않는다. 검사는 runAll과 같은 스크립트 락을 잡고 하므로
 *  실행 중에는 검사가 skip된다(로그만). 락을 놓은 뒤 runAll을 부르고, runAll은 자기 락으로 중복을 막는다.
 * ════════════════════════════════════════════════════════════════════════════ */
function checkAndRunOnUpdate(e) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) { Logger.log('[감지] skip — 실행 중'); return null; }
  let reason = null;
  try {
    const master = openMaster_();

    /* 1 출고계획 파일 */
    const orderHistory = getJsonProp_(CFG.PROP.ORDER_PROCESSED, {});
    const changedOrders = findOrderFiles_().filter(f => orderHistory[f.id] !== f.lastUpdated.toISOString());
    if (changedOrders.length) reason = '출고계획 반영: ' + changedOrders.map(f => f.name).join(', ');

    /* 1-2 실적 파일 */
    if (!reason) {
      const act = findActualFiles_();
      const stamps = getJsonProp_(CFG.PROP.ACTUAL_PROCESSED, {});
      const changedAct = [act.prod, act.ship].filter(f => f && stamps[f.id] !== f.lastUpdated.toISOString());
      if (changedAct.length) reason = '실적 반영: ' + changedAct.map(f => f.name).join(', ');
    }

    /* 2 출하계획 해시 */
    if (!reason) {
      const ship = openShip_({ create: false });
      if (ship && hashValues_(ship, SHEET.SHIP) !== getProp_(CFG.PROP.HASH_SHIP)) reason = '출하계획 갱신(판독 확정)';
    }

    /* 3 기준정보 탭별 해시 */
    if (!reason) {
      const changedTabs = [];
      Object.keys(SHEET.MASTER).forEach(k => {
        const tab = SHEET.MASTER[k];
        if (hashValues_(master, tab) !== getProp_(CFG.PROP.HASH_MASTER_PREFIX + tab)) changedTabs.push(tab);
      });
      if (changedTabs.length) reason = `기준정보 변경 (${changedTabs.slice(0, 3).join(', ')}${changedTabs.length > 3 ? ` 외 ${changedTabs.length - 3}` : ''})`;
    }
  } finally {
    lock.releaseLock();
  }
  if (!reason) { Logger.log('[감지] 변경 없음'); return null; }
  Logger.log(`[감지] ${reason} → runAll`);
  return runAll(reason);
}

/**
 * 자동 재계획 트리거 설치 — 기존 checkAndRunOnUpdate 트리거를 모두 지우고 2개만 남긴다:
 *  ① 기준정보 파일 설치형 onChange(변경 즉시) ② 1분 시간 기반(출고계획·실적·출하계획 파일 변경 감지 + onChange 누락 대비)
 *  push만으로는 트리거가 생기지 않으므로 최초 1회 실행(메뉴 [설정 → 자동 재계획 트리거 설치] 또는 편집기).
 */
function installTrigger() {
  removeTriggers_();
  const master = openMaster_();
  ScriptApp.newTrigger('checkAndRunOnUpdate').forSpreadsheet(master).onChange().create();
  ScriptApp.newTrigger('checkAndRunOnUpdate').timeBased().everyMinutes(1).create();
  const n = ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'checkAndRunOnUpdate').length;
  Logger.log(`[트리거] 설치 완료 — checkAndRunOnUpdate ${n}개 (기준정보 onChange · 1분 시간)`);
  try { SpreadsheetApp.getUi().alert('자동 재계획 트리거', `설치 완료: 기준정보 변경 감지(onChange) + 1분 주기 파일 감지, 총 ${n}개`, SpreadsheetApp.getUi().ButtonSet.OK); } catch (e) { /* 편집기에서 실행 시 UI 없음 */ }
  return n;
}

/** checkAndRunOnUpdate 트리거 전부 제거 (자동 재계획 중단용) */
function removeTriggers_() {
  let n = 0;
  ScriptApp.getProjectTriggers().forEach(t => { if (t.getHandlerFunction() === 'checkAndRunOnUpdate') { ScriptApp.deleteTrigger(t); n++; } });
  if (n) Logger.log(`[트리거] 기존 ${n}개 제거`);
  return n;
}
function uninstallTrigger() { const n = removeTriggers_(); Logger.log(`[트리거] 해제 ${n}개`); return n; }

/** 알림 채널 단독 테스트 — '👤 [수동실행]' 라벨로 1건 발송. 미설정이면 "알림 생략" 로그 */
function testNotifyOnly() {
  let resultUrl = '';
  try { resultUrl = openResult_().getUrl(); } catch (e) { resultUrl = ''; }
  const sent = notify_(`알림 채널 테스트입니다. (${fmtDate_(new Date(), 'yyyy-MM-dd HH:mm')})\n채널: ${CFG.NOTIFY.channel}${resultUrl ? '\n\n📄 생성된 파일:\n최신 계획: ' + resultUrl : ''}`, '👤 [수동실행]');
  Logger.log(sent ? '[알림 테스트] 전송 완료 — 채팅방에서 초록 라벨과 "바로 열기" 링크를 확인하세요' : '[알림 테스트] 알림 생략 — 웹훅 URL(스크립트 속성 CHAT_WEBHOOK_URL)을 등록하세요');
  try { SpreadsheetApp.getUi().alert('알림 테스트', sent ? '전송 완료 — 채팅방을 확인하세요.' : '알림 생략 — 스크립트 속성 CHAT_WEBHOOK_URL이 없습니다.', SpreadsheetApp.getUi().ButtonSet.OK); } catch (e) { /* UI 없음 */ }
  return sent;
}

/**
 * 렌더링 단독 테스트 — 변환·백업·이력 없이 계획 → publishPlan_만 실행해 렌더러별 소요를 로그로.
 * 검증: Σ[일별생산] = Σ[작업목록](품목·호기), 호기 탭 kg 합 = 작업목록 총 kg, 각 탭에 다른 호기 작업 없음
 */
function testPublishOnly() {
  clearWarnings_();
  const t0 = Date.now();
  const data = loadData_();
  const plan = generateProductionPlan_(data);
  const pub = publishPlan_(plan.planId, plan.jobs, { data, plan, reason: '테스트', updatedAt: new Date() });
  const rs = openResult_();
  // 호기 탭 대사
  let tabKg = 0, mixed = 0, seqMismatch = 0;
  CFG.MACHINES.forEach(m => {
    const t = readTable_(rs, m, false);
    // 호기 탭은 3행이 헤더 — readTable_는 첫 비어 있지 않은 행(1행 요약)을 헤더로 잡으므로 직접 읽는다
    const sh = rs.getSheetByName(m);
    const last = sh.getLastRow();
    if (last < 4) return;
    const vals = sh.getRange(4, 1, last - 3, MACHINE_TAB_COLS_.length).getValues();
    const mine = pub.jobRows.filter(j => j.호기 === m).sort((a, b) => a.순번 - b.순번);
    vals.forEach((r, i) => {
      tabKg += toNumber_(r[3]);
      if (!mine[i] || mine[i].순번 !== toNumber_(r[0])) seqMismatch++;
      if (mine[i] && normalizeItemCode_(r[1]) !== mine[i].품목코드) mixed++;
    });
  });
  const jobKg = pub.jobRows.reduce((s, j) => s + j['생산량(kg)'], 0);
  const dailyKg = pub.daily.reduce((s, d) => s + d['생산량(kg)'], 0);
  Logger.log(`■ 검증: 호기 탭 kg 합 ${Math.round(tabKg).toLocaleString()} vs 작업목록 ${Math.round(jobKg).toLocaleString()} ${Math.abs(tabKg - jobKg) < 1 ? '✓' : '✗'} · Σ일별생산 ${Math.round(dailyKg).toLocaleString()} ${Math.abs(dailyKg - jobKg) < 1 ? '✓' : '✗'} · 순번 불일치 ${seqMismatch} · 다른 호기 혼입 ${mixed} · 납기위험(재고 음수) ${pub.negatives.length}`);
  Logger.log(`■ 렌더링 소요: ` + Object.keys(pub.timing).map(k => `${k} ${pub.timing[k]}s`).join(' · ') + ` · 전체 ${Math.round((Date.now() - t0) / 100) / 10}초`);
  Logger.log(`■ 결과 파일: ${rs.getUrl()}`);
  const ws = getWarnings_();
  if (ws.length) { Logger.log(`■ 경고 ${ws.length}건`); ws.forEach(w => Logger.log(`  ⚠ [${w.stage}] ${w.message}`)); }
  return pub;
}

/* ════════════════════════════════════════════════════════════════════════════
 *  실적 반영 + 추가발주 재계획 (규칙 11)
 * ----------------------------------------------------------------------------
 *  기준일 = [설정] AS_OF_DATE (비어 있으면 오늘). 실적이 있는 구간(기준일 당일 오전 실적 포함)은
 *  [생산실적]·[출하실적] 값을 쓰고 계획값을 쓰지 않는다.
 *   - 계획 시작 재고 = 기초재고 + Σ생산실적 − Σ출하실적 (실적 파일 전체 합산)
 *   - 미출하 잔량(계획 − 실적 > 0)은 기준일 이후 첫 근무일 출하로 이월하고 [요약]에 "이월 N건"
 *   - 출하계획에 없는 (품목, 출고일, 고객사) 출고는 이월분 출하로 보고 앞선 잔량에서 차감
 *   - 실적 구간 작업 상태 `완료`, 호기별 현재 걸린 품목은 실적의 마지막 품목으로 이어받음
 *   - 실적 파일이 없거나 비어 있으면 이 규칙을 통째로 건너뛴다 (전 구간 계획 모드)
 *  실적 두 파일은 보빈(LOT) 1개 = 1행이므로 반드시 집계해서 쓴다.
 * ════════════════════════════════════════════════════════════════════════════ */

/**
 * 실적 파일 2개를 찾아 집계한다 (읽기만 — 옮기거나 쓰지 않는다).
 * @returns {{has, prodFile, shipFile, prodRows, shipRows, prodTotal, prodByItemDate, prodByMachineItemDate,
 *            shipTotal, shipByKey, shipByItemDate, lastItemByMachine, warnings}}
 */
function loadActuals_(data) {
  const files = findActualFiles_();
  const a = {
    has: false, prodFile: '', shipFile: '', prodRows: 0, shipRows: 0,
    prodTotal: {}, prodByItemDate: {}, prodByMachineItemDate: {},
    shipTotal: {}, shipByKey: {}, shipByItemDate: {},
    lastItemByMachine: {}, lastDateByMachine: {},
  };

  /* 생산실적 — 품목코드 × 실적일자 (× 작업호기)로 양품수량 합산 */
  if (files.prod) {
    const t = readExternalTab_(files.prod, SHEET.ACTUAL_PROD);
    t.rows.forEach((r, i) => {
      const code = normalizeItemCode_(field_(r, ['품목코드', '품번']));
      const mach = normalizeMachine_(field_(r, ['작업호기', '호기']));
      const qty = toNumber_(field_(r, ['양품수량', '수량']));
      let d = null;
      try { d = normalizeDateCell_(field_(r, ['실적일자', '계획일자']), CFG.TZ, `[생산실적] ${i + 2}행 실적일자`); }
      catch (e) { warn_('실적', e.message); return; }
      if (!code || !d || !qty) return;
      const dk = dateKey_(d);
      a.prodTotal[code] = (a.prodTotal[code] || 0) + qty;
      a.prodByItemDate[code + '|' + dk] = (a.prodByItemDate[code + '|' + dk] || 0) + qty;
      if (mach) {
        const k = mach + '|' + code + '|' + dk;
        a.prodByMachineItemDate[k] = (a.prodByMachineItemDate[k] || 0) + qty;
        if (!a.lastDateByMachine[mach] || d.getTime() >= a.lastDateByMachine[mach].getTime()) {
          a.lastDateByMachine[mach] = d; a.lastItemByMachine[mach] = code;
        }
      }
      a.prodRows++;
    });
    a.prodFile = files.prod.name;
    if (a.prodRows) a.has = true;
    Logger.log(`[실적] 생산실적 ${files.prod.name} — ${a.prodRows}행 집계 · 품목 ${Object.keys(a.prodTotal).length}종 · 총 ${Math.round(Object.keys(a.prodTotal).reduce((s, c) => s + a.prodTotal[c], 0)).toLocaleString()}kg`);
  }

  /* 출하실적 — 품목코드 × 출고일자 × 납품처(=고객사)로 수량(KG) 합산 */
  if (files.ship) {
    const t = readExternalTab_(files.ship, SHEET.ACTUAL_SHIP);
    t.rows.forEach((r, i) => {
      const code = normalizeItemCode_(field_(r, ['품목코드', '품번']));
      const cust = normalizeCustomer_(field_(r, ['납품처', '고객사']));
      const qty = toNumber_(field_(r, ['수량(KG)', '수량(kg)', '수량']));
      let d = null;
      try { d = normalizeDateCell_(field_(r, ['출고일자', '출하일']), CFG.TZ, `[출하실적] ${i + 2}행 출고일자`); }
      catch (e) { warn_('실적', e.message); return; }
      if (!code || !d || !qty) return;
      const dk = dateKey_(d);
      a.shipTotal[code] = (a.shipTotal[code] || 0) + qty;
      a.shipByItemDate[code + '|' + dk] = (a.shipByItemDate[code + '|' + dk] || 0) + qty;
      const k = code + '|' + dk + '|' + (cust || '-');
      a.shipByKey[k] = (a.shipByKey[k] || 0) + qty;
      a.shipRows++;
    });
    a.shipFile = files.ship.name;
    if (a.shipRows) a.has = true;
    Logger.log(`[실적] 출하실적 ${files.ship.name} — ${a.shipRows}행 집계 · 총 ${Math.round(Object.keys(a.shipTotal).reduce((s, c) => s + a.shipTotal[c], 0)).toLocaleString()}kg`);
  }

  if (!a.has) Logger.log('[실적] 실적 파일 없음 — 규칙 11 건너뜀 (전 구간 계획 모드)');
  return a;
}

/**
 * 집계된 실적을 계획 입력으로 변환한다 (규칙 11).
 * @returns {{applied, asOf, planStart, stockStart, carryOver, demandShipRows, machineAvailable, currentItem, warm, completedJobs, raw}}
 */
function applyActuals_(data, a) {
  const res = { applied: false, asOf: data.asOf, planStart: null, stockStart: null, carryOver: [], demandShipRows: null,
    machineAvailable: {}, currentItem: {}, warm: {}, completedJobs: [], raw: a };
  if (!a || !a.has) return res;
  const asOf = data.asOf;
  const asOfKey = dateKey_(asOf);

  /* 1) 계획 시작 재고 = 기초재고 + Σ생산실적 − Σ출하실적 */
  const codes = {};
  Object.keys(data.capa).forEach(c => { codes[c] = true; });
  Object.keys(data.stock).forEach(c => { codes[c] = true; });
  Object.keys(a.prodTotal).forEach(c => { codes[c] = true; });
  Object.keys(a.shipTotal).forEach(c => { codes[c] = true; });
  const stockStart = {};
  Object.keys(codes).forEach(c => { stockStart[c] = (data.stock[c] || 0) + (a.prodTotal[c] || 0) - (a.shipTotal[c] || 0); });
  res.stockStart = stockStart;

  /* 2) 미출하 잔량 → 이월. 기준일 전 출하계획과 출하실적을 (품목, 출하일, 고객사)로 대조 */
  const pending = [];
  const planKeys = {};
  data.ship.forEach(r => { planKeys[r.품목코드 + '|' + dateKey_(r.출하일) + '|' + r.고객사] = true; });
  data.ship.filter(r => r.출하일.getTime() < asOf.getTime()).forEach(r => {
    const shipped = a.shipByKey[r.품목코드 + '|' + dateKey_(r.출하일) + '|' + r.고객사] || 0;
    const remain = r.출하량 - shipped;
    if (remain > 1e-9) pending.push({ code: r.품목코드, cust: r.고객사, date: r.출하일, qty: remain });
  });
  /* 출하계획에 없는 (품목, 출고일, 고객사) 출고 = 이월분 출하 → 앞선 잔량에서 차감 */
  Object.keys(a.shipByKey).forEach(k => {
    if (planKeys[k]) return;
    const p = k.split('|');
    if (p[1] >= asOfKey) return;                       // 기준일 이후 실적은 이월 대상이 아니다
    let extra = a.shipByKey[k];
    pending.filter(x => x.code === p[0]).sort((x, y) => x.date.getTime() - y.date.getTime()).forEach(x => {
      if (extra <= 1e-9 || x.qty <= 1e-9) return;
      const cut = Math.min(x.qty, extra);
      x.qty -= cut; extra -= cut;
    });
  });
  const carryDate = firstWorkingDayAfter_(asOf, data.holidays);
  res.carryOver = pending.filter(x => x.qty > 1e-9).map(x => ({
    품목코드: x.code, 고객사: x.cust, 원출하일: x.date, 출하일: carryDate, 출하량: Math.round(x.qty * 10) / 10,
  }));

  /* 3) 수요 대상 = 기준일 이후 출하계획 + 이월분 */
  res.demandShipRows = data.ship.filter(r => r.출하일.getTime() >= asOf.getTime())
    .concat(res.carryOver.map(x => ({ 품목코드: x.품목코드, 고객사: x.고객사, 출하일: x.출하일, 출하량: x.출하량, _carry: true })));

  /* 4) 호기 상태 — 기준일 당일 실적만큼 시간을 소진한 뒤부터 배정, 현재 걸린 품목은 실적 마지막 품목 */
  const dayStart = new Date(asOf.getFullYear(), asOf.getMonth(), asOf.getDate(), CFG.PLAN.START_HOUR, 0, 0, 0);
  data.machineList.forEach(m => {
    let hr = 0;
    Object.keys(data.capa).forEach(c => {
      const kg = a.prodByMachineItemDate[m + '|' + c + '|' + asOfKey] || 0;
      if (kg) hr += kg / data.capa[c].kgPerHr;
    });
    res.machineAvailable[m] = hr > 0 ? addWorkHours_(dayStart, hr, data.holidays) : alignToWorkTime_(dayStart, data.holidays);
    if (a.lastItemByMachine[m]) { res.currentItem[m] = a.lastItemByMachine[m]; res.warm[m] = true; }
  });
  res.planStart = alignToWorkTime_(dayStart, data.holidays);

  /* 5) 실적 구간 작업 (상태 `완료`) — 호기 × 품목 × 실적일자로 1행. 같은 날 같은 호기는 순서대로 이어 붙인다 */
  const offset = {};
  Object.keys(a.prodByMachineItemDate).sort().forEach(k => {
    const p = k.split('|');
    const m = p[0], c = p[1], dk = p[2];
    if (dk > asOfKey) return;                          // 기준일 이후 실적은 계획 구간이라 제외
    const kg = a.prodByMachineItemDate[k];
    const capa = data.capa[c];
    const runHr = capa ? kg / capa.kgPerHr : 0;
    const base = toDate_(dk, '[생산실적] 실적일자');
    const ok = m + '|' + dk;
    const off = offset[ok] || 0;
    /* 같은 호기·같은 날 실적이 여러 품목이면 1분씩 밀어 순번을 정한다.
     * 실제 가동시간만큼 밀면 날짜를 넘어가 [일별생산]이 실적일자와 어긋나므로 분 단위로만 민다. */
    const s = new Date(base.getFullYear(), base.getMonth(), base.getDate(), CFG.PLAN.START_HOUR, off, 0, 0);
    offset[ok] = off + 1;
    res.completedJobs.push({
      호기: m, 품목코드: c, 고객사: '', 출하일: null,
      시작일시: s, 종료일시: new Date(s.getTime() + Math.round(runHr * 3600000)),
      '소요(h)': Math.round(runHr * 100) / 100, '생산량(kg)': kg,
      교체: 'N', 청크: '', 상태: '완료',
      _kind: '완료', _readyHr: 0, _changeHr: 0, _runHr: runHr, _actualDate: base, _late: false,
    });
  });

  res.applied = true;
  Logger.log(`[실적] 반영 — 기준일 ${asOfKey} · 완료 작업 ${res.completedJobs.length}건 · 이월 ${res.carryOver.length}건 · 계획 시작 재고: ` +
    Object.keys(stockStart).sort().map(c => `${c} ${Math.round(stockStart[c]).toLocaleString()}`).join(' · '));
  if (res.carryOver.length) res.carryOver.forEach(x => warn_('이월', `${x.품목코드} ${x.고객사} ${dateKey_(x.원출하일)} 미출하 ${Math.round(x.출하량).toLocaleString()}kg → ${dateKey_(x.출하일)}로 이월`));
  return res;
}

/**
 * 작업 상태 판정 — 재계획 diff (신규 · 변경 · 확정 · 계획). `완료`는 실적으로 대체된 구간이라 그대로 둔다.
 * 키 = (호기, 품목코드, 출하일, 청크). 뷰는 이 상태 컬럼 하나만 보고 하이라이트를 그린다.
 */
function applyJobStatus_(jobs, prevJobs, asOf) {
  const keyOf = (j) => [j.호기, j.품목코드, j.출하일 ? dateKey_(j.출하일) : '-', j.청크 || ''].join('|');
  /* 같은 키가 여러 건일 수 있으므로(보충 작업 등) 큐로 담아 하나씩 소비한다.
   * 첫 건만 담으면 같은 계획을 다시 매칭해도 두 번째부터 `변경`으로 잘못 판정된다. */
  const prev = {};
  (prevJobs || []).forEach(p => {
    if (p.상태 === '완료') return;      // 실적 기록은 매칭 대상이 아니다. 담아두면 출하일 없는 보충 작업과 키가 겹쳐 큐가 어긋난다
    const k = keyOf(p);
    (prev[k] = prev[k] || []).push(p);
  });
  const hasPrev = !!(prevJobs && prevJobs.length);
  const counts = { 완료: 0, 확정: 0, 신규: 0, 변경: 0, 계획: 0 };
  jobs.forEach(j => {
    if (j.상태 === '완료') { counts.완료++; return; }
    if (asOf && j.시작일시.getTime() < asOf.getTime()) { j.상태 = '확정'; counts.확정++; return; }
    if (!hasPrev) { j.상태 = '계획'; counts.계획++; return; }
    const q = prev[keyOf(j)];
    const p = (q && q.length) ? q.shift() : null;
    if (!p) { j.상태 = '신규'; counts.신규++; return; }
    const changed = Math.abs(p['생산량(kg)'] - j['생산량(kg)']) > 0.5
      || Math.abs(p.시작일시.getTime() - j.시작일시.getTime()) > 60000
      || Math.abs(p.종료일시.getTime() - j.종료일시.getTime()) > 60000;
    j.상태 = changed ? '변경' : '계획';
    counts[j.상태]++;
  });
  Logger.log(`[상태] 완료 ${counts.완료} · 확정 ${counts.확정} · 신규 ${counts.신규} · 변경 ${counts.변경} · 계획 ${counts.계획}` + (hasPrev ? '' : ' (이전 계획 없음 — 전부 계획)'));
  return counts;
}

/**
 * 추가발주 재계획 — 데모 하이라이트. 출하계획을 다시 읽어 기준일 이후만 재계획하고,
 * 이전 계획과 비교해 신규·변경을 표시한 뒤 🔄 알림을 보낸다. (runAll이 실적 반영·diff를 모두 수행)
 */
function replanFromToday() {
  return runAll('추가발주 반영 재계획');
}

/**
 * 데모 예비 헬퍼 — 추가발주 행을 출하계획 [출하계획] 탭에 upsert한다.
 * 기본값: 고객사C가 7000260을 2,940kg (보빈 7개 = 420 × 7) — 9/17 1,680 + 9/22 1,260
 * @param {Array<{품목코드,고객사,출하일,출하량}>} [rows]
 */
function demoAddOrder(rows) {
  const ship = openShip_({ create: true });
  const def = [
    { 품목코드: '7000260', 고객사: '고객사C', 출하일: new Date(2026, 8, 17), 출하량: 1680 },
    { 품목코드: '7000260', 고객사: '고객사C', 출하일: new Date(2026, 8, 22), 출하량: 1260 },
  ];
  const use = (rows && rows.length ? rows : def).map(r => ({
    품목코드: normalizeItemCode_(r.품목코드), 고객사: normalizeCustomer_(r.고객사),
    출하일: r.출하일 instanceof Date ? r.출하일 : toDate_(r.출하일, 'demoAddOrder 출하일'), 출하량: toNumber_(r.출하량),
  }));
  const up = upsertShipRows_(ship, use);
  const total = use.reduce((s, r) => s + r.출하량, 0);
  Logger.log(`[데모] 추가발주 ${use.length}건 ${Math.round(total).toLocaleString()}kg upsert — 추가 ${up.added} · 변경 ${up.updated} · 동일 ${up.unchanged} → [출하계획] ${up.total}행 ${Math.round(up.totalKg).toLocaleString()}kg`);
  use.forEach(r => Logger.log(`  - ${r.품목코드} ${r.고객사} ${dateKey_(r.출하일)} ${Math.round(r.출하량).toLocaleString()}kg`));
  Logger.log('■ 다음: [▶ 생산계획 → 🔄 지금 재계획] 또는 replanFromToday() 실행 (1분 트리거가 설치돼 있으면 자동으로 재계획됨)');
  return { rows: use, upsert: up };
}

/** 실적 집계 단독 테스트 — 실적 파일을 읽어 집계·계획 시작 재고·이월을 로그로 확인 (계획은 만들지 않는다) */
function testActualsOnly() {
  clearWarnings_();
  const data = loadData_();
  const raw = loadActuals_(data);
  if (!raw.has) { Logger.log('■ 실적 파일이 없어 규칙 11을 건너뜁니다 — 탐색 루트 아래에 `생산실적`·`출하실적` 파일을 올리세요'); return raw; }
  const act = applyActuals_(data, raw);
  const exp = { '1600190': 2876, '1900160': 4384, '1900190': 417, '2400190': 2546, '3000173': 3889, '3700260': 4415, '7000260': 8396, '7000320': 1682 };
  Logger.log(`■ 기준일 ${dateKey_(data.asOf)} · 생산실적 ${raw.prodRows}행 · 출하실적 ${raw.shipRows}행`);
  Logger.log('■ 계획 시작 재고 (기초재고 + 생산실적 − 출하실적):');
  Object.keys(act.stockStart).sort().forEach(c => {
    const v = Math.round(act.stockStart[c]);
    const e = exp[c];
    Logger.log(`   ${c}: 기초 ${Math.round(data.stock[c] || 0).toLocaleString()} + 생산 ${Math.round(raw.prodTotal[c] || 0).toLocaleString()} − 출하 ${Math.round(raw.shipTotal[c] || 0).toLocaleString()} = ${v.toLocaleString()}` + (e === undefined ? '' : (v === e ? ' ✓' : ` ✗ (기대 ${e.toLocaleString()})`)));
  });
  Logger.log(`■ 이월 ${act.carryOver.length}건` + (act.carryOver.length ? '' : ' ✓ (샘플 데이터 기대값 0건)'));
  act.carryOver.forEach(x => Logger.log(`   ${x.품목코드} ${x.고객사} ${dateKey_(x.원출하일)} → ${dateKey_(x.출하일)} ${Math.round(x.출하량).toLocaleString()}kg`));
  Logger.log(`■ 완료 작업 ${act.completedJobs.length}건 · 호기별 현재 품목(실적 마지막): ` + data.machineList.map(m => `${m} ${act.currentItem[m] || data.currentItem[m] || '-'}`).join(' · '));
  Logger.log(`■ 수요 대상 출하 행: ${act.demandShipRows.length}건 (기준일 이후 + 이월)`);
  const ws = getWarnings_();
  if (ws.length) { Logger.log(`■ 경고 ${ws.length}건`); ws.forEach(w => Logger.log(`  ⚠ [${w.stage}] ${w.message}`)); }
  return { raw, act };
}

/* ────────────────────────────────────────────────────────────────────────────
 *  출하계획 초기화 (재변환 준비)
 *  [출하계획] 탭을 헤더 4열만 남기고 비우고, 출고계획 처리이력(ORDER_PROCESSED)·동기화 이력(ORDER_SYNC_JSON)·
 *  출하계획 해시(HASH_SHIP)를 지운다. 다음 convertOrderToPlan(testConvertOnly·runAll)이 출고계획을 처음부터
 *  다시 읽어 출하계획을 새로 만들므로, 출고계획에서 지워진 행도 반영된다(upsert만으로는 옛 행이 남기 때문).
 *  파일·링크·결과 파일·백업 사본은 건드리지 않는다(링크 불변 원칙). 출하계획 파일이 아직 없으면 속성만 지운다.
 * ──────────────────────────────────────────────────────────────────────────── */
function resetShipPlan() {
  const ship = openShip_({ create: false });
  let cleared = 0;
  if (ship) {
    cleared = readShipRows_(ship).length;
    writeShipRows_(ship, []);
  }
  /* ORDER_SYNC(원본↔사본 짝)는 **지우지 않는다** — 지우면 동기화가 기존 사본을 못 찾아
   * `출고계획/`에 사본을 새로 만들어 파일이 여러 개로 갈라진다(사용자 지시: 덮어쓰기).
   * 초기화의 목적인 "출고계획을 처음부터 다시 읽기"는 ORDER_PROCESSED 삭제만으로 달성된다. */
  const keys = [CFG.PROP.ORDER_PROCESSED, CFG.PROP.HASH_SHIP];
  const props = PropertiesService.getScriptProperties();
  keys.forEach(k => props.deleteProperty(k));
  Logger.log(`[초기화] 출하계획 ${cleared}행 비움${ship ? '' : ' (출하계획 파일 없음)'} · 속성 삭제: ${keys.join(', ')}`
    + ' (동기화 이력은 유지 — 사본을 새로 만들지 않고 덮어쓰기 위함) → 다음 변환에서 출고계획을 처음부터 다시 읽습니다');
  return { cleared, deletedProps: keys, shipUrl: ship ? ship.getUrl() : '' };
}

/** 초기화 단독 테스트 — resetShipPlan() 실행 후 출하계획 행수 0·속성 삭제를 로그로 확인 */
function testResetShipPlan() {
  clearWarnings_();
  const r = resetShipPlan();
  const ship = openShip_({ create: false });
  const rows = ship ? readShipRows_(ship).length : 0;
  const props = PropertiesService.getScriptProperties();
  const left = r.deletedProps.filter(k => props.getProperty(k) !== null);
  Logger.log(`■ 초기화 결과: 비운 행 ${r.cleared} → 현재 [출하계획] ${rows}행 ${rows === 0 ? '✓' : '✗'} · 속성 ${left.length ? '남음 ✗ ' + left.join(', ') : '전부 삭제 ✓'}${ship ? ' · 파일 링크 유지: ' + ship.getUrl() : ' · 출하계획 파일 없음'}`);
  Logger.log('■ 다음: testConvertOnly(또는 testResetAndConvert)를 실행하면 출고계획을 처음부터 다시 읽어 출하계획을 새로 만듭니다');
  return { reset: r, rowsAfter: rows, propsLeft: left };
}

/** 초기화 + 재변환 테스트 — 비운 뒤 바로 convertOrderToPlan을 돌려 처음 변환 결과가 그대로 재현되는지(전부 "추가") 확인 */
function testResetAndConvert() {
  clearWarnings_();
  const r = resetShipPlan();
  Logger.log(`■ 초기화: 출하계획 ${r.cleared}행 비움 · 처리이력 삭제 → 재변환 시작`);
  const c = testConvertOnly();
  Logger.log(`■ 재변환 대사: 추가 ${c.added} · 변경 ${c.updated} · 동일 ${c.unchanged}${c.updated === 0 && c.unchanged === 0 && c.added > 0 ? ' ✓ (전부 신규 추가 — 초기화 정상)' : ' ✗ (초기화 뒤에는 전부 "추가"여야 함)'}${c.shipRows === r.cleared ? ` · 행수 ${c.shipRows} = 초기화 전 ${r.cleared} ✓` : ` · 행수 ${c.shipRows} ≠ 초기화 전 ${r.cleared} (출고계획이 바뀌었으면 정상)`}`);
  return { reset: r, convert: c };
}

/* ────────────────────────────────────────────────────────────────────────────
 *  변환 단독 테스트
 * ──────────────────────────────────────────────────────────────────────────── */

/** 외부 소스 폴더 동기화만 실행해 결과를 로그로 */
function testSyncOnly() {
  clearWarnings_();
  const settings = loadSettings_(openMaster_());
  const r = syncOrderFiles_(settings);
  Logger.log(JSON.stringify(r));
  getWarnings_().forEach(w => Logger.log(`  ⚠ [${w.stage}] ${w.message}`));
  return r;
}

/**
 * 변환만 실행하고 [출하계획] 탭을 검증한다. 더미데이터 기대값: 104행 · 103,050kg (판독 파트 실제 파일이면 다를 수 있음)
 * 두 번째 실행은 처리이력이 같아 "skip"이어야 한다(멱등). 강제 재처리: testConvertOnly({force:true})
 */
function testConvertOnly(opts) {
  clearWarnings_();
  const r = convertOrderToPlan(opts || {});
  const ship = openShip_({ create: false });
  if (!ship) { Logger.log('■ 출하계획 파일 없음 — 변환이 0건이었거나 출고계획 파일을 찾지 못함'); return r; }
  const rows = readShipRows_(ship);
  const kg = rows.reduce((s, x) => s + x.출하량, 0);
  const byCust = {}, byItem = {};
  let minD = null, maxD = null;
  rows.forEach(x => {
    byCust[x.고객사] = (byCust[x.고객사] || 0) + x.출하량;
    byItem[x.품목코드] = (byItem[x.품목코드] || 0) + x.출하량;
    if (!minD || x.출하일 < minD) minD = x.출하일;
    if (!maxD || x.출하일 > maxD) maxD = x.출하일;
  });
  Logger.log(`■ [출하계획] ${rows.length}행 · ${Math.round(kg).toLocaleString()}kg` + (rows.length === 104 && Math.round(kg) === 103050 ? ' ✓ (샘플 출고계획 기대값 일치)' : ' (샘플 출고계획 기준은 104행·103,050kg — 다른 출고계획 파일이면 기대값이 다름)'));
  Logger.log(`  출하일 범위: ${minD ? dateKey_(minD) : '-'} ~ ${maxD ? dateKey_(maxD) : '-'}`);
  Logger.log(`  고객사별: ${Object.keys(byCust).sort().map(c => `${c} ${Math.round(byCust[c]).toLocaleString()}`).join(' · ')}`);
  Logger.log(`  품목별: ${Object.keys(byItem).sort().map(c => `${c} ${Math.round(byItem[c]).toLocaleString()}`).join(' · ')}`);
  Logger.log(`  출하계획 파일: ${ship.getUrl()}`);
  const ws = getWarnings_();
  if (ws.length) { Logger.log(`■ 경고 ${ws.length}건`); ws.forEach(w => Logger.log(`  ⚠ [${w.stage}] ${w.message}`)); }
  return r;
}
