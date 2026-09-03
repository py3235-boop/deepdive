/**
 * 00_Config.gs — 자원 ID 상수, 설정 탭 접근, 문자열 정규화
 *
 * 폴더·시트 ID 는 여기 기본값을 두고, 관리 시트의 `설정` 탭에 같은 키가 있으면
 * 그 값이 이긴다. 대회 시연 중 자원을 갈아끼워야 할 때 코드를 안 고치도록.
 */

/** 코드에 박힌 기본값. 설정 탭이 비어 있으면 이걸 쓴다. */
const 기본설정 = {
  // 폴더
  '폴더.최상위': '1Tbijj1nyLdoFWgs81NwI5i2_Zd6xyGOJ',
  '폴더.첨부': '1FkGsE23INagy_CSv08Rs8yjcDBbgHqTn',
  '폴더.고객사A': '1-M695KoAMZIkwra-L3bgU-GIFYq6gIny',
  '폴더.고객사B': '1Nh93PZBAY6kvz5u7YcJAUy1pT9gb4ub4',
  '폴더.고객사C': '1bNnZg5NBH-0InvqEJgbgsdNy2p2pv1fC',
  '폴더.미분류': '', // 10_Setup.gs 가 만들고 설정 탭에 기록한다
  '폴더.필요파일': '1CR5EMz66dKttfkVdqrtOumQFO8-KuN6u',
  '폴더.출력': '18sf7-eLFpLT68p_yZCKsQA6cJmRER9Q_',

  // 참조 시트
  '시트.변환규칙': '1jbLaR3AcfpDRrQX3zCyrabBAwNIAM3q2Kw53IveaZPI',
  '시트.양식': '1LchPRjXKXx83iQYUfsbZ99zuQwWJLKyIuUAuyg9Le9U',
};

/** 관리 시트 탭 이름 */
const 탭 = {
  설정: '설정',
  고객사식별: '고객사식별',
  처리로그: '처리로그',
  미매핑: '미매핑',
};

/** 취급하는 고객사. 여기 순서가 UI·탭 순서가 된다. */
const 고객사목록 = ['고객사A', '고객사B', '고객사C'];

/** 미분류 폴더 이름 (10_Setup.gs 가 이 이름으로 만든다) */
const 미분류폴더명 = '미분류';

// ─────────────────────────────────────────────────────────────
// 설정 탭 접근
// ─────────────────────────────────────────────────────────────

/** 관리 시트(이 스크립트가 붙어 있는 시트) */
function 관리시트() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

/** 설정 탭을 한 번만 읽어 캐시한다. 실행 단위로 유효. */
let _설정캐시 = null;

function 설정전체() {
  if (_설정캐시) return _설정캐시;

  const map = Object.assign({}, 기본설정);
  const sheet = 관리시트().getSheetByName(탭.설정);
  if (sheet) {
    // 1행은 헤더(키/값/설명), 2행부터 데이터
    const rows = sheet.getDataRange().getValues().slice(1);
    rows.forEach(function (r) {
      const 키 = String(r[0] || '').trim();
      const 값 = String(r[1] || '').trim();
      if (키 && 값) map[키] = 값; // 빈 값은 기본값을 덮지 않는다
    });
  }
  _설정캐시 = map;
  return map;
}

/** 설정값 하나 읽기. 없으면 기본값, 그것도 없으면 예외. */
function 설정값(키) {
  const v = 설정전체()[키];
  if (v === undefined || v === '') {
    throw new Error('설정값이 비어 있습니다: ' + 키 + ' (설정 탭 또는 기본설정에 채워주세요)');
  }
  return v;
}

/** 설정값을 설정 탭에 기록한다. 있으면 갱신, 없으면 추가. */
function 설정값쓰기(키, 값, 설명) {
  const sheet = 관리시트().getSheetByName(탭.설정);
  if (!sheet) throw new Error('설정 탭이 없습니다. 먼저 설치()를 실행하세요.');

  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0] || '').trim() === 키) {
      sheet.getRange(i + 1, 2).setValue(값);
      _설정캐시 = null;
      return;
    }
  }
  sheet.appendRow([키, 값, 설명 || '']);
  _설정캐시 = null;
}

// ─────────────────────────────────────────────────────────────
// 문자열 정규화
// ─────────────────────────────────────────────────────────────

/**
 * 비교용 키 정규화. 공백 전부 제거 + 소문자.
 * 파일명 표기가 `고객사B`(공백없음) / `고객사 A`(공백있음) 로 섞여 있어서
 * 키워드 매칭 전에 반드시 이걸 통과시켜야 한다.
 */
function 키정규화(s) {
  return String(s == null ? '' : s)
    .replace(/\u00A0/g, ' ')  // NBSP
    .replace(/\s+/g, '')
    .toLowerCase();
}

/**
 * 규격 표기 정규화. 소선경 소수부를 3자리로 맞춘다.
 *   `7/0.26`  → `7/0.260`   (발주서 표기 → 변환규칙 표기)
 *   `30/0.173` → `30/0.173`  (이미 3자리면 그대로)
 * 형태가 `숫자/숫자` 가 아니면 공백만 정리해서 그대로 돌려준다.
 */
function 규격정규화(s) {
  const raw = String(s == null ? '' : s).replace(/\s+/g, '');
  const m = raw.match(/^(\d+)\s*\/\s*(\d*\.?\d+)$/);
  if (!m) return raw;
  const 소선수 = parseInt(m[1], 10);
  const 소선경 = parseFloat(m[2]);
  if (isNaN(소선수) || isNaN(소선경)) return raw;
  return 소선수 + '/' + 소선경.toFixed(3);
}

/** 숫자 파싱. `3,780` `12,384원` 같은 표기에서 숫자만 뽑는다. 실패하면 null. */
function 숫자파싱(s) {
  if (typeof s === 'number') return isNaN(s) ? null : s;
  const cleaned = String(s == null ? '' : s).replace(/[^\d.\-]/g, '');
  if (cleaned === '' || cleaned === '-' || cleaned === '.') return null;
  const n = Number(cleaned);
  return isNaN(n) ? null : n;
}

/** 고객사 A 발주코드의 접미사를 떼어 변환규칙의 `코드` 로 만든다. `ZWCSS-00030-P` → `ZWCSS-00030` */
function 발주코드에서코드(s) {
  return String(s == null ? '' : s).trim().replace(/-[A-Za-z]$/, '');
}
