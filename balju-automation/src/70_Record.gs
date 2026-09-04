/**
 * 70_Record.gs — 월별 발주서 파일에 고객사별 탭으로 기록한다.
 *
 * 출력 구조 (한 파일에 영구 누적하지 않는다):
 *     발주서/
 *     ├── 2026-09 발주서   탭: 고객사A · 고객사B · 고객사C
 *     └── 2026-10 발주서   10월 건이 처음 들어올 때 자동 생성
 *
 * 각 탭 레이아웃은 `고객사 발주서 양식` 시트를 그대로 따른다.
 * 왼쪽에 고객사 원본, 빈 열 하나, 오른쪽에 변환 결과 — 사람이 좌우를 눈으로 대조한다.
 *
 * **행마다 월을 판정한다.** 고객사A 는 납기 블록이 여러 개라 한 파일에 여러 달이 섞일 수 있다.
 * 섞이면 각 월 파일로 나눠 넣는다.
 */

/** 월별 출력 파일 이름 */
function _월파일이름(월) {
  return 월 + ' 발주서';
}

/** 데이터 시작 행 (1: 병합 라벨, 2: 필드명, 3부터 데이터) */
const _데이터시작행 = 3;

/**
 * 산출물 파일에는 **양식 열 + 출처 링크 한 열**만 둔다.
 * 채널·처리일시·검산 같은 추적 정보는 관리 시트 `처리로그` 에 남으므로 여기 두면 중복이다.
 * 대체됨 표시도 별도 열 없이 **색과 취소선으로만** 한다.
 */

/** 출처 열 이름. 양식 열 뒤에 빈 칸 하나 띄우고 붙는다 (전체폭 + 2 열) */
const _출처헤더 = '출처';

/**
 * 오른쪽 정렬할 필드 — 실제로 더하고 빼는 수치만.
 * `품목코드`(7000260)·`일자`(20260908) 는 숫자로 보이지만 코드라서 가운데 정렬한다.
 */
const _수치필드 = { 수량: true, 중량: true, 조장: true, 금액: true };

/**
 * 필드 하나의 가로 정렬 — **전부 가운데.**
 * 수치 열도 가운데로 맞춘다 (사용자 요청). 천단위 콤마는 `_숫자서식적용` 이 따로 입힌다.
 */
function _정렬(필드) {
  return 'center';
}

/**
 * 좌우 블록과 출처 열에만 배경색을 칠한다. **사이 빈 열은 건너뛴다.**
 *
 * 전체 폭을 한 번에 칠하면 좌우를 가르는 빈 열까지 색이 들어가 구분이 흐려진다.
 * 색을 지울 때는 `색` 에 null 을 넘긴다.
 */
function _블록배경(sheet, 시작행, 행수, 레이아웃, 색) {
  if (행수 <= 0) return;

  sheet.getRange(시작행, 1, 행수, 레이아웃.왼쪽.length).setBackground(색);

  if (레이아웃.오른쪽 && 레이아웃.오른쪽.length) {
    sheet.getRange(시작행, 레이아웃.구분열 + 2, 행수, 레이아웃.오른쪽.length).setBackground(색);
  }

  sheet.getRange(시작행, 레이아웃.전체폭 + 2, 행수, 1).setBackground(색);
}

/** 천단위 콤마 서식 */
const _숫자서식 = '#,##0';

/**
 * 수치 열(수량·중량·조장·금액)에 천단위 콤마를 입힌다.
 * @param 블록열 필드 배열과 그 블록이 시작하는 열 번호(1부터)
 */
function _숫자서식적용(sheet, 시작행, 행수, 필드들, 시작열) {
  if (!필드들 || !행수) return;
  필드들.forEach(function (필드, i) {
    if (!_수치필드[필드]) return;
    sheet.getRange(시작행, 시작열 + i, 행수, 1).setNumberFormat(_숫자서식);
  });
}

/**
 * 열 폭을 내용에 맞춘다. 글자가 잘리지 않게 여유를 두고, 너무 넓어지지 않게 상한을 둔다.
 * autoResizeColumns 만 쓰면 한글이 아슬아슬하게 잘리는 경우가 있어 패딩을 더한다.
 */
function _열폭맞추기(sheet, 시작열, 열수, 최소, 최대) {
  if (열수 <= 0) return;
  sheet.autoResizeColumns(시작열, 열수);
  for (let c = 시작열; c < 시작열 + 열수; c++) {
    const 폭 = sheet.getColumnWidth(c) + 16;   // 여유
    sheet.setColumnWidth(c, Math.max(최소 || 60, Math.min(최대 || 220, 폭)));
  }
}

/**
 * 처리 완료한 파일 이름에 붙이는 꼬리표.
 *
 * 식별키만으로도 중복은 막히지만, 파일명에 표시하면 **드라이브에서 눈으로 구분**되고
 * 메일 첨부로 저장된 파일이 나중에 드라이브 스캔에 다시 잡히는 것도 막힌다
 * (메일 경로와 드라이브 경로의 식별키가 서로 달라서 이 표시가 없으면 재처리된다).
 */
const _완료꼬리 = '_완료';

/** 파일명이 이미 완료 표시인지 */
function 완료표시됨(이름) {
  return String(이름 || '').indexOf(_완료꼬리) >= 0;
}

/**
 * 파일 이름에 `_완료` 를 붙인다. 확장자는 살린다.
 *   `고객사B 9월 발주서.png` → `고객사B 9월 발주서_완료.png`
 * 실패해도 조용히 넘어간다 — 식별키가 이미 중복을 막고 있으므로 치명적이지 않다.
 */
function 완료표시(파일) {
  try {
    const 이름 = 파일.getName();
    if (완료표시됨(이름)) return 이름;

    const 점 = 이름.lastIndexOf('.');
    const 새이름 = 점 > 0
      ? 이름.slice(0, 점) + _완료꼬리 + 이름.slice(점)
      : 이름 + _완료꼬리;

    파일.setName(새이름);
    return 새이름;
  } catch (e) {
    Logger.log('완료 표시 실패(무시): ' + e);
    return null;
  }
}

/**
 * 출처 열에 쓸 짧은 라벨. 파일명·메일제목은 길어서 표를 늘어뜨린다.
 * 클릭하면 원본으로 가니 **무엇이었는지**만 알려주면 된다.
 */
const _형식라벨 = {
  html: '메일',
  pdf: 'PDF',
  이미지: '이미지',
  docx: '문서',
  xlsx: '엑셀',
  csv: 'CSV',
};

function _출처라벨(추출결과, 채널) {
  const 형식 = 추출결과 && 추출결과.형식;
  return _형식라벨[형식] || (채널 || '원본');
}

/**
 * 클릭하면 원본으로 가는 링크 셀을 만든다.
 * 메일이면 Gmail 해당 메일, 첨부·드라이브 파일이면 그 파일로 연결한다.
 */
function _출처링크(URL, 라벨) {
  const 이름 = String(라벨 || '').trim();
  if (!URL) return 이름;
  // HYPERLINK 인자 안의 따옴표는 두 번 써서 이스케이프한다
  const 안전 = (이름 || URL).replace(/"/g, '""');
  return '=HYPERLINK("' + String(URL).replace(/"/g, '""') + '","' + 안전 + '")';
}

/** 색 규칙 — 시트를 열자마자 상태가 눈에 보이게 한다 */
const _색 = {
  대체됨배경: '#f1f3f4',
  대체됨글자: '#9aa0a6',
  경고배경: '#fff8e1',   // 검산 참고용 불일치 — 기록은 되지만 확인이 필요
  누락배경: '#ffe0b2',   // 품목명이 비어 있는 셀
  합계배경: '#f3f3f3',   // 원본 발주서의 합계 행
  추가배경: '#ede7f6',   // 추가 발주로 나중에 들어온 행 (연보라 — 다른 색들과 겹치지 않게)

  // 헤더 — 왼쪽(고객사 원본)과 오른쪽(딥다이브 변환)을 색으로 가른다
  왼쪽라벨: '#a4c2f4',   // 파랑
  왼쪽필드: '#d6e2f7',
  오른쪽라벨: '#b6d7a8', // 초록
  오른쪽필드: '#dfeada',
  출처헤더: '#efefef',

  테두리진: '#666666',
  테두리연: '#b7b7b7',
};

/**
 * 블록별로 테두리를 그린다.
 *
 * 왼쪽·오른쪽·출처를 **따로** 그리고 사이 빈 열은 비워 둔다.
 * 전체를 한 번에 그리면 좌우가 한 덩어리로 보여 구분이 안 된다.
 */
function _테두리(sheet, 시작행, 행수, 레이아웃, 진하게) {
  const 색 = 진하게 ? _색.테두리진 : _색.테두리연;
  const 굵기 = 진하게
    ? SpreadsheetApp.BorderStyle.SOLID_MEDIUM
    : SpreadsheetApp.BorderStyle.SOLID;

  sheet.getRange(시작행, 1, 행수, 레이아웃.왼쪽.length)
    .setBorder(true, true, true, true, true, true, 색, 굵기);

  // 원본 탭처럼 오른쪽 블록이 없는 레이아웃도 있다 (폭 0 은 getRange 가 거부한다)
  if (레이아웃.오른쪽 && 레이아웃.오른쪽.length) {
    sheet.getRange(시작행, 레이아웃.구분열 + 2, 행수, 레이아웃.오른쪽.length)
      .setBorder(true, true, true, true, true, true, 색, 굵기);
  }

  sheet.getRange(시작행, 레이아웃.전체폭 + 2, 행수, 1)
    .setBorder(true, true, true, true, false, true, 색, 굵기);
}

/**
 * 헤더 2행의 서식. 탭이 이미 있어도 다시 적용한다 (예전에 만든 탭도 따라오게).
 */
function _헤더서식(sheet, 레이아웃) {
  const 왼폭 = 레이아웃.왼쪽.length;
  const 오폭 = 레이아웃.오른쪽 ? 레이아웃.오른쪽.length : 0;

  sheet.getRange(1, 1, 1, 왼폭).setBackground(_색.왼쪽라벨);
  sheet.getRange(2, 1, 1, 왼폭).setBackground(_색.왼쪽필드);

  if (오폭) {
    sheet.getRange(1, 왼폭 + 2, 1, 오폭).setBackground(_색.오른쪽라벨);
    sheet.getRange(2, 왼폭 + 2, 1, 오폭).setBackground(_색.오른쪽필드);
  }
  sheet.getRange(2, 레이아웃.전체폭 + 2).setBackground(_색.출처헤더);

  sheet.getRange(1, 1, 2, 레이아웃.전체폭 + 2)
    .setVerticalAlignment('middle');

  _테두리(sheet, 1, 2, 레이아웃, true);
}

/**
 * 발주 한 줄을 식별하는 키의 구성 필드.
 *
 * 같은 키가 다시 오면 **수정 발주**로 보고 이전 행을 `대체됨` 으로 바꾼다.
 * 키가 다르면 **추가 발주**로 보고 이전 행은 그대로 `유효` 로 둔다.
 * 파일 단위로 대체하면 추가 발주일 때 멀쩡한 이전 발주가 무효가 되므로 행 단위로 판단한다.
 *
 * 고객사B 는 일자 필드가 없어 **그 달 안에서 규격 하나당 한 줄**만 유효해진다
 * (월 단위 계획서라서 그게 맞다).
 */
const _행키필드 = {
  고객사A: ['발주코드', '일자'],
  고객사B: ['규격'],
  고객사C: ['제품명', '규격', '일자'],
};

/** 값 하나를 키 조각으로 정규화한다. 규격은 표기 차이를 흡수한다. */
function _키조각(필드, 값) {
  if (필드 === '규격') return 규격정규화(값);
  return 키정규화(값);
}

/**
 * 합계 행 전용 키.
 *
 * 그냥 두면 고객사B 는 키가 `규격` 하나뿐인데 합계 행의 규격이 비어서 키가 빈 문자열이 되고,
 * 대체 판정에서 걸러져 **수정 발주 때마다 합계 행이 쌓인다.** 그래서 따로 만든다.
 * 일자를 붙여 고객사A 의 블록별 합계는 각각 구분된다.
 */
function _합계키(일자) {
  return '__합계__|' + 키정규화(일자);
}

/** 매핑 행에서 식별 키를 만든다. */
function _행키(고객사, 매핑행) {
  if (매핑행.원본 && 매핑행.원본.합계행) {
    return _합계키(_필드값(매핑행.원본, '일자'));
  }
  const 필드들 = _행키필드[고객사] || [];
  return 필드들.map(function (f) {
    return _키조각(f, _필드값(매핑행.원본, f));
  }).join('|');
}

/**
 * 시트에 이미 적힌 행에서 같은 식별 키를 만든다.
 * 왼쪽 블록 어느 칸에든 `합계`/`소계` 라벨이 있으면 합계 행으로 본다.
 */
function _시트행키(고객사, 레이아웃, 행배열) {
  const 왼쪽값들 = 레이아웃.왼쪽.map(function (_, i) { return 행배열[i]; });
  const 합계인가 = 왼쪽값들.some(function (v) {
    return _합계라벨.test(String(v == null ? '' : v).trim());
  });

  if (합계인가) {
    const i일자 = 레이아웃.왼쪽.indexOf('일자');
    return _합계키(i일자 >= 0 ? 행배열[i일자] : '');
  }

  const 필드들 = _행키필드[고객사] || [];
  return 필드들.map(function (f) {
    const i = 레이아웃.왼쪽.indexOf(f);
    return _키조각(f, i >= 0 ? 행배열[i] : '');
  }).join('|');
}

// ─────────────────────────────────────────────────────────────
// 월 판정
// ─────────────────────────────────────────────────────────────

/** 일자 문자열에서 'YYYY-MM'. 못 뽑으면 null. */
function _일자에서월(값) {
  const s = String(값 == null ? '' : 값).trim();
  if (!s) return null;

  let m = s.match(_일자패턴8);
  if (m) return m[1] + '-' + m[2];

  m = s.match(_일자패턴한글);
  if (m) return m[1] + '-' + ('0' + m[2]).slice(-2);

  m = s.match(/^(20\d{2})[-\/.](\d{1,2})/);
  if (m) return m[1] + '-' + ('0' + m[2]).slice(-2);

  return null;
}

/** 매핑 행 하나의 월. 행에 일자가 없으면 null. */
function _행의월(매핑행) {
  const 후보 = [
    매핑행.변환 && 매핑행.변환.일자,
    매핑행.원본 && 매핑행.원본.일자,
  ];
  for (let i = 0; i < 후보.length; i++) {
    const 월 = _일자에서월(후보[i]);
    if (월) return 월;
  }
  return null;
}

/**
 * 행들을 월별로 묶는다.
 * 행에 일자가 있으면 그 월, 없으면 파일 단위 힌트(파일명·제목의 `N월`), 그것도 없으면 미판정.
 */
function _월별로묶기(매핑행들, 기본월) {
  const 묶음 = {};
  const 미판정 = [];

  (매핑행들 || []).forEach(function (행) {
    const 월 = _행의월(행) || 기본월 || null;
    if (!월) { 미판정.push(행); return; }
    if (!묶음[월]) 묶음[월] = [];
    묶음[월].push(행);
  });

  return { 묶음: 묶음, 미판정: 미판정 };
}

// ─────────────────────────────────────────────────────────────
// 월별 파일 / 탭 확보
// ─────────────────────────────────────────────────────────────

/**
 * 해당 월의 발주서 파일을 얻는다. 없으면 만든다.
 * 매번 폴더를 검색하지 않도록 월→파일ID 를 `설정` 탭에 캐시한다.
 */
function 월별파일확보(월) {
  const 설정키 = '출력파일.' + 월;
  const 캐시 = 설정전체()[설정키];

  if (캐시) {
    try {
      // 휴지통에 있는 파일도 ID 로는 열린다. 열리는 것과 살아 있는 것은 다르다
      if (!DriveApp.getFileById(캐시).isTrashed()) {
        return SpreadsheetApp.openById(캐시);
      }
    } catch (e) {
      // 캐시가 죽었으면 (지웠거나 권한 변경) 아래로 내려가 다시 찾는다
    }
  }

  const 폴더 = DriveApp.getFolderById(설정값('폴더.출력'));
  const 이름 = _월파일이름(월);

  // ⚠️ getFilesByName 은 **휴지통에 있는 파일도 돌려준다.**
  // 걸러내지 않으면 산출물_초기화 로 버린 파일을 다시 찾아 거기에 기록하고,
  // 알림 링크가 휴지통 파일을 가리키게 된다 (실제로 그렇게 됐다).
  const it = 폴더.getFilesByName(이름);
  let 파일 = null;
  while (it.hasNext()) {
    const 후보 = it.next();
    if (후보.isTrashed()) continue;
    파일 = 후보;
    break;
  }

  if (!파일) {
    const ss = SpreadsheetApp.create(이름);
    파일 = DriveApp.getFileById(ss.getId());
    // create() 는 내 드라이브 루트에 만든다. 출력 폴더로 옮긴다.
    폴더.addFile(파일);
    try {
      DriveApp.getRootFolder().removeFile(파일);
    } catch (e) {
      // 루트에서 떼는 데 실패해도 출력 폴더에는 들어갔다. 진행한다.
      Logger.log('내 드라이브 루트에서 떼지 못했습니다(무시): ' + e);
    }
  }

  설정값쓰기(설정키, 파일.getId(), 월 + ' 발주서 파일');
  return SpreadsheetApp.openById(파일.getId());
}

/**
 * 파일 안에 고객사 탭을 확보한다. 없으면 양식 레이아웃대로 헤더 2행을 만들어 생성한다.
 *   1행: 병합 라벨  [고객사 X][빈][딥다이브]
 *   2행: 필드명
 */
function 탭확보(ss, 고객사, 레이아웃) {
  let sheet = ss.getSheetByName(고객사);
  if (sheet) {
    _헤더서식(sheet, 레이아웃);   // 예전에 만든 탭에도 서식을 입힌다
    return sheet;
  }

  sheet = ss.insertSheet(고객사);

  const 왼폭 = 레이아웃.왼쪽.length;
  const 오폭 = 레이아웃.오른쪽.length;

  // 1행 — 병합 라벨
  sheet.getRange(1, 1, 1, 왼폭).merge()
    .setValue(레이아웃.왼쪽라벨).setHorizontalAlignment('center').setFontWeight('bold');
  sheet.getRange(1, 왼폭 + 2, 1, 오폭).merge()
    .setValue(레이아웃.오른쪽라벨).setHorizontalAlignment('center').setFontWeight('bold');

  // 2행 — 필드명
  sheet.getRange(2, 1, 1, 왼폭).setValues([레이아웃.왼쪽]);
  sheet.getRange(2, 왼폭 + 2, 1, 오폭).setValues([레이아웃.오른쪽]);
  sheet.getRange(2, 1, 1, 레이아웃.전체폭).setFontWeight('bold').setHorizontalAlignment('center');

  // 출처 링크 열 (양식 열 뒤에 빈 칸 하나 띄우고)
  sheet.getRange(2, 레이아웃.전체폭 + 2)
    .setValue(_출처헤더)
    .setFontWeight('bold').setFontColor('#888888').setHorizontalAlignment('center');
  sheet.setColumnWidth(레이아웃.전체폭 + 1, 24);   // 좌우 구분용 빈 열은 좁게
  sheet.setColumnWidth(레이아웃.전체폭 + 2, 70);   // 출처 열도 좁게

  _헤더서식(sheet, 레이아웃);
  sheet.setFrozenRows(2);

  // 기본 시트가 남아 있으면 지운다
  const 여분 = ss.getSheets().filter(function (s) {
    return /^(Sheet1|시트1)$/.test(s.getName()) && s.getLastRow() === 0;
  });
  여분.forEach(function (s) { try { ss.deleteSheet(s); } catch (e) {} });

  return sheet;
}

// ─────────────────────────────────────────────────────────────
// 기록
// ─────────────────────────────────────────────────────────────

/**
 * 처리 결과를 월별 파일에 기록한다.
 *
 * @param {Object} 처리결과 60_Verify.gs 의 처리() 결과
 * @param {Object} 문맥 { 채널, 원본이름, 식별키, 출처URL }
 *   식별키 — 중복 처리 방지용. 드라이브 파일이면 파일 ID, 메일이면 메시지ID#부분
 *   출처URL — 산출물의 `출처` 열에 넣을 링크. 메일이면 Gmail 주소, 파일이면 파일 주소
 * @return {Object} 기록 요약
 */
function 기록(처리결과, 문맥) {
  문맥 = 문맥 || {};
  const 요약 = {
    고객사: 처리결과.고객사,
    기록수: 0, 미매핑수: 0, 품목명누락수: 0, 월미판정수: 0, 대체수: 0,
    기존행있음: false,   // 이 달 이 고객사 탭에 이미 행이 있었나 (추가/수정 발주 판단용)
    대상파일: [], 보류: false, 메모: [],
  };

  const 채널 = 문맥.채널 || '드라이브';
  const 원본이름 = 문맥.원본이름 || '(이름없음)';
  const 일시 = new Date();

  // ── 중복 처리 방지 ──
  if (문맥.식별키 && 이미처리했나(문맥.식별키)) {
    요약.보류 = true;
    요약.메모.push('이미 처리한 건입니다 (식별키 ' + 문맥.식별키 + ') — 건너뜁니다');
    return 요약;
  }

  // ── 레이아웃 없으면 기록 불가 ──
  if (!처리결과.레이아웃) {
    요약.보류 = true;
    요약.메모.push('양식 레이아웃이 없어 기록할 수 없습니다');
    미매핑_기록(일시, 채널, 원본이름, 처리결과.고객사, '양식 레이아웃 없음', '');
    처리로그_기록(일시, 채널, 원본이름, 처리결과.고객사, '', '보류', 요약.메모.join(' / '), 문맥.식별키);
    return 요약;
  }

  // ── 검산이 치명적으로 실패하면 이 건 전체를 보류한다 ──
  if (처리결과.검산결과 && 처리결과.검산결과.치명적) {
    요약.보류 = true;
    요약.메모.push(처리결과.검산결과.요약);
    처리결과.매핑결과.행들.forEach(function (행) {
      미매핑_기록(일시, 채널, 원본이름, 처리결과.고객사,
        '검산 실패로 보류: ' + 처리결과.검산결과.요약, JSON.stringify(행.원본));
    });
    처리로그_기록(일시, 채널, 원본이름, 처리결과.고객사, '', '보류(검산)',
      검산_설명(처리결과.검산결과), 문맥.식별키);
    try {
      발주알림(처리결과, 요약, { 채널: 채널, 원본이름: 원본이름 });
    } catch (e) {
      Logger.log('알림 실패(무시): ' + e);
    }
    return 요약;
  }

  // ── 매핑 실패 행은 미매핑으로 ──
  const 성공행들 = [];
  처리결과.매핑결과.행들.forEach(function (행) {
    if (!행.성공) {
      요약.미매핑수++;
      미매핑_기록(일시, 채널, 원본이름, 처리결과.고객사, 행.사유, JSON.stringify(행.원본));
      return;
    }
    if (행.품목명누락) 요약.품목명누락수++;
    성공행들.push(행);
  });

  // ── 월별로 나눠 각 월 파일에 기록 ──
  const 나눔 = _월별로묶기(성공행들, 처리결과.월힌트);
  요약.월미판정수 = 나눔.미판정.length;
  나눔.미판정.forEach(function (행) {
    미매핑_기록(일시, 채널, 원본이름, 처리결과.고객사,
      '월을 판정할 수 없습니다 (일자 필드도, 파일명·제목의 N월도 없음)',
      JSON.stringify(행.원본));
  });

  const 검산문구 = 처리결과.검산결과 ? 처리결과.검산결과.요약 : '';

  const 경고표시 = !!(처리결과.검산결과 && 처리결과.검산결과.통과 === false);

  Object.keys(나눔.묶음).sort().forEach(function (월) {
    const 행들 = 나눔.묶음[월];
    const ss = 월별파일확보(월);
    const 레이아웃 = 처리결과.레이아웃;
    const sheet = 탭확보(ss, 처리결과.고객사, 레이아웃);

    // 쓰기 전에 이미 행이 있었나 — 첫 발주 / 추가 발주를 가르는 근거
    if (sheet.getLastRow() >= _데이터시작행) 요약.기존행있음 = true;

    // ── 추가 발주는 기존 행을 대체하지 않는다 ──
    // 같은 품목·같은 납기일이라 키가 겹치지만, 뜻은 "덮어써라" 가 아니라 "더해라" 다.
    // 그 납기일 블록 안에 끼워 넣고 합계를 다시 계산한다.
    if (문맥.발주종류 === '추가') {
      const 결과 = _추가발주기록(sheet, 레이아웃, 행들, {
        문맥: 문맥, 채널: 채널, 처리결과: 처리결과, 경고표시: 경고표시,
        추가발주: true,   // 새로 넣는 행에 연보라 배경을 입힌다
      });
      요약.기록수 += 결과.기록수;
      요약.추가발주 = true;
      요약.대상파일.push({
        월: 월, 파일: _월파일이름(월), 파일ID: ss.getId(),
        탭: 처리결과.고객사, 행수: 결과.기록수, 시작행: 결과.첫행, 대체수: 0,
        합계갱신: 결과.합계갱신,
      });
      _원본과탭정리(ss, 처리결과, 레이아웃, 행들, 채널, 원본이름, 요약);
      return;
    }

    // ── 같은 키의 기존 행을 대체됨으로 바꾼다 (수정 발주) ──
    const 대체수 = _기존행대체(sheet, 처리결과.고객사, 레이아웃, 행들);
    요약.대체수 = (요약.대체수 || 0) + 대체수;

    // 산출물에는 양식 열만 쓴다 (추적 정보는 관리 시트 처리로그에 남는다)
    const 값들 = 행들.map(function (행) {
      return 행을배열로(레이아웃, 행);
    });

    const 시작행 = Math.max(sheet.getLastRow() + 1, _데이터시작행);
    sheet.getRange(시작행, 1, 값들.length, 레이아웃.전체폭).setValues(값들);

    // ── 출처 링크 ──
    // 같은 원본에서 나온 행들이므로 같은 링크를 넣는다.
    // 라벨은 형식만 짧게 (`메일`·`PDF`·`이미지`). 파일명 전체는 표를 늘어뜨린다
    sheet.getRange(시작행, 레이아웃.전체폭 + 2, 값들.length, 1)
      .setValue(_출처링크(문맥.출처URL, _출처라벨(처리결과.추출결과, 채널)));

    // ── 색·정렬 표시 ── (출처 열까지 함께)
    const 범위 = sheet.getRange(시작행, 1, 값들.length, 레이아웃.전체폭 + 2);

    // 새로 쓴 행은 기본 서식으로 되돌린다 (대체됨 회색이 남아 있을 수 있으므로)
    // 배경은 블록 단위로만 — 좌우를 가르는 빈 열에는 색을 넣지 않는다
    _블록배경(sheet, 시작행, 값들.length, 레이아웃, null);
    범위.setFontColor(null).setFontLine('none');

    // 전부 가운데 정렬
    const 한줄정렬 = 레이아웃.왼쪽.map(_정렬)
      .concat(['center'])                    // 좌우 구분 빈 열
      .concat(레이아웃.오른쪽.map(_정렬))
      .concat(['center', 'center']);         // 출처 앞 빈 열, 출처
    범위.setHorizontalAlignments(값들.map(function () { return 한줄정렬; }));

    if (경고표시) {
      // 검산이 참고용으로 불일치한 건 — 기록은 하되 확인이 필요하다는 표시
      _블록배경(sheet, 시작행, 값들.length, 레이아웃, _색.경고배경);
    }

    // 품목명이 비어 있는 셀만 주황으로 (오른쪽 블록의 첫 열이 품목)
    const 품목열 = 레이아웃.구분열 + 2 + Math.max(레이아웃.오른쪽.indexOf('품목'), 0);
    행들.forEach(function (행, i) {
      if (행.품목명누락) sheet.getRange(시작행 + i, 품목열).setBackground(_색.누락배경);

      // 원본 발주서의 합계 행 — 회색 배경 + 굵게. 데이터 행과 구분한다
      // (빈 열은 건너뛴다. 전체 폭을 칠하면 좌우 구분이 흐려진다)
      if (행.합계행 || (행.원본 && 행.원본.합계행)) {
        _블록배경(sheet, 시작행 + i, 1, 레이아웃, _색.합계배경);
        sheet.getRange(시작행 + i, 1, 1, 레이아웃.전체폭).setFontWeight('bold');
      }
    });

    // 새로 쓴 행에 테두리
    _테두리(sheet, 시작행, 값들.length, 레이아웃, false);

    // 중량·수량 같은 수치 열에 천단위 콤마
    _숫자서식적용(sheet, 시작행, 값들.length, 레이아웃.왼쪽, 1);
    _숫자서식적용(sheet, 시작행, 값들.length, 레이아웃.오른쪽, 레이아웃.구분열 + 2);

    // 글자가 잘리지 않게 열 폭을 내용에 맞춘다.
    // 왼쪽·오른쪽만 맞추고 구분 빈 열과 출처 열은 고정 폭을 유지한다
    범위.setWrap(false);
    _열폭맞추기(sheet, 1, 레이아웃.왼쪽.length, 60, 220);
    _열폭맞추기(sheet, 레이아웃.구분열 + 2, 레이아웃.오른쪽.length, 60, 220);
    sheet.setColumnWidth(레이아웃.전체폭 + 1, 24);
    sheet.setColumnWidth(레이아웃.전체폭 + 2, 70);

    _원본과탭정리(ss, 처리결과, 레이아웃, 행들, 채널, 원본이름, 요약);

    요약.기록수 += 행들.length;
    요약.대상파일.push({
      월: 월, 파일: _월파일이름(월), 파일ID: ss.getId(),
      탭: 처리결과.고객사, 행수: 행들.length, 시작행: 시작행, 대체수: 대체수,
    });
  });

  처리로그_기록(일시, 채널, 원본이름, 처리결과.고객사,
    (처리결과.파서결과.경로 || '') + ' 경로',
    '기록 ' + 요약.기록수 + '행' + (요약.미매핑수 ? ' / 미매핑 ' + 요약.미매핑수 : ''),
    검산문구 + (요약.품목명누락수 ? ' / 품목명 누락 ' + 요약.품목명누락수 + '행' : ''),
    문맥.식별키);

  // Google Chat 알림. 여기서 보내면 드라이브·메일첨부·메일본문·수동지정 경로를 다 덮는다.
  // 알림 실패가 발주서 처리를 막지 않도록 예외를 삼킨다
  try {
    발주알림(처리결과, 요약, { 채널: 채널, 원본이름: 원본이름 });
  } catch (e) {
    Logger.log('알림 실패(무시): ' + e);
  }

  return 요약;
}

/** 기록 뒤 공통 마무리 — 원본 탭 기록과 탭 순서 정리 */
function _원본과탭정리(ss, 처리결과, 레이아웃, 행들, 채널, 원본이름, 요약) {
  try {
    원본기록(ss, 처리결과.고객사, 레이아웃, 처리결과.추출결과, 행들,
      { 채널: 채널, 원본이름: 원본이름 });
  } catch (e) {
    Logger.log('원본 탭 기록 실패(무시): ' + e);
    요약.메모.push('원본 탭 기록 실패: ' + e);
  }

  try {
    탭순서정리(ss);
  } catch (e) {
    Logger.log('탭 순서 정리 실패(무시): ' + e);
  }
}

/**
 * 시트에서 같은 납기일의 살아 있는 행들을 찾는다 (취소선 그어진 행은 뺀다).
 *
 * `일자` 필드가 없는 고객사(B)는 전체가 한 블록이다.
 * @return {{데이터행번호:number[], 합계행:number, 수량합:number, 중량합:number}}
 */
function _시트블록찾기(sheet, 레이아웃, 일자) {
  const 빈결과 = { 데이터행번호: [], 합계행: 0, 수량합: 0, 중량합: 0 };
  const 마지막 = sheet.getLastRow();
  if (마지막 < _데이터시작행) return 빈결과;

  const 행수 = 마지막 - _데이터시작행 + 1;
  const 격자 = sheet.getRange(_데이터시작행, 1, 행수, 레이아웃.전체폭).getValues();
  const 취소선 = sheet.getRange(_데이터시작행, 1, 행수, 1).getFontLines();

  const i일자 = 레이아웃.왼쪽.indexOf('일자');
  const i수량 = 레이아웃.왼쪽.indexOf('수량');
  const i중량 = 레이아웃.왼쪽.indexOf('중량');
  const 찾는일자 = 키정규화(일자);

  const 결과 = { 데이터행번호: [], 합계행: 0, 수량합: 0, 중량합: 0 };

  격자.forEach(function (행배열, i) {
    if (취소선[i][0] === 'line-through') return;          // 이미 대체된 행은 없는 셈
    if (i일자 >= 0 && 키정규화(행배열[i일자]) !== 찾는일자) return;

    const 합계인가 = 레이아웃.왼쪽.some(function (_, c) {
      return _합계라벨.test(String(행배열[c] == null ? '' : 행배열[c]).trim());
    });

    const 행번호 = _데이터시작행 + i;
    if (합계인가) { 결과.합계행 = 행번호; return; }

    결과.데이터행번호.push(행번호);
    if (i수량 >= 0) 결과.수량합 += (숫자파싱(행배열[i수량]) || 0);
    if (i중량 >= 0) 결과.중량합 += (숫자파싱(행배열[i중량]) || 0);
  });

  return 결과;
}

/**
 * 추가 발주를 기록한다.
 *
 * 기존 행을 대체하지 않는다. 같은 납기일 블록을 찾아 **그 합계 행 바로 앞에 끼워 넣고**,
 * 합계를 시트의 살아 있는 행 전부로 다시 계산한다.
 *
 * 들어온 발주서의 합계(예: 6 / 2,552)는 **그 문서만의 합계**라 시트에 그대로 쓰면 안 된다.
 * 블록 전체 합계(29 / 12,356)로 갱신하는 것이 맞다.
 *
 * 같은 납기일 블록이 아직 없으면 평소처럼 맨 뒤에 붙인다 (합계 행 포함).
 */
function _추가발주기록(sheet, 레이아웃, 행들, 옵션) {
  const 결과 = { 기록수: 0, 첫행: 0, 합계갱신: [] };
  const i일자 = 레이아웃.왼쪽.indexOf('일자');
  const i수량 = 레이아웃.왼쪽.indexOf('수량');
  const i중량 = 레이아웃.왼쪽.indexOf('중량');

  // 납기일별로 묶는다. 들어온 합계 행은 버린다 (블록 합계로 다시 계산하므로)
  const 일자별 = {};
  const 순서 = [];
  행들.forEach(function (행) {
    if (행.원본 && 행.원본.합계행) return;
    const 일자 = i일자 >= 0 ? String(_필드값(행.원본, '일자') || '').trim() : '';
    if (!일자별[일자]) { 일자별[일자] = []; 순서.push(일자); }
    일자별[일자].push(행);
  });

  순서.forEach(function (일자) {
    const 새행들 = 일자별[일자];
    // 끼워 넣을 때마다 행 번호가 밀리므로 묶음마다 새로 찾는다
    const 블록 = _시트블록찾기(sheet, 레이아웃, 일자);

    let 시작행;
    if (블록.합계행) {
      // 합계 행 바로 앞에 자리를 만든다 → 추가분이 그 납기일 블록 안에 들어간다
      sheet.insertRowsBefore(블록.합계행, 새행들.length);
      시작행 = 블록.합계행;
    } else {
      시작행 = Math.max(sheet.getLastRow() + 1, _데이터시작행);
    }

    _행묶음쓰기(sheet, 시작행, 새행들, 레이아웃, 옵션);
    결과.기록수 += 새행들.length;
    if (!결과.첫행) 결과.첫행 = 시작행;

    if (블록.합계행) {
      // 합계 재계산 — 기존 살아 있는 행 + 이번에 넣은 행
      const 새수량 = 블록.수량합 + 새행들.reduce(function (a, r) {
        return a + (숫자파싱(_필드값(r.원본, '수량')) || 0);
      }, 0);
      const 새중량 = 블록.중량합 + 새행들.reduce(function (a, r) {
        return a + (숫자파싱(_필드값(r.원본, '중량')) || 0);
      }, 0);

      const 합계행번호 = 블록.합계행 + 새행들.length;   // 삽입한 만큼 밀렸다
      _합계행갱신(sheet, 합계행번호, 레이아웃, i수량, i중량, 새수량, 새중량);
      결과.합계갱신.push({ 일자: 일자, 수량: 새수량, 중량: 새중량 });
    } else {
      // 블록이 없었으면 들어온 합계를 그대로 하나 붙인다
      const 들어온합계 = 행들.filter(function (r) {
        return r.원본 && r.원본.합계행 &&
          (i일자 < 0 || String(_필드값(r.원본, '일자') || '').trim() === 일자);
      });
      if (들어온합계.length) {
        const 합계시작 = 시작행 + 새행들.length;
        _행묶음쓰기(sheet, 합계시작, 들어온합계, 레이아웃, 옵션);
        결과.기록수 += 들어온합계.length;
      }
    }
  });

  return 결과;
}

/** 합계 행의 수량·중량 칸을 좌우 모두 갱신한다 */
function _합계행갱신(sheet, 행번호, 레이아웃, i수량, i중량, 수량, 중량) {
  function 쓰기(필드, 값) {
    const 왼 = 레이아웃.왼쪽.indexOf(필드);
    if (왼 >= 0) sheet.getRange(행번호, 왼 + 1).setValue(값);
    const 오 = 레이아웃.오른쪽.indexOf(필드);
    if (오 >= 0) sheet.getRange(행번호, 레이아웃.구분열 + 2 + 오).setValue(값);
  }
  if (i수량 >= 0) 쓰기('수량', 수량);
  if (i중량 >= 0) 쓰기('중량', 중량);

  sheet.getRange(행번호, 1, 1, 레이아웃.전체폭).setFontWeight('bold');
  _블록배경(sheet, 행번호, 1, 레이아웃, _색.합계배경);
  _숫자서식적용(sheet, 행번호, 1, 레이아웃.왼쪽, 1);
  _숫자서식적용(sheet, 행번호, 1, 레이아웃.오른쪽, 레이아웃.구분열 + 2);
}

/**
 * 행 묶음을 시트에 쓰고 서식을 입힌다. 평소 기록과 추가 발주가 함께 쓴다.
 * @param 옵션 { 문맥, 채널, 처리결과, 경고표시 }
 */
function _행묶음쓰기(sheet, 시작행, 행들, 레이아웃, 옵션) {
  if (!행들.length) return;

  const 값들 = 행들.map(function (행) { return 행을배열로(레이아웃, 행); });
  sheet.getRange(시작행, 1, 값들.length, 레이아웃.전체폭).setValues(값들);

  sheet.getRange(시작행, 레이아웃.전체폭 + 2, 값들.length, 1)
    .setValue(_출처링크(옵션.문맥.출처URL,
      _출처라벨(옵션.처리결과.추출결과, 옵션.채널)));

  const 범위 = sheet.getRange(시작행, 1, 값들.length, 레이아웃.전체폭 + 2);
  _블록배경(sheet, 시작행, 값들.length, 레이아웃, null);
  범위.setFontColor(null).setFontLine('none').setFontWeight('normal').setWrap(false);

  const 한줄정렬 = 레이아웃.왼쪽.map(_정렬)
    .concat(['center'])
    .concat(레이아웃.오른쪽.map(_정렬))
    .concat(['center', 'center']);
  범위.setHorizontalAlignments(값들.map(function () { return 한줄정렬; }));

  // 검산 경고가 더 급한 신호라 추가 표시보다 우선한다
  if (옵션.경고표시) {
    _블록배경(sheet, 시작행, 값들.length, 레이아웃, _색.경고배경);
  } else if (옵션.추가발주) {
    // 나중에 추가된 행임을 한눈에 알아보게 (연보라)
    _블록배경(sheet, 시작행, 값들.length, 레이아웃, _색.추가배경);
  }

  const 품목열 = 레이아웃.구분열 + 2 + Math.max(레이아웃.오른쪽.indexOf('품목'), 0);
  행들.forEach(function (행, i) {
    if (행.품목명누락) sheet.getRange(시작행 + i, 품목열).setBackground(_색.누락배경);
    if (행.합계행 || (행.원본 && 행.원본.합계행)) {
      _블록배경(sheet, 시작행 + i, 1, 레이아웃, _색.합계배경);
      sheet.getRange(시작행 + i, 1, 1, 레이아웃.전체폭).setFontWeight('bold');
    }
  });

  _테두리(sheet, 시작행, 값들.length, 레이아웃, false);
  _숫자서식적용(sheet, 시작행, 값들.length, 레이아웃.왼쪽, 1);
  _숫자서식적용(sheet, 시작행, 값들.length, 레이아웃.오른쪽, 레이아웃.구분열 + 2);

  _열폭맞추기(sheet, 1, 레이아웃.왼쪽.length, 60, 220);
  _열폭맞추기(sheet, 레이아웃.구분열 + 2, 레이아웃.오른쪽.length, 60, 220);
  sheet.setColumnWidth(레이아웃.전체폭 + 1, 24);
  sheet.setColumnWidth(레이아웃.전체폭 + 2, 70);
}

/**
 * 새로 들어온 행들과 **같은 식별 키를 가진 기존 행**을 대체됨으로 표시한다.
 * 표시는 회색 배경 + 취소선. 상태 열이 없으므로 **취소선 여부로 이미 대체된 행을 가려낸다.**
 *
 * 키가 다른 행은 건드리지 않는다 — 추가 발주가 무효화되면 안 된다.
 * @return {number} 이번에 새로 대체 표시한 행 수
 */
function _기존행대체(sheet, 고객사, 레이아웃, 새행들) {
  const 마지막 = sheet.getLastRow();
  if (마지막 < _데이터시작행) return 0;

  const 행수 = 마지막 - _데이터시작행 + 1;
  const 폭 = 레이아웃.전체폭 + 2;   // 출처 링크 열까지 함께 칠한다

  const 격자 = sheet.getRange(_데이터시작행, 1, 행수, 레이아웃.전체폭).getValues();
  const 취소선 = sheet.getRange(_데이터시작행, 1, 행수, 1).getFontLines();

  // 이번에 들어온 키 집합
  const 새키들 = {};
  새행들.forEach(function (행) { 새키들[_행키(고객사, 행)] = true; });

  let 대체수 = 0;
  격자.forEach(function (행배열, i) {
    if (취소선[i][0] === 'line-through') return;              // 이미 대체 표시됨
    const 키 = _시트행키(고객사, 레이아웃, 행배열);
    if (!키 || !새키들[키]) return;

    // 배경은 블록 단위로만 (빈 열 제외), 글자 서식은 전체 폭에
    _블록배경(sheet, _데이터시작행 + i, 1, 레이아웃, _색.대체됨배경);
    sheet.getRange(_데이터시작행 + i, 1, 1, 폭)
      .setFontColor(_색.대체됨글자)
      .setFontLine('line-through');
    대체수++;
  });

  return 대체수;
}

// ─────────────────────────────────────────────────────────────
// 원본 탭
// ─────────────────────────────────────────────────────────────

/**
 * 탭 순서를 정해진 순서로 맞춘다.
 *
 *   고객사A · 고객사A 원본 · 고객사B · 고객사B 원본 · 고객사C · 고객사C 원본
 *
 * `insertSheet(이름, 위치)` 의 위치 인자와 `Sheet.getIndex()` 의 기준이 엇갈려
 * "고객사 탭 바로 뒤" 에 꽂는 계산이 어긋났다. 인덱스를 맞추려 애쓰는 대신
 * **기록이 끝난 뒤 원하는 순서로 다시 정렬**한다. 처리 순서와 무관하게 결과가 같아진다.
 */
function 탭순서정리(ss) {
  const 원하는순서 = [];
  고객사목록.forEach(function (고객사) {
    원하는순서.push(고객사);
    원하는순서.push(_원본탭이름(고객사));
  });

  let 위치 = 1;   // moveActiveSheet 는 1 부터 센다
  원하는순서.forEach(function (이름) {
    const sheet = ss.getSheetByName(이름);
    if (!sheet) return;
    try {
      ss.setActiveSheet(sheet);
      ss.moveActiveSheet(위치);
      위치++;
    } catch (e) {
      Logger.log('탭 이동 실패(무시): ' + 이름 + ' — ' + e);
    }
  });

  // 순서에 없는 탭(있다면)은 뒤에 그대로 남는다
}

/** 원본 탭 이름 */
function _원본탭이름(고객사) {
  return 고객사 + ' 원본';
}

/**
 * `고객사X 원본` 탭을 확보한다. 없으면 **고객사 탭 바로 뒤에** 만든다.
 *
 * 양식은 필요한 열만 뽑아 쓰므로 원본의 나머지 정보(업체코드·출고처 등)가 버려진다.
 * 추출이 제대로 됐는지 사람이 대조하려면 원본이 같이 있어야 한다.
 */
function 원본탭확보(ss, 고객사) {
  const 이름 = _원본탭이름(고객사);
  let sheet = ss.getSheetByName(이름);
  if (sheet) return sheet;

  // 위치는 지정하지 않는다. 기록이 끝난 뒤 탭순서정리() 가 제자리로 옮긴다
  sheet = ss.insertSheet(이름);

  // 제목 줄·출처 줄을 두지 않는다. 표만 남긴다 (추적 정보는 관리 시트 처리로그에 있다)
  sheet.setFrozenRows(1);
  return sheet;
}

/**
 * 원본 내용을 `고객사X 원본` 탭에 덧붙인다.
 *
 * 추출에 **표가 있으면 그 표를 그대로** 넣는다 — 고객사A 는 메일 본문 표가 그대로 오므로
 * 양식이 버린 `업체코드`·`업체명`·`코드`·`출고코드`·`출고처` 까지 전부 남는다.
 *
 * 표가 없으면(고객사B 이미지 OCR, 고객사C PDF 는 표 구조를 잃는다) **고객사 탭의 왼쪽 블록과
 * 같은 모양**으로 넣는다. 줄 텍스트를 그대로 붙이면 읽기 어렵기 때문이다.
 */
function 원본기록(ss, 고객사, 레이아웃, 추출결과, 매핑행들, 문맥) {
  const sheet = 원본탭확보(ss, 고객사);

  // 표만 넣는다. 첫 기록은 1행부터, 이후 기록은 한 줄 띄우고 이어 붙인다
  const 이전마지막 = sheet.getLastRow();
  const 시작 = 이전마지막 > 0 ? 이전마지막 + 2 : 1;

  const 표있음 = !!(추출결과 && 추출결과.표들 && 추출결과.표들.length);

  let 행 = 시작;

  // 표와 그 배경색을 짝지어 둔다 (배경색은 메일 HTML 표에서만 온다)
  const 묶음들 = 표있음
    ? 추출결과.표들.map(function (표, i) {
        return { 표: 표, 배경: (추출결과.표배경들 || [])[i] || null };
      })
    : [{ 표: _왼쪽블록격자(레이아웃, 매핑행들), 배경: null }];

  묶음들.forEach(function (묶음, ti) {
    const 표 = 묶음.표;
    if (!표 || !표.length) return;

    if (묶음들.length > 1) {
      sheet.getRange(행, 1).setValue('표 ' + (ti + 1)).setFontColor('#999999');
      행++;
    }

    // setValues 는 직사각형만 받는다. 가장 긴 행에 맞춰 빈 칸을 채운다
    const 폭 = 표.reduce(function (m, r) { return Math.max(m, r.length); }, 0);
    if (!폭) return;
    const 채운격자 = 표.map(function (r) {
      const 복사 = r.slice();
      while (복사.length < 폭) 복사.push('');
      return 복사;
    });

    const 범위 = sheet.getRange(행, 1, 채운격자.length, 폭);
    범위.setValues(채운격자);
    범위.setBorder(true, true, true, true, true, true, _색.테두리연,
      SpreadsheetApp.BorderStyle.SOLID);
    범위.setHorizontalAlignment('center').setWrap(false);

    // 원본 표의 배경색을 그대로 옮긴다. 색이 없는 칸은 null (기본값 유지)
    let 원본색있음 = false;
    if (묶음.배경 && 묶음.배경.length) {
      const 배경격자 = 채운격자.map(function (_, r) {
        const 줄 = (묶음.배경[r] || []).slice();
        while (줄.length < 폭) 줄.push('');
        return 줄.map(function (c) {
          if (c) 원본색있음 = true;
          return c || null;
        });
      });
      범위.setBackgrounds(배경격자);
    }

    // 원본에 색이 없으면 첫 행을 머리글로 칠한다 (고객사 탭 왼쪽 블록과 같은 색)
    const 머리 = sheet.getRange(행, 1, 1, 폭).setFontWeight('bold');
    if (!원본색있음) 머리.setBackground(_색.왼쪽필드);

    // 왼쪽 블록으로 재구성한 경우는 값이 숫자다. 천단위 콤마를 입힌다.
    // (원본 표를 그대로 옮긴 경우는 이미 `3,780` 처럼 문자열이라 손대지 않는다)
    if (!표있음 && 레이아웃) {
      _숫자서식적용(sheet, 행 + 1, 채운격자.length - 1, 레이아웃.왼쪽, 1);
    }

    행 += 채운격자.length;
  });

  if (행 === 시작) {
    sheet.getRange(행, 1).setValue('(원본 내용이 없습니다)').setFontColor('#c5221f');
    행++;
  }

  _열폭맞추기(sheet, 1, Math.min(Math.max(sheet.getLastColumn(), 1), 15), 60, 260);
  return 행 - 시작;
}

/**
 * 양식의 **왼쪽 블록(고객사 원본 필드)** 모양으로 격자를 만든다. 첫 행은 필드명.
 * 고객사 탭의 왼쪽과 같은 값·같은 순서다.
 */
function _왼쪽블록격자(레이아웃, 매핑행들) {
  if (!레이아웃) return [];
  const 격자 = [레이아웃.왼쪽.slice()];
  (매핑행들 || []).forEach(function (행) {
    격자.push(레이아웃.왼쪽.map(function (f) { return _필드값(행.원본, f); }));
  });
  return 격자;
}

// ─────────────────────────────────────────────────────────────
// 관리 시트 로그
// ─────────────────────────────────────────────────────────────

/** 처리로그 탭 헤더 (식별키 열은 중복 방지에 쓴다) */
const _처리로그헤더 = ['일시', '채널', '원본', '고객사', '판정근거', '결과', '비고', '식별키'];

/** 탭 헤더가 기대 컬럼을 다 갖도록 보정한다 (예전에 만든 탭도 따라오게). */
function _헤더보정(sheet, 기대헤더) {
  const 폭 = Math.max(sheet.getLastColumn(), 기대헤더.length);
  const 현재 = sheet.getRange(1, 1, 1, 폭).getValues()[0]
    .map(function (v) { return String(v == null ? '' : v).trim(); });

  let 고칠것 = false;
  기대헤더.forEach(function (이름, i) {
    if (현재[i] !== 이름) 고칠것 = true;
  });
  if (고칠것) {
    sheet.getRange(1, 1, 1, 기대헤더.length).setValues([기대헤더]).setFontWeight('bold');
  }
  return sheet;
}

function 처리로그_기록(일시, 채널, 원본, 고객사, 판정근거, 결과, 비고, 식별키) {
  const sheet = 관리시트().getSheetByName(탭.처리로그);
  if (!sheet) return;
  _헤더보정(sheet, _처리로그헤더);
  sheet.appendRow([일시 || new Date(), 채널 || '', 원본 || '', 고객사 || '',
    판정근거 || '', 결과 || '', 비고 || '', 식별키 || '']);
}

function 미매핑_기록(일시, 채널, 원본, 고객사, 사유, 원본행) {
  const sheet = 관리시트().getSheetByName(탭.미매핑);
  if (!sheet) return;
  sheet.appendRow([일시 || new Date(), 채널 || '', 원본 || '', 고객사 || '',
    사유 || '', 원본행 || '']);
}

/** 이 식별키를 이미 처리했는지 — 처리로그의 식별키 열로 판단한다. */
function 이미처리했나(식별키) {
  if (!식별키) return false;
  const sheet = 관리시트().getSheetByName(탭.처리로그);
  if (!sheet || sheet.getLastRow() < 2) return false;

  _헤더보정(sheet, _처리로그헤더);
  const 열 = _처리로그헤더.indexOf('식별키') + 1;
  const 값들 = sheet.getRange(2, 열, sheet.getLastRow() - 1, 1).getValues();

  for (let i = 0; i < 값들.length; i++) {
    if (String(값들[i][0]).trim() === String(식별키).trim()) return true;
  }
  return false;
}

// ─────────────────────────────────────────────────────────────
// 검증
// ─────────────────────────────────────────────────────────────

// ─────────────────────────────────────────────────────────────
// 드라이브 수집
// ─────────────────────────────────────────────────────────────

/**
 * 고객사 첨부파일 폴더를 훑어 처리하고 월별 파일에 기록한다.
 *
 * `_완료` 가 붙은 파일은 건너뛴다. 기록이 끝난 파일에는 `_완료` 를 붙인다.
 * @return {Object} 구조화된 요약 (전체 실행이 이 값을 쓴다)
 */
function 드라이브수집() {
  const 규칙 = 변환규칙_로드();
  const 양식 = 양식_로드();

  const 요약 = {
    처리파일: 0, 건너뜀: 0, 실패: 0,
    기록행수: 0, 대체수: 0, 미매핑수: 0, 품목명누락수: 0,
    상세: [],
  };

  고객사목록.forEach(function (고객사) {
    const 폴더ID = 설정전체()['폴더.' + 고객사];
    if (!폴더ID) {
      요약.상세.push({ 고객사: 고객사, 상태: '폴더 ID 설정 없음' });
      return;
    }

    const 목록 = 추출_폴더(폴더ID);
    if (!목록.length) {
      요약.상세.push({ 고객사: 고객사, 상태: '폴더가 비어 있습니다' });
      return;
    }

    목록.forEach(function (항목) {
      const 이름 = 항목.파일.getName();
      const 건 = { 고객사: 고객사, 이름: 이름 };

      // 이미 완료 표시된 파일은 건드리지 않는다 (메일 첨부로 저장된 것 포함)
      if (완료표시됨(이름)) {
        건.상태 = '완료 표시된 파일 — 건너뜀';
        요약.건너뜀++;
        요약.상세.push(건);
        return;
      }

      if (!항목.결과.성공) {
        건.상태 = '추출 실패: ' + (항목.결과.메타.실패사유 || '?');
        요약.실패++;
        요약.상세.push(건);
        return;
      }

      const R = 처리(고객사, 항목.결과, 규칙, 양식);
      const 기록요약 = 기록(R, {
        채널: '드라이브',
        원본이름: 이름,
        식별키: 항목.파일.getId(),
        출처URL: 항목.파일.getUrl(),
        발주종류: 발주종류판정([이름]),
      });

      if (기록요약.보류) {
        건.상태 = '보류: ' + 기록요약.메모.join(' / ');
        요약.건너뜀++;
        요약.상세.push(건);
        return;
      }

      건.새이름 = 완료표시(항목.파일);
      건.상태 = '기록 ' + 기록요약.기록수 + '행';
      건.기록요약 = 기록요약;
      건.검산 = R.검산결과 ? R.검산결과.요약 : '';

      요약.처리파일++;
      요약.기록행수 += 기록요약.기록수;
      요약.대체수 += (기록요약.대체수 || 0);
      요약.미매핑수 += 기록요약.미매핑수;
      요약.품목명누락수 += 기록요약.품목명누락수;
      요약.상세.push(건);
    });
  });

  return 요약;
}

/** 드라이브 수집 결과를 사람이 읽을 글로 */
function 드라이브수집_설명(요약) {
  const 줄 = [];
  줄.push('처리 ' + 요약.처리파일 + '파일 / 건너뜀 ' + 요약.건너뜀 +
    (요약.실패 ? ' / 실패 ' + 요약.실패 : ''));
  줄.push('기록 ' + 요약.기록행수 + '행' +
    (요약.대체수 ? ' / 대체됨 ' + 요약.대체수 : '') +
    (요약.미매핑수 ? ' / 미매핑 ' + 요약.미매핑수 : '') +
    (요약.품목명누락수 ? ' / 품목명누락 ' + 요약.품목명누락수 : ''));

  let 현재고객사 = null;
  요약.상세.forEach(function (건) {
    if (건.고객사 !== 현재고객사) {
      현재고객사 = 건.고객사;
      줄.push('');
      줄.push('[' + 건.고객사 + ']');
    }
    if (!건.이름) { 줄.push('   ' + 건.상태); return; }
    줄.push('   ' + 건.이름 + '  →  ' + 건.상태);
    if (건.새이름) 줄.push('      파일명 변경: ' + 건.새이름);
    if (건.검산) 줄.push('      ' + 건.검산);
    (건.기록요약 && 건.기록요약.대상파일 || []).forEach(function (t) {
      줄.push('      → [' + t.파일 + '] 탭 ' + t.탭 + ' ' + t.시작행 + '행부터 ' + t.행수 + '행' +
        (t.대체수 ? ' (기존 ' + t.대체수 + '행 대체됨)' : ''));
    });
  });

  return 줄.join('\n');
}

/** 드라이브 수집을 돌리고 로그에 찍는다 */
function 기록_테스트() {
  const 글 = 드라이브수집_설명(드라이브수집());
  Logger.log(글);
  return 글;
}

/** 시트 메뉴용 래퍼 */
function 기록_테스트_메뉴() {
  const 결과 = 기록_테스트();
  const ui = SpreadsheetApp.getUi();
  ui.alert('드라이브 수집', 결과.slice(0, 4000), ui.ButtonSet.OK);
}

// ─────────────────────────────────────────────────────────────
// 산출물 초기화 (개발·시연 준비용)
// ─────────────────────────────────────────────────────────────

/**
 * 지금까지 만든 산출물을 지우고 처음 상태로 되돌린다.
 *
 * 지우는 것: 월별 발주서 파일(휴지통으로), 설정의 출력파일.* 캐시,
 *            처리로그·미매핑 데이터 행 (헤더는 남긴다)
 * 지우지 않는 것: 변환규칙·양식 시트, 고객사식별 규칙, 첨부 폴더의 원본 파일
 *
 * 처리로그를 비우므로 **같은 파일·메일을 다시 처리할 수 있게 된다.**
 */
function 산출물_초기화() {
  const 로그 = [];

  // 1. 월별 출력 파일을 휴지통으로
  const 설정 = 설정전체();
  Object.keys(설정).forEach(function (키) {
    if (키.indexOf('출력파일.') !== 0) return;
    try {
      DriveApp.getFileById(설정[키]).setTrashed(true);
      로그.push('휴지통으로: ' + 키 + ' (' + 설정[키] + ')');
    } catch (e) {
      로그.push('이미 없음: ' + 키);
    }
  });

  // 2. 설정에서 출력파일 캐시 행 삭제
  const 설정시트 = 관리시트().getSheetByName(탭.설정);
  if (설정시트) {
    const 격자 = 설정시트.getDataRange().getValues();
    for (let r = 격자.length - 1; r >= 1; r--) {
      if (String(격자[r][0] || '').indexOf('출력파일.') === 0) {
        설정시트.deleteRow(r + 1);
      }
    }
    _설정캐시 = null;
  }

  // 3. 처리로그·미매핑 데이터 행 비우기 (헤더 유지)
  [탭.처리로그, 탭.미매핑].forEach(function (이름) {
    const s = 관리시트().getSheetByName(이름);
    if (!s) return;
    const 마지막 = s.getLastRow();
    if (마지막 > 1) {
      s.deleteRows(2, 마지막 - 1);
      로그.push(이름 + ' ' + (마지막 - 1) + '행 삭제');
    } else {
      로그.push(이름 + ' 이미 비어 있음');
    }
  });

  로그.push('');
  로그.push('처리로그를 비웠으므로 같은 파일·메일을 다시 처리할 수 있습니다.');

  const 요약 = 로그.join('\n');
  Logger.log(요약);
  return 요약;
}

/** 시트 메뉴용 래퍼. 되돌릴 수 없으니 한 번 묻는다. */
function 산출물_초기화_메뉴() {
  const ui = SpreadsheetApp.getUi();
  const 답 = ui.alert(
    '산출물 초기화',
    '월별 발주서 파일을 휴지통으로 보내고, 처리로그·미매핑을 비웁니다.\n\n'
    + '변환규칙·양식·고객사식별 규칙과 첨부 원본 파일은 그대로 둡니다.\n\n'
    + '계속할까요?',
    ui.ButtonSet.YES_NO
  );
  if (답 !== ui.Button.YES) return;
  ui.alert('산출물 초기화 결과', 산출물_초기화(), ui.ButtonSet.OK);
}
