/**
 * 90_웹앱.gs — 웹앱 진입점과 클라이언트가 부르는 서버 함수들
 *
 * 화면에서 하는 일:
 *   1. 현재 상태 보기 (규칙 건수, 산출물 링크, 트리거)
 *   2. 전체 실행 (드라이브 + 메일)
 *   3. 미분류 건 처리 — 사람이 고객사를 지정하면 해당 폴더로 옮기고 재처리
 *   4. 미매핑 목록 확인
 *   5. 수동 업로드 — 파일을 직접 올려 처리
 */

function doGet() {
  return HtmlService.createTemplateFromFile('UI')
    .evaluate()
    .setTitle('발주서 자동화')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/** UI.html 안에서 다른 파일을 끼워넣을 때 쓴다 */
function include(파일명) {
  return HtmlService.createHtmlOutputFromFile(파일명).getContent();
}

// ─────────────────────────────────────────────────────────────
// 상태
// ─────────────────────────────────────────────────────────────

/** 화면 상단에 뿌릴 현재 상태 */
function 웹앱_상태() {
  const 결과 = { 고객사: 고객사목록, 오류: null };

  try {
    const 규칙 = 변환규칙_로드();
    const 양식 = 양식_로드();

    결과.규칙 = 고객사목록.map(function (c) {
      const L = 양식.레이아웃[c];
      return {
        고객사: c,
        규칙수: Object.keys(규칙[c] || {}).length,
        레이아웃: L ? (L.왼쪽.length + '+' + L.오른쪽.length + '열') : '없음',
      };
    });
    결과.단위중량 = 규칙.단위중량;
    결과.경고 = 규칙.경고.concat(양식.경고);

    _식별규칙캐시 = null;
    결과.식별규칙수 = 식별규칙_로드().규칙들.length;

    결과.검색식 = 메일_검색식();
    결과.트리거 = 트리거_상태();

    // 산출물 목록
    결과.산출물 = [];
    const 설정 = 설정전체();
    Object.keys(설정).sort().reverse().forEach(function (키) {
      if (키.indexOf('출력파일.') !== 0) return;
      결과.산출물.push({
        월: 키.replace('출력파일.', ''),
        URL: 'https://docs.google.com/spreadsheets/d/' + 설정[키] + '/edit',
      });
    });

    결과.미분류수 = _미분류파일들().length;
    결과.미매핑수 = Math.max(_탭행수(탭.미매핑) - 1, 0);
    결과.처리로그수 = Math.max(_탭행수(탭.처리로그) - 1, 0);
    결과.관리시트URL = 관리시트().getUrl();
  } catch (e) {
    결과.오류 = String(e && e.message ? e.message : e);
  }

  return 결과;
}

function _탭행수(이름) {
  const s = 관리시트().getSheetByName(이름);
  return s ? s.getLastRow() : 0;
}

// ─────────────────────────────────────────────────────────────
// 전체 실행
// ─────────────────────────────────────────────────────────────

/** 드라이브 + 메일 전 과정 실행 */
function 웹앱_전체실행() {
  try {
    const R = 전체_실행();
    return { 성공: true, 요약: R.요약, 상세: R.상세, 문제: R.문제 };
  } catch (e) {
    return { 성공: false, 요약: '실행 실패: ' + e, 상세: String(e && e.stack || e), 문제: 1 };
  }
}

// ─────────────────────────────────────────────────────────────
// 미분류 처리
// ─────────────────────────────────────────────────────────────

function _미분류파일들() {
  const ID = 설정전체()['폴더.미분류'];
  if (!ID) return [];
  const 목록 = [];
  try {
    const it = DriveApp.getFolderById(ID).getFiles();
    while (it.hasNext()) {
      const f = it.next();
      if (완료표시됨(f.getName())) continue;
      목록.push({ ID: f.getId(), 이름: f.getName(), URL: f.getUrl() });
    }
  } catch (e) {
    // 폴더가 없으면 빈 목록
  }
  return 목록;
}

/** 미분류 폴더에 남은 파일 목록 */
function 웹앱_미분류목록() {
  return _미분류파일들();
}

/**
 * 사람이 고객사를 지정하면 해당 폴더로 옮기고 처리한다.
 * 판별 실패로 미분류에 들어간 건을 사람이 구제하는 경로다.
 */
function 웹앱_미분류지정(파일ID, 고객사) {
  try {
    if (고객사목록.indexOf(고객사) < 0) {
      return { 성공: false, 메시지: '알 수 없는 고객사: ' + 고객사 };
    }

    const 파일 = DriveApp.getFileById(파일ID);
    const 이름 = 파일.getName();

    // 고객사 폴더로 이동
    const 대상 = 고객사폴더(고객사);
    대상.addFile(파일);
    try {
      DriveApp.getFolderById(설정값('폴더.미분류')).removeFile(파일);
    } catch (e) {
      Logger.log('미분류 폴더에서 떼지 못했습니다(무시): ' + e);
    }

    // 처리
    const 규칙 = 변환규칙_로드();
    const 양식 = 양식_로드();
    const 추출결과 = 추출({ blob: 파일.getBlob(), 파일명: 이름, 원본ID: 파일ID });

    if (!추출결과.성공) {
      return { 성공: false, 메시지: '추출 실패: ' + (추출결과.메타.실패사유 || '?') };
    }

    const R = 처리(고객사, 추출결과, 규칙, 양식);
    const 기록요약 = 기록(R, {
      채널: '웹앱(미분류 지정)',
      원본이름: 이름,
      식별키: 파일ID,
      출처URL: 파일.getUrl(),
    });

    if (기록요약.보류) {
      return { 성공: false, 메시지: '보류: ' + 기록요약.메모.join(' / ') };
    }

    완료표시(파일);
    return {
      성공: true,
      메시지: 고객사 + ' 로 지정 → 기록 ' + 기록요약.기록수 + '행' +
        (기록요약.대체수 ? ' (기존 ' + 기록요약.대체수 + '행 대체)' : ''),
    };
  } catch (e) {
    return { 성공: false, 메시지: String(e && e.message ? e.message : e) };
  }
}

// ─────────────────────────────────────────────────────────────
// 미매핑 목록
// ─────────────────────────────────────────────────────────────

/** 미매핑 탭의 최근 항목 */
function 웹앱_미매핑목록(최대) {
  const s = 관리시트().getSheetByName(탭.미매핑);
  if (!s || s.getLastRow() < 2) return [];

  const 개수 = Math.min(최대 || 50, s.getLastRow() - 1);
  const 시작 = s.getLastRow() - 개수 + 1;
  const 격자 = s.getRange(시작, 1, 개수, 6).getValues();

  return 격자.map(function (r) {
    return {
      일시: r[0] instanceof Date ? Utilities.formatDate(r[0], 'Asia/Seoul', 'MM-dd HH:mm') : String(r[0]),
      채널: String(r[1] || ''),
      원본: String(r[2] || ''),
      고객사: String(r[3] || ''),
      사유: String(r[4] || ''),
    };
  }).reverse();
}

// ─────────────────────────────────────────────────────────────
// 수동 업로드
// ─────────────────────────────────────────────────────────────

/**
 * 파일을 직접 올려 처리한다.
 * @param {string} base64 파일 내용 (data URL 의 뒷부분)
 * @param {string} 파일명
 * @param {string} 고객사 비우면 키워드로 자동 판별
 */
function 웹앱_업로드(base64, 파일명, 고객사) {
  try {
    const blob = Utilities.newBlob(Utilities.base64Decode(base64), null, 파일명);
    const 규칙 = 변환규칙_로드();
    const 양식 = 양식_로드();
    const 추출결과 = 추출({ blob: blob, 파일명: 파일명 });

    // 고객사를 안 골랐으면 판별
    let 판정근거 = '사용자 지정';
    if (!고객사) {
      const 판정 = 고객사식별({ 파일명: 파일명, 문서내용: 추출결과.원문 });
      고객사 = 판정.고객사;
      판정근거 = 판정.근거;
    }

    if (!고객사) {
      // 판별 실패 → 미분류 폴더에 보관
      const 저장 = 고객사폴더(null).createFile(blob);
      return {
        성공: false,
        메시지: '고객사를 판별하지 못해 미분류 폴더에 보관했습니다. ' + 판정근거,
        미분류: true, 파일ID: 저장.getId(),
      };
    }

    if (!추출결과.성공) {
      return { 성공: false, 메시지: '추출 실패: ' + (추출결과.메타.실패사유 || '?') };
    }

    // 고객사 폴더에 저장
    const 저장된 = 고객사폴더(고객사).createFile(blob);

    const R = 처리(고객사, 추출결과, 규칙, 양식);
    const 기록요약 = 기록(R, {
      채널: '웹앱(업로드)',
      원본이름: 파일명,
      식별키: 저장된.getId(),
      출처URL: 저장된.getUrl(),
    });

    if (기록요약.보류) {
      return { 성공: false, 메시지: '보류: ' + 기록요약.메모.join(' / ') };
    }

    완료표시(저장된);

    return {
      성공: true,
      메시지: 고객사 + ' → 기록 ' + 기록요약.기록수 + '행' +
        (기록요약.대체수 ? ' (기존 ' + 기록요약.대체수 + '행 대체)' : '') +
        (기록요약.미매핑수 ? ' / 미매핑 ' + 기록요약.미매핑수 + '행' : ''),
      검산: R.검산결과 ? R.검산결과.요약 : '',
      판정근거: 판정근거,
    };
  } catch (e) {
    return { 성공: false, 메시지: String(e && e.message ? e.message : e) };
  }
}

// ─────────────────────────────────────────────────────────────
// 트리거 조작
// ─────────────────────────────────────────────────────────────

function 웹앱_트리거설치(분) {
  try {
    return { 성공: true, 메시지: 트리거_설치(분) };
  } catch (e) {
    return { 성공: false, 메시지: String(e) };
  }
}

function 웹앱_트리거제거() {
  try {
    return { 성공: true, 메시지: 트리거_제거() };
  } catch (e) {
    return { 성공: false, 메시지: String(e) };
  }
}
