/**
 * 88_Notify.gs — Google Chat 알림
 *
 * 발주서가 처리돼 월별 파일에 기록되면 채팅방으로 알린다.
 * 확인이 필요한 건(미매핑·검산 실패·보류)은 다른 표시로 보낸다.
 *
 * ⚠️ 웹훅 URL 은 그 자체가 열쇠다 (가진 사람은 누구나 그 채팅방에 글을 쓸 수 있다).
 *
 * 그래서 **스크립트 속성(Script Properties)** 에 둔다. 코드에도 시트에도 남지 않는다.
 * 시트를 `링크가 있는 모든 사용자` 로 공개해도 스크립트 속성은 보이지 않는다
 * (스크립트 편집 권한이 있어야 볼 수 있다).
 *
 * 예전에는 `설정` 탭에 뒀다. 거기 남아 있으면 읽어서 스크립트 속성으로 옮기고 시트에서 지운다.
 *
 * 설정 키
 *   알림.끔   `true` 로 두면 URL 이 있어도 보내지 않는다 (시연 중 조용히 하고 싶을 때)
 */

/** 스크립트 속성의 웹훅 키 */
const _웹훅속성키 = '알림.챗웹훅';

/**
 * 웹훅 URL 을 읽는다. 없거나 꺼져 있으면 빈 문자열.
 *
 * `설정` 탭에 예전 값이 남아 있으면 **스크립트 속성으로 옮기고 시트에서 지운다.**
 * 시트를 공개해도 URL 이 새지 않게 하기 위해서다.
 */
function _챗웹훅URL() {
  if (String(설정전체()['알림.끔'] || '').toLowerCase() === 'true') return '';

  const 속성 = PropertiesService.getScriptProperties();
  const 저장된 = String(속성.getProperty(_웹훅속성키) || '').trim();
  if (저장된) return 저장된;

  // 예전 방식(설정 탭)에서 옮겨온다
  const 시트값 = String(설정전체()[_웹훅속성키] || '').trim();
  if (시트값) {
    속성.setProperty(_웹훅속성키, 시트값);
    try {
      설정값쓰기(_웹훅속성키, '', '스크립트 속성으로 옮겼습니다 (시트를 공개해도 새지 않게)');
      Logger.log('웹훅 URL 을 설정 탭에서 스크립트 속성으로 옮겼습니다');
    } catch (e) {
      Logger.log('설정 탭에서 웹훅을 지우지 못했습니다(무시): ' + e);
    }
    return 시트값;
  }

  return '';
}

/** 웹훅 URL 을 스크립트 속성에 저장한다 (시트에는 쓰지 않는다) */
function _챗웹훅저장(URL) {
  const 속성 = PropertiesService.getScriptProperties();
  if (URL) 속성.setProperty(_웹훅속성키, URL);
  else 속성.deleteProperty(_웹훅속성키);

  // 예전 값이 시트에 남아 있으면 지운다
  try {
    if (String(설정전체()[_웹훅속성키] || '').trim()) {
      설정값쓰기(_웹훅속성키, '', '스크립트 속성에 있습니다 (시트에는 두지 않습니다)');
    }
  } catch (e) {
    Logger.log('설정 탭 정리 실패(무시): ' + e);
  }
}

/**
 * 채팅방에 글을 보낸다.
 *
 * **실패해도 예외를 던지지 않는다.** 알림이 안 갔다고 발주서 처리가 멈추면 안 된다.
 * @return {boolean} 보냈으면 true
 */
function 챗알림(텍스트) {
  const URL = _챗웹훅URL();
  if (!URL) return false;

  try {
    const 응답 = UrlFetchApp.fetch(URL, {
      method: 'post',
      contentType: 'application/json; charset=UTF-8',
      payload: JSON.stringify({ text: String(텍스트 || '') }),
      muteHttpExceptions: true,
    });
    const 코드 = 응답.getResponseCode();
    if (코드 >= 200 && 코드 < 300) return true;

    Logger.log('챗 알림 실패 ' + 코드 + ': ' + 응답.getContentText().slice(0, 300));
    return false;
  } catch (e) {
    Logger.log('챗 알림 예외(무시): ' + e);
    return false;
  }
}

/**
 * 기록이 끝난 뒤 알림을 보낸다.
 *
 * @param {Object} 처리결과 60_Verify.gs 의 처리() 결과
 * @param {Object} 기록요약 70_Record.gs 의 기록() 결과
 * @param {Object} 문맥 { 채널, 원본이름 }
 */
function 발주알림(처리결과, 기록요약, 문맥) {
  if (!_챗웹훅URL()) return false;   // 설정이 없으면 문자열도 만들지 않는다

  문맥 = 문맥 || {};
  const 고객사 = 처리결과.고객사 || '(미상)';
  const 원본 = 문맥.원본이름 || '(이름없음)';
  const 채널 = 문맥.채널 || '';

  const 문제 = (기록요약.미매핑수 || 0) + (기록요약.월미판정수 || 0);
  const 검산실패 = !!(처리결과.검산결과 && 처리결과.검산결과.통과 === false);
  const 확인필요 = 기록요약.보류 || 문제 > 0 || 검산실패;

  const 줄 = [];

  if (확인필요) {
    // 문제가 있는 건은 무엇이 문제인지 붙여 보낸다
    줄.push('⚠️ *[' + 고객사 + '] 발주서 확인 필요*');
    줄.push('원본: ' + 원본 + (채널 ? ' (' + 채널 + ')' : ''));

    if (기록요약.보류) {
      줄.push('*보류* — ' + (기록요약.메모 || []).join(' / '));
    } else {
      const 기록문 = ['기록 ' + 기록요약.기록수 + '행'];
      if (기록요약.미매핑수) 기록문.push('*미매핑 ' + 기록요약.미매핑수 + '행*');
      if (기록요약.월미판정수) 기록문.push('*월 미판정 ' + 기록요약.월미판정수 + '행*');
      if (기록요약.품목명누락수) 기록문.push('품목명 누락 ' + 기록요약.품목명누락수 + '행');
      줄.push(기록문.join(' · '));
    }

    if (처리결과.검산결과 && 처리결과.검산결과.요약) {
      줄.push('검산: ' + 처리결과.검산결과.요약);
    }
  } else {
    const 종류 = _발주종류(기록요약);
    줄.push('🔔 *[' + 고객사 + '] ' + 종류.제목 + '*');
    줄.push(_월표시(기록요약.대상파일) + ' 통합 문서에 ' + 종류.동작);
  }

  // 기록한 월별 파일 링크
  (기록요약.대상파일 || []).forEach(function (t) {
    줄.push('<https://docs.google.com/spreadsheets/d/' + t.파일ID + '/edit#gid=0|' +
      t.파일 + ' 열기>');
  });

  if (확인필요) {
    줄.push('<' + 관리시트().getUrl() + '|관리 시트에서 확인>');
  }

  return 챗알림(줄.join('\n'));
}

/**
 * 이번 건이 첫 발주인지 · 추가 발주인지 · 수정 발주인지 가린다.
 *
 *   대체수 > 0            같은 키가 다시 왔다  → **수정 발주** (이전 행이 취소선으로 바뀐다)
 *   대체수 0 + 기존행있음  새 품목·새 납기다     → **추가 발주**
 *   그 외                 그 달 첫 건          → 그냥 발주서
 */
function _발주종류(기록요약) {
  if (기록요약.대체수) {
    return {
      제목: '수정 발주서 도착!',
      동작: '반영되었습니다. (이전 ' + 기록요약.대체수 + '행은 취소선 처리)',
    };
  }
  if (기록요약.기존행있음) {
    return { 제목: '추가 발주서 도착!', 동작: '추가되었습니다.' };
  }
  return { 제목: '발주서 도착!', 동작: '추가되었습니다.' };
}

/**
 * 기록한 달을 `9월` 처럼 표시한다. 여러 달에 걸치면 `9월·10월`.
 * 대상 파일이 없으면 이번 달로 본다.
 */
function _월표시(대상파일) {
  const 달들 = (대상파일 || []).map(function (t) {
    const m = String(t.월 || '').split('-')[1];
    const n = parseInt(m, 10);
    return n ? n + '월' : '';
  }).filter(function (s) { return s !== ''; });

  // 중복 제거
  const 유일 = [];
  달들.forEach(function (d) { if (유일.indexOf(d) < 0) 유일.push(d); });

  if (유일.length) return 유일.join('·');
  return (new Date().getMonth() + 1) + '월';
}

// ─────────────────────────────────────────────────────────────
// 설정·검증
// ─────────────────────────────────────────────────────────────

/** 웹훅 만드는 법 안내문 (여러 곳에서 쓴다) */
function _웹훅안내문() {
  return [
    '만드는 법:',
    '1. Google Chat 에서 알림받을 스페이스를 연다',
    '2. 스페이스 이름 클릭 → 앱 및 통합 → 웹훅 관리',
    '3. 웹훅 추가 → 이름 지정 → 저장',
    '4. 나오는 URL 을 복사한다',
    '',
    '※ 1:1 대화방에는 웹훅을 만들 수 없습니다. 스페이스여야 합니다.',
  ].join('\n');
}

/**
 * 시트 메뉴용. 웹훅 URL 을 입력받아 `설정` 탭에 넣고 바로 시험 메시지를 보낸다.
 * 사용자가 설정 탭을 직접 편집하지 않아도 되게 한다.
 */
function 알림_웹훅설정_메뉴() {
  const ui = SpreadsheetApp.getUi();
  const 현재 = String(PropertiesService.getScriptProperties()
    .getProperty(_웹훅속성키) || 설정전체()[_웹훅속성키] || '').trim();

  const 안내 = [
    (현재 ? '지금 설정된 URL: ' + 현재.slice(0, 60) + '…' : '아직 설정되지 않았습니다.'),
    '',
    '웹훅 URL 을 붙여넣고 확인을 누르세요.',
    '(비워두고 확인하면 알림이 꺼집니다)',
    '',
    _웹훅안내문(),
  ].join('\n');

  const 답 = ui.prompt('Google Chat 웹훅 설정', 안내, ui.ButtonSet.OK_CANCEL);
  if (답.getSelectedButton() !== ui.Button.OK) return;

  const URL = String(답.getResponseText() || '').trim();

  if (!URL) {
    _챗웹훅저장('');
    ui.alert('알림 설정', '웹훅을 지웠습니다. 알림을 보내지 않습니다.', ui.ButtonSet.OK);
    return;
  }

  if (URL.indexOf('chat.googleapis.com') < 0) {
    const 계속 = ui.alert('확인',
      ['Google Chat 웹훅 주소로 보이지 않습니다.',
       '보통 chat.googleapis.com 으로 시작합니다.',
       '',
       '그래도 저장할까요?',
       '',
       URL.slice(0, 200)].join('\n'),
      ui.ButtonSet.YES_NO);
    if (계속 !== ui.Button.YES) return;
  }

  // 시트가 아니라 스크립트 속성에 저장한다 — 시트를 공개해도 URL 이 새지 않게
  _챗웹훅저장(URL);
  설정값쓰기('알림.끔', '', 'true 로 두면 알림을 보내지 않는다');

  const 보냄 = 챗알림(
    ['🔔 *발주서 자동화 — 알림 설정 완료*',
     '이제 발주서가 처리될 때마다 여기로 알려드립니다.',
     '<' + 관리시트().getUrl() + '|관리 시트 열기>'].join('\n'));

  ui.alert('알림 설정',
    보냄
      ? '설정 탭에 저장했고 시험 메시지를 보냈습니다.\n채팅방을 확인해주세요.'
      : ['설정 탭에는 저장했지만 메시지를 보내지 못했습니다.',
         'URL 이 맞는지, 웹훅이 살아 있는지 확인해주세요.',
         '(실행 로그에 사유가 남습니다)'].join('\n'),
    ui.ButtonSet.OK);
}

/** 시트 메뉴용. 웹훅이 살아 있는지 시험 메시지를 보낸다. */
function 알림_테스트_메뉴() {
  const ui = SpreadsheetApp.getUi();
  const URL = _챗웹훅URL();

  if (!URL) {
    ui.alert('알림 설정 없음',
      ['아직 웹훅이 설정되지 않았습니다.',
       '메뉴의 `알림 설정 (Google Chat 웹훅)` 을 먼저 실행해주세요.',
       '',
       _웹훅안내문()].join('\n'),
      ui.ButtonSet.OK);
    return;
  }

  const 보냄 = 챗알림(
    ['🔔 *발주서 자동화 — 알림 시험*',
     '이 메시지가 보이면 알림 설정이 정상입니다.',
     '<' + 관리시트().getUrl() + '|관리 시트 열기>'].join('\n'));

  ui.alert('알림 시험',
    보냄 ? '보냈습니다. 채팅방을 확인해주세요.'
         : '보내지 못했습니다. 실행 로그에서 사유를 확인하세요 (URL 이 맞는지, 웹훅이 살아 있는지).',
    ui.ButtonSet.OK);
}
