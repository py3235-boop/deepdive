/**
 * 공통 날짜 유틸. year/month는 전부 1-based 월(1~12)을 쓴다.
 */

function dateKey_(date) {
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function daysInMonth_(year, month) {
  return new Date(year, month, 0).getDate(); // month(1-based)의 마지막 날
}

function formatDayLabel_(day) {
  return day + '일';
}

/**
 * 그 달에서 특정 요일(0=일 ... 6=토)에 해당하는 날짜를 전부 반환한다.
 */
function getWeekdayDatesInMonth_(year, month, weekday) {
  const total = daysInMonth_(year, month);
  const dates = [];
  for (let d = 1; d <= total; d++) {
    const date = new Date(year, month - 1, d);
    if (date.getDay() === weekday) dates.push(date);
  }
  return dates;
}

/**
 * 실행이력 — generatePlan() 등을 실행할 때마다 [실행일시, 실행자, 함수, 결과, 메시지]를 한 줄씩
 * "실행이력" 탭(이 스프레드시트 안)에 누적 기록한다. 메시지 칸에는 검증사항이 있으면 건수만이
 * 아니라 "어떤 품목이 왜 경고인지" 실제 내용까지 그대로 적어서, Apps Script 실행 로그(보기 →
 * 로그, 브라우저 닫으면 사라짐)를 안 봐도 시트에서 바로 확인할 수 있게 한다.
 */
function appendExecutionLog_(functionName, result, message) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = '실행이력';
  let sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.appendRow(['실행일시', '실행자', '함수', '결과', '메시지']);
    sheet.getRange(1, 1, 1, 5).setFontWeight('bold');
    sheet.setColumnWidth(1, 140);
    sheet.setColumnWidth(2, 200);
    sheet.setColumnWidth(3, 150);
    sheet.setColumnWidth(4, 60);
    sheet.setColumnWidth(5, 600);
    sheet.setFrozenRows(1);
  }

  let executor;
  try {
    executor = Session.getActiveUser().getEmail() || '(이메일 비공개)';
  } catch (e) {
    executor = '(확인 불가)';
  }

  sheet.appendRow([
    Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss'),
    executor,
    functionName,
    result,
    message || '',
  ]);

  const lastRow = sheet.getLastRow();
  sheet.getRange(lastRow, 5).setWrap(true).setVerticalAlignment('top');
}

/**
 * 구글챗 웹훅 알림 — generatePlan()/applyActualShipment() 실행 결과를 구글챗 스페이스로 보낸다.
 *
 * 사전 준비(한 번만): 구글챗 스페이스 → 앱 및 통합 → 웹훅 추가 → 생성된 URL을 Apps Script 편집기
 * 좌측 톱니바퀴(프로젝트 설정) → 스크립트 속성에 CHAT_WEBHOOK_URL로 등록.
 * CONFIG.CHAT_WEBHOOK_URL이 비어있으면(등록 안 했으면) 조용히 아무것도 안 하고 넘어간다 — 필수
 * 연동이 아니라 있으면 켜지는 부가 알림이라 없어도 나머지 기능에 영향 없음.
 */
function notifyChat_(text) {
  if (!CONFIG.CHAT_WEBHOOK_URL) return;
  try {
    const res = UrlFetchApp.fetch(CONFIG.CHAT_WEBHOOK_URL, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({ text: text }),
      muteHttpExceptions: true, // 실패해도 예외로 안 터지므로, 아래서 응답 코드를 직접 확인해서 로그에 남긴다
    });
    const code = res.getResponseCode();
    if (code < 200 || code >= 300) {
      Logger.log('구글챗 알림 전송 실패(HTTP ' + code + '): ' + res.getContentText());
    } else {
      Logger.log('구글챗 알림 전송 완료');
    }
  } catch (e) {
    // 알림 실패가 본 작업(계획 생성/실적 반영)을 막으면 안 되므로 로그만 남기고 넘어간다.
    Logger.log('구글챗 알림 전송 실패: ' + e.message);
  }
}
