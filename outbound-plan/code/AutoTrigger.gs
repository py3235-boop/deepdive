/**
 * 발주서 파일 변경 자동 감지 — 데모/운영용. 발주서는 이 스프레드시트 자체가 아니라 별도 드라이브
 * 파일(ORDER_FOLDER_ID 폴더 안)이라 onEdit/onChange 같은 단순 트리거로는 못 잡는다. 그래서 시간
 * 기반 트리거를 쓰는데, Apps Script의 시간 기반 트리거는 최소 간격이 1분(`everyMinutes`가 받는
 * 값도 1/5/10/15/30분뿐)이라 트리거 자체를 더 촘촘하게 걸 수는 없다. 대신 트리거는 1분마다만
 * 실행되더라도, 그 안에서 짧은 간격(기본 5초)으로 여러 번 확인해서 체감 지연을 줄인다.
 *
 * 실제로 다시 계산할지 말지는 여기서 안 정한다 — 파일이 "바뀐 것 같으면" 일단 generatePlan()을
 * 부르고, 발주 내용이 진짜 그대로면 그 안의 발주지문 비교가 알아서 스킵한다.
 */

const ORDER_WATCH_HANDLER_ = 'checkOrderFileUpdate_';
const ORDER_WATCH_PROP_ = 'ORDER_FILE_LAST_SEEN';
const ORDER_WATCH_POLL_MS_ = 5000; // 트리거 한 번 실행되는 동안 이 간격(5초)으로 반복 확인
const ORDER_WATCH_DURATION_MS_ = 55000; // 다음 트리거(1분 뒤)와 안 겹치게 55초 안에서만 반복

/**
 * 트리거 핸들러 — 사람이 직접 실행할 일은 없고, installOrderWatchTrigger()가 건 트리거가 부른다.
 * 1분에 한 번만 불리지만, 그 안에서 5초 간격으로 최대 55초 동안 계속 확인해서 파일 변경을 훨씬
 * 빨리(최악의 경우도 1분이 아니라 5초 안팎) 잡아낸다.
 */
function checkOrderFileUpdate_() {
  const startedAt = Date.now();
  while (Date.now() - startedAt < ORDER_WATCH_DURATION_MS_) {
    if (_checkOrderFileOnce_()) return; // 변경 감지 + generatePlan() 실행 완료 — 이번 트리거는 끝
    Utilities.sleep(ORDER_WATCH_POLL_MS_);
  }
}

/** 발주서 폴더를 한 번만 확인한다. 변경을 감지해서 generatePlan()까지 실행했으면 true, 아니면 false. */
function _checkOrderFileOnce_() {
  const file = _findLatestOrderFile_();
  if (!file) return false; // 테스트 모드거나 발주서 파일이 아직 없음 — 할 일 없음

  const props = PropertiesService.getScriptProperties();
  const updatedAt = String(file.getLastUpdated().getTime());
  if (props.getProperty(ORDER_WATCH_PROP_) === updatedAt) return false; // 지난 확인 이후 안 바뀜

  props.setProperty(ORDER_WATCH_PROP_, updatedAt);
  generatePlan(null, null, '자동');
  return true;
}

/**
 * 메뉴 실행용 — 발주서 자동 감지 트리거를 설치한다. 이미 있으면 지우고 새로 건다.
 * 트리거 자체는 1분 간격(Apps Script 최소 단위)이지만, 그 안에서 5초마다 반복 확인하므로
 * 실제 감지까지는 대부분 5~10초 안팎이면 된다(최악의 경우에도 1분을 넘지 않음).
 */
function installOrderWatchTrigger() {
  _removeOrderWatchTrigger_();
  ScriptApp.newTrigger(ORDER_WATCH_HANDLER_)
    .timeBased()
    .everyMinutes(1)
    .create();
  SpreadsheetApp.getActiveSpreadsheet().toast(
    '발주서 자동 감지 트리거를 설치했습니다(5초 간격으로 확인, 최악의 경우도 1분 이내). 데모가 끝나면 "자동 감지 중지"로 꺼주세요.',
    '설치 완료',
    8
  );
}

/** 메뉴 실행용 — 자동 감지 트리거를 없앤다(데모 끝나고 정리할 때). */
function removeOrderWatchTrigger() {
  _removeOrderWatchTrigger_();
  SpreadsheetApp.getActiveSpreadsheet().toast('발주서 자동 감지 트리거를 껐습니다.', '완료', 6);
}

function _removeOrderWatchTrigger_() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === ORDER_WATCH_HANDLER_) ScriptApp.deleteTrigger(t);
  });
}
