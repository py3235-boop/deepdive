/**
 * 출고계획 설정값.
 *
 * 스프레드시트 ID, 웹훅 URL 등 민감정보는 코드에 직접 적지 않는다 — 절대 커밋/공유되는 이 파일에
 * 하드코딩하지 말고, Apps Script 편집기 좌측 톱니바퀴(프로젝트 설정) → "스크립트 속성"에 키/값으로
 * 등록해서 PropertiesService로만 읽는다.
 *   등록 예: ORDER_FOLDER_ID = 1AbCdEfG... (발주서 파일이 있는 드라이브 폴더 ID)
 */
const CONFIG = {
  // 'YYYY-MM 발주서' 파일이 있는 구글 드라이브 폴더 ID. 스크립트 속성에 ORDER_FOLDER_ID로 등록.
  // 비어있으면 테스트 모드로 동작 — 이 스프레드시트 자체 안의 "고객사 A/B/C" 탭을 발주서로 간주한다.
  ORDER_FOLDER_ID: PropertiesService.getScriptProperties().getProperty('ORDER_FOLDER_ID') || null,
  // 공휴일 / 기준정보 / 제품 중량 파일 등이 들어있는 구글 드라이브(공유문서함) 폴더 ID.
  // 스크립트 속성에 REFERENCE_FOLDER_ID로 등록. 비어있으면 생산capa 검증/공휴일 시프트를 스킵한다.
  REFERENCE_FOLDER_ID: PropertiesService.getScriptProperties().getProperty('REFERENCE_FOLDER_ID') || null,
  // 구글챗 스페이스에 등록한 수신 웹훅 URL. 스크립트 속성에 CHAT_WEBHOOK_URL로 등록.
  // 비어있으면 알림을 그냥 스킵한다(에러 아님) — 필수 기능이 아니라 있으면 켜지는 부가 기능.
  CHAT_WEBHOOK_URL: PropertiesService.getScriptProperties().getProperty('CHAT_WEBHOOK_URL') || null,
  // resetPlan()이 만드는 빈 "기본 시트" 탭 이름. 실제 계획 데이터는 이 이름이 아니라 매번
  // "N월 출고계획"이라는 별도 탭에 쓰고 항상 맨 앞(1번째)으로 옮긴다(WritePlan.gs 참고).
  PLAN_SHEET_NAME: '기본 시트',
  // 하루 상한은 항상 업체별 독립으로 정확히 트럭 1대 — 물량이 많으면 트럭을 더 싣는 게 아니라
  // 그 업체의 배송일 자체를 늘린다.
  TRUCK_KG: 5000,
  // 생산capa 등의 이유로 어떤 날짜의 배정량이 이 값보다 작으면, 배정이 다 끝난 뒤 그 품목의 다음
  // (더 나중) 배정일로 합친다(트럭 상한/생산capa를 넘기지 않을 때만) — 트럭 1대(TRUCK_KG)의 절반.
  MIN_SHIPMENT_KG: 2500,
  LIGHT_DAY_TRUCK_THRESHOLD: 2, // 여러 업체를 합쳐 이 대수 미만인 날을 "가벼운 날"로 선호(하드캡 아님, 여유일 고를 때만 씀) — 이미 이 대수에 도달했으면 트럭 하나 더 실을 여유가 없다고 봄
  DEEP_DIVE_MARKER: '딥다이브', // 발주서 시트 1행에서 이 텍스트를 찾아 그 열부터 표가 시작한다고 봄
};

// 실제 업체명 → 배송 타입. 발주서 시트명의 "고객사 X" 표기에서 뽑아낸 "고객사X"와 정확히 일치해야
// 한다. 업체 표기는 "X사"가 아니라 "고객사X" 형태로 통일해서 쓴다.
// 세 업체 모두 트럭버킷(요일별 기준일 + 트럭 1대 상한 + 라운드로빈) — 요일만 다르다.
const VENDOR_TYPE_MAP = {
  '고객사A': 'monday_bucket',
  '고객사B': 'wednesday_bucket',
  '고객사C': 'friday_bucket',
};

// 트럭버킷 타입 → 기준 요일(0=일 ... 6=토). TruckBucket.gs가 이 표로 요일을 찾는다.
const BUCKET_WEEKDAYS = {
  monday_bucket: 1,
  wednesday_bucket: 3,
  friday_bucket: 5,
};

// 배송 타입별 처리 함수는 Allocate.gs(date_as_is/friday_even), TruckBucket.gs(*_bucket) 참고
const ALLOCATION_TYPES = ['date_as_is', 'friday_even', 'monday_bucket', 'wednesday_bucket', 'friday_bucket'];

// 발주서 시트 2행 헤더 인식 키워드 (code/weight는 필수, spec/date는 있으면 쓰고 없어도 무방)
const ORDER_COLUMN_KEYWORDS = {
  code: ['품목코드', 'sc코드'],
  spec: ['규격', '사양'],
  weight: ['중량', '총중량'],
  date: ['납기일', '납품일자', '납기', '일자'],
};
