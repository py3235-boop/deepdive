/**
 * 출고계획 설정값
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
  PLAN_SHEET_NAME: '출고계획', // 출고계획 결과를 쓸 탭 이름 — 이 이름의 탭(없으면 첫 탭)에 그대로 덮어씀
  // 하루 상한은 항상 업체별 독립으로 정확히 트럭 1대 — 물량이 많으면 트럭을 더 싣는 게 아니라
  // 그 업체의 배송일 자체를 늘린다.
  TRUCK_KG: 5000,
  LIGHT_DAY_TRUCK_THRESHOLD: 2, // 여러 업체를 합쳐 이 대수 이하인 날을 "가벼운 날"로 선호(하드캡 아님, 여유일 고를 때만 씀)
  DEEP_DIVE_MARKER: '딥다이브', // 발주서 시트 1행에서 이 텍스트를 찾아 그 열부터 표가 시작한다고 봄
};

// 실제 업체명 → 배송 타입. 발주서 시트명의 "고객사 X" 표기에서 뽑아낸 "X사"와 정확히 일치해야 한다.
// A/B사는 트럭버킷(요일별 기준일 + 트럭 1대 상한 + 라운드로빈), C사는 금요일 단순 균등분배.
const VENDOR_TYPE_MAP = {
  'A사': 'monday_bucket',
  'B사': 'wednesday_bucket',
  'C사': 'friday_even',
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
