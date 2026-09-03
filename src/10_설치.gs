/**
 * 10_설치.gs — 최초 1회 설치. 관리 시트 탭을 만들고, 미분류 폴더를 만들고, 초기 규칙을 채운다.
 * 여러 번 실행해도 안전하다 (이미 있으면 건드리지 않는다).
 */

/** 시트를 열 때 메뉴를 붙인다. */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('발주서 자동화')
    .addItem('▶ 전체 실행 (드라이브 + 메일)', '전체_실행_메뉴')
    .addItem('↻ 초기화 후 전체 재실행', '전체_재실행_메뉴')
    .addSeparator()
    .addItem('설치 / 점검', '설치_메뉴')
    .addItem('트리거 설치 (주기 수집)', '트리거_설치_메뉴')
    .addItem('트리거 제거', '트리거_제거_메뉴')
    .addSeparator()
    .addSubMenu(
      SpreadsheetApp.getUi().createMenu('단계별 실행')
        .addItem('1. 규칙 로드', '규칙_로드_테스트_메뉴')
        .addItem('2. 추출', '추출_테스트_메뉴')
        .addItem('3. 처리 (추출→매핑→검산)', '처리_테스트_메뉴')
        .addItem('4. 식별', '식별_테스트_메뉴')
        .addItem('5. 드라이브 수집', '기록_테스트_메뉴')
        .addItem('6. 메일 수집', '메일수집_테스트_메뉴')
    )
    .addSeparator()
    .addItem('미분류 목록 보기', '미분류_목록_메뉴')
    .addItem('산출물 초기화', '산출물_초기화_메뉴')
    .addToUi();
}

/**
 * 설치. 이미 있는 것은 그대로 두고 없는 것만 만든다.
 *
 * 여기서는 절대 대화상자를 띄우지 않는다. 스크립트 편집기에서 실행하면 대화상자를
 * 보여줄 화면이 없어서 실행이 응답을 기다리며 매달릴 수 있다.
 * 팝업이 필요하면 시트 메뉴에서 부르는 설치_메뉴() 를 쓴다.
 */
function 설치() {
  const 결과 = [];

  결과.push(...탭_준비());
  결과.push(미분류폴더_준비());
  결과.push(고객사식별_초기화());

  const 요약 = 결과.join('\n');
  Logger.log(요약);
  return 요약;
}

/** 시트 메뉴용 래퍼. 설치 후 결과를 팝업으로 보여준다. */
function 설치_메뉴() {
  const 요약 = 설치();
  const ui = SpreadsheetApp.getUi();
  ui.alert('설치 / 점검 결과', 요약, ui.ButtonSet.OK);
}

/** 관리 시트에 필요한 탭 4개를 만든다. 기본 시트(Sheet1/시트1)는 설정 탭으로 재사용한다. */
function 탭_준비() {
  const ss = 관리시트();
  const 로그 = [];

  const 정의 = [
    { 이름: 탭.설정, 헤더: ['키', '값', '설명'] },
    { 이름: 탭.고객사식별, 헤더: ['고객사', '검사대상', '키워드', '가중치'] },
    { 이름: 탭.처리로그, 헤더: ['일시', '채널', '원본', '고객사', '판정근거', '결과', '비고'] },
    { 이름: 탭.미매핑, 헤더: ['일시', '채널', '원본', '고객사', '사유', '원본행'] },
  ];

  정의.forEach(function (d) {
    let sheet = ss.getSheetByName(d.이름);
    if (!sheet) {
      // 손 안 댄 기본 시트가 남아 있으면 이름만 바꿔 재사용한다
      const 여분 = ss.getSheets().filter(function (s) {
        return /^(Sheet1|시트1)$/.test(s.getName()) && s.getLastRow() === 0;
      })[0];
      if (여분) {
        sheet = 여분.setName(d.이름);
      } else {
        sheet = ss.insertSheet(d.이름);
      }
      sheet.getRange(1, 1, 1, d.헤더.length).setValues([d.헤더]).setFontWeight('bold');
      sheet.setFrozenRows(1);
      로그.push('탭 생성: ' + d.이름);
    } else {
      로그.push('탭 있음: ' + d.이름);
    }
  });

  return 로그;
}

/** 고객사 첨부파일 폴더 아래에 미분류 폴더를 만들고 ID 를 설정 탭에 기록한다. */
function 미분류폴더_준비() {
  const 기존 = 설정전체()['폴더.미분류'];
  if (기존) {
    try {
      DriveApp.getFolderById(기존);
      return '미분류 폴더 있음: ' + 기존;
    } catch (e) {
      // 설정에 적힌 ID 가 죽었으면 새로 만든다
    }
  }

  const 부모 = DriveApp.getFolderById(설정값('폴더.첨부'));
  const it = 부모.getFoldersByName(미분류폴더명);
  const folder = it.hasNext() ? it.next() : 부모.createFolder(미분류폴더명);

  설정값쓰기('폴더.미분류', folder.getId(), '고객사 판별 실패 건이 들어가는 폴더');
  return '미분류 폴더 준비: ' + folder.getId();
}

/**
 * 고객사식별 탭에 초기 규칙을 채운다. 이미 데이터가 있으면 건드리지 않는다.
 *
 * 가중치 설계 — 파일명·메일제목·메일본문은 OCR 을 타지 않으므로 높게,
 * 문서내용은 이미지로 들어오면 OCR 오인식 위험이 있으므로 낮게.
 * 발신자 행은 주소를 모르니 키워드를 비워 두고, 로더가 빈 키워드를 건너뛴다.
 */
function 고객사식별_초기화() {
  const sheet = 관리시트().getSheetByName(탭.고객사식별);
  if (!sheet) throw new Error('고객사식별 탭이 없습니다.');
  if (sheet.getLastRow() > 1) {
    return '고객사식별 규칙 있음 (' + (sheet.getLastRow() - 1) + '행) — 건드리지 않음';
  }

  // 문서내용 판별용 고유 마커. 발주서 데이터에서 확인한 값.
  // 주의: `딥다이브`(3사 공통), `AVSS 2.0SQ`·`37/0.260`(A·C 공통)은 구분에 쓸 수 없다.
  const 문서마커 = {
    고객사A: 'ZWCSS',
    고객사B: '동선발주',
    고객사C: 'EP-0101',
  };

  const rows = [];
  고객사목록.forEach(function (고객사) {
    rows.push([고객사, '발신자', '', 10]); // 주소 확보되면 키워드만 채우면 된다
    rows.push([고객사, '파일명', 고객사, 10]);
    rows.push([고객사, '메일제목', 고객사, 8]);
    rows.push([고객사, '메일본문', 고객사, 6]);
    rows.push([고객사, '문서내용', 문서마커[고객사] || '', 4]);
  });

  sheet.getRange(2, 1, rows.length, 4).setValues(rows);
  return '고객사식별 초기 규칙 ' + rows.length + '행 입력';
}
