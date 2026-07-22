/**
 * 시트 초기화/점검 — 7개 시트 생성 + 헤더/포맷/데이터검증/설정 시드 (스펙 6장)
 * 여러 번 실행해도 안전(멱등): 기존 데이터는 건드리지 않고 빠진 부분만 채운다.
 */

var SEED_KEYWORDS = [
  [KW_TYPE.STRONG_EXCLUDE, '알바트로스'], // 지융미 공용 세미나 — 아텍 교수 참여분도 아텍 예산 아님, 무조건 제외
  [KW_TYPE.PREFIX, 'Kakao테크포임팩트수업'],
  [KW_TYPE.PREFIX, 'ArtWithImpact'],
  [KW_TYPE.STRONG, '아텍'],
  [KW_TYPE.STRONG, 'ATC'],
  [KW_TYPE.PROFESSOR, '최용순'],
  [KW_TYPE.PROFESSOR, '김주섭'],
  [KW_TYPE.PROFESSOR, '김상용'],
  [KW_TYPE.PROFESSOR, '정다샘'],
  [KW_TYPE.PROFESSOR, '정해동'],
  [KW_TYPE.PROFESSOR, '오준호'],
  [KW_TYPE.PROFESSOR, '이선옥'],
  [KW_TYPE.PROFESSOR, '주진호'],
  [KW_TYPE.OTHER_DEPT, '캡스톤'],
  [KW_TYPE.OTHER_DEPT, '지융미'],
  [KW_TYPE.OTHER_DEPT, '신방'],
  [KW_TYPE.OTHER_DEPT, '미엔'],
  [KW_TYPE.OTHER_DEPT, '미컴원']
];

// 스펙 4장 B의 라우팅 규칙 — 상세분류명은 addMasterDetailLabels()로 마스터에 채운 라벨과 동일
// (OneTimeSetup.js MASTER_LABELS_2026_ 참고). 실데이터 401건 검증으로 확인된 매칭.
var SEED_RULES = [
  ['FGC9361000', '42381110', '포함', 'Kakao테크포임팩트수업', 'Kakao회의비'],
  ['FGC9361000', '42381110', '포함', '아텍', 'AWI회의비'],
  ['FGC9361000', '42381110', '포함', 'ArtWithImpact', 'AWI회의비'],
  ['TBC0184', '42371110', '기본값', '', '온라인홍보'],
  ['GBC2061', '43251010', '포함', '작품 제작', 'ATC컨퍼런스작품제작']
];

// 배정예산 마스터: 우영님 구글시트, '2026 예산' 탭이 회계연도 내 살아있는 마스터
// (스프레드시트에는 예산/ATC/대학혁신/기금인출/기금/연도별 정리본 등 다른 탭도 있으니 혼동 주의)
// 회계연도가 바뀌면 MASTER_SHEET_NAME을 '2027 예산' 등으로 직접 갱신해야 함
var SEED_GENERAL = [
  [SETTING_KEY.MASTER_ID, '1avOTbMw9dw4tMZC6sMLClbHoKbKMuoCTpL3qo-mwqAI'],
  [SETTING_KEY.MASTER_SHEET, '2026 예산']
];

function initializeSheets() {
  var ss = SpreadsheetApp.getActive();

  // 데이터 시트들
  setupDataSheet_(SHEET.LEDGER, LEDGER_HEADERS);
  setupDataSheet_(SHEET.SUMMARY, null);
  var review = setupDataSheet_(SHEET.REVIEW, REVIEW_HEADERS);
  var detailQ = setupDataSheet_(SHEET.DETAIL_QUEUE, DETAILQ_HEADERS);
  var schol = setupDataSheet_(SHEET.SCHOLARSHIP, SCHOLARSHIP_HEADERS);
  setupSettingsSheet_();
  setupDataSheet_(SHEET.AUDIT, AUDIT_HEADERS);

  // 고정 드롭다운: 검토대기함 '판정', 장학금 대기함 '귀속'
  setColumnValidation_(review, REVIEW_HEADERS.indexOf('판정') + 1, VERDICT_VALUES);
  setColumnValidation_(schol, SCHOLARSHIP_HEADERS.indexOf('귀속') + 1, TAG_VALUES);

  // 시트 순서 정리
  var order = [SHEET.SUMMARY, SHEET.LEDGER, SHEET.REVIEW, SHEET.DETAIL_QUEUE, SHEET.SCHOLARSHIP, SHEET.SETTINGS, SHEET.AUDIT];
  order.forEach(function (name, i) {
    var sh = ss.getSheetByName(name);
    if (sh) {
      ss.setActiveSheet(sh);
      ss.moveActiveSheet(i + 1);
    }
  });
  ss.setActiveSheet(ss.getSheetByName(SHEET.SUMMARY));

  SpreadsheetApp.getUi().alert(
    '시트 초기화/점검 완료\n\n' +
    '배정예산 마스터(구글시트 "2026 예산" 탭)가 설정 시트 J~K열에 기본 연결되었습니다.\n\n' +
    '다음 단계:\n' +
    '1. 마스터 시트에 "상세분류명" 컬럼을 추가해 상세분류별 짧은 라벨을 입력\n' +
    '   (라벨을 넣지 않으면 상세분류/근거 설명문 첫 줄을 자동 사용)\n' +
    '2. 설정 시트 D~H열 라우팅 규칙의 상세분류명을 1번에서 정한 라벨로 교체\n' +
    '3. 메뉴 [AXT 예산 > 세인트 파일 업로드] 실행\n\n' +
    '※ 회계연도가 바뀌면 설정 시트의 MASTER_SHEET_NAME 값을 다음 연도 탭명으로 직접 갱신해야 합니다.'
  );
}

/** 헤더 + 텍스트/금액 포맷 세팅 */
function setupDataSheet_(name, headers) {
  var sh = ensureSheet_(name, headers);
  if (!headers) return sh;

  sh.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  if (sh.getFrozenRows() < 1) sh.setFrozenRows(1);

  // 코드/일자류는 텍스트 포맷 (숫자 자동변환 방지), 사용금액은 통화 포맷
  var textCols = ['업서트키', '자금', '약정항목', '참조전표', '개별항목', '전표일자', '지급일자'];
  var maxRows = sh.getMaxRows();
  headers.forEach(function (h, i) {
    if (textCols.indexOf(h) >= 0) {
      sh.getRange(2, i + 1, maxRows - 1, 1).setNumberFormat('@');
    } else if (h === '사용금액') {
      sh.getRange(2, i + 1, maxRows - 1, 1).setNumberFormat('#,##0');
    }
  });
  return sh;
}

function setColumnValidation_(sheet, col, values) {
  if (col < 1) return;
  var rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(values, true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange(2, col, sheet.getMaxRows() - 1, 1).setDataValidation(rule);
}

/** 설정 시트: 3개 블록(키워드/라우팅규칙/일반설정) 헤더 + 시드 */
function setupSettingsSheet_() {
  var sh = ensureSheet_(SHEET.SETTINGS, null);

  // 섹션 제목 + 블록 헤더
  sh.getRange(1, 1).setValue('■ 키워드 (사이드바 또는 직접 편집)').setFontWeight('bold');
  sh.getRange(KW_HEADER_ROW, 1, 1, 2).setValues([['유형', '키워드']]).setFontWeight('bold');

  sh.getRange(1, RULE_COL_START).setValue(
    '■ 상세분류 라우팅 규칙 (매칭방식: 포함/prefix/기본값) — 상세분류 열은 마스터 시트 "상세분류명" 컬럼의 라벨과 정확히 일치해야 함. ' +
    '라벨 자체를 바꾸려면 마스터 시트에서 수정하고, 여기 값도 함께 갱신할 것 (메뉴의 [라우팅 규칙 검증]으로 어긋난 값 확인 가능)'
  ).setFontWeight('bold').setWrap(true);
  sh.getRange(KW_HEADER_ROW, RULE_COL_START, 1, RULE_HEADERS.length).setValues([RULE_HEADERS]).setFontWeight('bold');

  sh.getRange(1, GENERAL_COL_START).setValue('■ 일반 설정').setFontWeight('bold');
  sh.getRange(KW_HEADER_ROW, GENERAL_COL_START, 1, 2).setValues([['설정키', '값']]).setFontWeight('bold');

  if (sh.getFrozenRows() < 2) sh.setFrozenRows(2);

  // 시드: 각 블록이 비어 있을 때만 채움 (기존 운영 데이터 보존)
  var dataRow = KW_HEADER_ROW + 1;
  if (!normStr_(sh.getRange(dataRow, 1).getValue())) {
    sh.getRange(dataRow, 1, SEED_KEYWORDS.length, 2).setValues(SEED_KEYWORDS);
  }
  if (!normStr_(sh.getRange(dataRow, RULE_COL_START).getValue())) {
    sh.getRange(dataRow, RULE_COL_START, SEED_RULES.length, RULE_HEADERS.length).setValues(SEED_RULES);
  }
  if (!normStr_(sh.getRange(dataRow, GENERAL_COL_START).getValue())) {
    sh.getRange(dataRow, GENERAL_COL_START, SEED_GENERAL.length, 2).setValues(SEED_GENERAL);
  }

  // 자금/약정항목 코드 컬럼은 텍스트 포맷
  var maxRows = sh.getMaxRows();
  sh.getRange(dataRow, RULE_COL_START, maxRows - KW_HEADER_ROW, 2).setNumberFormat('@');

  // 매칭방식 드롭다운
  var rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(RULE_MODES, true)
    .setAllowInvalid(true)
    .build();
  sh.getRange(dataRow, RULE_COL_START + 2, maxRows - KW_HEADER_ROW, 1).setDataValidation(rule);

  // 키워드 유형 드롭다운 — 사이드바 없이 셀 직접 편집 시 오타 방지
  var typeRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(KW_TYPE_LIST, true)
    .setAllowInvalid(false)
    .build();
  sh.getRange(dataRow, 1, maxRows - KW_HEADER_ROW, 1).setDataValidation(typeRule);

  return sh;
}

/** 메뉴: 라우팅 규칙(D~H열)의 상세분류 값이 마스터 라벨과 실제로 일치하는지 점검 */
function validateRoutingRules() {
  var ui = SpreadsheetApp.getUi();
  var rules = getRoutingRules_();
  if (!rules.length) { ui.alert('등록된 라우팅 규칙이 없습니다.'); return; }

  var problems = [];
  rules.forEach(function (r, i) {
    var candidates = getDetailCandidates_(r.fund, r.itemCode);
    if (candidates.length === 0) candidates = getFundDetailCandidates_(r.fund);
    if (candidates.indexOf(r.detail) < 0) {
      problems.push('행 ' + (i + 1) + ': ' + r.fund + '(' + r.itemCode + ') → "' + r.detail +
        '" — 마스터 후보에 없음 (후보: ' + (candidates.join(', ') || '없음') + ')');
    }
  });

  ui.alert(
    problems.length
      ? '어긋난 라우팅 규칙 ' + problems.length + '건\n\n' + problems.join('\n')
      : '모든 라우팅 규칙이 마스터 라벨과 일치합니다.'
  );
}
