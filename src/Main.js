/**
 * 메뉴 및 사이드바 진입점
 */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('AXT 예산')
    .addItem('세인트 파일 업로드', 'showUploadDialog')
    .addSeparator()
    .addItem('검토대기함 → 확정 반영', 'applyReviewDecisions')
    .addItem('장학금 태깅 → 반영', 'applyScholarshipTags')
    .addItem('상세분류 선택 → 반영', 'applyDetailSelections')
    .addSeparator()
    .addItem('월별집계 재계산', 'menuRebuildSummary')
    .addItem('선택 셀 드릴다운 (월별집계)', 'showDrilldown')
    .addSeparator()
    .addItem('키워드 관리 (사이드바)', 'showKeywordSidebar')
    .addItem('라우팅 규칙 검증', 'validateRoutingRules')
    .addItem('상세분류명 변경 반영', 'renameDetailLabel')
    .addItem('시트 초기화/점검', 'initializeSheets')
    .addSeparator()
    .addItem('[1회성] 마스터 상세분류명 라벨 채우기', 'menuAddMasterDetailLabels')
    .addItem('[1회성] 키워드 유형 영문 변환', 'migrateKeywordTypeLabels')
    .addToUi();
}

function menuAddMasterDetailLabels() {
  var result = addMasterDetailLabels();
  SpreadsheetApp.getUi().alert(
    '마스터 상세분류명 라벨 채우기 완료\n\n' +
    result.rows + '개 행에 라벨을 입력했습니다 (열 ' + result.labelCol + ').\n' +
    '설정 시트 D~H열의 라우팅 규칙에 이 라벨을 반영해 주세요.'
  );
}

function menuRebuildSummary() {
  var result = rebuildMonthlySummary();
  SpreadsheetApp.getActive().toast(
    '월별집계 재계산 완료 — 라인 ' + result.lines + '개 / 월 ' + result.months + '개 (코드별 보기도 함께 갱신됨)',
    'AXT 예산', 5
  );
}

function showKeywordSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle('AXT 키워드 관리');
  SpreadsheetApp.getUi().showSidebar(html);
}
