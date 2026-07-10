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
    .addItem('시트 초기화/점검', 'initializeSheets')
    .addToUi();
}

function menuRebuildSummary() {
  var result = rebuildMonthlySummary();
  SpreadsheetApp.getActive().toast(
    '월별집계 재계산 완료 — 라인 ' + result.lines + '개 / 월 ' + result.months + '개',
    'AXT 예산', 5
  );
}

function showKeywordSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle('AXT 키워드 관리');
  SpreadsheetApp.getUi().showSidebar(html);
}
