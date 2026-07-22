/**
 * 마스터 시트에서 상세분류명(라벨)을 수정했을 때, 이미 원장/대기함/라우팅 규칙에
 * 텍스트로 박혀 있는 옛 라벨을 새 라벨로 일괄 교체하는 유지보수 기능.
 *
 * 마스터의 상세분류명은 매 실행마다 실시간으로 읽히므로, 새로 분류되는 건이나
 * 아직 실적이 없는 배정예산 라인에는 라벨 변경이 즉시 반영된다.
 * 하지만 원장/대기함에 이미 옛 라벨로 기록된 행과, 설정 시트 라우팅 규칙의
 * 상세분류 참조값은 그 시점의 텍스트로 저장돼 있어 자동으로 바뀌지 않는다 —
 * 이 함수가 그 간극을 메운다.
 */
function renameDetailLabel() {
  var ui = SpreadsheetApp.getUi();

  var r1 = ui.prompt('상세분류명 변경 반영', '마스터에서 바꾸기 전(기존) 라벨을 정확히 입력하세요:', ui.ButtonSet.OK_CANCEL);
  if (r1.getSelectedButton() !== ui.Button.OK) return;
  var oldLabel = normStr_(r1.getResponseText());
  if (!oldLabel) { ui.alert('기존 라벨을 입력해 주세요.'); return; }

  var r2 = ui.prompt('상세분류명 변경 반영', '"' + oldLabel + '" → 바꾼 후(새) 라벨을 입력하세요:', ui.ButtonSet.OK_CANCEL);
  if (r2.getSelectedButton() !== ui.Button.OK) return;
  var newLabel = normStr_(r2.getResponseText());
  if (!newLabel) { ui.alert('새 라벨을 입력해 주세요.'); return; }
  if (newLabel === oldLabel) { ui.alert('기존 라벨과 새 라벨이 같습니다.'); return; }

  var ss = SpreadsheetApp.getActive();
  var summary = [];
  var total = 0;

  // 원장 / 장학금 분류 대기함 / 상세분류 확인 대기열의 상세분류 컬럼 일괄 교체
  var targets = [
    { sheet: SHEET.LEDGER, headers: LEDGER_HEADERS, col: '상세분류' },
    { sheet: SHEET.SCHOLARSHIP, headers: SCHOLARSHIP_HEADERS, col: '상세분류' },
    { sheet: SHEET.DETAIL_QUEUE, headers: DETAILQ_HEADERS, col: '선택상세분류' }
  ];
  targets.forEach(function (t) {
    var sh = ss.getSheetByName(t.sheet);
    if (!sh) return;
    var last = sh.getLastRow();
    var colIdx = t.headers.indexOf(t.col) + 1;
    if (last < 2 || colIdx < 1) return;

    var range = sh.getRange(2, colIdx, last - 1, 1);
    var vals = range.getValues();
    var changed = 0;
    for (var i = 0; i < vals.length; i++) {
      if (normStr_(vals[i][0]) === oldLabel) { vals[i][0] = newLabel; changed++; }
    }
    if (changed) {
      range.setValues(vals);
      total += changed;
      summary.push(t.sheet + '.' + t.col + ': ' + changed + '건');
    }
  });

  // 설정 시트 라우팅 규칙(상세분류 열)도 함께 교체
  var settingsSh = getSettingsSheet_();
  var lastRow = settingsSh.getLastRow();
  var detailColIdx = RULE_COL_START + RULE_HEADERS.indexOf('상세분류');
  if (lastRow > KW_HEADER_ROW && detailColIdx >= RULE_COL_START) {
    var ruleRange = settingsSh.getRange(KW_HEADER_ROW + 1, detailColIdx, lastRow - KW_HEADER_ROW, 1);
    var ruleVals = ruleRange.getValues();
    var ruleChanged = 0;
    for (var j = 0; j < ruleVals.length; j++) {
      if (normStr_(ruleVals[j][0]) === oldLabel) { ruleVals[j][0] = newLabel; ruleChanged++; }
    }
    if (ruleChanged) {
      ruleRange.setValues(ruleVals);
      total += ruleChanged;
      summary.push(SHEET.SETTINGS + '.라우팅규칙: ' + ruleChanged + '건');
    }
  }

  var result = rebuildMonthlySummary();

  ui.alert(
    '상세분류명 변경 반영 완료\n\n"' + oldLabel + '" → "' + newLabel + '"\n\n' +
    (summary.length ? summary.join('\n') : '변경된 행이 없습니다 (해당 라벨을 참조하는 곳이 없음).') +
    '\n\n월별집계 재계산 완료 (라인 ' + result.lines + '개).\n' +
    '※ [라우팅 규칙 검증] 메뉴로 다른 어긋남이 없는지 한 번 더 확인해 보세요.'
  );
}
