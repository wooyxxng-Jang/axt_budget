/**
 * 코드별 보기 — 월별집계 '코드' 열(회의비/ATC/활동비 등, 사람이 드롭다운으로 지정)로
 * 라인을 묶어 배정액/사용합계/잔액을 코드 단위로 확인하는 파생 뷰.
 * - 월별집계를 그대로 읽어 재구성만 하므로, rebuildMonthlySummary()가 실행될 때마다
 *   함께 갱신된다 (Summary.js 참고). 이 시트를 직접 편집해도 다음 재계산 시 덮어써진다.
 */

var CODEVIEW_HEADERS = ['코드', '자금', '자금명', '약정항목', '약정항목 명', '상세분류', '배정액', '사용합계', '잔액'];

function rebuildCodeView_() {
  var ss = SpreadsheetApp.getActive();
  var summarySh = ss.getSheetByName(SHEET.SUMMARY);
  var sh = ensureSheet_(SHEET.CODE_VIEW, null);

  sh.clearContents();
  sh.getRange(1, 1, 1, CODEVIEW_HEADERS.length).setValues([CODEVIEW_HEADERS]).setFontWeight('bold');
  sh.setFrozenRows(1);
  sh.setFrozenColumns(1);
  if (!summarySh) return { lines: 0, codes: 0 };

  var last = summarySh.getLastRow();
  var lastCol = summarySh.getLastColumn();
  if (last < 2 || lastCol < 1) return { lines: 0, codes: 0 };

  var headerRow = summarySh.getRange(1, 1, 1, lastCol).getValues()[0].map(normStr_);
  var idx = {};
  ['코드', '자금', '자금명', '약정항목', '약정항목 명', '상세분류', '배정액', '사용합계'].forEach(function (h) {
    idx[h] = headerRow.indexOf(h);
  });
  if (idx['코드'] < 0) return { lines: 0, codes: 0 }; // 코드 열 도입 이전(최초 실행 전)

  var data = summarySh.getRange(2, 1, last - 1, lastCol).getValues();
  var groups = {}; // code → summary rows[]
  var order = [];
  data.forEach(function (row) {
    var code = normStr_(row[idx['코드']]);
    if (!code) return;
    if (!groups[code]) { groups[code] = []; order.push(code); }
    groups[code].push(row);
  });

  var out = [];
  var totalRowIdx = [];
  order.forEach(function (code) {
    var budgetSum = 0, totalSum = 0;
    groups[code].forEach(function (row) {
      var budget = Number(row[idx['배정액']]) || 0;
      var used = Number(row[idx['사용합계']]) || 0;
      budgetSum += budget;
      totalSum += used;
      out.push([
        code, row[idx['자금']], row[idx['자금명']], row[idx['약정항목']], row[idx['약정항목 명']],
        row[idx['상세분류']], budget, used, budget - used
      ]);
    });
    totalRowIdx.push(out.length);
    out.push([code + ' 합계', '', '', '', '', '', budgetSum, totalSum, budgetSum - totalSum]);
  });

  if (out.length) {
    sh.getRange(2, 1, out.length, CODEVIEW_HEADERS.length).setValues(out);
    sh.getRange(2, 7, out.length, 3).setNumberFormat('#,##0');
    totalRowIdx.forEach(function (i) {
      sh.getRange(2 + i, 1, 1, CODEVIEW_HEADERS.length).setFontWeight('bold');
    });
  }
  return { lines: out.length, codes: order.length };
}
