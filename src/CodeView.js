/**
 * 코드별 보기 — 월별집계 '코드' 열(회의비/ATC/활동비 등, 사람이 드롭다운으로 지정)로
 * 라인을 묶어 배정액/사용합계/잔액을 코드 단위로 확인하는 파생 뷰.
 * - 월별집계를 그대로 읽어 재구성만 하므로, rebuildMonthlySummary()가 실행될 때마다
 *   함께 갱신된다 (Summary.js 참고).
 * - '소진 계획'/'비고'는 사람이 직접 적는 값이라, 재계산 시에도 같은 라인(코드+자금+
 *   약정항목+상세분류) 또는 같은 코드의 합계 행을 기준으로 보존한다.
 * - 코드 그룹 사이는 빈 행 하나로 구분해 가독성을 높인다.
 */

var CODEVIEW_HEADERS = ['코드', '자금', '자금명', '약정항목', '약정항목 명', '상세분류', '배정액', '사용합계', '잔액', '소진 계획', '비고'];

function rebuildCodeView_() {
  var ss = SpreadsheetApp.getActive();
  var summarySh = ss.getSheetByName(SHEET.SUMMARY);
  var sh = ensureSheet_(SHEET.CODE_VIEW, null);

  var notesMap = readExistingCodeViewNotes_(sh);

  sh.clearContents();
  sh.getRange(1, 1, 1, CODEVIEW_HEADERS.length).setValues([CODEVIEW_HEADERS])
    .setFontWeight('bold').setBackground(SHEET_HEADER_BG_);
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
  order.forEach(function (code, gi) {
    var budgetSum = 0, totalSum = 0;
    groups[code].forEach(function (row) {
      var fund = row[idx['자금']], itemCode = row[idx['약정항목']], detail = row[idx['상세분류']];
      var budget = Number(row[idx['배정액']]) || 0;
      var used = Number(row[idx['사용합계']]) || 0;
      budgetSum += budget;
      totalSum += used;
      var prev = notesMap[codeViewLineKey_(code, fund, itemCode, detail)] || {};
      out.push([
        code, fund, row[idx['자금명']], itemCode, row[idx['약정항목 명']],
        detail, budget, used, budget - used, prev.plan || '', prev.note || ''
      ]);
    });
    totalRowIdx.push(out.length);
    var prevTotal = notesMap[codeViewTotalKey_(code)] || {};
    out.push([
      code + ' 합계', '', '', '', '', '', budgetSum, totalSum, budgetSum - totalSum,
      prevTotal.plan || '', prevTotal.note || ''
    ]);
    if (gi < order.length - 1) {
      out.push(new Array(CODEVIEW_HEADERS.length).fill('')); // 코드 그룹 사이 구분용 빈 행
    }
  });

  if (out.length) {
    sh.getRange(2, 1, out.length, CODEVIEW_HEADERS.length).setValues(out);
    sh.getRange(2, 7, out.length, 3).setNumberFormat('#,##0');
    totalRowIdx.forEach(function (i) {
      sh.getRange(2 + i, 1, 1, CODEVIEW_HEADERS.length)
        .setFontWeight('bold').setBackground(TOTAL_ROW_BG_)
        .setBorder(null, null, true, null, null, null);
    });
  }
  return { lines: out.length, codes: order.length };
}

function codeViewLineKey_(code, fund, itemCode, detail) {
  return code + '|' + normStr_(fund) + '|' + normStr_(itemCode) + '|' + normStr_(detail);
}
function codeViewTotalKey_(code) {
  return code + '|__TOTAL__';
}

/** 재계산 직전의 코드별 보기 시트에서 라인/합계 행별 '소진 계획'/'비고' 값을 읽어온다 */
function readExistingCodeViewNotes_(sh) {
  var map = {};
  var last = sh.getLastRow();
  if (last < 2) return map;
  var lastCol = sh.getLastColumn();
  if (lastCol < 1) return map;
  var headerRow = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(normStr_);
  var cCode = headerRow.indexOf('코드');
  var cFund = headerRow.indexOf('자금');
  var cItem = headerRow.indexOf('약정항목');
  var cDetail = headerRow.indexOf('상세분류');
  var cPlan = headerRow.indexOf('소진 계획');
  var cNote = headerRow.indexOf('비고');
  if (cCode < 0 || cPlan < 0 || cNote < 0) return map; // 이전 버전 시트(컬럼 도입 이전)

  var vals = sh.getRange(2, 1, last - 1, lastCol).getValues();
  vals.forEach(function (row) {
    var codeCell = normStr_(row[cCode]);
    if (!codeCell) return; // 구분용 빈 행
    var plan = normStr_(row[cPlan]);
    var note = normStr_(row[cNote]);
    if (!plan && !note) return;
    var key = / 합계$/.test(codeCell)
      ? codeViewTotalKey_(codeCell.replace(/ 합계$/, ''))
      : codeViewLineKey_(codeCell, row[cFund], cItem >= 0 ? row[cItem] : '', row[cDetail]);
    map[key] = { plan: plan, note: note };
  });
  return map;
}
