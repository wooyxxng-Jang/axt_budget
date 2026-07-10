/**
 * 원장 드릴다운 (스펙 6장)
 * - 월별집계에서 월(yyyy-MM) 또는 사용합계 셀을 선택하고 메뉴 실행
 * - 해당 라인(자금/약정항목/상세분류) × 월의 원장 구성 내역을 모달로 표시
 */

function showDrilldown() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getActiveSheet();

  if (sh.getName() !== SHEET.SUMMARY) {
    ui.alert('월별집계 시트에서 금액 셀을 선택한 뒤 실행해 주세요.');
    return;
  }
  var cell = sh.getActiveRange();
  var row = cell.getRow(), col = cell.getColumn();
  if (row < 2) {
    ui.alert('데이터 행(2행 이하)의 월별 금액 셀을 선택해 주세요.');
    return;
  }

  var headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(normStr_);
  var header = headers[col - 1];
  var isMonth = /^\d{4}-\d{2}$/.test(header) || header === '(일자없음)';
  var isTotal = header === '사용합계';
  if (!isMonth && !isTotal) {
    ui.alert('월(yyyy-MM) 컬럼 또는 사용합계 컬럼의 셀을 선택해 주세요.\n(선택된 컬럼: ' + header + ')');
    return;
  }

  var lineVals = sh.getRange(row, 1, 1, 5).getValues()[0];
  var fund = normStr_(lineVals[0]);
  var itemCode = normStr_(lineVals[2]);
  var detail = normStr_(lineVals[4]);

  var cRef = LEDGER_HEADERS.indexOf('참조전표');
  var cItemNo = LEDGER_HEADERS.indexOf('개별항목');
  var cFund = LEDGER_HEADERS.indexOf('자금');
  var cItem = LEDGER_HEADERS.indexOf('약정항목');
  var cDate = LEDGER_HEADERS.indexOf('전표일자');
  var cAmt = LEDGER_HEADERS.indexOf('사용금액');
  var cText = LEDGER_HEADERS.indexOf('텍스트');
  var cDetail = LEDGER_HEADERS.indexOf('상세분류');
  var cPath = LEDGER_HEADERS.indexOf('반영경로');

  var matches = [];
  var total = 0;
  readLedgerData_().forEach(function (r) {
    if (normStr_(r[cFund]) !== fund || normStr_(r[cItem]) !== itemCode) return;
    if (effectiveDetail_(r[cDetail], fund, itemCode) !== detail) return;
    var mk = monthKey_(r[cDate]) || '(일자없음)';
    if (isMonth && mk !== header) return;
    var amt = Number(r[cAmt]) || 0;
    total += amt;
    matches.push({
      date: normStr_(r[cDate]),
      ref: normStr_(r[cRef]),
      itemNo: normStr_(r[cItemNo]),
      text: normStr_(r[cText]),
      amt: amt,
      path: normStr_(r[cPath])
    });
  });
  matches.sort(function (a, b) { return a.date < b.date ? -1 : a.date > b.date ? 1 : 0; });

  var title = fund + ' / ' + itemCode + (detail ? ' / ' + detail : '') + ' — ' + (isTotal ? '전체' : header);
  var html = buildDrilldownHtml_(title, matches, total);
  ui.showModalDialog(
    HtmlService.createHtmlOutput(html).setWidth(780).setHeight(520),
    '원장 드릴다운'
  );
}

function buildDrilldownHtml_(title, matches, total) {
  var rows = matches.map(function (m) {
    return '<tr>' +
      '<td>' + esc_(m.date) + '</td>' +
      '<td>' + esc_(m.ref) + '</td>' +
      '<td>' + esc_(m.itemNo) + '</td>' +
      '<td class="txt">' + esc_(m.text) + '</td>' +
      '<td class="num">' + formatNum_(m.amt) + '</td>' +
      '<td>' + esc_(m.path) + '</td>' +
      '</tr>';
  }).join('');

  return '<style>' +
    'body{font-family:Arial,"Malgun Gothic",sans-serif;font-size:12px;margin:12px}' +
    'h3{margin:0 0 4px 0;font-size:14px}' +
    '.meta{color:#666;margin-bottom:10px}' +
    'table{border-collapse:collapse;width:100%}' +
    'th,td{border:1px solid #ccc;padding:4px 6px;text-align:left;vertical-align:top}' +
    'th{background:#f1f3f4;position:sticky;top:0}' +
    'td.num,th.num{text-align:right;white-space:nowrap}' +
    'td.txt{word-break:break-all}' +
    'tfoot td{font-weight:bold;background:#fafafa}' +
    '</style>' +
    '<h3>' + esc_(title) + '</h3>' +
    '<div class="meta">' + matches.length + '건 / 합계 ' + formatNum_(total) + '원</div>' +
    '<table><thead><tr>' +
    '<th>전표일자</th><th>참조전표</th><th>개별항목</th><th>텍스트</th><th class="num">사용금액</th><th>반영경로</th>' +
    '</tr></thead><tbody>' +
    (rows || '<tr><td colspan="6">해당 내역이 없습니다.</td></tr>') +
    '</tbody><tfoot><tr><td colspan="4">합계</td><td class="num">' + formatNum_(total) + '</td><td></td></tr></tfoot></table>';
}

function esc_(s) {
  return String(s === null || s === undefined ? '' : s)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

function formatNum_(n) {
  return String(Math.round(Number(n) || 0)).replace(/\B(?=(\d{3})+(?!\d))/g, ',');
}
