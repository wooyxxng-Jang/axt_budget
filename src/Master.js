/**
 * 배정예산 마스터(기존 구글시트) 참조 (스펙 2장)
 * - 설정 시트의 MASTER_SPREADSHEET_ID / MASTER_SHEET_NAME 으로 접근
 * - 자금코드 / 약정항목코드 / 상세분류 / 배정액 컬럼을 헤더명으로 유연하게 탐색
 * - 마스터 미연결 시에도 비목 레벨 집계는 동작하도록 available=false 로 폴백
 */

var MASTER_CACHE_ = null;

/**
 * @return {{available: boolean,
 *           lines: Array<{fund,fundName,itemCode,itemName,detail,budget}>,
 *           detailsByGroup: Object<string, string[]>}}
 */
function getMasterIndex_() {
  if (MASTER_CACHE_) return MASTER_CACHE_;

  var empty = { available: false, lines: [], detailsByGroup: {} };
  var id = getSetting_(SETTING_KEY.MASTER_ID);
  if (!id) {
    MASTER_CACHE_ = empty;
    return MASTER_CACHE_;
  }

  var ss;
  try {
    ss = SpreadsheetApp.openById(id);
  } catch (e) {
    throw new Error('배정예산 마스터 시트를 열 수 없습니다. 설정 시트의 ' + SETTING_KEY.MASTER_ID + ' 값을 확인해 주세요.\n' + e.message);
  }
  var sheetName = getSetting_(SETTING_KEY.MASTER_SHEET);
  var sh = sheetName ? ss.getSheetByName(sheetName) : ss.getSheets()[0];
  if (!sh) {
    throw new Error('배정예산 마스터에서 시트를 찾을 수 없습니다: ' + sheetName);
  }

  var data = sh.getDataRange().getValues();
  if (data.length < 2) {
    MASTER_CACHE_ = empty;
    return MASTER_CACHE_;
  }

  // 헤더 탐색 (첫 10행 내에서 자금/약정항목 계열 헤더가 있는 행)
  var headerRowIdx = -1, colMap = null;
  for (var r = 0; r < Math.min(10, data.length); r++) {
    var m = mapMasterHeader_(data[r]);
    if (m) { headerRowIdx = r; colMap = m; break; }
  }
  if (!colMap) {
    throw new Error('배정예산 마스터에서 헤더(자금코드/약정항목코드/상세분류/배정액)를 찾을 수 없습니다.');
  }

  var lines = [];
  var detailsByGroup = {};
  for (var i = headerRowIdx + 1; i < data.length; i++) {
    var row = data[i];
    var fund = normStr_(row[colMap.fund]);
    var itemCode = normStr_(row[colMap.itemCode]);
    if (!fund && !itemCode) continue;
    var line = {
      fund: fund,
      fundName: colMap.fundName >= 0 ? normStr_(row[colMap.fundName]) : '',
      itemCode: itemCode,
      itemName: colMap.itemName >= 0 ? normStr_(row[colMap.itemName]) : '',
      detail: colMap.detail >= 0 ? normStr_(row[colMap.detail]) : '',
      budget: colMap.budget >= 0 ? parseAmount_(row[colMap.budget]) : ''
    };
    lines.push(line);
    var gk = groupKey_(fund, itemCode);
    if (!detailsByGroup[gk]) detailsByGroup[gk] = [];
    if (line.detail && detailsByGroup[gk].indexOf(line.detail) < 0) {
      detailsByGroup[gk].push(line.detail);
    }
  }

  MASTER_CACHE_ = { available: true, lines: lines, detailsByGroup: detailsByGroup };
  return MASTER_CACHE_;
}

/** 헤더 행 후보 → 컬럼 매핑 (실패 시 null) */
function mapMasterHeader_(row) {
  var norm = row.map(function (c) { return normStr_(c).replace(/\s+/g, ''); });
  function find(names) {
    for (var i = 0; i < names.length; i++) {
      var idx = norm.indexOf(names[i]);
      if (idx >= 0) return idx;
    }
    return -1;
  }
  var fund = find(['자금코드', '자금']);
  var itemCode = find(['약정항목코드', '약정항목']);
  if (fund < 0 || itemCode < 0) return null;
  return {
    fund: fund,
    fundName: find(['자금명']),
    itemCode: itemCode,
    itemName: find(['약정항목명', '약정항목 명', '비목명']),
    detail: find(['상세분류', '세부항목', '세부라인']),
    budget: find(['배정액', '배정예산', '예산액'])
  };
}

/** 해당 자금+약정항목 그룹의 상세분류 후보 목록 */
function getDetailCandidates_(fund, itemCode) {
  var master = getMasterIndex_();
  return master.detailsByGroup[groupKey_(fund, itemCode)] || [];
}

/**
 * 원장 행의 '유효 상세분류' — 월별집계/드릴다운 공통 규칙
 *  - 원장에 상세분류가 기록돼 있으면 그대로
 *  - 없고 그룹 상세분류가 1개면 그 값
 *  - 없고 그룹 상세분류가 여러 개면 '(미지정)'
 */
function effectiveDetail_(recordedDetail, fund, itemCode) {
  var d = normStr_(recordedDetail);
  if (d) return d;
  var candidates = getDetailCandidates_(fund, itemCode);
  if (candidates.length === 1) return candidates[0];
  if (candidates.length > 1) return '(미지정)';
  return '';
}
