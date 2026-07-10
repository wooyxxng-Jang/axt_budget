/**
 * 배정예산 마스터(기존 구글시트) 참조 (스펙 2장)
 * - 설정 시트의 MASTER_SPREADSHEET_ID / MASTER_SHEET_NAME 으로 접근 (읽기 전용)
 *
 * 실제 마스터 레이아웃 (2026 예산 시트 기준):
 *  - 컬럼: 자금명 | 약정항목 | 자금/약정항목코드 | 상세분류 / 근거 | 배정예산 | 3월~2월 | 잔액
 *  - 코드는 'FCC4361000(43221040)' 형태로 자금+약정항목이 한 셀에 결합
 *  - 같은 그룹의 상세분류(세부라인)가 여러 행일 때 자금명/약정항목/코드 칸이 비어 있음(병합 스타일)
 *  - '상세분류명' 컬럼(짧은 라벨, 운영자가 추가)이 있으면 그것을 라벨로 사용.
 *    없거나 비어 있으면 '상세분류 / 근거' 설명문의 첫 줄(40자)로 대체, 그것도 없으면 '라인N'
 */

var MASTER_CACHE_ = null;

/**
 * @return {{available: boolean,
 *           lines: Array<{fund,fundName,itemCode,itemName,detail,budget}>,
 *           detailsByGroup: Object<string, string[]>}}
 */
function getMasterIndex_() {
  if (MASTER_CACHE_) return MASTER_CACHE_;

  var id = getSetting_(SETTING_KEY.MASTER_ID);
  if (!id) {
    MASTER_CACHE_ = { available: false, lines: [], detailsByGroup: {} };
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

  MASTER_CACHE_ = parseMasterValues_(sh.getDataRange().getValues());
  return MASTER_CACHE_;
}

/**
 * 마스터 시트 2차원 배열 → 라인/그룹 인덱스 (순수 함수, 로컬 테스트 가능)
 */
function parseMasterValues_(data) {
  var empty = { available: false, lines: [], detailsByGroup: {} };
  if (!data || data.length < 2) return empty;

  // 헤더 행 탐색 (첫 10행 내)
  var headerRowIdx = -1, colMap = null;
  for (var r = 0; r < Math.min(10, data.length); r++) {
    var m = mapMasterHeader_(data[r]);
    if (m) { headerRowIdx = r; colMap = m; break; }
  }
  if (!colMap) {
    throw new Error('배정예산 마스터에서 헤더(자금/약정항목코드, 배정예산 등)를 찾을 수 없습니다.');
  }

  var lines = [];
  var detailsByGroup = {};
  // 병합 스타일 빈 칸 forward-fill 상태
  var curFund = '', curItemCode = '', curFundName = '', curItemName = '';

  for (var i = headerRowIdx + 1; i < data.length; i++) {
    var row = data[i];
    var codeCell = normStr_(row[colMap.code]);
    var descCell = colMap.desc >= 0 ? normStr_(row[colMap.desc]) : '';
    var labelCell = colMap.label >= 0 ? normStr_(row[colMap.label]) : '';
    var budgetCell = colMap.budget >= 0 ? row[colMap.budget] : '';
    var hasBudget = budgetCell !== '' && budgetCell !== null && budgetCell !== undefined;

    if (codeCell) {
      var parsed = parseCombinedCode_(codeCell);
      if (parsed) {
        curFund = parsed.fund;
        curItemCode = parsed.itemCode;
        if (colMap.fundName >= 0 && normStr_(row[colMap.fundName])) curFundName = normStr_(row[colMap.fundName]);
        if (colMap.itemName >= 0 && normStr_(row[colMap.itemName])) curItemName = normStr_(row[colMap.itemName]);
      }
    } else {
      // 코드 없는 행이라도 자금명/약정항목명이 새로 나오면 갱신 (다음 코드 행 대비)
      if (colMap.fundName >= 0 && normStr_(row[colMap.fundName])) curFundName = normStr_(row[colMap.fundName]);
      if (colMap.itemName >= 0 && normStr_(row[colMap.itemName])) curItemName = normStr_(row[colMap.itemName]);
    }

    // 라인으로 인정하는 조건: 현재 그룹이 있고, 배정액 또는 설명/라벨이 존재
    if (!curFund || !curItemCode) continue;
    if (!hasBudget && !descCell && !labelCell) continue;

    var gk = groupKey_(curFund, curItemCode);
    if (!detailsByGroup[gk]) detailsByGroup[gk] = [];

    var label = labelCell || firstLine_(descCell, 40);
    if (!label) label = '라인' + (detailsByGroup[gk].length + 1);
    // 그룹 내 라벨 중복 시 번호 붙여 구분
    var base = label, n = 2;
    while (detailsByGroup[gk].indexOf(label) >= 0) {
      label = base + ' (' + n + ')';
      n++;
    }

    lines.push({
      fund: curFund,
      fundName: curFundName,
      itemCode: curItemCode,
      itemName: curItemName,
      detail: label,
      budget: hasBudget ? parseAmount_(budgetCell) : ''
    });
    detailsByGroup[gk].push(label);
  }

  // 상세분류가 1개뿐인 그룹은 라벨을 비목 레벨로 흡수 (라벨명 그대로 두되 후보는 1개)
  return { available: true, lines: lines, detailsByGroup: detailsByGroup };
}

/**
 * 'FCC4361000(43221040)' 형태(괄호 누락 등 변형 포함) → {fund, itemCode}
 * 약정항목코드는 항상 8자리 숫자라는 점을 이용해 뒤 8자리를 분리한다.
 */
function parseCombinedCode_(s) {
  var flat = normStr_(s).replace(/[\s()\-]/g, '');
  var m = flat.match(/^([A-Za-z0-9]+?)(\d{8})$/);
  if (!m) return null;
  if (!/[A-Za-z]/.test(m[1])) return null; // 자금코드는 영문 포함
  return { fund: m[1].toUpperCase(), itemCode: m[2] };
}

/** 여러 줄 텍스트의 첫 번째 비어있지 않은 줄 (maxLen자 초과 시 말줄임) */
function firstLine_(s, maxLen) {
  var t = String(s === null || s === undefined ? '' : s);
  var linesArr = t.split(/\r?\n/);
  for (var i = 0; i < linesArr.length; i++) {
    var line = linesArr[i].trim();
    if (line) {
      return line.length > maxLen ? line.substring(0, maxLen) + '…' : line;
    }
  }
  return '';
}

/** 헤더 행 후보 → 컬럼 매핑 (실패 시 null) */
function mapMasterHeader_(row) {
  var norm = row.map(function (c) { return normStr_(c).replace(/\s+/g, ''); });
  function find(names) {
    for (var i = 0; i < names.length; i++) {
      for (var j = 0; j < norm.length; j++) {
        if (norm[j] === names[i]) return j;
      }
    }
    return -1;
  }
  function findContains(pattern) {
    for (var j = 0; j < norm.length; j++) {
      if (norm[j] && pattern.test(norm[j])) return j;
    }
    return -1;
  }
  // 결합 코드 컬럼: '자금/약정항목코드' 등 '코드'가 들어간 헤더
  var code = findContains(/자금.*약정.*코드|약정.*자금.*코드|^코드$/);
  if (code < 0) return null;
  var budget = find(['배정예산', '배정액', '예산액']);
  if (budget < 0) return null;

  // '상세분류명'(짧은 라벨) 컬럼과 '상세분류 / 근거'(설명문) 컬럼 구분
  var label = find(['상세분류명', '상세분류라벨']);
  var desc = -1;
  for (var j = 0; j < norm.length; j++) {
    if (j !== label && /상세분류/.test(norm[j])) { desc = j; break; }
  }

  return {
    code: code,
    fundName: find(['자금명']),
    itemName: find(['약정항목', '약정항목명', '비목명']),
    label: label,
    desc: desc,
    budget: budget
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
