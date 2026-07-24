/**
 * 설정 시트 — 키워드 리스트 / 상세분류 라우팅 규칙 / 일반 설정 (스펙 3, 3-1, 4장)
 *
 * 시트 레이아웃:
 *  - A~B열  : 키워드 블록   (2행 헤더 '유형','키워드', 3행부터 데이터)
 *  - D~H열  : 상세분류 라우팅 규칙 (2행 헤더, 3행부터 데이터)
 *  - J~K열  : 일반 설정     (2행 헤더 '설정키','값', 3행부터 데이터)
 * 사이드바 UI 없이 셀을 직접 편집해도 동작한다.
 */

var KW_HEADER_ROW = 2;
var RULE_COL_START = 4; // D열
var RULE_HEADERS = ['자금', '약정항목', '매칭방식', '키워드', '상세분류'];
var RULE_MODES = ['포함', 'prefix', '기본값'];
var GENERAL_COL_START = 10; // J열

function getSettingsSheet_() {
  return ensureSheet_(SHEET.SETTINGS, null);
}

// ---------------------------------------------------------------------------
// 키워드 블록
// ---------------------------------------------------------------------------

/** 유형별 키워드 목록 { strongExclude, prefix, strong, professor, otherDept } */
function getKeywordLists_() {
  var raw = readKeywordRows_();
  var out = { strongExclude: [], prefix: [], strong: [], professor: [], otherDept: [] };
  raw.forEach(function (r) {
    if (r.type === KW_TYPE.STRONG_EXCLUDE) out.strongExclude.push(r.value);
    else if (r.type === KW_TYPE.PREFIX) out.prefix.push(r.value);
    else if (r.type === KW_TYPE.STRONG) out.strong.push(r.value);
    else if (r.type === KW_TYPE.PROFESSOR) out.professor.push(r.value);
    else if (r.type === KW_TYPE.OTHER_DEPT) out.otherDept.push(r.value);
  });
  return out;
}

/** 키워드 블록 원본 행 [{type, value}] */
function readKeywordRows_() {
  var sh = getSettingsSheet_();
  var last = sh.getLastRow();
  if (last <= KW_HEADER_ROW) return [];
  var vals = sh.getRange(KW_HEADER_ROW + 1, 1, last - KW_HEADER_ROW, 2).getValues();
  var rows = [];
  vals.forEach(function (v) {
    var type = normStr_(v[0]), value = normStr_(v[1]);
    if (type && value) rows.push({ type: type, value: value });
  });
  return rows;
}

/** 키워드 블록 전체 재기록 (유형 순서대로 정렬해 유지) */
function writeKeywordRows_(rows) {
  var sh = getSettingsSheet_();
  var last = sh.getLastRow();
  if (last > KW_HEADER_ROW) {
    sh.getRange(KW_HEADER_ROW + 1, 1, last - KW_HEADER_ROW, 2).clearContent();
  }
  rows.sort(function (a, b) {
    var ta = KW_TYPE_LIST.indexOf(a.type), tb = KW_TYPE_LIST.indexOf(b.type);
    if (ta !== tb) return ta - tb;
    return a.value.localeCompare(b.value, 'ko');
  });
  if (rows.length) {
    var vals = rows.map(function (r) { return [r.type, r.value]; });
    sh.getRange(KW_HEADER_ROW + 1, 1, vals.length, 2).setValues(vals);
  }
}

// --- 사이드바에서 호출하는 서버 API -----------------------------------------

/** 사이드바 초기 로딩용: 유형 목록 + 유형별 키워드 */
function sbGetKeywords() {
  var rows = readKeywordRows_();
  var data = {};
  KW_TYPE_LIST.forEach(function (t) { data[t] = []; });
  rows.forEach(function (r) {
    if (!data[r.type]) data[r.type] = [];
    data[r.type].push(r.value);
  });
  return { types: KW_TYPE_LIST, data: data };
}

function sbAddKeyword(type, value) {
  validateKwType_(type);
  value = normStr_(value);
  if (!value) throw new Error('키워드를 입력해 주세요.');
  var rows = readKeywordRows_();
  var dup = rows.some(function (r) {
    return r.type === type && r.value.toUpperCase() === value.toUpperCase();
  });
  if (dup) throw new Error('이미 등록된 키워드입니다: ' + value);
  rows.push({ type: type, value: value });
  writeKeywordRows_(rows);
  return sbGetKeywords();
}

function sbUpdateKeyword(type, oldValue, newValue) {
  validateKwType_(type);
  newValue = normStr_(newValue);
  if (!newValue) throw new Error('새 키워드를 입력해 주세요.');
  var rows = readKeywordRows_();
  var found = false;
  rows.forEach(function (r) {
    if (!found && r.type === type && r.value === normStr_(oldValue)) {
      r.value = newValue;
      found = true;
    }
  });
  if (!found) throw new Error('수정할 키워드를 찾지 못했습니다: ' + oldValue);
  var seen = {};
  rows = rows.filter(function (r) {
    var k = r.type + '|' + r.value.toUpperCase();
    if (seen[k]) return false;
    seen[k] = true;
    return true;
  });
  writeKeywordRows_(rows);
  return sbGetKeywords();
}

function sbDeleteKeyword(type, value) {
  validateKwType_(type);
  var rows = readKeywordRows_();
  var next = rows.filter(function (r) { return !(r.type === type && r.value === normStr_(value)); });
  if (next.length === rows.length) throw new Error('삭제할 키워드를 찾지 못했습니다: ' + value);
  writeKeywordRows_(next);
  return sbGetKeywords();
}

function validateKwType_(type) {
  if (KW_TYPE_LIST.indexOf(type) < 0) throw new Error('알 수 없는 키워드 유형: ' + type);
}

/** 월별집계 '코드' 열 드롭다운용 등록된 코드 목록 (키워드 관리 사이드바에서 'Summary Code' 유형으로 등록/삭제) */
function getSummaryCodes_() {
  return readKeywordRows_()
    .filter(function (r) { return r.type === KW_TYPE.SUMMARY_CODE; })
    .map(function (r) { return r.value; });
}

// ---------------------------------------------------------------------------
// 상세분류 라우팅 규칙 블록 (스펙 4장 B)
// ---------------------------------------------------------------------------

/** [{fund, itemCode, mode, keyword, detail}] */
function getRoutingRules_() {
  var sh = getSettingsSheet_();
  var last = sh.getLastRow();
  if (last <= KW_HEADER_ROW) return [];
  var vals = sh.getRange(KW_HEADER_ROW + 1, RULE_COL_START, last - KW_HEADER_ROW, RULE_HEADERS.length).getValues();
  var rules = [];
  vals.forEach(function (v) {
    var fund = normStr_(v[0]), itemCode = normStr_(v[1]);
    var mode = normStr_(v[2]) || '포함';
    var keyword = normStr_(v[3]), detail = normStr_(v[4]);
    if (!fund || !itemCode || !detail) return;
    if (mode !== '기본값' && !keyword) return;
    rules.push({ fund: fund, itemCode: itemCode, mode: mode, keyword: keyword, detail: detail });
  });
  return rules;
}

// ---------------------------------------------------------------------------
// 일반 설정 블록
// ---------------------------------------------------------------------------

function getSetting_(key) {
  var sh = getSettingsSheet_();
  var last = sh.getLastRow();
  if (last <= KW_HEADER_ROW) return '';
  var vals = sh.getRange(KW_HEADER_ROW + 1, GENERAL_COL_START, last - KW_HEADER_ROW, 2).getValues();
  for (var i = 0; i < vals.length; i++) {
    if (normStr_(vals[i][0]) === key) return normStr_(vals[i][1]);
  }
  return '';
}
