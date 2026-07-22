/**
 * AXT 예산 집행 자동화 — 공통 상수 및 유틸리티
 * 스펙: AXT_예산자동화_스펙.md
 */

var SHEET = {
  LEDGER: '원장',
  SUMMARY: '월별집계',
  REVIEW: '검토대기함',
  DETAIL_QUEUE: '상세분류 확인 대기열',
  SCHOLARSHIP: '장학금 분류 대기함',
  SETTINGS: '설정',
  AUDIT: '감사로그'
};

// 세인트 다운로드 파일의 핵심 컬럼 (스펙 2장)
var SRC_COLS = [
  '수입/지출', '자금', '자금명', '약정항목', '약정항목 명',
  '전표상태', '전표유형명', '참조전표', '개별항목',
  '전표일자', '사용금액', '텍스트', '지급일자'
];

var LEDGER_HEADERS = ['업서트키'].concat(SRC_COLS, ['상세분류', '판정근거', '반영경로', '반영일시']);
var REVIEW_HEADERS = ['업서트키'].concat(SRC_COLS, ['보류사유', '판정', '상태', '처리일시']);
var SCHOLARSHIP_HEADERS = ['업서트키'].concat(SRC_COLS, ['귀속', '상세분류', '상태', '처리일시']);
var DETAILQ_HEADERS = ['업서트키'].concat(SRC_COLS, ['판정근거', '후보상세분류', '선택상세분류', '상태', '처리일시']);
var AUDIT_HEADERS = ['업서트키'].concat(SRC_COLS, ['제외사유', '기록일시']);

// 키워드 유형 (설정 시트에서 관리, 스펙 3장)
// 값은 설정 시트의 '유형' 열에 그대로 저장되는 문자열이므로 영문으로 통일한다.
// 기존 한글 값이 남아있는 시트는 migrateKeywordTypeLabels()로 1회 변환한다 (OneTimeSetup.js).
var KW_TYPE = {
  STRONG_EXCLUDE: 'Strong Exclude',  // 교수성함 등과 무관하게 무조건 제외 (예: 알바트로스세미나)
  PREFIX: 'Project Prefix',          // 아텍 고유 프로그램명
  STRONG: 'Strong Include',          // 아텍/ATC 등
  PROFESSOR: 'Professor Name',
  OTHER_DEPT: 'Other Dept'
};
var KW_TYPE_LIST = [KW_TYPE.STRONG_EXCLUDE, KW_TYPE.PREFIX, KW_TYPE.STRONG, KW_TYPE.PROFESSOR, KW_TYPE.OTHER_DEPT];

// 마이그레이션용: 이전(한글) 유형 값 → 새 영문 값
var KW_TYPE_LEGACY_MAP = {
  '강한제외키워드': KW_TYPE.STRONG_EXCLUDE,
  '프로젝트prefix': KW_TYPE.PREFIX,
  '강한키워드': KW_TYPE.STRONG,
  '교수성함': KW_TYPE.PROFESSOR,
  '타학과키워드': KW_TYPE.OTHER_DEPT
};

var STATUS = { PENDING: '대기', DONE: '반영완료', TO_DETAIL: '상세분류대기' };
var PATH = { AUTO: '자동', REVIEW: '검토확정', SCHOLARSHIP: '장학금태깅', DETAIL: '상세분류확정' };
var VERDICT_VALUES = ['포함', '제외'];
var TAG_VALUES = ['아텍', '타학과'];

// 설정 시트 일반 설정 키
var SETTING_KEY = {
  MASTER_ID: 'MASTER_SPREADSHEET_ID',
  MASTER_SHEET: 'MASTER_SHEET_NAME'
};

// 장학금(등록금 유형) 판별 기준 (스펙 5장)
var SCHOLARSHIP_DOC_TYPE = '등록금';

// ---------------------------------------------------------------------------
// 유틸리티
// ---------------------------------------------------------------------------

/** upsert 키: 참조전표 + 개별항목 (스펙 2장) */
function makeKey_(refDoc, itemNo) {
  return normStr_(refDoc) + '|' + normStr_(itemNo);
}

/** 자금+약정항목 그룹 키 (상세분류 라우팅 단위, 스펙 4장) */
function groupKey_(fund, itemCode) {
  return normStr_(fund) + '|' + normStr_(itemCode);
}

function normStr_(v) {
  return String(v === null || v === undefined ? '' : v).trim();
}

/** "1,223,800", "1,000-", "(500)" 등 금액 문자열 → 숫자 */
function parseAmount_(v) {
  if (typeof v === 'number') return v;
  var s = normStr_(v).replace(/[,\s원₩]/g, '');
  if (!s) return 0;
  var neg = false;
  if (/^\(.*\)$/.test(s)) { neg = true; s = s.slice(1, -1); }
  if (/-$/.test(s)) { neg = true; s = s.slice(0, -1); } // SAP식 후행 마이너스
  if (/^-/.test(s)) { neg = true; s = s.slice(1); }
  var n = Number(s);
  if (isNaN(n)) return 0;
  return neg ? -n : n;
}

/** 다양한 날짜 표기 → 'yyyy-MM-dd' 문자열 (실패 시 '') */
function parseDateStr_(v) {
  if (v instanceof Date && !isNaN(v.getTime())) {
    return Utilities.formatDate(v, 'Asia/Seoul', 'yyyy-MM-dd');
  }
  var s = normStr_(v);
  if (!s) return '';
  var m = s.match(/^(\d{4})[.\-\/]?\s?(\d{1,2})[.\-\/]?\s?(\d{1,2})/);
  if (!m) return '';
  var y = m[1], mo = ('0' + m[2]).slice(-2), d = ('0' + m[3]).slice(-2);
  return y + '-' + mo + '-' + d;
}

/** 'yyyy-MM-dd' → 'yyyy-MM' (월별집계 컬럼 키) */
function monthKey_(dateStr) {
  var s = normStr_(dateStr);
  return /^\d{4}-\d{2}/.test(s) ? s.substring(0, 7) : '';
}

function nowStr_() {
  return Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss');
}

/** 시트를 가져오고 없으면 생성 + 헤더 세팅 */
function ensureSheet_(name, headers) {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
  }
  if (headers && headers.length) {
    var firstRow = sh.getRange(1, 1, 1, headers.length).getValues()[0];
    var mismatch = headers.some(function (h, i) { return normStr_(firstRow[i]) !== h; });
    if (mismatch) {
      sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
      if (sh.getFrozenRows() < 1) sh.setFrozenRows(1);
    }
  }
  return sh;
}

/** 키 컬럼(A열, 2행부터) → Set 형태 인덱스 */
function buildKeySet_(sheet) {
  var set = {};
  var last = sheet.getLastRow();
  if (last < 2) return set;
  var vals = sheet.getRange(2, 1, last - 1, 1).getValues();
  for (var i = 0; i < vals.length; i++) {
    var k = normStr_(vals[i][0]);
    if (k) set[k] = true;
  }
  return set;
}

/** 소스 행 객체 → 시트 행 배열(업서트키 + SRC_COLS 순서) */
function srcToRowPrefix_(src) {
  var row = [makeKey_(src['참조전표'], src['개별항목'])];
  for (var i = 0; i < SRC_COLS.length; i++) row.push(src[SRC_COLS[i]]);
  return row;
}
