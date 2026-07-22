/**
 * 배정예산 마스터에 '상세분류명' 라벨 컬럼을 채우는 1회성 유지보수 스크립트.
 *
 * 마스터의 '상세분류 / 근거'는 긴 설명문이라 대기함 드롭다운/월별집계 행 이름으로 쓰기 부적합.
 * 대신 짧은 라벨을 담는 '상세분류명' 컬럼을 만들고, 실데이터(세인트 401건) 검증 결과를 바탕으로
 * 32개 그룹(42개 세부라인) 각각에 라벨을 채운다. 라벨은 '2026 예산' 탭의 2~43행 순서와 정확히 대응.
 *
 * 실행: Apps Script 편집기에서 addMasterDetailLabels 함수를 1회 실행 (이미 라벨이 있으면 스킵, 재실행 안전)
 */

var MASTER_LABELS_2026_ = [
  '교내장학금-기금', '교육연구용 기계장치', 'Kakao회의비', 'AWI회의비', '기타운영비',
  'Kakao제작지원', 'AWI이월장학', 'Kakao오리엔테이션', 'AWI결과발표회', '특별강의료',
  '회의비', '일반행사비', '기타운영비', '테크니컬조교1학기', '테크니컬조교2학기',
  '미디어조교', '온사이트멘토링', 'ATC컨퍼런스작품제작', '일반업무추진비', '교육연구용 기계장치',
  '특별강의료', '유지보수비', '실험실습교재비', '수업준비지원비', '교육연구용 기계장치',
  '특별강의료', '유지보수비', '기타운영비', '실습실조교', '행사지원조교',
  '실험실습교재비', '수업준비지원비', '온라인홍보', '귀빈접대비', '학생활동지원비',
  '학생지도경비', '일반행사비', '학생회활동', 'ATC컨퍼런스', '학생지도경비',
  '학생활동지원비', '학생지도경비'
];

function addMasterDetailLabels() {
  var masterId = getSetting_(SETTING_KEY.MASTER_ID);
  var sheetName = getSetting_(SETTING_KEY.MASTER_SHEET) || '2026 예산';
  if (!masterId) throw new Error('설정 시트에 ' + SETTING_KEY.MASTER_ID + '가 없습니다.');

  var ss = SpreadsheetApp.openById(masterId);
  var sh = ss.getSheetByName(sheetName);
  if (!sh) throw new Error('마스터에서 시트를 찾을 수 없습니다: ' + sheetName);

  var headerRow = 1;
  var lastCol = sh.getLastColumn();
  var headers = sh.getRange(headerRow, 1, 1, lastCol).getValues()[0];

  var descCol = -1, labelCol = -1;
  for (var i = 0; i < headers.length; i++) {
    var h = normStr_(headers[i]).replace(/\s+/g, '');
    if (h.indexOf('상세분류') >= 0 && descCol < 0) descCol = i + 1;
    if (h === '상세분류명') labelCol = i + 1;
  }
  if (descCol < 0) throw new Error("'상세분류' 헤더를 찾을 수 없습니다.");

  if (labelCol < 0) {
    sh.insertColumnAfter(descCol);
    labelCol = descCol + 1;
    sh.getRange(headerRow, labelCol).setValue('상세분류명');
  }

  var existing = sh.getRange(headerRow + 1, labelCol, MASTER_LABELS_2026_.length, 1).getValues();
  var alreadyFilled = existing.some(function (r) { return normStr_(r[0]); });
  if (alreadyFilled) {
    throw new Error('상세분류명 컬럼(' + labelCol + '열)에 이미 값이 있습니다. 실수로 덮어쓰지 않도록 중단했습니다. ' +
      '필요하면 해당 열을 비우고 다시 실행하세요.');
  }

  sh.getRange(headerRow + 1, labelCol, MASTER_LABELS_2026_.length, 1)
    .setValues(MASTER_LABELS_2026_.map(function (v) { return [v]; }));

  return { descCol: descCol, labelCol: labelCol, rows: MASTER_LABELS_2026_.length };
}

/**
 * 설정 시트 키워드 '유형' 열의 예전 한글 값(강한제외키워드/프로젝트prefix/...)을
 * 새 영문 값(Strong Exclude/Project Prefix/...)으로 1회 변환한다.
 * 이미 영문이거나 빈 값인 행은 건드리지 않음 — 여러 번 실행해도 안전.
 */
function migrateKeywordTypeLabels() {
  var sh = getSettingsSheet_();
  var last = sh.getLastRow();
  if (last <= KW_HEADER_ROW) return { changed: 0 };

  var range = sh.getRange(KW_HEADER_ROW + 1, 1, last - KW_HEADER_ROW, 1);
  var vals = range.getValues();
  var changed = 0;
  for (var i = 0; i < vals.length; i++) {
    var old = normStr_(vals[i][0]);
    if (old && KW_TYPE_LEGACY_MAP[old]) {
      vals[i][0] = KW_TYPE_LEGACY_MAP[old];
      changed++;
    }
  }
  if (changed) range.setValues(vals);

  SpreadsheetApp.getUi().alert(
    changed
      ? '키워드 유형 ' + changed + '건을 영문으로 변환했습니다.'
      : '변환할 한글 유형 값이 없습니다 (이미 영문이거나 비어 있음).'
  );
  return { changed: changed };
}
