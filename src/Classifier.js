/**
 * 부서 귀속 판정 로직 — 키워드 우선순위 (스펙 3장)
 *
 * 판정 순서 (먼저 매치되는 규칙 우선):
 *  1. 프로젝트 전용 prefix (아텍 고유 프로그램명) → 자동 포함
 *  2. 강한 확정 키워드 (아텍/ATC 등)            → 자동 포함
 *  3. 교수 성함 포함 + 타학과 키워드 미포함      → 자동 포함
 *  4. 타학과 키워드 포함 + 교수 성함 미포함      → 자동 제외
 *  5. 둘 다 있거나(충돌) 둘 다 없음             → 검토대기함
 */

var CLS = { INCLUDE: 'INCLUDE', EXCLUDE: 'EXCLUDE', REVIEW: 'REVIEW' };

/**
 * @param {string} text 전표 텍스트
 * @param {Object} kw getKeywordLists_() 결과 {prefix, strong, professor, otherDept}
 * @return {{verdict: string, reason: string}}
 */
function classifyText_(text, kw) {
  var t = normStr_(text);
  var tUpper = t.toUpperCase();

  // 1. 프로젝트 전용 prefix — 아텍만 쓰는 고유 프로그램명이므로 텍스트 포함 여부로 판정
  var hit = findKeyword_(tUpper, kw.prefix);
  if (hit) return { verdict: CLS.INCLUDE, reason: '프로젝트prefix: ' + hit };

  // 2. 강한 확정 키워드
  hit = findKeyword_(tUpper, kw.strong);
  if (hit) return { verdict: CLS.INCLUDE, reason: '강한키워드: ' + hit };

  // 3~5. 교수 성함 vs 타학과 키워드
  var prof = findKeyword_(tUpper, kw.professor);
  var other = findKeyword_(tUpper, kw.otherDept);

  if (prof && !other) return { verdict: CLS.INCLUDE, reason: '교수성함: ' + prof };
  if (other && !prof) return { verdict: CLS.EXCLUDE, reason: '타학과키워드: ' + other };
  if (prof && other) {
    return { verdict: CLS.REVIEW, reason: '충돌 (교수성함: ' + prof + ' / 타학과키워드: ' + other + ')' };
  }
  return { verdict: CLS.REVIEW, reason: '판정 단서 없음' };
}

/**
 * 텍스트(대문자 정규화됨)에 포함된 첫 번째 키워드 반환
 * - 영문 대소문자 무시
 * - 2단어 이상 키워드는 공백 무시 비교 병행
 *   (실데이터에 'ArtWithImpact'가 'Art with impact'로 적히는 사례가 있음)
 * @return {string|null}
 */
function findKeyword_(textUpper, keywords) {
  var textNoSpace = textUpper.replace(/\s+/g, '');
  for (var i = 0; i < keywords.length; i++) {
    var k = normStr_(keywords[i]);
    if (!k) continue;
    var kUpper = k.toUpperCase();
    if (textUpper.indexOf(kUpper) >= 0) return k;
    var kNoSpace = kUpper.replace(/\s+/g, '');
    // 공백 제거 비교는 짧은 키워드의 우연한 일치(오탐)를 막기 위해 4자 이상만
    if (kNoSpace.length >= 4 && textNoSpace.indexOf(kNoSpace) >= 0) return k;
  }
  return null;
}
