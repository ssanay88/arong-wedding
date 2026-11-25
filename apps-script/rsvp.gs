/**
 * 청첩장 RSVP 수신용 Google Apps Script
 *
 * 설치 방법
 * 1. 구글 시트를 새로 만든다.
 * 2. 시트 메뉴 [확장 프로그램] → [Apps Script] 를 열고, 이 파일 내용을 전부 붙여넣고 저장한다.
 * 3. [배포] → [새 배포] → 유형: 웹 앱
 *    - 실행 사용자: 나
 *    - 액세스 권한: 모든 사용자
 *    → 배포 후 권한 승인, 나오는 "웹 앱 URL"(https://script.google.com/macros/s/.../exec) 을 복사한다.
 * 4. index.html CONFIG.rsvpEndpoint 에 그 URL 을 넣는다.
 *
 * 코드를 수정한 뒤에는 [배포] → [배포 관리] → 연필 → 버전: 새 버전 으로 다시 배포해야 반영된다.
 * (새 배포를 만들면 URL 이 바뀌므로 index.html 도 같이 바꿔야 함)
 */

const SHEET_NAME = "RSVP";
const HEADER = ["제출 시각", "구분", "참석 여부", "성함", "인원"];
const SIDE = { groom: "신랑측", bride: "신부측" };
const ATTEND = { yes: "참석", no: "불참" };

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);

    const d = JSON.parse(e.postData.contents);
    const side = SIDE[d.side];
    const attend = ATTEND[d.attend];
    const name = String(d.name || "").trim().slice(0, 30);
    const count = Math.min(Math.max(parseInt(d.count, 10) || 1, 1), 5);
    if (!side || !attend || !name) return json({ ok: false, error: "invalid" });

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
    if (sh.getLastRow() === 0) {
      sh.appendRow(HEADER);
      sh.setFrozenRows(1);
    }

    // 불참이면 인원 0 으로 기록 (식사 인원 합계 = 인원 열 SUM)
    sh.appendRow([new Date(), side, attend, safe(name), attend === "참석" ? count : 0]);
    return json({ ok: true });
  } catch (err) {
    return json({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

// =, +, -, @ 로 시작하는 입력이 수식으로 실행되지 않게 막음
function safe(s) {
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
