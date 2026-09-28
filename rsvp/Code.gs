/**
 * 홈커밍 초대장 — 참석 여부 응답 수집용 Google Apps Script
 *
 * ▣ 설정 방법 (5분)
 * 1. Google 드라이브에서 새 스프레드시트를 만듭니다. (예: "홈커밍 참석 응답")
 * 2. 메뉴 [확장 프로그램] → [Apps Script] 를 엽니다.
 * 3. 기본 Code.gs 내용을 모두 지우고 이 파일 내용을 붙여넣은 뒤 저장합니다.
 * 4. 오른쪽 위 [배포] → [새 배포] → 유형 선택(톱니바퀴) → [웹 앱]
 *      - 다음 사용자 인증 정보로 실행: 나
 *      - 액세스 권한이 있는 사용자: 모든 사용자
 *    [배포] 를 누르고 권한을 허용합니다.
 * 5. 발급된 "웹 앱 URL"(https://script.google.com/macros/s/.../exec)을 복사해
 *    index.html 의 CONFIG.rsvpEndpoint 에 붙여넣습니다.
 *
 * 응답은 스프레드시트의 "참석응답" 시트에 한 줄씩 쌓입니다.
 * 관리자는 이 스프레드시트를 열어 확인·필터·엑셀 다운로드할 수 있습니다.
 * (스프레드시트 공유 설정은 관리자만 볼 수 있게 유지하세요.)
 *
 * 코드를 수정한 경우 [배포] → [배포 관리] → 수정(연필) → 버전 "새 버전" 으로 다시 배포해야 반영됩니다.
 */

const SHEET_NAME = '참석응답';
const HEADERS = ['제출시각', '참석여부', '성명', '학번', '소속', '연락처', '이메일', '기기정보'];

// 새 응답이 올 때마다 메일로 알림을 받으려면 주소를 넣으세요. (비우면 알림 없음)
const NOTIFY_EMAIL = '';

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const d = JSON.parse(e.postData.contents);

    const required = ['attend', 'name', 'year', 'org', 'phone', 'email'];
    for (const k of required) {
      if (!d[k] || String(d[k]).trim() === '') return json({ ok: false, error: 'missing_' + k });
    }
    if (['참석', '불참'].indexOf(d.attend) < 0) return json({ ok: false, error: 'bad_attend' });
    if (!/^\S+@\S+\.\S+$/.test(d.email)) return json({ ok: false, error: 'bad_email' });

    const sheet = getSheet();
    // 수식 주입 방지: = + - @ 로 시작하면 앞에 ' 를 붙임
    const clean = v => {
      const s = String(v == null ? '' : v).slice(0, 200);
      return /^[=+\-@]/.test(s) ? "'" + s : s;
    };
    sheet.appendRow([
      new Date(), clean(d.attend), clean(d.name), clean(d.year) + '학번',
      clean(d.org), "'" + String(d.phone).slice(0, 20), clean(d.email), clean(d.userAgent),
    ]);

    if (NOTIFY_EMAIL) {
      MailApp.sendEmail(NOTIFY_EMAIL, `[홈커밍] ${d.name}님 ${d.attend} 응답`,
        `${d.name} (${d.year}학번 / ${d.org})\n${d.attend}\n${d.phone} / ${d.email}`);
    }
    return json({ ok: true });
  } catch (err) {
    return json({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

// 브라우저에서 웹 앱 URL을 직접 열었을 때 동작 확인용
function doGet() {
  return json({ ok: true, message: 'RSVP endpoint is running' });
}

function getSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(HEADERS);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold').setBackground('#efe2d6');
  }
  return sheet;
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
