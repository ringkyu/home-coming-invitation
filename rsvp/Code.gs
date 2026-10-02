/**
 * 페르시아어·이란학과 창립 50주년 홈커밍데이 — 참석 여부 응답 수집용 Google Apps Script
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
 * 응답은 스프레드시트의 "50주년 참석응답" 시트에 한 줄씩 쌓입니다.
 * 관리자는 이 스프레드시트를 열어 확인·필터·엑셀 다운로드할 수 있습니다.
 * (스프레드시트 공유 설정은 관리자만 볼 수 있게 유지하세요.)
 *
 * 초대장의 "축하 메시지 남기기"(방명록)는 "방명록" 시트에 쌓이고 초대장에 공개됩니다.
 * 부적절한 메시지는 해당 줄의 "숨김" 칸에 아무 값이나 입력하면 초대장에서 사라집니다.
 *
 * 코드를 수정한 경우 [배포] → [배포 관리] → 수정(연필) → 버전 "새 버전" 으로 다시 배포해야 반영됩니다.
 */

const SHEET_NAME = '50주년 참석응답';
const HEADERS = ['제출시각', '참석여부', '성명', '학번', '연락처', '후리스사이즈', '기기정보'];
const SIZES = ['S', 'M', 'L', 'XL', '2XL', '3XL'];

const GB_SHEET_NAME = '방명록';
const GB_HEADERS = ['작성시각', '이름', '메시지', '숨김'];

// 새 응답이 올 때마다 메일로 알림을 받으려면 주소를 넣으세요. (비우면 알림 없음)
const NOTIFY_EMAIL = '';

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const d = JSON.parse(e.postData.contents);
    if (d.type === 'guestbook') return addGuestbook(d);

    const required = ['attend', 'name', 'year', 'phone'];
    for (const k of required) {
      if (!d[k] || String(d[k]).trim() === '') return json({ ok: false, error: 'missing_' + k });
    }
    if (['참석', '불참'].indexOf(d.attend) < 0) return json({ ok: false, error: 'bad_attend' });
    if (d.attend === '참석' && SIZES.indexOf(d.size) < 0) return json({ ok: false, error: 'bad_size' });

    const sheet = getSheet(SHEET_NAME, HEADERS);
    sheet.appendRow([
      new Date(), clean(d.attend), clean(d.name), clean(d.year) + '학번',
      "'" + String(d.phone).slice(0, 20), d.attend === '참석' ? clean(d.size) : '', clean(d.userAgent),
    ]);

    if (NOTIFY_EMAIL) {
      MailApp.sendEmail(NOTIFY_EMAIL, `[50주년 홈커밍] ${d.name}님 ${d.attend} 응답`,
        `${d.name} (${d.year}학번)\n${d.attend}${d.size ? ' / 후리스 ' + d.size : ''}\n${d.phone}`);
    }
    return json({ ok: true });
  } catch (err) {
    return json({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

// ?action=guestbook : 방명록 목록 (최신순, 숨김 제외)
// 그 외 : 브라우저에서 웹 앱 URL을 직접 열었을 때 동작 확인용
function doGet(e) {
  if (e && e.parameter && e.parameter.action === 'guestbook') {
    return json({ ok: true, entries: listGuestbook() });
  }
  return json({ ok: true, message: 'RSVP endpoint is running' });
}

// 방명록 저장 (doPost 의 잠금 안에서 호출됨)
function addGuestbook(d) {
  const name = String(d.name || '').trim(), message = String(d.message || '').trim();
  if (!name) return json({ ok: false, error: 'missing_name' });
  if (!message) return json({ ok: false, error: 'missing_message' });
  getSheet(GB_SHEET_NAME, GB_HEADERS).appendRow([new Date(), clean(name, 20), clean(message, 300), '']);
  return json({ ok: true });
}

function listGuestbook() {
  const sheet = getSheet(GB_SHEET_NAME, GB_HEADERS);
  return sheet.getDataRange().getValues().slice(1)
    .filter(r => r[1] !== '' && r[2] !== '' && r[3] === '')
    .reverse()
    .slice(0, 300)
    .map(r => ({
      time: r[0] instanceof Date ? r[0].toISOString() : String(r[0]),
      name: String(r[1]),
      message: String(r[2]),
    }));
}

// 수식 주입 방지: = + - @ 로 시작하면 앞에 ' 를 붙임
function clean(v, max) {
  const s = String(v == null ? '' : v).slice(0, max || 200);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

function getSheet(name, headers) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(headers);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#efe2d6');
  }
  return sheet;
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
