/**
 * 옥계초 AI 수학 연수 — 실습별 작품 제출 저장소
 * 구글 시트에 붙어 있는 Apps Script입니다.
 * 배포: 배포 → 새 배포 → 유형 '웹 앱' → 실행: 나 / 액세스 권한: 모든 사용자
 */
const SHEET = 'posts';
const HEAD = ['id', 'createdAt', 'prac', 'name', 'url', 'token', 'title'];
const PRACS = ['v1', 'v2', 'v3', 'v4', 'i1', 'i2', 'n1', 'n2'];

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET);
  if (!sh) {
    sh = ss.insertSheet(SHEET);
    sh.appendRow(HEAD);
    sh.setFrozenRows(1);
  }
  if (sh.getRange(1, 7).getValue() !== 'title') sh.getRange(1, 7).setValue('title');
  return sh;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// 목록 읽기: 최신 글이 위로, 토큰은 내보내지 않음
function doGet() {
  const rows = sheet_().getDataRange().getValues().slice(1);
  const posts = rows
    .filter(r => r[0])
    .map(r => ({ id: String(r[0]), createdAt: Number(r[1]), prac: String(r[2]), name: String(r[3]), url: String(r[4]), title: String(r[6] || '') }))
    .sort((a, b) => b.createdAt - a.createdAt)
    .slice(0, 500);
  return json_({ ok: true, posts });
}

// 작품 제출 / 내 글 지우기
function doPost(e) {
  let body;
  try { body = JSON.parse(e.postData.contents); } catch (err) { return json_({ ok: false, error: 'bad_json' }); }
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = sheet_();
    if (body.action === 'add') {
      const name = String(body.name || '').trim().slice(0, 20);
      const title = String(body.title || '').trim().slice(0, 60);
      const url = String(body.url || '').trim().slice(0, 500);
      const prac = String(body.prac || '');
      const token = String(body.token || '').slice(0, 64);
      if (!name || !/^https?:\/\/\S+\.\S+/i.test(url) || PRACS.indexOf(prac) < 0 || !token) {
        return json_({ ok: false, error: 'invalid' });
      }
      const id = Utilities.getUuid().slice(0, 8);
      // 수식으로 해석되지 않도록 앞에 ' 를 붙여 저장
      sh.appendRow([id, Date.now(), prac, "'" + name, "'" + url, token, title ? "'" + title : '']);
      return json_({ ok: true, id });
    }
    if (body.action === 'delete') {
      const data = sh.getDataRange().getValues();
      for (let i = data.length - 1; i >= 1; i--) {
        if (String(data[i][0]) === String(body.id) && String(data[i][5]) === String(body.token)) {
          sh.deleteRow(i + 1);
          return json_({ ok: true });
        }
      }
      return json_({ ok: false, error: 'not_found' });
    }
    return json_({ ok: false, error: 'unknown_action' });
  } finally {
    lock.releaseLock();
  }
}
