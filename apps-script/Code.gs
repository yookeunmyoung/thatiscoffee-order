// thatiscoffee 단체 주문 모으기 (구글 시트 확장 프로그램 > Apps Script에 붙여 넣기)
// 배포: 배포 > 새 배포 > 웹 앱, 실행 계정 "나", 액세스 권한 "모든 사용자"
const SHEET = '주문', ARCHIVE = '지난주문';

function sheet_(name) {
  const ss = SpreadsheetApp.getActive();
  let s = ss.getSheetByName(name);
  if (!s) {
    s = ss.insertSheet(name);
    s.appendRow(['시간', '이름', '주문', '항목', '금액', '요청']);
  }
  return s;
}

// 시트에서 수식으로 해석되지 않도록 앞의 = + - @ 제거
function clean_(v, max) {
  return String(v == null ? '' : v).replace(/^[=+\-@\s]+/, '').slice(0, max);
}

function out_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function list_() {
  return sheet_(SHEET).getDataRange().getValues().slice(1).map(r => {
    let items = [];
    try { items = JSON.parse(r[3] || '[]'); } catch (e) {}
    return { time: r[0], name: String(r[1]), text: String(r[2]), items, sum: Number(r[4]) || 0, memo: String(r[5]) };
  });
}

function doGet() {
  return out_({ ok: true, orders: list_() });
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const d = JSON.parse(e.postData.contents);
    const s = sheet_(SHEET);

    // 새 모임 시작: 지금 주문을 '지난주문' 시트로 옮기고 비움
    if (d.action === 'reset') {
      const rows = s.getDataRange().getValues().slice(1);
      if (rows.length) {
        const a = sheet_(ARCHIVE);
        a.getRange(a.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
        s.deleteRows(2, rows.length);
      }
      return out_({ ok: true, orders: [] });
    }

    const name = clean_(d.name, 30).trim();
    if (!name) return out_({ ok: false, error: '이름을 입력해 주세요.' });

    // 같은 이름의 이전 주문은 지우고 새 주문으로 바꿈
    const v = s.getDataRange().getValues();
    for (let i = v.length - 1; i >= 1; i--) if (String(v[i][1]) === name) s.deleteRow(i + 1);

    if (d.action !== 'cancel') {
      const items = (Array.isArray(d.items) ? d.items : []).slice(0, 80)
        .map(x => [clean_(x[0], 60), Number(x[1]) || 0, Number(x[2]) || 0]);
      s.appendRow([new Date(), name, clean_(d.text, 2000), JSON.stringify(items), Number(d.sum) || 0, clean_(d.memo, 300)]);
    }
    return out_({ ok: true, orders: list_() });
  } catch (err) {
    return out_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}
