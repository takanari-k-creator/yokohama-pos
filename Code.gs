const ADMIN_KEY = 'ここを好きな合言葉に変える';   // 本物の合言葉は、スプレッドシート側のApps Scriptにだけ書く（GitHubには書かない）

/*
 横浜・まちのコンビニPOSシミュレーター：データの受け取り・保存・読み出し（Google Apps Script）
 ※ もとのコンビニPOSシミュレーターとはデータの形式が違うので、別のスプレッドシートで使ってください。

 index.html とのやりとり
  ・買い物の送信  POST {room, sid, p}          → {ok:true}
  ・データの消去  POST {action:'clear', room, key} → {ok:true}（管理キーが合っているときだけ）
  ・データの読込  GET  ?room=クラス名            → {ok:true, rows:[[sid, 購入コード], ...]}

 スプレッドシートの「data」シートに、1回の買い物を1行として保存します。
 受け付けるのは決まった形式の購入コードだけです（氏名などは扱いません）。
*/

const SHEET_NAME = 'data';
const ROOM_RE = /^[0-9A-Za-z_-]{1,20}$/;
const SID_RE  = /^[0-9a-z]{3}$/;
// 客層(1)性別(1)曜日(1)時(2)分(2)天気(1)気温+10(2)支払(1)来店のきっかけ(1)学校の予定(1) - [商品(36進1桁)個数(1桁)]...
const CODE_RE = /^[0-5][01][0-6]\d\d\d\d[0-2]\d\d[0-4][0-3][0-3]-(?:[0-9a-m][1-9]){1,23}$/;

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(['受信日時', 'クラス(room)', '生徒ID(sid)', '購入コード']);
    sh.setFrozenRows(1);
    sh.getRange('B:D').setNumberFormat('@');   // 「012」などの生徒IDが数値に変換されないよう、文字列として保存する
  }
  return sh;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// データの読み出し：GET ?room=クラス名
function doGet(e) {
  const room = String((e && e.parameter && e.parameter.room) || '');
  if (!ROOM_RE.test(room)) return json_({ ok: false, error: 'bad room' });
  const sh = sheet_();
  const last = sh.getLastRow();
  const rows = [];
  if (last >= 2) {
    sh.getRange(2, 2, last - 1, 3).getValues().forEach(r => {
      if (String(r[0]) === room) rows.push([String(r[1]), String(r[2])]);
    });
  }
  return json_({ ok: true, rows: rows });
}

// 買い物の送信・データの消去：POST
function doPost(e) {
  let req;
  try { req = JSON.parse(e.postData.contents); } catch (err) { return json_({ ok: false, error: 'bad json' }); }
  const room = String(req.room || '');
  if (!ROOM_RE.test(room)) return json_({ ok: false, error: 'bad room' });

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);   // 同時に送られても行が混ざらないように順番待ちする
  try {
    const sh = sheet_();

    // データの消去（管理キーが必要）
    if (req.action === 'clear') {
      if (String(req.key || '') !== ADMIN_KEY) return json_({ ok: false, error: 'bad key' });
      const last = sh.getLastRow();
      if (last >= 2) {
        const rooms = sh.getRange(2, 2, last - 1, 1).getValues();
        for (let i = rooms.length - 1; i >= 0; i--) {   // 下の行から消すと行番号がずれない
          if (String(rooms[i][0]) === room) sh.deleteRow(i + 2);
        }
      }
      return json_({ ok: true });
    }

    // 買い物の送信
    const sid = String(req.sid || '');
    const code = String(req.p || '').toLowerCase();
    if (!SID_RE.test(sid)) return json_({ ok: false, error: 'bad sid' });
    if (!CODE_RE.test(code)) return json_({ ok: false, error: 'bad code' });
    sh.appendRow([new Date(), room, sid, code]);
    return json_({ ok: true });
  } finally {
    lock.releaseLock();
  }
}
