// ============================================================
//  経費ノート — Google Apps Script バックエンド v3.1.1（複式簿記・領収書の写真に対応）
//
//  ■ 既存データについて
//    ・「経費」「開業費」「収入」「減価償却」シートの既存の行・列は一切消しません。
//    ・足りない列は「右端に追加」するだけです（既存列の並びは変えません）。
//    ・「シート1」「収入・補助金」「請求書」など、このアプリが使わないシートには触れません。
//    ・新しく「仕訳帳」「設定」「変更履歴」「領収書」シートを自動で作ります。
//    ・領収書の写真は Googleドライブの「経費ノート_領収書」フォルダ（年ごと）に保存します。
//
//  ■ 領収書の文字認識（OCR）を使うには
//    Apps Script の左メニュー「サービス ＋」→「Drive API」（v3）を追加してください。
//    追加しなくても写真の保存はできます（文字の自動読み取りだけが使えません）。
//
//  ■ 使い方
//    このファイルの中身をすべてコピーして Apps Script に貼り付け、
//    「デプロイを管理 → 鉛筆アイコン → バージョン：新バージョン → デプロイ」で更新してください。
//    （「新しいデプロイ」にするとURLが変わるので注意）
// ============================================================

var APP_VERSION    = '3.1.1';
var SPREADSHEET_ID = '1GYGiuBr2t87Nvp1nUGsgxgVAJCmzmPmOMnr2-1BAcGY'; // 経費ノート
var TZ             = 'Asia/Tokyo';
var BACKUP_FOLDER  = '経費ノート_バックアップ';
var RECEIPT_FOLDER = '経費ノート_領収書';

// ── シートの列定義（左から順番。既存シートは見出し名で列を探すので、並びが違っても動きます） ──
var SCHEMAS = {
  records: {
    sheet: '経費', color: '#2d6a4f',
    cols: [
      ['id','ID'], ['date','日付'], ['category','勘定科目'], ['memo','内容'], ['payment','支払方法'],
      ['amount','金額'], ['ratio','按分割合%'], ['adjusted','按分後金額'], ['note','メモ'],
      ['kubun','所得区分'], ['updatedAt','更新日時'], ['receipt','領収書ID']
    ]
  },
  kaigyohi: {
    sheet: '開業費', color: '#e07a2f',
    cols: [
      ['id','ID'], ['date','日付'], ['category','カテゴリ'], ['memo','内容'], ['amount','金額'],
      ['note','メモ'], ['moved','経費移動済み'], ['payment','支払方法'], ['updatedAt','更新日時'], ['receipt','領収書ID']
    ]
  },
  incomes: {
    sheet: '収入', color: '#1e40af',
    cols: [
      ['id','ID'], ['date','日付'], ['type','種類'], ['memo','内容'], ['amount','金額'], ['note','メモ'],
      ['receive','入金先'], ['kubun','所得区分'], ['updatedAt','更新日時'], ['receipt','領収書ID']
    ]
  },
  deprec: {
    sheet: '減価償却', color: '#6d28d9',
    cols: [
      ['id','ID'], ['date','取得日'], ['name','品名'], ['cost','取得価額'], ['years','耐用年数'],
      ['method','償却方法'], ['ratio','事業割合%'], ['annual','年間償却額'], ['note','メモ'],
      ['startDate','事業供用日'], ['assetType','資産区分'], ['account','勘定科目'],
      ['openingBook','開業時帳簿価額'], ['payment','支払方法'], ['kubun','所得区分'],
      ['disposedDate','除却日'], ['manualAnnual','手入力償却額'], ['updatedAt','更新日時'], ['receipt','領収書ID']
    ]
  },
  journals: {
    sheet: '仕訳帳', color: '#0f766e',
    cols: [
      ['id','ID'], ['slip','伝票番号'], ['date','日付'], ['debit','借方科目'], ['credit','貸方科目'],
      ['amount','金額'], ['memo','摘要'], ['kubun','所得区分'], ['kind','種別'], ['note','メモ'],
      ['updatedAt','更新日時']
    ]
  },
  receipts: {
    sheet: '領収書', color: '#92400e',
    cols: [
      ['id','ID'], ['takenAt','撮影日時'], ['fileName','ファイル名'], ['url','URL'], ['fileId','ファイルID'],
      ['status','状態'], ['linkKind','取引の種類'], ['linkId','取引ID'], ['ocrText','読み取った文字'],
      ['note','メモ'], ['thumb','サムネイル'], ['updatedAt','更新日時']
    ]
  }
};
var RECEIPT_KINDS = { records:1, kaigyohi:1, incomes:1, deprec:1 };
var NUMBER_KEYS = { amount:1, ratio:1, adjusted:1, cost:1, years:1, annual:1, openingBook:1, manualAnnual:1 };
var DATE_KEYS   = { date:1, startDate:1, disposedDate:1 };
var BOOL_KEYS   = { moved:1 };
var DATETIME_KEYS = { updatedAt:1, takenAt:1 };

var SETTINGS_SHEET = '設定';
var LOG_SHEET      = '変更履歴';

// ============================================================
//  入口
// ============================================================
function doPost(e) {
  var result;
  try {
    var data = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    checkKey_(data.key);
    result = handle_(data.action, data);
  } catch (err) {
    result = { ok: false, error: String(err && err.message ? err.message : err) };
  }
  return json_(result);
}

function doGet(e) {
  var p = (e && e.parameter) || {};
  var result;
  try {
    checkKey_(p.key);
    if (p.action === 'ping')      result = ping_();
    else if (p.action === 'init') { initSheets(); result = { ok: true, message: 'シートを準備しました（既存データはそのままです）' }; }
    else                          result = getAllData();
  } catch (err) {
    result = { ok: false, error: String(err && err.message ? err.message : err) };
  }
  return json_(result);
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// 合言葉（任意）。スクリプトプロパティ APP_KEY を設定したときだけ有効になります。
function checkKey_(key) {
  var required = PropertiesService.getScriptProperties().getProperty('APP_KEY');
  if (required && String(key || '') !== required) throw new Error('合言葉（アプリキー）が違います。アプリの設定画面で確認してください。');
}

function handle_(action, d) {
  switch (action) {
    case 'ping':            return ping_();
    case 'init':            initSheets(); return { ok: true };
    case 'getAll':          return getAllData();

    // ── 経費（旧バージョンと同じ名前も使えます） ──
    case 'addRecord':       return addRow_('records', d.record);
    case 'updateRecord':    return updateRow_('records', d.record);
    case 'delRecord':       return deleteRow_('records', d.id);
    // ── 開業費 ──
    case 'addKaigyohi':     return addRow_('kaigyohi', d.record);
    case 'updateKaigyohi':  return updateRow_('kaigyohi', d.record);
    case 'delKaigyohi':     return deleteRow_('kaigyohi', d.id);
    case 'moveToRecord':    return moveKaigyohiToRecord_(d.id, d.category, d.date);
    case 'moveToKaigyohi':  return moveRecordToKaigyohi_(d.id);
    // ── 収入 ──
    case 'addIncome':       return addRow_('incomes', d.record);
    case 'updateIncome':    return updateRow_('incomes', d.record);
    case 'delIncome':       return deleteRow_('incomes', d.id);
    // ── 固定資産（減価償却） ──
    case 'addDeprec':       return addRow_('deprec', d.record);
    case 'updateDeprec':    return updateRow_('deprec', d.record);
    case 'delDeprec':       return deleteRow_('deprec', d.id);
    case 'recordToAsset':   return recordToAsset_(d.id, d.asset);
    // ── 振替伝票（手動の仕訳） ──
    case 'addJournal':      return addRow_('journals', d.record);
    case 'updateJournal':   return updateRow_('journals', d.record);
    case 'delJournal':      return deleteRow_('journals', d.id);
    case 'saveSlip':        return saveSlip_(d.slip, d.lines);
    case 'delSlip':         return deleteSlip_(d.slip);
    // ── まとめて登録（定期取引など） ──
    case 'batchAdd':        return batchAdd_(d.items);
    // ── 設定 ──
    case 'getSettings':     return { ok: true, settings: readSettings_() };
    case 'saveSettings':    return saveSettings_(d.settings);
    // ── 書き出し・バックアップ ──
    case 'writeReport':     return writeReport_(d.name, d.rows);
    case 'backup':          return backupNow();
    case 'setupAutoBackup': return setupAutoBackup();
    case 'getLog':          return getLog_(d.limit);
    // ── 領収書の写真 ──
    case 'uploadReceipt':   return uploadReceipt_(d.receipt, d.image, d.mime);
    case 'getReceiptImage': return getReceiptImage_(d.id);
    case 'updateReceipt':   return updateReceipt_(d.receipt);
    case 'delReceipt':      return deleteReceipt_(d.id);
    case 'reocrReceipt':    return reocrReceipt_(d.id);
    // ── 為替 ──
    case 'getUSDRate':      return getUSDRate(d.date);
    default:                return { ok: false, error: 'unknown action: ' + action };
  }
}

function ping_() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  return { ok: true, version: APP_VERSION, spreadsheet: ss.getName(), serverTime: Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss'),
           ocr: (typeof Drive !== 'undefined' && !!Drive.Files) };
}

// ============================================================
//  シートの準備（何度実行しても安全。データは消しません）
// ============================================================
function initSheets() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  for (var kind in SCHEMAS) ensureSheet_(ss, kind);
  ensureKVSheet_(ss, SETTINGS_SHEET, ['キー', '値（JSON）'], '#374151');
  ensureKVSheet_(ss, LOG_SHEET, ['日時', '操作', 'シート', 'ID', '変更前', '変更後'], '#6b7280');
  return { ok: true };
}

function ensureKVSheet_(ss, name, headers, color) {
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold').setBackground(color).setFontColor('#ffffff');
    sh.setFrozenRows(1);
  }
  return sh;
}

// シートと列を用意し、{key: 列番号(1始まり)} を返す
function ensureSheet_(ss, kind) {
  var schema = SCHEMAS[kind];
  var sh = ss.getSheetByName(schema.sheet);
  var headers = schema.cols.map(function (c) { return c[1]; });
  if (!sh) {
    sh = ss.insertSheet(schema.sheet);
  }
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers])
      .setFontWeight('bold').setBackground(schema.color).setFontColor('#ffffff');
    sh.setFrozenRows(1);
  }
  var lastCol = Math.max(sh.getLastColumn(), 1);
  var row1 = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(function (v) { return String(v).trim(); });
  var map = {};
  schema.cols.forEach(function (c, i) {
    var key = c[0], name = c[1];
    var idx = row1.indexOf(name);
    if (idx >= 0) { map[key] = idx + 1; return; }
    // 本来の位置の見出しが空欄なら、そこに見出しだけ書く（例：開業費のG列）
    if (i < row1.length && row1[i] === '' && !isUsedIndex_(map, i + 1)) {
      sh.getRange(1, i + 1).setValue(name).setFontWeight('bold');
      row1[i] = name;
      map[key] = i + 1;
      return;
    }
    // 見つからなければ右端に列を追加（既存の列は動かさない）
    var newCol = row1.length + 1;
    sh.getRange(1, newCol).setValue(name).setFontWeight('bold').setBackground(schema.color).setFontColor('#ffffff');
    row1.push(name);
    map[key] = newCol;
  });
  return { sheet: sh, map: map, width: row1.length };
}

function isUsedIndex_(map, col) {
  for (var k in map) if (map[k] === col) return true;
  return false;
}

// ============================================================
//  読み込み
// ============================================================
function getAllData() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  initSheets();
  var out = { ok: true, version: APP_VERSION, serverTime: Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss') };
  for (var kind in SCHEMAS) out[kind] = readAll_(ss, kind);
  // 取引に紐づいた領収書は、通信量を減らすためサムネイルを送らない
  out.receipts.forEach(function (r) { if (r.status === '登録済') r.thumb = ''; });
  out.ocr = (typeof Drive !== 'undefined' && !!Drive.Files);
  out.settings = readSettings_();
  return out;
}

function readAll_(ss, kind) {
  var t = ensureSheet_(ss, kind);
  var sh = t.sheet;
  var last = sh.getLastRow();
  if (last < 2) return [];
  var values = sh.getRange(2, 1, last - 1, t.width).getValues();
  var list = [];
  values.forEach(function (row) {
    var o = rowToObj_(row, t.map);
    if (o.id) list.push(o);
  });
  return list;
}

function rowToObj_(row, map) {
  var o = {};
  for (var key in map) {
    var v = row[map[key] - 1];
    if (DATE_KEYS[key])        o[key] = normDate_(v);
    else if (NUMBER_KEYS[key]) o[key] = (v === '' || v === null) ? '' : toNum(v);
    else if (BOOL_KEYS[key])   o[key] = (v === true || String(v).toUpperCase() === 'TRUE');
    else if (DATETIME_KEYS[key]) o[key] = (v instanceof Date) ? Utilities.formatDate(v, TZ, 'yyyy-MM-dd HH:mm:ss') : String(v || '');
    else                       o[key] = (v === null || v === undefined) ? '' : String(v);
  }
  if (o.id) o.id = String(o.id).replace(/\.0$/, '');
  return o;
}

// 日付を yyyy-MM-dd にそろえる（Date型・「2026/06/11」・「2026-6-1」に対応）
function normDate_(v) {
  if (!v) return '';
  if (v instanceof Date) return Utilities.formatDate(v, TZ, 'yyyy-MM-dd');
  var s = String(v).trim();
  var m = s.match(/^(\d{4})[\/\-.](\d{1,2})[\/\-.](\d{1,2})/);
  if (m) return m[1] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[3]).slice(-2);
  return s;
}

function toNum(v) {
  if (v instanceof Date) return 0;
  if (typeof v === 'number') return v;
  var n = parseFloat(String(v).replace(/[^0-9.\-]/g, ''));
  return isNaN(n) ? 0 : n;
}

// ============================================================
//  書き込み（追加・更新・削除）— すべてロックして同時書き込みを防ぎ、変更履歴を残す
// ============================================================
function withLock_(fn) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try { return fn(); }
  finally { lock.releaseLock(); }
}

function addRow_(kind, obj) {
  if (!obj) return { ok: false, error: 'データがありません' };
  return withLock_(function () {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var r = appendObj_(ss, kind, obj);
    if (!r.ok) return r;
    log_(ss, '追加', SCHEMAS[kind].sheet, r.id, null, r.saved);
    return { ok: true, id: r.id, newId: r.id };
  });
}

function appendObj_(ss, kind, obj) {
  var t = ensureSheet_(ss, kind);
  var o = sanitize_(kind, obj);
  if (!o.id) o.id = newId_();
  var lockErr = checkLocked_(o);
  if (lockErr) return { ok: false, error: lockErr };
  if (findRow_(t, o.id) > 0) return { ok: false, error: '同じIDのデータがすでにあります（二重登録防止）: ' + o.id };
  o.updatedAt = now_();
  var row = new Array(t.width);
  for (var i = 0; i < t.width; i++) row[i] = '';
  for (var key in t.map) if (o.hasOwnProperty(key)) row[t.map[key] - 1] = o[key];
  t.sheet.appendRow(row);
  if (RECEIPT_KINDS[kind] && o.receipt) linkReceipt_(ss, o.receipt, kind, o);
  return { ok: true, id: o.id, saved: o };
}

function updateRow_(kind, obj) {
  if (!obj || !obj.id) return { ok: false, error: 'IDがありません' };
  return withLock_(function () {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var t = ensureSheet_(ss, kind);
    var rowNo = findRow_(t, obj.id);
    if (rowNo < 0) return { ok: false, error: '更新するデータが見つかりません（ID: ' + obj.id + '）' };
    var range = t.sheet.getRange(rowNo, 1, 1, t.width);
    var row = range.getValues()[0];
    var before = rowToObj_(row, t.map);
    var o = sanitize_(kind, obj);
    var lockErr = checkLocked_(before) || checkLocked_(o);
    if (lockErr) return { ok: false, error: lockErr };
    o.updatedAt = now_();
    // 渡された項目だけ上書き（知らない列・渡されていない項目はそのまま残す）
    for (var key in t.map) if (o.hasOwnProperty(key) && key !== 'id') row[t.map[key] - 1] = o[key];
    range.setValues([row]);
    var after = rowToObj_(row, t.map);
    if (RECEIPT_KINDS[kind] && o.hasOwnProperty('receipt') && String(before.receipt || '') !== String(o.receipt || '')) {
      if (before.receipt) unlinkReceipt_(ss, before.receipt, obj.id);
      if (o.receipt) linkReceipt_(ss, o.receipt, kind, after);
    } else if (RECEIPT_KINDS[kind] && after.receipt) {
      linkReceipt_(ss, after.receipt, kind, after); // 金額や内容が変わったらファイル名も更新
    }
    log_(ss, '修正', SCHEMAS[kind].sheet, obj.id, before, after);
    return { ok: true, id: String(obj.id) };
  });
}

function deleteRow_(kind, id) {
  if (!id) return { ok: false, error: 'IDがありません' };
  return withLock_(function () {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var r = deleteById_(ss, kind, id);
    return r;
  });
}

function deleteById_(ss, kind, id) {
  var t = ensureSheet_(ss, kind);
  var rowNo = findRow_(t, id);
  if (rowNo < 0) return { ok: true, notFound: true };
  var before = rowToObj_(t.sheet.getRange(rowNo, 1, 1, t.width).getValues()[0], t.map);
  var lockErr = checkLocked_(before);
  if (lockErr) return { ok: false, error: lockErr };
  t.sheet.deleteRow(rowNo);
  if (RECEIPT_KINDS[kind] && before.receipt) unlinkReceipt_(ss, before.receipt, id); // 写真は消さずに「未登録」に戻す
  log_(ss, '削除', SCHEMAS[kind].sheet, id, before, null);
  return { ok: true };
}

function findRow_(t, id) {
  var last = t.sheet.getLastRow();
  if (last < 2) return -1;
  var ids = t.sheet.getRange(2, t.map.id, last - 1, 1).getValues();
  var target = String(id).replace(/\.0$/, '');
  for (var i = 0; i < ids.length; i++) {
    if (String(ids[i][0]).replace(/\.0$/, '') === target) return i + 2;
  }
  return -1;
}

function sanitize_(kind, obj) {
  var o = {};
  SCHEMAS[kind].cols.forEach(function (c) {
    var key = c[0];
    if (!obj.hasOwnProperty(key) || key === 'updatedAt') return;
    var v = obj[key];
    if (DATE_KEYS[key])        o[key] = normDate_(v);
    else if (NUMBER_KEYS[key]) o[key] = (v === '' || v === null || v === undefined) ? '' : toNum(v);
    else if (BOOL_KEYS[key])   o[key] = (v === true || String(v).toUpperCase() === 'TRUE');
    else                       o[key] = safeText_(v);
  });
  if (o.id) o.id = String(o.id);
  return o;
}

// 「=」「+」「-」「@」で始まる文字はスプレッドシートが数式と解釈するので、先頭に ' を付けて文字として保存する
function safeText_(v) {
  if (v === null || v === undefined) return '';
  var s = String(v);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

function newId_() {
  Utilities.sleep(2); // 同じミリ秒のIDが重ならないように
  return String(new Date().getTime());
}

function now_() { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss'); }

// 締めた年（設定の lockedYears）のデータは変更不可
function checkLocked_(o) {
  if (!o || !o.date) return '';
  var years = (readSettings_().lockedYears || []).map(String);
  var y = String(o.date).slice(0, 4);
  if (years.indexOf(y) >= 0) return y + '年は「締め」済みのため変更できません。設定画面で締めを解除してから修正してください。';
  return '';
}

// ============================================================
//  開業費 ⇔ 経費 の移動
// ============================================================
// 開業費 → 経費（開業費の行は「移動済み」にして残す）
function moveKaigyohiToRecord_(id, category, date) {
  return withLock_(function () {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var kt = ensureSheet_(ss, 'kaigyohi');
    var rowNo = findRow_(kt, id);
    if (rowNo < 0) return { ok: false, error: '開業費が見つかりません' };
    var range = kt.sheet.getRange(rowNo, 1, 1, kt.width);
    var row = range.getValues()[0];
    var k = rowToObj_(row, kt.map);
    if (k.moved) return { ok: false, error: 'この開業費はすでに経費へ移動済みです' };
    var usd = extractUsd_(k.note);
    var pay = k.payment || extractPay_(k.note) || 'クレジットカード';
    var rec = {
      id: newId_(), date: date || k.date, category: category || k.category,
      memo: k.memo, payment: pay, amount: k.amount, ratio: 100, adjusted: k.amount,
      note: (usd ? usd + ' / ' : '') + '開業費より移動 (ID:' + k.id + ')' + (k.note && !usd ? ' / ' + cleanTags_(k.note) : ''),
      receipt: k.receipt || ''
    };
    var r = appendObj_(ss, 'records', rec);
    if (!r.ok) return r;
    row[kt.map.moved - 1] = true;
    row[kt.map.updatedAt - 1] = now_();
    range.setValues([row]);
    log_(ss, '移動(開業費→経費)', '開業費', k.id, k, rec);
    return { ok: true, newId: r.id };
  });
}

// 経費 → 開業費（開業費から来たものなら元の開業費に戻す）
function moveRecordToKaigyohi_(id) {
  return withLock_(function () {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var rt = ensureSheet_(ss, 'records');
    var rowNo = findRow_(rt, id);
    if (rowNo < 0) return { ok: false, error: '経費が見つかりません' };
    var r = rowToObj_(rt.sheet.getRange(rowNo, 1, 1, rt.width).getValues()[0], rt.map);
    var lockErr = checkLocked_(r);
    if (lockErr) return { ok: false, error: lockErr };
    var m = String(r.note || '').match(/\(ID:(\d+)\)/);
    if (m) {
      var kt = ensureSheet_(ss, 'kaigyohi');
      var kRow = findRow_(kt, m[1]);
      if (kRow > 0) {
        kt.sheet.getRange(kRow, kt.map.moved).setValue(false);
        kt.sheet.getRange(kRow, kt.map.updatedAt).setValue(now_());
        var kObj = rowToObj_(kt.sheet.getRange(kRow, 1, 1, kt.width).getValues()[0], kt.map);
        var rcpt = r.receipt || kObj.receipt;
        if (rcpt) { kt.sheet.getRange(kRow, kt.map.receipt).setValue(rcpt); kObj.receipt = rcpt; linkReceipt_(ss, rcpt, 'kaigyohi', kObj); }
        rt.sheet.deleteRow(rowNo);
        log_(ss, '移動(経費→開業費・復元)', '経費', r.id, r, { kaigyohiId: m[1] });
        return { ok: true, restored: true, kaigyohiId: m[1] };
      }
    }
    // 按分している経費は「按分後の金額」を開業費にする（旧版は全額を移していました）
    var amount = toNum(r.adjusted) || toNum(r.amount);
    var usd = extractUsd_(r.note);
    var ratioNote = (toNum(r.ratio) && toNum(r.ratio) < 100) ? ' 按分: ¥' + toNum(r.amount) + ' × ' + toNum(r.ratio) + '%' : '';
    var k = {
      id: newId_(), date: r.date, category: baseAccount_(r.category), memo: r.memo, amount: amount,
      note: '[PAY:' + (r.payment || 'クレジットカード') + ']' + (usd ? ' ' + usd : '') + ratioNote + ' / 経費より移動: ' + r.category,
      moved: false, payment: r.payment || '', receipt: r.receipt || ''
    };
    var a = appendObj_(ss, 'kaigyohi', k);
    if (!a.ok) return a;
    rt.sheet.deleteRow(rowNo);
    log_(ss, '移動(経費→開業費)', '経費', r.id, r, k);
    return { ok: true, restored: false, newKaigyohiId: a.id };
  });
}

// 経費 → 固定資産（10万円以上の買い物を減価償却に回す）
function recordToAsset_(id, asset) {
  return withLock_(function () {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var rt = ensureSheet_(ss, 'records');
    var rowNo = findRow_(rt, id);
    if (rowNo < 0) return { ok: false, error: '経費が見つかりません' };
    var r = rowToObj_(rt.sheet.getRange(rowNo, 1, 1, rt.width).getValues()[0], rt.map);
    var lockErr = checkLocked_(r);
    if (lockErr) return { ok: false, error: lockErr };
    var a = asset || {};
    a.id = a.id || newId_();
    a.date = a.date || r.date;
    a.name = a.name || r.memo;
    a.cost = a.cost || r.amount;
    a.payment = a.payment || r.payment;
    a.receipt = a.receipt || r.receipt || '';
    a.note = (a.note ? a.note + ' / ' : '') + '経費より移動 (元ID:' + r.id + ')';
    var res = appendObj_(ss, 'deprec', a);
    if (!res.ok) return res;
    rt.sheet.deleteRow(rowNo);
    log_(ss, '移動(経費→固定資産)', '経費', r.id, r, a);
    return { ok: true, newId: res.id };
  });
}

function baseAccount_(cat) {
  var s = String(cat || '');
  var special = { '地代家賃_駐車場':'地代家賃', '損害保険料_車':'損害保険料', '旅費交通費_車':'旅費交通費', '開業費→経費':'その他開業費' };
  if (special[s]) return special[s];
  return s.split('_')[0] || 'その他開業費';
}
function extractUsd_(note) {
  var m = String(note || '').match(/USD: \$[0-9.]+ × TTM: [0-9.]+円 = ¥[0-9,]+（[^）]+）/);
  return m ? m[0] : '';
}
function extractPay_(note) {
  var m = String(note || '').match(/\[PAY:([^\]]+)\]/);
  return m ? m[1] : '';
}
function cleanTags_(note) {
  return String(note || '').replace(/\[PAY:[^\]]+\]\s*/, '').trim();
}

// ============================================================
//  振替伝票（複数行の仕訳を同じ伝票番号でまとめて保存）
// ============================================================
function saveSlip_(slip, lines) {
  if (!lines || !lines.length) return { ok: false, error: '仕訳の行がありません' };
  return withLock_(function () {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var t = ensureSheet_(ss, 'journals');
    var slipNo = slip || ('S' + newId_());
    // 既存の同じ伝票の行を消してから書き直す（修正時）
    var removed = [];
    var last = t.sheet.getLastRow();
    if (slip && last >= 2) {
      var vals = t.sheet.getRange(2, 1, last - 1, t.width).getValues();
      for (var i = vals.length - 1; i >= 0; i--) {
        var o = rowToObj_(vals[i], t.map);
        if (String(o.slip) === String(slip)) {
          var le = checkLocked_(o);
          if (le) return { ok: false, error: le };
          removed.push(o);
          t.sheet.deleteRow(i + 2);
        }
      }
    }
    var saved = [];
    for (var j = 0; j < lines.length; j++) {
      var ln = lines[j];
      ln.slip = slipNo;
      ln.id = newId_();
      var r = appendObj_(ss, 'journals', ln);
      if (!r.ok) return r;
      saved.push(r.saved);
    }
    log_(ss, slip ? '振替伝票修正' : '振替伝票追加', '仕訳帳', slipNo, removed.length ? removed : null, saved);
    return { ok: true, slip: slipNo, lines: saved };
  });
}

function deleteSlip_(slip) {
  if (!slip) return { ok: false, error: '伝票番号がありません' };
  return withLock_(function () {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var t = ensureSheet_(ss, 'journals');
    var last = t.sheet.getLastRow();
    if (last < 2) return { ok: true };
    var vals = t.sheet.getRange(2, 1, last - 1, t.width).getValues();
    var removed = [];
    for (var i = vals.length - 1; i >= 0; i--) {
      var o = rowToObj_(vals[i], t.map);
      if (String(o.slip) === String(slip)) {
        var le = checkLocked_(o);
        if (le) return { ok: false, error: le };
        removed.push(o);
        t.sheet.deleteRow(i + 2);
      }
    }
    log_(ss, '振替伝票削除', '仕訳帳', slip, removed, null);
    return { ok: true, removed: removed.length };
  });
}

// まとめて追加（定期取引の一括登録など）。items = [{kind:'records'|'incomes'|..., record:{...}}]
function batchAdd_(items) {
  if (!items || !items.length) return { ok: false, error: '登録するデータがありません' };
  return withLock_(function () {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var ids = [];
    for (var i = 0; i < items.length; i++) {
      var it = items[i];
      if (!SCHEMAS[it.kind]) return { ok: false, error: '不明な種類: ' + it.kind, savedIds: ids };
      var r = appendObj_(ss, it.kind, it.record);
      if (!r.ok) return { ok: false, error: r.error, savedIds: ids };
      log_(ss, '追加(まとめて)', SCHEMAS[it.kind].sheet, r.id, null, r.saved);
      ids.push(r.id);
    }
    return { ok: true, ids: ids };
  });
}

// ============================================================
//  設定（按分割合・開業日・支払方法の対応など。スマホとPCで共通）
// ============================================================
function readSettings_() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sh = ensureKVSheet_(ss, SETTINGS_SHEET, ['キー', '値（JSON）'], '#374151');
  var last = sh.getLastRow();
  var s = {};
  if (last < 2) return s;
  sh.getRange(2, 1, last - 1, 2).getValues().forEach(function (r) {
    if (!r[0]) return;
    try { s[String(r[0])] = JSON.parse(r[1]); } catch (e) { s[String(r[0])] = r[1]; }
  });
  return s;
}

function saveSettings_(partial) {
  if (!partial) return { ok: false, error: '設定がありません' };
  return withLock_(function () {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sh = ensureKVSheet_(ss, SETTINGS_SHEET, ['キー', '値（JSON）'], '#374151');
    var last = sh.getLastRow();
    var keys = last >= 2 ? sh.getRange(2, 1, last - 1, 1).getValues().map(function (r) { return String(r[0]); }) : [];
    var before = readSettings_();
    for (var k in partial) {
      var val = JSON.stringify(partial[k]);
      var idx = keys.indexOf(k);
      if (idx >= 0) sh.getRange(idx + 2, 2).setValue(val);
      else { sh.appendRow([k, val]); keys.push(k); }
    }
    var changed = {};
    for (var k2 in partial) changed[k2] = { before: before[k2], after: partial[k2] };
    log_(ss, '設定変更', SETTINGS_SHEET, Object.keys(partial).join(','), null, changed);
    return { ok: true, settings: readSettings_() };
  });
}

// ============================================================
//  変更履歴
// ============================================================
function log_(ss, op, sheetName, id, before, after) {
  try {
    var sh = ensureKVSheet_(ss, LOG_SHEET, ['日時', '操作', 'シート', 'ID', '変更前', '変更後'], '#6b7280');
    sh.appendRow([now_(), op, sheetName, String(id || ''), before ? JSON.stringify(before) : '', after ? JSON.stringify(after) : '']);
  } catch (e) { /* 履歴の失敗で本処理を止めない */ }
}

function getLog_(limit) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sh = ensureKVSheet_(ss, LOG_SHEET, ['日時', '操作', 'シート', 'ID', '変更前', '変更後'], '#6b7280');
  var last = sh.getLastRow();
  if (last < 2) return { ok: true, log: [] };
  var n = Math.min(limit || 100, last - 1);
  var vals = sh.getRange(last - n + 1, 1, n, 6).getValues().reverse();
  return { ok: true, log: vals.map(function (r) {
    return { at: (r[0] instanceof Date) ? Utilities.formatDate(r[0], TZ, 'yyyy-MM-dd HH:mm:ss') : String(r[0]),
             op: String(r[1]), sheet: String(r[2]), id: String(r[3]), before: String(r[4]), after: String(r[5]) };
  }) };
}

// ============================================================
//  帳簿・決算書をシートに書き出す（「出力_」で始まるシートだけを書き換えます）
// ============================================================
function writeReport_(name, rows) {
  if (!name || String(name).indexOf('出力_') !== 0) return { ok: false, error: '書き出し先のシート名は「出力_」で始めてください' };
  if (!rows || !rows.length) return { ok: false, error: '書き出すデータがありません' };
  return withLock_(function () {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sh = ss.getSheetByName(name) || ss.insertSheet(name);
    sh.clear();
    var width = 0;
    rows.forEach(function (r) { if (r.length > width) width = r.length; });
    var grid = rows.map(function (r) { var a = r.slice(); while (a.length < width) a.push(''); return a; });
    sh.getRange(1, 1, grid.length, width).setValues(grid);
    sh.getRange(1, 1, 1, width).setFontWeight('bold');
    sh.setFrozenRows(1);
    log_(ss, '帳簿書き出し', name, '', null, { rows: rows.length });
    return { ok: true, url: ss.getUrl() + '#gid=' + sh.getSheetId() };
  });
}

// ============================================================
//  バックアップ（スプレッドシートをまるごとコピー）
// ============================================================
function backupNow() {
  var file = DriveApp.getFileById(SPREADSHEET_ID);
  var folders = DriveApp.getFoldersByName(BACKUP_FOLDER);
  var folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(BACKUP_FOLDER);
  var name = '経費ノート_バックアップ_' + Utilities.formatDate(new Date(), TZ, 'yyyyMMdd_HHmmss');
  var copy = file.makeCopy(name, folder);
  try { log_(SpreadsheetApp.openById(SPREADSHEET_ID), 'バックアップ', '', copy.getId(), null, { name: name }); } catch (e) {}
  return { ok: true, name: name, url: copy.getUrl(), folderUrl: folder.getUrl() };
}

// 毎週月曜の深夜に自動バックアップ（Apps Script の画面からこの関数を1回実行してください）
function setupAutoBackup() {
  var exists = ScriptApp.getProjectTriggers().some(function (t) { return t.getHandlerFunction() === 'autoBackup'; });
  if (!exists) {
    ScriptApp.newTrigger('autoBackup').timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(3).create();
  }
  return { ok: true, message: exists ? '自動バックアップはすでに設定されています' : '毎週月曜 3時台に自動バックアップするよう設定しました' };
}
function autoBackup() { backupNow(); }

// 合言葉を作る（任意。Apps Script の画面から1回実行し、ログに出た文字をアプリに入力）
function setupAppKey() {
  var key = Utilities.getUuid().replace(/-/g, '').slice(0, 16);
  PropertiesService.getScriptProperties().setProperty('APP_KEY', key);
  Logger.log('アプリキー: ' + key + '  ← アプリの「設定 → 接続」に入力してください');
  return key;
}
function removeAppKey() { PropertiesService.getScriptProperties().deleteProperty('APP_KEY'); }

// ============================================================
//  領収書の写真（Googleドライブに保存・文字の読み取り・取引との紐づけ）
// ============================================================
// 写真を受け取ってドライブに保存し、文字を読み取って「領収書」シートに記録する
// receipt = { id, takenAt, thumb, ocr(true/false), note }, image = base64（"data:" なし）, mime = image/jpeg など
function uploadReceipt_(rc, image, mime) {
  rc = rc || {};
  if (!image) return { ok: false, error: '写真のデータがありません' };
  var id = String(rc.id || ('R' + newId_()));
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var t = ensureSheet_(ss, 'receipts');
  var exist = findRow_(t, id);
  if (exist > 0) return { ok: true, duplicate: true, receipt: rowToObj_(t.sheet.getRange(exist, 1, 1, t.width).getValues()[0], t.map) };

  mime = /^image\/(jpeg|png|gif)$/.test(String(mime)) ? String(mime) : 'image/jpeg';
  var ext = mime === 'image/png' ? '.png' : mime === 'image/gif' ? '.gif' : '.jpg';
  var takenAt = normDateTime_(rc.takenAt) || now_();
  var name = '未登録_' + takenAt.replace(/[^0-9]/g, '').slice(0, 14) + '_' + id + ext;
  var bytes = Utilities.base64Decode(String(image).replace(/^data:[^,]*,/, ''));
  if (bytes.length > 10 * 1024 * 1024) return { ok: false, error: '写真が大きすぎます（10MBまで）' };
  var blob = Utilities.newBlob(bytes, mime, name);
  var file = receiptFolder_(takenAt.slice(0, 4)).createFile(blob);

  var ocr = rc.ocr === false ? { text: '', error: '' } : ocrImage_(blob, id);
  var obj = {
    id: id, takenAt: takenAt, fileName: name, url: file.getUrl(), fileId: file.getId(),
    status: '未登録', linkKind: '', linkId: '', ocrText: String(ocr.text || '').slice(0, 5000),
    note: rc.note || '', thumb: String(rc.thumb || '').length < 45000 ? String(rc.thumb || '') : ''
  };
  var res = withLock_(function () {
    var again = findRow_(t, id);
    if (again > 0) return { dup: true, row: again };
    var a = appendObj_(ss, 'receipts', obj);
    if (a.ok) log_(ss, '領収書の保存', '領収書', id, null, { fileName: name, url: obj.url });
    return a;
  });
  if (res.dup) { // 同時に2回送られた場合は、あとから作ったファイルを捨てる
    try { file.setTrashed(true); } catch (e) {}
    return { ok: true, duplicate: true, receipt: rowToObj_(t.sheet.getRange(res.row, 1, 1, t.width).getValues()[0], t.map) };
  }
  if (!res.ok) { try { file.setTrashed(true); } catch (e) {} return res; }
  return { ok: true, receipt: obj, ocrError: ocr.error || '' };
}

// 写真の文字を読み取る（Googleドライブの文字認識。Drive API サービスが必要）
function ocrImage_(blob, id) {
  if (typeof Drive === 'undefined' || !Drive.Files) {
    return { text: '', error: '文字の読み取り機能（Drive API サービス）が追加されていません。写真は保存しました。手順書の「領収書の文字読み取り」を見て追加してください。' };
  }
  try {
    var doc = Drive.Files.create({ name: '_ocr_' + id, mimeType: 'application/vnd.google-apps.document' }, blob, { ocrLanguage: 'ja' });
    var text = '';
    try { text = DocumentApp.openById(doc.id).getBody().getText(); }
    finally { try { DriveApp.getFileById(doc.id).setTrashed(true); } catch (e) {} }
    text = String(text || '').trim();
    return { text: text, error: text ? '' : '文字を読み取れませんでした（写真が暗い・ぼやけている・傾いている可能性があります）' };
  } catch (err) {
    var m = String(err && err.message ? err.message : err);
    if (/権限|permission|authorization|auth\/documents|auth\/drive/i.test(m)) {
      return { text: '', error: '文字の読み取りに必要な「Googleドキュメント」の許可がありません（写真は保存しました）。Apps Script の画面で setupReceipts を実行し、許可画面で「すべて選択」にチェックを入れて許可してください。そのあと「デプロイを管理」で新バージョンにしてください。', needAuth: true };
    }
    return { text: '', error: '文字の読み取りに失敗しました：' + m };
  }
}

// 「経費ノート_領収書／2026」のような年ごとのフォルダを返す
function receiptFolder_(year) {
  var props = PropertiesService.getScriptProperties();
  var root = null;
  var rootId = props.getProperty('RECEIPT_FOLDER_ID');
  if (rootId) { try { root = DriveApp.getFolderById(rootId); if (root.isTrashed && root.isTrashed()) root = null; } catch (e) { root = null; } }
  if (!root) {
    var it = DriveApp.getFoldersByName(RECEIPT_FOLDER);
    root = it.hasNext() ? it.next() : DriveApp.createFolder(RECEIPT_FOLDER);
    props.setProperty('RECEIPT_FOLDER_ID', root.getId());
  }
  var y = /^\d{4}$/.test(String(year)) ? String(year) : Utilities.formatDate(new Date(), TZ, 'yyyy');
  var sub = root.getFoldersByName(y);
  return sub.hasNext() ? sub.next() : root.createFolder(y);
}

// 取引と写真を紐づけ、ファイル名を「日付_金額円_内容.jpg」に変える（あとで探しやすくするため）
function linkReceipt_(ss, rid, kind, o) {
  var t = ensureSheet_(ss, 'receipts');
  var rowNo = findRow_(t, rid);
  if (rowNo < 0) return;
  var range = t.sheet.getRange(rowNo, 1, 1, t.width);
  var row = range.getValues()[0];
  var rc = rowToObj_(row, t.map);
  var amount = kind === 'deprec' ? o.cost : o.amount;
  var ext = (String(rc.fileName).match(/\.[a-z]+$/i) || ['.jpg'])[0];
  var name = [o.date || '日付なし', amount !== '' && amount !== undefined ? Math.round(toNum(amount)) + '円' : '', o.memo || o.name || '']
    .filter(function (x) { return x !== ''; }).join('_').replace(/[\\\/:*?"<>|\r\n\t]/g, ' ').slice(0, 90) + ext;
  if (rc.fileId && name !== rc.fileName) {
    try { DriveApp.getFileById(rc.fileId).setName(name); row[t.map.fileName - 1] = name; } catch (e) {}
  }
  row[t.map.status - 1] = '登録済';
  row[t.map.linkKind - 1] = SCHEMAS[kind].sheet;
  row[t.map.linkId - 1] = String(o.id);
  row[t.map.updatedAt - 1] = now_();
  range.setValues([row]);
}

// 取引から写真を外す（写真は消さずに「未登録」に戻す）
function unlinkReceipt_(ss, rid, fromId) {
  var t = ensureSheet_(ss, 'receipts');
  var rowNo = findRow_(t, rid);
  if (rowNo < 0) return;
  var range = t.sheet.getRange(rowNo, 1, 1, t.width);
  var row = range.getValues()[0];
  var rc = rowToObj_(row, t.map);
  if (fromId && rc.linkId && String(rc.linkId) !== String(fromId)) return; // ほかの取引に付け替え済み
  row[t.map.status - 1] = '未登録';
  row[t.map.linkKind - 1] = '';
  row[t.map.linkId - 1] = '';
  row[t.map.updatedAt - 1] = now_();
  if (rc.fileId && String(rc.fileName).indexOf('未登録_') !== 0) {
    try { var nm = '未登録_' + rc.fileName; DriveApp.getFileById(rc.fileId).setName(nm); row[t.map.fileName - 1] = nm; } catch (e) {}
  }
  range.setValues([row]);
}

// アプリで写真を表示するために画像データを返す
function getReceiptImage_(id) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var t = ensureSheet_(ss, 'receipts');
  var rowNo = findRow_(t, id);
  if (rowNo < 0) return { ok: false, error: '領収書が見つかりません' };
  var rc = rowToObj_(t.sheet.getRange(rowNo, 1, 1, t.width).getValues()[0], t.map);
  try {
    var blob = DriveApp.getFileById(rc.fileId).getBlob();
    return { ok: true, mime: blob.getContentType(), data: Utilities.base64Encode(blob.getBytes()), url: rc.url, fileName: rc.fileName };
  } catch (e) {
    return { ok: false, error: 'ドライブの写真を開けませんでした（削除された可能性があります）：' + e.message, url: rc.url };
  }
}

function updateReceipt_(rc) {
  if (!rc || !rc.id) return { ok: false, error: 'IDがありません' };
  return updateRow_('receipts', { id: rc.id, note: rc.note || '' });
}

// 写真を削除（取引に紐づいているものは削除できない）。ファイルはドライブのゴミ箱へ（30日間は戻せます）
function deleteReceipt_(id) {
  return withLock_(function () {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var t = ensureSheet_(ss, 'receipts');
    var rowNo = findRow_(t, id);
    if (rowNo < 0) return { ok: true, notFound: true };
    var rc = rowToObj_(t.sheet.getRange(rowNo, 1, 1, t.width).getValues()[0], t.map);
    if (rc.status === '登録済') return { ok: false, error: 'この写真は取引に紐づいています。取引の修正画面で写真を「外す」してから削除してください。' };
    try { DriveApp.getFileById(rc.fileId).setTrashed(true); } catch (e) {}
    t.sheet.deleteRow(rowNo);
    rc.thumb = '';
    log_(ss, '領収書の削除', '領収書', id, rc, null);
    return { ok: true };
  });
}

function normDateTime_(v) {
  if (!v) return '';
  if (v instanceof Date) return Utilities.formatDate(v, TZ, 'yyyy-MM-dd HH:mm:ss');
  var m = String(v).match(/^(\d{4})-(\d{2})-(\d{2})[ T](\d{2}):(\d{2})(?::(\d{2}))?/);
  return m ? m[1] + '-' + m[2] + '-' + m[3] + ' ' + m[4] + ':' + m[5] + ':' + (m[6] || '00') : '';
}

// 【Apps Script の画面から1回実行】領収書フォルダを作り、文字読み取りが使えるか確認する（初回は許可画面が出ます）
function setupReceipts() {
  // 許可画面で外れている項目があれば、ここでもう一度許可を求める（このときは「すべて選択」にチェック）
  if (ScriptApp.requireAllScopes) ScriptApp.requireAllScopes(ScriptApp.AuthMode.FULL);
  initSheets();
  var folder = receiptFolder_(Utilities.formatDate(new Date(), TZ, 'yyyy'));
  var driveApi = (typeof Drive !== 'undefined' && !!Drive.Files);
  // Googleドキュメントの許可を実際に確かめる（一時的な文書を作ってすぐゴミ箱へ）
  var docsOk = false, docsErr = '';
  try { var d = DocumentApp.create('_経費ノート_許可の確認'); DriveApp.getFileById(d.getId()).setTrashed(true); docsOk = true; }
  catch (e) { docsErr = String(e && e.message ? e.message : e); }
  var ocrReady = driveApi && docsOk;
  var msg = '領収書フォルダ：' + folder.getUrl() +
    '\nDrive API サービス：' + (driveApi ? 'OK' : '未追加（サービス ＋ → Drive API を追加してください）') +
    '\nGoogleドキュメントの許可：' + (docsOk ? 'OK' : 'なし（' + docsErr + '）') +
    '\n文字の読み取り：' + (ocrReady ? '使えます' : '使えません');
  Logger.log(msg);
  return { ok: true, folderUrl: folder.getUrl(), ocr: ocrReady, message: msg };
}

// 保存済みの写真の文字をもう一度読み取る（許可を直したあとや、読み取りに失敗したとき）
function reocrReceipt_(id) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var t = ensureSheet_(ss, 'receipts');
  var rowNo = findRow_(t, id);
  if (rowNo < 0) return { ok: false, error: '領収書が見つかりません' };
  var rc = rowToObj_(t.sheet.getRange(rowNo, 1, 1, t.width).getValues()[0], t.map);
  var blob;
  try { blob = DriveApp.getFileById(rc.fileId).getBlob(); }
  catch (e) { return { ok: false, error: 'ドライブの写真を開けませんでした：' + e.message }; }
  var ocr = ocrImage_(blob, id);
  if (ocr.text) {
    withLock_(function () {
      var r2 = findRow_(t, id);
      if (r2 > 0) { t.sheet.getRange(r2, t.map.ocrText).setValue(safeText_(String(ocr.text).slice(0, 5000))); t.sheet.getRange(r2, t.map.updatedAt).setValue(now_()); }
    });
    rc.ocrText = String(ocr.text).slice(0, 5000);
  }
  if (rc.status === '登録済') rc.thumb = '';
  return { ok: true, receipt: rc, ocrError: ocr.error || '' };
}

// ============================================================
//  USD/JPY レート取得（GOOGLEFINANCE の終値。TTMの近似値）
// ============================================================
function getUSDRate(dateStr) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var tmp = ss.getSheetByName('_fx_tmp');
    if (!tmp) { tmp = ss.insertSheet('_fx_tmp'); tmp.hideSheet(); }
    var p = String(dateStr).split('-');
    var d = new Date(parseInt(p[0], 10), parseInt(p[1], 10) - 1, parseInt(p[2], 10));
    var start = new Date(d.getTime() - 7 * 24 * 60 * 60 * 1000); // 土日祝に備えて1週間前から取得
    var end = new Date(d.getTime() + 24 * 60 * 60 * 1000);
    var f = function (x) { return x.getFullYear() + ',' + (x.getMonth() + 1) + ',' + x.getDate(); };
    tmp.getRange('A1').setFormula('=IFERROR(GOOGLEFINANCE("CURRENCY:USDJPY","close",DATE(' + f(start) + '),DATE(' + f(end) + ')),"")');
    SpreadsheetApp.flush();
    Utilities.sleep(1500);
    var vals = tmp.getRange('A1:B12').getValues();
    tmp.getRange('A1').clearContent();
    var target = normDate_(dateStr);
    var rate = 0, rateDate = '';
    // 取引日当日、なければ直前の営業日の終値を使う
    for (var i = 1; i < vals.length; i++) {
      var rd = (vals[i][0] instanceof Date) ? Utilities.formatDate(vals[i][0], TZ, 'yyyy-MM-dd') : '';
      if (typeof vals[i][1] === 'number' && vals[i][1] > 0 && rd && rd <= target) {
        rate = vals[i][1];
        rateDate = rd;
      }
    }
    if (rate > 0) return { ok: true, rate: Math.round(rate * 100) / 100, rateDate: rateDate };
    return { ok: false, error: 'レートを取得できませんでした' };
  } catch (err) {
    return { ok: false, error: String(err) };
  }
}
