/**
 * 打ち込みゼロ日報 — 通知・記録・クラウド同期用 Google Apps Script（v2）
 *
 * 役割：
 *  1. 日報データを Googleスプレッドシートに自動記録
 *  2. 代表者・経理へメールで自動通知
 *  3. 引き継ぎ状況・現場一覧を全端末で共有（クラウド同期）
 *
 * v1から更新した場合は、コードを貼り替えて「新バージョン」で再デプロイしてください
 * （URLは変わりません）。詳しくは「通知連携_セットアップガイド.md」を参照。
 */

// 動作確認用（URLをブラウザで開くと "OK" と表示されれば公開成功）
function doGet(e) {
  // クラウド同期：保存済みの状態をJSON（またはJSONP）で返す
  if (e && e.parameter && e.parameter.action === 'getState') {
    var body = JSON.stringify(getStoredState());
    if (e.parameter.callback) {
      return ContentService
        .createTextOutput(e.parameter.callback + '(' + body + ')')
        .setMimeType(ContentService.MimeType.JAVASCRIPT);
    }
    return ContentService.createTextOutput(body).setMimeType(ContentService.MimeType.JSON);
  }
  return ContentService.createTextOutput('OK - genba-nippo endpoint is running');
}

// アプリからのPOSTを受け取る本体
function doPost(e) {
  try {
    var d = JSON.parse(e.postData.contents);

    // クラウド同期：状態をマージ保存
    if (d.kind === 'state') {
      var lock = LockService.getScriptLock();
      lock.tryLock(10000);
      try { mergeState(d.state || {}); }
      finally { lock.releaseLock(); }
      return json({ ok: true });
    }

    // 通常の日報：シート記録＋メール通知
    logToSheet(d);
    sendMail(d);
    return json({ ok: true });
  } catch (err) {
    return json({ ok: false, error: String(err) });
  }
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/* ---------- 日報の記録・通知 ---------- */

function logToSheet(d) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('日報ログ') || ss.insertSheet('日報ログ');
  if (sh.getLastRow() === 0) {
    sh.appendRow(['受信日時', '種別', '日付', '現場', '作成者', '作成日時',
      '最終更新者', '最終更新', '人工', '進捗%', 'IP', '作業者', '本文']);
  }
  var workers = '';
  if (d.workers) {
    workers = Object.keys(d.workers).map(function (n) {
      return n + '(' + d.workers[n] + ')';
    }).join(', ');
  }
  sh.appendRow([
    new Date(), d.type || '', d.date || '', d.site || '', d.creator || '',
    d.createdAt || '', d.updatedBy || '', d.updatedAt || '',
    d.nin || '', d.pct || '', d.ip || '', workers, d.text || ''
  ]);
}

function sendMail(d) {
  if (!d.recipients) return;
  var label = d.type === 'update' ? '【日報 更新】'
            : d.type === 'test'   ? '【通知テスト】'
            : '【日報 作成】';
  var subject = label + ' ' + (d.date || '') + ' ' + (d.site || '');
  var body =
      '種別　　：' + (d.type || '') + '\n' +
      '日付　　：' + (d.date || '') + '\n' +
      '現場　　：' + (d.site || '') + '\n' +
      '作成者　：' + (d.creator || '') + '\n' +
      (d.updatedBy ? '最終更新者：' + d.updatedBy + '\n' : '') +
      '人工　　：' + (d.nin || '') + '　進捗：' + (d.pct || '') + '%\n' +
      'IP　　　：' + (d.ip || '') + '\n' +
      '\n---------------- 日報本文 ----------------\n' +
      (d.text || '');
  MailApp.sendEmail(d.recipients, subject, body);
}

/* ---------- クラウド同期（状態の保存・マージ） ---------- */

// 状態は「_sync」シートのA1セルにJSONで保存（大きくても保持できる）
function stateSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName('_sync') || ss.insertSheet('_sync');
}

function getStoredState() {
  var v = stateSheet().getRange('A1').getValue();
  if (!v) return { sites: [], statStore: {}, ninStore: {}, siteTags: {},
                   workers: [], creators: [], siteMeta: {} };
  try { return JSON.parse(v); }
  catch (e) { return { sites: [], statStore: {}, ninStore: {}, siteTags: {},
                       workers: [], creators: [], siteMeta: {} }; }
}

function saveStoredState(s) {
  stateSheet().getRange('A1').setValue(JSON.stringify(s));
}

// 端末から届いた状態を、既存とマージ（リストは和集合、現場ごとは新しい方を採用）
function mergeState(inc) {
  var s = getStoredState();
  ['sites', 'workers', 'creators'].forEach(function (key) {
    var cur = s[key] || [], list = inc[key] || [];
    list.forEach(function (x) { if (cur.indexOf(x) < 0) cur.push(x); });
    s[key] = cur;
  });
  s.statStore = s.statStore || {};
  s.ninStore  = s.ninStore  || {};
  s.siteMeta  = s.siteMeta  || {};
  s.siteTags  = s.siteTags  || {};

  var im = inc.siteMeta || {};
  Object.keys(im).forEach(function (site) {
    var rem = s.siteMeta[site];
    if (!rem || (im[site].at || 0) >= (rem.at || 0)) {          // 新しい方を採用
      if (inc.statStore && (site in inc.statStore)) s.statStore[site] = inc.statStore[site];
      if (inc.ninStore  && (site in inc.ninStore))  s.ninStore[site]  = inc.ninStore[site];
      s.siteMeta[site] = im[site];
    }
  });
  var it = inc.siteTags || {};
  Object.keys(it).forEach(function (site) { s.siteTags[site] = it[site]; });

  saveStoredState(s);
}
