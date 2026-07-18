/**
 * 打ち込みゼロ日報 — 通知・記録用 Google Apps Script
 *
 * 役割：
 *  - アプリから送られた日報データを Googleスプレッドシートに自動記録
 *  - 代表者・経理へメールで自動通知
 *
 * 使い方は「通知連携_セットアップガイド.md」を参照してください。
 * （このスクリプトは Googleスプレッドシートから開いた Apps Script に貼り付けます）
 */

// 動作確認用（ブラウザでURLを開くと "OK" と表示されれば公開成功）
function doGet(e) {
  return ContentService.createTextOutput('OK - genba-nippo notify endpoint is running');
}

// アプリからのPOSTを受け取る本体
function doPost(e) {
  try {
    var d = JSON.parse(e.postData.contents);
    logToSheet(d);
    sendMail(d);
    return ContentService
      .createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// スプレッドシートへ1行追記（会社の日報台帳が自動でできる）
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

// 代表者・経理へメール通知
function sendMail(d) {
  if (!d.recipients) return;                 // 宛先が無ければ送らない
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
