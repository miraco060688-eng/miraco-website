/**
 * MIRACO イベント申し込み受付 — Google Apps Script
 * ==========================================================
 * apply.html から送信された申し込みを受け取り、
 *   1. スプレッドシートに記録
 *   2. 申込者へ確認メールを送信
 *   3. 高木さんへ通知メールを送信
 * します。予約システム（booking-gas.js）とは別プロジェクトです。
 *
 * 【設定手順】
 * 1. Googleスプレッドシートを新規作成する（名前は「MIRACO イベント申込」など）
 * 2. そのスプレッドシートで「拡張機能」→「Apps Script」を開く
 * 3. このコード全体を貼り付けて保存
 * 4. 「デプロイ」→「新しいデプロイ」
 *      種類          : ウェブアプリ
 *      実行するユーザー: 自分
 *      アクセスできる人: 全員（匿名ユーザーを含む）
 * 5. 表示されたウェブアプリURLをコピー
 * 6. apply.html の const GAS_URL = 'YOUR_EVENT_GAS_URL' を、そのURLに書き換える
 *
 * 【テスト方法】
 * スクリプトエディタで testApply() を選んで「実行」
 * → 自分宛にメールが2通届き、スプレッドシートに1行増えればOK
 *
 * 【注意】
 * コードを直したあとは「デプロイ」→「デプロイを管理」→鉛筆アイコン→
 * バージョン「新バージョン」を選んで再デプロイしないと反映されません。
 */

// ============================================================
// 設定
// ============================================================
const OWNER_EMAIL = 'miraco060688@gmail.com';   // 通知メールの受信先
const OWNER_NAME  = 'MIRACO 高木菜摘';
const SHEET_NAME  = 'イベント申込';
const SITE_URL    = 'https://miracobackoffice.com';
const LINE_URL    = 'https://lin.ee/v2A0ds8L';
// ============================================================


/**
 * doPost — 申し込みを受け取る
 * apply.html から Content-Type: text/plain で JSON が送られてくる
 */
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);

    recordToSheet(data);
    if (data.email) sendConfirmationToApplicant(data);
    sendNotificationToOwner(data);

    return json({ status: 'ok' });

  } catch (err) {
    console.error('doPost error:', err);
    // 記録に失敗しても申し込みを取りこぼさないよう、生データをメールで送る
    try {
      GmailApp.sendEmail(
        OWNER_EMAIL,
        '【MIRACO】イベント申込の処理でエラーが発生しました',
        'エラー内容：\n' + err.toString() +
        '\n\n受信した生データ：\n' + (e && e.postData ? e.postData.contents : '(なし)')
      );
    } catch (_) { /* メールも失敗した場合は何もできない */ }

    return json({ status: 'error', message: err.toString() });
  }
}

function doGet() {
  return json({ status: 'ok', service: 'MIRACO event application' });
}

function json(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}


// ============================================================
// スプレッドシートに記録
// ============================================================
function recordToSheet(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    const headers = [
      '受付日時', 'イベント', '開催日時', 'お名前', 'ふりがな',
      'メール', '電話', 'お仕事', 'AI経験', 'ききたいこと', '状態',
    ];
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length)
         .setFontWeight('bold')
         .setBackground('#f3ede4')
         .setFontColor('#3d3530');
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(2, 260);
    sheet.setColumnWidth(10, 320);
  }

  sheet.appendRow([
    new Date(),
    data.eventName || '',
    data.eventDate || '',
    data.name      || '',
    data.kana      || '',
    data.email     || '',
    data.tel       || '',
    data.job       || '',
    data.level     || '',
    data.message   || '',
    '申込受付',
  ]);
}


// ============================================================
// 申込者への確認メール
// ============================================================
function sendConfirmationToApplicant(data) {
  const subject = '【MIRACO】お申し込みを承りました（' + (data.eventName || 'イベント') + '）';

  const plain =
    data.name + ' 様\n\n' +
    'このたびはお申し込みいただきありがとうございます。\n' +
    '以下の内容で承りました。\n\n' +
    '【イベント】' + (data.eventName  || '') + '\n' +
    '【日時】    ' + (data.eventDate  || '') + '\n' +
    '【場所】    ' + (data.eventPlace || '') + '\n' +
    '【参加費】  ' + (data.eventFee   || '') + '\n' +
    '【持ち物】  スマートフォンだけでOKです\n\n' +
    '会場の詳細は、開催前日までにこのメールアドレス宛にご案内します。\n' +
    'ご不明な点は、このメールへの返信またはLINEからお気軽にどうぞ。\n\n' +
    OWNER_NAME + '\n' + SITE_URL;

  GmailApp.sendEmail(data.email, subject, plain, {
    name:     OWNER_NAME,
    replyTo:  OWNER_EMAIL,
    htmlBody: buildApplicantHtml(data),
  });
}

function buildApplicantHtml(data) {
  const row = function (label, value) {
    if (!value) return '';
    return '<tr>' +
      '<td style="padding:10px 0;border-bottom:1px solid #ede8e1;color:#7a706a;font-size:12px;width:76px;vertical-align:top">' + label + '</td>' +
      '<td style="padding:10px 0;border-bottom:1px solid #ede8e1;color:#3d3530">' + esc(value) + '</td>' +
      '</tr>';
  };

  return '' +
'<!DOCTYPE html><html lang="ja"><body style="margin:0;padding:0;background:#f5f3f0;font-family:\'Noto Sans JP\',sans-serif">' +
'<div style="max-width:520px;margin:32px auto;background:#fff;border-radius:6px;overflow:hidden;box-shadow:0 2px 16px rgba(0,0,0,0.07)">' +
  '<div style="background:#faf6f0;padding:28px 36px 20px;border-bottom:2px solid #C9A96E">' +
    '<p style="font-size:11px;letter-spacing:0.16em;color:#C9A96E;margin:0 0 4px">MIRACO</p>' +
    '<p style="font-size:20px;font-weight:300;color:#1a1612;margin:0">お申し込みを承りました</p>' +
  '</div>' +
  '<div style="padding:28px 36px;color:#3d3530;font-size:14px;line-height:1.9">' +
    '<p style="margin:0 0 16px">' + esc(data.name) + ' 様</p>' +
    '<p style="margin:0 0 22px">このたびはお申し込みいただきありがとうございます。<br>以下の内容で承りました。</p>' +
    '<table style="border-collapse:collapse;width:100%;margin:0 0 24px">' +
      row('イベント', data.eventName) +
      row('日時',     data.eventDate) +
      row('場所',     data.eventPlace) +
      row('参加費',   data.eventFee) +
      row('持ち物',   'スマートフォンだけでOKです') +
    '</table>' +
    '<div style="background:#faf6f0;padding:16px 20px;border-radius:4px;font-size:13px;line-height:1.9">' +
      '会場の詳細は、開催前日までにこのメールアドレス宛にご案内します。<br>' +
      'ご不明な点は、このメールへの返信または' +
      '<a href="' + LINE_URL + '" style="color:#C9A96E">LINE</a>' +
      'からお気軽にどうぞ。' +
    '</div>' +
    '<p style="margin:24px 0 0;font-size:13px;color:#7a706a;line-height:1.8">' +
      'MIRACO（ミラコ）高木 菜摘<br>' +
      '<a href="' + SITE_URL + '" style="color:#C9A96E;text-decoration:none">miracobackoffice.com</a>' +
    '</p>' +
  '</div>' +
'</div></body></html>';
}


// ============================================================
// 高木さんへの通知メール
// ============================================================
function sendNotificationToOwner(data) {
  const subject = '【MIRACO イベント申込】' + (data.name || '') + ' 様 ／ ' + (data.eventDate || '');

  const plain =
    '新しい申し込みが入りました。\n\n' +
    'イベント　　：' + (data.eventName || '') + '\n' +
    '日時　　　　：' + (data.eventDate || '') + '\n' +
    'お名前　　　：' + (data.name  || '') + '（' + (data.kana || '—') + '）\n' +
    'メール　　　：' + (data.email || '') + '\n' +
    '電話　　　　：' + (data.tel   || '—') + '\n' +
    'お仕事　　　：' + (data.job   || '—') + '\n' +
    'AI経験　　　：' + (data.level || '—') + '\n' +
    'ききたいこと：' + (data.message || '—') + '\n';

  const row = function (label, value) {
    return '<tr>' +
      '<td style="padding:8px 0;border-bottom:1px solid #eee;color:#7a706a;font-size:12px;width:92px;vertical-align:top">' + label + '</td>' +
      '<td style="padding:8px 0;border-bottom:1px solid #eee">' + esc(value || '—') + '</td>' +
      '</tr>';
  };

  const html =
'<div style="font-family:sans-serif;max-width:520px;color:#1a1612;font-size:14px;line-height:1.8;padding:20px">' +
  '<h2 style="font-size:16px;border-bottom:2px solid #C9A96E;padding-bottom:8px;margin-bottom:20px">イベント申し込みが入りました</h2>' +
  '<table style="border-collapse:collapse;width:100%">' +
    row('イベント', data.eventName) +
    row('日時',     data.eventDate) +
    row('お名前',   (data.name || '') + (data.kana ? '（' + data.kana + '）' : '')) +
    '<tr><td style="padding:8px 0;border-bottom:1px solid #eee;color:#7a706a;font-size:12px">メール</td>' +
    '<td style="padding:8px 0;border-bottom:1px solid #eee"><a href="mailto:' + esc(data.email) + '" style="color:#C9A96E">' + esc(data.email) + '</a></td></tr>' +
    row('電話',     data.tel) +
    row('お仕事',   data.job) +
    row('AI経験',   data.level) +
    row('ききたいこと', data.message) +
  '</table>' +
  '<p style="margin-top:16px;font-size:12px;color:#7a706a">スプレッドシート「' + SHEET_NAME + '」にも記録されています。</p>' +
'</div>';

  GmailApp.sendEmail(OWNER_EMAIL, subject, plain, {
    name:     'MIRACO 申込通知',
    replyTo:  data.email || OWNER_EMAIL,
    htmlBody: html,
  });
}


function esc(s) {
  return String(s == null ? '' : s)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');
}


// ============================================================
// ★ テスト用 — スクリプトエディタで直接実行する
// ============================================================
function testApply() {
  const dummy = {
    type:       'event',
    eventName:  'AI活用を楽しむ会 ── スマホでAIを使いこなす',
    eventDate:  '2026年7月31日（金）18:00〜20:00',
    eventPlace: '高岡市内（会場はお申し込み後にご案内）',
    eventFee:   'テスト金額',
    name:       'テスト 太郎',
    kana:       'てすと たろう',
    email:      OWNER_EMAIL,   // 自分宛に送る
    tel:        '090-0000-0000',
    job:        '飲食店経営',
    level:      '少し触ったことがある',
    message:    'スマホで議事録を作れるようになりたいです。',
  };

  recordToSheet(dummy);
  sendConfirmationToApplicant(dummy);
  sendNotificationToOwner(dummy);
  Logger.log('テスト完了。メール2通とスプレッドシートを確認してください。');
}
