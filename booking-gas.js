/**
 * MIRACO 予約システム — Google Apps Script
 * ==========================================
 * 【設定手順】
 * 1. https://script.google.com で新しいプロジェクトを作成
 * 2. このコード全体を貼り付けて保存
 * 3. 下の「設定」セクションを確認・編集
 * 4. 「デプロイ」→「新しいデプロイ」
 *    種類: ウェブアプリ
 *    次のユーザーとして実行: 自分
 *    アクセスできるユーザー: 全員（匿名ユーザーを含む）
 * 5. 表示されたウェブアプリURLを booking.html の CONFIG.GAS_URL に貼り付け
 *
 * 【テスト方法】
 * スクリプトエディタで testSendMail() を選択して「実行」ボタンを押す
 * → メールが届けばOK
 */

// ============================================================
// 設定
// ============================================================
const OWNER_EMAIL  = 'vi.nk6327@gmail.com';  // 通知メールの受信先
const SHEET_NAME   = '予約一覧';              // スプレッドシートのシート名
const CALENDAR_ID  = 'primary';              // Googleカレンダー ID（primaryで自分のカレンダー）

// booking.html の CONFIG と同じ値にする
const WEEKDAY_SLOTS = ['18:00', '21:00'];
const WEEKEND_SLOTS = ['10:00', '15:00', '20:00'];
// ============================================================

/**
 * doGet — 時間スロットの空き確認
 * 呼び出し例: ?action=slots&date=2026-06-27
 */
function doGet(e) {
  const action = e.parameter.action;
  const date   = e.parameter.date;

  if (action === 'slots' && date) {
    const busySlots = getBusySlots(date);
    const output = ContentService
      .createTextOutput(JSON.stringify({ busySlots }))
      .setMimeType(ContentService.MimeType.JSON);
    return output;
  }

  return ContentService
    .createTextOutput(JSON.stringify({ status: 'ok' }))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * doPost — 予約受付・メール送信・スプレッドシート記録
 * Content-Type: text/plain で受け取り、JSON.parse して処理する
 */
function doPost(e) {
  try {
    const raw  = e.postData.contents;
    const data = JSON.parse(raw);

    // ── スプレッドシートに記録 ──
    recordToSheet(data);

    // ── 参加者への確認メール ──
    if (data.email) sendConfirmationToParticipant(data);

    // ── オーナーへの通知メール ──
    sendNotificationToOwner(data);

    return ContentService
      .createTextOutput(JSON.stringify({ status: 'ok' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    console.error('doPost error:', err);
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ============================================================
// カレンダー空き確認
// ============================================================
function getBusySlots(dateStr) {
  try {
    const calendar = CalendarApp.getCalendarById(CALENDAR_ID)
                  || CalendarApp.getDefaultCalendar();

    const [year, month, day] = dateStr.split('-').map(Number);
    const dow   = new Date(year, month - 1, day).getDay();
    const slots = (dow === 0 || dow === 6) ? WEEKEND_SLOTS : WEEKDAY_SLOTS;

    const busySlots = [];
    slots.forEach(timeStr => {
      const [h, m] = timeStr.split(':').map(Number);
      const start  = new Date(year, month - 1, day, h, m, 0);
      const end    = new Date(year, month - 1, day, h, m + 30, 0);
      const events = calendar.getEvents(start, end);
      if (events.length > 0) busySlots.push(timeStr);
    });

    return busySlots;
  } catch (err) {
    console.error('getBusySlots error:', err);
    return []; // エラー時は全スロットを表示
  }
}

// ============================================================
// スプレッドシート記録
// ============================================================
function recordToSheet(data) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  let sheet   = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    const headers = ['受付日時','日付','時間','形式','お名前','メール','会社・業種','相談内容'];
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length)
         .setFontWeight('bold')
         .setBackground('#f3ede4')
         .setFontColor('#3d3530');
    sheet.setFrozenRows(1);
  }

  sheet.appendRow([
    new Date(),
    data.date    || '',
    data.time    || '',
    data.meeting || '',
    data.name    || '',
    data.email   || '',
    data.company || '',
    data.message || '',
  ]);
}

// ============================================================
// 参加者への確認メール
// ============================================================
function sendConfirmationToParticipant(data) {
  GmailApp.sendEmail(
    data.email,
    '【MIRACO】無料相談のご予約を承りました',
    // プレーンテキスト（HTMLが表示されない環境用）
    `${data.name} 様\n\nMIRACOの無料相談にお申し込みいただきありがとうございます。\n\n` +
    `【ご予約内容】\n日時：${data.date} ${data.time}〜\n形式：${data.meeting}\n\n` +
    `当日のGoogle MeetのURLは開催前日までにメールにてお送りします。\n\n` +
    `MIRACO 高木菜摘\nhttps://miracobackoffice.com`,
    {
      name: 'MIRACO 高木菜摘',
      htmlBody: buildParticipantHtml(data),
    }
  );
}

function buildParticipantHtml(data) {
  return `
<!DOCTYPE html>
<html lang="ja">
<body style="margin:0;padding:0;background:#f5f3f0;font-family:'Noto Sans JP',sans-serif">
<div style="max-width:520px;margin:32px auto;background:#fff;border-radius:6px;overflow:hidden;box-shadow:0 2px 16px rgba(0,0,0,0.07)">
  <div style="background:#faf6f0;padding:28px 36px 20px;border-bottom:2px solid #C9A96E">
    <p style="font-size:11px;letter-spacing:0.16em;color:#C9A96E;margin:0 0 4px">MIRACO</p>
    <p style="font-size:20px;font-weight:300;color:#1a1612;margin:0">ご予約を承りました</p>
  </div>
  <div style="padding:28px 36px;color:#3d3530;font-size:14px;line-height:1.8">
    <p style="margin:0 0 16px">${data.name} 様</p>
    <p style="margin:0 0 20px">MIRACOの無料相談にお申し込みいただきありがとうございます。<br>以下の内容でご予約を承りました。</p>
    <table style="border-collapse:collapse;width:100%;margin:0 0 24px">
      <tr>
        <td style="padding:10px 0;border-bottom:1px solid #ede8e1;color:#7a706a;font-size:12px;width:80px;vertical-align:top">日時</td>
        <td style="padding:10px 0;border-bottom:1px solid #ede8e1;font-weight:400">${data.date}&nbsp;&nbsp;${data.time}〜</td>
      </tr>
      <tr>
        <td style="padding:10px 0;border-bottom:1px solid #ede8e1;color:#7a706a;font-size:12px;vertical-align:top">形式</td>
        <td style="padding:10px 0;border-bottom:1px solid #ede8e1">${data.meeting}</td>
      </tr>
      <tr>
        <td style="padding:10px 0;color:#7a706a;font-size:12px;vertical-align:top">お名前</td>
        <td style="padding:10px 0">${data.name}</td>
      </tr>
    </table>
    <div style="background:#faf6f0;padding:16px 20px;border-radius:4px;font-size:13px;color:#3d3530;line-height:1.9">
      📌 当日のGoogle MeetのURLは開催前日までにメールにてお送りします。<br>
      ご不明な点はこのメールへの返信またはLINEでお気軽にご連絡ください。
    </div>
    <p style="margin:24px 0 0;font-size:13px;color:#7a706a;line-height:1.8">
      MIRACO（ミラコ）高木 菜摘<br>
      <a href="https://miracobackoffice.com" style="color:#C9A96E;text-decoration:none">miracobackoffice.com</a>
    </p>
  </div>
</div>
</body>
</html>`;
}

// ============================================================
// オーナー（高木さん）への通知メール
// ============================================================
function sendNotificationToOwner(data) {
  GmailApp.sendEmail(
    OWNER_EMAIL,
    `【MIRACO新規予約】${data.date} ${data.time}〜 ${data.name} 様`,
    `新しい予約が入りました。\n\n` +
    `日時：${data.date} ${data.time}〜\nお名前：${data.name}\nメール：${data.email}\n` +
    `会社・業種：${data.company || '—'}\n相談内容：${data.message || '—'}`,
    {
      name: 'MIRACO予約システム',
      htmlBody: `
<div style="font-family:sans-serif;max-width:480px;color:#1a1612;font-size:14px;line-height:1.8;padding:20px">
  <h2 style="font-size:16px;border-bottom:2px solid #C9A96E;padding-bottom:8px;margin-bottom:20px">
    新規予約が入りました
  </h2>
  <table style="border-collapse:collapse;width:100%">
    <tr><td style="padding:8px 0;border-bottom:1px solid #eee;color:#7a706a;font-size:12px;width:80px">日時</td>
        <td style="padding:8px 0;border-bottom:1px solid #eee"><strong>${data.date}</strong>&nbsp;&nbsp;${data.time}〜</td></tr>
    <tr><td style="padding:8px 0;border-bottom:1px solid #eee;color:#7a706a;font-size:12px">お名前</td>
        <td style="padding:8px 0;border-bottom:1px solid #eee">${data.name}</td></tr>
    <tr><td style="padding:8px 0;border-bottom:1px solid #eee;color:#7a706a;font-size:12px">メール</td>
        <td style="padding:8px 0;border-bottom:1px solid #eee">
          <a href="mailto:${data.email}" style="color:#C9A96E">${data.email}</a></td></tr>
    <tr><td style="padding:8px 0;border-bottom:1px solid #eee;color:#7a706a;font-size:12px">会社・業種</td>
        <td style="padding:8px 0;border-bottom:1px solid #eee">${data.company || '—'}</td></tr>
    <tr><td style="padding:8px 0;color:#7a706a;font-size:12px;vertical-align:top">相談内容</td>
        <td style="padding:8px 0">${data.message || '—'}</td></tr>
  </table>
  <p style="margin-top:16px;font-size:12px;color:#7a706a">
    Googleスプレッドシートにも記録されています。
  </p>
</div>`,
    }
  );
}

// ============================================================
// ★ テスト関数 — スクリプトエディタで直接実行してメール確認
// ============================================================
function testSendMail() {
  const dummyData = {
    date:    '2026年7月1日（火）',
    time:    '18:00',
    meeting: '無料オンライン相談（30分 / Google Meet）',
    name:    'テスト 太郎',
    email:   OWNER_EMAIL,   // 自分のメールに送信
    company: 'テスト株式会社',
    message: 'これはテストメールです。',
  };

  sendConfirmationToParticipant(dummyData);
  sendNotificationToOwner(dummyData);
  Logger.log('テストメール送信完了。受信トレイを確認してください。');
}

function testCalendar() {
  const today = new Date();
  const yr    = today.getFullYear();
  const mo    = String(today.getMonth() + 1).padStart(2, '0');
  const dy    = String(today.getDate() + 3).padStart(2, '0');
  const dateStr = `${yr}-${mo}-${dy}`;
  const busy  = getBusySlots(dateStr);
  Logger.log(`${dateStr} の埋まっているスロット: ${JSON.stringify(busy)}`);
}
