/*******************************************************************
 * MIRACO 無料相談 予約バックエンド（Google Apps Script）
 * -----------------------------------------------------------------
 * 役割：
 *   1) 予約ページ（booking.html）からの「空き状況」問い合わせに答える
 *      → Google カレンダーに予定がある時間帯の【前後 BUFFER_MINUTES 分】は
 *        予約不可（busy）として返す。
 *   2) 予約フォーム送信を受け取り、カレンダーに予定を登録し、
 *      （任意で）スプレッドシート記録・通知メールを送る。
 *
 * ポイント：
 *   予約ページが開かれるたびに、その時点のカレンダーを毎回読みます。
 *   つまり「常に最新のカレンダーを監視している」状態になります。
 *   常駐プログラムは不要です。
 *
 * ★ このファイルの設定（下の CONFIG）を自分用に書き換えてから
 *   「デプロイ → ウェブアプリ」で公開し、発行された URL を
 *   booking.html の CONFIG.GAS_URL に貼り付けてください。
 *   詳しい手順は同じフォルダの SETUP.md を参照。
 *******************************************************************/

const CONFIG = {
  // 使用するカレンダー。自分のメインカレンダーなら 'primary'。
  // 別カレンダーを使う場合はそのカレンダーID（〜@group.calendar.google.com）。
  CALENDAR_ID: 'primary',

  // 予定の【前後この分数】は予約不可にする。3時間 = 180分。
  BUFFER_MINUTES: 180,

  // 1回の相談の長さ（分）。
  SLOT_MINUTES: 30,

  // タイムゾーン。
  TIMEZONE: 'Asia/Tokyo',

  // 予約可能な時間スロット（booking.html の CONFIG と必ず一致させること）。
  WEEKDAY_SLOTS: ['18:00', '21:00'],          // 平日（月〜金）
  WEEKEND_SLOTS: ['10:00', '15:00', '20:00'], // 土・日

  // 終日予定（例：「休み」など）も予約不可にするか。
  BLOCK_ALL_DAY_EVENTS: true,

  // 予約を受けたらカレンダーに予定を作るか（作ると、その予約自体も
  // 次の人の空き判定に反映され、前後3時間が自動でふさがる）。
  CREATE_EVENT_ON_BOOKING: true,

  // 予約を記録するスプレッドシートのID（任意。空なら記録しない）。
  SHEET_ID: '',

  // 予約通知メールの送信先（任意。空なら送らない。自分のメール推奨）。
  NOTIFY_EMAIL: '',
};

/** 空き状況の問い合わせ（GET）と疎通確認 */
function doGet(e) {
  const action = e && e.parameter ? e.parameter.action : '';
  if (action === 'slots') {
    return json_(getBusySlots_(e.parameter.date));
  }
  if (action === 'month') {
    return json_(getFullDays_(e.parameter.ym));
  }
  return json_({ ok: true, service: 'MIRACO booking backend' });
}

/**
 * 指定月の「満席日（全スロットが前後バッファ込みで埋まっている日）」を返す。
 * カレンダー画面で「満」マークを出すために使う。
 * @param {string} ym 'YYYY-MM'
 */
function getFullDays_(ym) {
  const m = String(ym || '').match(/^(\d{4})-(\d{2})$/);
  if (!m) return { fullDays: [] };
  const y = +m[1], mo = +m[2] - 1;

  const cal = CalendarApp.getCalendarById(CONFIG.CALENDAR_ID) || CalendarApp.getDefaultCalendar();
  if (!cal) return { fullDays: [] };

  const buf = CONFIG.BUFFER_MINUTES * 60000;
  const daysInMonth = new Date(y, mo + 1, 0).getDate();
  const rangeStart = new Date(y, mo, 1, 0, 0, 0);
  const rangeEnd = new Date(y, mo, daysInMonth, 23, 59, 59);
  // 月全体を1回だけ読む（前後バッファ分を広げる）
  const events = cal.getEvents(new Date(rangeStart.getTime() - buf), new Date(rangeEnd.getTime() + buf))
    .filter(function (ev) { return ev.isAllDayEvent() ? CONFIG.BLOCK_ALL_DAY_EVENTS : true; });

  const full = [];
  for (let d = 1; d <= daysInMonth; d++) {
    const dow = new Date(y, mo, d).getDay();
    const isWeekend = (dow === 0 || dow === 6);
    const slots = isWeekend ? CONFIG.WEEKEND_SLOTS : CONFIG.WEEKDAY_SLOTS;
    const allBusy = slots.length > 0 && slots.every(function (t) {
      const hm = t.split(':').map(Number);
      const s = new Date(y, mo, d, hm[0], hm[1], 0);
      const e2 = new Date(s.getTime() + CONFIG.SLOT_MINUTES * 60000);
      const ws = new Date(s.getTime() - buf), we = new Date(e2.getTime() + buf);
      return events.some(function (ev) {
        if (ev.isAllDayEvent()) return true;
        return ev.getStartTime() < we && ev.getEndTime() > ws;
      });
    });
    if (allBusy) full.push(y + '-' + pad2_(mo + 1) + '-' + pad2_(d));
  }
  return { fullDays: full };
}

function pad2_(n) { return String(n).padStart(2, '0'); }

/**
 * 指定日の busy スロットを返す。
 * 「そのスロットの前後 BUFFER_MINUTES 分の窓」に予定が重なっていれば busy。
 * @param {string} dateStr 'YYYY-MM-DD'
 */
function getBusySlots_(dateStr) {
  const m = String(dateStr || '').match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!m) return { busySlots: [] };
  const y = +m[1], mo = +m[2] - 1, d = +m[3];

  const dow = new Date(y, mo, d).getDay();
  const isWeekend = (dow === 0 || dow === 6);
  const slots = isWeekend ? CONFIG.WEEKEND_SLOTS : CONFIG.WEEKDAY_SLOTS;

  const cal = CalendarApp.getCalendarById(CONFIG.CALENDAR_ID) || CalendarApp.getDefaultCalendar();
  if (!cal) return { busySlots: [] };

  const buf = CONFIG.BUFFER_MINUTES * 60000;
  const dayStart = new Date(y, mo, d, 0, 0, 0);
  const dayEnd   = new Date(y, mo, d, 23, 59, 59);

  // バッファ分だけ広げてイベント取得（前日夜・翌日早朝の予定も考慮）
  const events = cal.getEvents(new Date(dayStart.getTime() - buf), new Date(dayEnd.getTime() + buf))
    .filter(function (ev) {
      if (ev.isAllDayEvent()) return CONFIG.BLOCK_ALL_DAY_EVENTS;
      return true;
    });

  const busy = slots.filter(function (t) {
    const hm = t.split(':').map(Number);
    const slotStart = new Date(y, mo, d, hm[0], hm[1], 0);
    const slotEnd   = new Date(slotStart.getTime() + CONFIG.SLOT_MINUTES * 60000);
    // このスロットの前後バッファを含めた「予約するとかぶる窓」
    const winStart = new Date(slotStart.getTime() - buf);
    const winEnd   = new Date(slotEnd.getTime() + buf);
    // 終日予定はその日全体を対象に
    return events.some(function (ev) {
      if (ev.isAllDayEvent()) return true;
      return ev.getStartTime() < winEnd && ev.getEndTime() > winStart;
    });
  });

  return { busySlots: busy };
}

/** 予約フォームの送信（POST） */
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);

    // 予約直前にもう一度サーバ側で空きを確認（二重予約・直前予定を防ぐ）
    if (data.startISO) {
      const start = new Date(data.startISO);
      if (isNaN(start.getTime()) || !isSlotFree_(start)) {
        return json_({ ok: false, error: 'slot_unavailable' });
      }
      if (CONFIG.CREATE_EVENT_ON_BOOKING) {
        const cal = CalendarApp.getCalendarById(CONFIG.CALENDAR_ID) || CalendarApp.getDefaultCalendar();
        const end = new Date(start.getTime() + CONFIG.SLOT_MINUTES * 60000);
        cal.createEvent(
          '【無料相談】' + (data.name || ''),
          start, end,
          {
            description: [
              'お名前: ' + (data.name || ''),
              'メール: ' + (data.email || ''),
              '会社/屋号/業種: ' + (data.company || ''),
              'ご相談内容: ' + (data.message || ''),
            ].join('\n'),
            guests: data.email || '',
            sendInvites: true,
          }
        );
      }
    }

    if (CONFIG.SHEET_ID) {
      const sh = SpreadsheetApp.openById(CONFIG.SHEET_ID).getSheets()[0];
      sh.appendRow([new Date(), data.date, data.time, data.name, data.email, data.company, data.message]);
    }
    if (CONFIG.NOTIFY_EMAIL) {
      MailApp.sendEmail(
        CONFIG.NOTIFY_EMAIL,
        '【MIRACO】新しい無料相談の予約',
        [
          '日時: ' + (data.date || '') + ' ' + (data.time || ''),
          'お名前: ' + (data.name || ''),
          'メール: ' + (data.email || ''),
          '会社/屋号/業種: ' + (data.company || ''),
          'ご相談内容: ' + (data.message || ''),
        ].join('\n')
      );
    }
    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

/** 指定開始時刻のスロットが（前後バッファ込みで）空いているか */
function isSlotFree_(start) {
  const cal = CalendarApp.getCalendarById(CONFIG.CALENDAR_ID) || CalendarApp.getDefaultCalendar();
  if (!cal) return true;
  const buf = CONFIG.BUFFER_MINUTES * 60000;
  const end = new Date(start.getTime() + CONFIG.SLOT_MINUTES * 60000);
  const winStart = new Date(start.getTime() - buf);
  const winEnd   = new Date(end.getTime() + buf);
  const events = cal.getEvents(winStart, winEnd).filter(function (ev) {
    if (ev.isAllDayEvent()) return CONFIG.BLOCK_ALL_DAY_EVENTS;
    return true;
  });
  return !events.some(function (ev) {
    if (ev.isAllDayEvent()) return true;
    return ev.getStartTime() < winEnd && ev.getEndTime() > winStart;
  });
}

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
