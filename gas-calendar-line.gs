// ========================================
// Google Apps Script: カレンダー → LINE通知
// ========================================
// 設定方法:
// 1. https://script.google.com で新規プロジェクト作成
// 2. このコードを貼り付け
// 3. LINE_TOKEN と LINE_USER_ID を設定
// 4. トリガーを設定（setupTriggers関数を1回実行）
// ========================================

// ===== 設定 =====
const LINE_TOKEN = 'YOUR_CHANNEL_ACCESS_TOKEN';  // LINE Messaging APIのチャネルアクセストークン
const LINE_USER_ID = 'YOUR_LINE_USER_ID';         // 送信先のLINEユーザーID
const CALENDAR_ID = 'primary';                     // カレンダーID（primaryはデフォルト）
const REMINDER_MINUTES = 10;                       // 予定の何分前に通知するか
const MORNING_HOUR = 7;                            // 朝の一覧通知の時刻（7時）

// ===== メイン: 朝の予定一覧を送信 =====
function sendDailySchedule() {
  const today = new Date();
  const events = getEventsForDate(today);

  if (events.length === 0) {
    sendLine(formatDate(today) + '\n\n予定はありません 🎉');
    return;
  }

  let msg = '【' + formatDate(today) + 'の予定】\n';
  events.forEach(e => {
    if (e.isAllDay) {
      msg += '終日　' + e.title + '\n';
    } else {
      msg += e.startTime + '〜' + e.endTime + '　' + e.title + '\n';
    }
  });
  msg += '\n計 ' + events.length + '件';

  sendLine(msg);
}

// ===== メイン: 直前リマインド通知 =====
function sendReminders() {
  const now = new Date();
  const checkFrom = new Date(now.getTime());
  const checkTo = new Date(now.getTime() + REMINDER_MINUTES * 60 * 1000);

  const calendar = CalendarApp.getCalendarById(CALENDAR_ID);
  const events = calendar.getEvents(checkFrom, checkTo);

  events.forEach(event => {
    if (event.isAllDayEvent()) return;

    const start = event.getStartTime();
    const diffMin = Math.round((start.getTime() - now.getTime()) / 60000);

    // REMINDER_MINUTES分前〜(REMINDER_MINUTES-5)分前の範囲のみ通知（重複防止）
    if (diffMin >= 0 && diffMin <= REMINDER_MINUTES && diffMin > REMINDER_MINUTES - 5) {
      const startStr = Utilities.formatDate(start, Session.getScriptTimeZone(), 'HH:mm');
      const endStr = Utilities.formatDate(event.getEndTime(), Session.getScriptTimeZone(), 'HH:mm');
      let msg = '⏰ まもなく予定\n\n';
      msg += startStr + '〜' + endStr + '\n';
      msg += '📌 ' + event.getTitle();
      if (event.getLocation()) {
        msg += '\n📍 ' + event.getLocation();
      }
      sendLine(msg);
    }
  });
}

// ===== ヘルパー: 指定日のイベント取得 =====
function getEventsForDate(date) {
  const calendar = CalendarApp.getCalendarById(CALENDAR_ID);
  const startOfDay = new Date(date.getFullYear(), date.getMonth(), date.getDate(), 0, 0, 0);
  const endOfDay = new Date(date.getFullYear(), date.getMonth(), date.getDate(), 23, 59, 59);
  const events = calendar.getEvents(startOfDay, endOfDay);
  const tz = Session.getScriptTimeZone();

  return events.map(e => ({
    title: e.getTitle(),
    isAllDay: e.isAllDayEvent(),
    startTime: e.isAllDayEvent() ? '' : Utilities.formatDate(e.getStartTime(), tz, 'HH:mm'),
    endTime: e.isAllDayEvent() ? '' : Utilities.formatDate(e.getEndTime(), tz, 'HH:mm'),
    location: e.getLocation() || ''
  }));
}

// ===== ヘルパー: 日付フォーマット =====
function formatDate(date) {
  const days = ['日', '月', '火', '水', '木', '金', '土'];
  const y = date.getFullYear();
  const m = String(date.getMonth() + 1).padStart(2, '0');
  const d = String(date.getDate()).padStart(2, '0');
  const w = days[date.getDay()];
  return y + '年' + m + '月' + d + '日（' + w + '）';
}

// ===== ヘルパー: LINE Messaging API送信 =====
function sendLine(message) {
  const url = 'https://api.line.me/v2/bot/message/push';
  const payload = {
    to: LINE_USER_ID,
    messages: [{ type: 'text', text: message }]
  };

  const options = {
    method: 'post',
    contentType: 'application/json',
    headers: { 'Authorization': 'Bearer ' + LINE_TOKEN },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  const res = UrlFetchApp.fetch(url, options);
  const code = res.getResponseCode();
  if (code !== 200) {
    Logger.log('LINE送信エラー: ' + code + ' ' + res.getContentText());
  }
}

// ===== トリガー設定（初回1回だけ実行） =====
function setupTriggers() {
  // 既存トリガー削除
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));

  // 朝の一覧通知（毎日7時）
  ScriptApp.newTrigger('sendDailySchedule')
    .timeBased()
    .everyDays(1)
    .atHour(MORNING_HOUR)
    .create();

  // リマインド通知（5分ごとにチェック）
  ScriptApp.newTrigger('sendReminders')
    .timeBased()
    .everyMinutes(5)
    .create();

  Logger.log('トリガー設定完了！');
}

// ===== テスト用 =====
function testSendDaily() {
  sendDailySchedule();
}

function testSendLine() {
  sendLine('テスト通知です 🔔');
}
