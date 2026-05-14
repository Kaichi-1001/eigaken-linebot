// ====== 設定項目 ======
const SPREADSHEET_ID = PropertiesService.getScriptProperties().getProperty('ssid'); 
const LINE_ACCESS_TOKEN = PropertiesService.getScriptProperties().getProperty('LINE_ACCESS_TOKEN');
const WEB_APP_URL = PropertiesService.getScriptProperties().getProperty('web_url');
const ADMIN_EMAIL = PropertiesService.getScriptProperties().getProperty('admin_email');

// ====== 環境初期化（最初に1回だけ実行する関数） ======
function setupEnvironment() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  
  // シート作成とヘッダー設定関数
  const createSheet = (name, headers) => {
    let sheet = ss.getSheetByName(name);
    if (!sheet) sheet = ss.insertSheet(name);
    sheet.clear(); // 初期化
    sheet.appendRow(headers);
    sheet.setFrozenRows(1);
  };

  createSheet('Groups', ['GroupID']);
  createSheet('Screenings', ['ID', 'MovieTitle', 'Datetime']);
  createSheet('Events', ['ID', 'Title', 'Datetime', 'Details', 'MeetingTime', 'Budget', 'Votes']);
  
  // デフォルトで存在する「シート1」等を削除
  const sheets = ss.getSheets();
  if (sheets.length > 3) ss.deleteSheet(sheets[0]);
}

// ====== LINE Webhook処理 ======
function doPost(e) {
  const json = JSON.parse(e.postData.contents);
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('Groups');
  
  json.events.forEach(event => {
    if (event.type === 'join' && (event.source.type === 'group' || event.source.type === 'room')) {
      const groupId = event.source.groupId || event.source.roomId;
      sheet.appendRow([groupId]);
      pushMessage(groupId, "映画研LineBotです！登録完了しました。");
    }
  });
  return ContentService.createTextOutput(JSON.stringify({content: "ok"})).setMimeType(ContentService.MimeType.JSON);
}

// ====== Webアプリ画面表示 ======
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('映画研イベント・上映会管理')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// ====== LINEメッセージ送信関数 ======
function pushMessage(groupId, text) {
  UrlFetchApp.fetch('https://api.line.me/v2/bot/message/push', {
    method: 'post',
    headers: {
      'Content-Type': 'application/json',
      'Authorization': 'Bearer ' + LINE_ACCESS_TOKEN
    },
    payload: JSON.stringify({
      to: groupId,
      messages: [{ type: 'text', text: text }]
    })
  });
}

function broadcastToGroups(text) {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('Groups');
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0]) pushMessage(data[i][0], text);
  }
}

// ====== 定期実行トリガー ======
function trigger_1st_EventAnnounce() {
  const subject = "【映画研LINEBot】1日のグループ送信のお願い";
  const body = `映画研LINEBotです。1日になりました。以下の文章をコピーして映画研グループに貼り付けてください。\n\n翌月のイベント作成、投票期間がスタートしました。以下のURLから作成を行ってください。\n${WEB_APP_URL}`;

  GmailApp.sendEmail(ADMIN_EMAIL, subject, body);
}

function trigger_11th_ScreeningAnnounce() {
  const now = new Date();
  const n = now.getMonth() + 1; // 現在の月(1〜12)
  const targetMonth = (n + 1) % 12 + 1; // 12月の場合は1月にする計算

  const subject = "【映画研LINEBot】11日のグループ送信のお願い";
  const body = `映画研LINEBotです。\n11日になりました。以下の文章をコピーしてグループに張り付けて送信してください。\n\n${targetMonth}月の上映会作成がスタートしました。以下のURLから作成を行ってください。\n${WEB_APP_URL}`;

  GmailApp.sendEmail(ADMIN_EMAIL, subject, body);
}

function trigger_9th_ScreeningClose() {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('Screenings');
  const data = sheet.getDataRange().getValues();
  
  let listText = "";
  if (data.length <= 1) {
    listText = "今月の上映会は登録されていません。\n";
  } else {
    for (let i = 1; i < data.length; i++) {
      let title = data[i][1];
      let dateObj = new Date(data[i][2]);
      // 日付を yyyy/MM/dd 形式に変換
      let dateStr = Utilities.formatDate(dateObj, 'JST', 'yyyy/MM/dd');
      listText += `"${title}" ${dateStr}\n`;
    }
  }

  const subject = "【映画研LINEBot】9日の上映会リストと予約のお願い";
  const body = `${listText}\n講義室予約を教務webから行い、このリストをグループで送信して告知してください。`;

  GmailApp.sendEmail(ADMIN_EMAIL, subject, body);
}

function trigger_15th_EventDecision() {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('Events');
  const data = sheet.getDataRange().getValues();
  let bestEvent = null;
  let maxVotes = 0;
  
  for (let i = 1; i < data.length; i++) {
    let votes = Number(data[i][6]);
    if (votes >= 5 && votes > maxVotes) {
      maxVotes = votes;
      bestEvent = data[i];
    }
  }
  
  if (bestEvent) {
    let dateObj = new Date(bestEvent[2]);
    // 日付を yyyy/mm/dd time 形式に変換
    let dateStr = Utilities.formatDate(dateObj, 'JST', 'yyyy/MM/dd HH:mm');
    
    broadcastToGroups(`【イベント決定！】\n来月のイベントが決定しました！\nタイトル: ${bestEvent[1]}\n日時: ${dateStr}\n詳細: ${bestEvent[3]}\n集合時間: ${bestEvent[4]}\n予算: ${bestEvent[5]}\n獲得票数: ${bestEvent[6]}票\n\nイベント作成者は、ノートを作成してください！`);
  } else {
    broadcastToGroups("【お知らせ】\n5票以上のイベントがなかったため、来月のイベントは見送りとなりました。");
  }
}

// ====== Webアプリ用API ======
function addEvent(title, datetime, details, meetingTime, budget) {
  if (!title || !datetime) throw new Error("タイトルと日時は必須です");
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('Events');
  const id = Utilities.getUuid();
  sheet.appendRow([id, title, datetime, details, meetingTime, budget, 0, ""]); // 最後に投票者記録用の空欄
  return "イベントを追加しました";
}

function addScreening(movieTitle, datetime) {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('Screenings');
  const data = sheet.getDataRange().getValues();
  const targetDate = new Date(datetime).toDateString();
  for (let i = 1; i < data.length; i++) {
    if (new Date(data[i][2]).toDateString() === targetDate) {
      throw new Error("同じ日に既に先約の上映会があります！別の日を指定してください。");
    }
  }
  sheet.appendRow([Utilities.getUuid(), movieTitle, datetime]);
  return "上映会を追加しました";
}

function getEvents() {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('Events');
  const data = sheet.getDataRange().getValues();
  const events = [];
  
  for (let i = 1; i < data.length; i++) {
    let voters = data[i][7] ? String(data[i][7]).split(',') : [];
    events.push({ 
      id: data[i][0], 
      title: data[i][1], 
      datetime: data[i][2],
      votes: data[i][6],
    });
  }
  return JSON.stringify(events);
}

function voteForEvent(id) {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('Events');
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === id) {
      let voters = data[i][7] ? String(data[i][7]).split(',') : [];
      // 票数を+1する
      let currentVotes = Number(data[i][6]) || 0;
      sheet.getRange(i + 1, 7).setValue(currentVotes + 1);
      return "投票しました！";
    }
  }
  throw new Error("イベントが見つかりません");
}

function unvoteForEvent(id) {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('Events');
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === id) {
      let currentVotes = Number(data[i][6]) || 0;
      if (currentVotes > 0) {
        sheet.getRange(i + 1, 7).setValue(currentVotes - 1);
      }
      return "投票を取り消しました。";
    }
  }
  throw new Error("イベントが見つかりません");
}