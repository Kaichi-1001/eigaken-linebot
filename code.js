// ====== 設定項目 ======
const SPREADSHEET_ID = '1DVsttehX74SCDrxkNFf2SJXA4uG3C2arPxLVPbF36f4';
const LINE_ACCESS_TOKEN = 'TmbcHFc6lphQU4IjFHxyHhYKM0VfXDtR6qSUW4ro4XFujnEMSKcREPbC9dNEsbS0KY/ZXlPDcjn/w4/WlPoYGY9U+OJZIV9HkHFpcucjOiLcO1CrJ2OuUQkVij1KLVfm7GGslGLWZksu5pa2iOI0eAdB04t89/1O/w1cDnyilFU=';
const WEB_APP_URL = 'https://script.google.com/macros/s/AKfycbzjZWZFQIEXQHVh4RXZ3YqToVgMsyUKLDeiI_B54h1zcM2fOaUca8KAh94jCoA282P2/exec';

// ====== LINE Webhook処理 ======
function doPost(e) {
  const json = JSON.parse(e.postData.contents);
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('Groups');
  
  json.events.forEach(event => {
    // グループに追加されたらGroupIDを保存
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

// ====== 定期実行トリガー（ボットが行うこと） ======

// 毎月1日：イベント作成アナウンス
function trigger_1st_EventAnnounce() {
  const text = `【お知らせ】\n来月のイベント作成期間が始まりました！\n以下のURLからイベントの提案と投票を行ってください。\n${WEB_APP_URL}`;
  broadcastToGroups(text);
}

// 毎月11日：上映会作成アナウンス
function trigger_11th_ScreeningAnnounce() {
  const text = `【お知らせ】\n来月の上映会登録期間が始まりました！\n以下のURLから見たい映画を登録してください。\n${WEB_APP_URL}`;
  broadcastToGroups(text);
}

// 毎月9日正午：上映会締め切りと一覧表示
function trigger_9th_ScreeningClose() {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('Screenings');
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) {
    broadcastToGroups("今月の上映会は登録されていません。");
    return;
  }
  let text = "【上映会一覧（確定）】\n";
  for (let i = 1; i < data.length; i++) {
    text += `・${data[i][1]} (${data[i][2]})\n`;
  }
  broadcastToGroups(text);
  // リセットする場合はここでシートのデータを消去する処理を入れます
}

// 毎月15日正午：イベント決定
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
    const text = `【イベント決定！】\n来月のイベントが決定しました！\nタイトル: ${bestEvent[1]}\n日時: ${bestEvent[2]}\n詳細: ${bestEvent[3]}\n集合時間: ${bestEvent[4]}\n予算: ${bestEvent[5]}\n獲得票数: ${bestEvent[6]}票`;
    broadcastToGroups(text);
  } else {
    broadcastToGroups("【お知らせ】\n5票以上のイベントがなかったため、来月のイベントは見送りとなりました。");
  }
}

// ====== Webアプリ用API（google.script.run用） ======

// イベント追加
function addEvent(title, datetime, details, meetingTime, budget) {
  if (!title || !datetime) throw new Error("タイトルと日時は必須です");
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('Events');
  const id = Utilities.getUuid();
  sheet.appendRow([id, title, datetime, details, meetingTime, budget, 0]);
  return "イベントを追加しました";
}

// 上映会追加 (重複チェック付き)
function addScreening(movieTitle, datetime) {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('Screenings');
  const data = sheet.getDataRange().getValues();
  
  const targetDate = new Date(datetime).toDateString();
  for (let i = 1; i < data.length; i++) {
    let existingDate = new Date(data[i][2]).toDateString();
    if (existingDate === targetDate) {
      throw new Error("同じ日に既に先約の上映会があります！別の日を指定してください。");
    }
  }
  const id = Utilities.getUuid();
  sheet.appendRow([id, movieTitle, datetime]);
  return "上映会を追加しました";
}

// イベント一覧取得
function getEvents() {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('Events');
  const data = sheet.getDataRange().getValues();
  const events = [];
  
  for (let i = 1; i < data.length; i++) {
    events.push({ 
      id: data[i][0], 
      title: data[i][1], 
      datetime: data[i][2], // JSON.stringifyが自動で安全な形式にしてくれます
      votes: data[i][6] 
    });
  }
  
  // 【重要】配列をそのまま返さず、安全な文字列(JSON)に変換して返す！
  return JSON.stringify(events);
}

// 投票処理
function voteForEvent(id) {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('Events');
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === id) {
      let currentVotes = Number(data[i][6]) || 0;
      sheet.getRange(i + 1, 7).setValue(currentVotes + 1);
      return "投票しました！";
    }
  }
  throw new Error("イベントが見つかりません");
}