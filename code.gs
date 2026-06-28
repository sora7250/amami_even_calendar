// 【修正点】2つのスプレッドシートIDを個別に設定できるようにしました！
const EVENT_SS_ID = "1DYB-TwuLtmk-C8aNE-mXOq1x6IXJsoAECo4SywafCgM";
const ACTION_SS_ID = "1Pz2fb5an1YAPlYK0-Y17RVVXfXn2CrBdY5NoQ_mVgSw";

function doGet() {
  return HtmlService.createTemplateFromFile('index')
      .evaluate()
      .setTitle('奄美イベントカレンダー')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * フロントエンドからの非同期リクエストを処理するAPI
 */
function doPost(e) {
  const params = JSON.parse(e.postData.contents);
  const action = params.action;
  
  if (action === 'getEvents') {
    return ContentService.createTextOutput(JSON.stringify(getEventsData()))
        .setMimeType(ContentService.MimeType.JSON);
  }
  
  if (action === 'addLike') {
    return ContentService.createTextOutput(JSON.stringify(addLikeAction(params.eventId)))
        .setMimeType(ContentService.MimeType.JSON);
  }
  
  if (action === 'addComment') {
    return ContentService.createTextOutput(JSON.stringify(addCommentAction(params.eventId, params.comment)))
        .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * 2つのスプレッドシートからデータを取得して結合する
 */
function getEventsData() {
  // 1. イベントデータシートから読み込み
  const eventSs = SpreadsheetApp.openById(EVENT_SS_ID);
  const eventSheet = eventSs.getSheets()[0]; // 最初のシート（シート1）を取得
  const eventRows = eventSheet.getDataRange().getValues();
  
  // 2. ユーザーアクションシートから読み込み
  const actionSs = SpreadsheetApp.openById(ACTION_SS_ID);
  const actionSheet = actionSs.getSheets()[0]; // 最初のシート（シート1）を取得
  const actionRows = actionSheet.getDataRange().getValues();
  
  // CSV固有のヘッダー行（1〜4行目などのタイトルや空行）をスキップしてデータ行を特定する
  // 共有いただいたCSVの5行目が「イベントID,イベント名...」のヘッダーになっています
  const eventDataStartIndex = eventRows.findIndex(row => row[0] === "イベントID" || row[0] === "eventId");
  const actionDataStartIndex = actionRows.findIndex(row => row[0] === "イベントID" || row[0] === "eventId");
  
  const cleanEventRows = eventRows.slice(eventDataStartIndex + 1);
  const cleanActionRows = actionRows.slice(actionDataStartIndex + 1);
  
  // アクションデータの集計（いいね数とコメント一覧）
  const likesMap = {};
  const commentsMap = {};
  
  cleanActionRows.forEach(row => {
    const eId = String(row[0]).trim();
    if (!eId) return;
    
    // アクションデータの列構成に応じてマッピング
    // [イベントID, イベント名, いいね数, コメント] もしくはログ形式に対応
    const type = row[1]; 
    const content = row[3] || row[2]; // コメント内容
    
    // 新規登録用ロジック（もし既存シートがログ形式[eventId, type, content, timestamp]の場合）
    if (type === 'like') {
      likesMap[eId] = (likesMap[eId] || 0) + 1;
    } else if (type === 'comment' && content) {
      if (!commentsMap[eId]) commentsMap[eId] = [];
      commentsMap[eId].push({
        content: content,
        timestamp: row[3] instanceof Date ? Utilities.formatDate(row[3], "JST", "yyyy/MM/dd HH:mm") : "直近のコメント"
      });
    } else {
      // 共有いただいた初期シートの構造（いいね数、コメント列が固定）に対応するフォールバック
      const initialLikes = parseInt(row[2]) || 0;
      if (initialLikes > 0) {
        likesMap[eId] = (likesMap[eId] || 0) + initialLikes;
      }
      if (row[3] && String(row[3]).trim() !== "") {
        if (!commentsMap[eId]) commentsMap[eId] = [];
        commentsMap[eId].push({
          content: row[3],
          timestamp: "ログ"
        });
      }
    }
  });
  
  // イベントデータの整形
  const events = cleanEventRows.map(row => {
    const eId = String(row[0]).trim();
    if (!eId) return null;
    
    return {
      eventId: eId,
      title: row[1],
      date: row[2] instanceof Date ? Utilities.formatDate(row[2], "JST", "yyyy/MM/dd") : row[2],
      time: row[3],
      location: row[4],
      address: row[5],
      organizer: row[6],
      link: row[7],
      remark: row[8] || "", // 備考を追加
      likes: likesMap[eId] || 0,
      comments: commentsMap[eId] || []
    };
  }).filter(e => e !== null);
  
  // 日付の若い順番（昇順）にソート
  events.sort((a, b) => new Date(a.date) - new Date(b.date));
  
  return events;
}

/**
 * ユーザーアクションシートに「いいね」ログを追記
 */
function addLikeAction(eventId) {
  const ss = SpreadsheetApp.openById(ACTION_SS_ID);
  const actionSheet = ss.getSheets()[0];
  
  // ログ形式として末尾に追記 (イベントID, タイプ, 値/タイムスタンプ)
  actionSheet.appendRow([eventId, 'like', 1, new Date()]);
  
  // 最新のいいね数を集計して返す
  const rows = actionSheet.getDataRange().getValues();
  const eventDataStartIndex = rows.findIndex(row => row[0] === "イベントID" || row[0] === "eventId");
  const cleanRows = rows.slice(eventDataStartIndex + 1);
  
  let totalLikes = 0;
  cleanRows.forEach(row => {
    if (String(row[0]).trim() == String(eventId).trim()) {
      if (row[1] === 'like') {
        totalLikes += 1;
      } else {
        totalLikes += (parseInt(row[2]) || 0); // 初期値のいいね数も合算
      }
    }
  });
  
  return { success: true, newLikes: totalLikes };
}

/**
 * ユーザーアクションシートに「コメント」ログを追記
 */
function addCommentAction(eventId, commentContent) {
  if (!commentContent || commentContent.trim() === "") return { success: false };
  
  const ss = SpreadsheetApp.openById(ACTION_SS_ID);
  const actionSheet = ss.getSheets()[0];
  const now = new Date();
  
  // 末尾に追記 (イベントID, タイプ, 空白, コメント内容/タイムスタンプ)
  actionSheet.appendRow([eventId, 'comment', '', commentContent, now]);
  
  return { 
    success: true, 
    comment: {
      content: commentContent,
      timestamp: Utilities.formatDate(now, "JST", "yyyy/MM/dd HH:mm")
    }
  };
}
