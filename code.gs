// 【一般ユーザー用】Code.gs 

// 2つのスプレッドシートIDを個別に設定（ご自身のスプレッドシートIDを貼り付けてください）
const EVENT_SS_ID = "1DYB-TwuLtmk-C8aNE-mXOq1x6IXJsoAECo4SywafCgM";
const ACTION_SS_ID = "1Pz2fb5an1YAPlYK0-Y17RVVXfXn2CrBdY5NoQ_mVgSw";

/**
 * 外部（GitHub Pagesなど）からアクセスされた際の画面配信、または通信確認用
 */
function doGet() {
  return ContentService.createTextOutput("Amami Calendar Public API Working");
}

/**
 * フロントエンド（index.html）からの非同期リクエスト（fetch）を受け付けるAPI窓口
 */
function doPost(e) {
  const params = JSON.parse(e.postData.contents);
  const action = params.action;
  
  // 1. イベント一覧とユーザーアクション（いいね・コメント）を結合して取得
  if (action === 'getEvents') {
    return ContentService.createTextOutput(JSON.stringify(getEventsData()))
        .setMimeType(ContentService.MimeType.JSON);
  }
  
  // 2. いいねアクションの記録とリアルタイム集計
  if (action === 'addLike') {
    return ContentService.createTextOutput(JSON.stringify(addLikeAction(params.eventId)))
        .setMimeType(ContentService.MimeType.JSON);
  }
  
  // 3. コメントアクションの記録
  if (action === 'addComment') {
    return ContentService.createTextOutput(JSON.stringify(addCommentAction(params.eventId, params.comment)))
        .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * 2つのスプレッドシートからデータを取得・結合し、一般公開用のクリーンなデータを返す関数
 */
function getEventsData() {
  // ① イベントデータシートから読み込み
  const eventSs = SpreadsheetApp.openById(EVENT_SS_ID);
  const eventSheet = eventSs.getSheets()[0]; // 最初のシートを取得
  const eventRows = eventSheet.getDataRange().getValues();
  
  // ② ユーザーアクションシートから読み込み
  const actionSs = SpreadsheetApp.openById(ACTION_SS_ID);
  const actionSheet = actionSs.getSheets()[0]; // 最初のシートを取得
  const actionRows = actionSheet.getDataRange().getValues();
  
  // CSV等のヘッダー行（見出し行）の位置を特定してスキップ
  const eventDataStartIndex = eventRows.findIndex(row => row[0] === "イベントID" || row[0] === "eventId");
  const actionDataStartIndex = actionRows.findIndex(row => row[0] === "イベントID" || row[0] === "eventId");
  
  const cleanEventRows = eventRows.slice(eventDataStartIndex + 1);
  const cleanActionRows = actionRows.slice(actionDataStartIndex + 1);
  
  // いいね数とコメント一覧を格納するマップを準備
  const likesMap = {};
  const commentsMap = {};
  
  // アクションログ（蓄積データ）の集計処理
  cleanActionRows.forEach(row => {
    const eId = String(row[0]).trim();
    if (!eId) return;
    
    const type = row[1];
    const content = row[3] || row[2]; // コメント内容の取得
    
    if (type === 'like') {
      likesMap[eId] = (likesMap[eId] || 0) + 1;
    } else if (type === 'comment' && content) {
      if (!commentsMap[eId]) commentsMap[eId] = [];
      // 要件定義「新しい順に上から表示」のため、配列の先頭に追加（unshift）
      commentsMap[eId].push({
        content: content,
        timestamp: row[4] instanceof Date ? Utilities.formatDate(row[4], "JST", "yyyy/MM/dd HH:mm") : "直近のコメント"
      });
    } else {
      // 初期シートデータ（いいね数値が固定で入っている場合）の互換処理
      const initialLikes = parseInt(row[2]) || 0;
      if (initialLikes > 0) {
        likesMap[eId] = (likesMap[eId] || 0) + initialLikes;
      }
      if (row[3] && String(row[3]).trim() !== "" && !row[1]) {
        if (!commentsMap[eId]) commentsMap[eId] = [];
        commentsMap[eId].push({
          content: row[3],
          timestamp: "ログデータ"
        });
      }
    }
  });
  
  // イベントデータの整形と非表示（hide）フィルター処理
  const events = cleanEventRows.map(row => {
    const eId = String(row[0]).trim();
    if (!eId) return null;
    
    // J列（10列目）のステータスが "hide" の場合は、一般画面には一切配信しない
    if (row[9] === "hide") return null;
    
    return {
      eventId: eId,
      title: row[1],
      date: row[2] instanceof Date ? Utilities.formatDate(row[2], "JST", "yyyy/MM/dd") : row[2],
      time: row[3],
      location: row[4],
      address: row[5],
      organizer: row[6],
      link: row[7],
      remark: row[8] || "", // 備考
      likes: likesMap[eId] || 0,
      comments: commentsMap[eId] || []
    };
  }).filter(e => e !== null);
  
  // 要件定義「日付の若い順番に上から表示」のため、昇順でソート
  events.sort((a, b) => new Date(a.date) - new Date(b.date));
  
  return events;
}

/**
 * ユーザーアクションシートに「いいね」を追記し、最新の合算数値をリアルタイムで返す関数
 */
function addLikeAction(eventId) {
  const ss = SpreadsheetApp.openById(ACTION_SS_ID);
  const actionSheet = ss.getSheets()[0];
  
  // 末尾に行を追加 [イベントID, タイプ, 値, メッセージ, タイムスタンプ]
  actionSheet.appendRow([eventId, 'like', 1, '', new Date()]);
  
  // 追加後の最新のいいね合計数をリアルタイムに再集計
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
 * ユーザーアクションシートに「コメント」ログを追記し、フロントに返す関数
 */
function addCommentAction(eventId, commentContent) {
  if (!commentContent || commentContent.trim() === "") return { success: false };
  
  const ss = SpreadsheetApp.openById(ACTION_SS_ID);
  const actionSheet = ss.getSheets()[0];
  const now = new Date();
  
  // 末尾に追記 [イベントID, タイプ, 値, コメント内容, タイムスタンプ]
  actionSheet.appendRow([eventId, 'comment', '', commentContent, now]);
  
  return { 
    success: true, 
    comment: {
      content: commentContent,
      timestamp: Utilities.formatDate(now, "JST", "yyyy/MM/dd HH:mm")
    }
  };
}
