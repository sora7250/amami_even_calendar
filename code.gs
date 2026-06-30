function doPost(e) {
  const params = JSON.parse(e.postData.contents);
  const action = params.action;
  
  // --- 一般ユーザー用機能（既存） ---
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

  // --- 管理者機能：認証チェック ---
  if (action === 'verifyAdmin') {
    const isCorrect = (params.password === ADMIN_PASSWORD);
    return ContentService.createTextOutput(JSON.stringify({ success: isCorrect }))
        .setMimeType(ContentService.MimeType.JSON);
  }

  // --- 【不具合修正】管理者機能：管理画面専用の全イベントデータ取得（hideも含めて返す） ---
  if (action === 'getAdminEvents') {
    if (params.password !== ADMIN_PASSWORD) return ContentService.createTextOutput(JSON.stringify({ success: false }));
    return ContentService.createTextOutput(JSON.stringify(getAdminEventsData()))
        .setMimeType(ContentService.MimeType.JSON);
  }

  // --- 管理者機能：新規イベントの直接追加 ---
  if (action === 'addEvent') {
    if (params.password !== ADMIN_PASSWORD) return ContentService.createTextOutput(JSON.stringify({ success: false, msg: "認証エラー" }));
    
    const ss = SpreadsheetApp.openById(EVENT_SS_ID);
    const sheet = ss.getSheets()[0];
    const newEventId = "ev_" + new Date().getTime();
    
    sheet.appendRow([
      newEventId,
      params.title,
      params.date,
      params.time,
      params.location,
      params.address,
      params.organizer,
      params.link,
      params.remark,
      "show"
    ]);
    
    return ContentService.createTextOutput(JSON.stringify({ success: true }))
        .setMimeType(ContentService.MimeType.JSON);
  }

  // --- 管理者機能：イベントの非表示 ---
  if (action === 'hideEvent') {
    if (params.password !== ADMIN_PASSWORD) return ContentService.createTextOutput(JSON.stringify({ success: false, msg: "認証エラー" }));
    
    const ss = SpreadsheetApp.openById(EVENT_SS_ID);
    const sheet = ss.getSheets()[0];
    const rows = sheet.getDataRange().getValues();
    
    for (let i = 0; i < rows.length; i++) {
      if (String(rows[i][0]).trim() === String(params.eventId).trim()) {
        sheet.getRange(i + 1, 10).setValue("hide"); 
        break;
      }
    }
    return ContentService.createTextOutput(JSON.stringify({ success: true }))
        .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * 【新規追加】管理画面専用のデータ取得関数
 * サイト上で「非表示(hide)」にしたものも含めてリストアップし、管理できるようにします
 */
function getAdminEventsData() {
  const eventSs = SpreadsheetApp.openById(EVENT_SS_ID);
  const eventSheet = eventSs.getSheets()[0];
  const eventRows = eventSheet.getDataRange().getValues();
  
  const eventDataStartIndex = eventRows.findIndex(row => row[0] === "イベントID" || row[0] === "eventId");
  const cleanEventRows = eventRows.slice(eventDataStartIndex + 1);
  
  const events = cleanEventRows.map(row => {
    const eId = String(row[0]).trim();
    if (!eId) return null;
    
    return {
      eventId: eId,
      title: row[1],
      date: row[2] instanceof Date ? Utilities.formatDate(row[2], "JST", "yyyy/MM/dd") : row[2],
      status: row[9] || "show" // 表示ステータス(showかhide)もフロントに渡す
    };
  }).filter(e => e !== null);
  
  // 日付の若い順にソート
  events.sort((a, b) => new Date(a.date) - new Date(b.date));
  return events;
}
