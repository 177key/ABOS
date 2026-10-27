
/*************************************************
 * Discord通知キュー
 *************************************************/

const NOTIFY_QUEUE_SHEET = "通知キュー";

/**
 * 通知キューシートを作成
 */
function createNotificationQueueSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  if (ss.getSheetByName(NOTIFY_QUEUE_SHEET)) {
    console.log("通知キューは既に存在します");
    return;
  }

  const sheet = ss.insertSheet(NOTIFY_QUEUE_SHEET);

  const headers = [
    "通知ID",
    "通知先",
    "通知内容",
    "状態",
    "作成日時",
    "送信日時",
    "DiscordメッセージID",
    "エラー"
  ];

  sheet.getRange(1, 1, 1, headers.length)
    .setValues([headers])
    .setFontWeight("bold")
    .setBackground("#d9ead3");

  sheet.setFrozenRows(1);

  // 通知ID、通知先、JSON、DiscordメッセージID
  sheet.getRange("A2:C").setNumberFormat("@");
  sheet.getRange("G2:G").setNumberFormat("@");

  sheet.setColumnWidth(1, 300);
  sheet.setColumnWidth(2, 160);
  sheet.setColumnWidth(3, 450);
  sheet.setColumnWidth(4, 110);
  sheet.setColumnWidth(5, 170);
  sheet.setColumnWidth(6, 170);
  sheet.setColumnWidth(7, 220);
  sheet.setColumnWidth(8, 300);

  console.log("通知キューを作成しました");
}


/**
 * 通知をキューに登録する
 *
 * destination: "総合", "U001" など
 * message: { content: "...", embeds: [...] }
 */
function enqueueDiscordNotification(destination, message) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(NOTIFY_QUEUE_SHEET);

  if (!sheet) {
    throw new Error("通知キューシートがありません");
  }

  if (!destination || typeof destination !== "string") {
    throw new Error("通知先が不正です");
  }

  if (!message || typeof message !== "object" ||
      !(message.content || 
        (Array.isArray(message.embeds) &&
         message.embeds.length > 0))) {
    throw new Error("通知内容が空です");
  }

  const json = JSON.stringify(message);

  if (json.length > 40000) {
    throw new Error("通知内容が大きすぎます");
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const notificationId = Utilities.getUuid();

    sheet.appendRow([
      notificationId,
      destination.trim(),
      json,
      "未送信",
      new Date(),
      "",
      "",
      ""
    ]);

    SpreadsheetApp.flush();

    console.log(
      `通知登録：${notificationId} → ${destination}`
    );

    return notificationId;

  } finally {
    lock.releaseLock();
  }
}
