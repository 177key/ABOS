
/**
 * Discord BotからのPOSTリクエストを受け付ける。
 * 現段階では疎通テストのみ。
 */

function doPost(e) {
  try {
    const data = JSON.parse(
      e && e.postData ? e.postData.contents : "{}"
    );

    // APIトークン認証
    const token = PropertiesService
      .getScriptProperties()
      .getProperty("DISCORD_PURCHASE_API_TOKEN");

    if (!token || data.token !== token) {
      return purchaseApiJson_({
        ok: false,
        error: "UNAUTHORIZED"
      });
    }

    // 疎通確認
    if (data.action === "ping") {
      return purchaseApiJson_({
        ok: true,
        action: "ping",
        message: "通信成功"
      });
    }

    // 購入内容の検証
    if (data.action === "validate") {
      return purchaseApiJson_(
        validateBetApi_(data)
      );
    }

    if (data.action === "purchase") {
      return purchaseApiJson_(
        purchaseBetApi_(data)
      );
    }

    return purchaseApiJson_({
      ok: false,
      error: "UNKNOWN_ACTION"
    });

  } catch (error) {
    return purchaseApiJson_({
      ok: false,
      error: String(error.message)
    });
  }
}

// 購入点数・必要ptを計算する
function validateBetApi_(data) {
  const betType = String(data.betType || "").trim();
  const method = String(data.method || "").trim();
  const selection = String(data.selection || "").trim();
  const pointPerBet = Number(data.pointPerBet);

  const multi =
    data.multi === true ||
    String(data.multi).toUpperCase() === "TRUE";

  // 最低限の入力検証
  if (!betType || !method || !selection) {
    throw new Error("券種・方式・買い目を指定してください。");
  }

  if (
    !Number.isSafeInteger(pointPerBet) ||
    pointPerBet <= 0 ||
    pointPerBet % 100 !== 0
  ) {
    throw new Error(
      "購入ptは100pt単位の正の整数で指定してください。"
    );
  }

  // 既存の購入処理と同じ展開ロジックを使用
  const tickets = expandBets_(
    betType,
    method,
    multi,
    selection
  );

  if (!Array.isArray(tickets) || tickets.length === 0) {
    throw new Error("有効な買い目がありません。");
  }

  const ticketCount = tickets.length;
  const totalPurchase = ticketCount * pointPerBet;

  if (!Number.isSafeInteger(totalPurchase)) {
    throw new Error("購入ptの計算結果が大きすぎます。");
  }

  return {
    ok: true,
    betType: betType,
    method: method,
    multi: multi,
    selection: selection,
    pointPerBet: pointPerBet,
    ticketCount: ticketCount,
    totalPurchase: totalPurchase
  };
}

// JSON形式のレスポンス
function purchaseApiJson_(object) {
  return ContentService
    .createTextOutput(JSON.stringify(object))
    .setMimeType(ContentService.MimeType.JSON);
}


/**
 * Discord経由の購入確定API
 *
 * 想定する既存関数：
 * purchaseBetWithNotification(
 *   participantId, venue, raceNumber,
 *   betType, method, multi, selection, pointPerBet
 * )
 */
function purchaseBetApi_(data) {
  const requestId = String(data.requestId || "").trim();

  if (!/^[A-Za-z0-9-]{20,64}$/.test(requestId)) {
    return {
      ok: false,
      error: "リクエストIDが不正です。"
    };
  }

  const participantId = String(data.participantId || "").trim();
  const venue = String(data.venue || "").trim();
  const raceNumber = Number(data.raceNumber);
  const betType = String(data.betType || "").trim();
  const method = String(data.method || "").trim();
  const selection = String(data.selection || "").trim();
  const pointPerBet = Number(data.pointPerBet);

  const multi = data.multi === true ||
    String(data.multi).toUpperCase() === "TRUE";

  if (
    !participantId || !venue || !betType ||
    !method || !selection ||
    !Number.isSafeInteger(raceNumber) ||
    raceNumber < 1 || raceNumber > 12 ||
    !Number.isSafeInteger(pointPerBet) ||
    pointPerBet <= 0 ||
    pointPerBet % 100 !== 0
  ) {
    return {
      ok: false,
      error: "購入内容が不正です。"
    };
  }

  // リクエストIDを先に確保して二重処理を防ぐ
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  let sheet;
  let historyRow;

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    sheet = ss.getSheetByName("購入API履歴");

    if (!sheet) {
      sheet = ss.insertSheet("購入API履歴");
      sheet.appendRow([
        "リクエストID",
        "参加者ID",
        "状態",
        "受付日時",
        "結果"
      ]);
    }

    const lastRow = sheet.getLastRow();

    if (lastRow >= 2) {
      const ids = sheet.getRange(
        2, 1, lastRow - 1, 1
      ).getDisplayValues().flat();

      if (ids.includes(requestId)) {
        return {
          ok: false,
          duplicate: true,
          error: "この購入リクエストは既に受付済みです。"
        };
      }
    }

    historyRow = sheet.getLastRow() + 1;

    sheet.appendRow([
      requestId,
      participantId,
      "処理中",
      new Date(),
      ""
    ]);

    SpreadsheetApp.flush();

  } finally {
    lock.releaseLock();
  }

  try {
    // 既存の購入処理をそのまま利用
    const result = purchaseBetWithNotification(
      participantId,
      venue,
      raceNumber,
      betType,
      method,
      multi,
      selection,
      pointPerBet
    );

    sheet.getRange(historyRow, 3).setValue("成功");
    sheet.getRange(historyRow, 5).setValue("購入処理完了");

    return {
      ok: true,
      requestId: requestId,
      message: "馬券の購入が確定しました。",
      purchaseResult: result == null ? null : result
    };

  } catch (error) {
    // エラー発生時は自動的に再購入しない
    const errorMessage = String(error.message || error);

    sheet.getRange(historyRow, 3)
      .setValue("要確認");

    sheet.getRange(historyRow, 5)
      .setValue(errorMessage);

    return {
      ok: false,
      requestId: requestId,
      error: errorMessage
    };
  }
}
