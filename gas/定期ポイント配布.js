
/*************************************************
 * 毎週月曜日の定期ポイント配布
 *************************************************/

const WEEKLY_PT_CONFIG = {
  pointSheet: "ポイント管理",
  logSheet: "定期配布履歴",
  timezone: "Asia/Tokyo"
};


/**
 * YYYY-MM-DD形式の日付を検証
 */
function parseWeeklyDate_(dateStr) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(dateStr)) {
    throw new Error("日付の形式が不正です：" + dateStr);
  }

  const [y, m, d] = dateStr.split("-").map(Number);
  const date = new Date(Date.UTC(y, m - 1, d));

  if (
    date.getUTCFullYear() !== y ||
    date.getUTCMonth() + 1 !== m ||
    date.getUTCDate() !== d
  ) {
    throw new Error("存在しない日付です：" + dateStr);
  }

  return date;
}


/**
 * 指定日の週の月曜日を取得
 */
function getWeeklyMonday_(dateStr) {
  const date = parseWeeklyDate_(dateStr);
  const day = date.getUTCDay();

  const offset = (day + 6) % 7;

  date.setUTCDate(date.getUTCDate() - offset);

  return Utilities.formatDate(
    date, "UTC", "yyyy-MM-dd"
  );
}


/**
 * 配布対象か判定し、現在の残高を取得
 * この関数では書き込まない
 */
function getWeeklyPointPlan_() {
  const config = WEEKLY_PT_CONFIG;
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const today = Utilities.formatDate(
    new Date(), config.timezone, "yyyy-MM-dd"
  );

  const thisMonday = getWeeklyMonday_(today);

  // 月曜日以外は配布しない
  if (today !== thisMonday) {
    return {
      eligible: false,
      reason: "今日は月曜日ではありません",
      date: thisMonday
    };
  }

    // 現在開催中の大会
  const tournament = getCurrentTournament();

  // 大会ごとの月曜配布pt
  const amount = tournament.weeklyPt;

  const pointSheet = ss.getSheetByName(config.pointSheet);
  const logSheet = ss.getSheetByName(config.logSheet);

  if (!pointSheet || !logSheet) {
    throw new Error("必要なシートが見つかりません");
  }

  // 履歴を確認する
  const logLastRow = logSheet.getLastRow();

  if (logLastRow >= 2) {
    const logs = logSheet.getRange(
      2, 1, logLastRow - 1, 5
    ).getValues();

    for (const row of logs) {
      const rawDate = row[0];

      const logDate = rawDate instanceof Date
        ? Utilities.formatDate(
            rawDate, config.timezone, "yyyy-MM-dd"
          )
        : String(rawDate).trim().replace(/\//g, "-");

      if (logDate === thisMonday) {
        return {
          eligible: false,
          reason: "この週の配布記録が既に存在します（状態：" +
            String(row[4]) + "）",
          date: thisMonday
        };
      }
    }
  }

  // ポイント管理シートを読み込む
  const lastRow = pointSheet.getLastRow();

  if (lastRow < 2) {
    throw new Error("参加者が登録されていません");
  }

  const data = pointSheet.getRange(
    2, 1, lastRow - 1, 2
  ).getValues();

  const participants = [];
  const ids = new Set();

  for (let i = 0; i < data.length; i++) {
    const id = String(data[i][0]).trim();

    if (!id) continue;

    if (ids.has(id)) {
      throw new Error("参加者IDが重複しています：" + id);
    }

    ids.add(id);

    const balance = Number(data[i][1]);

    if (
      data[i][1] === "" ||
      !Number.isSafeInteger(balance) ||
      balance < 0 ||
      !Number.isSafeInteger(balance + amount)
    ) {
      throw new Error(
        `${id}の通常ptが不正です`
      );
    }

    participants.push({
      id,
      rowNumber: i + 2,
      oldBalance: balance,
      newBalance: balance + amount
    });
  }

  if (participants.length === 0) {
    throw new Error("配布対象の参加者がいません");
  }

  return {
    eligible: true,
    date: thisMonday,
    amount: amount,
    count: participants.length,
    total: config.amount * participants.length,
    participants,
    pointSheet,
    logSheet,
    tournamentId: tournament.tournamentId,
  };
}


/**
 * 月曜日の定期配布を実行する
 */
function distributeWeeklyPoints() {
  const lock = LockService.getScriptLock();
  let locked = false;

  try {
    lock.waitLock(30000);
    locked = true;

    const plan = getWeeklyPointPlan_();

    if (!plan.eligible) {
      const result = {
        success: true,
        distributed: false,
        reason: plan.reason,
        date: plan.date
      };

      console.log(result);
      return result;
    }

    const { pointSheet, logSheet, participants } = plan;

    const logRow = logSheet.getLastRow() + 1;

    // 履歴は文字列で保存
    logSheet.getRange(logRow, 1)
      .setNumberFormat("@");

    // 配布開始を先に記録
    // 途中終了した場合は自動再実行を防ぐ
    logSheet.getRange(logRow, 1, 1, 6).setValues([[
      plan.date,
      new Date(),
      plan.count,
      plan.total,
      "処理中",
      plan.tournamentId
    ]]);

    SpreadsheetApp.flush();

    const modified = [];

    try {
      for (const p of participants) {
        // 書込前に復元情報を登録
        modified.push(p);

        pointSheet.getRange(p.rowNumber, 2)
          .setValue(p.newBalance);
      }

      SpreadsheetApp.flush();

      // すべて成功して初めて完了にする
      logSheet.getRange(logRow, 5)
        .setValue("完了");

      SpreadsheetApp.flush();

    } catch (e) {
      let restored = true;

      for (let i = modified.length - 1; i >= 0; i--) {
        const p = modified[i];

        try {
          pointSheet.getRange(p.rowNumber, 2)
            .setValue(p.oldBalance);
        } catch (restoreError) {
          restored = false;
          console.error(restoreError.message);
        }
      }

      try {
        SpreadsheetApp.flush();
      } catch (flushError) {
        restored = false;
      }

      try {
        logSheet.getRange(logRow, 5).setValue(
          restored ? "失敗・復元済" : "要確認"
        );
        SpreadsheetApp.flush();
      } catch (logError) {
        console.error(logError.message);
      }

      throw new Error(
        "配布処理に失敗しました：" + e.message
      );
    }

    const result = {
      success: true,
      distributed: true,
      date: plan.date,
      count: plan.count,
      amountPerPerson: plan.amount,
      total: plan.total
    };

    console.log(JSON.stringify(result, null, 2));

    return result;

  } catch (e) {
    console.error("定期配布エラー：" + e.message);

    return {
      success: false,
      error: e.message
    };

  } finally {
    if (locked) lock.releaseLock();
  }
}


/**
 * 毎週月曜日に実行するトリガーを登録
 * 初回のみ手動実行
 */
function setupWeeklyPointTrigger() {
  const functionName = "distributeWeeklyPoints";

  // 同じ関数のトリガー重複を防ぐ
  for (const trigger of ScriptApp.getProjectTriggers()) {
    if (trigger.getHandlerFunction() === functionName) {
      ScriptApp.deleteTrigger(trigger);
    }
  }

  ScriptApp.newTrigger(functionName)
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(8)
    .inTimezone(WEEKLY_PT_CONFIG.timezone)
    .create();

  console.log("毎週月曜日の定期配布トリガーを登録しました");
}


/**
 * 定期配布履歴シートを作成する
 * 既に存在する場合は何もしない
 */
function createWeeklyDistributionSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = "定期配布履歴";

  // 既存シートの確認
  const existing = ss.getSheetByName(sheetName);

  if (existing) {
    console.log(
      `「${sheetName}」は既に存在します。`
    );
    return;
  }

  // 新規シート作成
  const sheet = ss.insertSheet(sheetName);

  const headers = [
    "配布対象日",
    "実行日時",
    "配布人数",
    "配布総pt",
    "状態",
    "大会ID"
  ];

  // 見出しを設定
  sheet.getRange(1, 1, 1, headers.length)
    .setValues([headers]);

  // 見出しのデザイン
  sheet.getRange(1, 1, 1, headers.length)
    .setBackground("#d9ead3")
    .setFontWeight("bold");

  // 見出し行を固定
  sheet.setFrozenRows(1);

  // A列は日付文字列として保存
  sheet.getRange("A2:A").setNumberFormat("@");

  // B列は日時
  sheet.getRange("B2:B")
    .setNumberFormat("yyyy/MM/dd HH:mm:ss");

  // 列幅を調整
  sheet.setColumnWidth(1, 130);
  sheet.setColumnWidth(2, 180);
  sheet.setColumnWidth(3, 110);
  sheet.setColumnWidth(4, 130);
  sheet.setColumnWidth(5, 130);
  sheet.setColumnWidth(6, 110);

  console.log(
    `「${sheetName}」を作成しました！`
  );
}
