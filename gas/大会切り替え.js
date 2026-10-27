
/*************************************************
 * 大会終了・次大会初期化
 *
 * 大会情報：
 * A 大会ID / B 大会名 / C 終了日
 * D 初期pt / E 月曜配布pt / F ボーナスpt
 * G 終了処理済 / H 初期化済
 *
 * 大会結果：
 * A 大会ID / B 参加者ID / C 参加有無
 * D 終了時通常pt / E 終了時利益pt
 * F 付与前ボーナスpt / G 今回ボーナスpt
 * H 最終通常pt / I 最終利益pt
 * J 最終ボーナスpt / K 最終合計pt / L 順位
 *************************************************/

const CHANGE_CONFIG = {
  tournamentSheet: "大会情報",
  pointSheet: "ポイント管理",
  historySheet: "馬券購入履歴",
  bonusSheet: "ボーナス管理",
  resultSheet: "大会結果",
  timezone: "Asia/Tokyo"
};


/** 日付をyyyyMMddに統一 */
function changeDateKey_(value) {
  if (value instanceof Date) {
    if (isNaN(value.getTime())) {
      throw new Error("不正な日付です");
    }
    return Utilities.formatDate(
      value, CHANGE_CONFIG.timezone, "yyyyMMdd"
    );
  }

  const s = String(value).trim();
  const m = s.match(
    /^(\d{4})[/-](\d{1,2})[/-](\d{1,2})$/
  );

  if (!m) throw new Error("日付形式が不正です：" + s);

  const y = Number(m[1]);
  const month = Number(m[2]);
  const day = Number(m[3]);
  const d = new Date(Date.UTC(y, month - 1, day));

  if (
    d.getUTCFullYear() !== y ||
    d.getUTCMonth() + 1 !== month ||
    d.getUTCDate() !== day
  ) {
    throw new Error("存在しない日付です：" + s);
  }

  return String(y).padStart(4, "0") +
    String(month).padStart(2, "0") +
    String(day).padStart(2, "0");
}


/** 必須シート */
function changeSheet_(ss, name) {
  const sheet = ss.getSheetByName(name);
  if (!sheet) {
    throw new Error(`シート「${name}」がありません`);
  }
  return sheet;
}


/** TRUE/FALSE/空欄を解釈 */
function changeFlag_(value) {
  return value === true ||
    String(value).trim().toUpperCase() === "TRUE";
}


/** ポイントの安全な数値変換 */
function changePoints_(value, label) {
  if (value === "" || value === null) {
    throw new Error(label + "が空欄です");
  }
  const n = Number(value);

  if (!Number.isSafeInteger(n) || n < 0) {
    throw new Error(label + "が不正です");
  }
  return n;
}


/** 大会情報を時系列順に取得 */
function readChangeTournaments_(sheet) {
  const last = sheet.getLastRow();
  if (last < 2) throw new Error("大会情報がありません");

  const rows = sheet.getRange(
    2, 1, last - 1, 8
  ).getValues();

  const tournaments = [];
  let previousEnd = "";
  const ids = new Set();

  rows.forEach((r, i) => {
    const id = String(r[0]).trim();
    if (!id) return;

    if (ids.has(id)) {
      throw new Error("大会ID重複：" + id);
    }
    ids.add(id);

    const end = changeDateKey_(r[2]);

    if (previousEnd && end <= previousEnd) {
      throw new Error("大会終了日が時系列順ではありません");
    }

    previousEnd = end;

    tournaments.push({
      id,
      name: String(r[1]).trim(),
      end,
      initialPt: changePoints_(r[3], id + " 初期pt"),
      weeklyPt: changePoints_(r[4], id + " 月曜pt"),
      bonusPt: changePoints_(r[5], id + " ボーナスpt"),
      finished: changeFlag_(r[6]),
      initialized: changeFlag_(r[7]),
      rowNumber: i + 2
    });
  });

  return tournaments;
}


/** 結果保存用シートを自動作成 */
function ensureChangeResultSheet_(ss) {
  const name = CHANGE_CONFIG.resultSheet;
  let sheet = ss.getSheetByName(name);

  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.getRange(1, 1, 1, 12).setValues([[
      "大会ID", "参加者ID", "参加有無",
      "終了時通常pt", "終了時利益pt",
      "付与前ボーナスpt", "今回ボーナスpt",
      "最終通常pt", "最終利益pt",
      "最終ボーナスpt", "最終合計pt", "順位"
    ]]);
    sheet.getRange(1, 1, 1, 12)
      .setBackground("#d9ead3")
      .setFontWeight("bold");
    sheet.setFrozenRows(1);
    sheet.autoResizeColumns(1, 12);
  }

  return sheet;
}


/** 馬券購入履歴から大会参加者を判定 */
function getChangeEntrants_(ss, tournamentId) {
  const sheet = changeSheet_(
    ss, CHANGE_CONFIG.historySheet
  );
  const entrants = new Set();

  if (sheet.getLastRow() < 2) return entrants;

  const rows = sheet.getRange(
    2, 1, sheet.getLastRow() - 1, 11
  ).getValues();

  for (const r of rows) {
    if (String(r[1]).trim() !== tournamentId) continue;

    if (!changeFlag_(r[10])) {
      throw new Error(
        "未精算馬券が残っています：" + String(r[0])
      );
    }

    entrants.add(String(r[2]).trim());
  }

  return entrants;
}


/** 18部門のボーナス保持者を集計 */
function getChangeBonusCounts_(ss, entrants) {
  const sheet = changeSheet_(
    ss, CHANGE_CONFIG.bonusSheet
  );

  const data = sheet.getRange(1, 1, 23, 4)
    .getValues();

  const counts = new Map();

  const startRows = [3, 15];

  for (const start of startRows) {
    for (let i = 0; i < 9; i++) {
      const row = data[start + i - 1];
      const ids = String(row[1] ?? "").trim();

      if (!ids) continue;

      // 同一部門での重複受賞は1回
      const uniqueIds = new Set(
        ids.split("\n").map(x => x.trim()).filter(Boolean)
      );

      for (const id of uniqueIds) {
        if (!entrants.has(id)) {
          throw new Error(
            "未参加者がボーナス記録にあります：" + id
          );
        }

        counts.set(id, (counts.get(id) || 0) + 1);
      }
    }
  }

  return counts;
}


/** ポイント管理の参加者と残高を取得 */
function getChangeAccounts_(ss) {
  const sheet = changeSheet_(
    ss, CHANGE_CONFIG.pointSheet
  );
  const last = sheet.getLastRow();
  const accounts = [];
  const ids = new Set();

  if (last < 2) throw new Error("参加者がいません");

  const data = sheet.getRange(
    2, 1, last - 1, 4
  ).getValues();

  data.forEach((r, i) => {
    const id = String(r[0]).trim();
    if (!id) return;

    if (ids.has(id)) {
      throw new Error("参加者IDが重複：" + id);
    }
    ids.add(id);

    accounts.push({
      id,
      rowNumber: i + 2,
      normal: changePoints_(r[1], id + " 通常pt"),
      profit: changePoints_(r[2], id + " 利益pt"),
      bonus: changePoints_(r[3], id + " ボーナスpt")
    });
  });

  return accounts;
}


/** 大会結果のスナップショットを作成 */
function buildChangeResults_(
  accounts, entrants, bonusCounts, tournament
) {
  const results = accounts.map(a => {
    const joined = entrants.has(a.id);
    const award = joined
      ? (bonusCounts.get(a.id) || 0) *
        tournament.bonusPt
      : 0;

    const finalBonus = a.bonus + award;
    const total = a.normal + a.profit + finalBonus;

    if (
      !Number.isSafeInteger(finalBonus) ||
      !Number.isSafeInteger(total)
    ) {
      throw new Error(a.id + "の最終ptが大きすぎます");
    }

    return [
      tournament.id,
      a.id,
      joined,
      a.normal,
      a.profit,
      a.bonus,
      award,
      a.normal,
      a.profit,
      finalBonus,
      total,
      ""
    ];
  });

  const participating = results
    .filter(r => r[2] === true)
    .sort((a, b) => b[10] - a[10]);

  let lastScore = null;
  let rank = 0;

  participating.forEach((r, i) => {
    if (r[10] !== lastScore) {
      rank = i + 1;
      lastScore = r[10];
    }
    r[11] = rank;
  });

  return results;
}


/**
 * 既存の大会結果を読み取る。
 * 保存途中のデータを検出した場合は自動再開しない。
 */
function readSavedChangeResults_(
  sheet, tournamentId, accounts
) {
  if (sheet.getLastRow() < 2) return null;

  const data = sheet.getRange(
    2, 1, sheet.getLastRow() - 1, 12
  ).getValues();

  const rows = data.filter(
    r => String(r[0]).trim() === tournamentId
  );

  if (rows.length === 0) return null;

  const expected = new Set(accounts.map(a => a.id));
  const actual = new Set();

  for (const r of rows) {
    const id = String(r[1]).trim();

    if (!expected.has(id) || actual.has(id)) {
      throw new Error(
        "保存済み大会結果と参加者一覧が一致しません"
      );
    }

    actual.add(id);

    for (const index of [3, 4, 5, 6, 7, 8, 9, 10]) {
      changePoints_(r[index], id + " 保存済みpt");
    }
  }

  if (actual.size !== expected.size) {
    throw new Error(
      "大会結果の保存が不完全です。手動確認してください"
    );
  }

  return rows;
}


/*************************************************
 * 1. 大会終了処理
 *
 * ・未精算チェック
 * ・各部門のボーナス計算
 * ・大会結果保存
 * ・ボーナスpt付与
 * ・G列をTRUE
 *************************************************/
function finishTournament_(ss, tournament) {
  const tournamentSheet = changeSheet_(
    ss, CHANGE_CONFIG.tournamentSheet
  );
  const resultSheet = ensureChangeResultSheet_(ss);
  const pointSheet = changeSheet_(
    ss, CHANGE_CONFIG.pointSheet
  );

  const accounts = getChangeAccounts_(ss);
  const entrants = getChangeEntrants_(
    ss, tournament.id
  );

  for (const id of entrants) {
    if (!accounts.some(a => a.id === id)) {
      throw new Error(
        "ポイント管理に参加者がいません：" + id
      );
    }
  }

  let results = readSavedChangeResults_(
    resultSheet, tournament.id, accounts
  );

  if (!results) {
    const bonusCounts = getChangeBonusCounts_(
      ss, entrants
    );

    results = buildChangeResults_(
      accounts, entrants, bonusCounts, tournament
    );

    // ポイントを変更する前に最終結果を保存
    const nextRow = resultSheet.getLastRow() + 1;
    resultSheet.getRange(
      nextRow, 1, results.length, 12
    ).setValues(results);

    SpreadsheetApp.flush();
  }

  // 保存済みの最終値を使い、絶対値で更新する。
  // 再実行時にもボーナスを二重加算しない。
  const resultById = new Map(
    results.map(r => [String(r[1]).trim(), r])
  );

  for (const a of accounts) {
    const r = resultById.get(a.id);
    pointSheet.getRange(a.rowNumber, 2, 1, 3)
      .setValues([[r[7], r[8], r[9]]]);
  }

  SpreadsheetApp.flush();

  // 全更新後に終了処理済をTRUEにする
  tournamentSheet.getRange(
    tournament.rowNumber, 7
  ).setValue(true);

  SpreadsheetApp.flush();

  return {
    tournamentId: tournament.id,
    participantCount: results.filter(r => r[2]).length,
    results: results.filter(r => r[2])
      .sort((a, b) => a[11] - b[11])
      .map(r => ({
        participantId: r[1],
        rank: r[11],
        normalPt: r[7],
        profitPt: r[8],
        bonusPt: r[9],
        awardedBonusPt: r[6],
        totalPt: r[10]
      }))
  };
}


/*************************************************
 * 2. 次大会の初期化
 *
 * ・全登録メンバーのポイントをリセット
 * ・ボーナス最高記録をクリア
 * ・H列をTRUE
 *************************************************/
function initializeTournament_(ss, tournament) {
  const tournamentSheet = changeSheet_(
    ss, CHANGE_CONFIG.tournamentSheet
  );
  const pointSheet = changeSheet_(
    ss, CHANGE_CONFIG.pointSheet
  );
  const bonusSheet = changeSheet_(
    ss, CHANGE_CONFIG.bonusSheet
  );

  // 初期化済みなら何もしない
  if (changeFlag_(
    tournamentSheet.getRange(
      tournament.rowNumber, 8
    ).getValue()
  )) {
    return false;
  }

  const accounts = getChangeAccounts_(ss);

  // 通常pt=初期pt、利益pt=0、ボーナスpt=0
  // E列の合計ptの数式は変更しない
  for (const a of accounts) {
    pointSheet.getRange(a.rowNumber, 2, 1, 3)
      .setValues([[
        tournament.initialPt, 0, 0
      ]]);
  }

  // 最高オッズ・最高利益ptをクリア
  // 券種名のA列は残す
  bonusSheet.getRange(3, 2, 9, 3).clearContent();
  bonusSheet.getRange(15, 2, 9, 3).clearContent();

  SpreadsheetApp.flush();

  tournamentSheet.getRange(
    tournament.rowNumber, 8
  ).setValue(true);

  SpreadsheetApp.flush();

  return true;
}


/*************************************************
 * 3. 毎日の大会切り替えチェック
 *************************************************/

function checkTournamentChange() {
  const lock = LockService.getScriptLock();
  let locked = false;

  try {
    lock.waitLock(30000);
    locked = true;

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const tournamentSheet = changeSheet_(
      ss, CHANGE_CONFIG.tournamentSheet
    );

    const tournaments = readChangeTournaments_(
      tournamentSheet
    );

    const today = Utilities.formatDate(
      new Date(), CHANGE_CONFIG.timezone, "yyyyMMdd"
    );

    // 終了日が今日以降の最初の大会
    const currentIndex = tournaments.findIndex(
      t => t.end >= today
    );

    if (currentIndex === -1) {
      throw new Error(
        "現在の大会がありません。次大会を登録してください"
      );
    }

    const current = tournaments[currentIndex];

    // 過去大会の処理漏れを確認
    for (let i = 0; i < currentIndex - 1; i++) {
      if (
        !tournaments[i].finished ||
        !tournaments[i + 1].initialized
      ) {
        throw new Error(
          "過去の大会に処理漏れがあります：" +
          tournaments[i].id +
          "。手動確認が必要です"
        );
      }
    }

    let finishResult = null;
    let initialized = false;
    let finalNoticeQueued = false;
    let startNoticeQueued = false;

    if (currentIndex > 0) {
      const previous = tournaments[currentIndex - 1];

      // ① 前大会の終了処理
      if (!previous.finished) {
        finishResult = finishTournament_(
          ss, previous
        );
      }

      // 終了完了をシートから再確認
      const finishCompleted = changeFlag_(
        tournamentSheet.getRange(
          previous.rowNumber, 7
        ).getValue()
      );

      if (!finishCompleted) {
        throw new Error(
          "前大会の終了処理が未完了です"
        );
      }

      // ② 終了通知
      // 戻り値がなければ保存済み結果から復元
      const finalResultForNotice =
        finishResult ||
        getSavedFinalChangeResult_(
          ss, previous
        );

      finalNoticeQueued =
        enqueueChangeFinalNotice_(
          ss, previous, finalResultForNotice
        );

      // ③ 次大会の初期化
      if (!current.initialized) {
        initialized = initializeTournament_(
          ss, current
        );
      }

      // 初期化完了をシートから再確認
      const initializationCompleted = changeFlag_(
        tournamentSheet.getRange(
          current.rowNumber, 8
        ).getValue()
      );

      if (!initializationCompleted) {
        throw new Error(
          "次大会の初期化が未完了です"
        );
      }

      // ④ 次大会開始通知
      startNoticeQueued =
        enqueueChangeStartNotice_(
          ss, current
        );

    } else if (!current.initialized) {
      // 最初の大会の既存ルールを維持
      throw new Error(
        "最初の大会の初期化済フラグがありません。" +
        "既に運用中ならH列をTRUEにしてください"
      );
    }

    const result = {
      success: true,
      currentTournamentId: current.id,
      finishedTournament: finishResult,
      initialized,
      finalNoticeQueued,
      startNoticeQueued
    };

    console.log(JSON.stringify(result, null, 2));
    return result;

  } catch (e) {
    console.error(
      "大会切り替えエラー：" + e.message
    );

    return {
      success: false,
      error: e.message
    };

  } finally {
    if (locked) {
      lock.releaseLock();
    }
  }
}


/*************************************************
 * 4. 日次トリガー作成
 *************************************************/
function setupTournamentChangeTrigger() {
  const functionName = "checkTournamentChange";

  for (const trigger of ScriptApp.getProjectTriggers()) {
    if (trigger.getHandlerFunction() === functionName) {
      ScriptApp.deleteTrigger(trigger);
    }
  }

  ScriptApp.newTrigger(functionName)
    .timeBased()
    .everyDays(1)
    .atHour(0)
    .inTimezone(CHANGE_CONFIG.timezone)
    .create();

  console.log(
    "大会切り替えの日次トリガーを登録しました"
  );
}
