
/**
 * 確定情報が登録済みで、
 * 未精算馬券が残っているレースを取得する。
 *
 * シートの変更は行わない。
 */
function findUnsettledConfirmedRaces_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const raceSheet = ss.getSheetByName("地方レース情報");
  const historySheet = ss.getSheetByName("馬券購入履歴");

  if (!raceSheet || !historySheet) {
    throw new Error(
      "地方レース情報または馬券購入履歴がありません"
    );
  }

  if (
    raceSheet.getLastRow() < 2 ||
    historySheet.getLastRow() < 2
  ) {
    return [];
  }

  // 馬券購入履歴：A～K列
  const historyRows = historySheet.getRange(
    2, 1, historySheet.getLastRow() - 1, 11
  ).getValues();

  // レース別の精算状況
  const betStatus = new Map();

  for (const row of historyRows) {
    const raceId = String(row[3] || "").trim();
    if (!raceId) continue;

    if (!betStatus.has(raceId)) {
      betStatus.set(raceId, {
        unsettled: 0,
        settled: 0,
        invalid: 0
      });
    }

    const status = betStatus.get(raceId);
    const flag = row[10];
    const text = String(flag).trim().toUpperCase();

    if (flag === true || text === "TRUE") {
      status.settled++;
    } else if (flag === false || text === "FALSE") {
      status.unsettled++;
    } else {
      status.invalid++;
    }
  }

  // 地方レース情報：A～AE列
  const raceRows = raceSheet.getRange(
    2, 1, raceSheet.getLastRow() - 1, 31
  ).getValues();

  const candidates = [];
  const seen = new Set();

  for (const row of raceRows) {
    const raceId = String(row[1] || "").trim();

    if (!/^\d{8}[a-z]{3}\d{2}$/.test(raceId)) {
      continue;
    }

    // AE列：確定情報更新時刻
    const completedAt = row[30];

    if (
      completedAt === "" ||
      completedAt === null
    ) {
      continue;
    }

    if (seen.has(raceId)) {
      throw new Error(
        "レース情報に重複IDがあります：" + raceId
      );
    }
    seen.add(raceId);

    const status = betStatus.get(raceId);

    // 購入履歴のないレースは精算不要
    if (!status) continue;

    // 精算状態に不明な値がある場合は止める
    if (status.invalid > 0) {
      console.warn(
        `${raceId}：精算フラグが不正な馬券があります`
      );
      continue;
    }

    // 同一レースに精算済み・未精算が混在
    // 現行のsettleRace()では安全に処理できない
    if (
      status.unsettled > 0 &&
      status.settled > 0
    ) {
      console.warn(
        `${raceId}：精算済みと未精算が混在しています。` +
        "手動確認が必要です"
      );
      continue;
    }

    // 全馬券が未精算のレースだけ対象
    if (
      status.unsettled > 0 &&
      status.settled === 0
    ) {
      candidates.push(raceId);
    }
  }

  return candidates.sort();
}


/**
 * 確定済みの未精算レースを自動精算する。
 *
 * 払戻取得処理とは別に実行できる。
 */
function settleConfirmedRacesAutomatically() {
  const raceIds = findUnsettledConfirmedRaces_();

  const summary = {
    candidates: raceIds.length,
    success: [],
    failed: []
  };

  for (const raceId of raceIds) {
    try {
      // 既存の精算＋Discord通知
      const result = settleRaceWithNotification(raceId);

      if (result && result.success === true) {
        summary.success.push(raceId);

        console.log(
          "自動精算成功：" + raceId
        );
      } else {
        summary.failed.push({
          raceId,
          error: result?.error ||
            result?.message ||
            "精算処理が失敗しました"
        });

        console.warn(
          "自動精算失敗：" + raceId
        );
      }

    } catch (error) {
      summary.failed.push({
        raceId,
        error: String(error.message || error)
      });

      console.error(
        "自動精算エラー：" +
        raceId + " / " + error.message
      );
    }
  }

  console.log(
    JSON.stringify(summary, null, 2)
  );

  return summary;
}



/**
 * 10分ごとの自動実行用
 *
 * ① 公式払戻情報を取得
 * ② 確定済み・未精算レースを精算
 */
function runLocalRaceRefundAndSettlement() {
  console.log("払戻取得・自動精算を開始");

  // 既存の払戻情報取得
  const syncResult = updateLocalRaceRefundSync();

  // 取得処理が明示的に失敗を返した場合は中止
  if (
    syncResult &&
    syncResult.success === false
  ) {
    throw new Error(
      "払戻情報取得に失敗しました：" +
      (syncResult.error || "原因不明")
    );
  }

  // 確定済みレースの精算
  const settlementResult =
    settleConfirmedRacesAutomatically();

  console.log("払戻取得・自動精算が終了");

  return {
    success: settlementResult.failed.length === 0,
    syncResult,
    settlementResult
  };
}
