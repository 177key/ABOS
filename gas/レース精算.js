
/*************************************************
 * レース精算システム
 *
 * 依存する既存関数：
 * expandBets_(券種, 方式, マルチ, 買い目)
 *************************************************/

const SETTLE_CONFIG = {
  raceSheet: "地方レース情報",
  historySheet: "馬券購入履歴",
  pointSheet: "ポイント管理",

  // M～U：的中組合せ
  winningStartCol: 13,

  // V～AD：100円あたりの払戻額
  payoutStartCol: 22,

  // AE：確定情報更新時刻
  completedCol: 31
};

const SETTLE_TYPES = [
  "単勝", "複勝", "馬連",
  "馬単", "ワイド", "枠連",
  "枠単", "3連複", "3連単"
];

const SETTLE_UNORDERED = new Set([
  "馬連", "ワイド", "枠連", "3連複"
]);

const SETTLE_SIZE = {
  "単勝": 1,
  "複勝": 1,
  "馬連": 2,
  "馬単": 2,
  "ワイド": 2,
  "枠連": 2,
  "枠単": 2,
  "3連複": 3,
  "3連単": 3
};


/**
 * 指定レースの精算処理
 *
 * @param {string} raceId レースID
 * @return {Object} 精算結果
 */
function settleRace(raceId) {
  const lock = LockService.getScriptLock();
  let locked = false;

  try {
    lock.waitLock(30000);
    locked = true;

    const ss = SpreadsheetApp.getActiveSpreadsheet();

    const raceSheet = ss.getSheetByName(
      SETTLE_CONFIG.raceSheet
    );
    const historySheet = ss.getSheetByName(
      SETTLE_CONFIG.historySheet
    );
    const pointSheet = ss.getSheetByName(
      SETTLE_CONFIG.pointSheet
    );
    const bonusSheet = ss.getSheetByName(
      BONUS_CONFIG.sheetName
    );

    if (!bonusSheet) {
      throw new Error(
        "ボーナス管理シートが見つかりません"
      );
    }

    if (!raceSheet || !historySheet || !pointSheet) {
      throw new Error("必要なシートが見つかりません");
    }

    raceId = String(raceId).trim();

    if (!/^\d{8}[a-z]{3}\d{2}$/.test(raceId)) {
      throw new Error(
        "レースIDの形式が不正です：" + raceId
      );
    }

    // 1. レース情報を取得
    const raceInfo = getSettlementRaceInfo_(
      raceSheet, raceId
    );

    // 2. 払戻・返還情報を読み込む
    const payoutMap = buildSettlementPayoutMap_(
      raceInfo.row
    );
    const refundInfo = buildRefundInfo_(raceInfo.row);

    // 3. 購入履歴を参加者別にまとめる
    const participants = collectSettlementBets_(
      historySheet, raceId
    );

    // 4. ポイント残高を取得
    const pointData = pointSheet
      .getDataRange().getValues();

    const pointMap = new Map();

    for (let i = 1; i < pointData.length; i++) {
      const id = String(pointData[i][0]).trim();

      if (!id) continue;

      if (pointMap.has(id)) {
        throw new Error(
          `ポイント管理に参加者ID ${id} が重複しています`
        );
      }

      pointMap.set(id, {
        rowNumber: i + 1,
        normal: pointData[i][1],
        profit: pointData[i][2]
      });
    }

    // 5. 精算結果をメモリ上で計算
    // この段階ではシートを書き換えない
    const results = [];
    const historyUpdates = [];
    const pointUpdates = [];

    for (const [participantId, bets] of participants) {
      const pointInfo = pointMap.get(participantId);

      if (!pointInfo) {
        throw new Error(
          `ポイント管理に${participantId}が存在しません`
        );
      }

      const normal = Number(pointInfo.normal);
      const profit = Number(pointInfo.profit);

      if (
        !Number.isSafeInteger(normal) ||
        !Number.isSafeInteger(profit) ||
        normal < 0 ||
        profit < 0
      ) {
        throw new Error(
          `${participantId}のポイント残高が不正です`
        );
      }

      let totalOriginalPurchase = 0;
      let totalPurchase = 0;
      let totalPayout = 0;
      let totalRefund = 0;
      let hitBetCount = 0;

      const betResults = [];

      for (const bet of bets) {
        const result = calculateSettlementBet_(
          bet,
          payoutMap,
          refundInfo
        );

        totalOriginalPurchase += result.originalPurchase;
        totalPurchase += result.effectivePurchase;
        totalPayout += result.hitPayout;
        totalRefund += result.refundTotal;

        if (result.hit) {
          hitBetCount++;
        }

        betResults.push(result);

        historyUpdates.push({
          rowNumber: bet.rowNumber,
          oldValues: bet.settlementValues,
          newValues: [
            true,
            result.hit,
            result.totalReturned
          ]
        });
      }

      const netProfit = totalPayout - totalPurchase;
      const earnedProfit = Math.max(0, netProfit);

      if (
        ![
          totalOriginalPurchase,
          totalPurchase,
          totalPayout,
          totalRefund,
          earnedProfit,
          normal + totalPayout + totalRefund,
          profit + earnedProfit
        ].every(Number.isSafeInteger)
      ) {
        throw new Error(
          `${participantId}の精算額が安全な整数範囲を超えています`
        );
      }

      pointUpdates.push({
        rowNumber: pointInfo.rowNumber,
        oldValues: [
          pointInfo.normal,
          pointInfo.profit
        ],
        newValues: [
          normal + totalPayout + totalRefund,
          profit + earnedProfit
        ]
      });

      results.push({
        participantId: participantId,
        betCount: bets.length,
        hitBetCount: hitBetCount,
        originalPurchaseTotal: totalOriginalPurchase,
        purchaseTotal: totalPurchase,
        payoutTotal: totalPayout,
        refundTotal: totalRefund,
        totalReturned: totalPayout + totalRefund,
        netProfit: netProfit,
        earnedProfit: earnedProfit,
        bets: betResults,
        normalPtAfterSettlement: normal + totalPayout + totalRefund
      });
    }

    // 参加者ごとの精算結果からボーナス候補を計算
    const bonusCandidates = calculateBonusCandidates_(
      results,
      raceInfo.row
    );

    // 既存記録と比較し、更新予定データを作成
    const bonusUpdates = prepareBonusUpdates_(
      bonusSheet,
      bonusCandidates
    );

    // 6. 全検証成功後にシートを更新
    applySettlementUpdates_(
      historySheet,
      pointSheet,
      historyUpdates,
      pointUpdates,
      bonusSheet,
      bonusUpdates
    );

    const result = {
      success: true,
      raceId: raceId,
      participantCount: results.length,
      betCount: historyUpdates.length,
      totalPayout: results.reduce(
        (sum, r) => sum + r.payoutTotal, 0
      ),
      totalRefund: results.reduce(
        (sum, r) => sum + r.refundTotal, 0
      ),
      totalReturned: results.reduce(
        (sum, r) => sum + r.totalReturned, 0
      ),
      participants: results,
      bonusCandidates: bonusCandidates,
      bonusUpdates: bonusUpdates.map(u => ({
        rowNumber: u.rowNumber,
        holders: u.newValues[0],
        value: u.newValues[1],
        races: u.newValues[2]
    })),
        };

    console.log(
      "レース精算成功：" +
      JSON.stringify(result, null, 2)
    );

    return result;

  } catch (e) {
    console.error("レース精算失敗：" + e.message);

    return {
      success: false,
      raceId: String(raceId),
      error: e.message
    };

  } finally {
    if (locked) lock.releaseLock();
  }
}


/*************************************************
 * レース情報の確認
 *************************************************/

function getSettlementRaceInfo_(sheet, raceId) {
  const lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    throw new Error("レース情報がありません");
  }

  const data = sheet.getRange(
    2, 1, lastRow - 1, 31
  ).getValues();

  let found = null;

  for (let i = 0; i < data.length; i++) {
    const row = data[i];

    if (String(row[1]).trim() !== raceId) {
      continue;
    }

    if (found) {
      throw new Error(
        "レースIDが重複しています：" + raceId
      );
    }

    found = {
      rowNumber: i + 2,
      row: row
    };
  }

  if (!found) {
    throw new Error(
      "指定されたレースが存在しません：" + raceId
    );
  }

  const row = found.row;

  // A列：処理状態
  if (String(row[0]).trim() !== "確定") {
    throw new Error(
      "レースがまだ確定していません"
    );
  }

  // AE列：確定情報更新時刻
  if (
    row[SETTLE_CONFIG.completedCol - 1] === "" ||
    row[SETTLE_CONFIG.completedCol - 1] === null
  ) {
    throw new Error(
      "払戻情報の書き込みが完了していません"
    );
  }

  return found;
}


/*************************************************
 * 払戻情報を読み込む
 *
 * Map<券種, Map<買い目, 100円あたり払戻>>
 *************************************************/

function buildSettlementPayoutMap_(raceRow) {
  const payoutMap = new Map();

  for (let i = 0; i < SETTLE_TYPES.length; i++) {
    const type = SETTLE_TYPES[i];

    const winningRaw = raceRow[
      SETTLE_CONFIG.winningStartCol - 1 + i
    ];

    const payoutRaw = raceRow[
      SETTLE_CONFIG.payoutStartCol - 1 + i
    ];

    const winningText = settlementCellText_(
      winningRaw, type + "の的中組合せ"
    );

    const payoutText = settlementCellText_(
      payoutRaw, type + "の払戻額"
    );

    const typeMap = new Map();

    // 両方空欄：その券種の的中情報なし
    if (!winningText && !payoutText) {
      payoutMap.set(type, typeMap);
      continue;
    }

    if (!winningText || !payoutText) {
      throw new Error(
        `${type}の的中組合せと払戻額の片方が空欄です`
      );
    }

    const winners = winningText.split(",");
    const payouts = payoutText.split(",");

    if (winners.length !== payouts.length) {
      throw new Error(
        `${type}の的中組合せ数と払戻額数が一致しません`
      );
    }

    for (let j = 0; j < winners.length; j++) {
      const key = normalizeSettlementTicket_(
        type, winners[j].trim()
      );

      const amountText = payouts[j].trim();

      if (!/^\d+$/.test(amountText)) {
        throw new Error(
          `${type}の払戻額が不正です：${amountText}`
        );
      }

      const amount = Number(amountText);

      if (
        !Number.isSafeInteger(amount) ||
        amount <= 0
      ) {
        throw new Error(
          `${type}の払戻額が不正です：${amountText}`
        );
      }

      if (typeMap.has(key)) {
        throw new Error(
          `${type}の的中組合せが重複しています：${key}`
        );
      }

      typeMap.set(key, amount);
    }

    payoutMap.set(type, typeMap);
  }

  return payoutMap;
}


/*************************************************
 * 購入履歴を参加者別に取得
 *************************************************/

function collectSettlementBets_(sheet, raceId) {
  const lastRow = sheet.getLastRow();
  const participants = new Map();

  if (lastRow < 2) {
    return participants;
  }

  const data = sheet.getRange(
    2, 1, lastRow - 1, 13
  ).getValues();

  for (let i = 0; i < data.length; i++) {
    const row = data[i];

    if (String(row[3]).trim() !== raceId) {
      continue;
    }

    const betId = String(row[0]).trim();
    const participantId = String(row[2]).trim();

    if (!betId || !participantId) {
      throw new Error(
        `${i + 2}行目の馬券IDまたは参加者IDが空欄です`
      );
    }

    // 二重精算防止
    if (
      row[10] === true ||
      String(row[10]).toUpperCase() === "TRUE"
    ) {
      throw new Error(
        `馬券${betId}は精算済みです。` +
        "二重精算を防止するため処理を中止します"
      );
    }

    if (
      row[10] !== false &&
      String(row[10]).toUpperCase() !== "FALSE"
    ) {
      throw new Error(
        `馬券${betId}の精算済フラグが不正です`
      );
    }

    const bet = {
      rowNumber: i + 2,
      betId: betId,
      participantId: participantId,
      betType: String(row[4]).trim(),
      method: String(row[5]).trim(),
      multi: row[6] === true ||
        String(row[6]).toUpperCase() === "TRUE",
      selection: settlementCellText_(
        row[7], `馬券${betId}の買い目`
      ),
      pointPerBet: Number(row[8]),
      settlementValues: row.slice(10, 13)
    };

    if (!participants.has(participantId)) {
      participants.set(participantId, []);
    }

    participants.get(participantId).push(bet);
  }

  return participants;
}


/*************************************************
 * 1回の馬券購入を精算
 *************************************************/


function calculateSettlementBet_(
  bet,
  payoutMap,
  refundInfo
) {
  const {
    betId, betType, method,
    multi, selection, pointPerBet
  } = bet;

  if (
    !Number.isSafeInteger(pointPerBet) ||
    pointPerBet <= 0 ||
    pointPerBet % 100 !== 0
  ) {
    throw new Error(
      `馬券${betId}の購入ptが不正です`
    );
  }

  if (!payoutMap.has(betType)) {
    throw new Error(
      `未対応の券種です：${betType}`
    );
  }

  const tickets = expandBets_(
    betType, method, multi, selection
  );

  const winners = payoutMap.get(betType);

  let hitPayout = 0;
  let refundTotal = 0;

  const hitTickets = [];
  const refundTickets = [];

  for (const ticket of tickets) {
    const key = normalizeSettlementTicket_(
      betType, ticket
    );

    // 返還対象
    if (isRefundTicket_(betType, ticket, refundInfo)) {
      refundTotal += pointPerBet;
      refundTickets.push(key);
      continue;
    }

    // 不的中
    if (!winners.has(key)) {
      continue;
    }

    // 的中
    const payoutPer100 = winners.get(key);

    const returned =
      (payoutPer100 * pointPerBet) / 100;

    hitPayout += returned;

    hitTickets.push({
      selection: key,
      payoutPer100: payoutPer100,
      payout: returned
    });
  }

  const originalPurchase =
    tickets.length * pointPerBet;

  // 返還分を除いた購入額
  const effectivePurchase =
    originalPurchase - refundTotal;

  // 通常ptへ戻す総額
  const totalReturned =
    hitPayout + refundTotal;

  if (
    ![
      originalPurchase,
      effectivePurchase,
      hitPayout,
      refundTotal,
      totalReturned
    ].every(Number.isSafeInteger)
  ) {
    throw new Error(
      `馬券${betId}の精算金額が不正です`
    );
  }

  return {
    betId,
    betType,
    method,
    multi,
    selection,
    pointPerBet,

    ticketCount: tickets.length,
    hitCount: hitTickets.length,
    refundCount: refundTickets.length,

    originalPurchase,
    effectivePurchase,

    // 返還を除いた的中払戻
    hitPayout,

    // 返還された購入額
    refundTotal,

    // 通常ptへの合計加算額
    totalReturned,

    hit: hitTickets.length > 0,
    hitTickets,
    refundTickets
  };
}



/*************************************************
 * 買い目の正規化
 *
 * 馬連など：順序を無視
 * 馬単など：順序を維持
 *************************************************/

function normalizeSettlementTicket_(type, ticket) {
  const text = String(ticket)
    .replace(/\s/g, "")
    .replace(/＞/g, ">");

  const size = SETTLE_SIZE[type];

  if (!size) {
    throw new Error("未対応の券種：" + type);
  }

  const ordered =
    type === "馬単" ||
    type === "枠単" ||
    type === "3連単";

  const separator = ordered ? ">" : "-";
  const pattern = ordered
    ? /^\d+(>\d+)*$/
    : /^\d+(-\d+)*$/;

  if (!pattern.test(text)) {
    throw new Error(
      `的中組合せの形式が不正です：${text}`
    );
  }

  const numbers = text.split(separator).map(Number);

  if (
    numbers.length !== size ||
    numbers.some(n =>
      !Number.isSafeInteger(n) ||
      n < 1 ||
      n > (type.startsWith("枠") ? 8 : 18)
    )
  ) {
    throw new Error(
      `的中組合せの番号が不正です：${text}`
    );
  }

  const sameFrame =
    (type === "枠連" || type === "枠単") &&
    size === 2 &&
    numbers[0] === numbers[1];

  if (
    !sameFrame &&
    new Set(numbers).size !== numbers.length
  ) {
    throw new Error(
      `的中組合せに同一番号があります：${text}`
    );
  }

  if (SETTLE_UNORDERED.has(type)) {
    numbers.sort((a, b) => a - b);
  }

  return numbers.join(separator);
}


/*************************************************
 * スプレッドシートの文字列安全チェック
 *************************************************/

function settlementCellText_(value, label) {
  if (value === null || value === "") {
    return "";
  }

  // 例：「3-7」が日付に誤変換された場合
  if (value instanceof Date) {
    throw new Error(
      `${label}が日付に変換されています。` +
      "セルをテキスト形式に直してください"
    );
  }

  return String(value).trim();
}



/*************************************************
 * 精算結果を3シートに反映
 *************************************************/

function applySettlementUpdates_(
  historySheet,
  pointSheet,
  historyUpdates,
  pointUpdates,
  bonusSheet,
  bonusUpdates
) {
  const applied = [];

  // 更新対象の元データを保存する
  function writeWithBackup_(sheet, row, col, oldValues, newValues) {
    const range = sheet.getRange(
      row, col, 1, newValues.length
    );

    // 書き込み失敗後でも復元できるよう、
    // 書き込みを試みる前に復元情報を登録
    applied.push({
      sheet,
      row,
      col,
      oldValues
    });

    range.setValues([newValues]);
  }

  try {
    // 1. 購入履歴 K～M
    for (const u of historyUpdates) {
      writeWithBackup_(
        historySheet,
        u.rowNumber,
        11,
        u.oldValues,
        u.newValues
      );
    }

    // 2. ポイント管理 B～C
    for (const u of pointUpdates) {
      writeWithBackup_(
        pointSheet,
        u.rowNumber,
        2,
        u.oldValues,
        u.newValues
      );
    }

    // 3. ボーナス管理 B～D
    for (const u of bonusUpdates) {
      writeWithBackup_(
        bonusSheet,
        u.rowNumber,
        2,
        u.oldValues,
        u.newValues
      );
    }

    SpreadsheetApp.flush();

  } catch (error) {
    let rollbackFailed = false;

    // 逆順に元の値へ戻す
    for (let i = applied.length - 1; i >= 0; i--) {
      const u = applied[i];

      try {
        u.sheet.getRange(
          u.row,
          u.col,
          1,
          u.oldValues.length
        ).setValues([u.oldValues]);
      } catch (e) {
        rollbackFailed = true;
        console.error(
          `復元失敗（${u.sheet.getName()} ${u.row}行）: ` +
          e.message
        );
      }
    }

    try {
      SpreadsheetApp.flush();
    } catch (e) {
      rollbackFailed = true;
    }

    if (rollbackFailed) {
      throw new Error(
        "精算途中でエラーが発生し、" +
        "一部の復元にも失敗しました。" +
        "手動確認が必要です：" + error.message
      );
    }

    throw error;
  }
}




/**
 * レース情報から返還判定に必要な情報を取得
 *
 * raceRow：
 * 「地方レース情報」のA～AE列を含む配列
 */
function buildRefundInfo_(raceRow) {
  const cancelledText = settlementCellText_(
    raceRow[9], "取消・除外"
  );

  const mappingText = settlementCellText_(
    raceRow[10], "馬番・枠番対応"
  );

  // J列：取消・除外馬番
  const cancelled = new Set();

  if (cancelledText) {
    if (!/^\d+(,\d+)*$/.test(
      cancelledText.replace(/\s/g, "")
    )) {
      throw new Error(
        "取消・除外欄の形式が不正です"
      );
    }

    for (const s of cancelledText.split(",")) {
      cancelled.add(Number(s.trim()));
    }
  }

  // K列：馬番 -> 枠番
  const horseToFrame = new Map();
  const frameToHorses = new Map();

  if (!mappingText) {
    throw new Error("馬番・枠番対応が空欄です");
  }

  for (const entry of mappingText.split(",")) {
    const match = entry.trim().match(
      /^(\d+):(\d+)$/
    );

    if (!match) {
      throw new Error(
        "馬番・枠番対応の形式が不正です：" + entry
      );
    }

    const horse = Number(match[1]);
    const frame = Number(match[2]);

    if (
      horse < 1 || horse > 18 ||
      frame < 1 || frame > 8 ||
      horseToFrame.has(horse)
    ) {
      throw new Error(
        "馬番・枠番対応に不正な値があります"
      );
    }

    horseToFrame.set(horse, frame);

    if (!frameToHorses.has(frame)) {
      frameToHorses.set(frame, []);
    }

    frameToHorses.get(frame).push(horse);
  }

  for (const horse of cancelled) {
    if (!horseToFrame.has(horse)) {
      throw new Error(
        `取消・除外馬番${horse}が対応表にありません`
      );
    }
  }

  // 枠ごとの出走可能頭数
  const activeFrameCounts = new Map();

  for (const [frame, horses] of frameToHorses) {
    const activeCount = horses.filter(
      horse => !cancelled.has(horse)
    ).length;

    activeFrameCounts.set(frame, activeCount);
  }

  return {
    cancelled,
    horseToFrame,
    activeFrameCounts
  };
}


/**
 * 展開済みの買い目が返還対象か判定
 */
function isRefundTicket_(betType, ticket, refundInfo) {
  const numbers = String(ticket)
    .split(/[-＞>]/)
    .map(Number);

  // 馬番を指定する券種
  if (!betType.startsWith("枠")) {
    return numbers.some(
      n => refundInfo.cancelled.has(n)
    );
  }

  // 枠番を指定する券種
  // 出走馬が0頭になった枠を含む場合は返還
  for (const frame of numbers) {
    const count =
      refundInfo.activeFrameCounts.get(frame) || 0;

    if (count === 0) {
      return true;
    }
  }

  // 枠連・枠単のゾロ目
  // 同枠に出走可能馬が2頭未満となった場合
  if (
    numbers.length === 2 &&
    numbers[0] === numbers[1]
  ) {
    const count =
      refundInfo.activeFrameCounts.get(numbers[0]) || 0;

    if (count < 2) {
      return true;
    }
  }

  return false;
}



/*************************************************
 * ボーナス記録管理
 *************************************************/

const BONUS_CONFIG = {
  sheetName: "ボーナス管理",

  // 最高オッズ：3～11行
  oddsStartRow: 3,

  // 最高利益pt：15～23行
  profitStartRow: 15,

  types: [
    "単勝", "複勝", "馬連",
    "馬単", "ワイド", "枠連",
    "枠単", "3連複", "3連単"
  ]
};


/**
 * 精算済みの参加者データからボーナス候補を計算する。
 *
 * results は settleRace() 内で構築した
 * 参加者別精算結果の配列。
 */
function calculateBonusCandidates_(results, raceRow) {
  const candidates = [];

  // C列：開催日
  const rawDate = raceRow[2];

  // D列：競馬場、F列：レース番号、G列：レース名
  const venue = String(raceRow[3] ?? "").trim();
  const raceNumber = Number(raceRow[5]);
  const raceName = String(raceRow[6] ?? "").trim();

  // 開催日を yyyy/MM/dd 形式に統一
  let dateStr;

  if (rawDate instanceof Date) {
    dateStr = Utilities.formatDate(
      rawDate,
      "Asia/Tokyo",
      "yyyy/MM/dd"
    );
  } else {
    const match = String(rawDate).trim().match(
      /^(\d{4})[/-](\d{1,2})[/-](\d{1,2})$/
    );

    if (!match) {
      throw new Error("開催日の形式が不正です");
    }

    dateStr = `${match[1]}/${match[2].padStart(2, "0")}/${match[3].padStart(2, "0")}`;
  }

  if (!venue || !Number.isInteger(raceNumber) ||
      raceNumber < 1 || raceNumber > 12) {
    throw new Error(
      "ボーナス記録用の競馬場・レース番号が不正です"
    );
  }

  // 記録用のレース名
  const raceLabel =
    `${dateStr} ${venue}${raceNumber}R` +
    (raceName ? ` ${raceName}` : "");

  for (const participant of results) {

    // 参加者の馬券を券種別に集計
    const groups = new Map();

    for (const bet of participant.bets) {
      const type = bet.betType;

      if (!groups.has(type)) {
        groups.set(type, {
          purchase: 0,
          payout: 0,
          maxPayoutPer100: 0
        });
      }

      const group = groups.get(type);

      // 返還分は購入額から除外
      group.purchase += bet.effectivePurchase;

      // 返還を除いた的中払戻額
      group.payout += bet.hitPayout;

      // この馬券で的中した買い目の最高配当
      for (const hit of bet.hitTickets) {
        group.maxPayoutPer100 = Math.max(
          group.maxPayoutPer100,
          hit.payoutPer100
        );
      }
    }

    for (const [type, group] of groups) {

      const netProfit = group.payout - group.purchase;

      // 券種単位でプラス収支の場合のみ対象
      if (netProfit <= 0) {
        continue;
      }

      if (
        !Number.isSafeInteger(group.purchase) ||
        !Number.isSafeInteger(group.payout) ||
        !Number.isSafeInteger(group.maxPayoutPer100) ||
        !Number.isSafeInteger(netProfit)
      ) {
        throw new Error(
          `${participant.participantId}の` +
          `${type}ボーナス計算値が不正です`
        );
      }

      if (group.maxPayoutPer100 <= 0) {
        throw new Error(
          `${type}はプラス収支ですが的中情報がありません`
        );
      }

      candidates.push({
        participantId: participant.participantId,
        betType: type,
        raceLabel: raceLabel,

        // 最高オッズは100ptあたりの払戻額を
        // 100で割った倍率
        oddsValue: group.maxPayoutPer100 / 100,

        // 券種単位の利益pt
        profitValue: netProfit
      });
    }
  }

  return candidates;
}


/**
 * ボーナス管理シートと比較して更新予定データを作成。
 *
 * まだシートには書き込まない。
 */
function prepareBonusUpdates_(bonusSheet, candidates) {
  const data = bonusSheet.getRange(
    1, 1, 23, 4
  ).getValues();

  const updates = [];

  // 最高オッズ / 最高利益pt
  const categories = [
    {
      category: "odds",
      startRow: BONUS_CONFIG.oddsStartRow,
      valueKey: "oddsValue"
    },
    {
      category: "profit",
      startRow: BONUS_CONFIG.profitStartRow,
      valueKey: "profitValue"
    }
  ];

  for (const config of categories) {
    for (let i = 0; i < BONUS_CONFIG.types.length; i++) {
      const type = BONUS_CONFIG.types[i];
      const rowNumber = config.startRow + i;
      const row = data[rowNumber - 1];

      if (String(row[0]).trim() !== type) {
        throw new Error(
          `ボーナス管理${rowNumber}行目の券種が不正です`
        );
      }

      // B～D列の現在値
      const oldValues = row.slice(1, 4);

      const existingIds = String(
        oldValues[0] ?? ""
      ).trim();

      const existingValue = oldValues[1];

      const existingRaces = String(
        oldValues[2] ?? ""
      ).trim();

      const isEmpty =
        !existingIds &&
        (existingValue === "" ||
         existingValue === null) &&
        !existingRaces;

      if (
        !isEmpty &&
        (!existingIds || existingValue === "" ||
         existingValue === null || !existingRaces)
      ) {
        throw new Error(
          `ボーナス管理${rowNumber}行目の記録が不完全です`
        );
      }

      let currentValue = isEmpty
        ? 0
        : Number(existingValue);

      if (
        !Number.isFinite(currentValue) ||
        currentValue < 0
      ) {
        throw new Error(
          `ボーナス管理${rowNumber}行目の記録値が不正です`
        );
      }

      // 既存の同率保持者
      let holders = [];

      if (!isEmpty) {
        const ids = existingIds.split("\n");
        const races = existingRaces.split("\n");

        if (ids.length !== races.length) {
          throw new Error(
            `ボーナス管理${rowNumber}行目の` +
            "保持者とレースの件数が一致しません"
          );
        }

        holders = ids.map((id, index) => ({
          participantId: id.trim(),
          raceLabel: races[index].trim()
        }));

        if (holders.some(h =>
          !h.participantId || !h.raceLabel
        )) {
          throw new Error(
            `ボーナス管理${rowNumber}行目に空欄があります`
          );
        }
      }

      const targets = candidates.filter(
        c => c.betType === type
      );

      for (const target of targets) {
        const value = target[config.valueKey];

        if (value > currentValue) {
          // 新記録：従来の保持者を置き換える
          currentValue = value;
          holders = [{
            participantId: target.participantId,
            raceLabel: target.raceLabel
          }];

        } else if (value === currentValue) {

          // 同率：同じ参加者・同じレースの
          // 重複記録は作らない
          const exists = holders.some(h =>
            h.participantId === target.participantId &&
            h.raceLabel === target.raceLabel
          );

          if (!exists) {
            holders.push({
              participantId: target.participantId,
              raceLabel: target.raceLabel
            });
          }
        }
      }

      const newValues = holders.length
        ? [
            holders.map(h => h.participantId).join("\n"),
            currentValue,
            holders.map(h => h.raceLabel).join("\n")
          ]
        : oldValues;

      if (
        JSON.stringify(newValues) !==
        JSON.stringify(oldValues)
      ) {
        updates.push({
          rowNumber,
          oldValues,
          newValues
        });
      }
    }
  }

  return updates;
}
