
/*************************************************
 * 馬券購入システム
 *************************************************/

const BET_CONFIG = {
  historySheet: "馬券購入履歴",
  raceSheet: "地方レース情報",
  pointSheet: "ポイント管理",
  maxPurchase: 50000,
  unit: 100
};

const BET_TYPES = {
  "単勝": 1, "複勝": 1,
  "馬連": 2, "枠連": 2, "ワイド": 2,
  "馬単": 2, "枠単": 2,
  "3連複": 3, "3連単": 3
};

const UNORDERED_TYPES = new Set([
  "馬連", "枠連", "ワイド", "3連複"
]);

const PAIR_TYPES = new Set([
  "馬連", "枠連", "ワイド"
]);

const ORDERED_PAIR_TYPES = new Set([
  "馬単", "枠単"
]);

const FIXED_TRIPLE_MODES = {
  "1着固定ながし": [0],
  "2着固定ながし": [1],
  "3着固定ながし": [2],
  "1,2着固定ながし": [0, 1],
  "2,3着固定ながし": [1, 2],
  "1,3着固定ながし": [0, 2]
};


/**
 * 馬券を購入する
 *
 * @param {string} participantId 参加者ID
 * @param {string} venue 競馬場名
 * @param {number} raceNumber Rを付けないレース番号
 * @param {string} betType 券種
 * @param {string} method 方式
 * @param {boolean} multi マルチ
 * @param {string} selection 買い目
 * @param {number} pointPerBet 1点あたりの購入pt
 */
function purchaseBet(
  participantId,
  venue,
  raceNumber,
  betType,
  method,
  multi,
  selection,
  pointPerBet
) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    const history = requireSheet_(
      ss, BET_CONFIG.historySheet
    );
    const races = requireSheet_(
      ss, BET_CONFIG.raceSheet
    );
    const points = requireSheet_(
      ss, BET_CONFIG.pointSheet
    );

    // 現在開催中の大会情報を取得
    const tournament = getCurrentTournament();
    const tournamentId = tournament.tournamentId;

    participantId = String(participantId).trim();
    venue = String(venue).trim();
    betType = String(betType).trim();
    method = String(method).trim();
    selection = String(selection).replace(/\s/g, "");
    raceNumber = Number(raceNumber);
    pointPerBet = Number(pointPerBet);

    if (!participantId || !venue || !selection) {
      throw new Error("必須項目が空欄です");
    }

    if (!Number.isInteger(raceNumber) ||
        raceNumber < 1 || raceNumber > 12) {
      throw new Error("レース番号が不正です");
    }

    if (!Number.isInteger(pointPerBet) ||
        pointPerBet <= 0 ||
        pointPerBet % BET_CONFIG.unit !== 0) {
      throw new Error(
        "購入ptは100pt単位の正の整数にしてください"
      );
    }

    if (typeof multi !== "boolean") {
      throw new Error(
        "マルチはTRUEまたはFALSEで指定してください"
      );
    }

    // 買い目展開・方式の検証
    const tickets = expandBets_(
      betType, method, multi, selection
    );

    const count = tickets.length;
    const total = count * pointPerBet;

    if (count === 0) {
      throw new Error("購入点数が0点です");
    }

    //レース情報の取得
    const timezone = ss.getSpreadsheetTimeZone();

    const raceInfo = findRaceInfo_(
      races, venue, raceNumber, timezone
    );

    const raceId = raceInfo.raceId;

    // 実在する馬番・枠番か検証
    validateTicketNumbers_(tickets, betType, raceInfo);

    // 同一参加者・同一レースの購入済み総額
    const alreadyPurchased = calculateRacePurchase_(
      history, participantId, raceId
    );

    if (alreadyPurchased + total > BET_CONFIG.maxPurchase) {
      const remainingLimit =
        BET_CONFIG.maxPurchase - alreadyPurchased;

      throw new Error(
        "同一レースへの購入上限を超えています。" +
        `購入済み：${alreadyPurchased}pt、` +
        `今回：${total}pt、` +
        `追加可能：${remainingLimit}pt`
      );
    }

    // 参加者の通常ptを確認
    const pointData = points.getDataRange().getValues();
    let pointRow = -1;
    let balance = 0;

    for (let i = 1; i < pointData.length; i++) {
      if (String(pointData[i][0]).trim() === participantId) {
        pointRow = i + 1;
        balance = Number(pointData[i][1]);
        break;
      }
    }

    if (pointRow === -1) {
      throw new Error(
        `参加者ID ${participantId} が見つかりません`
      );
    }

    if (!Number.isFinite(balance) || balance < 0) {
      throw new Error("通常ptの残高データが不正です");
    }

    if (balance < total) {
      throw new Error(
        `ポイント不足：残高${balance}pt、必要${total}pt`
      );
    }

    // 新しい馬券ID
    const betId = generateBetId_(history);

    const now = new Date();
    const nextRow = Math.max(
      history.getLastRow() + 1, 2
    );

    // 購入履歴へ記録
    const record = [
      betId,          // A 馬券ID
      tournamentId,   // B 大会ID（自動取得）
      participantId,  // C 参加者ID
      raceId,         // D レースID
      betType,        // E 券種
      method,         // F 方式
      multi,          // G マルチ
      selection,      // H 買い目
      pointPerBet,    // I 1点あたりの購入pt
      now,            // J 購入日時
      false,          // K 精算済
      false,          // L 的中？
      0               // M 払戻pt
    ];

    // 書き込み失敗時には履歴を元に戻す
    try {
      // 買い目のセルをテキスト形式にする
      history.getRange(nextRow, 8)
        .setNumberFormat("@");

      // 購入履歴を書き込む
      history.getRange(nextRow, 1, 1, 13)
        .setValues([record]);

      points.getRange(pointRow, 2)
        .setValue(balance - total);

      SpreadsheetApp.flush();

    } catch (writeError) {
      // 同時実行防止ロック下での簡易ロールバック
      try {
        points.getRange(pointRow, 2).setValue(balance);
        history.getRange(nextRow, 1, 1, 13)
          .clearContent();
        SpreadsheetApp.flush();
      } catch (rollbackError) {
        throw new Error(
          "書き込みと復元に失敗しました。手動確認が必要です：" +
          rollbackError.message
        );
      }
      throw writeError;
    }

    const result = {
      success: true,
      betId: betId,
      raceId: raceId,
      count: count,
      total: total,
      remaining: balance - total,
      tickets: tickets
    };

    console.log("馬券購入成功");
    console.log(JSON.stringify(result, null, 2));

    return result;

  } catch (e) {
    console.error("馬券購入失敗：" + e.message);

    return {
      success: false,
      error: e.message
    };

  } finally {
    lock.releaseLock();
  }
}


/*************************************************
 * レース検索
 *************************************************/


function findRaceInfo_(sheet, venue, raceNumber, timezone) {
  const data = sheet.getDataRange().getValues();
  const now = new Date();

  const today = Utilities.formatDate(
    now, timezone, "yyyyMMdd"
  );

  for (let i = 1; i < data.length; i++) {
    const row = data[i];

    const raceId = String(row[1]).trim();
    const venueName = String(row[3]).trim();
    const number = Number(row[5]);

    if (
      !raceId.startsWith(today) ||
      venueName !== venue ||
      number !== raceNumber
    ) {
      continue;
    }

    // 確定済みレースは購入不可
    if (String(row[0]).trim() === "確定") {
      throw new Error("このレースは確定済みです");
    }

    // 開催日（C列）
    const raceDate = row[2] instanceof Date
      ? Utilities.formatDate(
          row[2], timezone, "yyyy/MM/dd"
        )
      : String(row[2]).trim();

    // 発走時刻（H列）
    const startTime = row[7] instanceof Date
      ? Utilities.formatDate(
          row[7], timezone, "HH:mm"
        )
      : String(row[7]).trim();

    if (!/^\d{1,2}:\d{2}$/.test(startTime)) {
      throw new Error("発走時刻の形式が不正です");
    }

    const dateMatch = raceDate.match(
      /^(\d{4})[/-](\d{1,2})[/-](\d{1,2})$/
    );

    if (!dateMatch) {
      throw new Error("開催日の形式が不正です");
    }

    const dateKey =
      dateMatch[1] +
      dateMatch[2].padStart(2, "0") +
      dateMatch[3].padStart(2, "0");

    if (dateKey !== today) {
      throw new Error("開催日とレースIDの日付が一致しません");
    }

    // 現在時刻との比較（スプレッドシートのTZ）
    const currentTime = Utilities.formatDate(
      now, timezone, "HH:mm:ss"
    );

    const [hour, minute] = startTime.split(":").map(Number);

    if (hour > 23 || minute > 59) {
      throw new Error("発走時刻の値が不正です");
    }

    const deadline =
      startTime.padStart(5, "0") + ":00";

    if (currentTime >= deadline) {
      throw new Error(
        `購入締切を過ぎています（発走${startTime}）`
      );
    }

    return {
      raceId: raceId,
      venue: venueName,
      raceNumber: number,
      startTime: startTime,
      declaredCount: Number(row[8]),
      cancelled: row[9],
      frameMapping: row[10]
    };
  }

  throw new Error(
    `本日の${venue}${raceNumber}Rが見つかりません`
  );
}



/*************************************************
 * 買い目展開
 *************************************************/

function expandBets_(type, method, multi, selection) {
  const size = BET_TYPES[type];

  if (!size) {
    throw new Error("未対応の券種です：" + type);
  }

  const unordered = UNORDERED_TYPES.has(type);
  const pairUnordered = PAIR_TYPES.has(type);

  const separator =
    type === "馬単" || type === "枠単" ||
    type === "3連単" ? ">" : "-";

  const groups = selection.split(separator)
    .map(parseNumberGroup_);

  let tickets = [];

  // 通常
  if (method === "通常") {
    if (groups.length !== size ||
        groups.some(g => g.length !== 1)) {
      throw new Error(
        "通常方式は1点分の買い目を指定してください"
      );
    }

    tickets.push(groups.map(g => g[0]));

  // ボックス
  } else if (method === "ボックス") {
    if (groups.length !== 1) {
      throw new Error(
        "ボックスはカンマ区切りで指定してください"
      );
    }

    const nums = uniqueNumbers_(groups[0]);

    tickets = combinations_(nums, size);

    if (!unordered) {
      tickets = tickets.flatMap(t => permutations_(t));
    }

  // 馬連・枠連・ワイドのながし
  } else if (method === "ながし") {
    if (!pairUnordered || groups.length !== 2 ||
        groups[0].length !== 1) {
      throw new Error(
        "ながしは馬連・枠連・ワイドで軸-相手の形式です"
      );
    }

    tickets = groups[1].map(n => [
      groups[0][0], n
    ]);

  // 3連複 1頭軸ながし
  } else if (method === "1頭軸ながし") {
    if (type !== "3連複" ||
        groups.length !== 2 ||
        groups[0].length !== 1) {
      throw new Error(
        "3連複の1頭軸ながしは軸-相手一覧です"
      );
    }

    tickets = combinations_(
      uniqueNumbers_(groups[1]), 2
    ).map(t => [groups[0][0], ...t]);

  // 3連複 2頭軸ながし
  } else if (method === "2頭軸ながし") {
    if (type !== "3連複" ||
        groups.length !== 3 ||
        groups[0].length !== 1 ||
        groups[1].length !== 1) {
      throw new Error(
        "3連複の2頭軸ながしは軸-軸-相手一覧です"
      );
    }

    tickets = groups[2].map(n => [
      groups[0][0], groups[1][0], n
    ]);

  // フォーメーション
  } else if (method === "フォーメーション") {
    if (size === 1 || groups.length !== size) {
      throw new Error(
        "フォーメーションのグループ数が不正です"
      );
    }

    tickets = cartesian_(groups);

  // 馬単・枠単の固定ながし
  } else if (method === "1着固定ながし" ||
             method === "2着固定ながし") {

    if (!ORDERED_PAIR_TYPES.has(type) ||
        groups.length !== 2) {
      // 3連単の場合は後の分岐へ
      if (type !== "3連単") {
        throw new Error(
          "この固定ながしは馬単・枠単・3連単用です"
        );
      }
      tickets = expandTripleFixed_(method, groups);
    } else {
      const axisFirst = method === "1着固定ながし";
      const axis = axisFirst ? groups[0] : groups[1];
      const others = axisFirst ? groups[1] : groups[0];

      if (axis.length !== 1) {
        throw new Error("軸は1つにしてください");
      }

      tickets = others.map(n =>
        axisFirst ? [axis[0], n] : [n, axis[0]]
      );
    }

  // 3連単のその他固定ながし
  } else if (Object.prototype.hasOwnProperty.call(
    FIXED_TRIPLE_MODES, method
  )) {
    if (type !== "3連単") {
      throw new Error(
        "この方式は3連単専用です"
      );
    }

    tickets = expandTripleFixed_(method, groups);

  } else {
    throw new Error("未対応の方式です：" + method);
  }

  // マルチ：固定ながしの着順入れ替え
  if (multi) {
    const allowed =
      (ORDERED_PAIR_TYPES.has(type) &&
       ["1着固定ながし", "2着固定ながし"]
         .includes(method)) ||
      (type === "3連単" &&
       Object.prototype.hasOwnProperty.call(
         FIXED_TRIPLE_MODES, method
       ));

    if (!allowed) {
      throw new Error(
        "この券種・方式ではマルチを使用できません"
      );
    }

    tickets = tickets.flatMap(t => permutations_(t));
  }

  // 番号の範囲と重複を検証し、同一買い目を除く
  const seen = new Set();
  const result = [];

  for (const ticket of tickets) {
    if (ticket.length !== size) {
      throw new Error("買い目の頭数が不正です");
    }

    for (const n of ticket) {
      const max = type.startsWith("枠") ? 8 : 18;
      if (!Number.isInteger(n) || n < 1 || n > max) {
        throw new Error("馬番・枠番の範囲が不正です");
      }
    }

    // 枠連・枠単のゾロ目は例外的に許可
    const sameFrame =
      (type === "枠連" || type === "枠単") &&
      size === 2 &&
      ticket[0] === ticket[1];

    if (!sameFrame &&
        new Set(ticket).size !== ticket.length) {
      continue;
    }

    const canonical = unordered
      ? [...ticket].sort((a, b) => a - b)
      : [...ticket];

    const key = canonical.join(
      unordered ? "-" : ">"
    );

    if (!seen.has(key)) {
      seen.add(key);
      result.push(canonical);
    }
  }

  if (result.length === 0) {
    throw new Error("有効な買い目がありません");
  }

  return result.map(t =>
    t.join(unordered ? "-" : ">")
  );
}


/*************************************************
 * 3連単 固定ながし展開
 *************************************************/

function expandTripleFixed_(method, groups) {
  const fixed = FIXED_TRIPLE_MODES[method];

  if (!fixed) {
    throw new Error("不正な3連単固定方式です");
  }

  const tickets = [];

  // 1着・2着・3着のうち1か所を固定
  if (fixed.length === 1) {
    if (groups.length !== 2) {
      throw new Error(
        "1か所固定は 軸>相手一覧 の形式です"
      );
    }

    const fixedPos = fixed[0];

    // 買い目は着順の位置に従う
    // 2着固定：相手一覧>軸
    const axisGroup =
      fixedPos === 0 ? groups[0] : groups[1];
    const otherGroup =
      fixedPos === 0 ? groups[1] : groups[0];

    if (axisGroup.length !== 1) {
      throw new Error("固定軸は1頭にしてください");
    }

    const pairs = permutationsOfLength_(
      uniqueNumbers_(otherGroup), 2
    );

    for (const pair of pairs) {
      const ticket = new Array(3);
      ticket[fixedPos] = axisGroup[0];

      const openPos = [0, 1, 2]
        .filter(p => p !== fixedPos);

      ticket[openPos[0]] = pair[0];
      ticket[openPos[1]] = pair[1];
      tickets.push(ticket);
    }

  // 2か所固定
  } else {
    if (groups.length !== 3) {
      throw new Error(
        "2か所固定は3グループ指定してください"
      );
    }

    const openPos = [0, 1, 2]
      .find(p => !fixed.includes(p));

    for (const p of fixed) {
      if (groups[p].length !== 1) {
        throw new Error(
          "固定する着順は各1頭にしてください"
        );
      }
    }

    for (const n of groups[openPos]) {
      const ticket = new Array(3);
      for (const p of fixed) {
        ticket[p] = groups[p][0];
      }
      ticket[openPos] = n;
      tickets.push(ticket);
    }
  }

  return tickets;
}


/*************************************************
 * 組合せ・順列・直積
 *************************************************/

function parseNumberGroup_(text) {
  if (!/^\d+(,\d+)*$/.test(text)) {
    throw new Error(
      "買い目の形式が不正です：" + text
    );
  }
  return text.split(",").map(Number);
}

function uniqueNumbers_(array) {
  return [...new Set(array)];
}

function combinations_(arr, k) {
  if (k === 0) return [[]];
  if (arr.length < k) return [];

  const result = [];

  for (let i = 0; i <= arr.length - k; i++) {
    const rest = combinations_(arr.slice(i + 1), k - 1);
    for (const r of rest) {
      result.push([arr[i], ...r]);
    }
  }

  return result;
}

function permutations_(arr) {
  if (arr.length === 0) return [[]];

  const result = [];

  for (let i = 0; i < arr.length; i++) {
    const rest = arr.filter((_, j) => j !== i);
    for (const p of permutations_(rest)) {
      result.push([arr[i], ...p]);
    }
  }

  return result;
}

function permutationsOfLength_(arr, k) {
  return combinations_(arr, k)
    .flatMap(x => permutations_(x));
}

function cartesian_(groups) {
  let result = [[]];

  for (const group of groups) {
    const next = [];
    for (const prefix of result) {
      for (const n of group) {
        next.push([...prefix, n]);
      }
    }
    result = next;
  }

  return result;
}


/*************************************************
 * ID生成・共通処理
 *************************************************/

function generateBetId_(sheet) {
  const last = sheet.getLastRow();
  let maxId = 0;

  if (last >= 2) {
    const ids = sheet.getRange(
      2, 1, last - 1, 1
    ).getValues();

    for (const [value] of ids) {
      const match = /^B(\d+)$/.exec(
        String(value).trim()
      );

      if (match) {
        maxId = Math.max(maxId, Number(match[1]));
      }
    }
  }

  return "B" + String(maxId + 1).padStart(6, "0");
}

function requireSheet_(ss, name) {
  const sheet = ss.getSheetByName(name);

  if (!sheet) {
    throw new Error(
      `シート「${name}」が見つかりません`
    );
  }

  return sheet;
}


/*************************************************
 * 出走馬・枠番の検証
 *************************************************/

function validateTicketNumbers_(tickets, betType, raceInfo) {
  const mapping = String(raceInfo.frameMapping || "");
  const frameByHorse = new Map();

  // 馬番:枠番 の対応を読み取る
  const entries = mapping.split(",")
    .map(s => s.trim())
    .filter(Boolean);

  if (entries.length === 0) {
    throw new Error("馬番・枠番対応が登録されていません");
  }

  for (const entry of entries) {
    const match = entry.match(/^(\d+)[:：](\d+)$/);

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
      frameByHorse.has(horse)
    ) {
      throw new Error("馬番・枠番対応に不正な値があります");
    }

    frameByHorse.set(horse, frame);
  }

  // 取消・除外された馬番を取り除く
  // J列は「3,5」などの馬番一覧とする
  const cancelledText = String(
    raceInfo.cancelled ?? ""
  ).trim();

  const cancelled = new Set();

  if (cancelledText) {
    if (!/^\d+(\s*,\s*\d+)*$/.test(cancelledText)) {
      throw new Error(
        "取消・除外欄の形式が不正です：" + cancelledText
      );
    }

    for (const value of cancelledText.split(",")) {
      cancelled.add(Number(value.trim()));
    }
  }

  for (const horse of cancelled) {
    if (!frameByHorse.has(horse)) {
      throw new Error(
        `取消・除外欄の馬番${horse}が存在しません`
      );
    }
    frameByHorse.delete(horse);
  }

  const activeHorses = new Set(frameByHorse.keys());
  const activeFrames = new Set(frameByHorse.values());

  const frameBet = betType.startsWith("枠");

  // 馬券の各組合せを検証
  for (const ticket of tickets) {
    const numbers = ticket.split(/[-＞>]/).map(Number);

    for (const n of numbers) {
      if (frameBet) {
        if (!activeFrames.has(n)) {
          throw new Error(
            `枠番${n}はこのレースで購入できません`
          );
        }
      } else {
        if (!activeHorses.has(n)) {
          throw new Error(
            `馬番${n}はこのレースで購入できません`
          );
        }
      }
    }

    // 枠連の同枠組合せは、同枠に2頭以上必要
    if (
      (betType === "枠連" || betType === "枠単") &&
      numbers.length === 2 &&
      numbers[0] === numbers[1]
    ) {
      const count = [...frameByHorse.values()]
        .filter(f => f === numbers[0]).length;

      if (count < 2) {
        throw new Error(
          `${numbers[0]}-${numbers[0]}は同枠に2頭以上必要です`
        );
      }
    }
  }

  return true;
}


/*************************************************
 * 同一レースの累計購入額を計算
 *************************************************/

function calculateRacePurchase_(
  historySheet,
  participantId,
  raceId
) {
  const lastRow = historySheet.getLastRow();

  if (lastRow < 2) {
    return 0;
  }

  const data = historySheet.getRange(
    2, 1, lastRow - 1, 13
  ).getValues();

  let total = 0;

  for (const row of data) {
    const rowParticipant = String(row[2]).trim();
    const rowRaceId = String(row[3]).trim();

    if (
      rowParticipant !== participantId ||
      rowRaceId !== raceId
    ) {
      continue;
    }

    const betType = String(row[4]).trim();
    const method = String(row[5]).trim();
    const multi = row[6] === true ||
      String(row[6]).toUpperCase() === "TRUE";
    const selection = String(row[7]).trim();
    const pointPerBet = Number(row[8]);

    if (
      !Number.isInteger(pointPerBet) ||
      pointPerBet <= 0
    ) {
      throw new Error(
        "過去の購入履歴に不正な購入ptがあります"
      );
    }

    // 過去の購入点数を再計算
    const tickets = expandBets_(
      betType, method, multi, selection
    );

    total += tickets.length * pointPerBet;
  }

  return total;
}
