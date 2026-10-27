
/**
 * Webアプリの画面を表示する
 * 既存のdoPost(e)はそのままでOK
 */
function doGet(e) {
  return HtmlService
    .createHtmlOutputFromFile("BetPage")
    .setTitle("エア馬券購入");
}

/**
 * 今日のレース一覧を取得する
 * 既存の「地方レース情報」シートを参照
 */
function getTodayRacesForWeb() {
  const sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName("地方レース情報");

  if (!sheet) {
    throw new Error("地方レース情報シートがありません。");
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  // A～L列を取得
  const rows = sheet.getRange(
    2, 1, lastRow - 1, 12
  ).getDisplayValues();

  const timezone = "Asia/Tokyo";
  const today = Utilities.formatDate(
    new Date(), timezone, "yyyyMMdd"
  );

  const races = [];

  for (const row of rows) {
    // B列のレースIDは YYYYMMDD + 競馬場コード + RR
    const raceId = String(row[1] || "").trim();

    if (!raceId.startsWith(today)) continue;

    const venue = String(row[3] || "").trim();
    const raceNumber = Number(row[5]);
    const name = String(row[6] || "").trim();
    const startTime = String(row[7] || "").trim();

    if (!venue || !Number.isInteger(raceNumber)) {
      continue;
    }


    races.push({
      raceId: raceId,
      venue: venue,
      raceNumber: raceNumber,
      name: name,
      startTime: startTime,
      runnerCount: Number(row[8]) || 0,

      // J列：取消・除外馬番
      excludedNumbers: String(row[9] || "")
        .split(/[,，、]/)
        .map(s => Number(s.trim()))
        .filter(n => Number.isInteger(n) && n > 0),

      // K列：馬番:枠番の対応
      framePairs: String(row[10] || "")
        .split(/[,，、]/)
        .map(s => {
          const match = s.trim().match(
            /^(\d+)\s*[:：]\s*(\d+)$/
          );
          if (!match) return null;

          return {
            horse: Number(match[1]),
            frame: Number(match[2])
          };
        })
        .filter(x => x !== null)
    });
  }

  races.sort((a, b) => {
    const byVenue = a.venue.localeCompare(
      b.venue, "ja"
    );

    return byVenue ||
      a.raceNumber - b.raceNumber;
  });

  return races;
}



/**
 * Web画面の選択内容を既存の購入形式に変換する
 */
function convertWebBet_(data) {
  const type = data.betType;
  const method = data.method;
  const groups = data.groups;

  const ordered = ["馬単", "枠単", "3連単"].includes(type);
  const separator = ordered ? ">" : "-";

  const join = nums => nums.join(",");
  let actualMethod = method;
  let selection = "";

  if (method === "通常") {
    selection = groups.map(g => join(g)).join(separator);

  } else if (method === "ボックス") {
    selection = join(groups[0]);

  } else if (method === "フォーメーション") {
    selection = groups.map(g => join(g)).join(separator);

  } else if (method === "流し") {
    const axis = groups[0];
    const others = groups[1];

    if (["馬連", "枠連", "ワイド"].includes(type)) {
      actualMethod = "ながし";
      selection = `${join(axis)}-${join(others)}`;

    } else if (type === "3連複") {
      if (data.axisCount === 1) {
        actualMethod = "1頭軸ながし";
        selection = `${join(axis)}-${join(others)}`;

      } else if (data.axisCount === 2) {
        if (axis.length !== 2) {
          throw new Error("軸馬を2頭選択してください。");
        }

        actualMethod = "2頭軸ながし";
        selection =
          `${axis[0]}-${axis[1]}-${join(others)}`;
      }

    } else if (["馬単", "枠単"].includes(type)) {
      const position = data.axisPosition || "1";

      actualMethod = position === "2"
        ? "2着固定ながし"
        : "1着固定ながし";

      selection = position === "2"
        ? `${join(others)}>${join(axis)}`
        : `${join(axis)}>${join(others)}`;


} else if (type === "3連単") {
  const axisCount = Number(data.axisCount);

  // マルチの場合でも、基準となる着順を使用
  const position = data.axisPosition ||
    (axisCount === 2 ? "1-2" : "1");

  if (axisCount === 1) {
    if (axis.length !== 1) {
      throw new Error("軸馬は1頭選択してください。");
    }

    const methods = {
      "1": "1着固定ながし",
      "2": "2着固定ながし",
      "3": "3着固定ながし"
    };

    actualMethod = methods[position];

    if (!actualMethod) {
      throw new Error("固定着順が不正です。");
    }

    // 既存のexpandTripleFixed_では、
    // 1着固定以外は「相手一覧>軸」
    selection = position === "1"
      ? `${join(axis)}>${join(others)}`
      : `${join(others)}>${join(axis)}`;

    } else if (axisCount === 2) {
      if (axis.length !== 2) {
        throw new Error("軸馬は2頭選択してください。");
      }

      const methods = {
        "1-2": "1,2着固定ながし",
        "1-3": "1,3着固定ながし",
        "2-3": "2,3着固定ながし"
      };

      actualMethod = methods[position];

      if (!actualMethod) {
        throw new Error("固定着順が不正です。");
      }

      // axis[0]が先の固定着順、
      // axis[1]が後の固定着順を担当する
      if (position === "1-2") {
        selection =
          `${axis[0]}>${axis[1]}>${join(others)}`;

      } else if (position === "1-3") {
        selection =
          `${axis[0]}>${join(others)}>${axis[1]}`;

      } else {
        selection =
          `${join(others)}>${axis[0]}>${axis[1]}`;
      }

    } else {
      throw new Error("軸の頭数が不正です。");
    }


    } else {
      throw new Error("この券種では流しを利用できません。");
    }

  } else {
    throw new Error("未対応の購入方式です。");
  }

  return {
    betType: type,
    method: actualMethod,
    multi: Boolean(data.multi),
    selection: selection
  };
}

/**
 * 買い目展開・購入点数計算
 * スプレッドシートは変更しない
 */
function previewWebBet(data) {
  const deadline = checkWebRaceDeadline(data.raceId);
  if (deadline.closed) {
    throw new Error(
      "このレースは締切時刻を過ぎています。"
    );
  }
    const converted = convertWebBet_(data);

  const pointPerBet = Number(data.pointPerBet);

  if (!Number.isSafeInteger(pointPerBet) ||
      pointPerBet <= 0 ||
      pointPerBet % 100 !== 0) {
    throw new Error(
      "1点あたりの購入ptは100pt単位で指定してください。"
    );
  }

  const tickets = expandBets_(
    converted.betType,
    converted.method,
    converted.multi,
    converted.selection
  );
  validateWebTickets_(data, tickets);

  const totalPurchase = tickets.length * pointPerBet;

  if (!Number.isSafeInteger(totalPurchase)) {
    throw new Error("購入総額が大きすぎます。");
  }

  return {
    ...converted,
    pointPerBet: pointPerBet,
    ticketCount: tickets.length,
    totalPurchase: totalPurchase,
    tickets: tickets
  };
}


function validateWebTickets_(data, tickets) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("地方レース情報");

  if (!sheet || sheet.getLastRow() < 2) {
    throw new Error("レース情報を取得できません。");
  }

  const raceId = String(data.raceId || "");
  const today = Utilities.formatDate(
    new Date(), "Asia/Tokyo", "yyyyMMdd"
  );

  if (!raceId.startsWith(today)) {
    throw new Error("本日のレースを選択してください。");
  }

  const rows = sheet.getRange(
    2, 1, sheet.getLastRow() - 1, 11
  ).getDisplayValues();

  const row = rows.find(r =>
    String(r[1]).trim() === raceId
  );

  if (!row) {
    throw new Error("指定されたレースが見つかりません。");
  }

  const excluded = new Set(
    String(row[9] || "")
      .split(/[,，、]/)
      .map(s => Number(s.trim()))
      .filter(n => Number.isInteger(n) && n > 0)
  );

  const pairs = String(row[10] || "")
    .split(/[,，、]/)
    .map(s => s.trim().match(/^(\d+)\s*[:：]\s*(\d+)$/))
    .filter(Boolean)
    .map(m => ({
      horse: Number(m[1]),
      frame: Number(m[2])
    }));

  const frameCounts = Array(9).fill(0);
  const activeHorses = new Set();

  for (const pair of pairs) {
    if (!excluded.has(pair.horse)) {
      activeHorses.add(pair.horse);
      if (pair.frame >= 1 && pair.frame <= 8) {
        frameCounts[pair.frame]++;
      }
    }
  }

  const isFrame = ["枠連", "枠単"].includes(data.betType);

  if (isFrame && pairs.length === 0) {
    throw new Error("枠番情報が未取得です。");
  }

  // K列がない場合のみ掲載頭数を利用
  if (!isFrame && pairs.length === 0) {
    for (let n = 1; n <= Number(row[8]); n++) {
      if (!excluded.has(n)) activeHorses.add(n);
    }
  }

  for (const ticket of tickets) {
    const nums = ticket.split(/[-＞>]/).map(Number);

    if (isFrame) {
      for (const n of nums) {
        if (!frameCounts[n]) {
          throw new Error(`${n}枠には出走可能な馬がいません。`);
        }
      }

      if (
        nums.length === 2 &&
        nums[0] === nums[1] &&
        frameCounts[nums[0]] < 2
      ) {
        throw new Error(
          `${nums[0]}枠のゾロ目には有効な出走馬が2頭必要です。`
        );
      }
    } else {
      for (const n of nums) {
        if (!activeHorses.has(n)) {
          throw new Error(
            `${n}番は出走できない、または存在しない馬番です。`
          );
        }
      }
    }
  }
}


/**
 * Webアプリ用の参加者一覧を取得
 */
function getParticipantsForWeb() {
  const sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName("参加者情報");

  if (!sheet) {
    throw new Error("参加者情報シートが見つかりません。");
  }

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();

  if (lastRow < 2) return [];

  const headers = sheet.getRange(
    1, 1, 1, lastCol
  ).getDisplayValues()[0];

  const idIndex = headers.indexOf("参加者ID");
  const nameIndex = headers.indexOf("参加者名");

  if (idIndex < 0 || nameIndex < 0) {
    throw new Error(
      "参加者IDまたは参加者名の列が見つかりません。"
    );
  }

  const rows = sheet.getRange(
    2, 1, lastRow - 1, lastCol
  ).getDisplayValues();

  return rows
    .map(row => ({
      id: String(row[idIndex] || "").trim(),
      name: String(row[nameIndex] || "").trim()
    }))
    .filter(p => p.id && p.name);
}


/**
 * Webアプリ用：参加者の通常ptを取得
 */
function getWebParticipantPoints(participantId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const participants = ss.getSheetByName("参加者情報");
  const pointSheet = ss.getSheetByName("ポイント管理");

  if (!participants || !pointSheet) {
    throw new Error("参加者情報またはポイント管理シートがありません。");
  }

  const id = String(participantId || "").trim();

  if (!id) {
    throw new Error("参加者IDが指定されていません。");
  }

  // 参加者情報に登録されていることを確認
  const known = getParticipantsForWeb()
    .find(p => p.id === id);

  if (!known) {
    throw new Error("登録されていない参加者です。");
  }

  const lastRow = pointSheet.getLastRow();

  if (lastRow < 2) {
    throw new Error("ポイント管理データがありません。");
  }

  const rows = pointSheet.getRange(
    2, 1, lastRow - 1, 2
  ).getValues();

  const row = rows.find(r =>
    String(r[0]).trim() === id
  );

  if (!row) {
    throw new Error("参加者のポイント情報がありません。");
  }

  const normalPt = Number(row[1]);

  if (!Number.isFinite(normalPt)) {
    throw new Error("通常ptが不正な値です。");
  }

  return {
    participantId: id,
    participantName: known.name,
    normalPt: normalPt
  };
}


/**
 * Web購入のリクエストIDを登録する
 * 同じIDで2回購入しないための処理
 */
function claimWebPurchaseRequest_(requestId) {
  if (
    typeof requestId !== "string" ||
    !/^[a-zA-Z0-9_-]{10,100}$/.test(requestId)
  ) {
    throw new Error("購入リクエストIDが不正です。");
  }

  const key = "WEB_PURCHASE_" + requestId;
  const lock = LockService.getScriptLock();

  lock.waitLock(30000);

  try {
    const props = PropertiesService.getScriptProperties();

    if (props.getProperty(key) !== null) {
      throw new Error(
        "この購入リクエストは既に受け付けています。" +
        "購入履歴を確認してください。"
      );
    }

    props.setProperty(key, "処理中");
  } finally {
    lock.releaseLock();
  }

  return key;
}

/**
 * Webアプリから馬券を購入する
 */
function purchaseWebBet(data) {
  if (!data || typeof data !== "object") {
    throw new Error("購入データがありません。");
  }

  const participantId = String(
    data.participantId || ""
  ).trim();

  // 参加者として登録されているか確認
  const participant = getParticipantsForWeb()
    .find(p => p.id === participantId);

  if (!participant) {
    throw new Error("参加者が見つかりません。");
  }

  // 既存の計算処理で最新の買い目を検証
  const preview = previewWebBet(data);

  // 重複リクエストをブロック
  const requestKey = claimWebPurchaseRequest_(
    data.requestId
  );

  const props = PropertiesService.getScriptProperties();

  try {
    // 既存の購入・Discord通知処理を利用
    // 引数の並びは既存関数と一致しているか要確認
    const result = purchaseBetWithNotification(
      participantId,
      data.venue,
      data.raceNumber,
      preview.betType,
      preview.method,
      preview.multi,
      preview.selection,
      preview.pointPerBet
    );

    // 購入処理の成功・失敗を正しく判定
    if (!result || result.success !== true) {
      const message =
        result?.message ||
        result?.error ||
        "購入処理が正常に完了しませんでした。";

      throw new Error(message);
    }

    // ここから下は購入成功時のみ実行される
    const points = getWebParticipantPoints(participantId);

    props.setProperty(requestKey, "購入処理完了");

    return {
      success: true,
      participantName: participant.name,
      ticketCount: preview.ticketCount,
      totalPurchase: preview.totalPurchase,
      remainingPt: points.normalPt
    };

  } catch (error) {
    // エラーが発生しても同じリクエストIDは再実行しない
    // 購入後の通知処理などで例外が発生した可能性があるため
    props.setProperty(requestKey, "要確認");

    throw new Error(
      "購入処理を完了できませんでした。" +
      "購入履歴を確認してから再操作してください。" +
      "\n詳細：" + error.message
    );
  }
}


/**
 * レースの締切状態を取得
 * 締切時刻は発走時刻とする
 */
function checkWebRaceDeadline(raceId) {
  const sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName("地方レース情報");

  if (!sheet || sheet.getLastRow() < 2) {
    throw new Error("レース情報がありません。");
  }

  const rows = sheet.getRange(
    2, 1, sheet.getLastRow() - 1, 8
  ).getDisplayValues();

  const row = rows.find(r =>
    String(r[1]).trim() === String(raceId).trim()
  );

  if (!row) {
    throw new Error("指定したレースが見つかりません。");
  }

  const id = String(row[1]).trim();
  const datePart = id.substring(0, 8);
  const timeText = String(row[7] || "").trim();

  if (!/^\d{8}/.test(id)) {
    throw new Error("レースIDの日付が不正です。");
  }

  // 例：14:30、14:30:00 に対応
  const match = timeText.match(
    /(?:^|\s)(\d{1,2}):(\d{2})(?::(\d{2}))?(?:\s|$)/
  );

  if (!match) {
    throw new Error(
      "発走時刻を読み取れません：" + timeText
    );
  }

  const hour = Number(match[1]);
  const minute = Number(match[2]);
  const second = Number(match[3] || 0);

  if (
    hour > 23 ||
    minute > 59 ||
    second > 59
  ) {
    throw new Error("発走時刻が不正です。");
  }

  const dateString =
    datePart.substring(0, 4) + "-" +
    datePart.substring(4, 6) + "-" +
    datePart.substring(6, 8) + " " +
    String(hour).padStart(2, "0") + ":" +
    String(minute).padStart(2, "0") + ":" +
    String(second).padStart(2, "0");

  // 日本時間で解釈
  const deadline = Utilities.parseDate(
    dateString,
    "Asia/Tokyo",
    "yyyy-MM-dd HH:mm:ss"
  );

  const now = new Date();

  const today = Utilities.formatDate(
    now, "Asia/Tokyo", "yyyyMMdd"
  );

  const closed =
    datePart !== today ||
    now.getTime() >= deadline.getTime();

  return {
    raceId: id,
    closed: closed,
    deadlineText: timeText,
    deadlineMs: deadline.getTime(),
    serverNowMs: now.getTime()
  };
}


