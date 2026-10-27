
const WEEKLY_NOTICE_SHEET = "週間通知履歴";
const WEEKLY_TZ = "Asia/Tokyo";

// 月曜日のトリガーから呼び出す関数
function runMondayWeeklyNotifications() {
  const now = new Date();
  if (now.getDay() !== 1) {
    throw new Error("月曜日以外には実行できません。");
  }

  // 既存の配布処理。二重配布防止機能はそのまま利用
  distributeWeeklyPoints();

  // 配布履歴が正常に記録されていることを確認
  const tournament = getCurrentTournament();
  const dateKey = weeklyDateKey_(now);
  const distribution = findWeeklyDistribution_(
    dateKey, tournament.tournamentId
  );

  if (!distribution) {
    Logger.log("配布実績がないため、通知を行いません。");
    return;
  }

  // 個別通知の登録
  const members = getWeeklyMembers_();
  for (const member of members) {
    const embed = {
      title: "週間ポイント配布",
      color: 0x3498db,
      description:
        `今週のポイントが配布されました！\n\n` +
        `## +${weeklyFmt_(tournament.weeklyPt)}pt\n` +
        `**配布後の通常pt：` +
        `${weeklyFmt_(member.normalPt)}pt**`,
      footer: { text: tournament.tournamentName }
    };

    // 総合ランキング
    enqueueWeeklyNoticeOnce_(
      dateKey,
      tournament.tournamentId,
      "総合",
      { embeds: [buildWeeklyProgressEmbed_(tournament)] }
    );

    // 最高オッズ記録
    enqueueWeeklyNoticeOnce_(
      dateKey,
      tournament.tournamentId,
      "最高オッズ",
      { embeds: [buildWeeklyRecordEmbed_(tournament, "odds")] }
    );

    // 最高利益pt記録
    enqueueWeeklyNoticeOnce_(
      dateKey,
      tournament.tournamentId,
      "最高利益ポイント",
      { embeds: [buildWeeklyRecordEmbed_(tournament, "profit")] }
    );

Logger.log("週間通知の登録処理が完了しました。");

  }

  // 全員分を登録してから総合ランキングを登録
  const progress = buildWeeklyProgressEmbed_(tournament);
  enqueueWeeklyNoticeOnce_(
    dateKey, tournament.tournamentId,
    "総合", { embeds: [progress] }
  );

  Logger.log("週間通知の登録処理が完了しました。");
}

// 今日の日付を YYYY-MM-DD で取得
function weeklyDateKey_(date) {
  return Utilities.formatDate(
    date, WEEKLY_TZ, "yyyy-MM-dd"
  );
}

function weeklyFmt_(value) {
  return Number(value).toLocaleString("ja-JP");
}

function weeklyNum_(value) {
  return Number(String(value).replace(/,/g, "")) || 0;
}

// 配布履歴を確認
function findWeeklyDistribution_(dateKey, tournamentId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("定期配布履歴");
  if (!sheet || sheet.getLastRow() < 2) return null;

  const rows = sheet.getRange(
    2, 1, sheet.getLastRow() - 1, 6
  ).getValues();

  return rows.find(row => {
    const date = row[0] instanceof Date
      ? weeklyDateKey_(row[0])
      : String(row[0]).replace(/\//g, "-");

    const status = String(row[4]);
    return date === dateKey &&
      String(row[5]) === String(tournamentId) &&
      !/失敗|エラー|中止/.test(status);
  }) || null;
}

// 参加者と現在の通常ptを取得
function getWeeklyMembers_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName("ポイント管理");

  const rows = sheet.getDataRange().getDisplayValues();

  return rows.slice(1)
    .filter(row => row[0])
    .map(row => ({
      id: String(row[0]).trim(),
      normalPt: weeklyNum_(row[1]),
      profitPt: weeklyNum_(row[2])
    }));
}

// 当該大会で馬券を購入した参加者のID集合
function getWeeklyParticipants_(tournamentId) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName("馬券購入履歴");

  const ids = new Set();
  if (!sheet || sheet.getLastRow() < 2) return ids;

  const rows = sheet.getRange(
    2, 2, sheet.getLastRow() - 1, 2
  ).getDisplayValues();

  for (const row of rows) {
    if (String(row[0]) === String(tournamentId)) {
      ids.add(String(row[1]).trim());
    }
  }
  return ids;
}

// ランキング（同点同順位）
function getWeeklyRanking_(tournamentId) {
  const buyers = getWeeklyParticipants_(tournamentId);

  const ranked = getWeeklyMembers_()
    .filter(member => buyers.has(member.id))
    .map(member => ({
      ...member,
      total: member.normalPt + member.profitPt
    }))
    .sort((a, b) =>
      b.total - a.total ||
      a.id.localeCompare(b.id)
    );

  let previousScore = null;
  let previousRank = 0;

  return ranked.map((member, index) => {
    const rank = member.total === previousScore
      ? previousRank : index + 1;

    previousScore = member.total;
    previousRank = rank;

    return { ...member, rank };
  });
}


// 総合チャンネル向け途中経過

function buildWeeklyProgressEmbed_(tournament) {
  const names = getWeeklyParticipantNames_();
  const ranking = getWeeklyRanking_(
    tournament.tournamentId
  );

  const today = Utilities.formatDate(
    new Date(), WEEKLY_TZ, "yyyy/MM/dd"
  );

  const rankLines = ranking.map(member => {
    const mark = member.rank === 1 ? "🥇"
      : member.rank === 2 ? "🥈"
      : member.rank === 3 ? "🥉" : "";

    const displayName = names.get(member.id) || member.id;

    return `${mark}${member.rank}位　` +
      `${displayName}　${weeklyFmt_(member.total)}pt`;
  });

  let rankText = rankLines.length
    ? rankLines.join("\n")
    : "まだ購入者がいません。";

  if (rankText.length > 2500) {
    rankText = rankLines.slice(0, 20).join("\n") +
      "\n（以下省略）";
  }

  return {
    title: `大会途中経過｜${today}`,
    color: 0xf1c40f,
    description: `**${tournament.tournamentName}**`,
    fields: [{
      name: "暫定総合ランキング",
      value: rankText.slice(0, 1024),
      inline: false
    }],
    footer: {
      text: "通常pt＋利益pt｜ボーナスptは未確定"
    },
    timestamp: new Date().toISOString()
  };
}


// 通知キューへの二重登録を防ぐための履歴
function getWeeklyNoticeSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(WEEKLY_NOTICE_SHEET);

  if (!sheet) {
    sheet = ss.insertSheet(WEEKLY_NOTICE_SHEET);
    sheet.appendRow([
      "通知キー", "通知先", "状態", "更新日時"
    ]);
  }
  return sheet;
}

function enqueueWeeklyNoticeOnce_(
  dateKey, tournamentId, destination, message
) {
  const key = [
    dateKey, tournamentId, destination
  ].join("|");

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  let sheet;
  let targetRow;

  try {
    sheet = getWeeklyNoticeSheet_();
    const last = sheet.getLastRow();

    if (last >= 2) {
      const keys = sheet.getRange(
        2, 1, last - 1, 1
      ).getDisplayValues().flat();

      if (keys.includes(key)) {
        Logger.log(`通知登録済み・要確認：${key}`);
        return;
      }
    }

    targetRow = sheet.getLastRow() + 1;
    sheet.appendRow([
      key, destination, "登録処理中", new Date()
    ]);
  } finally {
    lock.releaseLock();
  }

  try {
    enqueueDiscordNotification(destination, message);
    sheet.getRange(targetRow, 3, 1, 2)
      .setValues([["登録済", new Date()]]);
  } catch (error) {
    sheet.getRange(targetRow, 3, 1, 2)
      .setValues([["要確認", new Date()]]);
    throw error;
  }
}


function buildWeeklyRecordEmbed_(tournament, recordType) {
  const names = getWeeklyParticipantNames_();
  const sheet = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName("ボーナス管理");

  const isOdds = recordType === "odds";
  const startRow = isOdds ? 3 : 15;

  const rows = sheet.getRange(
    startRow, 1, 9, 4
  ).getDisplayValues();

  const lines = rows.map(row => {
    const betType = row[0];
    const holders = row[1]
      ? row[1]
          .split(/\r?\n/)
          .map(id => id.trim())
          .filter(Boolean)
          .map(id => names.get(id) || id)
          .join("・")
      : "—";

    if (!row[1] || row[2] === "") {
      return `**${betType}**：記録なし`;
    }

    const value = weeklyNum_(row[2]);

    const recordValue = isOdds
      ? `${value}倍`
      : `${weeklyFmt_(value)}pt`;

    return `**${betType}**：${recordValue}` +
      `（${holders}）`;
  });

  const today = Utilities.formatDate(
    new Date(), WEEKLY_TZ, "yyyy/MM/dd"
  );

  return {
    title: isOdds
      ? `最高オッズ｜週間途中経過 ${today}`
      : `最高利益pt｜週間途中経過 ${today}`,
    color: isOdds ? 0xf1c40f : 0x9b59b6,
    description:
      `**${tournament.tournamentName}**\n\n` +
      lines.join("\n"),
    timestamp: new Date().toISOString()
  };
}


/**
 * 参加者ID → 表示名 の対応表を取得する
 */
function getWeeklyParticipantNames_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName("参加者情報");

  if (!sheet) {
    throw new Error("参加者情報シートがありません。");
  }

  const rows = sheet.getDataRange().getDisplayValues();
  const headers = rows[0].map(h => h.trim());

  const idCol = headers.indexOf("参加者ID");
  const nameCol = headers.indexOf("参加者名");

  if (idCol < 0 || nameCol < 0) {
    throw new Error(
      "参加者情報シートに「参加者ID」「参加者名」の列が必要です。"
    );
  }

  const names = new Map();

  for (const row of rows.slice(1)) {
    const id = String(row[idCol]).trim();
    const name = String(row[nameCol]).trim();

    if (id) {
      names.set(id, name || id);
    }
  }

  return names;
}


