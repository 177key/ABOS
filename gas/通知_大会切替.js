
/**
 * 大会切替通知
 *
 * checkTournamentChange() がスクリプトロックを
 * 保持している間に呼び出す専用関数。
 *
 * enqueueDiscordNotification() は使わない。
 */

// 通知イベントから固定のUUID形式IDを生成
function changeNoticeId_(eventKey) {
  const bytes = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    "TOURNAMENT_NOTICE_V1:" + eventKey,
    Utilities.Charset.UTF_8
  ).slice(0, 16).map(b => (b + 256) % 256);

  // UUIDのバージョン・バリアント位置を整える
  bytes[6] = (bytes[6] & 15) | 80;
  bytes[8] = (bytes[8] & 63) | 128;

  const h = bytes.map(b =>
    b.toString(16).padStart(2, "0")
  ).join("");

  return [
    h.slice(0, 8),
    h.slice(8, 12),
    h.slice(12, 16),
    h.slice(16, 20),
    h.slice(20, 32)
  ].join("-");
}

/**
 * 通知を一度だけキューに登録する。
 *
 * 必ずcheckTournamentChange()のロック内で呼ぶ。
 * この関数自身ではロックを取得しない。
 */
function enqueueChangeNoticeOnce_(
  ss, eventKey, destination, message
) {
  const sheet = ss.getSheetByName(NOTIFY_QUEUE_SHEET);

  if (!sheet) {
    throw new Error("通知キューシートがありません");
  }

  const noticeId = changeNoticeId_(eventKey);
  const lastRow = sheet.getLastRow();

  if (lastRow >= 2) {
    const existing = sheet.getRange(
      2, 1, lastRow - 1, 1
    ).getDisplayValues();

    if (existing.some(
      row => String(row[0]).trim() === noticeId
    )) {
      console.log("通知登録済み：" + eventKey);
      return false;
    }
  }

  const json = JSON.stringify(message);

  if (json.length > 40000) {
    throw new Error("通知内容が大きすぎます");
  }

  sheet.appendRow([
    noticeId,
    destination,
    json,
    "未送信",
    new Date(),
    "",
    "",
    ""
  ]);

  SpreadsheetApp.flush();

  console.log("通知を登録：" + eventKey);
  return true;
}

/**
 * 参加者IDから表示名を取得
 */
function getChangeDisplayNames_(ss) {
  const sheet = ss.getSheetByName("参加者情報");
  if (!sheet || sheet.getLastRow() < 2) {
    return {};
  }

  const values = sheet.getDataRange()
    .getDisplayValues();

  const headers = values[0];
  const idCol = headers.indexOf("参加者ID");
  const nameCol = headers.indexOf("参加者名");

  if (idCol < 0 || nameCol < 0) {
    throw new Error(
      "参加者情報のID列・名前列が見つかりません"
    );
  }

  const names = {};

  for (const row of values.slice(1)) {
    const id = String(row[idCol] || "").trim();
    const name = String(row[nameCol] || "").trim();

    if (id) names[id] = name || id;
  }

  return names;
}

/**
 * 保存済みの大会結果を通知用の形式に復元する。
 * ポイント管理シートは変更しない。
 */
function getSavedFinalChangeResult_(ss, tournament) {
  const resultSheet = ensureChangeResultSheet_(ss);
  const accounts = getChangeAccounts_(ss);

  const rows = readSavedChangeResults_(
    resultSheet, tournament.id, accounts
  );

  if (!rows) {
    throw new Error(
      "大会結果が保存されていません：" + tournament.id
    );
  }

  const participants = rows
    .filter(r => r[2])
    .sort((a, b) => Number(a[11]) - Number(b[11]))
    .map(r => ({
      participantId: String(r[1]),
      rank: Number(r[11]),
      normalPt: Number(r[7]),
      profitPt: Number(r[8]),
      bonusPt: Number(r[9]),
      awardedBonusPt: Number(r[6]),
      totalPt: Number(r[10])
    }));

  return {
    tournamentId: tournament.id,
    participantCount: participants.length,
    results: participants
  };
}

/**
 * 最終結果発表Embedを作る
 */
function buildChangeFinalMessage_(ss, finishResult) {
  const names = getChangeDisplayNames_(ss);
  const pt = n => Number(n).toLocaleString("ja-JP");

  const lines = finishResult.results.map(r => {
    const name = String(
      names[r.participantId] || r.participantId
    ).replace(/[\r\n]/g, " ").slice(0, 30);

    const prefix =
      r.rank === 1 ? "🥇" :
      r.rank === 2 ? "🥈" :
      r.rank === 3 ? "🥉" : "・";

    return (
      `${prefix} ${r.rank}位 ${name}：` +
      `${pt(r.totalPt)}pt`
    );
  });

  // Discord Embedの文字数制限への対応
  const displayed = [];
  let length = 0;

  for (const line of lines) {
    if (length + line.length + 1 > 3400) break;

    displayed.push(line);
    length += line.length + 1;
  }

  const remaining = lines.length - displayed.length;

  const rankingText = displayed.length
    ? displayed.join("\n")
    : "参加者はいません。";

  const description = (
    "大会の最終結果が確定しました！\n" +
    "最高記録ボーナス反映後の順位です。\n\n" +
    rankingText +
    (remaining > 0
      ? `\n\nほか ${remaining}名（大会結果シート参照）`
      : "")
  );

  return {
    embeds: [{
      title: "🏆 エア馬券大会・最終結果発表",
      description,
      color: 0xF1C40F,
      fields: [
        {
          name: "大会ID",
          value: String(finishResult.tournamentId),
          inline: true
        },
        {
          name: "参加人数",
          value: `${finishResult.participantCount}名`,
          inline: true
        }
      ],
      footer: {
        text: "最終順位・ポイントは確定済みです"
      }
    }]
  };
}

/**
 * 新大会開始通知Embed
 */
function buildChangeStartMessage_(tournament) {
  const pt = n => Number(n).toLocaleString("ja-JP");

  return {
    embeds: [{
      title: "🏁 新しい大会がスタートしました",
      description:
        "前大会の集計が完了し、" +
        "新しい大会が始まりました！\n\n" +
        `初期通常pt：**${pt(tournament.initialPt)}pt**\n` +
        "利益pt・ボーナスpt：**0pt**\n" +
        "最高オッズ・最高利益pt記録：リセット済み\n\n" +
        "新しい大会もよろしくお願いします！",
      color: 0x2ECC71,
      fields: [{
        name: "大会ID",
        value: String(tournament.id),
        inline: true
      }]
    }]
  };
}

/**
 * 最終結果発表をキューへ登録
 */
function enqueueChangeFinalNotice_(
  ss, tournament, finishResult
) {
  return enqueueChangeNoticeOnce_(
    ss,
    "FINAL:" + tournament.id,
    "総合",
    buildChangeFinalMessage_(ss, finishResult)
  );
}

/**
 * 次大会開始通知をキューへ登録
 */
function enqueueChangeStartNotice_(ss, tournament) {
  return enqueueChangeNoticeOnce_(
    ss,
    "START:" + tournament.id,
    "総合",
    buildChangeStartMessage_(tournament)
  );
}
