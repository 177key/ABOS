
/*************************************************
 * レース精算・最高記録 Discord通知
 *
 * 既存関数：
 * settleRace(raceId)
 * enqueueDiscordNotification(destination, message)
 *************************************************/

const SETTLE_NOTIFY_TYPES = [
  "単勝", "複勝", "馬連",
  "馬単", "ワイド", "枠連",
  "枠単", "3連複", "3連単"
];


/**
 * レース精算＋通知キュー登録
 *
 * 精算が成功した場合のみ通知する。
 * 通知登録失敗でも精算結果は維持する。
 */
function settleRaceWithNotification(raceId) {
  const result = settleRace(raceId);

  if (!result.success) {
    return result;
  }

  const notificationResults = [];

  // レース名を取得
  let raceLabel;

  try {
    raceLabel = getSettleNotifyRaceLabel_(
      result.raceId
    );
  } catch (e) {
    console.error("レース名取得失敗：" + e.message);
    raceLabel = result.raceId;
  }

  // 参加者別の精算通知
  for (const participant of result.participants) {
    try {
      const messages = buildSettleNotifyMessages_(
        participant, raceLabel
      );

      for (const message of messages) {
        const id = enqueueDiscordNotification(
          participant.participantId,
          message
        );

        notificationResults.push({
          destination: participant.participantId,
          success: true,
          notificationId: id
        });
      }
    } catch (e) {
      console.error(
        `${participant.participantId}の通知登録失敗：` +
        e.message
      );

      notificationResults.push({
        destination: participant.participantId,
        success: false,
        error: e.message
      });
    }
  }

  // 最高記録更新通知
  let bonusMessages = [];

  try {
    bonusMessages = buildSettleBonusNotices_(result);
  } catch (e) {
    console.error("最高記録通知の生成に失敗：" + e.message);

    notificationResults.push({
      destination: "最高記録",
      success: false,
      error: e.message
    });
  }

  for (const notice of bonusMessages) {
    try {
      const id = enqueueDiscordNotification(
        notice.destination,
        notice.message
      );

      notificationResults.push({
        destination: notice.destination,
        success: true,
        notificationId: id
      });
    } catch (e) {
      console.error(
        "最高記録通知登録失敗：" + e.message
      );

      notificationResults.push({
        destination: notice.destination,
        success: false,
        error: e.message
      });
    }
  }

  return {
    ...result,
    notificationResults
  };
}


/**
 * レース情報から表示名を取得する
 */
function getSettleNotifyRaceLabel_(raceId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("地方レース情報");

  if (!sheet) {
    throw new Error("地方レース情報がありません");
  }

  const last = sheet.getLastRow();

  if (last < 2) {
    throw new Error("レース情報が空です");
  }

  const rows = sheet.getRange(
    2, 1, last - 1, 7
  ).getValues();

  for (const row of rows) {
    if (String(row[1]).trim() !== raceId) continue;

    const venue = String(row[3]).trim();
    const number = Number(row[5]);
    const name = String(row[6] ?? "").trim();

    return `${venue} ${number}R` +
      (name ? ` ${name}` : "");
  }

  throw new Error("レースIDが見つかりません");
}


/**
 * 参加者別の精算通知を作成
 *
 * 馬券が多い場合は複数メッセージに分割。
 */
function buildSettleNotifyMessages_(participant, raceLabel) {
  const fmt = n => Number(n).toLocaleString("ja-JP");
  const signed = n => (n > 0 ? "+" : "") + fmt(n);

  const bets = participant.bets;

  const totalReturned =
    participant.payoutTotal + participant.refundTotal;

  const color = participant.netProfit >= 0
    ? 0x2ecc71 : 0xe74c3c;

  const normalPt = participant.normalPtAfterSettlement;

  // ① レース情報と購入・払戻
  const summary = {
    title: `${raceLabel}｜精算結果`,
    color: color,
    fields: [
      {
        name: "購入pt",
        value: `${fmt(participant.purchaseTotal)}pt`,
        inline: true
      },
      {
        name: "払戻pt",
        value: `${fmt(participant.payoutTotal)}pt`,
        inline: true
      }
    ]
  };

  // 返還があった場合だけ追加
  if (participant.refundTotal > 0) {
    summary.fields.push({
      name: "返還pt",
      value: `${fmt(participant.refundTotal)}pt`,
      inline: true
    });
  }

  const profitSummary = {
    title: `収支：${signed(participant.netProfit)}pt`,
    description:
      `**精算後の通常pt：${fmt(participant.normalPtAfterSettlement)}pt**`,
    color: color
  };

  // 2つのEmbedを1つのメッセージにまとめる
  const messages = [{
    embeds: [
      summary,
      profitSummary
    ]
  }];


  // ===== 個別馬券の精算結果 =====

  // 複数の馬券を1つのEmbedにまとめる
  const detailLines = [];

  for (let i = 0; i < participant.bets.length; i++) {
    const bet = participant.bets[i];

    const isMulti =
      bet.multi === true ||
      String(bet.multi).toUpperCase() === "TRUE";

    const multiLabel = isMulti ? "（マルチ）" : "";
    const methodLabel =
      bet.method && bet.method !== "通常"
        ? `${bet.method} `
        : "";

    const betLabel =
      `${i + 1}. ${bet.betType}｜` +
      `${methodLabel}${bet.selection}${multiLabel}`;

    const lines = [`**${betLabel}**`];

    // 的中した買い目
    if (bet.hitTickets && bet.hitTickets.length > 0) {
      for (const ticket of bet.hitTickets) {
        lines.push(
          `🎯 **的中：${ticket.selection}**`
        );
      }
    }

    // 返還された買い目
    if (bet.refundTickets && bet.refundTickets.length > 0) {
      lines.push(
        `↩️ 返還：${bet.refundTickets.join("、")}`
      );
    }

    // 的中も返還もない場合
    if (
      (!bet.hitTickets || bet.hitTickets.length === 0) &&
      (!bet.refundTickets || bet.refundTickets.length === 0)
    ) {
      lines.push("不的中");
    }

    // 購入ptと的中払戻pt
    lines.push(
      `購入 ${fmt(bet.effectivePurchase)}pt` +
      ` → 払戻 ${fmt(bet.hitPayout)}pt`
    );

    // 返還ptがある場合だけ追加
    if (bet.refundTotal > 0) {
      lines.push(
        `返還 ${fmt(bet.refundTotal)}pt`
      );
    }

    detailLines.push(lines.join("\n"));
  }

  // Embedのdescriptionには4096文字の制限があるため、
  // 長くなった場合は複数のメッセージへ分割する
  const separator = "\n\n────────────\n\n";
  const MAX_DESCRIPTION_LENGTH = 3500;

  let currentDescription = "";
  let detailPage = 1;

  function addDetailMessage(description, page) {
    const detailEmbed = {
      title:
        `購入馬券の結果｜${raceLabel}` +
        (page > 1 ? `（${page}）` : ""),
      color: 0x95a5a6,
      description: description
    };

    messages.push({
      embeds: [detailEmbed]
    });
  }

  for (const detail of detailLines) {
    const nextDescription = currentDescription
      ? currentDescription + separator + detail
      : detail;

    if (
      currentDescription &&
      nextDescription.length > MAX_DESCRIPTION_LENGTH
    ) {
      addDetailMessage(currentDescription, detailPage);
      detailPage++;
      currentDescription = detail;
    } else {
      currentDescription = nextDescription;
    }
  }

  if (currentDescription) {
    addDetailMessage(currentDescription, detailPage);
  }

  return messages;

}


/**
 * ボーナス更新結果から通知を作る
 *
 * 更新前の値と更新後の値を比較し、
 * 今回の精算で記録した参加者だけを抽出する。
 */
function buildSettleBonusNotices_(result) {
  const notices = [];
  const candidates = result.bonusCandidates || [];
  const updates = result.bonusUpdates || [];
  const participantNames = getWeeklyParticipantNames_();
  const toDisplayName = id => {
    const key = String(id).trim();
    return participantNames.get(key) || key;
  };

  // 更新後の値だけでなく、更新前のシート状態が必要。
  // 精算後のシートには既に新しい値が反映されているため、
  // candidatesとupdatesを使って今回の記録保持者を特定。
  for (const update of updates) {
    const row = update.rowNumber;

    let destination;
    let category;
    let index;

    if (row >= 3 && row <= 11) {
      destination = "最高オッズ";
      category = "最高オッズ";
      index = row - 3;
    } else if (row >= 15 && row <= 23) {
      destination = "最高利益ポイント";
      category = "最高利益pt";
      index = row - 15;
    } else {
      continue;
    }

    const betType = SETTLE_NOTIFY_TYPES[index];
    const recordValue = Number(update.value);

    const holderIds = String(update.holders)
      .split("\n").map(s => s.trim());

    const holderRaces = String(update.races)
      .split("\n").map(s => s.trim());

    const newHolders = [];

    for (const candidate of candidates) {
      if (candidate.betType !== betType) continue;

      const value = category === "最高オッズ"
        ? candidate.oddsValue
        : candidate.profitValue;

      if (value !== recordValue) continue;

      const exists = holderIds.some((id, i) =>
        id === candidate.participantId &&
        holderRaces[i] === candidate.raceLabel
      );

      if (exists) {
        newHolders.push(candidate);
      }
    }

    if (newHolders.length === 0) continue;

    const fmt = n => Number(n).toLocaleString("ja-JP");

    for (const holder of newHolders) {
      const valueText = category === "最高オッズ"
        ? `${fmt(recordValue)}倍`
        : `${fmt(recordValue)}pt`;

      notices.push({
        destination,
        message: {
          embeds: [{
            title: `${category}記録更新！`,
            color: category === "最高オッズ"
              ? 0xf1c40f : 0x9b59b6,
            fields: [
              {
                name: "券種",
                value: betType,
                inline: true
              },
              {
                name: "記録",
                value: `**${valueText}**`,
                inline: true
              },
              {
                name: "記録保持者",
                value: toDisplayName(holder.participantId),
                inline: false
              },
              {
                name: "対象レース",
                value: holder.raceLabel,
                inline: false
              }
            ],
            timestamp: new Date().toISOString()
          }]
        }
      });
    }
  }

  return notices;
}
