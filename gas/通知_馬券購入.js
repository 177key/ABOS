/*************************************************
 * 馬券購入通知
 *************************************************/

/**
 * 馬券購入成功時にDiscordへ通知する
 *
 * @param {Object} purchase 購入情報
 * @param {Object} result purchaseBetの実行結果
 */
function notifyBetPurchase_(purchase, result) {
  if (!result.success) {
    return {
      success: false,
      reason: "購入失敗のため通知しません"
    };
  }

  const {
    participantId,
    venue,
    raceNumber,
    betType,
    method,
    multi,
    selection,
    pointPerBet
  } = purchase;

  const formatPt = n => Number(n).toLocaleString("ja-JP");

  const embed = {
    title: "馬券を購入しました！",
    color: 3447003,

    description:
      `**${venue} ${raceNumber}R**`,

    fields: [
      {
        name: "馬券ID",
        value: String(result.betId),
        inline: true
      },
      {
        name: "券種",
        value: betType,
        inline: true
      },
      {
        name: "方式",
        value: method,
        inline: true
      },
      {
        name: "買い目",
        value: "`" + selection + "`",
        inline: false
      },
      {
        name: "マルチ",
        value: multi ? "あり" : "なし",
        inline: true
      },
      {
        name: "購入点数",
        value: `${result.count}点`,
        inline: true
      },
      {
        name: "1点あたり",
        value: `${formatPt(pointPerBet)}pt`,
        inline: true
      },
      {
        name: "購入総額",
        value: `**${formatPt(result.total)}pt**`,
        inline: true
      },
      {
        name: "購入後の通常pt",
        value: `**${formatPt(result.remaining)}pt**`,
        inline: true
      }
    ],

    footer: {
      text: `レースID：${result.raceId}`
    },

    timestamp: new Date().toISOString()
  };

  // 参加者IDを通知先として使用
  const notificationId = enqueueDiscordNotification(
    participantId,
    { embeds: [embed] }
  );

  return {
    success: true,
    notificationId: notificationId
  };
}


/**
 * 馬券購入とDiscord通知をまとめて実行
 *
 * 通知失敗でも購入は取り消さない。
 */
function purchaseBetWithNotification(
  participantId,
  venue,
  raceNumber,
  betType,
  method,
  multi,
  selection,
  pointPerBet
) {
  // 1. 既存の購入処理
  const result = purchaseBet(
    participantId,
    venue,
    raceNumber,
    betType,
    method,
    multi,
    selection,
    pointPerBet
  );

  // 購入失敗なら通知しない
  if (!result.success) {
    return result;
  }

  // 2. 購入成功後に通知
  const purchase = {
    participantId,
    venue,
    raceNumber,
    betType,
    method,
    multi,
    selection,
    pointPerBet
  };

  try {
    const notification = notifyBetPurchase_(
      purchase, result
    );

    result.notification = {
      queued: true,
      notificationId: notification.notificationId
    };

  } catch (e) {
    // Discord送信失敗でも購入は成功扱い
    console.error(
      `馬券${result.betId}の通知に失敗：${e.message}`
    );

    result.notification = {
      success: false,
      error: e.message
    };
  }

  return result;
}
