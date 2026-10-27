
function diagnoseTournamentResults() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const tournamentSheet = changeSheet_(
    ss, CHANGE_CONFIG.tournamentSheet
  );

  const historySheet = changeSheet_(
    ss, CHANGE_CONFIG.historySheet
  );

  const resultSheet = ensureChangeResultSheet_(ss);

  const tournaments = readChangeTournaments_(
    tournamentSheet
  );

  // 最新の終了済み大会を診断
  const finished = tournaments.filter(t => t.finished);

  if (finished.length === 0) {
    Logger.log("終了済み大会がありません。");
    return;
  }

  const tournament = finished[finished.length - 1];

  Logger.log("対象大会：" + tournament.id);

  // 購入履歴から購入者を集計
  const historyRows = historySheet.getLastRow() >= 2
    ? historySheet.getRange(
        2, 1, historySheet.getLastRow() - 1, 11
      ).getValues()
    : [];

  const byTournament = new Map();

  for (const row of historyRows) {
    const tournamentId = String(row[1]).trim();
    const participantId = String(row[2]).trim();

    if (!byTournament.has(tournamentId)) {
      byTournament.set(tournamentId, new Set());
    }

    if (participantId) {
      byTournament.get(tournamentId).add(
        participantId
      );
    }
  }

  Logger.log("【購入履歴の大会別参加者】");

  for (const [id, members] of byTournament) {
    Logger.log(
      id + "：" + [...members].join(", ")
    );
  }

  // 保存済み結果を確認
  const savedRows = resultSheet.getLastRow() >= 2
    ? resultSheet.getRange(
        2, 1, resultSheet.getLastRow() - 1, 12
      ).getValues().filter(
        r => String(r[0]).trim() === tournament.id
      )
    : [];

  Logger.log("【保存済み大会結果】");

  for (const row of savedRows) {
    Logger.log(JSON.stringify({
      participantId: row[1],
      joined: row[2],
      normalPt: row[7],
      profitPt: row[8],
      bonusPt: row[9],
      totalPt: row[10],
      rank: row[11]
    }));
  }

  // 購入履歴の参加者と結果の参加判定を比較
  const historyEntrants =
    byTournament.get(tournament.id) || new Set();

  const savedEntrants = new Set(
    savedRows
      .filter(r => changeFlag_(r[2]))
      .map(r => String(r[1]).trim())
  );

  Logger.log("【比較結果】");

  Logger.log(
    "購入履歴の参加者：" +
    [...historyEntrants].join(", ")
  );

  Logger.log(
    "大会結果の参加者：" +
    [...savedEntrants].join(", ")
  );

  const missing = [...historyEntrants].filter(
    id => !savedEntrants.has(id)
  );

  Logger.log(
    "結果から漏れている参加者：" +
    (missing.length ? missing.join(", ") : "なし")
  );
}
