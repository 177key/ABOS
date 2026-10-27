
/**
 * 現在開催中の大会情報を取得する
 *
 * 前提：
 * ・大会情報は時系列順
 * ・大会は重複しない
 * ・前大会の終了日の翌日に次大会が始まる
 *
 * @return {Object} 現在の大会情報
 */
function getCurrentTournament() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("大会情報");

  if (!sheet) {
    throw new Error("大会情報シートが見つかりません");
  }

  const timezone = ss.getSpreadsheetTimeZone();
  const today = Utilities.formatDate(
    new Date(), timezone, "yyyyMMdd"
  );

  const lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    throw new Error("大会情報が登録されていません");
  }

  const data = sheet.getRange(
    2, 1, lastRow - 1, 6
  ).getValues();

  for (let i = 0; i < data.length; i++) {
    const row = data[i];

    const tournamentId = String(row[0]).trim();
    if (!tournamentId) continue;

    const rawEndDate = row[2];
    let endDate;

    if (rawEndDate instanceof Date) {
      endDate = Utilities.formatDate(
        rawEndDate, timezone, "yyyyMMdd"
      );
    } else {
      const match = String(rawEndDate).trim().match(
        /^(\d{4})[/-](\d{1,2})[/-](\d{1,2})$/
      );

      if (!match) {
        throw new Error(
          `大会情報${i + 2}行目の終了日が不正です`
        );
      }

      endDate =
        match[1] +
        match[2].padStart(2, "0") +
        match[3].padStart(2, "0");
    }

    // 終了日をまだ過ぎていない最初の大会
    if (endDate >= today) {
      const initialPt = Number(row[3]);
      const weeklyPt = Number(row[4]);
      const bonusPt = Number(row[5]);

      if (
        [initialPt, weeklyPt, bonusPt].some(
          n => !Number.isSafeInteger(n) || n < 0
        )
      ) {
        throw new Error(
          `大会${tournamentId}のポイント設定が不正です`
        );
      }

      return {
        tournamentId,
        tournamentName: String(row[1]).trim(),
        endDate,
        initialPt,
        weeklyPt,
        bonusPt,
        rowNumber: i + 2
      };
    }
  }

  throw new Error(
    "現在開催中の大会が見つかりません。" +
    "次の大会情報を登録してください"
  );
}
