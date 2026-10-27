
/**
 * 指定シートの古いデータ行を定期削除する。
 *
 * ・3000行を超えたら古い1000行を削除
 * ・1行目のヘッダーは残す
 * ・上から古い順に並んでいることを前提とする
 */

const SHEET_CLEANUP_CONFIG = {
  maxRows: 3000,
  deleteRows: 1000,

  sheets: [
    "通知キュー",
    "馬券購入履歴",
    "地方レース情報",
    "中央レース情報",
    "定期配布履歴",
    "地方レース取得ログ",
    "中央レース取得ログ"
  ]
};

function cleanupOldSheetRows() {
  const lock = LockService.getScriptLock();

  // 他の処理が動いていたら今回はスキップ
  if (!lock.tryLock(1000)) {
    console.log("他の処理が実行中のためスキップ");
    return;
  }

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    for (const name of SHEET_CLEANUP_CONFIG.sheets) {
      const sheet = ss.getSheetByName(name);

      if (!sheet) {
        console.log("シートが存在しません: " + name);
        continue;
      }

      const lastRow = sheet.getLastRow();

      if (lastRow <= SHEET_CLEANUP_CONFIG.maxRows) {
        continue;
      }

      // ヘッダーを除く先頭500行を削除
      sheet.deleteRows(
        2,
        SHEET_CLEANUP_CONFIG.deleteRows
      );

      console.log(
        name + ": " +
        lastRow + "行 → " +
        sheet.getLastRow() + "行"
      );
    }

    console.log("シート行数の整理が完了しました");

  } finally {
    lock.releaseLock();
  }
}
