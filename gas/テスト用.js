function diagnoseLocalRaceRefundSource() {
  const results = lrfPaybackRows_();

  const today = Utilities.formatDate(
    new Date(),
    LRF.tz,
    "yyyyMMdd"
  );

  Logger.log("実行日：" + today);
  Logger.log("CSV取得件数：" + results.length);

  const dates = {};

  for (const result of results) {
    const date = String(result.date8 || "");
    dates[date] = (dates[date] || 0) + 1;
  }

  Logger.log(
    "取得データの日付別件数：" +
    JSON.stringify(dates)
  );

  // レース名などの詳細は先頭5件だけ
  for (const result of results.slice(0, 5)) {
    Logger.log(JSON.stringify({
      date: result.date8,
      course: result.course,
      raceNumber: result.no,
      hitEntries: result.totalHitEntries,
      errors: result.errors
    }));
  }
}



function diagnosePaybackCsvStages() {
  // ① ZIPからCSVを取得
  const files = lrZipFiles_();

  if (!files.payback) {
    throw new Error("payback.csvがありません");
  }

  Logger.log("【ZIP取得】");
  Logger.log(
    "payback.csv のデータ型：" +
    typeof files.payback
  );

  // ② CSVを解析
  const rows = lrParseCsv_(files.payback, 1);

  Logger.log("【CSV読み込み】");
  Logger.log("行数：" + rows.length);

  // 先頭数行の構造を確認
  Logger.log(
    "先頭3行：" +
    JSON.stringify(rows.slice(0, 3))
  );

  // ③ 払戻データへ変換
  const parsed = lpbAnalyzeRows_(rows);

  Logger.log("【払戻解析】");
  Logger.log(
    "解析結果件数：" + parsed.results.length
  );
  Logger.log(
    "解析エラー件数：" + parsed.errors.length
  );

  Logger.log(
    "解析エラー先頭3件：" +
    JSON.stringify(parsed.errors.slice(0, 3))
  );

  Logger.log(
    "解析結果先頭3件：" +
    JSON.stringify(parsed.results.slice(0, 3))
  );
}

function diagnosePaybackDates() {
  const files = lrZipFiles_();
  const rows = lrParseCsv_(files.payback, 1);

  const counts = {};

  for (const row of rows) {
    const date = String(row[1] || "").trim();

    if (/^\d{8}$/.test(date)) {
      counts[date] = (counts[date] || 0) + 1;
    }
  }

  Logger.log(
    "CSV内の日付別レコード数：" +
    JSON.stringify(counts)
  );

  Logger.log(
    "現在のJST日付：" +
    Utilities.formatDate(
      new Date(), "Asia/Tokyo", "yyyyMMdd"
    )
  );
}


function diagnosePaybackCsvContents() {
  const files = lrZipFiles_();

  if (!files.payback) {
    throw new Error("payback.csvがありません");
  }

  const rows = lrParseCsv_(files.payback, 1);

  Logger.log("CSV総行数：" + rows.length);

  // CSVの先頭5行を確認
  rows.slice(0, 5).forEach((row, index) => {
    Logger.log(
      "行" + (index + 1) +
      " / 列数=" + row.length +
      " / 内容=" + JSON.stringify(row)
    );
  });

  // 2列目の日付候補を調査
  const dateSamples = [...new Set(
    rows.slice(0, 30)
      .map(r => String(r[1] || "").trim())
  )];

  Logger.log(
    "2列目の値の例：" +
    JSON.stringify(dateSamples)
  );
}

