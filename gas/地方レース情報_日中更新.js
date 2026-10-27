/**
 * 地方レース情報_日中更新.gs
 * 既存の「地方レース情報_事前登録_競馬場ID対応.gs」と同じプロジェクトに追加する。
 *
 * - 4時台の updateLocalRacePreInfo トリガーはそのまま維持
 * - 30分ごとに事前情報（既存CSV）を再照合
 * - 「競馬場ID」A列=競馬場名、C列=NAR Web用コードを参照
 * - NARのRaceListから取消・除外を取得し「地方レース情報」J列を更新
 * - A列が「確定」の行、M～AE列は一切変更しない
 * - 取得失敗時は既存のJ列を維持する
 */
const LRD = {
  tz: 'Asia/Tokyo',
  sheet: '地方レース情報',
  log: '地方レース取得ログ',
  // 4時の初期登録は既存トリガーに任せ、日中更新は5～23時に限定
  startHour: 5,
  endHour: 23,
  host: 'https://www.keiba.go.jp/KeibaWeb/TodayRaceInfo/RaceList'
};

function lrdToday_() {
  return Utilities.formatDate(new Date(), LRD.tz, 'yyyy/MM/dd');
}

function lrdDate_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, LRD.tz, 'yyyy/MM/dd');
  return String(v).trim().replace(/-/g, '/');
}

function lrdClean_(html) {
  return String(html)
    .replace(/<script\b[^>]*>[\s\S]*?<\/script>/gi, '')
    .replace(/<style\b[^>]*>[\s\S]*?<\/style>/gi, '')
    .replace(/<br\s*\/?>/gi, ' ')
    .replace(/<[^>]+>/g, '')
    .replace(/&nbsp;|&#160;|&#xA0;/gi, ' ')
    .replace(/&amp;/gi, '&')
    .replace(/\u3000/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

/**
 * レース一覧HTMLを解析。
 * 出力 Map<レース番号, '5,10'>（取消・除外を区別しない）。
 * 「レース番号を含む本体行」が見つからなければエラーにして、
 * HTMLがメンテナンス画面だった場合にJ列を空欄で上書きしない。
 */
function lrdParseRaceList_(html, expectedRaceNumbers) {
  const source = String(html);
  const trs = source.match(/<tr\b[^>]*>[\s\S]*?<\/tr>/gi) || [];
  const changes = new Map();
  const recognizedRaces = new Set();
  let changeCandidates = 0;
  let parsedChanges = 0;

  for (const tr of trs) {
    const cells = (tr.match(/<(?:td|th)\b[^>]*>[\s\S]*?<\/(?:td|th)>/gi) || [])
      .map(lrdClean_);
    if (!cells.length) continue;
    const raceMatch = String(cells[0]).match(/^(\d{1,2})\s*R$/i);
    if (!raceMatch) continue;
    const raceNo = Number(raceMatch[1]);

    // レース一覧の本体行：2列目に発走時刻、頭数などが並ぶ
    if (/^\d{1,2}:\d{2}$/.test(String(cells[1] || ''))) {
      recognizedRaces.add(raceNo);
      if (!changes.has(raceNo)) changes.set(raceNo, []);
      continue;
    }

    // 変更情報の行：例 ["1R", "5", "馬名", "出走取消", ...]
    const joined = cells.join(' ');
    if (!/出走取消|競走除外/.test(joined)) continue;
    changeCandidates++;
    const horse = Number(cells[1]);
    const statusCell = cells.find(c => /出走取消|競走除外/.test(c));
    if (!Number.isInteger(horse) || horse < 1 || horse > 18 || !statusCell) {
      throw new Error('変更行の解析に失敗: ' + joined.slice(0, 160));
    }
    if (!changes.has(raceNo)) changes.set(raceNo, []);
    const entry = String(horse);
    if (!changes.get(raceNo).includes(entry)) changes.get(raceNo).push(entry);
    parsedChanges++;
  }

  if (recognizedRaces.size === 0) {
    throw new Error('正常なレース一覧を検出できません（メンテナンス・構造変更の可能性）');
  }
  if (changeCandidates !== parsedChanges) {
    throw new Error('変更情報の抽出件数が一致しません');
  }
  for (const n of expectedRaceNumbers) {
    if (!recognizedRaces.has(n)) {
      throw new Error('レース一覧に' + n + 'Rがありません。部分取得として更新中止');
    }
  }
  const out = new Map();
  for (const n of expectedRaceNumbers) {
    const entries = changes.get(n) || [];
    entries.sort((a, b) => Number(a) - Number(b));
    out.set(n, entries.join(','));
  }
  return out;
}

/**
 * 「競馬場ID」シートのA列（競馬場名）・C列（NAR Web用コード）を読み込む。
 * B列の略号と「地方レース情報」E列のCSV用コードは変更しない。
 * 例: 佐賀 → 32（CSV側では55）
 */
function lrdWebCodeMap_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('競馬場ID');
  if (!sheet) throw new Error('「競馬場ID」シートがありません');
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) throw new Error('「競馬場ID」シートが空です');

  const values = sheet.getRange(2, 1, lastRow - 1, 3).getDisplayValues();
  const map = new Map();
  const usedCodes = new Map();
  for (let i = 0; i < values.length; i++) {
    const [nameRaw, , codeRaw] = values[i];
    const name = String(nameRaw).trim();
    const codeText = String(codeRaw).trim();
    if (!name && !codeText) continue;
    if (!name) throw new Error('競馬場IDの' + (i + 2) + '行目に競馬場名がありません');
    // 略号だけ登録した競馬場（現在開催していない場）はC列空欄でも可。
    if (!codeText) continue;
    if (!/^\d+$/.test(codeText) || Number(codeText) <= 0) {
      throw new Error('競馬場IDのC' + (i + 2) + 'が不正: ' + codeText);
    }
    const code = Number(codeText);
    if (map.has(name)) throw new Error('競馬場IDの競馬場名が重複: ' + name);
    if (usedCodes.has(code)) {
      throw new Error('NAR Webコードが重複: ' + code +
        '（' + usedCodes.get(code) + ' / ' + name + '）');
    }
    map.set(name, code);
    usedCodes.set(code, name);
  }
  if (map.size === 0) throw new Error('競馬場IDのC列に有効なWebコードがありません');
  return map;
}

/**
 * 今日の全競馬場をレース情報シートから列挙して、各場のRaceListを取得する。
 * 取得失敗した競馬場はJ列を一切変更しない。
 */
function updateLocalRaceChanges() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) throw new Error('他のレース更新処理が実行中です');
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(LRD.sheet);
    if (!sheet) throw new Error('「地方レース情報」シートがありません');
    const last = sheet.getLastRow();
    if (last < 2) return;
    const rows = sheet.getRange(2, 1, last - 1, 12).getValues();
    const today = lrdToday_();
    const webCodeMap = lrdWebCodeMap_();
    const groups = new Map();
    for (let i = 0; i < rows.length; i++) {
      const row = rows[i];
      if (lrdDate_(row[2]) !== today || String(row[0]).trim() === '確定') continue;
      const csvCode = Number(row[4]);
      const race = Number(row[5]);
      const rawCourse = String(row[3]).trim();
      const courseName = rawCourse === '帯広ば' ? '帯広' : rawCourse;
      if (!Number.isInteger(csvCode) || csvCode <= 0 ||
          !Number.isInteger(race) || race < 1 || race > 99) {
        throw new Error('E/F列の競馬場コード・Rが不正: ' + (i + 2) + '行');
      }
      const webCode = webCodeMap.get(courseName);
      if (!webCode) {
        throw new Error('競馬場IDのC列にNAR Webコードが未登録: ' +
          courseName + '（地方レース情報 ' + (i + 2) + '行目）');
      }
      // Webコード単位ではなく競馬場名ごとにまとめる。
      // 同じ場なのにCSVコードが異なるなどの不整合は拒否する。
      if (!groups.has(courseName)) {
        groups.set(courseName, { webCode, csvCode, entries: [] });
      }
      const group = groups.get(courseName);
      if (group.webCode !== webCode || group.csvCode !== csvCode) {
        throw new Error('同一競馬場のコードが不一致: ' + courseName);
      }
      group.entries.push({
        sheetRow: i + 2, raceNo: race,
        oldJ: String(row[9] || '')
      });
    }

    let updated = 0, unchanged = 0, failedCourses = 0;
    for (const [courseName, group] of groups) {
      const { webCode, csvCode, entries } = group;
      const url = LRD.host + '?k_babaCode=' + encodeURIComponent(webCode) +
        '&k_raceDate=' + encodeURIComponent(today);
      try {
        const response = UrlFetchApp.fetch(url, {
          muteHttpExceptions: true, followRedirects: true
        });
        if (response.getResponseCode() !== 200) {
          throw new Error('HTTP ' + response.getResponseCode());
        }
        const parsed = lrdParseRaceList_(response.getContentText('UTF-8'),
          entries.map(e => e.raceNo));
        // パースが全件成功した競馬場だけ書き込む
        for (const item of entries) {
          const newJ = parsed.get(item.raceNo);
          if (newJ === item.oldJ) { unchanged++; continue; }
          sheet.getRange(item.sheetRow, 10).setValue(newJ); // Jのみ
          updated++;
        }
        console.log(courseName + '（CSV=' + csvCode + ', Web=' + webCode + '）: ' +
          entries.length + 'レース / J更新' +
          entries.filter(e => parsed.get(e.raceNo) !== e.oldJ).length + '件');
      } catch (error) {
        failedCourses++;
        console.error(courseName + '（CSV=' + csvCode + ', Web=' + webCode +
          '）の変更情報取得失敗: ' + error.message);
      }
    }
    console.log('取消・除外更新完了: 更新' + updated + ' / 変更なし' +
      unchanged + ' / 取得失敗場' + failedCourses);
    const log = ss.getSheetByName(LRD.log);
    if (log) log.appendRow([new Date(), today,
      failedCourses ? '一部失敗' : '成功', 0, updated, unchanged, 0,
      '日中J列更新 / 取得失敗場' + failedCourses]);
    if (failedCourses) throw new Error('変更情報を取得できなかった競馬場があります: ' + failedCourses + '場');
  } finally {
    lock.releaseLock();
  }
}

/**
 * 30分定期用。CSVの事前更新 → RaceListの変更情報取得。
 * 片方の取得が失敗しても、もう片方の処理は試行する。
 */
function updateLocalRaceDaytime() {
  const hour = Number(Utilities.formatDate(new Date(), LRD.tz, 'H'));
  if (hour < LRD.startHour || hour > LRD.endHour) return;
  let errors = [];
  try {
    updateLocalRacePreInfo();
  } catch (e) {
    console.error('CSV事前情報の更新失敗: ' + e.message);
    errors.push('CSV: ' + e.message);
  }
  try {
    updateLocalRaceChanges();
  } catch (e) {
    console.error('取消・除外情報の更新失敗: ' + e.message);
    errors.push('取消除外: ' + e.message);
  }
  if (errors.length) throw new Error(errors.join(' / '));
}

/** 30分トリガーを1つだけ作成。既存の朝4時台トリガーは維持。 */
function installLocalRaceDaytimeTrigger() {
  const handler = 'updateLocalRaceDaytime';
  for (const trigger of ScriptApp.getProjectTriggers()) {
    if (trigger.getHandlerFunction() === handler) ScriptApp.deleteTrigger(trigger);
  }
  ScriptApp.newTrigger(handler).timeBased().everyMinutes(30).create();
  console.log('30分ごとの日中更新を登録（実処理はJST 5～23時のみ）');
}

/** 日中更新トリガーのみ停止。 */
function stopLocalRaceDaytimeTrigger() {
  for (const trigger of ScriptApp.getProjectTriggers()) {
    if (trigger.getHandlerFunction() === 'updateLocalRaceDaytime') {
      ScriptApp.deleteTrigger(trigger);
    }
  }
}

/** 外部アクセスせず、ダミーのHTMLから変更情報の解析をテスト。 */
function testLocalRaceChangesParser() {
  const sample = '<table>' +
    '<tr><td>1R</td><td>14:35</td><td>有</td><td></td></tr>' +
    '<tr><td>2R</td><td>15:10</td><td>有</td><td></td></tr>' +
    '<tr><td>1R</td><td>5</td><td>テスト馬</td><td>出走取消</td><td>馬体故障</td></tr>' +
    '<tr><td>1R</td><td>7</td><td>テスト馬2</td><td>競走除外</td><td>疾病</td></tr>' +
    '</table>';
  const result = lrdParseRaceList_(sample, [1, 2]);
  if (result.get(1) !== '5,7' || result.get(2) !== '') {
    throw new Error('解析テスト失敗: ' + JSON.stringify([...result]));
  }
  console.log('PASS: J列解析ロジックの単体テスト');
}
