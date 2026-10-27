/**
 * 地方レース情報：NAR公式ZIPのCSVによる当日事前情報登録
 * A～Lのみ更新。M～AEは絶対に書き換えない。
 * J(取消・除外)は今回のCSVだけでは得られないため、既存値を維持する。
 */
const LR = {
  sheet: '地方レース情報',
  log: '地方レース取得ログ',
  tz: 'Asia/Tokyo',
  url: 'https://www.keiba.go.jp/KeibaWeb/DataDownload/RaceDataDownload?type=daily',
  // 競馬場の略号は「競馬場ID」シート A列/B列から読み込む。
  code: {
    '帯広ば':3, '帯広':3, '門別':36, '盛岡':10, '水沢':11,
    '浦和':18, '船橋':19, '大井':20, '川崎':21, '金沢':22,
    '笠松':42, '名古屋':48, '園田':50, '姫路':51, '高知':31, '佐賀':55
  },
  headers: ['処理状態','レースID','開催日','競馬場','競馬場コード','R','レース名',
            '発走時刻','掲載頭数','取消・除外','馬番・枠番対応','事前情報最終更新']
};


/**
 * 「競馬場ID」のA列=競馬場名、B列=3文字略号。
 * 更新実行ごとに1度だけ読み込む。
 */
function getCourseSlugMap_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('競馬場ID');
  if (!sheet) throw new Error('「競馬場ID」シートがありません');
  if (sheet.getLastRow() < 2) throw new Error('「競馬場ID」シートが空です');

  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getDisplayValues();
  const map = new Map();
  const usedSlugs = new Map();
  for (const [rawName, rawSlug] of values) {
    const name = String(rawName).trim();
    const slug = String(rawSlug).trim();
    if (!name && !slug) continue;
    if (!name || !/^[a-z]{3}$/.test(slug)) {
      throw new Error('競馬場IDの不正な行: ' + name + ' / ' + slug);
    }
    if (map.has(name)) throw new Error('競馬場名が重複しています: ' + name);
    if (usedSlugs.has(slug)) {
      throw new Error('競馬場略号が重複: ' + slug + ' (' + usedSlugs.get(slug) + ', ' + name + ')');
    }
    map.set(name, slug);
    usedSlugs.set(slug, name);
  }
  if (map.size === 0) throw new Error('競馬場IDに有効な登録がありません');
  return map;
}

/** NAR側表記との軽微な差を吸収する。 */
function lrNormalizeCourseName_(name) {
  if (name === '帯広ば') return '帯広';
  return name;
}

function setupLocalRaceInfo() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const s = ss.getSheetByName(LR.sheet);
  if (!s) throw new Error('既存の「地方レース情報」シートが見つかりません');
  const headers = s.getRange(1,1,1,12).getDisplayValues()[0];
  const mismatches = LR.headers.map((x,i) => headers[i] === x ? null :
    `${String.fromCharCode(65+i)}列: 「${headers[i]}」(想定「${x}」)` ).filter(Boolean);
  // 既存シートの見出しと意味が合致しているか、ここでは重要な列のみ検証する。
  if (headers[1] !== 'レースID' || headers[2] !== '開催日' ||
      headers[3] !== '競馬場' || headers[10] !== '馬番・枠番対応') {
    throw new Error('列の並びが異なります。確認: ' + mismatches.join(' / '));
  }
  // 自動変換でレースIDや時間が壊れないよう文字列形式を指定。
  s.getRange('B2:D').setNumberFormat('@');
  s.getRange('H2:H').setNumberFormat('@');
  s.getRange('K2:K').setNumberFormat('@');
  s.getRange('L2:L').setNumberFormat('yyyy/MM/dd HH:mm:ss');
  let log = ss.getSheetByName(LR.log);
  if (!log) {
    log = ss.insertSheet(LR.log);
    log.appendRow(['処理日時','対象日','結果','追加','更新','変更なし','スキップ','備考']);
    log.setFrozenRows(1);
  }
  console.log('セットアップ完了。既存データは変更していません');
}

function lrZipFiles_() {
  const res = UrlFetchApp.fetch(LR.url, {
    muteHttpExceptions:true, followRedirects:true
  });
  if (res.getResponseCode() !== 200)
    throw new Error('ZIP取得失敗 HTTP '+res.getResponseCode());
  const blob = res.getBlob();
  const bytes = blob.getBytes();
  if (bytes.length < 4 || (bytes[0]&255)!==0x50 || (bytes[1]&255)!==0x4b)
    throw new Error('ZIPではないデータを受信しました（メンテナンス画面等の可能性）');
  const files = Utilities.unzip(blob);
  const csv = {};
  for (const f of files) {
    const name = f.getName().toLowerCase();
    for (const type of ['racelist','horselist','payback']) {
      if (name.endsWith('_'+type+'.csv')) csv[type] = f;
    }
  }
  if (!csv.racelist || !csv.horselist)
    throw new Error('ZIP内にracelist.csv / horselist.csvが見つかりません');
  return csv;
}

function lrParseCsv_(blob, minimumFields) {
  // 通常UTF-8。日本語の置換文字が多ければShift_JISでも試す。
  let text = blob.getDataAsString('UTF-8').replace(/^\uFEFF/,'');
  if ((text.match(/\uFFFD/g)||[]).length > 3) {
    text = blob.getDataAsString('Shift_JIS').replace(/^\uFEFF/,'');
  }
  return Utilities.parseCsv(text)
    .filter(row => row.length>=minimumFields && row.some(v=>String(v).trim()));
}

function lrDate_(raw) {
  const x=String(raw).trim();
  if (!/^\d{8}$/.test(x)) throw new Error('日付が不正: '+x);
  return x.slice(0,4)+'/'+x.slice(4,6)+'/'+x.slice(6,8);
}

function lrTime_(raw) {
  const x=String(raw).trim().padStart(4,'0');
  if (!/^\d{4}$/.test(x) || +x.slice(0,2)>23 || +x.slice(2)>59)
    throw new Error('発走時刻が不正: '+raw);
  return x.slice(0,2)+':'+x.slice(2);
}

function lrKey_(course,date8,no) { return `${course}|${date8}|${no}`; }

function lrBuildData_() {
  const files = lrZipFiles_();
  const races = lrParseCsv_(files.racelist,27);
  const horses = lrParseCsv_(files.horselist,6);
  const today8=Utilities.formatDate(new Date(),LR.tz,'yyyyMMdd');
  const courseSlugMap=getCourseSlugMap_();
  const map=new Map();
  for (const h of horses) {
    const course=String(h[0]).trim(), date8=String(h[1]).trim();
    const no=Number(h[2]), frame=Number(h[3]), horse=Number(h[5]);
    if (date8!==today8 || !Number.isInteger(no)) continue;
    if (!Number.isInteger(frame)||frame<1||frame>8||
        !Number.isInteger(horse)||horse<1||horse>18)
      throw new Error(`出馬表の枠番/馬番が不正: ${course} ${no}R`);
    const k=lrKey_(course,date8,no);
    if (!map.has(k)) map.set(k,new Map());
    if (map.get(k).has(horse)) throw new Error('馬番重複 '+k+' '+horse);
    map.get(k).set(horse,frame);
  }
  const result=[],seen=new Set();
  for (const r of races) {
    const course=String(r[0]).trim(), date8=String(r[1]).trim();
    if (date8!==today8) continue;
    const no=Number(r[2]);
    if (!Number.isInteger(no)||no<1||no>99) throw new Error('レース番号不正');
    const courseName=lrNormalizeCourseName_(course);
    const slug=courseSlugMap.get(courseName);
    if (!slug) throw new Error('「競馬場ID」に未登録の競馬場: '+course+'（正規化後: '+courseName+'）');
    if (!LR.code[course]) throw new Error('NARコードが未登録の競馬場: '+course);
    const k=lrKey_(course,date8,no);
    if(seen.has(k)) throw new Error('レース一覧重複: '+k);
    seen.add(k);
    const h=map.get(k);
    if (!h || !h.size) throw new Error('出馬表未掲載: '+k);
    const num=Number(r[26]);
    if (!Number.isInteger(num)||num<1||num>18||h.size<num)
      throw new Error(`頭数の不整合: ${k} 一覧=${num} 出馬表=${h.size}`);
    const pairs=[...h.entries()].sort((a,b)=>a[0]-b[0]);
    pairs.forEach(([horse,frame],i)=>{
      if (horse!==i+1 || (i>0 && frame<pairs[i-1][1]))
        throw new Error('馬番/枠番の並びが不正: '+k);
    });
    const id=date8+slug+String(no).padStart(2,'0');
    result.push({id, date:lrDate_(date8), course, code:LR.code[course],
      no, name:String(r[5]).trim(), time:lrTime_(r[3]), count:num,
      mapping:pairs.map(([horse,frame])=>horse+':'+frame).join(',')});
  }
  if (!result.length) throw new Error('当日のレースデータがありません');
  return result;
}

/**
 * 当日レースのB～Lだけを更新。Aの確定・M～AEの払戻情報は保護。
 * 取消/除外のJ列は既存値を維持（CSVに対応項目がないため）。
 */
function updateLocalRacePreInfo() {
  const lock=LockService.getScriptLock();
  if (!lock.tryLock(30000)) throw new Error('別の更新処理が実行中');
  try {
    const ss=SpreadsheetApp.getActiveSpreadsheet();
    const s=ss.getSheetByName(LR.sheet);
    if (!s) throw new Error('先にsetupLocalRaceInfo()を実行');
    const items=lrBuildData_(); // 全件の整合性を確認してから書く
    const last=s.getLastRow();
    const existing=last>=2?s.getRange(2,1,last-1,12).getValues():[];
    const ids=new Map();
    existing.forEach((r,i)=>{
      const id=String(r[1]).trim();
      if (!id) return;
      if (ids.has(id)) throw new Error('既存レースIDが重複: '+id);
      ids.set(id,i+2);
    });
    const now=new Date();
    let added=0,updated=0,unchanged=0,skipped=0;
    const toAppend=[];
    for (const x of items) {
      const row=ids.get(x.id);
      if (row) {
        const old=s.getRange(row,1,1,12).getValues()[0];
        if (String(old[0]).trim()==='確定') {skipped++;continue;}
        const target=[x.id,x.date,x.course,x.code,x.no,x.name,
          x.time,x.count,old[9],x.mapping]; // B～K（Jは維持）
        const oldDisplay=s.getRange(row,2,1,10).getDisplayValues()[0];
        const comparable=target.map(v=>String(v));
        if (comparable.every((v,i)=>v===oldDisplay[i])) {unchanged++;continue;}
        // B～Kだけ更新。Lは実際に変化した時刻。
        s.getRange(row,2,1,10).setValues([target]);
        s.getRange(row,12).setValue(now);
        updated++;
      } else {
        toAppend.push(['未確定',x.id,x.date,x.course,x.code,x.no,x.name,
          x.time,x.count,'',x.mapping,now]);
        added++;
      }
    }
    if (toAppend.length) {
      const start=Math.max(2,s.getLastRow()+1);
      s.getRange(start,1,toAppend.length,12).setValues(toAppend);
    }
    const log=ss.getSheetByName(LR.log);
    if(log) log.appendRow([new Date(),Utilities.formatDate(now,LR.tz,'yyyy/MM/dd'),
      '成功',added,updated,unchanged,skipped,'取消・除外(J)は今回自動更新しない']);
    console.log(`事前情報更新完了: 新規${added} / 更新${updated} / 変更なし${unchanged} / 確定保護${skipped}`);
  } catch(err) {
    const log=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LR.log);
    if(log) log.appendRow([new Date(),'', '失敗',0,0,0,0,String(err.message)]);
    throw err;
  } finally {lock.releaseLock();}
}

function installLocalRaceMorningTrigger() {
  const handler = "updateLocalRacePreInfo";

  // 既存の同じ関数のトリガーを削除
  for (const trigger of ScriptApp.getProjectTriggers()) {
    if (trigger.getHandlerFunction() === handler) {
      ScriptApp.deleteTrigger(trigger);
    }
  }

  // 毎朝4時台に実行
  ScriptApp.newTrigger(handler)
    .timeBased()
    .everyDays(1)
    .atHour(4)
    .inTimezone("Asia/Tokyo")
    .create();

  console.log("毎朝4時台の自動取得を登録しました");
}
