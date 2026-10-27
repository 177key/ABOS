/**
 * 地方レース情報：払戻登録（CSV取得時に即時確定）
 * 依存: 地方レース情報_事前登録_競馬場ID対応.gs
 *       地方レース情報_払戻CSV診断.gs
 *       (lrZipFiles_, lrParseCsv_, lpbAnalyzeRows_, LPB)
 *
 * payback.csvに正常な払戻情報が掲載された通常レースを即時確定登録する。
 * 公式の確定フラグは参照せず、CSV掲載と解析成功を登録条件にする運用仕様。
 * 公式2ページで取り止め・不成立が一致したレースは全返還として確定。
 * 特払は対象外。補助シート・観測履歴は作成しない。
 */
const LRF = {
  sheet: '地方レース情報',
  tz: 'Asia/Tokyo',
  statuses: ['未確定','確定'],
  typeOrder: ['単勝','複勝','馬連','馬単','ワイド','枠連','枠単','3連複','3連単']
};

function setupLocalRaceRefundSync() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss.getSheetByName(LRF.sheet)) throw new Error('地方レース情報がありません');
  const s = ss.getSheetByName(LRF.sheet);
  const expect = {1:'処理状態',2:'レースID',10:'取消・除外',11:'馬番・枠番対応'};
  for (const [col, name] of Object.entries(expect)) {
    if (String(s.getRange(1,Number(col)).getDisplayValue()).trim() !== name) {
      throw new Error('列構成不一致: '+col+'列は「'+name+'」を想定');
    }
  }
  console.log('払戻同期セットアップ完了（追加シート作成なし・既存データ変更なし）');
}

function lrfRows_(sheet,width) {
  return sheet.getLastRow()<2 ? [] : sheet.getRange(2,1,sheet.getLastRow()-1,width).getValues();
}
function lrfDay_(v) {
  if(v instanceof Date) return Utilities.formatDate(v,LRF.tz,'yyyyMMdd');
  const s=String(v||'').trim().replace(/[\/-]/g,'');
  if(!/^\d{8}$/.test(s)) throw new Error('開催日形式不正: '+v);
  return s;
}
function lrfClock_(v) {
  if(v instanceof Date) return Utilities.formatDate(v,LRF.tz,'HH:mm');
  const m=String(v||'').trim().match(/^(\d{1,2}):(\d{2})(?::\d{2})?$/);
  if(!m||+m[1]>23||+m[2]>59) throw new Error('発走時刻形式不正: '+v);
  return m[1].padStart(2,'0')+':'+m[2];
}
function lrfPaybackRows_() {
  if(typeof lrZipFiles_!=='function'||typeof lrParseCsv_!=='function'||typeof lpbAnalyzeRows_!=='function') {
    throw new Error('事前登録・払戻CSV診断の既存コードが必要です');
  }
  const files=lrZipFiles_();
  if(!files.payback) throw new Error('ZIPにpayback.csvがありません');
  const rows=lrParseCsv_(files.payback,1);
  const parsed=lpbAnalyzeRows_(rows);
  if(parsed.errors.length) throw new Error('CSVに解析対象外エラー: '+parsed.errors.slice(0,3).join(';'));
  return parsed.results;
}
function lrfIndex_(ss) {
  const s=ss.getSheetByName(LRF.sheet);
  if(!s) throw new Error('地方レース情報シートがありません');
  const records=lrfRows_(s,31);
  const byRaceId=new Map(), byNatural=new Map();
  const slugs=getCourseSlugMap_();
  for(let i=0;i<records.length;i++) {
    const r=records[i];
    const id=String(r[1]||'').trim();
    if(!id) continue;
    if(byRaceId.has(id)) throw new Error('重複レースID: '+id);
    byRaceId.set(id,{row:i+2, data:r});
    const course=lrNormalizeCourseName_(String(r[3]||'').trim());
    const key=course+'|'+lrfDay_(r[2])+'|'+Number(r[5]);
    if(byNatural.has(key)) throw new Error('開催日・場・Rの重複: '+key);
    byNatural.set(key,{id,row:i+2,data:r});
    const expected= lrfDay_(r[2]) + slugs.get(course) + String(Number(r[5])).padStart(2,'0');
    if(expected!==id) throw new Error('レースIDと競馬場略号の不整合: '+id+' / '+expected);
  }
  return {sheet:s,byRaceId,byNatural};
}
function lrfPayoutArrays_(result) {
  const winners=[], amounts=[];
  for(const type of LRF.typeOrder) {
    const entries=result.items[type]||[];
    winners.push(entries.map(x=>x.combination).join(','));
    amounts.push(entries.map(x=>x.payout).join(','));
  }
  return {winners,amounts};
}
/** HTMLの確定告知だけを抽出する。予告や曖昧な表現は除外。 */
function lrfNoticeText_(html) {
  return String(html).replace(/<script\b[^>]*>[\s\S]*?<\/script>/gi,' ')
    .replace(/<style\b[^>]*>[\s\S]*?<\/style>/gi,' ')
    .replace(/<!--[\s\S]*?-->/g,' ')
    .replace(/<[^>]*>/g,' ')
    .replace(/&nbsp;|&#160;|&#xA0;/gi,' ')
    .replace(/&amp;/gi,'&').replace(/\u3000/g,' ')
    .replace(/[\t\r\n]+/g,' ').replace(/\s+/g,' ').trim();
}
function lrfCanceledFromHtml_(html, course) {
  const text=lrfNoticeText_(html);
  const races=new Set(), evidence=[];
  for(const raw of text.split(/[。\n]/)) {
    const sentence=raw.trim();
    if(!sentence.includes(course) || !/(?:取り止め|取止め|競走不成立|不成立)/.test(sentence)) continue;
    if(/(?:予定|可能性|おそれ|場合|見込み|検討中|未定)/.test(sentence))continue;
    const end=sentence.search(/取り止め|取止め|競走不成立|不成立/);
    const relevant=sentence.slice(Math.max(0,end-110),end+16);
    const numbers=[...relevant.matchAll(/第\s*(\d{1,2})\s*競走/g)]
      .map(x=>Number(x[1]));
    if(!numbers.length || numbers.some(x=>x<1||x>99))continue;
    numbers.forEach(x=>races.add(x));
    evidence.push(sentence.slice(0,220));
  }
  return {races:[...races].sort((a,b)=>a-b),evidence};
}
function lrfFullRefundNotices_(course, webCode, today8) {
  const date=today8.slice(0,4)+'/'+today8.slice(4,6)+'/'+today8.slice(6,8);
  const fullDate=today8.slice(0,4)+'年'+Number(today8.slice(4,6))+'月'+Number(today8.slice(6,8))+'日';
  const results=[];
  for(const page of ['RaceList','RefundMoneyList']) {
    const url='https://www.keiba.go.jp/KeibaWeb/TodayRaceInfo/'+page+
      '?k_babaCode='+encodeURIComponent(webCode)+'&k_raceDate='+encodeURIComponent(date);
    const response=UrlFetchApp.fetch(url,{muteHttpExceptions:true,followRedirects:true,
      headers:{'User-Agent':'Mozilla/5.0'}});
    if(response.getResponseCode()!==200)throw new Error(page+' HTTP '+response.getResponseCode());
    const html=response.getContentText('UTF-8');
    const plain=lrfNoticeText_(html);
    if(plain.length<200 || !plain.includes(course) || !plain.includes(fullDate)) {
      throw new Error(page+' 日付・競馬場のページ確認に失敗');
    }
    results.push(lrfCanceledFromHtml_(html,course));
  }
  // 一方のページしか告知していないケースは安全側に倒して未確定維持。
  if(JSON.stringify(results[0].races)!==JSON.stringify(results[1].races)) {
    throw new Error('公式2ページの取り止め・不成立情報が一致しません: '+
      results[0].races+' / '+results[1].races);
  }
  return new Set(results[0].races);
}
/** K列「馬番:枠番,...」の全馬番を使う。I列の頭数を連番化しない。 */
function lrfAllHorseNumbers_(value) {
  const parts=String(value||'').split(',').map(x=>x.trim()).filter(Boolean);
  if(!parts.length)throw new Error('K列の馬番・枠番対応が空です');
  const numbers=new Set();
  for(const part of parts) {
    const match=part.match(/^(\d{1,2})\s*:\s*(\d{1,2})$/);
    if(!match)throw new Error('K列形式不正: '+part);
    const horse=Number(match[1]),frame=Number(match[2]);
    if(horse<1||horse>18||frame<1||frame>8||numbers.has(horse)) {
      throw new Error('K列の馬番・枠番が不正または重複: '+part);
    }
    numbers.add(horse);
  }
  return [...numbers].sort((a,b)=>a-b).join(',');
}
/**
 * シートに記録する全返還の内部表現。
 * M～Uは空欄、V～ADは全券種100（未発売券種も100の番兵値）。
 * 実際の払戻種類は判別できていないため精算側で券種発売を別途検査する。
 */
function lrfFullRefundData_(found) {
  return {horses:lrfAllHorseNumbers_(found.data[10]),
    winners:Array(9).fill(''),amounts:Array(9).fill(100)};
}
/** 今日の未確定レースについて公式2ページを取得。場ごとに1回。 */
function lrfFullRefundCandidates_(index,today,knownResults) {
  if(typeof lrdWebCodeMap_!=='function')throw new Error('日中更新コードの lrdWebCodeMap_ が必要です');
  const map=lrdWebCodeMap_(), groups=new Map(),actions=[],errors=[];
  for(const found of index.byRaceId.values()) {
    if(String(found.data[0]).trim()==='確定'||lrfDay_(found.data[2])!==today)continue;
    const course=lrNormalizeCourseName_(String(found.data[3]||'').trim());
    const no=Number(found.data[5]);
    if(!groups.has(course))groups.set(course,[]);
    groups.get(course).push({found,no});
  }
  for(const [course,entries] of groups) {
    const webCode=map.get(course);
    if(!webCode) {errors.push(course+': NAR Webコードがありません');continue;}
    let canceled;
    try {canceled=lrfFullRefundNotices_(course,webCode,today);}
    catch(e) {errors.push(course+': '+e.message);continue;}
    for(const {found,no} of entries) {
      if(!canceled.has(no))continue;
      const record=knownResults.get(found.id);
      // CSVに通常払戻情報があるなら矛盾。勝手にどちらかを採用しない。
      if(record && record.totalHitEntries>0) {
        errors.push(found.id+': 全返還告知と通常払戻CSVが競合');continue;
      }
      try {
        actions.push({kind:'全返還',id:found.id,row:found.row,
          data:lrfFullRefundData_(found)});
      } catch(e) {errors.push(found.id+': '+e.message);}
    }
  }
  return {actions,errors};
}

/**
 * dryRun=true: CSVと既存シートを照合するだけ（書き込みなし）。
 * dryRun=false: 有効な払戻情報を取得した未確定レースを即時登録。
 * 注意: CSV掲載は公式確定フラグの直接確認ではない。
 */
function lrfSync_(dryRun) {
  const lock=LockService.getScriptLock();
  if(!lock.tryLock(30000))throw new Error('別の更新処理が実行中です');
  try {
    const ss=SpreadsheetApp.getActiveSpreadsheet();
    const index=lrfIndex_(ss);
    const results=lrfPaybackRows_();
    const now=new Date(), today=Utilities.formatDate(now,LRF.tz,'yyyyMMdd');
    const byId=new Map(), actions=[];
    let held=0, errors=0, finalCount=0, alreadyFinal=0;
    for(const result of results) {
      if(result.date8!==today)continue;
      const course=lrNormalizeCourseName_(result.course);
      const found=index.byNatural.get(course+'|'+result.date8+'|'+result.no);
      if(!found) { console.error('シートに見当たらないレース: '+course+' '+result.no+'R'); errors++; continue; }
      if(byId.has(found.id))throw new Error('CSV内で同一レースが重複: '+found.id);
      byId.set(found.id,result);
    }
    for(const [id,result] of byId) {
      const found=index.byRaceId.get(id);
      if(String(found.data[0]).trim()==='確定') {alreadyFinal++;continue;}
      if(result.errors.length) {
        console.error('払戻データ解析エラー '+id+': '+result.errors.join(';'));
        errors++;continue;
      }
      if(result.totalHitEntries===0) {
        held++;
        console.log('特殊状態の可能性があるため保留: '+id);
        continue;
      }
      // 特払、発売なし等による空欄は0円とみなさず、そのまま空欄で保存する。
      // 未確定レースで有効な払戻データを初めて取得した時点で登録候補にする。
      actions.push({kind:'通常',row:found.row,id,data:lrfPayoutArrays_(result)});
      finalCount++;
    }
    const refunds=lrfFullRefundCandidates_(index,today,byId);
    refunds.errors.forEach(e=>console.error('全返還検証: '+e));
    // 該当レースが公式2ページで明示された場合は、通常の候補から外す。
    const canceledIds=new Set(refunds.actions.map(a=>a.id));
    for(let i=actions.length-1;i>=0;i--) {
      if(canceledIds.has(actions[i].id))actions.splice(i,1);
    }
    actions.push(...refunds.actions);
    // 全返還候補の解析エラーがあれば、その回の登録を中止する。
    errors+=refunds.errors.length;
    console.log('払戻照合 / CSV='+results.length+' / 登録候補='+actions.length+
      '（通常'+actions.filter(a=>a.kind==='通常').length+
      '・全返還'+refunds.actions.length+'） / 保留='+held+' / 確定済み除外='+alreadyFinal+' / エラー='+errors);
    if(dryRun) {
      actions.slice(0,15).forEach(a=>console.log('DRY RUN '+a.kind+' '+a.id));
      console.log('読み取り専用確認完了（シート・プロパティの書き込みなし）');
      return;
    }
    // CSVの照合が全件成功した場合のみ書き込みを開始する。
    if(errors)throw new Error('CSV照合にエラーがあるため登録中止。ログを確認: '+errors+'件');
    let written=0, refundWritten=0, skippedAlreadyFinal=0;
    for(const a of actions) {
      const row=a.row;
      // A列=確定になったレースは更新しない
      if(String(index.sheet.getRange(row,1).getDisplayValue()).trim()==='確定') {skippedAlreadyFinal++;continue;}
      if(a.kind==='全返還') {
        index.sheet.getRange(row,10).setValue(a.data.horses); // J
      }
      index.sheet.getRange(row,13,1,18)
        .setValues([[...a.data.winners,...a.data.amounts]]);
      index.sheet.getRange(row,31).setValue(now); // AE
      // すべての書き込みが完了した後で確定にする
      index.sheet.getRange(row,1).setValue('確定');
      written++;
      if(a.kind==='全返還')refundWritten++;
    }
    console.log('登録完了 通常='+(written-refundWritten)+' / 全返還='+refundWritten+
      ' / 直前確定によるスキップ='+skippedAlreadyFinal);
  } finally {lock.releaseLock();}
}
function previewLocalRaceRefundSync(){lrfSync_(true);}
function updateLocalRaceRefundSync(){lrfSync_(false);}
/**
 * 10分間隔の監視トリガー。CSV掲載時点で即時登録する運用仕様。
 */
function installLocalRaceRefundTrigger(){
  const handler='updateLocalRaceRefundSync';
  for(const t of ScriptApp.getProjectTriggers()) {
    if(t.getHandlerFunction()===handler)ScriptApp.deleteTrigger(t);
  }
  ScriptApp.newTrigger(handler).timeBased().everyMinutes(10).create();
  console.log('払戻同期10分トリガー登録（CSV掲載時点で即時登録）');
}
function stopLocalRaceRefundTrigger(){
  for(const t of ScriptApp.getProjectTriggers()){
    if(t.getHandlerFunction()==='updateLocalRaceRefundSync') ScriptApp.deleteTrigger(t);
  }
}
