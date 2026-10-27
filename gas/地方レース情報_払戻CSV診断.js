/**
 * 地方レース情報：NAR payback.csv 診断専用
 * シートへの書き込み・トリガー作成・確定処理はしない。
 * 既存の「地方レース情報_事前登録」内の LR, lrZipFiles_,
 * lrParseCsv_, getCourseSlugMap_, lrNormalizeCourseName_ を利用する。
 */
const LPB = {
  tz: 'Asia/Tokyo',
  maxLogRaces: 12,
  // payback.csv 1行 = 54項目（NAR公式説明書の2026年版）
  specs: [
    {name:'単勝',  count:1, items:[{n:[4], p:5}]},
    {name:'複勝',  count:1, items:[{n:[7],p:8},{n:[10],p:11},{n:[13],p:14}]},
    {name:'馬連',  count:2, items:[{n:[24,25],p:26}]},
    {name:'馬単',  count:2, items:[{n:[28,29],p:30}]},
    {name:'ワイド',count:2, items:[{n:[32,33],p:34},{n:[36,37],p:38},{n:[40,41],p:42}]},
    {name:'枠連',  count:2, items:[{n:[16,17],p:18}]},
    {name:'枠単',  count:2, items:[{n:[20,21],p:22}]},
    {name:'3連複', count:3, items:[{n:[44,45,46],p:47}]},
    {name:'3連単', count:3, items:[{n:[49,50,51],p:52}]}
  ]
};

function lpbToday_() {
  return Utilities.formatDate(new Date(), LPB.tz, 'yyyyMMdd');
}
function lpbRaceKey_(course, date8, raceNo) {
  return course + '|' + date8 + '|' + raceNo;
}
function lpbCleanNumber_(v) {
  return String(v == null ? '' : v).trim().replace(/[,，]/g, '');
}
function lpbPositiveInt_(raw, label) {
  const s = lpbCleanNumber_(raw);
  if (!/^\d+$/.test(s) || Number(s) <= 0 || !Number.isSafeInteger(Number(s))) {
    throw new Error(label + ' が不正: ' + JSON.stringify(raw));
  }
  return Number(s);
}
function lpbExtractOne_(row, item, typeName) {
  const numbers = item.n.map(i => lpbCleanNumber_(row[i]));
  const payout = lpbCleanNumber_(row[item.p]);
  if (numbers.every(x => x === '') && payout === '') return null;
  if (numbers.some(x => x === '') || payout === '') {
    throw new Error(typeName + ' の組番と払戻額が一部欠けています');
  }
  const nums = numbers.map((x, i) => lpbPositiveInt_(x, typeName+'組番'+i));
  const amount = lpbPositiveInt_(payout, typeName+'払戻');
  if (nums.some(n => n > 18)) throw new Error(typeName+' 組番が18を超えています');
  if ((typeName==='枠連'||typeName==='枠単') && nums.some(n=>n>8)) {
    throw new Error(typeName+' 枠番が8を超えています');
  }
  return {numbers:nums, payout:amount};
}
function lpbAnalyzeRows_(rows) {
  const grouped = new Map();
  const errors = [];
  let skippedNonToday = 0;
  const today = lpbToday_();
  for (let idx=0;idx<rows.length;idx++) {
    const row=rows[idx];
    const course=String(row[0]||'').trim();
    const date8=String(row[1]||'').trim();
    const no=Number(String(row[2]||'').trim());
    // 見出しや当日以外のレコードは読み飛ばす
    if (date8 !== today) { skippedNonToday++; continue; }
    if (!course || !Number.isInteger(no) || no<1 || no>99) {
      errors.push('CSV行'+(idx+1)+': レース識別情報が不正'); continue;
    }
    if (row.length < 54) {
      errors.push('CSV行'+(idx+1)+': 54列未満 ('+row.length+'列)'); continue;
    }
    const key=lpbRaceKey_(course,date8,no);
    if (!grouped.has(key)) grouped.set(key, {course,date8,no,rawRowCount:0,items:{},errors:[]});
    const result=grouped.get(key);
    result.rawRowCount++;
    for (const spec of LPB.specs) {
      if (!result.items[spec.name]) result.items[spec.name]=[];
      for (const item of spec.items) {
        try {
          const entry=lpbExtractOne_(row,item,spec.name);
          if (entry) result.items[spec.name].push(entry);
        } catch(e) {
          result.errors.push('CSV行'+(idx+1)+' '+e.message);
        }
      }
    }
  }
  for (const result of grouped.values()) {
    for (const spec of LPB.specs) {
      const map=new Map();
      for (const item of result.items[spec.name]) {
        const key=item.numbers.join(spec.name==='枠単'||spec.name==='馬単'||spec.name==='3連単'?'>':'-');
        // 同一的中組番が別行にも現れる際、単純に加算しない
        if (map.has(key) && map.get(key)!==item.payout) {
          result.errors.push(spec.name+' '+key+' に異なる払戻額が複数あります');
        }
        map.set(key,item.payout);
      }
      result.items[spec.name]=[...map].map(([combination,payout])=>({combination,payout}));
    }
    result.totalHitEntries = LPB.specs.reduce((n,s)=>n+result.items[s.name].length,0);
    // 行が存在しても、明示的な確定フラグがないため「確定」と断定しない
    result.diagnosticStatus=result.errors.length ? '要確認（解析エラー）' :
      result.totalHitEntries===0 ? '払戻未掲載または特殊状態' : '払戻データあり（確定未判定）';
  }
  return {results:[...grouped.values()], errors, skippedNonToday};
}
