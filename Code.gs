/****************************************************
 * 在庫・棚卸・入出庫・発注 統合アプリ
 * Google Apps Script V8 / Google Sheets
 * 初期データは InitialData.gs から一度だけ取込
 ****************************************************/
const APP = {
  SHEETS: {
    products: '商品マスタ',
    stocktakes: '棚卸履歴',
    movements: '入出庫履歴',
    orders: '発注ヘッダ',
    orderLines: '発注明細',
    audit: '操作ログ',
    settings: '設定',
    legacyOutflow: '過去出庫実績',
    backlog: '受注残管理'
  },
  HEADERS: {
    products: ['商品ID','商品名','サイズ','仕入先','補充ライン','単価','有効','元シート','メモ','棚卸状態','通貨'],
    stocktakes: ['棚卸ID','商品ID','棚卸日','実棚数量','棚卸区分','元データ','登録日時','登録者'],
    movements: ['移動ID','商品ID','日付','区分','数量','単価','参照番号','備考','登録日時','登録者','通貨'],
    orders: ['発注ID','作成日','仕入先','発注日','納期','分割回','分割数','状態','合計金額','備考','更新日時','通貨'],
    orderLines: ['明細ID','発注ID','商品ID','商品名','サイズ','数量','単価','金額','分割元明細ID','通貨'],
    audit: ['日時','ユーザー','操作','対象ID','内容'],
    settings: ['キー','値'],
    legacyOutflow: ['実績ID','商品ID','数量','出庫日','元データ','備考','登録日時'],
    backlog: ['受注残ID','商品ID','商品名','サイズ','初期受注残','現在受注残','元シート','更新日時']
  }
};

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('在庫・発注管理')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1');
}

function setupDatabase() {
  const ss = SpreadsheetApp.getActive();
  Object.keys(APP.SHEETS).forEach(k => {
    const name = APP.SHEETS[k];
    let sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    if (sh.getLastRow() === 0) {
      sh.getRange(1,1,1,APP.HEADERS[k].length).setValues([APP.HEADERS[k]]);
      sh.setFrozenRows(1);
    }
  });
  ensureProductSchema_();
  applyInitialUncountedFlags_();
  const settings = ss.getSheetByName(APP.SHEETS.settings);
  if (settings.getLastRow() < 2) {
    settings.getRange(2,1,3,2).setValues([
      ['GEMINI_MODEL','gemini-3.8-flash'],
      ['会社名','竹山商事'],
      ['予測期間日数','90']
    ]);
  }
  return 'データベースを準備しました。次に importInitialData() を1回実行してください。';
}

function importInitialData() {
  if (typeof INITIAL_DATA === 'undefined') throw new Error('InitialData.gs がありません。ZIP内のファイルを同じApps Scriptプロジェクトに追加してください。');
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const ss = SpreadsheetApp.getActive();
    const psh = ss.getSheetByName(APP.SHEETS.products);
    if (psh.getLastRow() > 1) return '商品マスタにデータがあるため取込を中止しました（二重登録防止）。';
    const d = INITIAL_DATA;
    if (d.products.length) psh.getRange(2,1,d.products.length,11).setValues(d.products.map(p => [
      p.id,p.name,p.size,p.supplier||'',p.reorderPoint||0,p.unitPrice||0,p.active!==false,p.sourceSheets.join(' / '),'初期取込','',normalizeCurrency_(p.currency)
    ]));
    const tsh = ss.getSheetByName(APP.SHEETS.stocktakes);
    if (d.stocktakes.length) tsh.getRange(2,1,d.stocktakes.length,8).setValues(d.stocktakes.map((x,i)=>[
      'STK-INIT-'+(i+1),x.productId,x.date,x.qty,x.kind,x.source,new Date(),'初期データ'
    ]));
    const msh = ss.getSheetByName(APP.SHEETS.movements);
    if (d.movements.length) msh.getRange(2,1,d.movements.length,11).setValues(d.movements.map((x,i)=>[
      'MOV-INIT-'+(i+1),x.productId,x.date,'入庫',x.qty,0,'',x.note+' / '+x.source,new Date(),'初期データ','JPY'
    ]));
    [psh,tsh,msh].forEach(sh=>sh.autoResizeColumns(1,Math.min(sh.getLastColumn(),9)));
    log_('初期データ取込','ALL',`商品${d.products.length}件 / 棚卸${d.stocktakes.length}件 / 入庫${d.movements.length}件`);
    return `初期データ取込完了：商品${d.products.length}件、棚卸履歴${d.stocktakes.length}件、入庫履歴${d.movements.length}件。`;
  } finally { lock.releaseLock(); }
}

function getBootstrap() {
  ensureSetup_();
  return {
    products: getProducts_(),
    orders: getOrders_(),
    summary: getSummary_(),
    settings: getSettings_(),
    backlogs: getBacklogRows_()
  };
}

// 画面通信用：Date等を含む発注履歴を安全なJSON値にして返す
function getBootstrapForApp20261119() {
  ensureSetup_();
  return { ok:true, products:getProducts_(), orders:JSON.parse(JSON.stringify(getOrders_())), settings:JSON.parse(JSON.stringify(getSettings_())), backlogs:JSON.parse(JSON.stringify(getBacklogRows_())) };
}

function getProducts() { ensureSetup_(); return getProducts_(); }
function getOrderBacklogs(){ensureSetup_();return getBacklogRows_();}

// 過去から最新までの入庫履歴を商品マスタと結合して返す（履歴シートは変更しない）
function getInboundHistory(){
  ensureSetup_();
  const ss=SpreadsheetApp.getActive();
  const movements=readRows_(ss.getSheetByName(APP.SHEETS.movements));
  const productMap={};
  getProducts_().forEach(p=>productMap[String(p.id)]=p);
  return movements.filter(x=>String(x['区分']||'')==='入庫').map(x=>{
    const p=productMap[String(x['商品ID'])]||{};
    return {
      id:String(x['移動ID']||''), productId:String(x['商品ID']||''),
      date:dateKey_(x['日付']), name:String(p.name||'（商品マスタ未登録）'),
      size:String(p.size||''), supplier:String(p.supplier||''),
      qty:Number(x['数量'])||0, unitPrice:Number(x['単価'])||0,
      currency:normalizeCurrency_(x['通貨']), reference:String(x['参照番号']||''),
      note:String(x['備考']||''), registeredAt:x['登録日時'] instanceof Date?Utilities.formatDate(x['登録日時'],Session.getScriptTimeZone(),'yyyy-MM-dd HH:mm:ss'):String(x['登録日時']||'')
    };
  }).sort((a,b)=>b.date.localeCompare(a.date)||b.registeredAt.localeCompare(a.registeredAt));
}

function deriveStockOutflow_(stockRows,movementRows,fromDate,toDate,productRows){
  const history={};
  (stockRows||[]).forEach((x,i)=>{
    const id=String(x['商品ID']||''), d=dateKey_(x['棚卸日']);
    if(!id||!d)return;
    if(!history[id])history[id]=[];
    history[id].push({date:d,qty:Number(x['実棚数量'])||0,seq:i});
  });
  const hasRange=!!(fromDate&&toDate);
  const from=hasRange?dateKey_(fromDate):'';
  const to=hasRange?dateKey_(toDate):'';
  const distance=(a,b)=>Math.abs(new Date(a+'T00:00:00').getTime()-new Date(b+'T00:00:00').getTime());
  const ids=new Set(Object.keys(history));
  // レポートでは棚卸履歴がない商品も「計算不可」として表示する
  (productRows||[]).forEach(p=>ids.add(String(p['商品ID']||'')));
  const result={};
  ids.forEach(id=>{
    if(!id)return;
    const rows=(history[id]||[]).slice().sort((a,b)=>a.date.localeCompare(b.date)||a.seq-b.seq);
    if(rows.length<2){
      result[id]={status:'unavailable',reason:rows.length===0?'棚卸履歴なし':'棚卸履歴が1回のみ',baseDate:null,baseStock:null,latestDate:null,latestStock:null,inbound:null,derivedOut:null,days:null};
      return;
    }
    let base,latest;
    if(hasRange){
      // 開始日・終了日にそれぞれ最も近い棚卸を優先。
      // 2つが同日または逆転する場合は、開始日より前後の順序を守る最小距離の組を採用。
      const nearestBase=rows.slice().sort((a,b)=>distance(a.date,from)-distance(b.date,from)||a.date.localeCompare(b.date)||a.seq-b.seq)[0];
      const nearestEnd=rows.slice().sort((a,b)=>distance(a.date,to)-distance(b.date,to)||b.date.localeCompare(a.date)||b.seq-a.seq)[0];
      if(nearestBase.date<nearestEnd.date){base=nearestBase;latest=nearestEnd;}
      else {
        const pairs=[];
        rows.forEach(a=>rows.forEach(b=>{
          if(a.date<b.date) pairs.push({a,b,score:distance(a.date,from)+distance(b.date,to)});
        }));
        pairs.sort((a,b)=>a.score-b.score||a.a.date.localeCompare(b.a.date)||b.b.date.localeCompare(a.b.date)||a.a.seq-b.a.seq||b.b.seq-a.b.seq);
        if(pairs.length){base=pairs[0].a;latest=pairs[0].b;}
      }
    }else{base=rows[0];latest=rows[rows.length-1];}
    if(!base||!latest||base.date>=latest.date){
      result[id]={status:'unavailable',reason:'異なる日付の棚卸が2回ありません',baseDate:null,baseStock:null,latestDate:null,latestStock:null,inbound:null,derivedOut:null,days:null};
      return;
    }
    const inbound=(movementRows||[])
      .filter(x=>String(x['商品ID']||'')===id&&x['区分']==='入庫')
      .filter(x=>{const d=dateKey_(x['日付']);return d>base.date&&d<=latest.date;})
      .reduce((a,x)=>a+(Number(x['数量'])||0),0);
    const days=Math.max(1,Math.round((new Date(latest.date+'T00:00:00')-new Date(base.date+'T00:00:00'))/86400000));
    result[id]={status:'calculated',reason:'',baseDate:base.date,baseStock:base.qty,latestDate:latest.date,latestStock:latest.qty,inbound,derivedOut:base.qty+inbound-latest.qty,days};
  });
  return result;
}

function getProducts_() {
  const ss=SpreadsheetApp.getActive();
  const ps=readRows_(ss.getSheetByName(APP.SHEETS.products));
  const st=readRows_(ss.getSheetByName(APP.SHEETS.stocktakes));
  const mv=readRows_(ss.getSheetByName(APP.SHEETS.movements));
  const stockMap={};
  st.forEach(x=>{
    const id=String(x['商品ID']);
    const d=dateKey_(x['棚卸日']);
    // 同じ日付に再棚卸した場合は、後から登録した数量を優先
    if (!stockMap[id] || d>=stockMap[id].date) stockMap[id]={date:d,qty:Number(x['実棚数量'])||0,kind:x['棚卸区分']||''};
  });
  const moveMap={};
  const derivedMap=deriveStockOutflow_(st,mv);
  mv.forEach(x=>{
    const id=String(x['商品ID']), d=dateKey_(x['日付']);
    if (!moveMap[id]) moveMap[id]=[];
    moveMap[id].push({date:d,type:x['区分'],qty:Number(x['数量'])||0});
  });
  return ps.filter(x=>String(x['有効']).toLowerCase()!=='false').map(x=>{
    const id=String(x['商品ID']), last=stockMap[id];
    const forcedUnknown=String(x['棚卸状態']||'')==='未棚卸';
    let balance=null;
    if (last && !forcedUnknown) {
      // 現在庫は最新棚卸の実棚数量をそのまま採用する。
      // 棚卸日以降の入出庫履歴は記録・分析用として保持し、現在庫には加算しない。
      balance=last.qty;
    }
    const usage=(moveMap[id]||[]).filter(m=>m.type==='出庫');
    const recent=usage.filter(m=>m.date>=shiftDate_(today_(),-90));
    const explicitDaily=recent.reduce((s,m)=>s+m.qty,0)/90;
    const derivedInfo=derivedMap[id]||{};
    const derivedDaily=Number(derivedInfo.derivedOut||0)/Math.max(1,Number(derivedInfo.days)||1);
    const daily=Math.max(explicitDaily,derivedDaily);
    const reorder=Number(x['補充ライン'])||0;
    const eta=daily>0 && balance!==null ? Math.max(0,Math.ceil((balance-reorder)/daily)) : null;
    const stockout=daily>0 && balance!==null ? Math.max(0,Math.ceil(balance/daily)) : null;
    return {
      id,name:x['商品名'],size:x['サイズ'],supplier:x['仕入先']||'',
      reorderPoint:reorder,unitPrice:Number(x['単価'])||0,currency:normalizeCurrency_(x['通貨']),
      balance,lastStocktakeDate:(last&&!forcedUnknown)?last.date:'',stocktakeKind:(last&&!forcedUnknown)?last.kind:'',
      stocktakeStatus:balance===null?'未棚卸':'棚卸済',
      status:balance===null?'未棚卸':(balance<=reorder?'要補充':'在庫あり'),
      dailyOut:daily,derivedOutQty:Number(derivedInfo.derivedOut)||0,derivedBaseStock:derivedInfo.baseStock??null,derivedLatestStock:derivedInfo.latestStock??null,derivedInbound:Number(derivedInfo.inbound)||0,derivedPeriodDays:Number(derivedInfo.days)||0,daysToReorder:eta,daysToStockout:stockout,
      sourceSheets:x['元シート']||''
    };
  });
}

function saveMovement(data) {
  validate_(data && data.productId,'商品を選択してください');
  const qty=Number(data.qty);
  if (!Number.isFinite(qty)||qty<=0) throw new Error('数量は1以上で入力してください');
  if (!['入庫','出庫'].includes(data.type)) throw new Error('区分が不正です');
  const date=dateKey_(data.date||today_());
  const lock=LockService.getScriptLock(); lock.waitLock(30000);
  try {
    const sh=SpreadsheetApp.getActive().getSheetByName(APP.SHEETS.movements);
    const id=makeId_('MOV');
    const productRow=readRows_(SpreadsheetApp.getActive().getSheetByName(APP.SHEETS.products)).find(x=>String(x['商品ID'])===String(data.productId))||{};
    const price=data.unitPrice!=='' && data.unitPrice!==null && data.unitPrice!==undefined ? Number(data.unitPrice) : Number(productRow['単価'])||0;
    if(!Number.isFinite(price)||price<0) throw new Error('単価は0以上の数値で入力してください');
    const currency=normalizeCurrency_(data.currency||productRow['通貨']);
    sh.appendRow([id,data.productId,date,data.type,qty,price,data.reference||'',data.note||'',new Date(),user_(),currency]);
    if(data.type==='入庫') adjustBacklog_(data.productId,-qty,productRow['商品名']||'',productRow['サイズ']||'');
    log_('入出庫登録',data.productId,`${data.type} ${qty} / ${date}`);
    return {id,products:getProducts_()};
  } finally {lock.releaseLock();}
}

function saveStocktake(data) {
  if (!data || !Array.isArray(data.items) || !data.items.length) throw new Error('棚卸対象がありません');
  const date=dateKey_(data.date||today_());
  const inputMap={};
  data.items.forEach(x=>{
    if (!x.productId || x.qty==='' || x.qty===null || x.qty===undefined) return;
    const q=Number(x.qty);
    if (!Number.isFinite(q)||q<0) throw new Error('棚卸数量は0以上で入力してください');
    inputMap[String(x.productId)]=q;
  });
  const ids=Object.keys(inputMap);
  if (!ids.length) throw new Error('数量が入力された商品がありません');
  const lock=LockService.getScriptLock(); lock.waitLock(30000);
  try {
    const ss=SpreadsheetApp.getActive(), sh=ss.getSheetByName(APP.SHEETS.stocktakes);
    const last=sh.getLastRow(), values=last>1?sh.getRange(2,1,last-1,8).getValues():[];
    const idSet=new Set(ids), existing={};
    // 同一商品・同一日付の行は更新し、別日付の棚卸履歴は残す
    values.forEach((r,i)=>{
      const key=String(r[1])+'|'+dateKey_(r[2]);
      if(idSet.has(String(r[1])) && dateKey_(r[2])===date) existing[String(r[1])]=i;
    });
    const now=new Date(), user=user_();
    ids.forEach(pid=>{
      const row=[makeId_('STK'),pid,date,inputMap[pid],data.kind||'通常棚卸',data.note||'',now,user];
      if(existing[pid]!==undefined) values[existing[pid]]=row;
      else values.push(row);
    });
    if(values.length){
      sh.getRange(2,1,Math.max(last-1,values.length),8).clearContent();
      sh.getRange(2,1,values.length,8).setValues(values);
    }
    const psh=ss.getSheetByName(APP.SHEETS.products), pvals=psh.getDataRange().getValues(), pmap={};
    pvals.slice(1).forEach((r,i)=>pmap[String(r[0])]=i+2);
    ids.forEach(pid=>{const pr=pmap[pid];if(pr)psh.getRange(pr,10).setValue('棚卸済');});
    log_('棚卸一括確定','MULTI',`${date} / ${ids.length}商品（同日分は置換）`);
    return {count:ids.length,products:getProducts_()};
  } finally {lock.releaseLock();}
}

function getReports(from,to) {
  ensureSetup_();
  from=dateKey_(from||shiftDate_(today_(),-365)); to=dateKey_(to||today_());
  if(from>to) throw new Error('開始日は終了日以前の日付にしてください');
  const ss=SpreadsheetApp.getActive();
  const products=readRows_(ss.getSheetByName(APP.SHEETS.products));
  const derivedMap=deriveStockOutflow_(readRows_(ss.getSheetByName(APP.SHEETS.stocktakes)),readRows_(ss.getSheetByName(APP.SHEETS.movements)),from,to,products);
  const names={}; products.forEach(p=>names[String(p['商品ID'])]=`${p['商品名']} ${p['サイズ']}`.trim());
  const mv=readRows_(ss.getSheetByName(APP.SHEETS.movements)).filter(x=>{
    const d=dateKey_(x['日付']); return d>=from && d<=to;
  });
  const monthly={}, byProduct={}, daily={}, dailyDetails=[];
  mv.forEach(x=>{
    const d=dateKey_(x['日付']), month=d.slice(0,7), id=String(x['商品ID']);
    if(!monthly[month])monthly[month]={month,in:0,out:0,amountIn:0,amountOut:0};
    if(!daily[d])daily[d]={date:d,in:0,out:0};
    const q=Number(x['数量'])||0, price=Number(x['単価'])||0;
    if(x['区分']==='入庫'){monthly[month].in+=q;monthly[month].amountIn+=q*price;daily[d].in+=q;}
    else {
      monthly[month].out+=q;monthly[month].amountOut+=q*price;daily[d].out+=q;
      dailyDetails.push({date:d,productId:id,name:names[id]||id,qty:q,unitPrice:price,amount:q*price,note:String(x['備考']||'')});
    }
    if(!byProduct[id])byProduct[id]={productId:id,name:names[id]||id,in:0,out:0};
    if(x['区分']==='入庫')byProduct[id].in+=q;else byProduct[id].out+=q;
  });
  // 日付が特定できない元台帳の累計出庫。期間内の日別/月別へは混ぜず別枠で返す。
  const legacyRows=readRows_(ensureLegacyOutflowSheet_()).map(x=>{
    const id=String(x['商品ID']||'');
    return {productId:id,name:names[id]||id,qty:Number(x['数量'])||0,source:String(x['元データ']||''),note:String(x['備考']||'')};
  }).filter(x=>x.qty>0);
  const legacyOutTotal=legacyRows.reduce((a,x)=>a+x.qty,0);
  const derivedOutflow=Object.keys(derivedMap).map(id=>({productId:id,name:names[id]||id,status:derivedMap[id].status,reason:derivedMap[id].reason,baseDate:derivedMap[id].baseDate,baseStock:derivedMap[id].baseStock,latestDate:derivedMap[id].latestDate,latestStock:derivedMap[id].latestStock,inbound:derivedMap[id].inbound,qty:derivedMap[id].derivedOut,days:derivedMap[id].days})); 
  const stocktakeRows=readRows_(ss.getSheetByName(APP.SHEETS.stocktakes)).filter(x=>{
    const d=dateKey_(x['棚卸日']); return d>=from && d<=to;
  });
  const stocktakeByDate={};
  stocktakeRows.forEach(x=>{
    const d=dateKey_(x['棚卸日']);
    if(!stocktakeByDate[d])stocktakeByDate[d]={date:d,items:0,qty:0};
    stocktakeByDate[d].items++;stocktakeByDate[d].qty+=Number(x['実棚数量'])||0;
  });
  const dailyRows=Object.values(daily).sort((a,b)=>a.date.localeCompare(b.date));
  const dailyOutTotal=dailyRows.reduce((a,x)=>a+x.out,0);
  const dailyDetailsSorted=dailyDetails.sort((a,b)=>a.date.localeCompare(b.date)||a.name.localeCompare(b.name,'ja'));
  return {from,to,monthly:Object.values(monthly).sort((a,b)=>a.month.localeCompare(b.month)),
    byProduct:Object.values(byProduct).sort((a,b)=>b.out-a.out),daily:dailyRows,dailyDetails:dailyDetailsSorted,
    dailyOutTotal,derivedOutflow,derivedOutTotal:derivedOutflow.filter(x=>x.status==='calculated').reduce((a,x)=>a+(Number(x.qty)||0),0),derivedOutCalculatedCount:derivedOutflow.filter(x=>x.status==='calculated').length,derivedOutUnavailableCount:derivedOutflow.filter(x=>x.status!=='calculated').length,legacyOut:legacyRows,legacyOutTotal,stocktakes:Object.values(stocktakeByDate).sort((a,b)=>a.date.localeCompare(b.date))};
}

function getSummary_() {
  const p=getProducts_();
  return {
    total:p.length,
    known:p.filter(x=>x.balance!==null).length,
    unknown:p.filter(x=>x.balance===null).length,
    low:p.filter(x=>x.balance!==null&&x.balance<=x.reorderPoint).length,
    totalStock:p.filter(x=>x.balance!==null).reduce((s,x)=>s+x.balance,0)
  };
}

function getOrders() { ensureSetup_(); return getOrders_(); }
function getOrders_() {
  const ss=SpreadsheetApp.getActive();
  const hs=readRows_(ss.getSheetByName(APP.SHEETS.orders));
  const ls=readRows_(ss.getSheetByName(APP.SHEETS.orderLines));
  return hs.map(h=>({...h,lines:ls.filter(l=>String(l['発注ID'])===String(h['発注ID']))}));
}

function saveOrder(data) {
  if(!data || !data.supplier) throw new Error('仕入先を入力してください');
  if(!Array.isArray(data.lines)||!data.lines.length) throw new Error('発注明細がありません');
  const id=makeId_('ORD'), date=today_();
  const lines=data.lines.filter(x=>Number(x.qty)>0).map(x=>{
    const qty=Math.floor(Number(x.qty)), price=Number(x.unitPrice)||0;
    return {productId:x.productId||'',name:x.name||'',size:x.size||'',qty,unitPrice:price,currency:normalizeCurrency_(x.currency),amount:qty*price,sourceLineId:''};
  });
  if(!lines.length) throw new Error('数量が1以上の明細がありません');
  const total=lines.reduce((s,x)=>s+x.amount,0);
  const ss=SpreadsheetApp.getActive();
  ss.getSheetByName(APP.SHEETS.orders).appendRow([id,date,data.supplier,data.orderDate||date,data.deliveryDate||'',1,1,'下書き',total,data.note||'',new Date(),orderCurrency_(lines)]);
  const sh=ss.getSheetByName(APP.SHEETS.orderLines);
  sh.getRange(sh.getLastRow()+1,1,lines.length,10).setValues(lines.map((x,i)=>[
    id+'-'+(i+1),id,x.productId,x.name,x.size,x.qty,x.unitPrice,x.amount,'',x.currency
  ]));
  lines.forEach(x=>adjustBacklog_(x.productId,x.qty,x.name,x.size));
  log_('発注作成',id,`${data.supplier} / ${total}`);
  return getOrders_();
}

/**
 * 金額均等分割。明細の数量は整数のまま、単価の高い単位から
 * 現在金額が最も少ない回に1個ずつ割り当てる決定的な貪欲法。
 */
function createSplitPlan(data) {
  const n=Math.max(2,Math.min(12,Math.floor(Number(data.count)||2)));
  if(!Array.isArray(data.lines)||!data.lines.length) throw new Error('分割対象明細がありません');
  const batches=Array.from({length:n},(_,i)=>({index:i+1,total:0,lines:[]}));
  const currencies=[...new Set(data.lines.filter(l=>Number(l.qty)>0).map(l=>normalizeCurrency_(l.currency)))];
  if(currencies.length>1) throw new Error('異なる通貨を混在した分割発注はできません。通貨ごとに発注を分けてください。');
  const units=[];
  data.lines.forEach((l,idx)=>{
    const qty=Math.max(0,Math.floor(Number(l.qty)||0)), price=Math.max(0,Number(l.unitPrice)||0);
    for(let k=0;k<qty;k++) units.push({lineIndex:idx,name:l.name,size:l.size,productId:l.productId||'',unitPrice:price,currency:normalizeCurrency_(l.currency)});
  });
  units.sort((a,b)=>b.unitPrice-a.unitPrice);
  units.forEach(u=>{
    batches.sort((a,b)=>a.total-b.total || a.index-b.index);
    const b=batches[0];
    let line=b.lines.find(x=>x.lineIndex===u.lineIndex);
    if(!line){line={lineIndex:u.lineIndex,productId:u.productId,name:u.name,size:u.size,qty:0,unitPrice:u.unitPrice,currency:u.currency,amount:0};b.lines.push(line);}
    line.qty++;line.amount+=u.unitPrice;b.total+=u.unitPrice;
  });
  batches.sort((a,b)=>a.index-b.index);
  const original=data.lines.reduce((s,l)=>s+Math.max(0,Math.floor(Number(l.qty)||0))*(Number(l.unitPrice)||0),0);
  const sum=batches.reduce((s,b)=>s+b.total,0);
  return {count:n,currency:currencies[0]||'JPY',originalTotal:original,distributedTotal:sum,difference:original-sum,batches,average:sum/n,spread:Math.max(...batches.map(x=>x.total))-Math.min(...batches.map(x=>x.total))};
}

function saveSplitPlan(data) {
  const plan=data.plan;
  if(!plan||!plan.batches||!plan.batches.length) throw new Error('分割案がありません');
  const ss=SpreadsheetApp.getActive(), hsh=ss.getSheetByName(APP.SHEETS.orders), lsh=ss.getSheetByName(APP.SHEETS.orderLines);
  const baseId=data.orderId||makeId_('ORD');
  const supplier=data.supplier||'';
  const allRows=[];
  plan.batches.forEach((b,i)=>{
    const id=baseId+'-P'+(i+1), total=b.lines.reduce((s,l)=>s+l.amount,0);
    hsh.appendRow([id,today_(),supplier,data.orderDates?.[i]||today_(),data.deliveryDates?.[i]||'',i+1,plan.count,'下書き',total,data.note||'',new Date(),orderCurrency_(b.lines)]);
    b.lines.forEach((l,j)=>allRows.push([id+'-'+(j+1),id,l.productId,l.name,l.size,l.qty,l.unitPrice,l.amount,'',normalizeCurrency_(l.currency)]));
  });
  if(allRows.length) lsh.getRange(lsh.getLastRow()+1,1,allRows.length,10).setValues(allRows);
  log_('分割発注保存',baseId,`${plan.count}回 / 合計${plan.distributedTotal}`);
  return getOrders_();
}

function updateOrderStatus(id,status) {
  const allowed=['下書き','発注済','一部入荷','完了','取消'];
  if(!allowed.includes(status)) throw new Error('状態が不正です');
  const sh=SpreadsheetApp.getActive().getSheetByName(APP.SHEETS.orders);
  const values=sh.getDataRange().getValues();
  for(let r=1;r<values.length;r++) if(String(values[r][0])===String(id)){
    sh.getRange(r+1,8).setValue(status);sh.getRange(r+1,11).setValue(new Date());
    log_('発注状態変更',id,status);return getOrders_();
  }
  throw new Error('発注が見つかりません');
}


/** 発注候補：直近90日出庫ペースから指定日数分を予測し、売れ筋には控えめな安全在庫を加える */
function getOrderSuggestions(supplier, horizonDays) {
  ensureSetup_();
  const days=Math.max(1,Math.min(365,Math.floor(Number(horizonDays)||30)));
  const ss=SpreadsheetApp.getActive();
  const products=getProducts_().filter(p=>!supplier || supplier==='__UNASSIGNED__'?!p.supplier:p.supplier===supplier);
  const legacy=readRows_(ensureLegacyOutflowSheet_());
  const legacyMap={};
  legacy.forEach(r=>{const id=String(r['商品ID']||'');legacyMap[id]=(legacyMap[id]||0)+(Number(r['数量'])||0);});
  const ranked=products.map(p=>({...p,legacyOutQty:legacyMap[p.id]||0}));
  const activeRates=ranked.map(p=>p.dailyOut||0).filter(x=>x>0).sort((a,b)=>a-b);
  const q80=activeRates.length?activeRates[Math.floor((activeRates.length-1)*0.8)]:Infinity;
  const q50=activeRates.length?activeRates[Math.floor((activeRates.length-1)*0.5)]:Infinity;
  return ranked.map(p=>{
    const rate=Number(p.dailyOut)||0;
    // 日付不明の元台帳は日割りせず、売れ筋判定の補助情報に限定する
    const fast=rate>0 && rate>=q80;
    const regular=rate>0 && !fast && rate>=q50;
    const uplift=fast?0.15:(regular?0.08:0);
    // 日付付き出庫が少ない商品は、元台帳の過去出庫と補充ラインも発注候補計算に利用する。
    // これにより、履歴が旧台帳側にしかない商品が一律0になるのを防ぐ。
    const demandBase=Math.max(rate*days, Number(p.legacyOutQty)||0, Number(p.reorderPoint)||0);
    const forecast=Math.ceil(demandBase*(1+uplift));
    const stock=p.balance===null?null:Number(p.balance);
    const suggested=stock===null?Math.max(0,forecast):Math.max(0,forecast-stock,Number(p.reorderPoint)||0-stock);
    return {id:p.id,name:p.name,size:p.size,supplier:p.supplier,balance:stock,stocktakeStatus:p.stocktakeStatus,
      dailyOut90:rate,legacyOutQty:p.legacyOutQty,forecastQty:forecast,upliftRate:uplift,suggestedQty:suggested,
      unitPrice:p.unitPrice,currency:p.currency,reorderPoint:p.reorderPoint,derivedOutQty:p.derivedOutQty||0,derivedPeriodDays:p.derivedPeriodDays||0,days};
  }).sort((a,b)=>b.suggestedQty-a.suggestedQty || b.dailyOut90-a.dailyOut90);
}

function getSettings_() {
  const sh=SpreadsheetApp.getActive().getSheetByName(APP.SHEETS.settings);
  const o={}; if(sh.getLastRow()>1) sh.getRange(2,1,sh.getLastRow()-1,2).getValues().forEach(r=>o[r[0]]=r[1]);
  o.hasGeminiKey=!!PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  return o;
}
function setGeminiApiKey(key) {
  if(!key || String(key).trim().length<20) throw new Error('APIキーを確認してください');
  PropertiesService.getScriptProperties().setProperty('GEMINI_API_KEY',String(key).trim());
  log_('Gemini APIキー設定','SETTINGS','キーをScript Propertiesに保存');
  return getSettings_();
}
function aiForecast(from,to) {
  const key=PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if(!key) throw new Error('設定画面でGemini APIキーを登録してください。無料枠の有無・上限はGoogle AI Studio側で確認できます。');
  const settings=getSettings_();
  const model='gemini-3.8-flash';
  const start=dateKey_(from||shiftDate_(today_(),-365)), end=dateKey_(to||today_());
  const products=getProducts_().filter(p=>p.balance!==null);
  const report=getReports(start,end);
  const candidates=products.map(p=>({id:p.id,name:p.name,size:p.size,supplier:p.supplier,balance:p.balance,
    reorderPoint:p.reorderPoint,dailyOut90:p.dailyOut,daysToReorder:p.daysToReorder,daysToStockout:p.daysToStockout,
    unitPrice:p.unitPrice})).filter(p=>p.dailyOut90>0||p.balance<=p.reorderPoint)
    .sort((a,b)=>(a.daysToStockout==null?99999:a.daysToStockout)-(b.daysToStockout==null?99999:b.daysToStockout)).slice(0,100);
  const prompt=[
    'あなたは剣道用品店の在庫管理を支援する需要予測AIです。日本語で、実務向けに簡潔に回答してください。',
    '以下のデータに基づき、①出庫傾向 ②欠品リスクが高い商品 ③補充を検討する商品 ④データ不足・注意点 ⑤次に行うべきことを整理してください。',
    '重要：在庫数・出庫数・金額を推測で作らない。未棚卸商品は在庫0と扱わない。商品ごとに最新棚卸日が違うため、在庫数はシステム計算値として扱う。',
    'dailyOut90は直近90日間のアプリ登録出庫数÷90。過去出庫参考値と補充ラインも候補判断に使う。出庫データがない場合はデータ不足と明記し、数値を推測しない。',
    '分析対象期間：'+start+' ～ '+end,
    '期間内の月別入出庫：'+JSON.stringify(report.monthly),
    '期間内の商品別出庫（アプリ登録・日付確定分）：'+JSON.stringify(report.byProduct.slice(0,100)),
    '棚卸差異から推定した出庫（基準在庫＋期間内入庫−最新棚卸在庫）：合計'+(report.derivedOutTotal||0)+'点 / 明細：'+JSON.stringify(report.derivedOutflow||[]),
    '元台帳の過去出庫実績（日付不明・期間集計には混ぜない）：合計'+report.legacyOutTotal+'点 / 明細：'+JSON.stringify(report.legacyOut),
    '重要：日付不明の過去出庫は日別・月別・90日平均に加算しない。商品別の過去需要参考情報として別枠で評価する。',
    '現在庫・補充リスク候補：'+JSON.stringify(candidates)
  ].join('\n');
  const url='https://generativelanguage.googleapis.com/v1beta/models/'+encodeURIComponent(model)+':generateContent';
  const res=UrlFetchApp.fetch(url,{method:'post',contentType:'application/json',headers:{'x-goog-api-key':key},payload:JSON.stringify({contents:[{role:'user',parts:[{text:prompt}]}],generationConfig:{maxOutputTokens:3000}}),muteHttpExceptions:true});
  let body={};try{body=JSON.parse(res.getContentText());}catch(e){throw new Error('GeminiからJSON以外の応答が返りました：'+res.getContentText().slice(0,300));}
  if(res.getResponseCode()>=300)throw new Error((body.error&&body.error.message)||('Gemini APIエラー（HTTP '+res.getResponseCode()+'）'));
  const text=(body.candidates&&body.candidates[0]&&body.candidates[0].content&&body.candidates[0].content.parts||[]).map(x=>x.text||'').join('\n');
  if(!text)throw new Error('AIから分析文が返りませんでした。APIキー・利用上限を確認してください。');
  log_('AI需要分析','ALL',model+' / '+start+'～'+end);
  return {text,model,from:start,to:end,candidateCount:candidates.length};
}

/** 発注書原本に基づく商品名・バリエーション・仕入先・単価の修正。履歴は変更しない。 */
function repairProductStructureAndPrices20261119() {
  ensureSetup_();
  const sh=SpreadsheetApp.getActive().getSheetByName(APP.SHEETS.products);
  const vals=sh.getDataRange().getValues(), headers=vals[0], col={};
  headers.forEach((h,i)=>col[String(h)]=i+1);
  const idCol=col['商品ID'], nameCol=col['商品名'], sizeCol=col['サイズ'], supplierCol=col['仕入先'], priceCol=col['単価'], activeCol=col['有効'], sourceCol=col['元シート'];
  const rowById={}; vals.slice(1).forEach((r,i)=>rowById[String(r[idCol-1])]=i+2);
  const updates=[];
  function setRow(id,name,size,supplier,price){
    const row=rowById[id]; if(!row)throw new Error('対象商品IDが見つかりません：'+id);
    sh.getRange(row,nameCol).setValue(name);sh.getRange(row,sizeCol).setValue(size);
    sh.getRange(row,supplierCol).setValue(supplier);sh.getRange(row,priceCol).setValue(price);
    if(activeCol)sh.getRange(row,activeCol).setValue(true);
    updates.push(id);
  }
  // 既存の在庫・棚卸履歴が紐づくP00046/P00047は、面ピットM/LとしてIDを維持
  setRow('P00046','面ピット','M','三信物産',1320);
  setRow('P00047','面ピット','L','三信物産',1320);
  // 学販小手は独立商品。生子あり/なしを新規バリエーションとして追加
  const newItems=[['学販小手','生子あり','三信物産',2740],['学販小手','生子なし','三信物産',2600]];
  let next=1; vals.slice(1).forEach(r=>{const m=String(r[idCol-1]).match(/^P(\d+)$/);if(m)next=Math.max(next,Number(m[1])+1);});
  const existingNames=vals.slice(1).map(r=>String(r[nameCol-1])+'|'+String(r[sizeCol-1]));
  const addRows=[];
  newItems.forEach(x=>{
    if(existingNames.includes(x[0]+'|'+x[1]))return;
    const row=Array(headers.length).fill('');
    row[idCol-1]='P'+String(next++).padStart(5,'0');row[nameCol-1]=x[0];row[sizeCol-1]=x[1];
    row[supplierCol-1]=x[2];row[col['補充ライン']-1]=0;row[priceCol-1]=x[3];row[activeCol-1]=true;
    if(sourceCol)row[sourceCol-1]='三信物産 発注書原本';
    if(col['メモ'])row[col['メモ']-1]='発注書原本より登録';
    addRows.push(row);
  });
  if(addRows.length)sh.getRange(sh.getLastRow()+1,1,addRows.length,headers.length).setValues(addRows);
  // 三栄武道具の胸3種類。2025-12-19原本の4,605円（2025-06-10原本4,515円より高い方）
  const chestIds=['P00307','P00308','P00309','P00310','P00311','P00312','P00313','P00314','P00315','P00316','P00317','P00318'];
  chestIds.forEach(id=>{
    const row=rowById[id];if(!row)return;
    sh.getRange(row,supplierCol).setValue('三栄武道具');sh.getRange(row,priceCol).setValue(4605);
  });
  // ガラムスポーツのジャージ関連商品はメーカーを明示
  const allRows=sh.getDataRange().getValues();let garam=0;
  allRows.slice(1).forEach((r,i)=>{
    const id=String(r[idCol-1]);
    if(/^P00319$/.test(id)||(/^P003[2-4]\d$/.test(id)&&Number(id.slice(1))<=349)){
      sh.getRange(i+2,supplierCol).setValue('ガラムスポーツ');garam++;
    }
  });
  log_('商品構成・単価修正','P00046-P00558',`商品名・バリエーション修正${updates.length}件 / 学販小手追加${addRows.length}件 / 胸単価4605円 / ガラム仕入先${garam}件`);
  return `修正完了：面ピットM/L、学販小手（生子あり/なし）を別商品化。胸12バリエーションは三栄武道具・4,605円（高い方）。ガラムスポーツ${garam}商品をメーカー登録。追加商品${addRows.length}件。棚卸・入出庫履歴は削除・変更していません。`;
}

function setProductSettings(data) {
  const ss=SpreadsheetApp.getActive();
  const sh=ss.getSheetByName(APP.SHEETS.products);
  const vals=sh.getDataRange().getValues();
  const map={}; vals.slice(1).forEach((r,i)=>map[String(r[0])]=i+2);
  let updated=0, priceChanges=0;
  (data||[]).forEach(p=>{
    const row=map[String(p.id)]; if(!row)return;
    const old=sh.getRange(row,6).getValue();
    const headers=vals[0], currencyCol=headers.indexOf('通貨')+1;
    const next=Number(p.unitPrice);
    if(!Number.isFinite(next)||next<0) throw new Error('単価は0以上の数値で入力してください：'+(p.name||p.id));
    const supplier=p.supplier||'';
    const reorder=Number(p.reorderPoint)||0;
    sh.getRange(row,4,1,3).setValues([[supplier,reorder,next]]);
    if(currencyCol) sh.getRange(row,currencyCol).setValue(normalizeCurrency_(p.currency));
    updated++;
    if(Number(old||0)!==next) {
      priceChanges++;
      const productName=String(sh.getRange(row,2).getValue()||p.id);
      log_('単価変更',String(p.id),`${productName}：${formatMoney_(Number(old)||0, normalizeCurrency_(p.currency))} → ${formatMoney_(next, normalizeCurrency_(p.currency))}（${today_()}適用）`);
    }
  });
  log_('商品設定更新','MULTI',`${updated}件 / 単価変更${priceChanges}件`);
  return getProducts_();
}


/**
 * 印伝風鍔止の23種類と「菖蒲１本入り（良品/不良）」の商品分類を修正。
 * 既存の棚卸履歴・入出庫履歴・単価・補充ラインは保持します。
 * 既存DBでこの関数を1回だけ手動実行してください。
 */
function repairIndenTsubadomeProducts() {
  ensureSetup_();
  const sh=SpreadsheetApp.getActive().getSheetByName(APP.SHEETS.products);
  const values=sh.getDataRange().getValues();
  const map={};
  values.slice(1).forEach((r,i)=>map[String(r[0])]=i+2);
  const variants=[
    '1.黒×金蜻蛉','2.黒×銀蜻蛉','3.黒×黒蜻蛉','4.黒×青蜻蛉','5.黒×黄蜻蛉',
    '6.黒×赤蜻蛉','7.黒×赤青蜻蛉','8.黒×黒赤蜻蛉','9.赤×白蜻蛉','10.黒×金小桜',
    '11.黒×銀小桜','12.黒×黒小桜','13.黒×青小桜','14.黒×紫小桜','15.黒×赤小桜',
    '16.黒×黒赤小桜','17.赤×白小桜','18.黒×青瓢箪','19.黒×赤瓢箪','20.黒×黒瓢箪',
    '21.黒×金波','22.黒×銀波','23.黒×黒波'
  ];
  variants.forEach((size,i)=>{
    const id='P'+String(461+i).padStart(5,'0');
    const row=map[id];
    if(row){ sh.getRange(row,2,1,2).setValues([['印伝風 鍔止',size]]); sh.getRange(row,7).setValue(true); }
    else {
      const newRow=sh.getLastRow()+1;
      sh.getRange(newRow,1,1,10).setValues([[id,'印伝風 鍔止',size,'',0,0,true,'2024 衣類 / 在庫衣類 / 棚卸 海外発注用','分類修正で追加','']]);
      map[id]=newRow;
    }
  });
  // P00554 was only a section heading, not a countable stock variant.
  if(map.P00554){ sh.getRange(map.P00554,2,1,2).setValues([['菖蒲１本入り','']]); sh.getRange(map.P00554,7).setValue(false); }
  // Preserve existing product IDs so the 273 / 376 stocktake history remains linked.
  if(map.P00555){ sh.getRange(map.P00555,2,1,2).setValues([['菖蒲１本入り','良品']]); sh.getRange(map.P00555,7).setValue(true); }
  if(map.P00556){ sh.getRange(map.P00556,2,1,2).setValues([['菖蒲１本入り','不良']]); sh.getRange(map.P00556,7).setValue(true); }
  log_('商品分類修正','P00461-P00556','印伝風鍔止23種類 / 菖蒲１本入り 良品・不良');
  return '修正完了：印伝風 鍔止 23種類、菖蒲１本入り（良品・不良）。既存の棚卸・入出庫履歴は保持しました。';
}

/**
 * 元Excel「作業用 衣類」の印伝風 鍔止 D列「出庫数」を日付不明実績として登録。
 * 日付・月は推定せず、アプリ入出庫履歴や棚卸履歴には追加しない。
 * 同じ実績IDがあれば再登録しないため、既存データを壊さず再実行可能。
 */
function importIndenLegacyOutflow20261119() {
  ensureSetup_();
  const sh=ensureLegacyOutflowSheet_();
  const rows=[
    ['LEGACY-INDEN-01','P00461',75,'','在庫2024~.xlsx / 作業用 衣類 / D列 出庫数','2025/10/20在庫欄に対応する元台帳。出庫日は特定不可'],
    ['LEGACY-INDEN-02','P00462',88,'','在庫2024~.xlsx / 作業用 衣類 / D列 出庫数','2025/10/20在庫欄に対応する元台帳。出庫日は特定不可'],
    ['LEGACY-INDEN-03','P00463',113,'','在庫2024~.xlsx / 作業用 衣類 / D列 出庫数','2025/10/20在庫欄に対応する元台帳。出庫日は特定不可'],
    ['LEGACY-INDEN-04','P00464',72,'','在庫2024~.xlsx / 作業用 衣類 / D列 出庫数','2025/10/20在庫欄に対応する元台帳。出庫日は特定不可'],
    ['LEGACY-INDEN-05','P00465',80,'','在庫2024~.xlsx / 作業用 衣類 / D列 出庫数','2025/10/20在庫欄に対応する元台帳。出庫日は特定不可'],
    ['LEGACY-INDEN-06','P00466',70,'','在庫2024~.xlsx / 作業用 衣類 / D列 出庫数','2025/10/20在庫欄に対応する元台帳。出庫日は特定不可'],
    ['LEGACY-INDEN-07','P00467',184,'','在庫2024~.xlsx / 作業用 衣類 / D列 出庫数','2025/10/20在庫欄に対応する元台帳。出庫日は特定不可'],
    ['LEGACY-INDEN-08','P00468',94,'','在庫2024~.xlsx / 作業用 衣類 / D列 出庫数','2025/10/20在庫欄に対応する元台帳。出庫日は特定不可']
  ];
  const existing=new Set(readRows_(sh).map(x=>String(x['実績ID'])));
  const add=rows.filter(r=>!existing.has(r[0]));
  if(add.length) sh.getRange(sh.getLastRow()+1,1,add.length,7).setValues(add.map(r=>[...r,new Date()]));
  log_('過去出庫実績取込','P00461-P00468',`印伝風 鍔止 ${add.length}件追加 / 合計776点（出庫日不明・別枠管理）`);
  return `取込完了：印伝風 鍔止の過去出庫実績776点を日付不明データとして登録しました。新規${add.length}件。日別・月別出庫や90日平均には加算していません。`;
}

// 2026-11-20に受領したExcel 3シートの「注文残」を初期受注残として使用。
// 商品・サイズごとの数量を保持し、初回だけ登録する。既存の在庫・入出庫・棚卸データには触れない。
const INITIAL_ORDER_BACKLOG_20261120 = [["2024 防具1","隼1　面","小小",0],["2024 防具1","隼1　面","小",40],["2024 防具1","隼1　面","中",10],["2024 防具1","隼1　面","大",0],["2024 防具1","隼1　胴","小小",30],["2024 防具1","隼1　胴","小",30],["2024 防具1","隼1　胴","中",0],["2024 防具1","隼1　胴","大",0],["2024 防具1","隼1　赤胴","小小",0],["2024 防具1","隼1　赤胴","小",0],["2024 防具1","隼1　赤胴","中",0],["2024 防具1","隼1　小手","小小",30],["2024 防具1","隼1　小手","小",60],["2024 防具1","隼1　小手","中",20],["2024 防具1","隼1　小手","大",30],["2024 防具1","隼1　垂","小小",0],["2024 防具1","隼1　垂","小",10],["2024 防具1","隼1　垂","中",0],["2024 防具1","隼1　垂","大",10],["2024 防具1","隼1　垂","中　濃紺",0],["2024 防具1","隼1　垂","中　エンジ",0],["2024 防具1","隼1　垂","大　濃紺",0],["2024 防具1","隼1　垂","大　エンジ",0],["2024 防具1","隼1　垂","大大　濃紺",0],["2024 防具1","隼1　垂","大大　エンジ",0],["2024 防具1","鷹1　小手","中",0],["2024 防具1","鷹1　小手","大",0],["2024 防具1","鷹1　小手","大大",0],["2024 防具1","鷹1　小手","中　濃紺",0],["2024 防具1","鷹1　小手","大　濃紺",0],["2024 防具1","鷹1　小手","大大　濃紺",0],["2024 防具1","究極　面","中",0],["2024 防具1","究極　面","大",0],["2024 防具1","究極　面","大大",0],["2024 防具1","鶴1　面","小",0],["2024 防具1","鶴1　面","中",0],["2024 防具1","鶴1　面","大",0],["2024 防具1","鶴1　小手(薙刀白)","小",0],["2024 防具1","鶴1　小手(薙刀白)","中",0],["2024 防具1","鶴1　小手(薙刀白)","大",0],["2024 防具1","鶴2　面","中",0],["2024 防具1","鶴2　面","大",0],["2024 防具1","鶴2　小手(薙刀紺)","小",20],["2024 防具1","鶴2　小手(薙刀紺)","中",40],["2024 防具1","鶴2　小手(薙刀紺)","大",20],["2024 防具1","鶴2　小手(薙刀紺)","学販小手　生子入り",300],["2024 防具1","鶴2　小手(薙刀紺)","学販小手　生子なし",450],["2024 防具1","鶴2　小手(薙刀紺)","メンピット　M",0],["2024 防具1","鶴2　小手(薙刀紺)","メンピット　L",800],["2024 防具1","隼1T　面(内輪固定式)","小小",0],["2024 防具1","隼1T　面(内輪固定式)","65～66㎝",25],["2024 防具1","隼1T　面(内輪固定式)","67～68㎝",44],["2024 防具1","隼1T　面(内輪固定式)","69～70㎝",44],["2024 防具1","隼1T　面(内輪固定式)","71～72㎝",4],["2024 防具2","燕　面 ￥7,450","60㎝",0],["2024 防具2","燕　面 ￥7,450","61㎝",0],["2024 防具2","燕　面 ￥7,450","62㎝",16],["2024 防具2","燕　面 ￥7,450","63㎝",16],["2024 防具2","燕　面 ￥7,450","64㎝",24],["2024 防具2","燕　面 ￥7,450","65㎝",8],["2024 防具2","燕　面 ￥7,450","66㎝",32],["2024 防具2","燕　面 ￥7,450","67㎝",32],["2024 防具2","燕　面 ￥7,450","68㎝",0],["2024 防具2","燕　面 ￥7,450","69㎝",0],["2024 防具2","燕　面 ￥7,450","70㎝",8],["2024 防具2","燕　面 ￥7,450","71㎝",16],["2024 防具2","燕　面 ￥7,450","72㎝",0],["2024 防具2","燕　面 ￥7,450","73㎝",0],["2024 防具2","燕　小手 ￥4,177","SS",20],["2024 防具2","燕　小手 ￥4,177","S",180],["2024 防具2","燕　小手 ￥4,177","M",360],["2024 防具2","燕　小手 ￥4,177","L",360],["2024 防具2","燕　小手 ￥4,177","X",40],["2024 防具2","燕　小手 ￥4,177","XL",20],["2024 防具2","燕　垂 ￥3,571","SS",0],["2024 防具2","燕　垂 ￥3,571","S",0],["2024 防具2","燕　垂 ￥3,571","M",90],["2024 防具2","燕　垂 ￥3,571","L",30],["2024 防具2","燕　垂 ￥3,571","X",0],["2024 防具2","燕　垂 ￥3,571","XL",0],["2024 防具2","鷹1T　面　濃紺 ￥8,405","60㎝",0],["2024 防具2","鷹1T　面　濃紺 ￥8,405","61㎝",0],["2024 防具2","鷹1T　面　濃紺 ￥8,405","62㎝",0],["2024 防具2","鷹1T　面　濃紺 ￥8,405","63㎝",0],["2024 防具2","鷹1T　面　濃紺 ￥8,405","64㎝",0],["2024 防具2","鷹1T　面　濃紺 ￥8,405","65㎝",0],["2024 防具2","鷹1T　面　濃紺 ￥8,405","66㎝",24],["2024 防具2","鷹1T　面　濃紺 ￥8,405","67㎝",28],["2024 防具2","鷹1T　面　濃紺 ￥8,405","68㎝",0],["2024 防具2","鷹1T　面　濃紺 ￥8,405","69㎝",0],["2024 防具2","鷹1T　面　濃紺 ￥8,405","70㎝",0],["2024 防具2","鷹1T　面　濃紺 ￥8,405","71㎝",8],["2024 防具2","鷹1T　面　濃紺 ￥8,405","72㎝",0],["2024 防具2","鷹1T　面　濃紺 ￥8,405","73㎝",0],["2024 防具2","鷹1T　面　エンジ ￥8,405","60㎝",0],["2024 防具2","鷹1T　面　エンジ ￥8,405","61㎝",0],["2024 防具2","鷹1T　面　エンジ ￥8,405","62㎝",0],["2024 防具2","鷹1T　面　エンジ ￥8,405","63㎝",0],["2024 防具2","鷹1T　面　エンジ ￥8,405","64㎝",0],["2024 防具2","鷹1T　面　エンジ ￥8,405","65㎝",0],["2024 防具2","鷹1T　面　エンジ ￥8,405","66㎝",0],["2024 防具2","鷹1T　面　エンジ ￥8,405","67㎝",0],["2024 防具2","鷹1T　面　エンジ ￥8,405","68㎝",0],["2024 防具2","鷹1T　面　エンジ ￥8,405","69㎝",0],["2024 防具2","鷹1T　面　エンジ ￥8,405","70㎝",0],["2024 防具2","鷹1T　面　エンジ ￥8,405","71㎝",0],["2024 防具2","鷹1T　面　エンジ ￥8,405","72㎝",0],["2024 防具2","鷹1T　面　エンジ ￥8,405","73㎝",0],["2024 防具2","鷹1T　小手(赤姫) ￥4,498","SS",0],["2024 防具2","鷹1T　小手(赤姫) ￥4,498","S",20],["2024 防具2","鷹1T　小手(赤姫) ￥4,498","M",140],["2024 防具2","鷹1T　小手(赤姫) ￥4,498","L",140],["2024 防具2","鷹1T　小手(赤姫) ￥4,498","X",20],["2024 防具2","鷹1T　小手(赤姫) ￥4,498","XL",20],["2024 防具2","鷹1T　垂　濃紺 ￥4,522","SS",0],["2024 防具2","鷹1T　垂　濃紺 ￥4,522","S",0],["2024 防具2","鷹1T　垂　濃紺 ￥4,522","M",0],["2024 防具2","鷹1T　垂　濃紺 ￥4,522","L",60],["2024 防具2","鷹1T　垂　濃紺 ￥4,522","X",0],["2024 防具2","鷹1T　垂　濃紺 ￥4,522","XL",0],["2024 防具2","鷹1T　垂　エンジ(赤姫) ￥4,522","SS",0],["2024 防具2","鷹1T　垂　エンジ(赤姫) ￥4,522","S",0],["2024 防具2","鷹1T　垂　エンジ(赤姫) ￥4,522","M",0],["2024 防具2","鷹1T　垂　エンジ(赤姫) ￥4,522","L",0],["2024 防具2","鷹1T　垂　エンジ(赤姫) ￥4,522","X",0],["2024 防具2","鷹1T　垂　エンジ(赤姫) ￥4,522","XL",0],["2024 防具2","共用胴　濃紺鬼雲 ￥6,137","SS",0],["2024 防具2","共用胴　濃紺鬼雲 ￥6,137","S",0],["2024 防具2","共用胴　濃紺鬼雲 ￥6,137","M",60],["2024 防具2","共用胴　濃紺鬼雲 ￥6,137","L",60],["2024 防具2","共用胴　濃紺鬼雲 ￥6,137","X",15],["2024 防具2","共用胴　濃紺鬼雲 ￥6,137","XL",39],["2024 防具2","共用胴　エンジ鬼雲 ￥6,137","SS",0],["2024 防具2","共用胴　エンジ鬼雲 ￥6,137","S",0],["2024 防具2","共用胴　エンジ鬼雲 ￥6,137","M",0],["2024 防具2","共用胴　エンジ鬼雲 ￥6,137","L",0],["2024 防具2","共用胴　エンジ鬼雲 ￥6,137","X",0],["2024 防具2","共用胴　エンジ鬼雲 ￥6,137","XL",0],["2024 防具2","共用胴　濃紺ベタ ￥6,137","SS",0],["2024 防具2","共用胴　濃紺ベタ ￥6,137","S",0],["2024 防具2","共用胴　濃紺ベタ ￥6,137","M",0],["2024 防具2","共用胴　濃紺ベタ ￥6,137","L",0],["2024 防具2","共用胴　濃紺ベタ ￥6,137","X",0],["2024 防具2","共用胴　濃紺ベタ ￥6,137","XL",0],["2024 防具2","共用胴　エンジベタ ￥6,137","SS",0],["2024 防具2","共用胴　エンジベタ ￥6,137","S",0],["2024 防具2","共用胴　エンジベタ ￥6,137","M",0],["2024 防具2","共用胴　エンジベタ ￥6,137","L",0],["2024 防具2","共用胴　エンジベタ ￥6,137","X",0],["2024 防具2","共用胴　エンジベタ ￥6,137","XL",0],["2024 防具2","鷹2T　面　濃紺 ￥10,497","60㎝",0],["2024 防具2","鷹2T　面　濃紺 ￥10,497","61㎝",0],["2024 防具2","鷹2T　面　濃紺 ￥10,497","62㎝",0],["2024 防具2","鷹2T　面　濃紺 ￥10,497","63㎝",0],["2024 防具2","鷹2T　面　濃紺 ￥10,497","64㎝",0],["2024 防具2","鷹2T　面　濃紺 ￥10,497","65㎝",0],["2024 防具2","鷹2T　面　濃紺 ￥10,497","66㎝",0],["2024 防具2","鷹2T　面　濃紺 ￥10,497","67㎝",8],["2024 防具2","鷹2T　面　濃紺 ￥10,497","68㎝",0],["2024 防具2","鷹2T　面　濃紺 ￥10,497","69㎝",8],["2024 防具2","鷹2T　面　濃紺 ￥10,497","70㎝",0],["2024 防具2","鷹2T　面　濃紺 ￥10,497","71㎝",0],["2024 防具2","鷹2T　面　濃紺 ￥10,497","72㎝",0],["2024 防具2","鷹2T　面　濃紺 ￥10,497","73㎝",0],["2024 防具2","鷹2T　面　エンジ ￥10,497","60㎝",0],["2024 防具2","鷹2T　面　エンジ ￥10,497","61㎝",0],["2024 防具2","鷹2T　面　エンジ ￥10,497","62㎝",0],["2024 防具2","鷹2T　面　エンジ ￥10,497","63㎝",0],["2024 防具2","鷹2T　面　エンジ ￥10,497","64㎝",0],["2024 防具2","鷹2T　面　エンジ ￥10,497","65㎝",0],["2024 防具2","鷹2T　面　エンジ ￥10,497","66㎝",0],["2024 防具2","鷹2T　面　エンジ ￥10,497","67㎝",0],["2024 防具2","鷹2T　面　エンジ ￥10,497","68㎝",0],["2024 防具2","鷹2T　面　エンジ ￥10,497","69㎝",0],["2024 防具2","鷹2T　面　エンジ ￥10,497","70㎝",0],["2024 防具2","鷹2T　面　エンジ ￥10,497","71㎝",0],["2024 防具2","鷹2T　面　エンジ ￥10,497","72㎝",0],["2024 防具2","鷹2T　面　エンジ ￥10,497","73㎝",0],["2024 防具2","鷹2T　小手　茶鹿革 \\7,385","SS",0],["2024 防具2","鷹2T　小手　茶鹿革 \\7,385","S",0],["2024 防具2","鷹2T　小手　茶鹿革 \\7,385","M",0],["2024 防具2","鷹2T　小手　茶鹿革 \\7,385","L",40],["2024 防具2","鷹2T　小手　茶鹿革 \\7,385","X",10],["2024 防具2","鷹2T　小手　茶鹿革 \\7,385","XL",10],["2024 防具2","鷹2T　小手　クラレ ￥5,141","M",0],["2024 防具2","鷹2T　小手　クラレ ￥5,141","L",40],["2024 防具2","鷹2T　小手　クラレ ￥5,141","X",0],["2024 防具2","鷹2T　小手　クラレ ￥5,141","XL",0],["2024 防具2","鷹2T　垂　濃紺 ￥5,366","SS",0],["2024 防具2","鷹2T　垂　濃紺 ￥5,366","S",0],["2024 防具2","鷹2T　垂　濃紺 ￥5,366","M",0],["2024 防具2","鷹2T　垂　濃紺 ￥5,366","L",30],["2024 防具2","鷹2T　垂　濃紺 ￥5,366","X",0],["2024 防具2","鷹2T　垂　濃紺 ￥5,366","XL",0],["2024 防具2","鷹2T　垂　エンジ ￥5,366","SS",0],["2024 防具2","鷹2T　垂　エンジ ￥5,366","S",0],["2024 防具2","鷹2T　垂　エンジ ￥5,366","M",0],["2024 防具2","鷹2T　垂　エンジ ￥5,366","L",0],["2024 防具2","鷹2T　垂　エンジ ￥5,366","X",0],["2024 防具2","鷹2T　垂　エンジ ￥5,366","XL",0],["2024 防具2","無双　小手 ￥4,884","S",0],["2024 防具2","無双　小手 ￥4,884","M",0],["2024 防具2","無双　小手 ￥4,884","L",0],["2024 防具2","無双　小手 ￥4,884","XL",0],["2024 防具2","赤姫　面 ￥9,067","58㎝",0],["2024 防具2","赤姫　面 ￥9,067","59㎝",0],["2024 防具2","赤姫　面 ￥9,067","60㎝",0],["2024 防具2","赤姫　面 ￥9,067","61㎝",0],["2024 防具2","赤姫　面 ￥9,067","62㎝",0],["2024 防具2","赤姫　面 ￥9,067","63㎝",0],["2024 防具2","赤姫　面 ￥9,067","64㎝",0],["2024 防具2","赤姫　面 ￥9,067","65㎝",0],["2024 防具2","赤姫　面 ￥9,067","66㎝",0],["2024 防具2","赤姫　面 ￥9,067","67㎝",0],["2024 防具2","赤姫　面 ￥9,067","68㎝",0],["2024 防具2","赤姫　面 ￥9,067","69㎝",0],["2024 防具2","赤姫　面 ￥9,067","70㎝",0],["2024 防具2","赤姫　面 ￥9,067","71㎝",0],["2024 防具2","赤姫　黒胴 ￥8,630","SS",0],["2024 防具2","赤姫　黒胴 ￥8,630","S",0],["2024 防具2","赤姫　黒胴 ￥8,630","M",0],["2024 防具2","赤姫　黒胴 ￥8,630","L",0],["2024 防具2","赤姫　黒胴 ￥8,630","XL",0],["2024 防具2","赤姫　玉虫胴 ￥11,053","SS",0],["2024 防具2","赤姫　玉虫胴 ￥11,053","S",0],["2024 防具2","赤姫　玉虫胴 ￥11,053","M",0],["2024 防具2","赤姫　玉虫胴 ￥11,053","L",0],["2024 防具2","赤姫　玉虫胴 ￥11,053","XL",0],["2024 防具2","疾風　面 ￥8,405","58㎝",0],["2024 防具2","疾風　面 ￥8,405","59㎝",0],["2024 防具2","疾風　面 ￥8,405","60㎝",0],["2024 防具2","疾風　面 ￥8,405","61㎝",0],["2024 防具2","疾風　面 ￥8,405","62㎝",0],["2024 防具2","疾風　面 ￥8,405","63㎝",0],["2024 防具2","疾風　面 ￥8,405","64㎝",0],["2024 防具2","疾風　面 ￥8,405","65㎝",0],["2024 防具2","疾風　面 ￥8,405","66㎝",0],["2024 防具2","疾風　面 ￥8,405","67㎝",0],["2024 防具2","疾風　面 ￥8,405","68㎝",0],["2024 防具2","疾風　面 ￥8,405","69㎝",0],["2024 防具2","疾風　面 ￥8,405","70㎝",0],["2024 防具2","疾風　面 ￥8,405","71㎝",0],["2024 防具2","疾風　面 ￥8,405","72㎝",0],["2024 防具2","疾風　面 ￥8,405","73㎝",0],["2024 防具2","疾風　胴 \\6,137","SS",0],["2024 防具2","疾風　胴 \\6,137","S",0],["2024 防具2","疾風　胴 \\6,137","M",0],["2024 防具2","疾風　胴 \\6,137","L",0],["2024 防具2","疾風　胴 \\6,137","X",0],["2024 防具2","疾風　胴 \\6,137","XL",0],["2024 防具2","疾風　小手 ￥4,498","SS",0],["2024 防具2","疾風　小手 ￥4,498","S",0],["2024 防具2","疾風　小手 ￥4,498","M",0],["2024 防具2","疾風　小手 ￥4,498","L",0],["2024 防具2","疾風　小手 ￥4,498","X",0],["2024 防具2","疾風　小手 ￥4,498","XL",0],["2024 防具2","疾風　垂 ￥4,522","SS",0],["2024 防具2","疾風　垂 ￥4,522","S",0],["2024 防具2","疾風　垂 ￥4,522","M",0],["2024 防具2","疾風　垂 ￥4,522","L",0],["2024 防具2","疾風　垂 ￥4,522","X",0],["2024 防具2","疾風　垂 ￥4,522","XL",0],["2024 防具2","波千鳥胴　エンジ　溜石目$90","S",0],["2024 防具2","波千鳥胴　エンジ　溜石目$90","M",5],["2024 防具2","波千鳥胴　エンジ　溜石目$90","L",10],["2024 防具2","波千鳥胴　エンジ　溜石目$90","XL",5],["2024 防具2","波千鳥胴　古代紫　黒石目$90","S",0],["2024 防具2","波千鳥胴　古代紫　黒石目$90","M",20],["2024 防具2","波千鳥胴　古代紫　黒石目$90","L",45],["2024 防具2","波千鳥胴　古代紫　黒石目$90","XL",20],["2024 防具2","黒雲龍　胴  \\10,496","SS",0],["2024 防具2","黒雲龍　胴  \\10,496","S",0],["2024 防具2","黒雲龍　胴  \\10,496","M",0],["2024 防具2","黒雲龍　胴  \\10,496","L",0],["2024 防具2","黒雲龍　胴  \\10,496","XL",6],["2024 防具2","新1号　胴 \\11,000","M",0],["2024 防具2","新1号　胴 \\11,000","L",0],["2024 防具2","新1号　胴 \\11,000","XL",0],["2024 防具2","黒鮫　胴 ￥12,500","M",0],["2024 防具2","黒鮫　胴 ￥12,500","L",0],["2024 防具2","黒鮫　胴 ￥12,500","XL",0],["2024 防具2","V1　小手 ＄45","M",0],["2024 防具2","V1　小手 ＄45","L",0],["2024 防具2","V1　小手 ＄45","XL",10],["2024 防具2","7mm　面 ￥6,962","60㎝",0],["2024 防具2","7mm　面 ￥6,962","61㎝",0],["2024 防具2","7mm　面 ￥6,962","62㎝",0],["2024 防具2","7mm　面 ￥6,962","63㎝",0],["2024 防具2","7mm　面 ￥6,962","64㎝",0],["2024 防具2","7mm　面 ￥6,962","65㎝",0],["2024 防具2","7mm　面 ￥6,962","66㎝",0],["2024 防具2","7mm　面 ￥6,962","67㎝",0],["2024 防具2","7mm　面 ￥6,962","68㎝",0],["2024 防具2","7mm　面 ￥6,962","69㎝",0],["2024 防具2","7mm　面 ￥6,962","70㎝",0],["2024 防具2","7mm　面 ￥6,962","71㎝",0],["2024 防具2","7mm　面 ￥6,962","72㎝",0],["2024 防具2","7mm　面 ￥6,962","73㎝",0],["2024 防具2","7mm　小手 ￥3,213","SS",0],["2024 防具2","7mm　小手 ￥3,213","S",0],["2024 防具2","7mm　小手 ￥3,213","M",0],["2024 防具2","7mm　小手 ￥3,213","L",0],["2024 防具2","7mm　小手 ￥3,213","X",0],["2024 防具2","7mm　小手 ￥3,213","XL",0],["2024 防具2","7mm　垂 ￥2,892","SS",0],["2024 防具2","7mm　垂 ￥2,892","S",0],["2024 防具2","7mm　垂 ￥2,892","M",0],["2024 防具2","7mm　垂 ￥2,892","L",0],["2024 防具2","7mm　垂 ￥2,892","X",0],["2024 防具2","7mm　垂 ￥2,892","XL",0],["2024 防具2","(胸)鬼雲 碁盤刺＃29 ￥4,605","SS",0],["2024 防具2","(胸)鬼雲 碁盤刺＃29 ￥4,605","S",0],["2024 防具2","(胸)鬼雲 碁盤刺＃29 ￥4,605","M",0],["2024 防具2","(胸)鬼雲 碁盤刺＃29 ￥4,605","L",0],["2024 防具2","(胸)兜 碁盤刺＃29 ￥4,605","SS",0],["2024 防具2","(胸)兜 碁盤刺＃29 ￥4,605","S",0],["2024 防具2","(胸)兜 碁盤刺＃29 ￥4,605","M",0],["2024 防具2","(胸)兜 碁盤刺＃29 ￥4,605","L",0],["2024 防具2","(胸)山飾 碁盤刺＃29 ￥4,605","SS",0],["2024 防具2","(胸)山飾 碁盤刺＃29 ￥4,605","S",0],["2024 防具2","(胸)山飾 碁盤刺＃29 ￥4,605","M",0],["2024 防具2","(胸)山飾 碁盤刺＃29 ￥4,605","L",0],["2024 衣類","ガラムジャージ紺","000号",50],["2024 衣類","ガラムジャージ紺","00号",30],["2024 衣類","ガラムジャージ紺","0号",90],["2024 衣類","ガラムジャージ紺","0.5号",0],["2024 衣類","ガラムジャージ紺","1号",180],["2024 衣類","ガラムジャージ紺","1.5号",300],["2024 衣類","ガラムジャージ紺","2号",450],["2024 衣類","ガラムジャージ紺","2.5号",360],["2024 衣類","ガラムジャージ紺","3号",300],["2024 衣類","ガラムジャージ紺","3.5号",0],["2024 衣類","ガラムジャージ紺","4号",90],["2024 衣類","ガラムジャージ紺","4.5号",30],["2024 衣類","ガラムジャージ紺","5号",0],["2024 衣類","ガラムジャージ白","000号",10],["2024 衣類","ガラムジャージ白","00号",0],["2024 衣類","ガラムジャージ白","0号",60],["2024 衣類","ガラムジャージ白","0.5号",0],["2024 衣類","ガラムジャージ白","1号",30],["2024 衣類","ガラムジャージ白","1.5号",0],["2024 衣類","ガラムジャージ白","2号",90],["2024 衣類","ガラムジャージ白","2.5号",90],["2024 衣類","ガラムジャージ白","3号",0],["2024 衣類","ガラムジャージ白","3.5号",0],["2024 衣類","ガラムジャージ白","4号",0],["2024 衣類","ガラムジャージ白","4.5号",0],["2024 衣類","ガラムジャージ白","5号",0],["2024 衣類","紺一重","000号",0],["2024 衣類","紺一重","00号",0],["2024 衣類","紺一重","0号",0],["2024 衣類","紺一重","0.5号",0],["2024 衣類","紺一重","1号",10],["2024 衣類","紺一重","1.5号",0],["2024 衣類","紺一重","2号",30],["2024 衣類","紺一重","2.5号",0],["2024 衣類","紺一重","3号",350],["2024 衣類","紺一重","3.5号",0],["2024 衣類","紺一重","4号",175],["2024 衣類","紺一重","4.5号",0],["2024 衣類","紺一重","5号",20],["2024 衣類","白一重","000号",0],["2024 衣類","白一重","00号",0],["2024 衣類","白一重","0号",0],["2024 衣類","白一重","0.5号",0],["2024 衣類","白一重","1号",0],["2024 衣類","白一重","1.5号",15],["2024 衣類","白一重","2号",30],["2024 衣類","白一重","2.5号",30],["2024 衣類","白一重","3号",0],["2024 衣類","白一重","3.5号",0],["2024 衣類","白一重","4号",0],["2024 衣類","白一重","4.5号",15],["2024 衣類","白一重","5号",0],["2024 衣類","涼夏","00号",0],["2024 衣類","涼夏","0号",0],["2024 衣類","涼夏","0.5号",0],["2024 衣類","涼夏","1号",0],["2024 衣類","涼夏","1.5号",0],["2024 衣類","涼夏","2号",0],["2024 衣類","涼夏","2.5号",0],["2024 衣類","涼夏","3号",0],["2024 衣類","涼夏","3.5号",0],["2024 衣類","涼夏","4号",0],["2024 衣類","涼夏","4.5号",0],["2024 衣類","涼夏","5号",0],["2024 衣類","通気性　紺","00号",0],["2024 衣類","通気性　紺","0号",0],["2024 衣類","通気性　紺","0.5号",0],["2024 衣類","通気性　紺","1号",0],["2024 衣類","通気性　紺","1.5号",0],["2024 衣類","通気性　紺","2号",0],["2024 衣類","通気性　紺","2.5号",0],["2024 衣類","通気性　紺","3号",0],["2024 衣類","通気性　紺","3.5号",0],["2024 衣類","通気性　紺","4号",0],["2024 衣類","通気性　紺","4.5号",0],["2024 衣類","通気性　紺","5号",0],["2024 衣類","通気性　白","00号",0],["2024 衣類","通気性　白","0号",0],["2024 衣類","通気性　白","0.5号",0],["2024 衣類","通気性　白","1号",0],["2024 衣類","通気性　白","1.5号",0],["2024 衣類","通気性　白","2号",0],["2024 衣類","通気性　白","2.5号",0],["2024 衣類","通気性　白","3号",0],["2024 衣類","通気性　白","3.5号",0],["2024 衣類","通気性　白","4号",0],["2024 衣類","通気性　白","4.5号",0],["2024 衣類","通気性　白","5号",0],["2024 衣類","安吉ジャージ黒","00号",0],["2024 衣類","安吉ジャージ黒","0号",0],["2024 衣類","安吉ジャージ黒","0.5号",0],["2024 衣類","安吉ジャージ黒","1号",0],["2024 衣類","安吉ジャージ黒","1.5号",0],["2024 衣類","安吉ジャージ黒","2号",0],["2024 衣類","安吉ジャージ黒","2.5号",0],["2024 衣類","安吉ジャージ黒","3号",0],["2024 衣類","安吉ジャージ黒","3.5号",0],["2024 衣類","安吉ジャージ黒","4号",0],["2024 衣類","安吉ジャージ黒","4.5号",0],["2024 衣類","安吉ジャージ黒","5号",0],["2024 衣類","安吉ジャージ白","00号",0],["2024 衣類","安吉ジャージ白","0号",0],["2024 衣類","安吉ジャージ白","0.5号",0],["2024 衣類","安吉ジャージ白","1号",0],["2024 衣類","安吉ジャージ白","1.5号",0],["2024 衣類","安吉ジャージ白","2号",0],["2024 衣類","安吉ジャージ白","2.5号",0],["2024 衣類","安吉ジャージ白","3号",0],["2024 衣類","安吉ジャージ白","3.5号",0],["2024 衣類","安吉ジャージ白","4号",0],["2024 衣類","安吉ジャージ白","4.5号",0],["2024 衣類","安吉ジャージ白","5号",0],["2024 衣類","安吉ジャージ白","16号",25],["2024 衣類","安吉ジャージ白","17号",25],["2024 衣類","安吉ジャージ白","18号",25],["2024 衣類","安吉ジャージ白","19号",50],["2024 衣類","安吉ジャージ白","20号",0],["2024 衣類","安吉ジャージ白","21号",225],["2024 衣類","安吉ジャージ白","22号",175],["2024 衣類","安吉ジャージ白","23号",275],["2024 衣類","安吉ジャージ白","24号",350],["2024 衣類","安吉ジャージ白","25号",425],["2024 衣類","安吉ジャージ白","26号",375],["2024 衣類","安吉ジャージ白","27号",150],["2024 衣類","安吉ジャージ白","28号",25],["2024 衣類","TR袴　紺","16号",60],["2024 衣類","TR袴　紺","17号",0],["2024 衣類","TR袴　紺","18号",25],["2024 衣類","TR袴　紺","19号",100],["2024 衣類","TR袴　紺","20号",100],["2024 衣類","TR袴　紺","21号",125],["2024 衣類","TR袴　紺","22号",150],["2024 衣類","TR袴　紺","23号",225],["2024 衣類","TR袴　紺","24号",275],["2024 衣類","TR袴　紺","25号",375],["2024 衣類","TR袴　紺","26号",225],["2024 衣類","TR袴　紺","27号",175],["2024 衣類","TR袴　紺","28号",100],["2024 衣類","TR袴　黒","16号",0],["2024 衣類","TR袴　黒","17号",30],["2024 衣類","TR袴　黒","18号",0],["2024 衣類","TR袴　黒","19号",60],["2024 衣類","TR袴　黒","20号",0],["2024 衣類","TR袴　黒","21号",75],["2024 衣類","TR袴　黒","22号",75],["2024 衣類","TR袴　黒","23号",125],["2024 衣類","TR袴　黒","24号",125],["2024 衣類","TR袴　黒","25号",125],["2024 衣類","TR袴　黒","26号",100],["2024 衣類","TR袴　黒","27号",25],["2024 衣類","TR袴　黒","28号",0],["2024 衣類","TR袴　白","16号",25],["2024 衣類","TR袴　白","17号",10],["2024 衣類","TR袴　白","18号",10],["2024 衣類","TR袴　白","19号",25],["2024 衣類","TR袴　白","20号",25],["2024 衣類","TR袴　白","21号",25],["2024 衣類","TR袴　白","22号",25],["2024 衣類","TR袴　白","23号",25],["2024 衣類","TR袴　白","24号",25],["2024 衣類","TR袴　白","25号",15],["2024 衣類","TR袴　白","26号",30],["2024 衣類","TR袴　白","27号",0],["2024 衣類","TR袴　白","28号",0],["2024 衣類","TC袴","16号",0],["2024 衣類","TC袴","17号",0],["2024 衣類","TC袴","18号",0],["2024 衣類","TC袴","19号",0],["2024 衣類","TC袴","20号",0],["2024 衣類","TC袴","21号",0],["2024 衣類","TC袴","22号",0],["2024 衣類","TC袴","23号",0],["2024 衣類","TC袴","24号",0],["2024 衣類","TC袴","25号",0],["2024 衣類","TC袴","26号",0],["2024 衣類","TC袴","27号",0],["2024 衣類","TC袴","28号",0],["2024 衣類","＃8000袴","23号",0],["2024 衣類","＃8000袴","23.5号",0],["2024 衣類","＃8000袴","24号",0],["2024 衣類","＃8000袴","24.5号",0],["2024 衣類","＃8000袴","25号",0],["2024 衣類","＃8000袴","25.5号",0],["2024 衣類","＃8000袴","26号",0],["2024 衣類","＃8000袴","26.5号",0],["2024 衣類","＃8000袴","27号",0],["2024 衣類","＃8000袴","27.5号",0],["2024 衣類","＃8000袴","28号",0],["2024 衣類","＃10000袴","23号",0],["2024 衣類","＃10000袴","23.5号",0],["2024 衣類","＃10000袴","24号",0],["2024 衣類","＃10000袴","24.5号",0],["2024 衣類","＃10000袴","25号",0],["2024 衣類","＃10000袴","25.5号",0],["2024 衣類","＃10000袴","26号",0],["2024 衣類","＃10000袴","26.5号",0],["2024 衣類","＃10000袴","27号",0],["2024 衣類","＃10000袴","27.5号",0],["2024 衣類","＃10000袴","28号",0],["2024 衣類","印伝風 鍔止","1.黒×金蜻蛉",0],["2024 衣類","印伝風 鍔止","2.黒×銀蜻蛉",50],["2024 衣類","印伝風 鍔止","3.黒×黒蜻蛉",100],["2024 衣類","印伝風 鍔止","4.黒×青蜻蛉",50],["2024 衣類","印伝風 鍔止","5.黒×黄蜻蛉",0],["2024 衣類","印伝風 鍔止","6.黒×赤蜻蛉",50],["2024 衣類","印伝風 鍔止","7.黒×赤青蜻蛉",100],["2024 衣類","印伝風 鍔止","8.黒×黒赤蜻蛉",100],["2024 衣類","印伝風 鍔止","9.赤×白蜻蛉",0],["2024 衣類","印伝風 鍔止","10.黒×金小桜",50],["2024 衣類","印伝風 鍔止","11.黒×銀小桜",0],["2024 衣類","印伝風 鍔止","12.黒×黒小桜",0],["2024 衣類","印伝風 鍔止","13.黒×青小桜",100],["2024 衣類","印伝風 鍔止","14.黒×紫小桜",0],["2024 衣類","印伝風 鍔止","15.黒×赤小桜",0],["2024 衣類","印伝風 鍔止","16.黒×黒赤小桜",100],["2024 衣類","印伝風 鍔止","17.赤×白小桜",0],["2024 衣類","印伝風 鍔止","18.黒×青瓢箪",0],["2024 衣類","印伝風 鍔止","19.黒×赤瓢箪",0],["2024 衣類","印伝風 鍔止","20.黒×黒瓢箪",0],["2024 衣類","印伝風 鍔止","21.黒×金波",50],["2024 衣類","印伝風 鍔止","22.黒×銀波",0],["2024 衣類","印伝風 鍔止","23.黒×黒波",0]];
function ensureBacklogSheet_(){
  const ss=SpreadsheetApp.getActive(); let sh=ss.getSheetByName(APP.SHEETS.backlog);
  if(!sh) sh=ss.insertSheet(APP.SHEETS.backlog);
  if(sh.getLastRow()===0){sh.getRange(1,1,1,APP.HEADERS.backlog.length).setValues([APP.HEADERS.backlog]);sh.setFrozenRows(1);}
  if(sh.getLastRow()<2) seedInitialOrderBacklog_();
  return sh;
}
function backlogNorm_(v){return String(v||'').normalize('NFKC').replace(/[￥¥$＄\\][\s\d,\.]+.*$/,'').replace(/[\s　]+/g,'').replace(/[()（）]/g,'').toLowerCase();}
function backlogSizeNorm_(v){return String(v||'').normalize('NFKC').replace(/[\s　]+/g,'').toLowerCase();}
function seedInitialOrderBacklog_(){
  const ss=SpreadsheetApp.getActive(), sh=ss.getSheetByName(APP.SHEETS.backlog);
  const ps=readRows_(ss.getSheetByName(APP.SHEETS.products)); if(!ps.length)return 0;
  const map={}; ps.forEach(p=>{const k=backlogNorm_(p['商品名'])+'|'+backlogSizeNorm_(p['サイズ']); if(!map[k])map[k]=p;});
  const grouped={};
  INITIAL_ORDER_BACKLOG_20261120.forEach((r,i)=>{
    const [source,name,size,qty]=r, p=map[backlogNorm_(name)+'|'+backlogSizeNorm_(size)];
    const key=p?('ID:'+String(p['商品ID'])):('SRC:'+source+'|'+backlogNorm_(name)+'|'+backlogSizeNorm_(size));
    if(!grouped[key])grouped[key]={productId:p?String(p['商品ID']):'',name,size,initial:0,sources:[],index:i+1};
    grouped[key].initial+=Number(qty)||0; if(!grouped[key].sources.includes(source))grouped[key].sources.push(source);
  });
  const rows=Object.keys(grouped).map((k)=>{const x=grouped[k];return ['OB-INIT-'+String(x.index).padStart(4,'0'),x.productId,x.name,x.size,x.initial,x.initial,x.sources.join(' / '),new Date()];});
  if(rows.length)sh.getRange(2,1,rows.length,8).setValues(rows);
  log_('初期受注残取込','ALL',`Excel3シート ${INITIAL_ORDER_BACKLOG_20261120.length}行 / ${rows.length}バリエーション`);
  return rows.length;
}
function getBacklogRows_(){
  const sh=SpreadsheetApp.getActive().getSheetByName(APP.SHEETS.backlog); if(!sh)return [];
  return readRows_(sh).map(x=>({id:String(x['受注残ID']||''),productId:String(x['商品ID']||''),name:String(x['商品名']||''),size:String(x['サイズ']||''),initial:Number(x['初期受注残'])||0,remaining:Number(x['現在受注残'])||0,source:String(x['元シート']||'')}));
}
function adjustBacklog_(productId,delta,name,size){
  const sh=ensureBacklogSheet_(); const last=sh.getLastRow(); if(last<2)return;
  const vals=sh.getRange(2,1,last-1,8).getValues(); const pid=String(productId||'');
  let hit=-1;
  for(let i=0;i<vals.length;i++) if(pid && String(vals[i][1]||'')===pid){hit=i;break;}
  if(hit<0){const nk=backlogNorm_(name),sk=backlogSizeNorm_(size);for(let i=0;i<vals.length;i++)if(backlogNorm_(vals[i][2])===nk&&backlogSizeNorm_(vals[i][3])===sk){hit=i;break;}}
  if(hit>=0){const row=hit+2;const current=Math.max(0,(Number(vals[hit][5])||0)+Number(delta||0));sh.getRange(row,2).setValue(pid||vals[hit][1]);sh.getRange(row,6).setValue(current);sh.getRange(row,8).setValue(new Date());}
  else if(Number(delta)>0){sh.appendRow(['OB-'+Utilities.getUuid().slice(0,8).toUpperCase(),pid,name||'',size||'',0,Number(delta), 'アプリ発注',new Date()]);}
}

function ensureSetup_() {
  const ss=SpreadsheetApp.getActive();
  if(!ss.getSheetByName(APP.SHEETS.products)) setupDatabase();
  ensureProductSchema_();
  ensureLegacyOutflowSheet_();
  ensureBacklogSheet_();
  applyInitialUncountedFlags_();
}

function ensureLegacyOutflowSheet_() {
  const ss=SpreadsheetApp.getActive();
  let sh=ss.getSheetByName(APP.SHEETS.legacyOutflow);
  if(!sh) sh=ss.insertSheet(APP.SHEETS.legacyOutflow);
  if(sh.getLastRow()===0) sh.getRange(1,1,1,APP.HEADERS.legacyOutflow.length).setValues([APP.HEADERS.legacyOutflow]);
  return sh;
}

// 既存DBにも新しい棚卸状態列を安全に追加
function ensureProductSchema_() {
  const ss=SpreadsheetApp.getActive();
  const specs=[['商品マスタ','通貨'],['入出庫履歴','通貨'],['発注ヘッダ','通貨'],['発注明細','通貨']];
  specs.forEach(([sheet,header])=>{
    const sh=ss.getSheetByName(sheet); if(!sh)return;
    const headers=sh.getRange(1,1,1,Math.max(1,sh.getLastColumn())).getValues()[0].map(String);
    if(!headers.includes(header)) sh.getRange(1,sh.getLastColumn()+1).setValue(header);
  });
  const sh=ss.getSheetByName(APP.SHEETS.products); if(!sh)return;
  if(!String(sh.getRange(1,10).getValue())) sh.getRange(1,10).setValue('棚卸状態');
  const headers=sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0].map(String);
  const c=headers.indexOf('通貨')+1;
  if(c) { const n=Math.max(0,sh.getLastRow()-1); if(n) sh.getRange(2,c,n,1).getValues().forEach((r,i)=>{if(!r[0])sh.getRange(i+2,c).setValue('JPY')}); }
}

// 初期データで「0」と入っていたが、実際には未棚卸と確認された商品。
// 過去の棚卸履歴は削除せず、次回実棚を入力するまで在庫計算から除外する。
function applyInitialUncountedFlags_() {
  const ids=['P00049','P00149','P00150','P00163','P00164','P00227','P00372','P00373','P00429','P00487','P00508','P00514'];
  const sh=SpreadsheetApp.getActive().getSheetByName(APP.SHEETS.products);
  if(!sh || sh.getLastRow()<2) return;
  const values=sh.getRange(2,1,sh.getLastRow()-1,10).getValues();
  const idSet=new Set(ids);
  values.forEach((r,i)=>{
    if(idSet.has(String(r[0])) && !String(r[9]||'')) sh.getRange(i+2,10).setValue('未棚卸');
  });
}
function normalizeCurrency_(v){const s=String(v||'').trim().toUpperCase();return (s==='USD'||s==='$'||s==='ドル'||s==='US DOLLAR')?'USD':'JPY';}
function formatMoney_(n,c){return normalizeCurrency_(c)==='USD'?'$'+Number(n||0).toLocaleString('en-US',{minimumFractionDigits:2,maximumFractionDigits:2}):'¥'+Math.round(Number(n)||0).toLocaleString('ja-JP');}
function orderCurrency_(lines){const cs=[...new Set((lines||[]).map(x=>normalizeCurrency_(x.currency)))];return cs.length===1?cs[0]:'MIXED';}
function repairCurrencyAndGaramProducts20261119(){
  ensureSetup_(); const sh=SpreadsheetApp.getActive().getSheetByName(APP.SHEETS.products); const v=sh.getDataRange().getValues(), h=v[0].map(String), c={};h.forEach((x,i)=>c[x]=i+1);
  const variants=['1.黒×金蜻蛉','2.黒×銀蜻蛉','3.黒×黒蜻蛉','4.黒×青蜻蛉','5.黒×黄蜻蛉','6.黒×赤蜻蛉','7.黒×赤青蜻蛉','8.黒×黒赤蜻蛉','9.赤×白蜻蛉','10.黒×金小桜','11.黒×銀小桜','12.黒×黒小桜','13.黒×青小桜','14.黒×紫小桜','15.黒×赤小桜','16.黒×黒赤小桜','17.赤×白小桜','18.黒×青瓢箪','19.黒×赤瓢箪','20.黒×黒瓢箪','21.黒×金波','22.黒×銀波','23.黒×黒波'];
  const rows={};v.slice(1).forEach((r,i)=>rows[String(r[c['商品ID']-1])]=i+2); let count=0;
  variants.forEach((size,i)=>{const id='P'+String(461+i).padStart(5,'0'), row=rows[id]; if(!row)return; sh.getRange(row,c['商品名'],1,2).setValues([['印伝風 鍔止',size]]); sh.getRange(row,c['仕入先']).setValue('ガラムスポーツ'); sh.getRange(row,c['単価']).setValue(2.30); sh.getRange(row,c['有効']).setValue(true); if(c['通貨'])sh.getRange(row,c['通貨']).setValue('USD'); count++;});
  log_('ドル通貨・ガラム表記修正','P00461-P00483',`印伝風 鍔止 ${count}件 / USD $2.30 / ガラムスポーツ`); return `修正完了：印伝風 鍔止 ${count}件をガラムスポーツ・USD $2.30に設定しました。`;
}
function readRows_(sh) {
  if(!sh || sh.getLastRow()<2)return [];
  const v=sh.getDataRange().getValues(), h=v[0];
  return v.slice(1).filter(r=>r.some(x=>x!==''&&x!==null)).map(r=>{
    const o={};h.forEach((k,i)=>o[k]=r[i]);return o;
  });
}
function log_(action,id,detail) {
  const sh=SpreadsheetApp.getActive().getSheetByName(APP.SHEETS.audit);
  if(sh) sh.appendRow([new Date(),user_(),action,id,detail]);
}
function makeId_(prefix){return prefix+'-'+Utilities.getUuid().slice(0,8).toUpperCase();}
function user_(){try{return Session.getActiveUser().getEmail()||'不明';}catch(e){return '不明';}}
function today_(){return Utilities.formatDate(new Date(),Session.getScriptTimeZone(),'yyyy-MM-dd');}
function dateKey_(v){
  if(v instanceof Date)return Utilities.formatDate(v,Session.getScriptTimeZone(),'yyyy-MM-dd');
  const s=String(v||'').trim(); if(/^\d{4}-\d{2}-\d{2}$/.test(s))return s;
  const d=new Date(s); if(!isNaN(d.getTime()))return Utilities.formatDate(d,Session.getScriptTimeZone(),'yyyy-MM-dd');
  return s;
}
function shiftDate_(s,days){const d=new Date(s+'T00:00:00');d.setDate(d.getDate()+days);return Utilities.formatDate(d,Session.getScriptTimeZone(),'yyyy-MM-dd');}
function validate_(v,msg){if(!v)throw new Error(msg);}


/**
 * 元Excel「2024 防具2」の商品見出しに記載された単価を商品マスタへ反映。
 * 単価が0または空欄の商品だけを更新し、入力済み単価は上書きしません。
 * 棚卸履歴・入出庫履歴・補充ライン・仕入先は変更しません。
 * 既存DBで1回だけ実行してください。
 */
function repairUnitPricesFromOriginalExcel() {
  ensureSetup_();
  const priceMap = {"燕 面":7450,"燕 小手":4177,"燕 垂":3571,"鷹1T 面 濃紺":8405,"鷹1T 面 エンジ":8405,"鷹1T 小手(赤姫)":4498,"鷹1T 垂 濃紺":4522,"鷹1T 垂 エンジ(赤姫)":4522,"共用胴 濃紺鬼雲":6137,"共用胴 エンジ鬼雲":6137,"共用胴 濃紺ベタ":6137,"共用胴 エンジベタ":6137,"鷹2T 面 濃紺":10497,"鷹2T 面 エンジ":10497,"鷹2T 小手 クラレ":5141,"鷹2T 垂 濃紺":5366,"鷹2T 垂 エンジ":5366,"無双 小手":4884,"赤姫 面":9067,"赤姫 黒胴":8630,"赤姫 玉虫胴":11053,"疾風 面":8405,"疾風 小手":4498,"疾風 垂":4522,"黒鮫 胴":12500,"7mm 面":6962,"7mm 小手":3213,"7mm 垂":2892,"(胸)鬼雲 碁盤刺＃29":4605,"(胸)兜 碁盤刺＃29":4605,"(胸)山飾 碁盤刺＃29":4605};
  const sh = SpreadsheetApp.getActive().getSheetByName(APP.SHEETS.products);
  const values = sh.getDataRange().getValues();
  let updated = 0, alreadySet = 0, notInSource = 0;
  const norm = s => String(s || '').replace(/[\u3000\s]+/g, ' ').trim();
  values.slice(1).forEach((r, i) => {
    const name = norm(r[1]);
    const sources = String(r[7] || '');
    if (!sources.includes('2024 防具2') || !Object.prototype.hasOwnProperty.call(priceMap, name)) return;
    const current = Number(r[5]) || 0;
    if (current > 0) { alreadySet++; return; }
    sh.getRange(i + 2, 6).setValue(priceMap[name]);
    updated++;
  });
  log_('元Excel単価修正','2024 防具2',`単価更新${updated}件 / 既存単価保持${alreadySet}件`);
  return `単価反映完了：${updated}件を更新、${alreadySet}件は既存単価を保持しました。補充ライン・仕入先・棚卸履歴・入出庫履歴は変更していません。`;
}


/**
 * 元のメーカー発注書テンプレートで確認できた単価と仕入先を一括反映します。
 * 単価0の商品だけ更新し、既存単価・棚卸・入出庫・補充ラインは変更しません。
 * 1回だけ実行してください。
 */
function repairKnownPricesAndSuppliers20261119() {
  ensureSetup_();
  const sh=SpreadsheetApp.getActive().getSheetByName(APP.SHEETS.products);
  const values=sh.getDataRange().getValues();
  const norm=v=>String(v||'').replace(/[\u3000\s]+/g,' ').trim();
  const priceByKey={};
  const add=(name,size,price)=>priceByKey[norm(name)+'|'+norm(size)]=price;
  // 三信物産「2新価格」
  [['隼1 面','小小',7200],['隼1 面','小',7200],['隼1 面','中',7300],['隼1 面','大',7300],
   ['隼1 胴','小小',5100],['隼1 胴','小',5100],['隼1 胴','中',5200],['隼1 胴','大',5200],
   ['隼1 小手','小小',2800],['隼1 小手','小',2800],['隼1 小手','中',2900],['隼1 小手','大',2900],
   ['隼1 垂','小小',2900],['隼1 垂','小',2900],['隼1 垂','中',3000],['隼1 垂','大',3000],
   ['隼1 赤胴','小小',5400],['隼1 赤胴','小',5500],['隼1 赤胴','中',5500],
   ['鷹1 面','中 濃紺',8680],['鷹1 面','中 エンジ',8680],['鷹1 面','大 濃紺',8680],['鷹1 面','大 エンジ',8680],['鷹1 面','大大 濃紺',9110],['鷹1 面','大大 エンジ',9110],
   ['鷹1 小手','中',5120],['鷹1 小手','大',5120],['鷹1 小手','大大',5220],
   ['鷹1 垂','中 濃紺',3880],['鷹1 垂','大 濃紺',3880],['鷹1 垂','大大 濃紺',4000],
   ['鶴1 面','小',7800],['鶴1 面','中',7900],['鶴1 面','大',7900],
   ['鶴1 小手(薙刀白)','小',4800],['鶴1 小手(薙刀白)','中',4900],['鶴1 小手(薙刀白)','大',4900],
   ['鶴2 面','中',8200],['鶴2 面','大',8200],
   ['鶴2 小手(薙刀紺)','小',4800],['鶴2 小手(薙刀紺)','中',4900],['鶴2 小手(薙刀紺)','大',4900],
   ['隼1T 面(内輪固定式)','65~66cm',7200],['隼1T 面(内輪固定式)','67~68cm',7200],['隼1T 面(内輪固定式)','69~70cm',7200],['隼1T 面(内輪固定式)','71~72cm',7400]]
   .forEach(x=>add(x[0],x[1],x[2]));
  const supplierByPrefix=[
    {test:/^(隼1 |鷹1 |究極 |鶴1 |鶴2 |学販小手|隼1T )/,supplier:'三信物産'},
    {test:/^(燕 |鷹1T )/,supplier:'三栄武道具'}
  ];
  const prices=values.slice(1).map(r=>[Number(r[5])||0]);
  const suppliers=values.slice(1).map(r=>[String(r[3]||'')]);
  let updated=0,kept=0,supplierSet=0;
  values.slice(1).forEach((r,i)=>{
    const name=norm(r[1]), size=norm(r[2]), key=name+'|'+size;
    if(Object.prototype.hasOwnProperty.call(priceByKey,key)) {
      if((Number(r[5])||0)===0){prices[i][0]=priceByKey[key];updated++;} else kept++;
    }
    if(!String(r[3]||'').trim()) {
      const rule=supplierByPrefix.find(x=>x.test.test(name));
      if(rule){suppliers[i][0]=rule.supplier;supplierSet++;}
    }
  });
  if(values.length>1){
    sh.getRange(2,6,prices.length,1).setValues(prices);
    sh.getRange(2,4,suppliers.length,1).setValues(suppliers);
  }
  log_('メーカー単価・仕入先補完','2024 防具1 / 発注書テンプレート',`単価更新${updated}件 / 既存単価保持${kept}件 / 仕入先設定${supplierSet}件`);
  return `完了：単価${updated}件更新（既存単価${kept}件保持）、仕入先${supplierSet}件設定。棚卸・入出庫履歴は変更していません。`;
}


/** ガラムジャージの表記修正。商品ID・棚卸・入出庫履歴は変更しない */
function repairGaramVariantLabels20261119() {
  ensureSetup_();
  const sh=SpreadsheetApp.getActive().getSheetByName(APP.SHEETS.products);
  const rows=readRows_(sh);
  const targets={'P00345':'紺一剣','P00346':'白一剣'};
  let changed=0;
  rows.forEach(r=>{const id=String(r['商品ID']||'');if(targets[id] && String(r['サイズ']||'')!==targets[id]){
    const row=Number(r.__row);
    if(row) sh.getRange(row,3).setValue(targets[id]); else {
      const vals=sh.getDataRange().getValues(), head=vals[0], c=head.indexOf('サイズ')+1, idc=head.indexOf('商品ID')+1;
      for(let i=1;i<vals.length;i++) if(String(vals[i][idc-1])===id){sh.getRange(i+1,c).setValue(targets[id]);break;}
    } changed++;
  }});
  log_('ガラムジャージ表記修正','P00345/P00346',`変更${changed}件`);
  return `修正完了：紺一重→紺一剣、白一重→白一剣（${changed}件）。履歴・商品IDは保持しました。`;
}


/**
 * ガラムスポーツ発注書原本の商品を商品マスタへ反映する。
 * 既存商品ID・棚卸履歴・入出庫履歴・発注履歴は保持。
 * 既存商品の仕入先・表示名を補正し、紺一剣/白一剣の不足サイズのみ追加。
 * 同じ商品名・サイズは二重登録しないため、複数回実行しても安全。
 */
function repairGaramAllProducts20261119() {
  ensureSetup_();
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(APP.SHEETS.products);
  const values = sh.getDataRange().getValues();
  const headers = values[0].map(String);
  const col = {};
  headers.forEach((h, i) => col[h] = i + 1);
  const required = ['商品ID','商品名','サイズ','仕入先','補充ライン','単価','有効','元シート','メモ'];
  required.forEach(h => { if (!col[h]) throw new Error('商品マスタに必要な列がありません：' + h); });

  const rowById = {};
  const keyToRows = {};
  let nextId = 1;
  values.slice(1).forEach((r, i) => {
    const id = String(r[col['商品ID'] - 1] || '');
    if (id) rowById[id] = i + 2;
    const m = id.match(/^P(\d+)$/);
    if (m) nextId = Math.max(nextId, Number(m[1]) + 1);
    const key = String(r[col['商品名'] - 1] || '').trim() + '|' + String(r[col['サイズ'] - 1] || '').trim();
    if (!keyToRows[key]) keyToRows[key] = [];
    keyToRows[key].push(i + 2);
  });

  let assigned = 0, renamed = 0, added = 0;
  function updateExisting(id, newName) {
    const row = rowById[id];
    if (!row) return;
    const oldName = String(sh.getRange(row, col['商品名']).getValue() || '');
    if (newName && oldName !== newName) {
      sh.getRange(row, col['商品名']).setValue(newName);
      renamed++;
    }
    const supplierCell = sh.getRange(row, col['仕入先']);
    if (supplierCell.getValue() !== 'ガラムスポーツ') {
      supplierCell.setValue('ガラムスポーツ');
      assigned++;
    }
    if (col['有効']) sh.getRange(row, col['有効']).setValue(true);
  }

  // 既存のガラムジャージ紺・白はIDを維持したまま発注書の正式名称へ変更
  for (let n = 319; n <= 331; n++) updateExisting('P' + String(n).padStart(5,'0'), '織刺ジャージ剣道着（紺）');
  for (let n = 332; n <= 344; n++) updateExisting('P' + String(n).padStart(5,'0'), '織刺ジャージ剣道着（白）');
  // 追加仕様のジャージ関連バリエーション
  for (let n = 345; n <= 349; n++) updateExisting('P' + String(n).padStart(5,'0'), null);
  // ジャージ袴 16～28号
  for (let n = 374; n <= 386; n++) updateExisting('P' + String(n).padStart(5,'0'), 'ジャージ袴');
  // テトロン袴 紺・黒・白
  for (let n = 387; n <= 399; n++) updateExisting('P' + String(n).padStart(5,'0'), 'テトロン袴 紺');
  for (let n = 400; n <= 412; n++) updateExisting('P' + String(n).padStart(5,'0'), 'テトロン袴 黒');
  for (let n = 413; n <= 425; n++) updateExisting('P' + String(n).padStart(5,'0'), 'テトロン袴 白');
  // 印伝風鍔止 23柄
  for (let n = 461; n <= 483; n++) updateExisting('P' + String(n).padStart(5,'0'), '印伝風 鍔止');
  // 波千鳥胴 2色・各4サイズ
  for (let n = 496; n <= 499; n++) updateExisting('P' + String(n).padStart(5,'0'), '波千鳥胴 エンジ 溜石目');
  for (let n = 500; n <= 503; n++) updateExisting('P' + String(n).padStart(5,'0'), '波千鳥胴 古代紫 黒石目');
  // V-1小手
  for (let n = 517; n <= 519; n++) updateExisting('P' + String(n).padStart(5,'0'), 'V-1 小手');

  const sizes = ['000号','00号','0号','0.5号','1号','1.5号','2号','2.5号','3号','3.5号','4号','4.5号','5号'];
  const newGroups = [
    ['紺一剣(並)白片布付', sizes],
    ['白一剣(並)白片布付', sizes]
  ];
  const append = [];
  const known = new Set(Object.keys(keyToRows));
  newGroups.forEach(group => group[1].forEach(size => {
    const key = group[0] + '|' + size;
    if (known.has(key)) return;
    const row = Array(headers.length).fill('');
    row[col['商品ID'] - 1] = 'P' + String(nextId++).padStart(5,'0');
    row[col['商品名'] - 1] = group[0];
    row[col['サイズ'] - 1] = size;
    row[col['仕入先'] - 1] = 'ガラムスポーツ';
    row[col['補充ライン'] - 1] = 0;
    row[col['単価'] - 1] = 0; // 原本の各サイズ単価欄は空欄のため推測しない
    row[col['有効'] - 1] = true;
    row[col['元シート'] - 1] = 'ガラムスポーツ発注書原本';
    row[col['メモ'] - 1] = '2026-11-19 発注書原本より追加（単価未設定）';
    append.push(row);
    known.add(key);
  }));
  if (append.length) {
    sh.getRange(sh.getLastRow() + 1, 1, append.length, headers.length).setValues(append);
    added = append.length;
  }
  log_('ガラムスポーツ全商品登録','GARAM',`仕入先設定${assigned}件 / 商品名補正${renamed}件 / 新規追加${added}件`);
  return `ガラムスポーツ商品登録完了：既存仕入先設定${assigned}件、正式名称補正${renamed}件、新規追加${added}件。新規追加商品の単価は原本の各サイズ欄が空欄のため0円（未設定）です。商品ID・棚卸履歴・入出庫履歴・発注履歴は保持しています。`;
}


/** メーカー不明商品の手動メーカー紐づけ。商品ID・単価・在庫・各履歴は変更しない。 */
function setProductSuppliersManually(data) {
  if (!data || !Array.isArray(data.productIds) || !data.productIds.length) throw new Error('紐づける商品を選択してください');
  const supplier = String(data.supplier || '').trim();
  if (!supplier) throw new Error('メーカー名を入力または選択してください');
  if (supplier.length > 100) throw new Error('メーカー名は100文字以内で入力してください');
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const sh = SpreadsheetApp.getActive().getSheetByName(APP.SHEETS.products);
    if (!sh) throw new Error('商品マスタが見つかりません');
    const values = sh.getDataRange().getValues();
    const headers = values[0].map(String);
    const idCol = headers.indexOf('商品ID');
    const supplierCol = headers.indexOf('仕入先');
    if (idCol < 0 || supplierCol < 0) throw new Error('商品マスタの列構成が正しくありません');
    const targets = new Set(data.productIds.map(x => String(x)));
    let updated = 0;
    const names = [];
    for (let i = 1; i < values.length; i++) {
      const id = String(values[i][idCol] || '');
      if (!targets.has(id)) continue;
      if (String(values[i][supplierCol] || '').trim() !== supplier) {
        sh.getRange(i + 1, supplierCol + 1).setValue(supplier);
        updated++;
      }
      names.push(`${values[i][headers.indexOf('商品名')] || id} ${values[i][headers.indexOf('サイズ')] || ''}`.trim());
    }
    if (!names.length) throw new Error('選択した商品IDが商品マスタに見つかりません');
    log_('メーカー手動紐づけ', Array.from(targets).join(','), `${supplier} / 対象${names.length}件 / 変更${updated}件`);
    return { ok:true, supplier, count:names.length, updated, products:getProducts_() };
  } finally {
    lock.releaseLock();
  }
}
