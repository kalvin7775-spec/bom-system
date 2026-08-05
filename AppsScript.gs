/*************************************************************
 *  BOM 管理系統 — Google Apps Script 後端
 *  綁定於 Google 試算表（工具 → Apps Script）
 *
 *  資料表（沒有會自動建立）：
 *    料號主檔 / BOM主檔 / BOM明細 / 異動紀錄 / 系統
 *
 *  部署：部署 → 新增部署作業 → 類型「網頁應用程式」
 *        執行身分：我　　存取權：擁有 Google 帳戶的任何人 或 所有人
 *        取得 /exec 網址後填入 index.html 的 API_URL
 *************************************************************/

/* ====== 設定 ====== */
const PASSWORD = 'endosemio2026';                                  // 共用密碼，可自行修改
const SHEET_ID = '1bHzcmIVyN8fvIsP9YTc0vRXyMoQEKiDOETTV8QemtEk';   // BOM資料庫 試算表 ID
/* ================== */

const SH = {items:'料號主檔', boms:'BOM主檔', lines:'BOM明細', hist:'異動紀錄', sys:'系統'};

const HEAD = {
  items:['品號','品名','規格','單位','庫存數量','單位成本','安全庫存','供應商','交期','儲位','MSB','備註'],
  boms :['BOM品號','產品名稱','規格','來源','備註'],
  lines:['BOM品號','項次','階層','品號','品名','規格','用量','單位','備註'],
  hist :['時間','來源','新增','消失','庫存變動','單價變動','其他','操作者']
};

var _SS = null, _SH = {}, _SYS = null;
function ss(){
  if(!_SS) _SS = SHEET_ID ? SpreadsheetApp.openById(SHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
  return _SS;
}

/* 首次使用：在編輯器選這個函式按「執行」，建立所有工作表並完成授權 */
function setup(){
  ['items','boms','lines','hist','sys'].forEach(sheet);
  if(String(sysGet('rev',''))==='') { sysSet('rev', 0); sysFlush(); }
  const s = ss();
  const d = s.getSheetByName('工作表1');
  if(d && s.getSheets().length > 1 && d.getLastRow() === 0) s.deleteSheet(d);
  Logger.log('已建立工作表：' + s.getSheets().map(function(x){return x.getName();}).join(', '));
  return '完成';
}

function sheet(key){
  if(_SH[key]) return _SH[key];
  const s = ss();
  let sh = s.getSheetByName(SH[key]);
  if(!sh){
    sh = s.insertSheet(SH[key]);
    if(HEAD[key]){
      sh.getRange(1,1,1,HEAD[key].length).setValues([HEAD[key]])
        .setFontWeight('bold').setBackground('#eef2f7');
      sh.setFrozenRows(1);
    }
  }
  if(HEAD[key] && sh.getLastRow()===0){
    sh.getRange(1,1,1,HEAD[key].length).setValues([HEAD[key]])
      .setFontWeight('bold').setBackground('#eef2f7');
    sh.setFrozenRows(1);
  }
  _SH[key] = sh;
  return sh;
}

function rows(key){
  const sh = sheet(key);
  const last = sh.getLastRow(), lastC = sh.getLastColumn();
  if(last < 2) return [];
  const vals = sh.getRange(2,1,last-1,lastC).getValues();
  const head = sh.getRange(1,1,1,lastC).getValues()[0].map(String);
  return vals.map(r=>{ const o={}; head.forEach((h,i)=>o[h]=r[i]); return o; })
             .filter(o=>String(o[head[0]]||'').trim()!=='');
}

function writeRows(key, arr){
  const sh = sheet(key), head = HEAD[key];
  const last = sh.getLastRow();
  if(last > 1) sh.getRange(2,1,last-1,sh.getLastColumn()).clearContent();
  if(!arr.length) return;
  const data = arr.map(o=>head.map(h=>{ const v=o[h]; return (v===null||v===undefined)?'':v; }));
  sh.getRange(2,1,data.length,head.length).setValues(data);
}

/* ---------- 系統設定（rev / settings）：整張表一次讀寫，避免逐格往返 ---------- */
function sysLoad(){
  if(_SYS) return _SYS;
  const sh = sheet('sys');
  const last = sh.getLastRow();
  _SYS = {keys:[], map:{}};
  if(last >= 1){
    const vals = sh.getRange(1,1,last,2).getValues();
    for(let i=0;i<vals.length;i++){
      const k = String(vals[i][0]||'');
      if(!k) continue;
      _SYS.keys.push(k);
      _SYS.map[k] = vals[i][1];
    }
  }
  return _SYS;
}
function sysGet(k, dflt){
  const c = sysLoad();
  return (k in c.map) ? c.map[k] : dflt;
}
function sysSet(k, v){
  const c = sysLoad();
  if(!(k in c.map)) c.keys.push(k);
  c.map[k] = v;
  c.dirty = true;
}
function sysFlush(){
  const c = _SYS;
  if(!c || !c.dirty) return;
  const sh = sheet('sys');
  const data = c.keys.map(function(k){ return [k, c.map[k]]; });
  if(data.length) sh.getRange(1,1,data.length,2).setValues(data);
  c.dirty = false;
}

/* ---------- 讀取整份資料 ---------- */
function readState(){
  const items = rows('items').map(r=>({
    code:String(r['品號']).trim(), name:String(r['品名']||''), spec:String(r['規格']||''),
    unit:String(r['單位']||'pcs'), stock:num(r['庫存數量']), cost:numOrNull(r['單位成本']),
    safety:numOrNull(r['安全庫存']), supplier:String(r['供應商']||''), lead:String(r['交期']||''),
    loc:String(r['儲位']||''), msb:String(r['MSB']||''), note:String(r['備註']||'')
  }));
  const meta = {};
  rows('boms').forEach(r=>{
    const c = String(r['BOM品號']).trim();
    meta[c] = {name:String(r['產品名稱']||''), spec:String(r['規格']||''),
               source:String(r['來源']||''), note:String(r['備註']||'')};
  });
  const grp = {}, order = [];
  rows('lines').forEach((r,i)=>{
    const bc = String(r['BOM品號']).trim(), code = String(r['品號']).trim();
    if(!bc || !code) return;
    if(!grp[bc]){ grp[bc]=[]; order.push(bc); }
    grp[bc].push({_seq: r['項次']===''?i:num(r['項次']),
      level: Math.max(1, num(r['階層'])||1), code:code,
      name:String(r['品名']||''), spec:String(r['規格']||''),
      qty: r['用量']===''?1:num(r['用量']), unit:String(r['單位']||'pcs'),
      note:String(r['備註']||'')});
  });
  Object.keys(meta).forEach(c=>{ if(order.indexOf(c)<0){ order.push(c); grp[c]=grp[c]||[]; } });
  const boms = order.map(bc=>{
    const ls = grp[bc].sort((a,b)=>a._seq-b._seq).map(function(x){
      return {level:x.level, code:x.code, name:x.name, spec:x.spec, qty:x.qty, unit:x.unit, note:x.note};
    });
    const m = meta[bc] || {};
    return {id:bc, code:bc, name:m.name||'', spec:m.spec||'', source:m.source||'', note:m.note||'', lines:ls};
  });
  const history = rows('hist').map(r=>({
    ts:fmtTS(r['時間']), file:String(r['來源']||''),
    sum:{added:num(r['新增']), removed:num(r['消失']), stock:num(r['庫存變動']),
         cost:num(r['單價變動']), other:num(r['其他'])},
    by:String(r['操作者']||''), diff:{added:[],removed:[],stock:[],cost:[],other:[],
         sum:{added:num(r['新增']),removed:num(r['消失']),stock:num(r['庫存變動']),
              cost:num(r['單價變動']),other:num(r['其他'])}}
  }));
  let settings = {purchaseStop:true};
  try{ const s = sysGet('settings',''); if(s) settings = JSON.parse(s); }catch(e){}
  return {
    rev: Number(sysGet('rev',0)) || 0,
    updatedAt: fmtTS(sysGet('updatedAt','')),
    updatedBy: String(sysGet('updatedBy','')),
    items: items, boms: boms, history: history, settings: settings
  };
}

function num(v){ const n = parseFloat(v); return isNaN(n) ? 0 : n; }
/* 試算表可能把時間字串自動轉成日期物件，讀回時統一格式化 */
function fmtTS(v){
  if(v instanceof Date) return Utilities.formatDate(v,'Asia/Taipei','yyyy/MM/dd HH:mm:ss');
  return String(v===null||v===undefined?'':v);
}
function numOrNull(v){
  if(v===''||v===null||v===undefined) return null;
  const n = parseFloat(v); return isNaN(n) ? null : n;
}

/* ---------- 寫入整份資料 ---------- */
function writeState(st, who){
  /* 安全檢查：資料異常時直接拒絕，避免把整份資料清空 */
  if(!st || !Array.isArray(st.items) || !Array.isArray(st.boms))
    throw new Error('資料格式不正確，未寫入');
  const curItems = rows('items').length;
  if(st.items.length === 0 && curItems > 0)
    throw new Error('收到 0 筆料號但雲端現有 ' + curItems + ' 筆，已拒絕寫入（避免誤清空）');
  const lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try{
    writeRows('items', (st.items||[]).map(i=>({
      '品號':i.code, '品名':i.name||'', '規格':i.spec||'', '單位':i.unit||'pcs',
      '庫存數量':i.stock==null?0:i.stock, '單位成本':i.cost==null?'':i.cost,
      '安全庫存':i.safety==null?'':i.safety, '供應商':i.supplier||'', '交期':i.lead||'',
      '儲位':i.loc||'', 'MSB':i.msb||'', '備註':i.note||''
    })));
    writeRows('boms', (st.boms||[]).map(b=>({
      'BOM品號':b.code, '產品名稱':b.name||'', '規格':b.spec||'',
      '來源':b.source||'', '備註':b.note||''
    })));
    const lines = [];
    (st.boms||[]).forEach(b=>{
      (b.lines||[]).forEach((l,i)=>lines.push({
        'BOM品號':b.code, '項次':i+1, '階層':l.level, '品號':l.code,
        '品名':l.name||'', '規格':l.spec||'', '用量':l.qty, '單位':l.unit||'pcs', '備註':l.note||''
      }));
    });
    writeRows('lines', lines);
    writeRows('hist', (st.history||[]).slice(0,200).map(h=>({
      '時間':h.ts||'', '來源':h.file||'', '新增':(h.sum&&h.sum.added)||0,
      '消失':(h.sum&&h.sum.removed)||0, '庫存變動':(h.sum&&h.sum.stock)||0,
      '單價變動':(h.sum&&h.sum.cost)||0, '其他':(h.sum&&h.sum.other)||0, '操作者':h.by||''
    })));
    if(st.settings) sysSet('settings', JSON.stringify(st.settings));
    const rev = (Number(sysGet('rev',0))||0) + 1;
    sysSet('rev', rev);
    sysSet('updatedAt', Utilities.formatDate(new Date(),'Asia/Taipei','yyyy/MM/dd HH:mm:ss'));
    sysSet('updatedBy', who||'');
    sysFlush();
    SpreadsheetApp.flush();
    return rev;
  } finally { lock.releaseLock(); }
}

/* ---------- HTTP ---------- */
function out(obj){
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet(e){
  const p = (e && e.parameter) || {};
  if(p.action === 'ping') return out({ok:true, needPassword:true});
  if(String(p.pw||'') !== PASSWORD) return out({error:'密碼錯誤'});
  return out(readState());
}

function doPost(e){
  let body = {};
  try{ body = JSON.parse(e.postData.contents); }catch(err){ return out({error:'資料格式錯誤'}); }
  if(String(body.pw||'') !== PASSWORD) return out({error:'密碼錯誤'});

  const action = body.action;
  if(action === 'ping')  return out({ok:true, rev:Number(sysGet('rev',0))||0});
  if(action === 'load')  return out(readState());

  if(action === 'save'){
    const st = body.payload || {};
    const cur = Number(sysGet('rev',0))||0;
    if(body.baseRev !== undefined && body.baseRev !== null &&
       Number(body.baseRev) !== cur && !body.force){
      return out({conflict:true, serverRev:cur,
                  updatedAt:fmtTS(sysGet('updatedAt','')), updatedBy:String(sysGet('updatedBy',''))});
    }
    try{
      const rev = writeState(st, body.who||'');
      return out({ok:true, rev:rev, updatedAt:fmtTS(sysGet('updatedAt',''))});
    }catch(err){
      return out({error:'寫入失敗：'+err.message});
    }
  }
  return out({error:'未知的指令：'+action});
}
