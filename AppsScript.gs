/*************************************************************
 *  BOM 管理系統 — Google Apps Script 後端
 *  綁定於 Google 試算表（工具 → Apps Script）
 *
 *  資料表（沒有會自動建立）：
 *    料號主檔 / BOM主檔 / BOM明細 / 異動紀錄 / 系統 / 使用者 / 庫存異動 / 產品批序號
 *
 *  部署：部署 → 新增部署作業 → 類型「網頁應用程式」
 *        執行身分：我　　存取權：擁有 Google 帳戶的任何人 或 所有人
 *        取得 /exec 網址後填入 index.html 的 API_URL
 *************************************************************/

/* ====== 設定 ====== */
const SHEET_ID = '1bHzcmIVyN8fvIsP9YTc0vRXyMoQEKiDOETTV8QemtEk';   // BOM資料庫 試算表 ID

/* --- 登入系統 --- */
const ADMIN_EMAIL   = 'kalvin7775@gmail.com';   // 用此 Email 註冊會自動成為管理員
const DEFAULT_ROLE  = 'editor';                 // 新註冊者的預設權限：editor 可編輯／viewer 只能看
const TOKEN_SECRET  = 'endosemio-bom-2026-change-me';  // 登入憑證簽章金鑰，建議改成別人猜不到的字串
const TOKEN_DAYS    = 30;                       // 登入後幾天內免重新輸入密碼

/* --- 過渡期共用密碼：新帳號系統上線後仍可用舊密碼登入，但只有「唯讀」權限 --- */
const PASSWORD          = 'endosemio2026';      // 舊的共用密碼
const LEGACY_PW_ENABLED = true;                 // 全部同事都註冊完後，改成 false 即停用
/* ================== */

const SH = {items:'料號主檔', boms:'BOM主檔', lines:'BOM明細', hist:'異動紀錄', sys:'系統',
            users:'使用者', moves:'庫存異動', serial:'產品批序號'};

const HEAD = {
  items:['品號','品名','規格','單位','庫存數量','單位成本','安全庫存','供應商','交期','儲位','MSB','備註'],
  boms :['BOM品號','產品名稱','規格','來源','備註'],
  lines:['BOM品號','項次','階層','品號','品名','規格','用量','單位','備註'],
  hist :['時間','來源','新增','消失','庫存變動','單價變動','其他','操作者'],
  users:['ID','Email','姓名','密碼','角色','狀態','建立時間','最後登入','備註'],
  /* 一張單多列；欄位參考「簡易領料紀錄」，另加單頭欄位方便還原成 Word */
  moves:['單號','日期','類別','倉別','申請部門','異動原因','項次','品號','品名','數量','單位',
         '異動前庫存','異動後庫存','領用人','填表人','備註','單據備註','建立時間'],
  /* 產品批序號記錄表：欄位對齊原本的 Excel「產品批序號記錄表」 */
  serial:['項次','日期','產品名稱','數量','批號/序號','執行者','確認者','用途','備註1']
};

/* 角色權限：admin 全部；editor 可讀可寫；viewer 只能讀 */
const ROLES = {admin:'管理員', editor:'編輯', viewer:'唯讀'};
function canWrite(role){ return role === 'admin' || role === 'editor'; }
function isAdmin(role){ return role === 'admin'; }

var _SS = null, _SH = {}, _SYS = null;
function ss(){
  if(!_SS) _SS = SHEET_ID ? SpreadsheetApp.openById(SHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
  return _SS;
}

/* 首次使用：在編輯器選這個函式按「執行」，建立所有工作表並完成授權 */
function setup(){
  ['items','boms','lines','hist','sys','users','moves','serial'].forEach(sheet);
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
    staff: readStaff(),
    moves: readMoves(),
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

/* ============================================================
   產品批序號記錄表（serial.html 使用）
   讀寫整張表；版本號存在「系統」表的 serialRev，與 BOM 的 rev 互不干擾
   ============================================================ */
/* 批序號有些列沒填「項次」，不能沿用 rows() 以第一欄判斷有無資料，
   改以「批號/序號」或「產品名稱」有值為準，避免整列被濾掉 */
function serialRows(){
  const sh = sheet('serial');
  const last = sh.getLastRow(), lastC = sh.getLastColumn();
  if(last < 2) return [];
  const head = sh.getRange(1,1,1,lastC).getValues()[0].map(String);
  const vals = sh.getRange(2,1,last-1,lastC).getValues();
  return vals.map(function(r){
      const o = {}; head.forEach(function(h,i){ o[h] = r[i]; }); return o;
    }).filter(function(o){
      return String(o['批號/序號']||'').trim() !== '' ||
             String(o['產品名稱']||'').trim() !== '';
    });
}

function readSerial(){
  const list = serialRows().map(function(r){
    return {no:  r['項次']===''||r['項次']===null ? '' : r['項次'],
            date:    fmtCell(r['日期']),
            product: String(r['產品名稱']||''),
            qty: r['數量']===''||r['數量']===null ? '' : r['數量'],
            sn:      String(r['批號/序號']||''),
            doer:    String(r['執行者']||''),
            checker: String(r['確認者']||''),
            use:     String(r['用途']||''),
            note:    String(r['備註1']||'')};
  });
  return {ok:true, records:list,
          rev: Number(sysGet('serialRev',0))||0,
          updatedAt: fmtTS(sysGet('serialAt','')),
          updatedBy: String(sysGet('serialBy',''))};
}

/* 日期欄可能被試算表轉成日期物件，也可能是 20240513 / 202406 這種數字，統一轉字串 */
function fmtCell(v){
  if(v instanceof Date) return Utilities.formatDate(v,'Asia/Taipei','yyyy/MM/dd');
  if(typeof v === 'number') return String(Math.round(v));
  return String(v===null||v===undefined?'':v);
}

function writeSerial(list, who){
  if(!Array.isArray(list)) throw new Error('資料格式不正確，未寫入');
  const cur = serialRows().length;
  if(list.length === 0 && cur > 0)
    throw new Error('收到 0 筆批序號但雲端現有 ' + cur + ' 筆，已拒絕寫入（避免誤清空）');
  const lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try{
    writeRows('serial', list.map(function(r){
      return {'項次': r.no===undefined||r.no===null?'':r.no,
              '日期': r.date||'',
              '產品名稱': r.product||'',
              '數量': r.qty===undefined||r.qty===null?'':r.qty,
              '批號/序號': r.sn||'',
              '執行者': r.doer||'',
              '確認者': r.checker||'',
              '用途': r.use||'',
              '備註1': r.note||''};
    }));
    /* 日期欄一律當文字，避免 202406 被試算表當成數字或日期 */
    const sh = sheet('serial');
    if(list.length) sh.getRange(2,2,list.length,1).setNumberFormat('@');
    const rev = (Number(sysGet('serialRev',0))||0) + 1;
    sysSet('serialRev', rev);
    sysSet('serialAt', Utilities.formatDate(new Date(),'Asia/Taipei','yyyy/MM/dd HH:mm:ss'));
    sysSet('serialBy', who||'');
    sysFlush();
    SpreadsheetApp.flush();
    return rev;
  } finally { lock.releaseLock(); }
}

/* ============================================================
   庫存異動（入庫／領料／退料單）
   單號規則：PP + 西元年後2碼 + 月2碼 + 流水3碼，例 PP2608001
   ============================================================ */
const MOVE_TYPES = ['入庫','領料','退料'];
const MOVE_LIMIT = 3000;        /* 回傳給前端的最近筆數上限 */
const DEPTS      = ['製造部','業務部','研發部','品保部','管理部'];

/* ---------- 人員部門對照（依「領料單系統規範」第二節與工作流程指令的暱稱補充表） ---------- */
const STAFF_DEFAULT = [
  {name:'方信堯', alias:'金剛', dept:'製造部'},
  {name:'李芷羚', alias:'芷羚', dept:'製造部'},
  {name:'林力',   alias:'林力', dept:'製造部'},
  {name:'睿豪',   alias:'睿豪', dept:'製造部'},
  {name:'陳佳葳', alias:'佳葳', dept:'業務部'},
  {name:'羿伶',   alias:'羿伶', dept:'業務部'},
  {name:'陳匡佑', alias:'YOYO', dept:'研發部'},
  {name:'謝嘉元', alias:'嘉元', dept:'研發部'},
  {name:'承儀',   alias:'承儀', dept:'研發部'}
];
function readStaff(){
  try{ const s = sysGet('staff','');
       if(s){ const a = JSON.parse(s); if(Array.isArray(a)) return a; } }catch(e){}
  return STAFF_DEFAULT;
}
function acSetStaff(auth, p){
  if(!isAdmin(auth.role)) return {error:'只有管理員可以維護人員部門對照表'};
  const list = p && p.staff;
  if(!Array.isArray(list)) return {error:'資料格式不正確'};
  const clean = [];
  for(let i=0;i<list.length;i++){
    const n = String(list[i].name||'').trim();
    const d = String(list[i].dept||'').trim();
    if(!n) continue;
    if(DEPTS.indexOf(d) < 0) return {error:'「'+n+'」的部門不在允許清單中：'+DEPTS.join('／')};
    clean.push({name:n, alias:String(list[i].alias||'').trim(), dept:d});
  }
  sysSet('staff', JSON.stringify(clean));
  sysSet('updatedAt', nowTS());
  sysSet('updatedBy', (auth.name||'') + '（更新部門對照表）');
  sysFlush(); SpreadsheetApp.flush();
  return {ok:true, staff:clean};
}

/* 入庫與退料加庫存，領料扣庫存 */
function moveSign(type){ return type === '領料' ? -1 : 1; }

function readMoves(){
  const all = rows('moves');
  const use = all.length > MOVE_LIMIT ? all.slice(all.length - MOVE_LIMIT) : all;
  return use.map(function(r){
    return {no:String(r['單號']||''), date:fmtTS(r['日期']), type:String(r['類別']||''),
      house:String(r['倉別']||''), dept:String(r['申請部門']||''), reason:String(r['異動原因']||''),
      seq:num(r['項次']), code:String(r['品號']||''), name:String(r['品名']||''),
      qty:num(r['數量']), unit:String(r['單位']||'pcs'),
      before:num(r['異動前庫存']), after:num(r['異動後庫存']),
      taker:String(r['領用人']||''), by:String(r['填表人']||''),
      note:String(r['備註']||''), memo:String(r['單據備註']||''), at:fmtTS(r['建立時間'])};
  });
}

/* 送出一張單：配單號 → 加減庫存 → 逐列寫入庫存異動 */
function postOrder(auth, p){
  if(!canWrite(auth.role))
    return {error: auth.legacy ? '共用密碼為唯讀模式，無法送出單據。'
                              : '你的權限為「唯讀」，無法送出單據。', denied:true};
  let o = (p && p.order) || {};
  if(MOVE_TYPES.indexOf(o.type) < 0) return {error:'類別必須是入庫／領料／退料'};
  const lines = (o.lines || []).filter(function(l){
    return String(l.code||'').trim() !== '' && Number(l.qty) > 0; });
  if(!lines.length) return {error:'單據沒有有效明細（品號與數量都要填）'};
  /* 規範：一張單最多 24 個品項（Word 模板固定 24 列，不可多頁） */
  if(lines.length > 24)
    return {error:'一張單最多 24 個品項，目前 '+lines.length+' 項，請拆成多張單'};
  if(!String(o.reason||'').trim()) return {error:'「異動原因」必填，會印在 Word 單的表頭'};
  /* 申請部門沒填就依填表人查人員部門對照表 */
  let dept = String(o.dept||'').trim();
  if(!dept){
    const me = readStaff().find(function(s){
      return s.name === auth.name || s.alias === auth.name; });
    dept = me ? me.dept : '';
  }
  if(!dept) return {error:'找不到「'+(auth.name||'')+'」的申請部門，請先在「使用者管理 → 人員部門對照表」補上'};
  o = Object.assign({}, o, {dept:dept});

  const lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try{
    /* --- 取單號 --- */
    const d = o.date ? new Date(String(o.date).replace(/-/g,'/')) : new Date();
    const valid = isNaN(d.getTime()) ? new Date() : d;
    const pfx = 'PP' + Utilities.formatDate(valid,'Asia/Taipei','yyMM');
    let max = 0;
    rows('moves').forEach(function(r){
      const m = String(r['單號']||'').match(/^PP(\d{4})(\d{3})$/);
      if(m && ('PP'+m[1]) === pfx) max = Math.max(max, parseInt(m[2],10));
    });
    const no = pfx + String(max+1).padStart(3,'0');
    const dateStr = Utilities.formatDate(valid,'Asia/Taipei','yyyy/MM/dd');

    /* --- 讀料號主檔，準備改庫存 --- */
    const ish = sheet('items');
    const last = ish.getLastRow();
    if(last < 2) return {error:'料號主檔沒有資料，無法異動'};
    const head = ish.getRange(1,1,1,ish.getLastColumn()).getValues()[0].map(String);
    const cCode = head.indexOf('品號'), cStock = head.indexOf('庫存數量'),
          cName = head.indexOf('品名'), cUnit = head.indexOf('單位');
    if(cCode < 0 || cStock < 0) return {error:'料號主檔缺少「品號」或「庫存數量」欄'};
    const body = ish.getRange(2,1,last-1,head.length).getValues();
    const idx = {};
    for(let i=0;i<body.length;i++){
      const c = String(body[i][cCode]).trim();
      if(c && !(c in idx)) idx[c] = i;
    }
    const missing = [];
    lines.forEach(function(l){ if(!(String(l.code).trim() in idx)) missing.push(l.code); });
    if(missing.length) return {error:'下列品號不在料號主檔，請先建檔：' + missing.join('、')};

    /* --- 逐列算庫存並組出要寫的資料 --- */
    const sign = moveSign(o.type);
    const now  = nowTS();
    const out  = [], touched = {};
    lines.forEach(function(l,i){
      const code = String(l.code).trim(), r = idx[code];
      const before = touched[code] !== undefined ? touched[code] : num(body[r][cStock]);
      const after  = before + sign * Number(l.qty);
      touched[code] = after;
      out.push([no, dateStr, o.type, o.house||'', o.dept||'製造部', o.reason||'',
                i+1, code, (cName>=0?String(body[r][cName]):(l.name||'')),
                Number(l.qty), (l.unit || (cUnit>=0?String(body[r][cUnit]):'') || 'pcs'),
                before, after, o.taker||'', auth.name||o.by||'', l.note||'', o.memo||'', now]);
    });
    /* --- 寫回庫存 --- */
    Object.keys(touched).forEach(function(code){
      ish.getRange(idx[code]+2, cStock+1).setValue(touched[code]);
    });
    /* --- 附加到庫存異動 --- */
    const msh = sheet('moves');
    msh.getRange(msh.getLastRow()+1, 1, out.length, HEAD.moves.length).setValues(out);

    const rev = (Number(sysGet('rev',0))||0) + 1;
    sysSet('rev', rev);
    sysSet('updatedAt', now);
    sysSet('updatedBy', (auth.name||'') + '（' + o.type + ' ' + no + '）');
    sysFlush();
    SpreadsheetApp.flush();
    return {ok:true, no:no, rev:rev, date:dateStr,
            stock: Object.keys(touched).map(function(c){ return {code:c, stock:touched[c]}; })};
  } finally { lock.releaseLock(); }
}

/* ============================================================
   使用者與登入
   ============================================================ */

/* 讀出使用者，附帶試算表列號以便單列更新 */
function userRows(){
  const sh = sheet('users');
  const last = sh.getLastRow();
  if(last < 2) return [];
  const head = HEAD.users;
  const vals = sh.getRange(2,1,last-1,head.length).getValues();
  const out = [];
  for(let i=0;i<vals.length;i++){
    const o = {_row: i+2};
    head.forEach(function(h,j){ o[h] = vals[i][j]; });
    if(String(o['ID']||'').trim()==='' && String(o['Email']||'').trim()==='') continue;
    out.push(o);
  }
  return out;
}
function userByEmail(email){
  const e = String(email||'').trim().toLowerCase();
  return userRows().find(function(u){ return String(u['Email']).trim().toLowerCase() === e; }) || null;
}
function userById(id){
  const s = String(id||'');
  return userRows().find(function(u){ return String(u['ID']) === s; }) || null;
}
function setUserCell(row, field, value){
  sheet('users').getRange(row, HEAD.users.indexOf(field)+1).setValue(value);
}
function newId(){ return Utilities.getUuid().replace(/-/g,'').slice(0,10); }
function nowTS(){ return Utilities.formatDate(new Date(),'Asia/Taipei','yyyy/MM/dd HH:mm:ss'); }

/* 密碼以 SHA-256 加鹽雜湊存放，試算表裡看不到明碼 */
function hashPw(pw, salt){
  const raw = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,
                String(salt) + '|' + String(pw) + '|' + TOKEN_SECRET, Utilities.Charset.UTF_8);
  return Utilities.base64Encode(raw);
}
function b64u(s){ return Utilities.base64EncodeWebSafe(s).replace(/=+$/,''); }
function sign(payload){
  const raw = Utilities.computeHmacSha256Signature(payload, TOKEN_SECRET);
  return Utilities.base64EncodeWebSafe(raw).replace(/=+$/,'');
}
function makeToken(id){
  const payload = id + '|' + (Date.now() + TOKEN_DAYS*86400000);
  return b64u(payload) + '.' + sign(payload);
}
function readToken(token){
  const parts = String(token||'').split('.');
  if(parts.length !== 2) return null;
  let payload;
  try{ payload = Utilities.newBlob(Utilities.base64DecodeWebSafe(parts[0])).getDataAsString(); }
  catch(e){ return null; }
  if(sign(payload) !== parts[1]) return null;
  const seg = payload.split('|');
  if(Number(seg[1]) < Date.now()) return null;
  return seg[0];
}

/* 驗證請求身分 → {role, name, id, legacy} 或 {error} */
function authOf(body){
  const tk = body && body.token;
  if(tk){
    const id = readToken(tk);
    if(!id) return {error:'登入已過期，請重新登入', needLogin:true};
    const u = userById(id);
    if(!u) return {error:'帳號不存在，請重新登入', needLogin:true};
    const st = String(u['狀態']||'');
    if(st === 'pending')  return {error:'你的帳號尚待管理員審核', needLogin:true};
    if(st !== 'active')   return {error:'此帳號已停用，請聯絡管理員', needLogin:true};
    return {id:String(u['ID']), name:String(u['姓名']||''), email:String(u['Email']||''),
            role:String(u['角色']||'viewer')};
  }
  /* 過渡期：舊共用密碼 → 唯讀 */
  if(LEGACY_PW_ENABLED && body && String(body.pw||'') === PASSWORD){
    return {id:'', name:String(body.who||'共用密碼'), email:'', role:'viewer', legacy:true};
  }
  return {error:'請先登入', needLogin:true};
}
function profileOf(u){
  return {id:String(u['ID']), email:String(u['Email']), name:String(u['姓名']||''),
          role:String(u['角色']||'viewer'), roleName: ROLES[String(u['角色']||'viewer')] || '唯讀',
          status:String(u['狀態']||'')};
}

/* ---------- 註冊（註冊即開通，預設可編輯；ADMIN_EMAIL 自動成為管理員） ---------- */
function acRegister(p){
  const email = String(p.email||'').trim().toLowerCase();
  const name  = String(p.name||'').trim();
  const pw    = String(p.password||'');
  if(!email || !name || !pw) return {error:'姓名、Email 與密碼皆為必填'};
  if(!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email)) return {error:'Email 格式不正確'};
  if(pw.length < 6) return {error:'密碼至少 6 個字元'};
  if(userByEmail(email)) return {error:'此 Email 已註冊過，請直接登入或請管理員重設密碼'};

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try{
    if(userByEmail(email)) return {error:'此 Email 已註冊過'};
    const first  = userRows().length === 0;
    const isBoss = (email === String(ADMIN_EMAIL).toLowerCase()) || first;
    const id = newId();
    sheet('users').appendRow([id, email, name, hashPw(pw, id),
      isBoss ? 'admin' : DEFAULT_ROLE, 'active',
      nowTS(), '', isBoss ? '系統自動設為管理員' : '']);
    SpreadsheetApp.flush();
    return {ok:true, token: makeToken(id), profile: profileOf(userById(id))};
  } finally { lock.releaseLock(); }
}

/* ---------- 登入 ---------- */
function acLogin(p){
  const email = String(p.email||'').trim().toLowerCase();
  const pw    = String(p.password||'');
  if(!email || !pw) return {error:'請輸入 Email 與密碼'};
  const u = userByEmail(email);
  if(!u) return {error:'查無此 Email，請先註冊'};
  if(String(u['密碼']) !== hashPw(pw, String(u['ID']))) return {error:'密碼錯誤'};
  const st = String(u['狀態']||'');
  if(st === 'pending') return {error:'你的帳號尚待管理員審核，核准後才能登入'};
  if(st !== 'active')  return {error:'此帳號已停用，請聯絡管理員'};
  setUserCell(u._row, '最後登入', nowTS());
  return {ok:true, token: makeToken(String(u['ID'])), profile: profileOf(u)};
}

/* ---------- 修改自己的密碼 ---------- */
function acChangePw(auth, p){
  if(auth.legacy || !auth.id) return {error:'共用密碼模式無法修改密碼，請先註冊個人帳號'};
  const u = userById(auth.id);
  if(!u) return {error:'帳號不存在'};
  if(String(u['密碼']) !== hashPw(String(p.oldPassword||''), String(u['ID'])))
    return {error:'原密碼錯誤'};
  const np = String(p.newPassword||'');
  if(np.length < 6) return {error:'新密碼至少 6 個字元'};
  setUserCell(u._row, '密碼', hashPw(np, String(u['ID'])));
  return {ok:true, message:'密碼已更新'};
}

/* ---------- 管理員：使用者清單／審核／角色／停用／重設密碼／刪除 ---------- */
function acListUsers(auth){
  if(!isAdmin(auth.role)) return {error:'只有管理員可以檢視使用者清單'};
  return {ok:true, users: userRows().map(function(u){
    return {id:String(u['ID']), email:String(u['Email']), name:String(u['姓名']||''),
            role:String(u['角色']||'viewer'), status:String(u['狀態']||''),
            createdAt: fmtTS(u['建立時間']), lastLogin: fmtTS(u['最後登入']),
            note:String(u['備註']||'')};
  })};
}
function acSetUser(auth, p){
  if(!isAdmin(auth.role)) return {error:'只有管理員可以變更使用者設定'};
  const u = userById(p.id);
  if(!u) return {error:'查無此使用者'};
  if(p.role !== undefined && p.role !== null && p.role !== ''){
    if(!ROLES[p.role]) return {error:'角色不正確'};
    if(String(u['ID']) === String(auth.id) && p.role !== 'admin')
      return {error:'不能把自己降級，請由另一位管理員操作'};
    setUserCell(u._row, '角色', p.role);
  }
  if(p.status !== undefined && p.status !== null && p.status !== ''){
    if(['pending','active','disabled'].indexOf(p.status) < 0) return {error:'狀態不正確'};
    if(String(u['ID']) === String(auth.id) && p.status !== 'active')
      return {error:'不能停用自己的帳號'};
    setUserCell(u._row, '狀態', p.status);
  }
  if(p.note !== undefined) setUserCell(u._row, '備註', String(p.note||''));
  return {ok:true};
}
function acResetPw(auth, p){
  if(!isAdmin(auth.role)) return {error:'只有管理員可以重設密碼'};
  const u = userById(p.id);
  if(!u) return {error:'查無此使用者'};
  const np = String(p.newPassword||'');
  if(np.length < 6) return {error:'新密碼至少 6 個字元'};
  setUserCell(u._row, '密碼', hashPw(np, String(u['ID'])));
  return {ok:true, message:'已重設密碼，請把新密碼告知該同事'};
}
function acDelUser(auth, p){
  if(!isAdmin(auth.role)) return {error:'只有管理員可以刪除帳號'};
  const u = userById(p.id);
  if(!u) return {error:'查無此使用者'};
  if(String(u['ID']) === String(auth.id)) return {error:'不能刪除自己的帳號'};
  sheet('users').deleteRow(u._row);
  return {ok:true};
}

/* ---------- HTTP ---------- */
function out(obj){
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet(e){
  const p = (e && e.parameter) || {};
  if(p.action === 'ping')
    return out({ok:true, needLogin:true, legacyPw: LEGACY_PW_ENABLED});
  const auth = authOf(p);
  if(auth.error) return out(auth);
  return out(Object.assign(readState(), {auth:{role:auth.role, name:auth.name, legacy:!!auth.legacy}}));
}

function doPost(e){
  let body = {};
  try{ body = JSON.parse(e.postData.contents); }catch(err){ return out({error:'資料格式錯誤'}); }
  const action = body.action;

  /* --- 免登入即可呼叫 --- */
  if(action === 'ping')     return out({ok:true, legacyPw: LEGACY_PW_ENABLED});
  if(action === 'register') return out(acRegister(body));
  if(action === 'login')    return out(acLogin(body));

  /* --- 以下都需要身分 --- */
  const auth = authOf(body);
  if(auth.error) return out(auth);

  if(action === 'me')         return out({ok:true, role:auth.role, name:auth.name,
                                          email:auth.email, legacy:!!auth.legacy,
                                          roleName: ROLES[auth.role]||'唯讀'});
  if(action === 'changePw')   return out(acChangePw(auth, body));
  if(action === 'postOrder')  return out(postOrder(auth, body));
  if(action === 'setStaff')   return out(acSetStaff(auth, body));
  if(action === 'listUsers')  return out(acListUsers(auth));
  if(action === 'setUser')    return out(acSetUser(auth, body));
  if(action === 'resetPw')    return out(acResetPw(auth, body));
  if(action === 'delUser')    return out(acDelUser(auth, body));

  if(action === 'loadSerial') return out(readSerial());

  if(action === 'saveSerial'){
    if(!canWrite(auth.role))
      return out({error: auth.legacy
        ? '共用密碼為唯讀模式，無法儲存。請註冊個人帳號並請管理員給予「編輯」權限。'
        : '你的權限為「唯讀」，無法儲存資料。請聯絡管理員調整權限。', denied:true});
    const curRev = Number(sysGet('serialRev',0))||0;
    if(body.baseRev !== undefined && body.baseRev !== null &&
       Number(body.baseRev) !== curRev && !body.force){
      return out({conflict:true, serverRev:curRev,
                  updatedAt:fmtTS(sysGet('serialAt','')),
                  updatedBy:String(sysGet('serialBy',''))});
    }
    try{
      const rev = writeSerial(body.payload||[], auth.name || body.who || '');
      return out({ok:true, rev:rev, updatedAt:fmtTS(sysGet('serialAt',''))});
    }catch(err){
      return out({error:'寫入失敗：'+err.message});
    }
  }

  if(action === 'load')
    return out(Object.assign(readState(),
      {auth:{id:auth.id, role:auth.role, name:auth.name, email:auth.email, legacy:!!auth.legacy,
             roleName: ROLES[auth.role]||'唯讀'}}));

  if(action === 'save'){
    if(!canWrite(auth.role))
      return out({error: auth.legacy
        ? '共用密碼為唯讀模式，無法儲存。請註冊個人帳號並請管理員給予「編輯」權限。'
        : '你的權限為「唯讀」，無法儲存資料。請聯絡管理員調整權限。', denied:true});
    const st = body.payload || {};
    const cur = Number(sysGet('rev',0))||0;
    if(body.baseRev !== undefined && body.baseRev !== null &&
       Number(body.baseRev) !== cur && !body.force){
      return out({conflict:true, serverRev:cur,
                  updatedAt:fmtTS(sysGet('updatedAt','')), updatedBy:String(sysGet('updatedBy',''))});
    }
    try{
      const rev = writeState(st, auth.name || body.who || '');
      return out({ok:true, rev:rev, updatedAt:fmtTS(sysGet('updatedAt',''))});
    }catch(err){
      return out({error:'寫入失敗：'+err.message});
    }
  }
  return out({error:'未知的指令：'+action});
}
