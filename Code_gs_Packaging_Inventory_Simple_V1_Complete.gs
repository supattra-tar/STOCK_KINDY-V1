const DB_NAME = 'Packaging Inventory Simple Database';
const SHEETS = {
  Products: ['ProductCode','ProductName','Category','Description','Unit','ImageURL','MinimumStock','Location','Status','CreatedAt','UpdatedAt'],
  Categories: ['CategoryCode','CategoryName','Description','Status'],
  Receiving: ['ReceiveID','Date','ProductCode','ProductName','Quantity','Unit','ReceivedBy','Reference','Note','CreatedAt'],
  Issue: ['IssueID','Date','ProductCode','ProductName','Quantity','Unit','IssuedBy','Department','Purpose','CreatedAt','Reference'],
  Returns: ['ReturnID','Date','ProductCode','ProductName','Quantity','Unit','ReturnedBy','Reason','Condition','Status','CreatedAt'],
  Return_Stock: ['ReturnID','ProductCode','ProductName','Quantity','Unit','ReturnedBy','Condition','Status','Date','UpdatedAt'],
  Stock_Balance: ['ProductCode','ProductName','Unit','TotalReceived','TotalIssued','AvailableStock','ReturnStock','HoldStock','MinimumStock','StockStatus','UpdatedAt'],
  Stock_Movement: ['MovementID','Date','Type','DocumentID','ProductCode','ProductName','Quantity','Unit','Person','Note','CreatedAt'],
  Daily_Report: ['Date','IssueID','ProductCode','ProductName','Quantity','Unit','IssuedBy','Department','Purpose'],
  Settings: ['Key','Value']
};
const DEFAULT_CATEGORIES = [
  ['CAT001','กล่อง','กล่องบรรจุสินค้า','Active'],['CAT002','ขวด','ขวดบรรจุสินค้า','Active'],
  ['CAT003','ฝาขวด','ฝาปิดบรรจุภัณฑ์','Active'],['CAT004','สติ๊กเกอร์','ฉลากและสติ๊กเกอร์','Active'],
  ['CAT005','วัสดุบรรจุภัณฑ์','วัสดุประกอบบรรจุภัณฑ์','Active'],['CAT006','อุปกรณ์บรรจุ','อุปกรณ์สำหรับบรรจุสินค้า','Active'],
  ['CAT007','อื่นๆ','รายการอื่นๆ','Active']
];

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Packaging Inventory Simple V1.0')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT);
}

/** ใช้ครั้งแรกกับฐานข้อมูลใหม่เท่านั้น; ไม่ใช้ซ้ำเพื่อแก้ schema */
function setupDatabaseSimple() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('กรุณาเปิด Apps Script จาก Google Sheets ฐานข้อมูล');
  PropertiesService.getScriptProperties().setProperty('DATABASE_ID', ss.getId());
  Object.keys(SHEETS).forEach(name => {
    let sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    if (sh.getLastRow() === 0) sh.getRange(1,1,1,SHEETS[name].length).setValues([SHEETS[name]]);
    sh.setFrozenRows(1);
  });
  const cat = ss.getSheetByName('Categories');
  if (cat.getLastRow() === 1) cat.getRange(2,1,DEFAULT_CATEGORIES.length,4).setValues(DEFAULT_CATEGORIES);
  const settings = ss.getSheetByName('Settings');
  if (settings.getLastRow() === 1) settings.getRange(2,1,3,2).setValues([
    ['SystemName','Packaging Inventory Simple V1.0'],['DatabaseName',DB_NAME],['SetupCompleted',new Date()]
  ]);
  ensureIssueReferenceColumn();
  return 'ตรวจสอบฐานข้อมูลเรียบร้อย';
}
function getSS_() {
  const id = PropertiesService.getScriptProperties().getProperty('DATABASE_ID');
  if (id) return SpreadsheetApp.openById(id);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('ไม่พบฐานข้อมูล กรุณาตรวจสอบ Script Property DATABASE_ID');
  return ss;
}
function rows_(name) {
  const sh = getSS_().getSheetByName(name);
  if (!sh) throw new Error('ไม่พบชีต ' + name);
  const v = sh.getDataRange().getValues();
  if (v.length < 2) return [];
  const h = v[0].map(x=>String(x).trim());
  return v.slice(1).filter(r=>r.some(x=>x!=='' && x!==null)).map(r=>{
    const o={}; h.forEach((k,i)=>{if(k)o[k]=r[i];}); return o;
  });
}
function clientSafe_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone() || 'Asia/Bangkok', "yyyy-MM-dd'T'HH:mm:ss");
  if (Array.isArray(v)) return v.map(clientSafe_);
  if (v && typeof v==='object') { const o={}; Object.keys(v).forEach(k=>o[k]=clientSafe_(v[k])); return o; }
  return v;
}
function append_(name,obj) {
  const sh=getSS_().getSheetByName(name);
  if(!sh) throw new Error('ไม่พบชีต '+name);
  const lastCol=sh.getLastColumn();
  if(!lastCol) throw new Error('ชีต '+name+' ยังไม่มีหัวตาราง');
  const h=sh.getRange(1,1,1,lastCol).getValues()[0].map(x=>String(x).trim());
  sh.appendRow(h.map(k=>obj[k]===undefined?'':obj[k]));
}
function id_(p) { return p+'-'+Utilities.formatDate(new Date(),Session.getScriptTimeZone()||'Asia/Bangkok','yyyyMMdd-HHmmss')+'-'+Math.floor(Math.random()*900+100); }
function dateText_(d) {
  if(!d) return Utilities.formatDate(new Date(),Session.getScriptTimeZone()||'Asia/Bangkok','yyyy-MM-dd');
  if(d instanceof Date) return Utilities.formatDate(d,Session.getScriptTimeZone()||'Asia/Bangkok','yyyy-MM-dd');
  const s=String(d); const m=s.match(/^(\d{4}-\d{2}-\d{2})/); return m?m[1]:s.slice(0,10);
}
function num_(v) { const n=Number(v); if(!Number.isFinite(n)||n<=0) throw new Error('จำนวนต้องเป็นตัวเลขมากกว่า 0'); return n; }
function isActive_(v) { return String(v==null?'':v).trim().toLowerCase()==='active'; }
function product_(code) {
  const p=rows_('Products').find(x=>String(x.ProductCode).trim()===String(code||'').trim() && String(x.Status||'').trim().toLowerCase()!=='inactive');
  if(!p) throw new Error('ไม่พบรหัสสินค้า หรือสินค้าไม่ได้เปิดใช้งาน');
  return p;
}
function balanceFor_(code) {
  const same=x=>String(x.ProductCode).trim()===String(code).trim();
  const received=rows_('Receiving').filter(same).reduce((s,x)=>s+Number(x.Quantity||0),0);
  const issued=rows_('Issue').filter(same).reduce((s,x)=>s+Number(x.Quantity||0),0);
  const returns=rows_('Returns').filter(x=>same(x)&&!['Restocked','Rejected'].includes(String(x.Status||''))).reduce((s,x)=>s+Number(x.Quantity||0),0);
  const hold=rows_('Returns').filter(x=>same(x)&&String(x.Status)==='Hold').reduce((s,x)=>s+Number(x.Quantity||0),0);
  return {received,issued,available:received-issued,returns,hold};
}
function refreshBalance_(code) {
  const p=product_(code),b=balanceFor_(code),sh=getSS_().getSheetByName('Stock_Balance'),data=sh.getDataRange().getValues();
  const idx=data.findIndex((r,i)=>i>0&&String(r[0]).trim()===String(code).trim());
  const vals=[code,p.ProductName,p.Unit,b.received,b.issued,b.available,b.returns,b.hold,Number(p.MinimumStock||0),b.available<=Number(p.MinimumStock||0)?'Low':'Normal',new Date()];
  if(idx<0) sh.appendRow(vals); else sh.getRange(idx+1,1,1,vals.length).setValues([vals]);
  return vals;
}
function withLock_(fn) { const l=LockService.getScriptLock(); l.waitLock(20000); try{return fn();}finally{l.releaseLock();} }

function getInitialData() {
  const products=rows_('Products').filter(p=>String(p.Status||'').trim().toLowerCase()!=='inactive');
  const categories=rows_('Categories').filter(c=>isActive_(c.Status));
  const balances=products.map(p=>refreshBalance_(p.ProductCode));
  return clientSafe_({products,categories,balances,dashboard:getDashboard(),movements:getRecentMovements(12),recentIssues:getRecentIssues(10)});
}
function getDashboard() {
  const ps=rows_('Products').filter(p=>String(p.Status||'').trim().toLowerCase()!=='inactive');
  const bs=ps.map(p=>balanceFor_(p.ProductCode)),today=dateText_(new Date());
  const sum=(a,k)=>a.reduce((s,x)=>s+Number(x[k]||0),0);
  return {productCount:ps.length,available:sum(bs,'available'),
    receivedToday:rows_('Receiving').filter(x=>dateText_(x.Date)===today).reduce((s,x)=>s+Number(x.Quantity||0),0),
    issuedToday:rows_('Issue').filter(x=>dateText_(x.Date)===today).reduce((s,x)=>s+Number(x.Quantity||0),0),
    lowStock:ps.filter((p,i)=>bs[i].available<=Number(p.MinimumStock||0)).length,returnQty:sum(bs,'returns')};
}
function getProducts() {
  return clientSafe_(rows_('Products').filter(p=>String(p.Status||'').trim().toLowerCase()!=='inactive').map(p=>({...p,...balanceFor_(p.ProductCode)})));
}
function addProduct(data) {
  return withLock_(()=>{
    const code=String(data.ProductCode||'').trim(),name=String(data.ProductName||'').trim(),category=String(data.Category||'').trim(),unit=String(data.Unit||'').trim();
    if(!code||!name||!category||!unit) throw new Error('กรุณากรอกรหัสสินค้า ชื่อสินค้า หมวดหมู่ และหน่วย');
    if(rows_('Products').some(p=>String(p.ProductCode).trim().toLowerCase()===code.toLowerCase())) throw new Error('รหัสสินค้านี้มีอยู่แล้ว');
    append_('Products',{...data,ProductCode:code,ProductName:name,Category:category,Unit:unit,MinimumStock:Number(data.MinimumStock||0),Status:'Active',CreatedAt:new Date(),UpdatedAt:new Date()});
    refreshBalance_(code); return {success:true,message:'เพิ่มสินค้าเรียบร้อยแล้ว'};
  });
}
function receiveStock(data) {
  return withLock_(()=>{
    const p=product_(data.ProductCode),q=num_(data.Quantity),who=String(data.ReceivedBy||'').trim();
    if(!who) throw new Error('กรุณาระบุชื่อผู้รับสินค้า');
    const doc=id_('RCV'),date=dateText_(data.Date);
    append_('Receiving',{ReceiveID:doc,Date:date,ProductCode:p.ProductCode,ProductName:p.ProductName,Quantity:q,Unit:p.Unit,ReceivedBy:who,Reference:String(data.Reference||'').trim(),Note:data.Note||'',CreatedAt:new Date()});
    append_('Stock_Movement',{MovementID:id_('MOV'),Date:date,Type:'RECEIVE',DocumentID:doc,ProductCode:p.ProductCode,ProductName:p.ProductName,Quantity:q,Unit:p.Unit,Person:who,Note:[data.Reference,data.Note].filter(Boolean).join(' | '),CreatedAt:new Date()});
    refreshBalance_(p.ProductCode); return {success:true,message:'บันทึกรับเข้าเรียบร้อย',documentId:doc};
  });
}
function issueStock(data) {
  return withLock_(()=>{
    const p=product_(data.ProductCode),q=num_(data.Quantity),who=String(data.IssuedBy||'').trim();
    if(!who) throw new Error('กรุณาระบุชื่อผู้เบิก');
    const b=balanceFor_(p.ProductCode);
    if(q>b.available) throw new Error('จำนวนเบิกเกินยอดพร้อมใช้ คงเหลือ '+b.available+' '+p.Unit);
    const doc=id_('ISS'),date=dateText_(data.Date),reference=String(data.Reference||'').trim();
    append_('Issue',{IssueID:doc,Date:date,ProductCode:p.ProductCode,ProductName:p.ProductName,Quantity:q,Unit:p.Unit,IssuedBy:who,Department:data.Department||'',Purpose:data.Purpose||'',CreatedAt:new Date(),Reference:reference});
    append_('Stock_Movement',{MovementID:id_('MOV'),Date:date,Type:'ISSUE',DocumentID:doc,ProductCode:p.ProductCode,ProductName:p.ProductName,Quantity:q,Unit:p.Unit,Person:who,Note:[reference,data.Purpose].filter(Boolean).join(' | '),CreatedAt:new Date()});
    refreshBalance_(p.ProductCode); return {success:true,message:'บันทึกเบิกจ่ายเรียบร้อย',documentId:doc};
  });
}
function receiveReturn(data) {
  return withLock_(()=>{
    const p=product_(data.ProductCode),q=num_(data.Quantity),who=String(data.ReturnedBy||'').trim();
    if(!who) throw new Error('กรุณาระบุชื่อผู้คืน');
    const doc=id_('RET'),date=dateText_(data.Date),condition=data.Condition||'รอตรวจสอบ';
    append_('Returns',{ReturnID:doc,Date:date,ProductCode:p.ProductCode,ProductName:p.ProductName,Quantity:q,Unit:p.Unit,ReturnedBy:who,Reason:data.Reason||'',Condition:condition,Status:'Pending Inspection',CreatedAt:new Date()});
    append_('Return_Stock',{ReturnID:doc,ProductCode:p.ProductCode,ProductName:p.ProductName,Quantity:q,Unit:p.Unit,ReturnedBy:who,Condition:condition,Status:'Pending Inspection',Date:date,UpdatedAt:new Date()});
    refreshBalance_(p.ProductCode); return {success:true,message:'รับคืนเรียบร้อย (แยกจากยอดพร้อมใช้)',documentId:doc};
  });
}
function getDailyIssues(date) { const d=dateText_(date); return clientSafe_(rows_('Issue').filter(x=>dateText_(x.Date)===d)); }
function getRecentIssues(limit) { return clientSafe_(rows_('Issue').sort((a,b)=>String(b.CreatedAt||b.Date).localeCompare(String(a.CreatedAt||a.Date))).slice(0,Number(limit)||10)); }
function getRecentMovements(limit) { return clientSafe_(rows_('Stock_Movement').sort((a,b)=>String(b.CreatedAt||b.Date).localeCompare(String(a.CreatedAt||a.Date))).slice(0,Number(limit)||12)); }
function getReportData() {
  return clientSafe_({products:getProducts(),receiving:rows_('Receiving'),issues:rows_('Issue'),returns:rows_('Returns'),movements:rows_('Stock_Movement')});
}
function exportDailyIssuesCsv(date) {
  const h=getSS_().getSheetByName('Issue').getRange(1,1,1,getSS_().getSheetByName('Issue').getLastColumn()).getDisplayValues()[0].map(x=>String(x).trim());
  const data=getDailyIssues(date),csv=[h,...data.map(o=>h.map(k=>o[k]===undefined?'':o[k]))].map(r=>r.map(v=>'"'+String(v==null?'':v).replace(/"/g,'""')+'"').join(',')).join('\r\n');
  return {filename:'Daily_Issue_'+dateText_(date)+'.csv',content:'\uFEFF'+csv};
}
/** เพิ่มหัวคอลัมน์ Reference ต่อท้ายเท่านั้น ไม่ลบหรือเขียนทับข้อมูลเดิม */
function ensureIssueReferenceColumn() {
  const sh=getSS_().getSheetByName('Issue'); if(!sh) throw new Error('ไม่พบชีต Issue');
  const n=Math.max(sh.getLastColumn(),1),h=sh.getRange(1,1,1,n).getDisplayValues()[0].map(x=>String(x).trim());
  if(h.includes('Reference')) return 'มีคอลัมน์ Reference อยู่แล้ว';
  sh.getRange(1,n+1).setValue('Reference').setFontWeight('bold');
  return 'เพิ่มคอลัมน์ Reference สำเร็จ โดยคงข้อมูลเดิมไว้';
}
