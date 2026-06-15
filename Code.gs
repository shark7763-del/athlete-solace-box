/*** 選手解憂信箱 ｜ 後端 Apps Script ***
 * 部署方式（不用手動找試算表 ID）：
 * 1. 先執行一次 setup()，它會「自動建立」一個新試算表、欄位，並把網址印在執行記錄。
 *    （第一次會要求授權：選帳號 →「進階」→「前往（不安全）」→ 允許）
 * 2. 部署 → 新增部署作業 → 類型「網頁應用程式」：
 *      執行身分：我自己　／　誰可以存取：「任何人」
 *    取得 /exec 網址，貼回 index.html 最上方的 GAS_URL。
 * 3. 之後改程式碼要「管理部署 → 編輯 → 版本：新版本」才會生效。
 *
 * 進階：若想用「現成的某個試算表」，把它的 ID 填進下方 SHEET_ID 即可；
 *       留空則自動建立並記住（存在 Script Properties）。
 ************************************************/

const SHEET_ID = "";   // 留空＝自動建立；要指定現成試算表才填 ID
const SHEET_NAME = "Records";
const PROP_KEY = "SOLACE_SHEET_ID";

const HEADERS = [
  "id","ts","date","name","team","mood",
  "confidence","focus","motivation","anxiety","sleep","emotion","giveup",
  "stress","q1","q2","q3","wantReply","risk","status","reply","replyTs"
];

/* 取得（或自動建立）試算表，並記住其 ID */
function getSpreadsheet_(){
  const props = PropertiesService.getScriptProperties();
  // 1) 程式碼有指定 ID → 用它
  if(SHEET_ID){ return SpreadsheetApp.openById(SHEET_ID); }
  // 2) 之前已自動建立過 → 用記住的 ID
  const saved = props.getProperty(PROP_KEY);
  if(saved){
    try{ return SpreadsheetApp.openById(saved); }catch(e){ /* 被刪了就重建 */ }
  }
  // 3) 都沒有 → 自動建立一個新試算表並記住
  const ss = SpreadsheetApp.create("選手解憂信箱_資料");
  props.setProperty(PROP_KEY, ss.getId());
  return ss;
}

/* 初始化：建立工作表 + 標題列（手動執行一次） */
function setup(){
  const ss = getSpreadsheet_();
  let sh = ss.getSheetByName(SHEET_NAME);
  if(!sh) sh = ss.insertSheet(SHEET_NAME);
  sh.clear();
  sh.getRange(1,1,1,HEADERS.length).setValues([HEADERS]).setFontWeight("bold");
  sh.setFrozenRows(1);
  // 整欄鎖成純文字，避免日期/數字被 Sheets 自動轉型
  sh.getRange(1,1,sh.getMaxRows(),HEADERS.length).setNumberFormat("@");
  // 移除多餘的預設「工作表1」
  const def = ss.getSheetByName("工作表1") || ss.getSheetByName("Sheet1");
  if(def && ss.getSheets().length>1) ss.deleteSheet(def);
  const msg = "✅ 完成！資料試算表網址：\n" + ss.getUrl();
  Logger.log(msg);
  return msg;
}

function getSheet_(){
  const ss = getSpreadsheet_();
  let sh = ss.getSheetByName(SHEET_NAME);
  if(!sh){ sh = ss.insertSheet(SHEET_NAME); sh.getRange(1,1,1,HEADERS.length).setValues([HEADERS]).setFontWeight("bold"); sh.setFrozenRows(1); }
  return sh;
}

/* 讀取（GET）：回傳所有紀錄 */
function doGet(e){
  return json_({ ok:true, records: readAll_() });
}

/* 寫入（POST）：用 text/plain 傳 JSON，避免 CORS 預檢 */
function doPost(e){
  try{
    const body = JSON.parse(e.postData.contents || "{}");
    const action = body.action;
    if(action === "create")      return json_({ ok:true, record: createRec_(body.record) });
    if(action === "update")      return json_({ ok:true, updated: updateRec_(body.id, body.fields) });
    if(action === "clear")       return json_({ ok:true, cleared: clearAll_() });
    return json_({ ok:false, error:"未知 action：" + action });
  }catch(err){
    return json_({ ok:false, error: String(err) });
  }
}

/* ---- 資料操作 ---- */
function readAll_(){
  const sh = getSheet_();
  const last = sh.getLastRow();
  if(last < 2) return [];
  const rows = sh.getRange(2,1,last-1,HEADERS.length).getValues();
  return rows.filter(r=>r[0]).map(rowToObj_).reverse(); // 最新在前
}

function createRec_(rec){
  const sh = getSheet_();
  sh.appendRow(objToRow_(rec));
  return rec;
}

function updateRec_(id, fields){
  const sh = getSheet_();
  const last = sh.getLastRow();
  if(last < 2) return false;
  const ids = sh.getRange(2,1,last-1,1).getValues();
  for(let i=0;i<ids.length;i++){
    if(String(ids[i][0]) === String(id)){
      const rowNum = i+2;
      const cur = rowToObj_(sh.getRange(rowNum,1,1,HEADERS.length).getValues()[0]);
      const merged = Object.assign(cur, fields);
      sh.getRange(rowNum,1,1,HEADERS.length).setValues([objToRow_(merged)]);
      return true;
    }
  }
  return false;
}

function clearAll_(){
  const sh = getSheet_();
  const last = sh.getLastRow();
  if(last >= 2) sh.deleteRows(2, last-1);
  return true;
}

/* ---- 轉換：物件 <-> 列 ---- */
function objToRow_(o){
  o = o || {};
  const sc = o.scales || {};
  return [
    o.id||"", o.ts||"", o.date||"", o.name||"", o.team||"", (o.mood===0?0:(o.mood||"")),
    num_(sc.confidence), num_(sc.focus), num_(sc.motivation), num_(sc.anxiety),
    num_(sc.sleep), num_(sc.emotion), num_(sc.giveup),
    (o.stress||[]).join("、"), o.q1||"", o.q2||"", o.q3||"",
    o.wantReply?"是":"否", o.risk?"高風險":"", o.status||"todo", o.reply||"", o.replyTs||""
  ];
}
function rowToObj_(r){
  return {
    id:String(r[0]), ts:r[1], date:r[2], name:r[3], team:r[4], mood:Number(r[5]),
    scales:{ confidence:Number(r[6]), focus:Number(r[7]), motivation:Number(r[8]),
             anxiety:Number(r[9]), sleep:Number(r[10]), emotion:Number(r[11]), giveup:Number(r[12]) },
    stress: r[13] ? String(r[13]).split("、").filter(Boolean) : [],
    q1:r[14], q2:r[15], q3:r[16],
    wantReply: r[17]==="是", risk: r[18]==="高風險",
    status: r[19]||"todo", reply:r[20]||"", replyTs:r[21]||""
  };
}
function num_(v){ return (v===0||v) ? Number(v) : 3; }

function json_(obj){
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
