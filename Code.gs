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
  "id","athleteId","teamId","date","name","team",
  "moodCard","monsterTags","needType","message",
  "riskLevel","wantsReply","visibleToCoachSummary",
  "psychologistReply","status","followUpDate",
  "issueTypes","urgency","summary","keyQuotes","intervention",
  "createdAt","updatedAt","replyTs","demo"
];

// 有新留言時，通知這個信箱（運動心理教練）
const NOTIFY_EMAIL = "relax635@gmail.com";

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

/* 升級欄位：保留既有資料，只更新標題列與欄位格式 */
function upgradeHeaders(){
  const sh = getSheet_();
  sh.getRange(1,1,1,HEADERS.length).setValues([HEADERS]).setFontWeight("bold");
  sh.setFrozenRows(1);
  sh.getRange(1,1,sh.getMaxRows(),HEADERS.length).setNumberFormat("@");
  const msg = "✅ 已更新欄位標題，既有資料已保留。";
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
  const width = Math.max(HEADERS.length, sh.getLastColumn());
  const rows = sh.getRange(2,1,last-1,width).getValues();
  return rows.filter(r=>r[0]).map(rowToObj_).reverse(); // 最新在前
}

function createRec_(rec){
  const sh = getSheet_();
  sh.appendRow(objToRow_(rec));
  notifyCoach_(rec);   // 寄信通知運動心理教練
  return rec;
}

/* 有新留言就寄 email 通知運動心理教練 */
function notifyCoach_(rec){
  try{
    if(!NOTIFY_EMAIL) return;
    rec = rec || {};
    const high = rec.risk ? "（⚠ 高風險，請優先處理）" : "";
    const subject = "【選手解憂信箱】新留言：" + (rec.name || "匿名選手") + high;
    let body = "";
    body += "有一位選手剛剛送出解憂信箱留言：\n\n";
    body += "姓名：" + (rec.name || "") + "\n";
    body += "隊伍：" + (rec.team || "") + "\n";
    body += "日期：" + (rec.date || "") + "\n";
    body += "希望回覆：" + (rec.wantReply ? "是" : "否") + "\n";
    if(rec.risk){
      body += "\n⚠ 系統偵測到可能的高風險語句，請務必優先且立即關心，必要時通知主要教練或轉介專業協助。\n";
    }
    body += "\n--- 留言內容 ---\n";
    body += "最困擾的事：" + (rec.q1 || "（未填）") + "\n";
    body += "希望被理解：" + (rec.q2 || "（未填）") + "\n";
    body += "希望幫忙：" + (rec.q3 || "（未填）") + "\n";
    body += "\n請登入運動心理教練後台查看與回覆。";
    MailApp.sendEmail(NOTIFY_EMAIL, subject, body);
  }catch(err){
    // 通知失敗不影響資料寫入
    Logger.log("notifyCoach_ 失敗：" + err);
  }
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
  return [
    o.id||"", o.athleteId||"", o.teamId||"", o.date||"", o.name||"", o.team||"",
    o.moodCard||"", ((o.monsterTags||o.stress)||[]).join("、"), o.needType||"", o.message||o.q1||"",
    o.riskLevel||(o.risk?"immediate":"normal"), (o.wantsReply||o.wantReply)?"是":"否", o.visibleToCoachSummary||"",
    o.psychologistReply||o.reply||"", o.status||"new", o.followUpDate||"",
    ((o.analysis&&o.analysis.issueTypes)||o.issueTypes||[]).join("、"),
    (o.analysis&&o.analysis.urgency)||o.urgency||"",
    (o.analysis&&o.analysis.summary)||o.summary||"",
    ((o.analysis&&o.analysis.keyQuotes)||o.keyQuotes||[]).join("｜"),
    (o.analysis&&o.analysis.intervention)||o.intervention||"",
    o.createdAt||o.ts||"", o.updatedAt||"", o.replyTs||"", o.demo?"是":""
  ];
}
function rowToObj_(r){
  const oldStatusAt18 = ["todo","done","follow","refer"].indexOf(String(r[18]||"")) !== -1;
  const previousSchema = String(r[1]||"").indexOf("T") >= 0 && (r[6] === 0 || r[6] || r[19] || r[24]);
  if(previousSchema && !oldStatusAt18){
    const analysis = {
      issueTypes: r[19] ? String(r[19]).split("、").filter(Boolean) : [],
      urgency: r[20] || "",
      summary: r[21] || "",
      keyQuotes: r[22] ? String(r[22]).split("｜").filter(Boolean) : [],
      intervention: r[23] || ""
    };
    const stress = r[12] ? String(r[12]).split("、").filter(Boolean) : [];
    return {
      id:String(r[0]), ts:r[1], createdAt:r[1], updatedAt:"", date:r[2], name:r[3], team:r[4],
      mood:Number(r[5]), moodCard:"", scales:{ confidence:Number(r[6]), focus:Number(r[7]), motivation:Number(r[8]),
             anxiety:Number(r[9]), fatigue:Number(r[10]), sleep:Number(r[11]) },
      stress:stress, monsterTags:stress, needType:r[13]||"", q1:r[14], q2:r[15], q3:r[16], message:r[14]||"",
      wantReply: r[17]==="是", wantsReply:r[17]==="是", risk: r[18]==="高風險", riskLevel:r[18]==="高風險"?"immediate":"normal",
      analysis: analysis.urgency ? analysis : null, visibleToCoachSummary:analysis.summary||"",
      status: r[24]||"new", reply:r[25]||"", psychologistReply:r[25]||"", replyTs:r[26]||"", followUpDate:""
    };
  }
  if(r.length > 25 || !oldStatusAt18){
    const analysis = {
      issueTypes: r[16] ? String(r[16]).split("、").filter(Boolean) : [],
      urgency: r[17] || "",
      summary: r[18] || "",
      keyQuotes: r[19] ? String(r[19]).split("｜").filter(Boolean) : [],
      intervention: r[20] || ""
    };
    const moodName = r[6] || "";
    const monsters = r[7] ? String(r[7]).split("、").filter(Boolean) : [];
    return {
      id:String(r[0]), athleteId:r[1]||"", teamId:r[2]||"", date:r[3], name:r[4], team:r[5],
      moodCard:moodName, mood:moodIndex_(moodName), monsterTags:monsters, stress:monsters,
      needType:r[8]||"", message:r[9]||"", q1:r[9]||"", q2:"", q3:"",
      riskLevel:r[10]||"", risk:r[10]==="immediate", wantsReply:r[11]==="是", wantReply:r[11]==="是",
      visibleToCoachSummary:r[12]||"", psychologistReply:r[13]||"", reply:r[13]||"",
      status:r[14]||"new", followUpDate:r[15]||"",
      analysis: analysis.urgency ? analysis : null,
      createdAt:r[21]||"", ts:r[21]||"", updatedAt:r[22]||"", replyTs:r[23]||"", demo:r[24]==="是"
    };
  }
  return {
    id:String(r[0]), ts:r[1], date:r[2], name:r[3], team:r[4], mood:Number(r[5]),
    scales:{ confidence:Number(r[6]), focus:Number(r[7]), motivation:Number(r[8]),
             anxiety:Number(r[9]), fatigue:Number(r[10]), sleep:Number(r[11]) },
    stress: r[12] ? String(r[12]).split("、").filter(Boolean) : [],
    needType:r[13]||"", q1:r[14], q2:r[15], q3:r[16],
    wantReply: r[17]==="是", risk: r[18]==="高風險",
    message:r[14]||"", athleteId:"", teamId:"",
    riskLevel:r[17]==="高風險"?"immediate":"normal",
    analysis: null,
    status: r[18]||"todo", reply:r[19]||"", psychologistReply:r[19]||"", replyTs:r[20]||"", followUpDate:""
  };
}
function num_(v){ return (v===0||v) ? Number(v) : 3; }
function moodIndex_(name){
  const moods = ["火山快爆了","電池快沒電","腦袋很吵","想躲起來","今天還可以","我很想贏但很累","我不知道怎麼講"];
  const idx = moods.indexOf(String(name||""));
  return idx >= 0 ? idx : 4;
}

function json_(obj){
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
