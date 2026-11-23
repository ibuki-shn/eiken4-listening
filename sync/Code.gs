/** @OnlyCurrentDoc */
/*
 * 英検4級リスニングアプリの「記録の写し」を受け取る Google Apps Script。
 *
 * 何をする物か
 *  - アプリ（お母さんのスマホ）から送られてきた記録の写しを、タブ「受信データ」に端末ごと1行で保存する。
 *  - 家族の「見るだけ」表示が ?k=<鍵> で取りに来たら、n（解いた問題数）が一番多い端末の写しを返す。
 *  - タブ「毎日の記録」に、日ごとの解いた問題・正解・正答率を書き直す。
 *
 * デプロイ手順
 *  1. 記録を入れたいスプレッドシートを開き、拡張機能 > Apps Script を開く（シートに紐づいたスクリプト）。
 *  2. このファイルの中身を貼り、下の VIEW_KEY を長いランダムな文字列に変える。
 *  3. デプロイ > 新しいデプロイ > 種類「ウェブアプリ」。実行するユーザー=自分、アクセスできるユーザー=全員。
 *  4. 出てきた URL を、アプリ側の SYNC_URL に入れる。
 *
 * コードを直したとき
 *  - 「デプロイを管理」> 鉛筆 > バージョン「新バージョン」> デプロイ。こうすると URL が変わらない。
 *    （「新しいデプロイ」を選ぶと URL が変わってしまう）
 */

const VIEW_KEY = "ここに見る専用の鍵";
const SHEET_RAW = "受信データ";
const SHEET_DAILY = "毎日の記録";
const CHUNK = 40000;
const MAX_BODY = 3000000;
const MAX_DEVICES = 20;
const FIRST_DATA_COL = 6; // F列

function json_(o){
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function sheet_(name){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  return sh;
}

function rawSheet_(){
  const sh = sheet_(SHEET_RAW);
  if (sh.getLastRow() === 0){
    sh.getRange(1, 1, 1, 6).setValues([["端末id", "受信日時", "端末の時刻", "問題数", "分割数", "JSON（分割）"]]).setFontWeight("bold");
    sh.setFrozenRows(1);
  }
  return sh;
}

function logLen_(data){
  try{
    const a = JSON.parse(data["eiken4.log"]);
    return Array.isArray(a) ? a.length : 0;
  }catch(e){ return 0; }
}

function findRow_(sh, dev){
  const last = sh.getLastRow();
  if (last < 2) return 0;
  const ids = sh.getRange(2, 1, last - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) if (String(ids[i][0]) === dev) return i + 2;
  return 0;
}

function doPost(e){
  const lock = LockService.getScriptLock();
  try{
    lock.waitLock(20000);
  }catch(err){
    return json_({ok:false});
  }
  try{
    const text = (e && e.postData && e.postData.contents) || "";
    if (!text || text.length > MAX_BODY) return json_({ok:false});
    const o = JSON.parse(text);
    const sh = rawSheet_();

    if (o && o.del !== undefined){
      if (o.k !== VIEW_KEY || typeof o.del !== "string") return json_({ok:false});
      const r = findRow_(sh, o.del);
      if (r) sh.deleteRow(r);
      rebuildDaily_();
      return json_({ok:true});
    }

    const dev = o && o.dev;
    if (typeof dev !== "string" || !/^[a-z0-9]{8,40}$/.test(dev)) return json_({ok:false});
    const data = o.data;
    if (!data || typeof data !== "object" || Array.isArray(data)) return json_({ok:false});
    for (const k in data) if (typeof data[k] !== "string") return json_({ok:false});

    let row = findRow_(sh, dev);
    if (!row){
      if (Math.max(sh.getLastRow() - 1, 0) >= MAX_DEVICES) return json_({ok:false});
      row = Math.max(sh.getLastRow(), 1) + 1;
    }
    const s = JSON.stringify(data);
    const parts = [];
    for (let i = 0; i < s.length; i += CHUNK) parts.push("j" + s.substring(i, i + CHUNK));
    const t = Number(o.t);
    // 古い内容を消してから書く（分割数が減ったときに残らないように）
    const width = Math.max(sh.getLastColumn(), FIRST_DATA_COL + parts.length - 1);
    sh.getRange(row, 1, 1, width).clearContent();
    sh.getRange(row, 1, 1, 5).setValues([[dev, new Date(), isFinite(t) ? new Date(t) : "", logLen_(data), parts.length]]);
    sh.getRange(row, 2, 1, 2).setNumberFormat("yyyy/mm/dd hh:mm");
    sh.getRange(row, FIRST_DATA_COL, 1, parts.length).setNumberFormat("@").setValues([parts]);
    rebuildDaily_();
    return json_({ok:true});
  }catch(err){
    return json_({ok:false});
  }finally{
    lock.releaseLock();
  }
}

// シートの日時をミリ秒に（ウェブアプリから読むと instanceof Date が効かないことがある）
function ms_(v){
  if (v && typeof v.getTime === "function") return v.getTime();
  if (typeof v === "number") return v;
  const d = new Date(v);
  return isNaN(d.getTime()) ? 0 : d.getTime();
}

// 受信データの行から、端末ごとの情報を組み立て直す
function readDevices_(){
  const sh = rawSheet_();
  const last = sh.getLastRow();
  if (last < 2) return [];
  const width = Math.max(sh.getLastColumn(), FIRST_DATA_COL);
  const vals = sh.getRange(2, 1, last - 1, width).getValues();
  const out = [];
  vals.forEach(function(r){
    const cnt = Number(r[4]) || 0;
    if (!r[0] || cnt < 1) return;
    let s = "";
    for (let i = 0; i < cnt; i++) s += String(r[FIRST_DATA_COL - 1 + i] || "").substring(1);
    let data = null;
    try{ data = JSON.parse(s); }catch(e){ return; }
    out.push({
      dev: String(r[0]),
      recv: ms_(r[1]),
      t: ms_(r[2]),
      n: Number(r[3]) || 0,
      data: data
    });
  });
  return out;
}

function best_(devs){
  let b = null;
  devs.forEach(function(d){
    if (!b || d.n > b.n || (d.n === b.n && d.recv > b.recv)) b = d;
  });
  return b;
}

function doGet(e){
  try{
    const k = e && e.parameter && e.parameter.k;
    if (k !== VIEW_KEY) return json_({ok:false, err:"key"});
    const b = best_(readDevices_());
    if (!b) return json_({ok:true, empty:true});
    return json_({ok:true, dev:b.dev, updated:b.recv, t:b.t, n:b.n, data:b.data});
  }catch(err){
    return json_({ok:false, err:"server"});
  }
}

// いちばん問題数の多い端末の log から、日ごとに集計して書き直す
function rebuildDaily_(){
  const b = best_(readDevices_());
  const sh = sheet_(SHEET_DAILY);
  sh.clear();
  const head = ["日付", "曜日", "解いた問題", "正解", "正答率", "リスニング", "文法"];
  sh.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight("bold");
  sh.setFrozenRows(1);
  if (!b) return;
  let log = [];
  try{ log = JSON.parse(b.data["eiken4.log"]); }catch(e){}
  if (!Array.isArray(log)) return;
  const days = {};
  log.forEach(function(x){
    if (!x || typeof x.d !== "string") return;
    const d = days[x.d] || (days[x.d] = {a:0, c:0, l:0, g:0});
    d.a++;
    if (x.ok) d.c++;
    if (x.k === "l") d.l++;
    else if (x.k === "g") d.g++;
  });
  const keys = Object.keys(days).sort().reverse();
  if (!keys.length) return;
  const rows = keys.map(function(d){
    const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(d);
    const wd = m ? "月火水木金土日".charAt(Number(Utilities.formatDate(new Date(Date.UTC(+m[1], +m[2] - 1, +m[3], 3)), "Asia/Tokyo", "u")) - 1) : "";
    const v = days[d];
    return [d, wd, v.a, v.c, v.a ? v.c / v.a : 0, v.l, v.g];
  });
  sh.getRange(2, 1, rows.length, 1).setNumberFormat("@");
  sh.getRange(2, 1, rows.length, head.length).setValues(rows);
  sh.getRange(2, 5, rows.length, 1).setNumberFormat("0%");
}
