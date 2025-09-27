/************************************
 * YADEA AUDIT v6.5 — FINAL STABLE VERSION
 ************************************/

/*********** CONFIG ************/
const TEMP_FOLDER_NAME = '__temp_exports';
const SHEET_USERS = 'Users';
const SHEET_SHOPS = 'Shops';
const SHEET_AUDIT = 'StoreAudit';
const SHEET_ERROR_MAPPING = 'ErrorMapping';
const SHEET_TRAINING_LESSONS = 'TrainingLessons';
const SHEET_TRAINING_PROGRESS = 'TrainingProgress';
const SHEET_QUIZ_BANK = 'QuizBank';
const SHEET_DAILY_QUIZ_PASS = 'DailyQuizPass';
const SHEET_STOCK_IN = 'StockIn';
const SHEET_SALES_OUT = 'SalesOut';
const SHEET_PRODUCTS = 'Products';
const SHEET_TARGETS_SALE = 'Targets_Sale';
const DRIVE_ROOT_ID = '1FqcmYZJedEJdakQCY4RKQ68Iax8hfwcv';
const ALLOWED_EMAILS = ['cchatgptplus2025@gmail.com'];
const COLS = {
  Timestamp:1, Date:2, FullName:3, Phone:4, DVN:5, Shop:6, Sale:7,
  Region:8, Province:9, Points:10, Note:11, SI:12, SO:13, Filenames:14, Links:15,
  CreatedBy:16
};
/*********** UTILS ************/
function sha256B64_(s){ return Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(s||''))); }
function fmtYMD_(d){ return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd'); }
function monthRange_(yyyyMM){ const [y,m] = yyyyMM.split('-').map(Number);
return [new Date(y, m-1, 1), new Date(y, m, 1)]; }
function isActive_(v){ const s = String(v||'').trim().toLowerCase(); if(v===true||['true','1','yes','y','x','on','enable','enabled'].includes(s)) return true;
if(v===false||['false','0','no','lock','locked','off','disable','disabled'].includes(s)) return false; return true; }
function assureFolder_(parent, name){ const it = parent.getFoldersByName(name); return it.hasNext() ? it.next() : parent.createFolder(name);
}
function canon_(s) { return String(s||'').toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g,'').replace(/\s+/g,' ').trim(); }

/*********** AUTH & METADATA ***********/
function login(username, password){
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_USERS);
if(!sh) throw new Error('Chưa có sheet Users');
  const v = sh.getDataRange().getValues();
  const u = String(username||'').trim().toLowerCase();
  const pHash = sha256B64_(password);
for (let i=1;i<v.length;i++){
    const userRow = v[i];
    const usr = String(userRow[0]||'').trim().toLowerCase();
    if (usr !== u) continue;
const plainPasswordInSheet = String(userRow[1]||'').trim();
    if (!plainPasswordInSheet) continue;
    const correctHash = sha256B64_(plainPasswordInSheet);
    const full = String(userRow[2]||'').trim(), role = String(userRow[3]||'pg').trim().toLowerCase() ||
'pg',
          email = String(userRow[4]||'').trim(), active = isActive_(userRow[5]), region = String(userRow[6]||'').trim();
if (!active) throw new Error('Tài khoản đã bị khóa.');
    if (pHash === correctHash) {
      const token = Utilities.getUuid();
CacheService.getUserCache().put(token, JSON.stringify({username:u, fullName:full, role, email, region}), 6*3600);
      return { token };
} else {
      throw new Error('Mật khẩu không đúng.');
}
  }
  throw new Error('Không tìm thấy tài khoản.');
}

function requestPasswordReset(username) {
  const u = String(username || '').trim().toLowerCase();
if (!u) { throw new Error('Vui lòng nhập tên đăng nhập.'); }
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_USERS);
if (!sh) throw new Error('Không tìm thấy sheet Users.');
  const data = sh.getDataRange().getValues();
for (let i = 1; i < data.length; i++) {
    const row = data[i];
const userInSheet = String(row[0] || '').trim().toLowerCase();
    if (userInSheet === u) {
      const email = String(row[4] || '').trim();
const password = String(row[1] || '').trim();
      if (!email) { throw new Error('Tài khoản của bạn chưa được thiết lập email để khôi phục.');
}
      if (!password) { throw new Error('Không tìm thấy mật khẩu cho tài khoản này.');
}
      const subject = 'YADEA App - Khôi phục mật khẩu của bạn';
const body = `Chào ${row[2] || u},<br><br>Bạn đã yêu cầu khôi phục mật khẩu.<br>Tên đăng nhập của bạn là: <b>${u}</b><br>Mật khẩu của bạn là: <b>${password}</b><br><br>Vui lòng đăng nhập lại và bảo mật thông tin này.`;
MailApp.sendEmail({ to: email, subject: subject, htmlBody: body });
      return { message: 'Một email chứa mật khẩu đã được gửi đến địa chỉ ' + email + '.\nVui lòng kiểm tra hộp thư của bạn.'
};
    }
  }
  throw new Error('Không tìm thấy tên đăng nhập này trong hệ thống.');
}

function requireAuth_(token){
  if (!token) throw new Error('Hết phiên đăng nhập. Vui lòng đăng nhập lại.');
  const raw = CacheService.getUserCache().get(token);
if (!raw) throw new Error('Hết phiên đăng nhập. Vui lòng đăng nhập lại.');
  return JSON.parse(raw);
}

function getAdminInfo(token){
  const me = requireAuth_(token);
  let visibleTabs = [];
let meta = { regions:[], provinces:[], sales:[], svns: [] };

  const role = me.role;
  const isAdmin = (role === 'admin');
const isManager = (role === 'khu vực');

  if (isAdmin) {
    visibleTabs = ['settargets', 'upload', 'history', 'progress', 'quick', 'report', 'stockin', 'salesout', 'training'];
meta = getMetaFiltersPublic();
  } else if (isManager) {
    visibleTabs = ['upload', 'history', 'progress', 'quick', 'report', 'stockin', 'salesout', 'training'];
const dependentFilters = getDependentFilters(me.region);
    meta = { ...dependentFilters, regions: [me.region] };
} else { // Sale or PG
    visibleTabs = ['upload', 'history', 'quick', 'stockin', 'salesout', 'training']; // Đã thêm 'quick' và 'stockin'
meta = { regions:[], provinces:[], sales:[], svns: [] };
  }

  return { me, meta, userRole: role, visibleTabs };
}

function getMetaFiltersPublic(){
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_SHOPS);
  const usersSh = SpreadsheetApp.getActive().getSheetByName(SHEET_USERS);
  
  if(!sh) return {regions:[], provinces:[], sales:[], svns:[]};
const vals = sh.getDataRange().getValues();
  const regions = new Set(), provinces = new Set(), saleFullNames = new Set(), svns = new Set();
for (let i=1;i<vals.length;i++){
    if (vals[i][2]) svns.add(vals[i][2]);         // Column C: SVN
    if (vals[i][3]) saleFullNames.add(vals[i][3]);
// Column D: Tên Sale
    if (vals[i][5]) regions.add(vals[i][5]);
// Column F: Khu vực
    if (vals[i][6]) provinces.add(vals[i][6]);
// Column G: Tỉnh
  }

  const userMap = new Map();
if(usersSh){
    usersSh.getDataRange().getValues().slice(1).forEach(r => {
        userMap.set(r[2], r[0]) // key: fullName, value: username
    });
}

  const salesList = [...saleFullNames].sort().map(fullName => ({
      username: userMap.get(fullName) || fullName,
      fullName: fullName
  }));
return { regions:[...regions].sort(), provinces:[...provinces].sort(), sales: salesList, svns: [...svns].sort() };
}

function getDependentFilters(region) {
  if (!region) return { provinces: [], sales: [], svns: [] };
const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_SHOPS);
  if (!sh) return { provinces: [], sales: [], svns: [] };
  
  const vals = sh.getDataRange().getValues();
const provinces = new Set(), salesFullNames = new Set(), svns = new Set();
  const rLower = region.toLowerCase();
for (let i = 1; i < vals.length; i++) {
    if (String(vals[i][5] || '').toLowerCase() === rLower) { // Column F: Khu vực
      if (vals[i][2]) svns.add(vals[i][2]);
if (vals[i][3]) salesFullNames.add(vals[i][3]);
      if (vals[i][6]) provinces.add(vals[i][6]);
    }
  }

  const usersSh = SpreadsheetApp.getActive().getSheetByName(SHEET_USERS);
  const userMap = new Map();
if(usersSh){
    usersSh.getDataRange().getValues().slice(1).forEach(r => {
        userMap.set(r[2], r[0])
    });
}

  const salesList = [...salesFullNames].sort().map(fullName => ({
      username: userMap.get(fullName) || fullName,
      fullName: fullName
  }));
return { provinces: [...provinces].sort(), sales: salesList, svns: [...svns].sort() };
}

function getReportFilters(filters) {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_SHOPS);
if (!sh) return { provinces: [], sales: [], svns: [] };

  const allShops = sh.getDataRange().getValues().slice(1);
  let filteredShops = allShops;
if (filters && filters.region) {
    const rLower = filters.region.toLowerCase();
filteredShops = filteredShops.filter(row => String(row[5] || '').toLowerCase() === rLower); // Column F
  }

  const provinces = new Set(), salesFullNames = new Set(), svns = new Set();
filteredShops.forEach(row => {
    if (row[2]) svns.add(row[2]);
    if (row[3]) salesFullNames.add(row[3]);
    if (row[6]) provinces.add(row[6]);
  });
const usersSh = SpreadsheetApp.getActive().getSheetByName(SHEET_USERS);
  const userMap = new Map();
  if(usersSh){
    usersSh.getDataRange().getValues().slice(1).forEach(r => {
        userMap.set(r[2], r[0])
    });
}

  const salesList = [...salesFullNames].sort().map(fullName => ({
      username: userMap.get(fullName) || fullName,
      fullName: fullName
  }));
return {
    provinces: [...provinces].sort(),
    sales: salesList,
    svns: [...svns].sort()
  };
}

function getDvnsForSvn(svnCode) {
  if (!svnCode) return [];
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_SHOPS);
  if (!sh) return [];
  const vals = sh.getDataRange().getValues();
const dvns = new Set();
  const svnLower = svnCode.toLowerCase();
  for (let i = 1; i < vals.length; i++) {
    if (String(vals[i][2] || '').toLowerCase() === svnLower) {
      if (vals[i][0]) dvns.add(vals[i][0]);
}
  }
  return [...dvns].sort();
}

function getShopInfoByDvn(token, dvnCode) {
  requireAuth_(token);
  const dvnUpper = String(dvnCode || '').trim().toUpperCase();
if (!dvnUpper) return null;
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_SHOPS);
  if (!sh) throw new Error('Không tìm thấy sheet Shops.');
const values = sh.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
if (String(row[0] || '').trim().toUpperCase() === dvnUpper) {
      return {
        shopName: row[1] ||
'',      // Col B
        svn: row[2] ||
'',           // Col C
        saleName: row[3] ||
'',      // Col D
        saleUsername: row[4] ||
'',  // Col E
        region: row[5] ||
'',        // Col F
        province: row[6] ||
''       // Col G
      };
}
  }
  return null;
}

function saveImages(payload) {
  try {
    const me = requireAuth_(payload.__auth);
const date = new Date(payload.date);
    const dvnCode = String(payload.dvnCode || '').trim().toUpperCase();
if (!dvnCode || (payload.images || []).length === 0) { throw new Error('Thiếu mã DVN hoặc không có ảnh nào để tải lên.');
}
    const shopInfo = getShopInfoByDvn(payload.__auth, dvnCode) || {};
    const rootFolder = DriveApp.getFolderById(DRIVE_ROOT_ID);
const monthFolder = assureFolder_(rootFolder, Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM'));
    const regionFolder = assureFolder_(monthFolder, shopInfo.region || 'Chưa xác định Khu vực');
const saleFolder = assureFolder_(regionFolder, shopInfo.saleName || 'Chưa xác định Sale');
    const provinceFolder = assureFolder_(saleFolder, shopInfo.province || 'Chưa xác định Tỉnh');
const dvnFolder = assureFolder_(provinceFolder, dvnCode);
    const finalUploadFolder = assureFolder_(dvnFolder, fmtYMD_(date));
    const savedFileLinks = [], savedFileNames = [];
payload.images.forEach((img) => {
      const blob = Utilities.newBlob(Utilities.base64Decode(img.dataUrl.split(',')[1]), MimeType.JPEG, img.name);
      const file = finalUploadFolder.createFile(blob);
      savedFileLinks.push(file.getUrl());
      savedFileNames.push(file.getName());
    });
const auditSheet = SpreadsheetApp.getActive().getSheetByName(SHEET_AUDIT);
    if (!auditSheet) { throw new Error('Không tìm thấy sheet StoreAudit.');
}
    auditSheet.appendRow([ new Date(), date, me.fullName || me.username, payload.upEmp || '', dvnCode, shopInfo.shopName || '', shopInfo.saleName || '', shopInfo.region || '', shopInfo.province || '', payload.points, payload.note, payload.si || 0, payload.so || 0, savedFileNames.join('\n'), savedFileLinks.join('|'), me.username ]);
return { message: 'Đã tải lên thành công ' + payload.images.length + ' ảnh.' };
} catch (e) {
    console.error("Lỗi saveImages: " + e.toString());
throw new Error('Lỗi phía máy chủ: ' + e.message);
  }
}

/*********** QUIZ FUNCTIONS ***********/
function getDailyQuiz() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_QUIZ_BANK);
if (!sh || sh.getLastRow() < 2) return { questions: [] };
  const values = sh.getDataRange().getValues().slice(1);
const activeQuestions = values.filter(row => row[7] === true);
  for (let i = activeQuestions.length - 1; i > 0; i--) { const j = Math.floor(Math.random() * (i + 1));
[activeQuestions[i], activeQuestions[j]] = [activeQuestions[j], activeQuestions[i]]; }
  const selectedQuestions = activeQuestions.slice(0, 3).map(row => ({ id: row[0], text: row[1], options: { A: row[2], B: row[3], C: row[4], D: row[5] } }));
return { questions: selectedQuestions };
}

function submitQuiz(payload) {
  const me = requireAuth_(payload.authToken);
  const userAnswers = payload.answers;
const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_QUIZ_BANK);
  if (!sh) throw new Error("Không tìm thấy ngân hàng câu hỏi.");
  const questionIds = Object.keys(userAnswers);
if (questionIds.length === 0) throw new Error("Không có câu trả lời nào được gửi.");
  const allQuestions = sh.getDataRange().getValues().slice(1);
const correctAnswers = {};
  allQuestions.forEach(row => { correctAnswers[row[0]] = row[6]; });
  let allCorrect = questionIds.every(id => userAnswers[id] === correctAnswers[id]);
if (allCorrect) {
    const passSh = SpreadsheetApp.getActive().getSheetByName(SHEET_DAILY_QUIZ_PASS);
    if (passSh) passSh.appendRow([me.username, new Date()]);
    return { passed: true };
} else {
    return { passed: false, message: "Có câu trả lời chưa đúng. Vui lòng thử lại."
};
  }
}

function checkQuizStatus(token) {
  const me = requireAuth_(token);
  const todayStr = fmtYMD_(new Date());
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_DAILY_QUIZ_PASS);
if (!sh || sh.getLastRow() < 2) return { completed: false };
  const data = sh.getDataRange().getValues().slice(1);
return { completed: data.some(row => String(row[0]).toLowerCase() === me.username && fmtYMD_(new Date(row[1])) === todayStr) };
}

/*********** DATA RETRIEVAL (HISTORY, PROGRESS, REPORTS) ***********/
function getHistory(filter){
  const me = requireAuth_(filter.__auth);
  const isManager = me.role === 'khu vực';
const isSaleOrPG = (me.role === 'sale' || me.role === 'pg');

  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_AUDIT);
  if(!sh) return {rows:[]};
const allData = sh.getDataRange().getValues().slice(1);

  const from = filter.from ? new Date(filter.from) : null;
  const to = filter.to ?
new Date(filter.to) : null;
  const regionF = String(filter.region||'').trim().toLowerCase();
  const provinceF = String(filter.province||'').trim().toLowerCase();
  const saleF = String(filter.sale||'').trim();
const svnF = String(filter.svn || '').trim().toLowerCase();
  const dvnF = String(filter.dvn||'').trim().toLowerCase();
  const shopF = String(filter.shop||'').trim().toLowerCase();
  const limit = Number(filter.limit||200)||200;

  let out=[];
for (let i = allData.length - 1; i >= 0; i--) {
    if (out.length >= limit) break;
const r = allData[i];

    const createdBy = String(r[COLS.CreatedBy-1]||'').trim().toLowerCase();
    const rowRegion = String(r[COLS.Region-1]||'').trim().toLowerCase();
    if (isSaleOrPG && createdBy !== me.username) continue;
if (isManager && me.region && rowRegion !== me.region.toLowerCase()) continue;
    const ymd = new Date(r[COLS.Date-1]);
if (from && ymd < from) continue;
    if (to && ymd > to) continue;
if (regionF && rowRegion !== regionF) continue;
    if (provinceF && String(r[COLS.Province-1]||'').trim().toLowerCase() !== provinceF) continue;
if (saleF && String(r[COLS.Sale-1]||'').trim() !== saleF) continue;
    if (dvnF && !String(r[COLS.DVN-1]||'').trim().toLowerCase().includes(dvnF)) continue;
    if (shopF && !String(r[COLS.Shop-1]||'').trim().toLowerCase().includes(shopF)) continue;
const names = String(r[COLS.Filenames-1]||'').split('\n').filter(Boolean);
    const links = String(r[COLS.Links-1]||'').split('|').filter(Boolean);
    const images = names.map((name, k) => ({name: name || `Ảnh ${k+1}`, url: links[k] || ''}));
out.push({ date: fmtYMD_(ymd), dvn: r[COLS.DVN-1]||'', shop: r[COLS.Shop-1]||'', sale: r[COLS.Sale-1]||'', region: r[COLS.Region-1]||'', province: r[COLS.Province-1]||'', points: Number(r[COLS.Points-1]||0), note: r[COLS.Note-1]||'', images: images });
}
  return {rows:out};
}

function getProgress(filter){
  const me = requireAuth_(filter.__auth);
  const isManager = me.role === 'khu vực';
const month = filter.month || fmtYMD_(new Date()).slice(0,7);
  const [start, end] = monthRange_(month);
  const shopsSh = SpreadsheetApp.getActive().getSheetByName(SHEET_SHOPS);
  if(!shopsSh) return {summary:{month,total:0,done:0}, rows:[]};
const shops = shopsSh.getDataRange().getValues().slice(1);

  const auditSh = SpreadsheetApp.getActive().getSheetByName(SHEET_AUDIT);
  const audits = (auditSh && auditSh.getLastRow()>=2) ? auditSh.getRange(2,1,auditSh.getLastRow()-1,COLS.DVN).getValues() : [];
const lastByDVN = new Map();
  audits.forEach(r=>{
    const d = new Date(r[COLS.Date-1]);
    if (d>=start && d<end){
      const dvn = String(r[COLS.DVN-1]||'').toUpperCase();
      const ymd = fmtYMD_(d);
      if (!lastByDVN.has(dvn) || lastByDVN.get(dvn)<ymd) lastByDVN.set(dvn, ymd);
    }
  });
const regionF = String(filter.region||'').trim().toLowerCase();
  const provinceF = String(filter.province||'').trim().toLowerCase();
  const saleF = String(filter.sale||'').trim();
  const statusF = String(filter.status||'').trim().toLowerCase();
  const svnF = String(filter.svn||'').trim().toLowerCase();
const dvnF = String(filter.dvn||'').trim().toLowerCase();

  let rows=[];
  for (const shopRow of shops) {
    // [FIX 2] Bỏ qua các hàng trống bằng cách kiểm tra Mã DVN
    const dvn = String(shopRow[0]||'').trim();
if (!dvn) continue;

    const svn = String(shopRow[2]||'').trim().toLowerCase();
    const sale = String(shopRow[3]||'').trim();
    const reg = String(shopRow[5]||'').trim().toLowerCase(); 
    const prov = String(shopRow[6]||'').trim().toLowerCase();
if (isManager && me.region && reg !== me.region.toLowerCase()) continue;
    if (regionF && reg !== regionF) continue;
if (provinceF && prov !== provinceF) continue;
    if (saleF && sale !== saleF) continue;
if (svnF && svn !== svnF) continue;
    if (dvnF && dvn.toLowerCase() !== dvnF) continue;

    const lastDate = lastByDVN.get(dvn.toUpperCase())||'';
const status = lastDate ? 'done' : 'todo';
    if (statusF && status !== statusF) continue;
rows.push({ region: shopRow[5]||'', province: shopRow[6]||'', sale: shopRow[3]||'', dvn: shopRow[0]||'', shop: shopRow[1]||'', lastDate, status });
}
  rows.sort((a,b)=>{ if (a.status!==b.status) return a.status==='todo'?1:-1; return (a.region+a.sale+a.dvn).localeCompare(b.region+b.sale+a.dvn); });
  const total = rows.length;
  const done = rows.filter(r=>r.status==='done').length;
return {summary:{month,total,done}, rows};
}

const ERROR_CACHE = { mapping: null, timestamp: 0 };
function getErrorMapping_() {
    const now = Date.now();
if (ERROR_CACHE.mapping && (now - ERROR_CACHE.timestamp < 300000)) { return ERROR_CACHE.mapping; }
    const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_ERROR_MAPPING);
if (!sh) return [];
    const values = sh.getDataRange().getValues().slice(1);
    const keywordList = values.map(row => {
        const displayName = String(row[0] || '').trim();
        const searchKeywords = String(row[1] || '').split(',').map(kw => canon_(kw)).filter(Boolean);
        return (displayName && searchKeywords.length > 0) ? { name: displayName, searchKeys: searchKeywords } : null;
    }).filter(Boolean);
ERROR_CACHE.mapping = keywordList;
    ERROR_CACHE.timestamp = now;
    return keywordList;
}

function getQuickReportV2(filter){
  const me = requireAuth_(filter.__auth);
const isManager = me.role === 'khu vực';
  const month = filter.month || fmtYMD_(new Date()).slice(0,7);
  const [start, end] = monthRange_(month);
const shShops = SpreadsheetApp.getActive().getSheetByName(SHEET_SHOPS);
  if(!shShops) return {ok:false,message:'Chưa có sheet Shops'};

  const regionF = String(filter.region||'').trim().toLowerCase();
  const provinceF = String(filter.province||'').trim().toLowerCase();
const saleF = String(filter.sale||'').trim();
  const svnF = String(filter.svn||'').trim().toLowerCase();
  const dvnF = String(filter.dvn||'').trim().toLowerCase();

  const shops = shShops.getDataRange().getValues().slice(1);
  const shAudit = SpreadsheetApp.getActive().getSheetByName(SHEET_AUDIT);
const audits = (shAudit && shAudit.getLastRow()>1) ? shAudit.getRange(2, 1, shAudit.getLastRow()-1, COLS.Note).getValues() : [];

  const byDVN = new Map();
audits.forEach(r=>{
      const d = new Date(r[COLS.Date-1]);
      if (d>=start && d<end){
        const dvn = String(r[COLS.DVN-1]||'').toUpperCase();
        const ymd = fmtYMD_(d);
        if (!byDVN.has(dvn) || byDVN.get(dvn).date < ymd){
          byDVN.set(dvn, { date: ymd, points: Number(r[COLS.Points-1]||0), note: String(r[COLS.Note-1]||'').trim(), timestamp: r[COLS.Timestamp-1] });
        }
      }
  });
const allErrorTypes = getErrorMapping_();
  const buckets = Object.fromEntries(allErrorTypes.map(kw => [kw.name, 0]));
  buckets['Khác'] = 0;

  let total=0, done=0, sumPts=0, cntPts=0;
const denomBySale = {}, doneBySale = {};
  const auditedShopsDetails = [];
shops.forEach(s=>{
      const dvn = String(s[0]||'').trim().toLowerCase();
      if(!dvn) return; // Bỏ qua hàng trống
      const shopName = String(s[1]||'').trim();
      const svn = String(s[2]||'').trim().toLowerCase();
      const sale = String(s[3]||'').trim();
      const reg = String(s[5]||'').trim().toLowerCase();
      const prov = String(s[6]||'').trim().toLowerCase();

      if (isManager && me.region && reg !== me.region.toLowerCase()) return;
      if (regionF && reg !== regionF) return;
      if (provinceF 
&& prov !== provinceF) return;
      if (saleF 
&& sale !== saleF) return;
      if (svnF && svn !== svnF) return;
      if (dvnF && dvn !== dvnF) return;

      if (!dvn || !sale) return;

      total++;
      denomBySale[sale] = (denomBySale[sale]||0)+1;

      const last = byDVN.get(dvn.toUpperCase());
      if (last){
        done++;
        doneBySale[sale] = (doneBySale[sale]||0)+1;
  
auditedShopsDetails.push({ dvn: dvn.toUpperCase(), shopName: shopName, points: 
last.points, timestamp: last.timestamp });
        if (last.points){ sumPts+=last.points; cntPts++;
}

        const noteCanon = canon_(last.note);
let found = false;
if (noteCanon) {
          for (const errorType of allErrorTypes) {
            if (errorType.searchKeys.some(key => noteCanon.includes(key))) {
              buckets[errorType.name]++;
found = true;
              break;
            }
          }
        }
        if (!found) buckets['Khác']++;
}
  });

  const percentCompleted = total ? Math.round(done*1000/total)/10 : 0;
  const avgPointsPerShop = cntPts ? Math.round((sumPts/cntPts)*10)/10 : 0;
const errorBreakdown = Object.entries(buckets).map(([key, value]) => ({ name: key, count: value })).filter(item => item.count > 0);
const salesRanking = Object.keys(denomBySale).map(saleName=>{
      const totalAss = denomBySale[saleName]||0, completed = doneBySale[saleName]||0, percent = totalAss? Math.round(completed*1000/totalAss)/10 : 0;
      return { name: saleName, completed, total: totalAss, percent };
    }).sort((a,b)=> b.percent - a.percent || b.completed - a.completed);
const topShops = auditedShopsDetails.sort((a, b) => b.points - a.points || new Date(a.timestamp) - new Date(b.timestamp))
    .slice(0, 10)
    .map(s => ({ dvn: s.dvn, shopName: s.shopName, points: s.points, date: fmtYMD_(new Date(s.timestamp)) }));
return { ok:true, month, totalShops: total, totalAuditedShops: done, percentCompleted, avgPointsPerShop, errorBreakdown, salesRanking, topShops };
}

/*********** STABLE CSV EXPORT & OTHER FUNCTIONS ***********/
function getTempFolder_() {
    const root = DriveApp.getFolderById(DRIVE_ROOT_ID);
    return assureFolder_(root, TEMP_FOLDER_NAME);
}
function generateProgressCsvUrl(filter) {
  requireAuth_(filter.__auth);
  const data = getProgress(filter);
  if (!data.rows || !data.rows.length) throw new Error('Không có dữ liệu để xuất.');
const header = ['Khu vực','Tỉnh','Sale','DVN','Shop','Ngày cập nhật','Trạng thái'];
  const esc = s => (/[",\n]/.test(String(s??'')) ? '"' + String(s).replace(/"/g,'""') + '"' : String(s??''));
const lines = [header.join(',')];
  data.rows.forEach(r => lines.push([r.region, r.province, r.sale, r.dvn, r.shop, r.lastDate, (r.status==='done'?'Hoàn thành':'Chưa làm')].map(esc).join(',')));
const csvContent = '\uFEFF' + lines.join('\n');
  const tempFolder = getTempFolder_();
  const fileName = `YADEA_Progress_${filter.month}_${Date.now()}.csv`;
  const file = tempFolder.createFile(fileName, csvContent, MimeType.CSV);
file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return `https://drive.google.com/uc?export=download&id=${file.getId()}`;
}
function generateQuickCsvUrl(filter) {
  requireAuth_(filter.__auth);
  const rep = getQuickReportV2(filter);
if (!rep || !rep.ok) throw new Error(rep.message || 'Không có dữ liệu để xuất.');
const rows = [], esc = s => (/[",\n]/.test(String(s??'')) ? '"' + String(s).replace(/"/g,'""') + '"' : String(s??''));
  rows.push(['Tháng', rep.month]);
rows.push(['Số shop đã hoàn thành', rep.totalAuditedShops||0]);
  rows.push(['% hoàn thành', (rep.percentCompleted||0)+'%']);
  rows.push(['Điểm TB / shop', rep.avgPointsPerShop||0]);
  rows.push([]);
  rows.push(['Phân loại lỗi','Số lượng']);
(rep.errorBreakdown||[]).forEach(e=> rows.push([e.name, e.count||0]));
  rows.push([]);
  rows.push(['Xếp hạng nhân viên','% hoàn thành','Hoàn thành/Tổng']);
  (rep.salesRanking||[]).forEach(r=> rows.push([r.name, r.percent||0, (r.completed||0)+'/'+(r.total||0)]));
const csvContent = '\uFEFF' + rows.map(r => r.map(esc).join(',')).join('\n');
  const tempFolder = getTempFolder_();
  const fileName = `YADEA_QuickReport_${filter.month}_${Date.now()}.csv`;
const file = tempFolder.createFile(fileName, csvContent, MimeType.CSV);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return `https://drive.google.com/uc?export=download&id=${file.getId()}`;
}

// [FIX 3] Bổ sung các hàm xuất CSV còn thiếu
function generateStockInReportCsvUrl(filter) {
  requireAuth_(filter.__auth);
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_STOCK_IN);
if (!sh || sh.getLastRow() < 2) throw new Error('Không có dữ liệu S.I để xuất.');
const allData = sh.getRange(2, 1, sh.getLastRow() - 1, 12).getValues();
  const isManager = requireAuth_(filter.__auth).role === 'khu vực';
  const me = requireAuth_(filter.__auth);
const regionFilter = String(filter.region || '').trim().toLowerCase();
  const provinceFilter = String(filter.province || '').trim().toLowerCase();
  const saleFilter = String(filter.sale || '').trim().toLowerCase();
const svnFilter = String(filter.svn || '').trim().toLowerCase();
  const dvnFilter = String(filter.dvn || '').trim().toLowerCase();
  const shopFilter = String(filter.shop || '').trim().toLowerCase();
const filteredData = allData.filter(row => {
    const rowRegion = String(row[6] || '').trim().toLowerCase();
    if (isManager && me.region && rowRegion !== me.region.toLowerCase()) return false;
    if (regionFilter && rowRegion !== regionFilter) return false;
    if (provinceFilter && String(row[7] || '').trim().toLowerCase() !== provinceFilter) return false;
    if (saleFilter && String(row[11] || '').trim().toLowerCase() !== saleFilter) return false;
    if (svnFilter && String(row[4] || '').trim().toLowerCase() !== svnFilter) return false;
    if (dvnFilter && String(row[2] || '').trim().toLowerCase() !== dvnFilter) return false;
    if (shopFilter && !String(row[3] || '').trim().toLowerCase().includes(shopFilter)) 
return false;
    return true;
  });

  if (filteredData.length === 0) throw new Error('Không có dữ liệu S.I nào khớp với bộ lọc.');
const header = ['Ngày Nhập', 'Mã DVN', 'Tên Shop', 'Mã SVN', 'Tên Sale', 'Khu Vực', 'Tỉnh', 'Tên Xe', 'Số Lượng'];
const esc = s => (/[",\n]/.test(String(s??'')) ? '"' + String(s).replace(/"/g,'""') + '"' : String(s??''));
  const lines = [header.join(',')];
filteredData.forEach(r => lines.push([fmtYMD_(new Date(r[1])), r[2], r[3], r[4], r[5], r[6], r[7], r[8], r[9]].map(esc).join(',')));
  
  const csvContent = '\uFEFF' + lines.join('\n');
const tempFolder = getTempFolder_();
  const fileName = `YADEA_StockIn_Report_${Date.now()}.csv`;
  const file = tempFolder.createFile(fileName, csvContent, MimeType.CSV);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return `https://drive.google.com/uc?export=download&id=${file.getId()}`;
}

function generateSalesOutReportCsvUrl(filter) {
  requireAuth_(filter.__auth);
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_SALES_OUT);
  if (!sh || sh.getLastRow() < 2) throw new Error('Không có dữ liệu S.O để xuất.');
const allData = sh.getRange(2, 1, sh.getLastRow() - 1, 14).getValues();
  const isManager = requireAuth_(filter.__auth).role === 'khu vực';
  const me = requireAuth_(filter.__auth);
const regionFilter = String(filter.region || '').trim().toLowerCase();
  const provinceFilter = String(filter.province || '').trim().toLowerCase();
  const saleFilter = String(filter.sale || '').trim().toLowerCase();
const svnFilter = String(filter.svn || '').trim().toLowerCase();
  const dvnFilter = String(filter.dvn || '').trim().toLowerCase();
  const shopFilter = String(filter.shop || '').trim().toLowerCase();
const filteredData = allData.filter(row => {
    const rowRegion = String(row[6] || '').trim().toLowerCase();
    if (isManager && me.region && rowRegion !== me.region.toLowerCase()) return false;
    if (regionFilter && rowRegion !== regionFilter) return false;
    if (provinceFilter && String(row[7] || '').trim().toLowerCase() !== provinceFilter) return false;
    if (saleFilter && String(row[13] || '').trim().toLowerCase() !== saleFilter) return false;
    if (svnFilter && String(row[4] || '').trim().toLowerCase() !== svnFilter) return false;
    if (dvnFilter && String(row[2] || '').trim().toLowerCase() !== dvnFilter) return false;
    if (shopFilter && !String(row[3] || '').trim().toLowerCase().includes(shopFilter)) 
return false;
    return true;
  });

  if (filteredData.length === 0) throw new Error('Không có dữ liệu S.O nào khớp với bộ lọc.');
const header = ['Ngày Bán', 'Mã DVN', 'Tên Shop', 'Mã SVN', 'Tên Sale', 'Khu Vực', 'Tỉnh', 'PG/NV Bán', 'Tên Xe', 'Số Lượng', 'Doanh Thu'];
const esc = s => (/[",\n]/.test(String(s??'')) ? '"' + String(s).replace(/"/g,'""') + '"' : String(s??''));
  const lines = [header.join(',')];
filteredData.forEach(r => lines.push([fmtYMD_(new Date(r[1])), r[2], r[3], r[4], r[5], r[6], r[7], r[8], r[9], r[10], r[11]].map(esc).join(',')));

  const csvContent = '\uFEFF' + lines.join('\n');
const tempFolder = getTempFolder_();
  const fileName = `YADEA_SalesOut_Report_${Date.now()}.csv`;
  const file = tempFolder.createFile(fileName, csvContent, MimeType.CSV);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return `https://drive.google.com/uc?export=download&id=${file.getId()}`;
}


/*********** TRAINING PANEL FUNCTIONS V2 ***********/
function getTrainingCategories(token) {
    requireAuth_(token);
    const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_TRAINING_LESSONS);
if (!sh) return { ok: false, categories: [] };
    const values = sh.getRange(2, 3, sh.getLastRow() - 1, 1).getValues();
const categorySet = new Set(values.flat().filter(Boolean));
    return { ok: true, categories: [...categorySet].sort() };
}

function listTrainingLessons(filter) {
    requireAuth_(filter.__auth);
const categoryFilter = String(filter.category || '').trim();
    if (!categoryFilter) return { ok: false, items: [], message: 'Vui lòng chọn một danh mục.'
};
    const lessonsSh = SpreadsheetApp.getActive().getSheetByName(SHEET_TRAINING_LESSONS);
    if (!lessonsSh) return { ok: false, items: [], message: 'Training sheet not found.' };
const lessonsData = lessonsSh.getDataRange().getValues().slice(1);
    const items = lessonsData.map(r => {
        const active = r[9] === true || String(r[9]).trim().toLowerCase() === 'true';
        if (active && String(r[2] || '').trim() === categoryFilter && r[0] && r[1] && r[7]) {
            return { id: r[0], title: r[1], tags: String(r[3]||'').split(',').map(t=>t.trim()).filter(Boolean), thumbnail: r[6], fileUrl: r[7] };
        }
        return null;
    }).filter(Boolean);
return { ok: true, items: items, category: categoryFilter };
}

/*********** MAIN HANDLER ***********/
function doGet(e) {
  try {
    // 1. Tạo template từ file
    // 2. Dùng .evaluate() để xử lý và trả về một HtmlOutput
    // 3. SAU ĐÓ mới gọi các hàm cài đặt trên HtmlOutput đó
    return HtmlService.createTemplateFromFile('IndexMobile')
      .evaluate() // Bước này phải được thực hiện trước
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover');
  } catch(err) {
    // Trả về lỗi nếu có vấn đề để dễ dàng gỡ rối
    return ContentService.createTextOutput(JSON.stringify({error: err.message}));
  }
}

/*********** STOCK & SALES FUNCTIONS ***********/
function getProducts(token) {
  requireAuth_(token);
  const sh = SpreadsheetApp.getActive().getSheetByName('Products');
  if (!sh) return [];
const values = sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues();
  return values.map(row => (row[0] && !isNaN(row[1])) ? { name: row[0], price: Number(row[1]) } : null)
               .filter(Boolean)
               .sort((a,b) => a.name.localeCompare(b.name));
}

function saveStockIn(payload) {
  try {
    const me = requireAuth_(payload.__auth);
const { date, dvnCode, shopName, svn, sale, region, province, vehicles } = payload;
    const shopInfo = getShopInfoByDvn(payload.__auth, dvnCode) || {};
const saleUsername = shopInfo.saleUsername || '';

    if (!dvnCode || !vehicles || vehicles.length === 0) throw new Error('Vui lòng nhập Mã DVN và ít nhất một loại xe.');
if (!date) throw new Error('Vui lòng chọn ngày nhập kho.');
    const stockInSheet = SpreadsheetApp.getActive().getSheetByName(SHEET_STOCK_IN);
if (!stockInSheet) throw new Error('Không tìm thấy sheet StockIn.');
    const now = new Date();
    const stockInDate = new Date(date);
vehicles.forEach(vehicle => {
      if (vehicle.name && vehicle.quantity > 0) {
        stockInSheet.appendRow([ now, stockInDate, dvnCode, shopName, svn, sale, region, province, vehicle.name, vehicle.quantity, me.username, saleUsername ]);
      }
    });
return { message: `Đã nhập kho thành công ${vehicles.length} loại xe.` };
} catch (e) {
    console.error("Lỗi saveStockIn: " + e.toString());
    throw new Error('Lỗi máy chủ: ' + e.message);
}
}

function saveSalesOut(payload) {
  try {
    const me = requireAuth_(payload.__auth);
const { date, dvnCode, shopName, svn, sale, region, province, pgName, vehicles } = payload;
    const shopInfo = getShopInfoByDvn(payload.__auth, dvnCode) ||
{};
    const saleUsername = shopInfo.saleUsername || '';

    if (!dvnCode || !vehicles || vehicles.length === 0) throw new Error('Vui lòng nhập Mã DVN và ít nhất một loại xe.');
if (!date) throw new Error('Vui lòng chọn ngày bán.');
    const salesOutSheet = SpreadsheetApp.getActive().getSheetByName(SHEET_SALES_OUT);
if (!salesOutSheet) throw new Error('Không tìm thấy sheet SalesOut.');
    const productsData = getProducts(payload.__auth);
const priceMap = new Map(productsData.map(p => [p.name, p.price]));
    const now = new Date();
    const saleDate = new Date(date);
vehicles.forEach(vehicle => {
      if (vehicle.name && vehicle.quantity > 0) {
        const price = priceMap.get(vehicle.name) || 0;
        const revenue = price * vehicle.quantity;
        salesOutSheet.appendRow([ now, saleDate, dvnCode, shopName, svn, sale, region, province, pgName || me.fullName, vehicle.name, vehicle.quantity, revenue, me.username, saleUsername ]);
      }
    });
return { message: `Đã lưu thành công ${vehicles.length} loại xe đã bán.` };
} catch (e) {
    console.error("Lỗi saveSalesOut: " + e.toString());
    throw new Error('Lỗi máy chủ: ' + e.message);
}
}

/*********** BÁO CÁO NHẬP / BÁN HÀNG (S.I/S.O) - ĐÃ CẬP NHẬT ***********/

/**
 * [HELPER] Lấy dữ liệu target S.I và S.O trong tháng, nhóm theo Sale và Khu vực.
* @param {string} yyyyMM - Tháng cần lấy dữ liệu, định dạng 'yyyy-MM'.
* @returns {object} Trả về các target đã được tổng hợp.
 */
function getTargetsByMonth_(yyyyMM) {
  const targetsSh = SpreadsheetApp.getActive().getSheetByName(SHEET_TARGETS_SALE);
if (!targetsSh) return { bySale: new Map(), byRegion: new Map(), totals: { si: 0, so: 0 } };
const usersSh = SpreadsheetApp.getActive().getSheetByName(SHEET_USERS);
  const userToRegionMap = new Map();
  if (usersSh) {
    usersSh.getDataRange().getValues().slice(1).forEach(r => {
      if (r[0] && r[6]) userToRegionMap.set(String(r[0]).trim(), String(r[6]).trim());
    });
}

  const targetsBySale = new Map();
  const targetsByRegion = new Map();
  let totalSI = 0, totalSO = 0;
targetsSh.getDataRange().getValues().slice(1).forEach(row => {
    if (!row[0]) return;
    const rowMonth = Utilities.formatDate(new Date(row[0]), Session.getScriptTimeZone(), 'yyyy-MM');
    if (rowMonth === yyyyMM) {
      const saleUsername = String(row[1]).trim();
      const region = userToRegionMap.get(saleUsername) || 'Chưa xác định';
      const targetSI = Number(row[3] || 0);
      const targetSO = Number(row[4] || 0);

      targetsBySale.set(saleUsername, { si: targetSI, so: targetSO, region: region });
      
      const currentRegionTarget = targetsByRegion.get(region) 
|| { si: 0, so: 0 };
      currentRegionTarget.si += targetSI;
      currentRegionTarget.so += targetSO;
      targetsByRegion.set(region, currentRegionTarget);

      totalSI += targetSI;
      totalSO += targetSO;
    }
  });
return { bySale: targetsBySale, byRegion: targetsByRegion, totals: { si: totalSI, so: totalSO } };
}

function getStockInReport(filter) {
    const me = requireAuth_(filter.__auth);
    const isManager = me.role === 'khu vực';
const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_STOCK_IN);
    if (!sh || sh.getLastRow() < 2) return { barChartData: { labels: [], datasets: [] }, pieChartData: { labels: [], datasets: [] } };
const allData = sh.getRange(2, 1, sh.getLastRow() - 1, 12).getValues(); 

    const monthFilter = filter.month || fmtYMD_(new Date()).slice(0, 7);
const targets = getTargetsByMonth_(monthFilter);
    const actualsByRegion = new Map();
    let totalActualSI = 0;

    const regionFilter = String(filter.region || '').trim().toLowerCase();
const provinceFilter = String(filter.province || '').trim().toLowerCase();
    const saleFilter = String(filter.sale || '').trim().toLowerCase();
    const svnFilter = String(filter.svn || '').trim().toLowerCase();
const dvnFilter = String(filter.dvn || '').trim().toLowerCase();
    const shopFilter = String(filter.shop || '').trim().toLowerCase();
const filteredData = allData.filter(row => {
        const rowRegion = String(row[6] || '').trim().toLowerCase();
        if (isManager && me.region && rowRegion !== me.region.toLowerCase()) return false;
        if (regionFilter && rowRegion !== regionFilter) return false;
        if (provinceFilter && String(row[7] || '').trim().toLowerCase() !== provinceFilter) return false;
        if (saleFilter && String(row[11] || '').trim().toLowerCase() !== saleFilter) return false;
        if (svnFilter && String(row[4] || '').trim().toLowerCase() !== svnFilter) 
return false;
        if (dvnFilter && String(row[2] || '').trim().toLowerCase() !== dvnFilter) return false;
        if (shopFilter && !String(row[3] || '').trim().toLowerCase().includes(shopFilter)) return false;
        return true;
    });
const twelveMonthsAgo = new Date();
    twelveMonthsAgo.setMonth(twelveMonthsAgo.getMonth() - 11);
    twelveMonthsAgo.setDate(1);
    const monthlyTotals = {};
const monthLabels = Array.from({length: 12}, (_, i) => {
        const d = new Date(twelveMonthsAgo.getFullYear(), twelveMonthsAgo.getMonth() + i, 1);
        const monthKey = Utilities.formatDate(d, Session.getScriptTimeZone(), "yyyy-MM");
        monthlyTotals[monthKey] = 0;
        return monthKey;
    });
filteredData.forEach(row => {
        const entryDate = new Date(row[1]);
        if (entryDate >= twelveMonthsAgo) {
            const monthKey = Utilities.formatDate(entryDate, Session.getScriptTimeZone(), "yyyy-MM");
            if (monthlyTotals.hasOwnProperty(monthKey)) monthlyTotals[monthKey] += Number(row[9] || 0);
        }
    });
const monthData = monthLabels.map(key => monthlyTotals[key]);
    const ratios = [null, ...monthData.slice(1).map((curr, i) => monthData[i] > 0 ? (curr / monthData[i]) * 100 : null)];
const barChartData = { labels: monthLabels, datasets: [{ label: 'Số lượng nhập', data: monthData, ratios: ratios, backgroundColor: 'rgba(255, 159, 64, 0.5)', borderColor: 'rgba(255, 159, 64, 1)', borderWidth: 1 }] };
const [startOfMonth, endOfMonth] = monthRange_(monthFilter);
    const modelTotals = {};
    filteredData.forEach(row => {
        const entryDate = new Date(row[1]);
        const quantity = Number(row[9] || 0);
        const region = String(row[6] || 'Chưa xác định').trim();

        if (entryDate >= startOfMonth && entryDate < endOfMonth) {
            const model = String(row[8] || 'Khác');
            modelTotals[model] = (modelTotals[model] || 
0) + quantity;
            
            const currentRegionActual = actualsByRegion.get(region) || 0;
            actualsByRegion.set(region, currentRegionActual + quantity);
            totalActualSI += quantity;
        }
    });
const pieChartData = { labels: Object.keys(modelTotals), datasets: [{ label: 'Tỉ lệ model', data: Object.values(modelTotals), backgroundColor: ['#FF6384', '#36A2EB', '#FFCE56', '#4BC0C0', '#9966FF', '#FF9F40', '#E7E9ED', '#8DDF3C', '#F54E59', '#4472C4', '#F17F29', '#C55A11'] }]};
let progressData;
    if (regionFilter) {
        const regionName = filter.region;
const target = (targets.byRegion.get(regionName) || {si: 0}).si;
        const actual = actualsByRegion.get(regionName) || 0;
        const percent = target > 0 ?
(actual / target) * 100 : 0;
        progressData = { title: `Tiến độ S.I - Khu vực ${regionName}`, actual, target, percent };
} else {
        const target = targets.totals.si;
const percent = target > 0 ? (totalActualSI / target) * 100 : 0;
progressData = { title: 'Tiến độ S.I Toàn Quốc', actual: totalActualSI, target, percent };
}

    const regionalRanking = [];
    for (const [region, regionTargets] of targets.byRegion.entries()) {
        const actual = actualsByRegion.get(region) ||
0;
        const target = regionTargets.si;
        const percent = target > 0 ? (actual / target) * 100 : 0;
regionalRanking.push({ region, actual, target, percent });
    }
    regionalRanking.sort((a,b) => b.percent - a.percent);
return { barChartData, pieChartData, progressData, regionalRanking };
}

function getSalesOutReport(filter) {
    const me = requireAuth_(filter.__auth);
const isManager = me.role === 'khu vực';
    const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_SALES_OUT);
if (!sh || sh.getLastRow() < 2) return { quantityBarChartData: {}, quantityPieChartData: {}, revenueBarChartData: {}, revenuePieChartData: {} };
const allData = sh.getRange(2, 1, sh.getLastRow() - 1, 14).getValues();

    const monthFilter = filter.month || fmtYMD_(new Date()).slice(0, 7);
const targets = getTargetsByMonth_(monthFilter);
    const actualsByRegion = new Map();
    let totalActualSO = 0;
    
    const regionFilter = String(filter.region || '').trim().toLowerCase();
const provinceFilter = String(filter.province || '').trim().toLowerCase();
    const saleFilter = String(filter.sale || '').trim().toLowerCase();
    const svnFilter = String(filter.svn || '').trim().toLowerCase();
const dvnFilter = String(filter.dvn || '').trim().toLowerCase();
    const shopFilter = String(filter.shop || '').trim().toLowerCase();
const filteredData = allData.filter(row => {
        const rowRegion = String(row[6] || '').trim().toLowerCase();
        if (isManager && me.region && rowRegion !== me.region.toLowerCase()) return false;
        if (regionFilter && rowRegion !== regionFilter) return false;
        if (provinceFilter && String(row[7] || '').trim().toLowerCase() !== provinceFilter) return false;
        if (saleFilter && String(row[13] || '').trim().toLowerCase() !== saleFilter) return false;
        if (svnFilter && String(row[4] || '').trim().toLowerCase() !== svnFilter) 
return false;
        if (dvnFilter && String(row[2] || '').trim().toLowerCase() !== dvnFilter) return false;
        if (shopFilter && !String(row[3] || '').trim().toLowerCase().includes(shopFilter)) return false;
        return true;
    });
const twelveMonthsAgo = new Date();
    twelveMonthsAgo.setMonth(twelveMonthsAgo.getMonth() - 11);
    twelveMonthsAgo.setDate(1);
    const [startOfMonth, endOfMonth] = monthRange_(monthFilter);
    const monthlyQtyTotals = {};
const monthLabels = Array.from({length: 12}, (_, i) => {
        const d = new Date(twelveMonthsAgo.getFullYear(), twelveMonthsAgo.getMonth() + i, 1);
        const monthKey = Utilities.formatDate(d, Session.getScriptTimeZone(), "yyyy-MM");
        monthlyQtyTotals[monthKey] = 0;
        return monthKey;
    });
const modelQtyTotals = {};
    const monthlyRevTotals = {...monthlyQtyTotals};
    const modelRevTotals = {};
filteredData.forEach(row => {
        const entryDate = new Date(row[1]);
        const model = String(row[9] || 'Khác');
        const qty = Number(row[10] || 0);
        const rev = Number(row[11] || 0);
        const region = String(row[6] || 'Chưa xác định').trim();

        if (entryDate >= twelveMonthsAgo) {
            const monthKey = Utilities.formatDate(entryDate, Session.getScriptTimeZone(), "yyyy-MM");
   
if (monthlyQtyTotals.hasOwnProperty(monthKey)) {
                monthlyQtyTotals[monthKey] += qty;
                monthlyRevTotals[monthKey] += rev;
            }
        }
        if (entryDate >= startOfMonth && entryDate < endOfMonth) {
            modelQtyTotals[model] = (modelQtyTotals[model] || 0) 
+ qty;
            modelRevTotals[model] = (modelRevTotals[model] || 0) + rev;

            const currentRegionActual = actualsByRegion.get(region) || 0;
            actualsByRegion.set(region, currentRegionActual + qty);
            totalActualSO += qty;
}
    });

    const PIE_COLORS = ['#FF6384', '#36A2EB', '#FFCE56', '#4BC0C0', '#9966FF', '#FF9F40', '#E7E9ED', '#8DDF3C', '#F54E59', '#4472C4', '#F17F29', '#C55A11'];
const qtyMonthData = monthLabels.map(key => monthlyQtyTotals[key]);
    const qtyRatios = [null, ...qtyMonthData.slice(1).map((curr, i) => qtyMonthData[i] > 0 ? (curr / qtyMonthData[i]) * 100 : null)];
const quantityBarChartData = { labels: monthLabels, datasets: [{ label: 'Số lượng bán', data: qtyMonthData, ratios: qtyRatios, backgroundColor: 'rgba(54, 162, 235, 0.5)', borderColor: 'rgba(54, 162, 235, 1)', borderWidth: 1 }] };
const quantityPieChartData = { labels: Object.keys(modelQtyTotals), datasets: [{ label: 'Tỉ lệ model bán ra', data: Object.values(modelQtyTotals), backgroundColor: PIE_COLORS }] };
const revMonthData = monthLabels.map(key => monthlyRevTotals[key]);
    const revRatios = [null, ...revMonthData.slice(1).map((curr, i) => revMonthData[i] > 0 ? (curr / revMonthData[i]) * 100 : null)];
const revenueBarChartData = { labels: monthLabels, datasets: [{ label: 'Doanh thu (VNĐ)', data: revMonthData, ratios: revRatios, backgroundColor: 'rgba(75, 192, 192, 0.5)', borderColor: 'rgba(75, 192, 192, 1)', borderWidth: 1 }] };
const revenuePieChartData = { labels: Object.keys(modelRevTotals), datasets: [{ label: 'Doanh thu theo model', data: Object.values(modelRevTotals), backgroundColor: PIE_COLORS }] };
let progressData;
    if (regionFilter) {
        const regionName = filter.region;
const target = (targets.byRegion.get(regionName) || {so: 0}).so;
        const actual = actualsByRegion.get(regionName) || 0;
        const percent = target > 0 ?
(actual / target) * 100 : 0;
        progressData = { title: `Tiến độ S.O - Khu vực ${regionName}`, actual, target, percent };
} else {
        const target = targets.totals.so;
const percent = target > 0 ? (totalActualSO / target) * 100 : 0;
progressData = { title: 'Tiến độ S.O Toàn Quốc', actual: totalActualSO, target, percent };
}

    const regionalRanking = [];
    for (const [region, regionTargets] of targets.byRegion.entries()) {
        const actual = actualsByRegion.get(region) ||
0;
        const target = regionTargets.so;
        const percent = target > 0 ? (actual / target) * 100 : 0;
regionalRanking.push({ region, actual, target, percent });
    }
    regionalRanking.sort((a,b) => b.percent - a.percent);
return { quantityBarChartData, quantityPieChartData, revenueBarChartData, revenuePieChartData, progressData, regionalRanking };
}

/*********** TARGET TRACKING FUNCTIONS (REWRITTEN) ***********/
function getTargetProgress(filter) {
  const me = requireAuth_(filter.__auth);
const isManager = me.role === 'khu vực';
  const monthFilter = filter.month || fmtYMD_(new Date()).slice(0, 7);
  const [startOfMonth, endOfMonth] = monthRange_(monthFilter);
const usersSh = SpreadsheetApp.getActive().getSheetByName(SHEET_USERS);
  const userMap = new Map();
  if (usersSh) {
    usersSh.getDataRange().getValues().slice(1).forEach(r => {
      userMap.set(String(r[0]).trim(), { fullName: String(r[2]).trim(), region: String(r[6]).trim() });
    });
}

  const targetsSh = SpreadsheetApp.getActive().getSheetByName(SHEET_TARGETS_SALE);
  if (!targetsSh) throw new Error(`Không tìm thấy sheet '${SHEET_TARGETS_SALE}'`);
  const monthlyTargets = new Map();
targetsSh.getDataRange().getValues().slice(1).forEach(row => {
    if (!row[0]) return; 

    const rowMonth = Utilities.formatDate(new Date(row[0]), Session.getScriptTimeZone(), 'yyyy-MM');
    
    if (rowMonth === monthFilter) {
      const saleUsername = String(row[1]).trim();
      monthlyTargets.set(saleUsername, {
        targetSI: Number(row[3] || 0), // Column D
        targetSO: Number(row[4] || 0)  // Column E
      });
    }
  });
const actuals = new Map();

  const siSh = SpreadsheetApp.getActive().getSheetByName(SHEET_STOCK_IN);
  if (siSh && siSh.getLastRow() > 1) {
    siSh.getDataRange().getValues().slice(1).forEach(row => {
      const entryDate = new Date(row[1]);
      if (entryDate >= startOfMonth && entryDate < endOfMonth) {
        const saleUsername = String(row[11]).trim(); 
        if (!saleUsername) return;
        if (!actuals.has(saleUsername)) actuals.set(saleUsername, { actualSI: 0, actualSO: 0 });
        actuals.get(saleUsername).actualSI += Number(row[9] || 0);
     
}
    });
  }

  const soSh = SpreadsheetApp.getActive().getSheetByName(SHEET_SALES_OUT);
if (soSh && soSh.getLastRow() > 1) {
    soSh.getDataRange().getValues().slice(1).forEach(row => {
      const entryDate = new Date(row[1]);
      if (entryDate >= startOfMonth && entryDate < endOfMonth) {
        const saleUsername = String(row[13]).trim(); 
        if (!saleUsername) return;
        if (!actuals.has(saleUsername)) actuals.set(saleUsername, { actualSI: 0, actualSO: 0 });
        actuals.get(saleUsername).actualSO += Number(row[10] || 0);
      }
    });
}

  let results = [];
  for (const [saleUsername, targets] of monthlyTargets.entries()) {
    const saleInfo = userMap.get(saleUsername) ||
{ fullName: saleUsername, region: '' };
    const saleActuals = actuals.get(saleUsername) || { actualSI: 0, actualSO: 0 };
if (isManager && me.region && saleInfo.region.toLowerCase() !== me.region.toLowerCase()) continue;
    if (filter.region && saleInfo.region.toLowerCase() !== filter.region.toLowerCase()) continue;
if (filter.sale && saleUsername.toLowerCase() !== filter.sale.toLowerCase()) continue;

    const progressSI = targets.targetSI > 0 ?
(saleActuals.actualSI / targets.targetSI) * 100 : 0;
    const progressSO = targets.targetSO > 0 ?
(saleActuals.actualSO / targets.targetSO) * 100 : 0;

    results.push({
      saleUsername: saleInfo.fullName,
      region: saleInfo.region,
      targetSI: targets.targetSI,
      actualSI: saleActuals.actualSI,
      progressSI: progressSI,
      targetSO: targets.targetSO,
      actualSO: saleActuals.actualSO,
      progressSO: progressSO
    });
}
  results.sort((a,b) => a.saleUsername.localeCompare(b.saleUsername));

  const siRanking = [...results].sort((a, b) => b.progressSI - a.progressSI);
const soRanking = [...results].sort((a, b) => b.progressSO - a.progressSO);

  return { ok: true, data: results, siRanking: siRanking, soRanking: soRanking };
}


/*********** ADMIN SET TARGETS FUNCTIONS (NEW SIMPLER VERSION) ***********/
function getSalesByRegion(region) {
  const usersSh = SpreadsheetApp.getActive().getSheetByName(SHEET_USERS);
  if (!usersSh) return [];
const userMap = new Map(usersSh.getDataRange().getValues().slice(1).map(r => [r[2], r[0]])); 

  const shopsSh = SpreadsheetApp.getActive().getSheetByName(SHEET_SHOPS);
  if (!shopsSh) return [];
  const data = shopsSh.getDataRange().getValues().slice(1);
const sales = new Set();
  data.forEach(row => {
    const rowRegion = String(row[5] || '').trim().toLowerCase(); // Column F
    if (rowRegion === region.toLowerCase()) {
      if (row[3]) sales.add(row[3]);
    }
  });
return [...sales].sort().map(fullName => {
    return { username: userMap.get(fullName) || fullName, fullName: fullName };
  });
}

function getExistingSaleTargets(payload) {
    requireAuth_(payload.__auth);
    const { yearMonth } = payload;
if (!yearMonth) return {ok: false, message: 'Invalid month'};

    const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_TARGETS_SALE);
if (!sh) return {ok: false, message: 'Target sheet not found'};

    const data = sh.getDataRange().getValues().slice(1);

    const usersSh = SpreadsheetApp.getActive().getSheetByName(SHEET_USERS);
const userMap = new Map(usersSh ? usersSh.getDataRange().getValues().slice(1).map(r => [r[0], r[2]]) : []);
const targets = data
      .filter(r => {
          if (!r[0]) return false;
          // [FIX 1] So sánh ngày tháng sau khi đã định dạng
          const rowMonth = Utilities.formatDate(new Date(r[0]), Session.getScriptTimeZone(), 'yyyy-MM');
          return rowMonth === yearMonth;
      })
      .map(r => {
          const saleUsername 
= String(r[1]);
          return {
              username: saleUsername,
              fullName: userMap.get(saleUsername) || saleUsername,
              targetSI: r[3] || 0,
              targetSO: r[4] || 0
          };
      });
return { ok: true, data: targets };
}

function saveSingleSaleTarget(payload) {
  const me = requireAuth_(payload.__auth);
if (me.role !== 'admin') {
    throw new Error('Bạn không có quyền thực hiện chức năng này.');
}

  const { yearMonth, saleUsername, targetSI, targetSO } = payload;
if (!yearMonth || !saleUsername) {
    throw new Error('Dữ liệu không hợp lệ.');
}

  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_TARGETS_SALE);
  if (!sh) throw new Error(`Không tìm thấy sheet '${SHEET_TARGETS_SALE}'.`);
  
  const allData = sh.getDataRange().getValues();
let foundRow = -1;

  for (let i = 1; i < allData.length; i++) {
    if (allData[i][0] && Utilities.formatDate(new Date(allData[i][0]), Session.getScriptTimeZone(), 'yyyy-MM') == yearMonth && allData[i][1] == saleUsername) {
      foundRow = i + 1;
break;
    }
  }
  
  const usersSh = SpreadsheetApp.getActive().getSheetByName(SHEET_USERS);
const userMap = new Map(usersSh ? usersSh.getDataRange().getValues().slice(1).map(r => [r[0], r[2]]) : []);
  const fullName = userMap.get(saleUsername) || saleUsername;
let message = '';

  if (Number(targetSI) === 0 && Number(targetSO) === 0) {
    if (foundRow > -1) {
      sh.deleteRow(foundRow);
message = `Đã xóa target của ${fullName} cho tháng ${yearMonth}.`;
} else {
      message = 'Không có thay đổi.';
}
  } else {
    const newRowData = [yearMonth, saleUsername, fullName, targetSI, targetSO];
if (foundRow > -1) {
      sh.getRange(foundRow, 1, 1, 5).setValues([newRowData]);
message = `Đã cập nhật target của ${fullName} cho tháng ${yearMonth}.`;
    } else {
      sh.appendRow(newRowData);
message = `Đã thêm target của ${fullName} cho tháng ${yearMonth}.`;
    }
  }

  const lastRow = sh.getLastRow();
if (lastRow > 1) {
    const range = sh.getRange(2, 1, lastRow - 1, sh.getLastColumn());
range.sort([{column: 1, ascending: true}, {column: 3, ascending: true}]); 
  }
  
  return { message: message };
}
