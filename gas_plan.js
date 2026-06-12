const SPREADSHEET_ID = '1n2W0KYk6RHA5lsIYDg_avsKtiS9tUnbJKeOVnswIqOw';
const ADMIN_KEY      = 'ANGEL20261972';
const LINE_TOKEN     = 'SGl/bUCnFz3NpOQJJKJTU+zTKgkqtIfdAKE1FM4v6Eu6KKm8i+MmbXsegjW3ef8WLBxNzoIx6oZfh67alrl5OUTdyPezUDiVTz7nbLTwbLESCzzTAQnxcRuwBaKihcgUT1z+ZtQ7Z8QFVOJQhJ4VRAdB04t89/1O/w1cDnyilFU=';
const LINE_USER_ID   = 'U045fa7302eac96a6d54261d38a67f1b7';
const SHEET_NAME     = '客戶需求';

// 修復1：安全解碼中文（防止 double-decode 出錯）
function sd(s) {
  if (!s) return '';
  try { return decodeURIComponent(s); } catch(e) { return String(s); }
}

function doGet(e) {
  const p = e.parameter || {};
  const action = p.action || '';
  const res = d => ContentService.createTextOutput(JSON.stringify(d))
    .setMimeType(ContentService.MimeType.JSON);
  try {
    if (action === 'getPlans')   { if (p.key !== ADMIN_KEY) return res({ok:false,error:'驗證失敗'}); return res(getPlans()); }
    if (action === 'updatePlan') { if (p.key !== ADMIN_KEY) return res({ok:false,error:'驗證失敗'}); return res(updatePlan(p)); }
    if (action === 'submitPlan') return res(submitPlan(p));
    return res({ok:false, error:'未知 action'});
  } catch(err) { return res({ok:false, error:err.message}); }
}

function doPost(e) { return doGet({parameter: e.parameter||{}}); }

function submitPlan(p) {
  const sh = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
  if (!sh) return {ok:false, error:'找不到工作表：'+SHEET_NAME};

  // 修復1：全部欄位套用 sd() 解碼中文
  const name          = sd(p.name);
  const phone         = sd(p.phone);
  const business_type = sd(p.business_type);
  const features      = sd(p.features);
  const budget        = sd(p.budget);
  const price_range   = sd(p.price_range);
  const note          = sd(p.note);

  const id  = 'REQ' + Date.now();
  const now = new Date().toLocaleString('zh-TW', {timeZone:'Asia/Taipei'});

  sh.appendRow([id, name, phone, business_type, features, budget, price_range, note, '新需求', '', now]);

  sendLineMsg(
    `🔔 新客戶需求！\n\n姓名：${name}\n手機：${phone}\n業務：${business_type}\n` +
    `功能：${features||'未選擇'}\n預算：${budget||'不確定'}\n估價：${price_range||'—'}\n` +
    (note ? `說明：${note}\n` : '') + `\n請到後台查看 📋`
  );
  return {ok:true, id};
}

function getPlans() {
  const sh = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
  if (!sh) return {ok:true, data:[]};
  const rows = sh.getDataRange().getValues();
  if (rows.length <= 2) return {ok:true, data:[]};

  const data = rows.slice(2).filter(r => r[0]).map(r => ({
    id:            String(r[0]),
    name:          String(r[1]||''),
    phone:         String(r[2]||''),
    business_type: String(r[3]||''),
    features:      String(r[4]||''),
    budget:        String(r[5]||''),
    price_range:   String(r[6]||''),
    note:          String(r[7]||''),
    status:        String(r[8]||'').trim() || '新需求',  // 修復2：trim() 防空白
    admin_note:    String(r[9]||''),
    created_at:    r[10] ? String(r[10]).slice(0,10) : ''
  }));

  data.sort((a,b) => {
    const o = {'新需求':0,'已聯繫':1,'已成交':2,'未成交':3};
    const oa = o[a.status]??9, ob = o[b.status]??9;
    return oa !== ob ? oa-ob : b.created_at.localeCompare(a.created_at);
  });
  return {ok:true, data};
}

function updatePlan(p) {
  const sh = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
  if (!sh) return {ok:false, error:'找不到工作表'};
  const rows = sh.getDataRange().getValues();
  for (let i = 2; i < rows.length; i++) {
    if (String(rows[i][0]) === String(p.id)) {
      if (p.status)                sh.getRange(i+1, 9).setValue(p.status.trim());
      if (p.admin_note !== undefined) sh.getRange(i+1, 10).setValue(sd(p.admin_note));
      return {ok:true};
    }
  }
  return {ok:false, error:'找不到：'+p.id};
}

function sendLineMsg(text) {
  try {
    UrlFetchApp.fetch('https://api.line.me/v2/bot/message/push', {
      method:'POST',
      headers:{'Content-Type':'application/json','Authorization':'Bearer '+LINE_TOKEN},
      payload:JSON.stringify({to:LINE_USER_ID, messages:[{type:'text',text}]}),
      muteHttpExceptions:true
    });
  } catch(e) {}
}
