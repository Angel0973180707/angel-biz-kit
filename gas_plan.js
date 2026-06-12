/**
 * 天使智慧商務 · 需求管理 GAS
 * =============================================
 * 使用前設定：
 * 1. 建立新的 Google 試算表
 * 2. 新增一個分頁，名稱：客戶需求
 *    A1: 需求編號  B1: 姓名  C1: 手機  D1: 業務類型
 *    E1: 選擇功能  F1: 預算  G1: 估算報價  H1: 說明
 *    I1: 狀態  J1: 管理員備註  K1: 建立時間
 * 3. 建立新 GAS 專案，貼上此程式碼
 * 4. 部署為 Web 應用程式（Anyone can access）
 * 5. 複製 URL 填入 plan.html 和 plan-admin.html 的 GAS_URL
 * =============================================
 */

const SPREADSHEET_ID = 'YOUR_SPREADSHEET_ID_HERE'; // 填入你的試算表 ID
const ADMIN_KEY      = 'ANGEL20261972';
const LINE_TOKEN     = 'YOUR_LINE_BOT_TOKEN';       // LINE Bot Channel Access Token
const LINE_USER_ID   = 'YOUR_LINE_USER_ID';          // Angel 的 LINE User ID

const SHEET_NAME = '客戶需求';

// ── GET 路由 ──
function doGet(e) {
  const p = e.parameter || {};
  const action = p.action || '';
  const res = d => ContentService.createTextOutput(JSON.stringify(d))
    .setMimeType(ContentService.MimeType.JSON);

  try {
    if (action === 'getPlans') {
      if (p.key !== ADMIN_KEY) return res({ ok:false, error:'驗證失敗' });
      return res(getPlans(p));
    }
    if (action === 'updatePlan') {
      if (p.key !== ADMIN_KEY) return res({ ok:false, error:'驗證失敗' });
      return res(updatePlan(p));
    }
    if (action === 'submitPlan') {
      return res(submitPlan(p));
    }
    return res({ ok:false, error:'未知 action' });
  } catch(err) {
    return res({ ok:false, error: err.message });
  }
}

function doPost(e) {
  const p = e.parameter || {};
  return doGet({ parameter: p });
}

// ── 送出需求（客戶用）──
function submitPlan(p) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) return { ok:false, error:'找不到工作表：'+SHEET_NAME };

  const id = 'REQ' + Date.now();
  const now = new Date().toLocaleString('zh-TW', { timeZone:'Asia/Taipei' });

  sh.appendRow([
    id,
    p.name || '',
    p.phone || '',
    p.business_type || '',
    p.features || '',
    p.budget || '',
    p.price_range || '',
    p.note || '',
    '新需求',   // 狀態
    '',          // 管理員備註
    now          // 建立時間
  ]);

  // LINE 通知 Angel
  const msg = `🔔 新的客戶需求！\n\n` +
    `姓名：${p.name}\n` +
    `手機：${p.phone}\n` +
    `業務：${p.business_type}\n` +
    `功能：${p.features || '未選擇'}\n` +
    `預算：${p.budget || '不確定'}\n` +
    `估價：${p.price_range || '—'}\n` +
    (p.note ? `說明：${p.note}\n` : '') +
    `\n請到後台查看詳情 📋`;
  sendLineMsg(msg);

  return { ok:true, id };
}

// ── 查詢全部需求（後台用）──
function getPlans(p) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) return { ok:true, data:[] };

  const rows = sh.getDataRange().getValues();
  if (rows.length <= 1) return { ok:true, data:[] };

  const data = rows.slice(1).filter(r => r[0]).map(r => ({
    id:            r[0],
    name:          r[1],
    phone:         r[2],
    business_type: r[3],
    features:      r[4],
    budget:        r[5],
    price_range:   r[6],
    note:          r[7],
    status:        r[8] || '新需求',
    admin_note:    r[9] || '',
    created_at:    r[10] ? String(r[10]).slice(0,10) : ''
  }));

  // 排序：新需求在前，按時間倒序
  data.sort((a,b) => {
    const order = { '新需求':0, '已聯繫':1, '已成交':2, '未成交':3 };
    if (order[a.status] !== order[b.status]) return order[a.status]-order[b.status];
    return b.created_at.localeCompare(a.created_at);
  });

  return { ok:true, data };
}

// ── 更新需求狀態/備註（後台用）──
function updatePlan(p) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) return { ok:false, error:'找不到工作表' };

  const rows = sh.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === p.id) {
      if (p.status)     sh.getRange(i+1, 9).setValue(p.status);
      if (p.admin_note !== undefined) sh.getRange(i+1, 10).setValue(p.admin_note);
      return { ok:true };
    }
  }
  return { ok:false, error:'找不到需求 ID：'+p.id };
}

// ── LINE 推播 ──
function sendLineMsg(text) {
  if (!LINE_TOKEN || LINE_TOKEN === 'YOUR_LINE_BOT_TOKEN') return;
  try {
    UrlFetchApp.fetch('https://api.line.me/v2/bot/message/push', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Authorization': 'Bearer ' + LINE_TOKEN
      },
      payload: JSON.stringify({
        to: LINE_USER_ID,
        messages: [{ type:'text', text }]
      }),
      muteHttpExceptions: true
    });
  } catch(e) {}
}
