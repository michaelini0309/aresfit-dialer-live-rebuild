const CONFIG = Object.freeze({
  SPREADSHEET_ID: '1b1xaTfmWtlTlgx1nb7UJvBI2yx1FCKNpOWFgt9_RwKI',
  USERS_SHEET: 'Users',
  REGISTRY_SHEET: 'Updated Global Registry',
  SMOKE_SHEET: 'App Smoke Test',
  HEADER_ROW: 3,
  DEFAULT_LIMIT: 10,
  MAX_LIMIT: 50
});

function doGet(e) {
  const p = (e && e.parameter) || {};
  const callback = safeCallback_(p.callback);
  try {
    const action = String(p.action || '').trim();
    if (!action) return output_(callback, {ok:false,error:'Missing action'});
    if (action === 'health') return output_(callback, {ok:true,service:'AresFit live registry',time:isoNow_()});

    const user = findActiveUser_(p.email);
    if (!user) return output_(callback, {ok:false,error:'Email is not an active AresFit user'});

    if (action === 'smoke_test') {
      const sheet = getSheet_(CONFIG.SMOKE_SHEET);
      const ts = new Date();
      sheet.appendRow([
        ts,
        user.rep_id,
        user.email,
        'LIVE_APP_SMOKE_TEST',
        String(p.value || '').slice(0,500),
        String(p.app_build || '').slice(0,100)
      ]);
      return output_(callback, {ok:true,user:user,timestamp:Utilities.formatDate(ts,'Europe/London',"yyyy-MM-dd'T'HH:mm:ssXXX")});
    }

    if (action === 'get_queue') {
      const limit = Math.min(CONFIG.MAX_LIMIT, Math.max(1, Number(p.limit) || CONFIG.DEFAULT_LIMIT));
      const result = getAssignedQueue_(user, limit);
      return output_(callback, {ok:true,user:user,total:result.total,leads:result.leads});
    }

    return output_(callback, {ok:false,error:'Unknown action'});
  } catch (err) {
    return output_(callback, {ok:false,error:String(err && err.message ? err.message : err)});
  }
}

function findActiveUser_(email) {
  const target = String(email || '').trim().toLowerCase();
  if (!target) return null;
  const sheet = getSheet_(CONFIG.USERS_SHEET);
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return null;
  const h = headerIndex_(values[0]);
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    if (String(row[h.Email] || '').trim().toLowerCase() !== target) continue;
    const active = row[h.Active] === true || String(row[h.Active] || '').toUpperCase() === 'TRUE';
    if (!active) return null;
    return {
      rep_id: String(row[h.Rep_ID] || '').trim(),
      name: String(row[h.Name] || '').trim(),
      email: String(row[h.Email] || '').trim(),
      role: String(row[h.Role] || '').trim(),
      sender_name: String(row[h.Sender_Name] || '').trim(),
      sender_email: String(row[h.Sender_Email] || '').trim()
    };
  }
  return null;
}

function getAssignedQueue_(user, limit) {
  const sheet = getSheet_(CONFIG.REGISTRY_SHEET);
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow <= CONFIG.HEADER_ROW) return {total:0,leads:[]};
  const values = sheet.getRange(CONFIG.HEADER_ROW, 1, lastRow - CONFIG.HEADER_ROW + 1, lastCol).getDisplayValues();
  const headers = values[0];
  const h = headerIndex_(headers);
  const ownerName = normalise_(user.name);
  let total = 0;
  const leads = [];
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    if (normalise_(row[h.Current_Owner]) !== ownerName) continue;
    const suppression = normalise_(row[h.Suppression_Status]);
    const callability = normalise_(row[h.Current_Callability_State]);
    const allowedSuppression = ['', 'NO', 'CLEAR', 'NONE'];
    if (!allowedSuppression.includes(suppression)) continue;
    total++;
    if (leads.length >= limit) continue;
    leads.push({
      global_site_id: value_(row,h,'Global_Site_ID'),
      global_account_id: value_(row,h,'Global_Account_ID'),
      business_name: value_(row,h,'Business_Name'),
      phone: value_(row,h,'Phone'),
      email: value_(row,h,'Email'),
      decision_maker: value_(row,h,'Decision_Maker'),
      lifecycle_status: value_(row,h,'Current_Lifecycle_Status'),
      activity_status: value_(row,h,'Current_Activity_Status'),
      callability: value_(row,h,'Current_Callability_State'),
      latest_notes: value_(row,h,'Latest_Real_Notes'),
      callback_date: value_(row,h,'Callback_Date'),
      follow_up_date: value_(row,h,'Follow_Up_Date'),
      latest_activity_date: value_(row,h,'Latest_Activity_Date')
    });
  }
  return {total:total,leads:leads};
}

function getSheet_(name) {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sheet = ss.getSheetByName(name);
  if (!sheet) throw new Error('Missing sheet: ' + name);
  return sheet;
}

function headerIndex_(headers) {
  const map = {};
  headers.forEach(function(v,i){ map[String(v || '').trim()] = i; });
  return map;
}

function value_(row,h,key) {
  const idx = h[key];
  return idx == null ? '' : String(row[idx] || '').trim();
}

function normalise_(v) { return String(v || '').trim().toUpperCase(); }
function isoNow_() { return Utilities.formatDate(new Date(),'Europe/London',"yyyy-MM-dd'T'HH:mm:ssXXX"); }

function safeCallback_(name) {
  const cb = String(name || '').trim();
  return /^[A-Za-z_$][0-9A-Za-z_$\.]{0,120}$/.test(cb) ? cb : '';
}

function output_(callback, payload) {
  const json = JSON.stringify(payload);
  if (callback) {
    return ContentService.createTextOutput(callback + '(' + json + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}
