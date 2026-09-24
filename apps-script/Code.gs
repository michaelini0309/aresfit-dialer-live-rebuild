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
    if (String(p.app || '') === 'live-dialer') {
      return HtmlService.createHtmlOutputFromFile('LiveDialer')
        .setTitle('AresFit Live Registry Dialer');
    }
    const action = String(p.action || '').trim();
    if (!action) return output_(callback, {ok:false,error:'Missing action'});
    if (action === 'health') return output_(callback, {ok:true,service:'AresFit live registry',time:isoNow_()});

    const user = requireSignedInAresFitUser_();

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

// These server functions are intended to be called by google.script.run from
// the HtmlService app. Identity is taken from Google's authenticated session,
// never from an email string supplied by the browser.
function getQueueForSignedInUser(afterSiteId) {
  const user = requireSignedInAresFitUser_();
  const cache = CacheService.getScriptCache();
  const cacheKey = 'ARES_QUEUE_' + user.rep_id;
  let ids, offset = 0;
  if (afterSiteId) {
    offset = Number(afterSiteId);
    if (!Number.isInteger(offset) || offset < 0) throw new Error('Queue page token is invalid. Refresh the queue.');
    const cached = cache.get(cacheKey);
    if (!cached) throw new Error('Queue snapshot expired. Refresh the queue.');
    ids = JSON.parse(cached);
  }
  const result = getAssignedQueue_(user, CONFIG.MAX_LIMIT, ids || null, offset);
  if (!ids) cache.put(cacheKey, JSON.stringify(result.queue_ids), 21600);
  return {ok:true, user:{name:user.name, email:user.email, rep_id:user.rep_id, sender_name:user.sender_name, sender_email:user.sender_email}, total:result.total, leads:result.leads, next_cursor:result.next_cursor};
}

function logActivityForSignedInUser(request) {
  const user = requireSignedInAresFitUser_();
  const p = request || {};
  const siteId = String(p.global_site_id || '').trim();
  const eventId = String(p.event_id || '').trim();
  const outcome = String(p.outcome || '').trim();
  const channel = String(p.channel || 'call').trim();
  const note = String(p.note_text || '').trim();
  const storedNote = 'AresFit dialer event: ' + note;
  const allowedOutcomes = ['reached','no answer','voicemail','dead air','email sent','reply received','quote sent','quote needed','supplier needed','customer waiting','parked','lost','do not touch','HOLD','app issue'];
  const allowedChannels = ['call','email','WhatsApp','quote','supplier','Shopify','blocker','app issue'];
  if (!siteId || siteId.length > 80) throw new Error('A valid Global_Site_ID is required.');
  if (!/^[-A-Za-z0-9_:]{8,120}$/.test(eventId)) throw new Error('A unique event ID is required.');
  if (!allowedOutcomes.includes(outcome)) throw new Error('Choose a supported activity outcome.');
  if (!allowedChannels.includes(channel)) throw new Error('Choose a supported activity type.');
  if (!note || note.length > 2000) throw new Error('A call note of 1–2,000 characters is required.');
  const nextActionDate = String(p.next_action_date || '').trim();
  if (nextActionDate && !/^\d{4}-\d{2}-\d{2}(?:[ T]\d{2}:\d{2})?$/.test(nextActionDate)) throw new Error('Follow-up date must be YYYY-MM-DD or YYYY-MM-DD HH:MM.');
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const registry = getSheet_(CONFIG.REGISTRY_SHEET);
    const lastRow = registry.getLastRow();
    const lastCol = registry.getLastColumn();
    if (lastRow <= CONFIG.HEADER_ROW) throw new Error('Registry has no lead rows.');
    const rows = registry.getRange(CONFIG.HEADER_ROW, 1, lastRow - CONFIG.HEADER_ROW + 1, lastCol).getDisplayValues();
    const h = headerIndex_(rows[0]);
    ['Global_Site_ID','Current_Owner','Suppression_Status','Business_Name'].forEach(k => {
      if (h[k] == null) throw new Error('Registry is missing required column: ' + k);
    });
    const matches = [];
    for (let i=1; i<rows.length; i++) if (String(rows[i][h.Global_Site_ID]).trim() === siteId) matches.push({row:i+CONFIG.HEADER_ROW, values:rows[i]});
    if (matches.length !== 1) throw new Error(matches.length ? 'Global_Site_ID is not unique; event held.' : 'Lead was not found.');
    const lead = matches[0];
    if (normalise_(lead.values[h.Current_Owner]) !== normalise_(user.name)) throw new Error('Lead is not assigned to the signed-in rep.');
    const suppression = normalise_(lead.values[h.Suppression_Status]);
    if (!['','NO','CLEAR','NONE'].includes(suppression)) throw new Error('Lead is suppressed; event held.');

    const history = getSheet_('Activity History');
    const headerRow = 3;
    const historyLastCol = history.getLastColumn();
    const hh = headerIndex_(history.getRange(headerRow,1,1,historyLastCol).getDisplayValues()[0]);
    const required = ['Event_ID','Global_Site_ID','Business_Name','Rep','Event_Date','Event_Time','Date_Confidence','Event_Type','Outcome','Note_Text','Status_After','Stage_After','Follow_Up_Date','Match_Method','First_Source','Latest_Source','Source_Count','Source_SHA256s','Source_Observed_Date','Context_Only','Is_Latest_Event'];
    required.forEach(k => { if (hh[k] == null) throw new Error('Activity History is missing required column: ' + k); });
    const oldLastRow = history.getLastRow();
    const oldRows = oldLastRow >= 4 ? history.getRange(4,1,oldLastRow-3,historyLastCol).getDisplayValues() : [];
    const demoteRows = [];
    let duplicateRow = 0;
    for (let i=0; i<oldRows.length; i++) {
      if (String(oldRows[i][hh.Event_ID]).trim() === eventId) {
        if (String(oldRows[i][hh.Global_Site_ID]).trim() !== siteId) throw new Error('Event ID already belongs to another lead.');
        if (String(oldRows[i][hh.Outcome]).trim() !== outcome || String(oldRows[i][hh.Note_Text]).trim() !== storedNote || String(oldRows[i][hh.Follow_Up_Date] || '').trim() !== nextActionDate) throw new Error('Event ID was already used with different activity details.');
        duplicateRow = i+4;
      }
      if (String(oldRows[i][hh.Global_Site_ID]).trim() === siteId && normalise_(oldRows[i][hh.Is_Latest_Event]) === 'YES') demoteRows.push(i+4);
    }

    if (duplicateRow) {
      demoteRows.forEach(r => { if (r !== duplicateRow) history.getRange(r,hh.Is_Latest_Event+1).setValue('NO'); });
      if (normalise_(oldRows[duplicateRow-4][hh.Is_Latest_Event]) !== 'YES') history.getRange(duplicateRow,hh.Is_Latest_Event+1).setValue('YES');
      persistLatestLeadState_(registry,lead.row,h,historyRowLatestNote_(dateKey_(oldRows[duplicateRow-4][hh.Event_Date]),String(oldRows[duplicateRow-4][hh.Event_Time]||''),outcome,storedNote),String(oldRows[duplicateRow-4][hh.Event_Date]||''),String(oldRows[duplicateRow-4][hh.Follow_Up_Date]||''),outcome,channel,user.name);
      return {ok:true, duplicate:true, event_id:eventId, global_site_id:siteId};
    }

    const now = new Date();
    const date = Utilities.formatDate(now,'Europe/London','yyyy-MM-dd');
    const time = Utilities.formatDate(now,'Europe/London','HH:mm:ss');
    const values = new Array(historyLastCol).fill('');
    const set = (k,v) => { values[hh[k]] = v; };
    set('Event_ID',eventId); set('Global_Site_ID',siteId);
    set('Business_Name',lead.values[h.Business_Name]); set('Rep',user.name);
    set('Event_Date',date); set('Event_Time',time); set('Date_Confidence','APP_ACTION_SAVED_AT');
    set('Event_Type',channel === 'call' ? 'CALL' : channel.toUpperCase()); set('Outcome',outcome); set('Note_Text',storedNote);
    if (nextActionDate) set('Follow_Up_Date',nextActionDate);
    set('Match_Method','GLOBAL_ID_EXACT'); set('First_Source','AresFit Live Dialer');
    set('Latest_Source','AresFit Live Dialer'); set('Source_Count',1);
    set('Source_Observed_Date',date); set('Context_Only','NO'); set('Is_Latest_Event','YES');
    const appendRow = Math.max(4,history.getLastRow()+1);
    // The reconciled history may fill every allocated row. Extend the grid
    // before either formatting or writing the new event.
    if (appendRow > history.getMaxRows()) {
      history.insertRowsAfter(history.getMaxRows(), Math.max(100,appendRow-history.getMaxRows()));
    }
    if (appendRow > 4) history.getRange(appendRow-1,1,1,historyLastCol)
      .copyTo(history.getRange(appendRow,1,1,historyLastCol),SpreadsheetApp.CopyPasteType.PASTE_FORMAT,false);
    // Append the durable event first. A repeated client retry is safe because
    // Event_ID is checked under the script lock before this point.
    history.getRange(appendRow,1,1,historyLastCol).setValues([values]);
    demoteRows.forEach(r => history.getRange(r,hh.Is_Latest_Event+1).setValue('NO'));
    persistLatestLeadState_(registry,lead.row,h,historyRowLatestNote_(date,time,outcome,storedNote),date,nextActionDate,outcome,channel,user.name);
    return {ok:true, duplicate:false, event_id:eventId, global_site_id:siteId, timestamp:isoNow_()};
  } finally {
    lock.releaseLock();
  }
}

function historyRowLatestNote_(date,time,outcome,note) {
  return [date, time, outcome, note].filter(Boolean).join(' · ');
}

function persistLatestLeadState_(registry,rowNumber,h,latestNote,eventDate,nextActionDate,outcome,channel,repName) {
  const setIfPresent = (name,value) => { if (h[name] != null) registry.getRange(rowNumber,h[name]+1).setValue(value); };
  setIfPresent('Latest_Real_Notes',latestNote);
  setIfPresent('Latest_Activity_Date',eventDate);
  setIfPresent('Latest_Activity_Type',channel === 'call' ? 'CALL' : channel.toUpperCase());
  setIfPresent('Latest_Activity_Outcome',outcome);
  setIfPresent('Latest_Note_Date',eventDate);
  setIfPresent('Last_Touched_By',repName);
  setIfPresent('Last_Touched_Date',eventDate);
  setIfPresent('Last_Registry_Update',isoNow_());
  if (outcome === 'do not touch') {
    setIfPresent('Suppression_Status','YES');
    setIfPresent('Current_Callability_State','NOT_CALLABLE');
  }
  if (nextActionDate) {
    setIfPresent('Follow_Up_Date',nextActionDate);
    if (outcome === 'customer waiting') setIfPresent('Callback_Date',nextActionDate);
  }
}

function requireSignedInAresFitUser_() {
  const email = String(Session.getActiveUser().getEmail() || '').trim();
  if (!email) throw new Error('Google could not verify the signed-in identity. Open the AresFit app with your business account.');
  const user = findActiveUser_(email);
  if (!user) throw new Error('Signed-in Google account is not an active AresFit user.');
  return user;
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

function getAssignedQueue_(user, limit, snapshotIds, offset) {
  const sheet = getSheet_(CONFIG.REGISTRY_SHEET);
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow <= CONFIG.HEADER_ROW) return {total:0,leads:[],next_cursor:'',queue_ids:[]};
  const values = sheet.getRange(CONFIG.HEADER_ROW, 1, lastRow - CONFIG.HEADER_ROW + 1, lastCol).getDisplayValues();
  const headers = values[0];
  const h = headerIndex_(headers);
  ['Global_Site_ID','Current_Owner','Suppression_Status','Business_Name'].forEach(k => {
    if (h[k] == null) throw new Error('Registry is missing required column: ' + k);
  });
  const ownerName = normalise_(user.name);
  const all = [];
  const byId = {};
  const today = Utilities.formatDate(new Date(),'Europe/London','yyyy-MM-dd');
  const stoppedIds = getLatestStoppedSiteIds_();
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    if (normalise_(row[h.Current_Owner]) !== ownerName) continue;
    const suppression = normalise_(row[h.Suppression_Status]);
    const allowedSuppression = ['', 'NO', 'CLEAR', 'NONE'];
    if (!allowedSuppression.includes(suppression)) continue;
    const lead = {
      global_site_id: value_(row,h,'Global_Site_ID'),
      global_account_id: value_(row,h,'Global_Account_ID'),
      business_name: value_(row,h,'Business_Name'),
      phone: value_(row,h,'Phone'),
      email: value_(row,h,'Email'),
      decision_maker: value_(row,h,'Decision_Maker'),
      lifecycle_status: value_(row,h,'Current_Lifecycle_Status'),
      activity_status: value_(row,h,'Current_Activity_Status'),
      latest_outcome: value_(row,h,'Latest_Activity_Outcome'),
      closed_lost: value_(row,h,'Closed_Lost_Status'),
      callability: value_(row,h,'Current_Callability_State'),
      latest_notes: value_(row,h,'Latest_Real_Notes'),
      callback_date: value_(row,h,'Callback_Date'),
      follow_up_date: value_(row,h,'Follow_Up_Date'),
      latest_activity_date: value_(row,h,'Latest_Activity_Date')
    };
    if (!lead.global_site_id) continue;
    if (stoppedIds[lead.global_site_id] || normalise_(lead.callability) === 'NOT_CALLABLE') continue;
    if (normalise_(lead.lifecycle_status) === 'CLOSED_LOST' || normalise_(lead.closed_lost) === 'YES') continue;
    if (['LOST','NOT INTERESTED','NOT INT.','PERMANENT CLOSURE CONFIRMED'].includes(normalise_(lead.latest_outcome)) || normalise_(lead.activity_status) === 'PERMANENT CLOSURE CONFIRMED') continue;
    const callback = dateKey_(lead.callback_date);
    const followup = dateKey_(lead.follow_up_date);
    const callability = normalise_(lead.callability);
    // Rank explicit due dates first. Registry callability labels are shown as
    // historical evidence; this ranking does not certify TPS/CTPS compliance.
    if (callback && callback <= today) lead._rank = callback < today ? 0 : 1;
    else if (followup && followup <= today) lead._rank = followup < today ? 2 : 3;
    else if (['CALL_READY_SUPPORTED_BY_HISTORY','CALL_READY_SUPPORTED_BY_ACTIVITY'].includes(callability)) lead._rank = 4;
    else if (callback || followup) lead._rank = 5;
    else lead._rank = 6;
    lead._sortDate = callback && callback <= today ? callback : (followup || callback || '9999-99-99');
    all.push(lead); byId[lead.global_site_id]=lead;
  }
  all.sort((a,b) => a._rank-b._rank || a._sortDate.localeCompare(b._sortDate) || a.global_site_id.localeCompare(b.global_site_id));
  const queueIds = snapshotIds || all.map(l => l.global_site_id);
  const total = queueIds.length;
  const start = Number(offset)||0;
  const leads = queueIds.slice(start,start+limit).map(id=>byId[id]).filter(Boolean);
  leads.forEach(l => { delete l._rank; delete l._sortDate; });
  return {total:total,leads:leads,next_cursor:start+limit<total?String(start+limit):'',queue_ids:queueIds};
}

function getLatestStoppedSiteIds_() {
  const sheet = getSheet_('Activity History');
  const lastRow = sheet.getLastRow();
  if (lastRow < 4) return {};
  const lastCol = sheet.getLastColumn();
  const rows = sheet.getRange(3,1,lastRow-2,lastCol).getDisplayValues();
  const h = headerIndex_(rows[0]);
  if (h.Global_Site_ID == null || h.Outcome == null || h.Is_Latest_Event == null) return {};
  const stopped = {};
  for (let i=1; i<rows.length; i++) {
    const row=rows[i];
    if (normalise_(row[h.Is_Latest_Event]) !== 'YES') continue;
    const outcome=normalise_(row[h.Outcome]);
    if (outcome === 'LOST' || outcome === 'DO NOT TOUCH') stopped[String(row[h.Global_Site_ID]).trim()] = true;
  }
  return stopped;
}

function dateKey_(value) {
  const s = String(value || '').trim();
  if (!s) return '';
  let m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (m) return m[1]+'-'+('0'+m[2]).slice(-2)+'-'+('0'+m[3]).slice(-2);
  m = s.match(/^(\d{1,2})[/.](\d{1,2})[/.](\d{4})/);
  if (m) return m[3]+'-'+('0'+m[2]).slice(-2)+'-'+('0'+m[1]).slice(-2);
  return '';
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
