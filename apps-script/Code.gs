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
        .addMetaTag('viewport','width=device-width, initial-scale=1, viewport-fit=cover')
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
function getQueueForSignedInUser(cursor, dueOnly) {
  const user = requireSignedInAresFitUser_();
  const cache = CacheService.getScriptCache();
  const token = cursor || Utilities.getUuid();
  if (!/^[-a-zA-Z0-9]{20,80}$/.test(token)) throw new Error('Queue session is invalid. Reconnect to the registry.');
  const cacheKey = 'ARES_QUEUE_' + user.rep_id + '_' + token;
  let served = [];
  if (cursor) {
    const saved = cache.get(cacheKey);
    if (!saved) throw new Error('Queue session expired. Reconnect to the registry.');
    served = JSON.parse(saved);
  }
  const result = getAssignedQueue_(user, CONFIG.MAX_LIMIT, served, dueOnly === true);
  const sentIds = served.concat(result.leads.map(l => l.global_site_id));
  cache.put(cacheKey, JSON.stringify(sentIds), 21600);
  return {ok:true, user:{name:user.name, email:user.email, rep_id:user.rep_id, sender_name:user.sender_name, sender_email:user.sender_email}, total:result.total, leads:result.leads, remaining:result.remaining, next_cursor:result.remaining ? token : '', queue_cursor:token, next_due_at:result.next_due_at};
}

function logActivityForSignedInUser(request) {
  const user = requireSignedInAresFitUser_();
  const p = request || {};
  const siteId = String(p.global_site_id || '').trim();
  const eventId = String(p.event_id || '').trim();
  const outcome = String(p.outcome || '').trim();
  const channel = String(p.channel || 'call').trim();
  const note = String(p.note_text || '').trim();
  const storedNote = note;
  const nextAction = normalise_(p.next_action);
  const allowedOutcomes = ['reached','no answer','voicemail','dead air','email sent','reply received','quote sent','quote needed','supplier needed','customer waiting','parked','lost','do not touch','HOLD','app issue'];
  const allowedChannels = ['call','email','WhatsApp','quote','supplier','Shopify','blocker','app issue'];
  if (!siteId || siteId.length > 80) throw new Error('A valid Global_Site_ID is required.');
  if (!/^[-A-Za-z0-9_:]{8,120}$/.test(eventId)) throw new Error('A unique event ID is required.');
  if (!allowedOutcomes.includes(outcome)) throw new Error('Choose a supported activity outcome.');
  if (!allowedChannels.includes(channel)) throw new Error('Choose a supported activity type.');
  if (!note || note.length > 2000) throw new Error('A call note of 1–2,000 characters is required.');
  if (!['','CALL','EMAIL','WAIT','NONE'].includes(nextAction)) throw new Error('Next action is invalid.');
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
        const priorNote = String(oldRows[i][hh.Note_Text]).trim();
        if (String(oldRows[i][hh.Outcome]).trim() !== outcome ||
            (priorNote !== storedNote && priorNote !== 'AresFit dialer event: '+storedNote) ||
            String(oldRows[i][hh.Follow_Up_Date] || '').trim() !== nextActionDate) throw new Error('Event ID was already used with different activity details.');
        duplicateRow = i+4;
      }
      if (String(oldRows[i][hh.Global_Site_ID]).trim() === siteId && normalise_(oldRows[i][hh.Is_Latest_Event]) === 'YES') demoteRows.push(i+4);
    }

    if (duplicateRow) {
      const duplicateKey = dateKey_(oldRows[duplicateRow-4][hh.Event_Date]) + ' ' + String(oldRows[duplicateRow-4][hh.Event_Time] || '');
      const newer = oldRows.some((r,i) => String(r[hh.Global_Site_ID]).trim() === siteId &&
        normalise_(r[hh.Context_Only]) !== 'YES' &&
        (dateKey_(r[hh.Event_Date]) + ' ' + String(r[hh.Event_Time] || '') > duplicateKey ||
         (dateKey_(r[hh.Event_Date]) + ' ' + String(r[hh.Event_Time] || '') === duplicateKey && i+4 > duplicateRow)));
      if (newer) return {ok:true, duplicate:true, event_id:eventId, global_site_id:siteId};
      demoteRows.forEach(r => { if (r !== duplicateRow) history.getRange(r,hh.Is_Latest_Event+1).setValue('NO'); });
      if (normalise_(oldRows[duplicateRow-4][hh.Is_Latest_Event]) !== 'YES') history.getRange(duplicateRow,hh.Is_Latest_Event+1).setValue('YES');
      persistLatestLeadState_(registry,lead.row,h,historyRowLatestNote_(dateKey_(oldRows[duplicateRow-4][hh.Event_Date]),String(oldRows[duplicateRow-4][hh.Event_Time]||''),outcome,storedNote),String(oldRows[duplicateRow-4][hh.Event_Date]||''),String(oldRows[duplicateRow-4][hh.Follow_Up_Date]||''),outcome,channel,user.name,nextAction,lead.values,eventId);
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
    persistLatestLeadState_(registry,lead.row,h,historyRowLatestNote_(date,time,outcome,storedNote),date,nextActionDate,outcome,channel,user.name,nextAction,lead.values,eventId);
    return {ok:true, duplicate:false, event_id:eventId, global_site_id:siteId, timestamp:isoNow_()};
  } finally {
    lock.releaseLock();
  }
}

function historyRowLatestNote_(date,time,outcome,note) {
  return [date, time, outcome, note].filter(Boolean).join(' · ');
}

function persistLatestLeadState_(registry,rowNumber,h,latestNote,eventDate,nextActionDate,outcome,channel,repName,nextAction,priorRowValues,eventId) {
  const setIfPresent = (name,value) => { if (h[name] != null) registry.getRange(rowNumber,h[name]+1).setValue(value); };
  // Keep the previous registry context before replacing the latest-note cell.
  // Retried event IDs must not append the same text a second time.
  if (h.Rep_Contact_History != null && h.Latest_Real_Notes != null) {
    const historyCell = registry.getRange(rowNumber,h.Rep_Contact_History+1);
    const prior = priorRowValues || [];
    let history = String(prior[h.Rep_Contact_History] || '');
    const previousHistory = history;
    const previousLatest = String(prior[h.Latest_Real_Notes] || '');
    [previousLatest,latestNote].forEach(part => {
      if (part && !history.includes(part)) history += (history ? ' | ' : '') + part;
    });
    if (history !== previousHistory) historyCell.setValue(history);
  }
  // A technical note without a selected next action must not cancel a customer callback.
  if (!nextAction && !nextActionDate && (outcome === 'app issue' || outcome === 'HOLD' || channel === 'app issue' || channel === 'blocker')) return;
  setIfPresent('Latest_Real_Notes',latestNote);
  setIfPresent('Latest_Activity_Date',eventDate);
  setIfPresent('Latest_Activity_Type',channel === 'call' ? 'CALL' : channel.toUpperCase());
  setIfPresent('Latest_Activity_Outcome',outcome);
  setIfPresent('Latest_Note_Date',eventDate);
  setIfPresent('Last_Touched_By',repName);
  setIfPresent('Last_Touched_Date',eventDate);
  setIfPresent('Last_Registry_Update',isoNow_());
  // A fresh activity makes any earlier queue explanation stale.
  setIfPresent('Latest_Event_ID',eventId || '');
  setIfPresent('Next_Action',nextAction);
  setIfPresent('Next_Action_Due_At',nextActionDate);
  setIfPresent('Call_Reason',nextAction === 'CALL' ? 'Follow-up selected after this activity; read the original note.' : '');
  setIfPresent('Action_Source_Event_ID',nextAction ? (eventId || '') : '');
  // Keep the existing Michael columns compatible, without writing them for another rep.
  if (normalise_(repName) === 'MICHAEL') {
    setIfPresent('Michael_Next_Action',nextAction);
    setIfPresent('Michael_Call_Reason',nextAction === 'CALL' ? 'Follow-up selected after this activity; read the original note.' : '');
    setIfPresent('Michael_Reason_Event_Date',nextAction ? eventDate : '');
  }
  if (outcome === 'do not touch') {
    setIfPresent('Suppression_Status','YES');
    setIfPresent('Current_Callability_State','NOT_CALLABLE');
  }
  if (nextActionDate) {
    setIfPresent('Follow_Up_Date',nextActionDate);
    setIfPresent('Callback_Date',nextAction === 'CALL' ? nextActionDate : '');
  } else {
    setIfPresent('Follow_Up_Date','');
    setIfPresent('Callback_Date','');
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

function getAssignedQueue_(user, limit, servedIds, dueOnly) {
  const sheet = getSheet_(CONFIG.REGISTRY_SHEET);
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow <= CONFIG.HEADER_ROW) return {total:0,leads:[],next_cursor:'',queue_ids:[]};
  const values = sheet.getRange(CONFIG.HEADER_ROW, 1, lastRow - CONFIG.HEADER_ROW + 1, lastCol).getDisplayValues();
  const headers = values[0];
  const h = headerIndex_(headers);
  ['Global_Site_ID','Current_Owner','Suppression_Status','Business_Name','Last_Registry_Update','Latest_Activity_Date'].forEach(k => {
    if (h[k] == null) throw new Error('Registry is missing required column: ' + k);
  });
  const ownerName = normalise_(user.name);
  const all = [];
  const byId = {};
  const today = Utilities.formatDate(new Date(),'Europe/London','yyyy-MM-dd');
  const nowTime = Utilities.formatDate(new Date(),'Europe/London','HH:mm');
  const stoppedIds = getLatestStoppedSiteIds_();
  const reviewedQueue = getReviewedQueue_(user);
  const served = new Set(servedIds || []);
  let nextDue = '';
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    if (normalise_(row[h.Current_Owner]) !== ownerName) continue;
    const reviewed = reviewedQueue[String(row[h.Global_Site_ID] || '').trim()];
    const queueItem = reviewed && reviewed.activityDate === dateKey_(value_(row,h,'Latest_Activity_Date')) &&
      reviewed.registryVersion === value_(row,h,'Last_Registry_Update') ? reviewed : null;
    if (queueItem && queueItem.lane === 'HOLD') continue;
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
      master_category: value_(row,h,'Current_Master_Category'),
      manual_review: value_(row,h,'Manual_Review_Required'),
      provider_reject: value_(row,h,'Provider_Reject_Status'),
      no_gym: value_(row,h,'No_Gym_Status'),
      ownership_conflict: value_(row,h,'Ownership_Conflict_Status'),
      review_flag: value_(row,h,'Conflict_Review_Flag'),
      review_reason: value_(row,h,'Conflict_Review_Reason'),
      owner_evidence: value_(row,h,'Current_Owner_Evidence'),
      latest_notes: value_(row,h,'Latest_Real_Notes'),
      contact_history: value_(row,h,'Rep_Contact_History'),
      website: value_(row,h,'Website'),
      intent_level: value_(row,h,'Lead_Intent_Level'),
      intent_evidence: value_(row,h,'Intent_Signal_Evidence'),
      next_action: value_(row,h,'Next_Action') || (ownerName === 'MICHAEL' ? value_(row,h,'Michael_Next_Action') : ''),
      call_reason: value_(row,h,'Call_Reason') || (ownerName === 'MICHAEL' ? value_(row,h,'Michael_Call_Reason') : ''),
      reason_event_date: ownerName === 'MICHAEL' ? value_(row,h,'Michael_Reason_Event_Date') : '',
      action_source_event_id: value_(row,h,'Action_Source_Event_ID'),
      latest_event_id: value_(row,h,'Latest_Event_ID'),
      action_due_at: value_(row,h,'Next_Action_Due_At'),
      callback_date: value_(row,h,'Callback_Date'),
      follow_up_date: value_(row,h,'Follow_Up_Date'),
      latest_activity_date: value_(row,h,'Latest_Activity_Date')
    };
    if (!lead.global_site_id || !lead.business_name || String(lead.phone).replace(/\D/g,'').length < 10) continue;
    const callability = normalise_(lead.callability);
    if (stoppedIds[lead.global_site_id] || callability === 'NOT_CALLABLE' || callability === 'DO_NOT_CALL' || callability.startsWith('BLOCKED')) continue;
    if ((callability === 'EMAIL_ACTION_REQUIRED' || callability.startsWith('NOT_CALL_READY')) &&
        !(currentQueueAction_(lead) === 'CALL' && lead.action_due_at)) continue;
    const acceptedAssignment = ownerName === 'KRZYSZTOF' && normalise_(lead.master_category) === 'KRZYSZTOF_50_OWEN_COLD_CALL_READY' &&
      normalise_(lead.lifecycle_status) === 'CURRENTLY_ASSIGNED_CALL_READY' && normalise_(lead.review_flag) === 'NO' &&
      lead.review_reason === 'Direct Lead_ID, phone and other-active-rep parent-chain crossover controls clear at assignment.' &&
      lead.owner_evidence.includes('Exact '+lead.global_site_id+' selected') && lead.owner_evidence.includes('reassigned to Krzysztof');
    if ((normalise_(lead.manual_review) === 'YES' && !acceptedAssignment && !(queueItem && queueItem.lane === 'CALL')) ||
        normalise_(lead.provider_reject) === 'YES' || normalise_(lead.no_gym) === 'YES' || normalise_(lead.ownership_conflict) === 'YES') continue;
    if (['RESEARCH_HOLD','BOYS_POOL_AVAILABLE_NOT_CALL_READY','OWNERSHIP_CONFLICT_QUARANTINE'].includes(normalise_(lead.master_category))) continue;
    if (normalise_(lead.lifecycle_status) === 'CLOSED_LOST' || normalise_(lead.closed_lost) === 'YES') continue;
    if (['LOST','NOT INTERESTED','NOT INT.','PERMANENT CLOSURE CONFIRMED'].includes(normalise_(lead.latest_outcome)) || normalise_(lead.activity_status) === 'PERMANENT CLOSURE CONFIRMED') continue;
    const signal = queueSignal_(lead,queueItem,today,nowTime);
    if (!signal) continue;
    if (signal.lane === 'LATER_TODAY') {
      if (!nextDue || signal.sortDate < nextDue) nextDue = signal.sortDate;
      continue;
    }
    lead.queue_reason = signal.reason;
    lead.queue_reason_date = signal.sourceDate;
    lead.queue_lane = signal.lane;
    lead._rank = signal.rank;
    lead._sortDate = signal.sortDate;
    if (!byId[lead.global_site_id]) { all.push(lead); byId[lead.global_site_id]=lead; }
  }
  all.sort((a,b) => a._rank-b._rank || a._sortDate.localeCompare(b._sortDate) || a.global_site_id.localeCompare(b.global_site_id));
  const remaining = all.filter(l => !served.has(l.global_site_id));
  const candidates = dueOnly ? remaining.filter(l => l.queue_lane === 'DUE') : remaining;
  const leads = candidates.slice(0,limit);
  leads.forEach(l => { delete l._rank; delete l._sortDate; });
  return {total:all.length,leads:leads,remaining:remaining.length-leads.length,queue_ids:all.map(l=>l.global_site_id),next_due_at:nextDue};
}

function queueDays_(from,to) {
  const a = Date.parse(dateKey_(from)+'T00:00:00Z');
  const b = Date.parse(dateKey_(to)+'T00:00:00Z');
  return Number.isFinite(a) && Number.isFinite(b) ? Math.round((b-a)/86400000) : 9999;
}

function recentMissStreak_(notes) {
  let count = 0;
  const parts = String(notes || '').split(/\s+\|\s+/);
  for (let i=parts.length-1; i>=0; i--) {
    const part = parts[i].trim();
    if (!/^\d{1,2}\/\d{1,2}\s+\d{1,2}:\d{2}\s*-/.test(part)) continue;
    if (/^\d{1,2}\/\d{1,2}\s+\d{1,2}:\d{2}\s*-\s*(?:VM|NA|VOICEMAIL|NO ANSWER)\b/i.test(part)) count++;
    else break;
  }
  return count;
}

function recentRoute_(notes,today) {
  const parts = String(notes || '').split(/\s+\|\s+/);
  for (let i=parts.length-1; i>=0; i--) {
    const m = parts[i].match(/(\d{1,2})\/(\d{1,2})\s+\d{1,2}:\d{2}\s*-\s*(?:CALLBACK|CONTACTED)(?:\s+\[(?:DM|GK)\])?\s*-\s*(.+)$/i);
    if (!m) continue;
    const date = today.slice(0,4)+'-'+('0'+m[2]).slice(-2)+'-'+('0'+m[1]).slice(-2);
    if (queueDays_(date,today) > 30 || queueDays_(date,today) < 0) continue;
    const detail = String(m[3] || '').trim();
    if (/\b(?:email instead|best.*email|do not call|not interested)\b/i.test(detail)) return null;
    if (!/\b(?:DM|DECISION MAKER|OWNER|MANAGER|CALL BACK|RING BACK|AVAILABLE|IN THE CLUB|IN TOMORROW)\b/i.test(detail)) continue;
    if (/\b(?:fuck|cunt|nigga|faggot)\b/i.test(detail)) return {date,detail:''};
    return {date,detail:detail.slice(0,170)};
  }
  return null;
}

function currentQueueAction_(lead) {
  if (lead.action_source_event_id || lead.latest_event_id) {
    return lead.action_source_event_id && lead.action_source_event_id === lead.latest_event_id ? normalise_(lead.next_action) : '';
  }
  return dateKey_(lead.reason_event_date) === dateKey_(lead.latest_activity_date) && dateKey_(lead.reason_event_date) ? normalise_(lead.next_action) : '';
}

function queueDue_(value,today,nowTime) {
  const text = String(value || '').trim();
  const date = dateKey_(text);
  if (!date) return null;
  const time = text.match(/(?:T|\s)(\d{2}:\d{2})/);
  const range = text.match(/(\d{2}:\d{2})\s*[-\u2013]\s*(\d{2}:\d{2})/);
  const future = date > today || (date === today && time && time[1] > nowTime);
  return {date,time:time ? time[1] : '00:00',future,expiredWindow:date === today && range && range[2] < nowTime};
}

function queueSignal_(lead,queueItem,today,nowTime) {
  const age = queueDays_(lead.latest_activity_date,today);
  const outcome = normalise_(lead.latest_outcome);
  const status = normalise_(lead.activity_status);
  const category = normalise_(lead.master_category);
  const action = currentQueueAction_(lead);
  const sourceDate = dateKey_(lead.latest_activity_date);
  const lastText = String(lead.latest_outcome || '').toLowerCase();
  const lastNote = String(lead.latest_notes || '').split(/\s+\|\s+/).pop() || '';
  const actualMiss = /^(?:VM|NA|VOICEMAIL|NO ANSWER)\b/.test(outcome);
  const miss = actualMiss || ['VOICEMAIL','NO ANSWER'].includes(status);
  const latestEmail = /^(?:EMAILED|EMAIL SENT|EMAIL REQUESTED|EMAIL ROUTE)/.test(outcome);
  if (/\b(?:PROVIDER REJECT|NOT INTERESTED|NOT INT\.|DO NOT CALL|PERMANENT CLOSURE|DEAD AIR|NOT APPLICABLE)\b/.test(outcome+' '+status) ||
      (/\b(?:HOLD|PARKED|AWAITING REPLY|EMAIL SENT|DNC)\b/.test(status) && action !== 'CALL')) return null;
  const blockedText = /(?:email first|best (?:form|way) of contact is (?:via )?email|asked (?:me|us) to (?:send|email)|told (?:me|us) (?:to |an? )email|quote (?:send )?paused|supplier agreement hold|customer will reach out|not looking for anything|no (?:current|active|immediate) (?:need|requirement|opportunity|budget)|wait for (?:lee|customer|supplier))/i.test(lastText);
  const disqualifiedNote = /(?:permanently? closed|closed forever|permanent closure|not a gym|no gym|do not call|\bDNC\b|not interested|no longer operating|closed down)/i.test(lastNote);
  const needsNumberReview = /(?:research flag|new number|number needs (?:checking|updating))/i.test(lastNote);
  const testEvent = /(?:aresfit dialer event:\s*test\b|^test (?:note|log|logging)\b)/i.test(lastNote);
  if (['EMAIL','WAIT','NONE','HOLD'].includes(action) || blockedText || disqualifiedNote || needsNumberReview || testEvent) return null;
  if (latestEmail && action !== 'CALL') return null;
  if (/\b(?:gave|give|given)\b.*\bnumber\b|\bnew (?:owner )?(?:number|mobile)\b/i.test(lastText+' '+lastNote)) return null;
  if (queueItem && queueItem.lane === 'HOLD') return null;

  // Explicit next actions are valid only against the same source activity.
  if (action === 'CALL' && lead.call_reason && (!miss || lead.action_due_at)) {
    const due = queueDue_(lead.action_due_at || lead.callback_date || lead.follow_up_date,today,nowTime);
    if (due && (due.date > today || due.expiredWindow)) return null;
    if (!due && age < 1) return null;
    const laterToday = due && due.future;
    return {lane:laterToday?'LATER_TODAY':due?'DUE':'FOLLOW_UP',rank:laterToday?3500:due && due.date === today?500:1000,
      sortDate:due ? due.date+' '+due.time : sourceDate,
      reason:lead.call_reason,sourceDate};
  }

  const dueText = (action ? lead.action_due_at : '') || lead.callback_date || lead.follow_up_date;
  const due = dateKey_(dueText);
  const daysPastDue = queueDays_(due,today);
  const callbackEvidence = /CALLBACK|CALL BACK|DM AVAILABLE|RING BACK/.test(outcome+' '+status+' '+normalise_(lead.callability));
  const schedule = queueDue_(dueText,today,nowTime);
  if (due && (due > today || schedule.expiredWindow)) return null;
  if (callbackEvidence && !miss && due && daysPastDue >= 0 &&
      (!sourceDate || due >= sourceDate)) {
    // A missed one-off gatekeeper availability slot is not a standing buyer promise.
    if (daysPastDue > 14 && /\[GK\]|gatekeeper|\bGK\b/i.test(lastNote+' '+lead.latest_outcome)) return null;
    if (daysPastDue > 3 && /\[GK\]|gatekeeper|\bGK\b/i.test(lastNote+' '+lead.latest_outcome)) {
      return {lane:'RETRY',rank:4600,sortDate:due,reason:'Earlier gatekeeper availability has passed. Recheck availability; no buyer interest is confirmed.',sourceDate};
    }
    const laterToday = schedule.future;
    return {lane:laterToday?'LATER_TODAY':'DUE',rank:laterToday?3500:due === today?500:1500,
      sortDate:due+' '+schedule.time,
      reason:laterToday?'Callback due today at '+schedule.time+'.':due === today?'Requested callback is due today.':'Unresolved callback from '+due+' is overdue; read the original note.',
      sourceDate};
  }

  if (!miss && age <= 21 && normalise_(lead.intent_level) === 'ACTIVE_OPPORTUNITY' &&
      /CONTACTED|REACHED|CALLBACK/.test(outcome+' '+status) &&
      /\b(?:INTERESTED|EQUIPMENT|REPLACEMENT|BUDGET|PRICING|BUYING|NEEDS? KIT)\b/i.test(outcome+' '+lastNote) &&
      !/QUOTE|EMAIL|SUPPLIER|BLOCKED|WAITING|NOT INTERESTED/.test(outcome+' '+status)) {
    return {lane:'FOLLOW_UP',rank:2000,sortDate:sourceDate,
      reason:'Recent buyer conversation needs a follow-up call; check the dated notes before dialing.',sourceDate};
  }

  if (miss && age >= 3 && age <= 21) {
    const streak = recentMissStreak_(lead.latest_notes);
    if (streak >= 4 && age < 10) return null;
    const route = recentRoute_(lead.latest_notes,today);
    const oneOff = route && /\b(?:today|tomorrow|tonight|minutes?|monday|tuesday|wednesday|thursday|friday|saturday|sunday)\b|\b\d{1,2}\s*(?:am|pm)\b/i.test(route.detail);
    const targeted = route && queueDays_(route.date,today) <= 2 && !oneOff && streak < 3;
    const namedBuyerRoute = targeted && /\b(?:DM|DECISION MAKER|OWNER)\b/i.test(route.detail);
    return {lane:'RETRY',rank:targeted ? 2900+(namedBuyerRoute?0:120)+Math.min(2,streak)*150+Math.abs(age-5)*5 : streak >= 3 ? 5500 : 4500+Math.abs(age-7)*5,
      sortDate:sourceDate,
      reason:targeted ? 'Earlier decision-maker route ('+route.date+'): '+(route.detail || 'see the original call note')+'. Latest call did not connect; confirm the route still applies.' :
        streak >= 3 ? 'Several missed attempts; retry only after a longer gap.' : 'A recent call did not connect; retry after a sensible gap. No current buyer signal is recorded.',sourceDate};
  }

  // Cold volume requires an identifiable gym and a usable contact route.
  // A name or an old priority label alone is not a buyer signal.
  const coldStatus = ['UNCALLED','NEW','COLD','NOT CONTACTED'].includes(status);
  const gymName = /\b(?:GYM|FITNESS|CROSSFIT|STRENGTH|TRAINING|HEALTH CLUB|PERFORMANCE)\b/i.test(lead.business_name);
  const ambiguousSite = /\b(?:SCHOOL|COLLEGE|UNIVERSITY|GOLF|POOLS?|LEISURE CENTRE|SPORTS CENTRE)\b/i.test(lead.business_name);
  const usablePhone = /^0\d{9,10}$/.test(String(lead.phone || '').replace(/\D/g,''));
  const priorCall = /\b\d{1,2}\/\d{1,2}\s+\d{1,2}:\d{2}\s*-\s*(?:VM|NA|CONTACTED|CALLBACK|DEAD AIR|NOT INT\.)\b/i.test(lead.latest_notes+' | '+lead.contact_history);
  const assignedReady = category === 'KRZYSZTOF_50_OWEN_COLD_CALL_READY' && normalise_(lead.lifecycle_status) === 'CURRENTLY_ASSIGNED_CALL_READY';
  if (coldStatus && (gymName || assignedReady) && (!ambiguousSite || assignedReady) && usablePhone && (lead.website || assignedReady) &&
      !priorCall &&
      ['UNRESOLVED','HISTORICAL_ONLY','MICHAEL ACTIVE/TOUCHED ADDENDUM','KRZYSZTOF_50_OWEN_COLD_CALL_READY','ACTIVE_ASSIGNED'].includes(category)) {
    return {lane:'COLD',rank:4800,sortDate:lead.global_site_id,
      reason:assignedReady ? 'Reviewed cold assignment; no current buying signal is recorded.' : 'Cold gym candidate with a phone and website; no current buying signal is recorded.',sourceDate};
  }
  return null;
}

function getReviewedQueue_(user) {
  const sheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getSheetByName(user.name + ' Call Queue');
  if (!sheet) return {};
  const values = sheet.getDataRange().getDisplayValues();
  if (!values.length) throw new Error('Michael Call Queue is empty.');
  const h = headerIndex_(values[0]);
  ['Global_Site_ID','Lane','Priority','Due_At','Registry_Version','Activity_Date_At_Review'].forEach(k => {
    if (h[k] == null) throw new Error('Michael Call Queue is missing required column: ' + k);
  });
  const queue = {};
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const id = String(row[h.Global_Site_ID] || '').trim();
    if (!id) continue;
    if (queue[id]) throw new Error('Michael Call Queue contains duplicate ID: ' + id);
    const lane = normalise_(row[h.Lane]);
    const priority = Number(row[h.Priority]);
    if (lane === 'CALL' && (!Number.isInteger(priority) || priority < 1)) throw new Error('Michael Call Queue priority is invalid at row ' + (i+1));
    const due = String(row[h.Due_At] || '').trim();
    if (due && !dateKey_(due)) throw new Error('Michael Call Queue due date is invalid at row ' + (i+1));
    queue[id] = {lane, priority:lane === 'CALL' ? priority : 9999, due, registryVersion:String(row[h.Registry_Version] || '').trim(), activityDate:String(row[h.Activity_Date_At_Review] || '').trim()};
  }
  return queue;
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

