import fs from 'node:fs';
import vm from 'node:vm';
import assert from 'node:assert/strict';

const code=fs.readFileSync(new URL('../apps-script/Code.gs',import.meta.url),'utf8');
const headers=['Global_Site_ID','Current_Owner','Suppression_Status','Business_Name','Last_Registry_Update','Latest_Activity_Date','Phone','Website','Current_Activity_Status','Current_Callability_State','Current_Master_Category','Current_Lifecycle_Status','Latest_Activity_Outcome','Latest_Real_Notes','Rep_Contact_History','Callback_Date','Follow_Up_Date','Manual_Review_Required','No_Gym_Status','Provider_Reject_Status','Ownership_Conflict_Status','Lead_Intent_Level','Michael_Next_Action','Michael_Call_Reason','Michael_Reason_Event_Date','Next_Action','Next_Action_Due_At','Call_Reason','Action_Source_Event_ID','Latest_Event_ID'];
let source=[],nowTime='09:00',user={rep_id:'REP-MICHAEL',name:'Michael',email:'michael@aresfit.co.uk'},serial=0;
const cached=new Map();
const registry={getLastRow:()=>source.length+3,getLastColumn:()=>headers.length,getRange:()=>({getDisplayValues:()=>[headers,...source.map(r=>headers.map(h=>r[h]||''))]})};
const ctx=vm.createContext({console,Date,Set,JSON,SpreadsheetApp:{openById:()=>({getSheetByName:()=>registry})},CacheService:{getScriptCache:()=>({get:k=>cached.get(k),put:(k,v)=>cached.set(k,v)})},Utilities:{getUuid:()=>`00000000-0000-0000-0000-${String(++serial).padStart(12,'0')}`,formatDate:(_,zone,format)=>format==='HH:mm'?nowTime:'2026-09-29'}});
vm.runInContext(code,ctx);
ctx.requireSignedInAresFitUser_=()=>user;
ctx.getReviewedQueue_=()=>({});
ctx.getLatestStoppedSiteIds_=()=>({});
const row=(i,extra={})=>({Global_Site_ID:`AFS-${String(i).padStart(6,'0')}`,Current_Owner:'Michael',Suppression_Status:'NO',Business_Name:`Example Fitness ${i}`,Last_Registry_Update:'2026-09-28',Phone:'01234567890',Website:'https://example.org',Current_Activity_Status:'Uncalled',Current_Master_Category:'HISTORICAL_ONLY',...extra});
const callback=(i,date='2026-09-29 09:00')=>row(i,{Current_Activity_Status:'Callback',Current_Callability_State:'WAITING_ON_DM_CALLBACK',Latest_Activity_Date:'2026-09-25',Latest_Activity_Outcome:'Callback',Callback_Date:date,Latest_Real_Notes:'25/09 12:00 - Callback - owner asked us to call at agreed time'});
source=Array.from({length:160},(_,i)=>row(i+1));
source.push(callback(900,'2026-09-29 16:00'),callback(901),row(902,{Current_Owner:'Krzysztof',Current_Master_Category:'KRZYSZTOF_50_OWEN_COLD_CALL_READY',Current_Activity_Status:'Not Contacted'}));
const first=ctx.getQueueForSignedInUser('');
assert.equal(first.leads.length,50);assert.equal(first.leads[0].global_site_id,'AFS-000901');assert.equal(first.next_due_at,'2026-09-29 16:00');
assert.ok(!first.leads.some(l=>l.global_site_id==='AFS-000900'));
source.slice(49,70).forEach(r=>r.Suppression_Status='YES');
const second=ctx.getQueueForSignedInUser(first.queue_cursor);
assert.equal(second.leads.length,50,'refill page after 21 leads become suppressed');
assert.equal(new Set([...first.leads,...second.leads].map(l=>l.global_site_id)).size,100);
const separate=ctx.getQueueForSignedInUser('');
assert.notEqual(first.queue_cursor,separate.queue_cursor);
const third=ctx.getQueueForSignedInUser(first.queue_cursor);
assert.ok(!third.leads.some(l=>second.leads.some(p=>p.global_site_id===l.global_site_id)),'another tab must not reset existing cursor');
nowTime='16:00';
const timed=ctx.getQueueForSignedInUser(first.queue_cursor,true);
assert.deepEqual(Array.from(timed.leads,l=>l.global_site_id),['AFS-000900']);
assert.equal(ctx.getQueueForSignedInUser(first.queue_cursor,true).leads.length,0);
user={rep_id:'REP-KRZYSZTOF',name:'Krzysztof',email:'krzysztof@aresfit.co.uk'};
assert.equal(ctx.getQueueForSignedInUser('').leads[0].global_site_id,'AFS-000902');
assert.throws(()=>ctx.getQueueForSignedInUser(first.queue_cursor),/expired/,'rep cannot reuse another rep cursor');
user={rep_id:'REP-MICHAEL',name:'Michael',email:'michael@aresfit.co.uk'};
source=Array.from({length:79},(_,i)=>row(i+1));
const a=ctx.getQueueForSignedInUser(''),b=ctx.getQueueForSignedInUser(a.queue_cursor);
assert.equal(a.leads.length,50);assert.equal(b.leads.length,29);assert.equal(b.remaining,0);assert.equal(b.next_cursor,'');
const lead={business_name:'Example Gym',master_category:'HISTORICAL_ONLY',activity_status:'Callback',latest_activity_date:'2026-09-01',latest_outcome:'Callback',latest_notes:'01/09 10:00 - Callback - call at agreed time',callback_date:'2026-09-22 16:00'};
assert.equal(ctx.queueSignal_(lead,null,'2026-09-29','09:00').lane,'DUE','overdue callback should not vanish after three days');
assert.equal(ctx.queueSignal_({...lead,callback_date:'2026-09-29 07:00-12:00'},null,'2026-09-29','13:00'),null,'respect expired availability window');
assert.equal(ctx.queueSignal_({...lead,callback_date:'2026-09-30'},null,'2026-09-29','09:00'),null);
assert.ok(ctx.queueSignal_({...lead,next_action:'WAIT',reason_event_date:'2026-08-01'},null,'2026-09-29','09:00'),'old wait recommendation cannot override fresh callback');
assert.equal(ctx.queueSignal_({...lead,next_action:'WAIT',action_source_event_id:'EVT1',latest_event_id:'EVT1'},null,'2026-09-29','09:00'),null);
assert.ok(ctx.queueSignal_({...lead,next_action:'WAIT',action_source_event_id:'OLD',latest_event_id:'NEW'},null,'2026-09-29','09:00'));
assert.ok(ctx.queueSignal_({...lead,latest_notes:'test note | 25/09 10:00 - Callback - call owner'},null,'2026-09-29','09:00'),'old test note cannot suppress real latest activity');
assert.equal(ctx.queueSignal_({...lead,activity_status:'Dead Air'},null,'2026-09-29','09:00'),null);
assert.equal(ctx.queueSignal_({...lead,latest_outcome:'email first, customer will reach out'},null,'2026-09-29','09:00'),null);
const values=new Map();
const names=['Latest_Real_Notes','Rep_Contact_History','Next_Action','Next_Action_Due_At','Call_Reason','Action_Source_Event_ID','Latest_Event_ID','Michael_Next_Action','Michael_Call_Reason','Michael_Reason_Event_Date','Follow_Up_Date','Callback_Date'];
const h=Object.fromEntries(names.map((n,i)=>[n,i]));
const prior=names.map(n=>n==='Latest_Real_Notes'?'Original Spencer note':n==='Rep_Contact_History'?'Earlier notes':'');
ctx.persistLatestLeadState_({getRange:(r,c)=>({setValue:v=>values.set(names[c-1],v)})},1,h,'New original note','2026-09-29','2026-09-30 10:00','reached','call','Krzysztof','CALL',prior,'EVENT-123');
assert.equal(values.get('Next_Action'),'CALL');assert.equal(values.get('Action_Source_Event_ID'),'EVENT-123');
assert.ok(values.get('Rep_Contact_History').includes('Original Spencer note'));
assert.ok(values.get('Rep_Contact_History').includes('Earlier notes'));
assert.equal(values.has('Michael_Next_Action'),false,'Krzysztof must not write Michael action columns');
assert.equal(values.get('Callback_Date'),'2026-09-30 10:00');
values.clear();
ctx.persistLatestLeadState_({getRange:(r,c)=>({setValue:v=>values.set(names[c-1],v)})},1,h,'App issue note','2026-09-29','','app issue','app issue','Michael','',prior,'EVENT-TECH');
assert.equal(values.has('Callback_Date'),false,'technical note must preserve customer callback');
assert.equal(values.has('Latest_Event_ID'),false,'technical note must preserve customer action identity');
console.log('PASS: session isolation, 50+50 refill, actual 79 total, timed callbacks, due windows, ownership, stale actions, notes preservation and rep-neutral writes');

// Replaying an older acknowledged event must never demote a newer customer event.
const historyHeaders=['Event_ID','Global_Site_ID','Business_Name','Rep','Event_Date','Event_Time','Date_Confidence','Event_Type','Outcome','Note_Text','Status_After','Stage_After','Follow_Up_Date','Match_Method','First_Source','Latest_Source','Source_Count','Source_SHA256s','Source_Observed_Date','Context_Only','Is_Latest_Event'];
const historyRows=[{Event_ID:'DIALER-OLD-123',Global_Site_ID:'AFS-000001',Event_Date:'2026-09-28',Event_Time:'10:00:00',Outcome:'reached',Note_Text:'Old original note',Is_Latest_Event:'NO'},{Event_ID:'DIALER-NEW-123',Global_Site_ID:'AFS-000001',Event_Date:'2026-09-29',Event_Time:'10:00:00',Outcome:'reached',Note_Text:'New original note',Is_Latest_Event:'YES'}];
const history={getLastRow:()=>5,getLastColumn:()=>historyHeaders.length,getRange:(row)=>({getDisplayValues:()=>row===3?[historyHeaders]:historyRows.map(r=>historyHeaders.map(h=>r[h]||'')),setValue:()=>{throw new Error('Unexpected replay write')}})};
source=[row(1)];ctx.getSheet_=name=>name==='Activity History'?history:registry;ctx.LockService={getScriptLock:()=>({waitLock(){},releaseLock(){}})};
const replay=ctx.logActivityForSignedInUser({global_site_id:'AFS-000001',event_id:'DIALER-OLD-123',outcome:'reached',channel:'call',note_text:'Old original note'});
assert.equal(replay.duplicate,true);
console.log('PASS: older duplicate event cannot overwrite newer history');
ctx.getSheet_=()=>registry;

// Optional ephemeral provider read; source records never enter the repository.
if(process.argv[2]){
 const fixture=JSON.parse(fs.readFileSync(process.argv[2],'utf8'));
 source=fixture.rows;headers.splice(0,headers.length,...fixture.headers);
 const outcomes=[];
 for(const name of ['Michael','Krzysztof']){
  user={name,rep_id:'REP-'+name.toUpperCase(),email:name.toLowerCase()+'@aresfit.co.uk'};
  for(const day of ['2026-09-29','2026-09-30'])for(const time of ['09:00','16:00']){nowTime=time;ctx.Utilities.formatDate=(_,zone,format)=>format==='HH:mm'?nowTime:day;const q=ctx.getQueueForSignedInUser('');outcomes.push({rep:name,day,time,eligible:q.total,first:q.leads.slice(0,5).map(l=>({id:l.global_site_id,name:l.business_name,lane:l.queue_lane,reason:l.queue_reason,lastActivity:l.latest_activity_date})),nextDue:q.next_due_at});}
 }
 console.log(JSON.stringify(outcomes,null,2));
}
