/* ============================================================
   TaskTrack — single-file app logic
   Storage: localStorage with safe in-memory fallback
   ============================================================ */

import { isSupabaseConfigured, supabase } from './supabaseClient.js';
import { supabaseConfig } from './supabaseConfig.js';
import {
  applyRegenerationPreview,
  buildTaskDraftFromMeetingAction,
  canonicalMeetingTaskResult,
  createMeetingDraft,
  createRegenerationPreview,
  formatMeetingNotes,
  suggestMeetingActionOwner,
  validateGeneratedMeetingAnalysis,
} from './meetingDomain.js';
import { createMeetingRepository } from './meetingRepository.js';
import {
  renderMeetingAnalysisPreview,
  renderMeetingCaptureBanner,
  renderMeetingsView,
} from './meetingView.js';
import {
  appendCaptureSegment,
  canonicalCaptureTranscript,
  captureElapsedLabel,
  normalizeCaptureDraft,
  realtimeTranscriptEventToSegment,
} from './meetingCapture.js';
import { googleCalendarItems, googleEventCalendarSpan, googleEventDisplayEnd } from './googleCalendarView.js';
import { calendarFeedRelevantToMember, mergeSharedCalendarFeed } from './sharedCalendarFeed.js';

const REQUIRED_WORKSPACE_ID = supabaseConfig.sharedWorkspaceId || '';
const ASSET_BUCKET = 'workspace-assets';
const MEETING_CAPTURE_BUCKET = 'meeting-capture-audio';
const MAX_LOCAL_ASSET_BYTES = 2_500_000;
const MAX_REMOTE_ASSET_BYTES = 25_000_000;
const BLOCKED_ASSET_EXTENSIONS = new Set(['bat','cmd','com','cpl','dll','exe','hta','html','js','jse','mjs','msi','ps1','scr','svg','vbs','wsf']);
const ALLOWED_ASSET_EXTENSIONS = new Set(['7z','aac','avi','bmp','csv','doc','docx','gif','heic','jpeg','jpg','json','key','md','mov','mp3','mp4','odt','pages','pdf','png','ppt','pptx','psd','rtf','srt','tar','tif','tiff','txt','wav','webm','webp','xls','xlsx','xml','zip']);
const ALLOWED_ASSET_MIME_EXACT = new Set([
  'application/msword',
  'application/pdf',
  'application/vnd.ms-excel',
  'application/vnd.ms-powerpoint',
  'application/vnd.openxmlformats-officedocument.presentationml.presentation',
  'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  'text/csv',
  'text/plain'
]);

/* ---------- safe storage wrapper ---------- */
const Store = (() => {
  let mem = {}, usable = false;
  try { const k='__iw_t'; localStorage.setItem(k,'1'); localStorage.removeItem(k); usable=true; } catch(e){ usable=false; }
  return {
    persistent: usable,
    get(key){ try{ return usable ? localStorage.getItem(key) : (mem[key] ?? null); }catch(e){ return mem[key] ?? null; } },
    set(key,val){ try{ usable ? localStorage.setItem(key,val) : (mem[key]=val); }catch(e){ mem[key]=val; } }
  };
})();

/* ---------- palettes ---------- */
const STATUS = {
  not_started:{label:'Not Started',color:'#c4c4c4'},
  working:    {label:'Working on it',color:'#fdab3d'},
  review:     {label:'Review',color:'#579bfc'},
  stuck:      {label:'Stuck',color:'#e2445c'},
  done:       {label:'Done',color:'#00c875'},
  on_hold:    {label:'On Hold',color:'#a25ddc'}
};
const PRIORITY = {
  none:{label:'—',color:'#c4c4c4'},
  low:{label:'Low',color:'#579bfc'},
  medium:{label:'Medium',color:'#a25ddc'},
  high:{label:'High',color:'#401694'},
  critical:{label:'Critical',color:'#e2445c'}
};
const AVATAR_COLORS = ['#0073ea','#00c875','#fdab3d','#e2445c','#a25ddc','#00c2ff','#ff5ac4','#7e3b8a','#037f4c','#cab641'];
const GROUP_COLORS = ['#0073ea','#00c875','#fdab3d','#e2445c','#a25ddc','#579bfc','#037f4c','#ff5ac4'];
const LABEL_COLORS = ['#00c875','#fdab3d','#e2445c','#c4c4c4','#579bfc','#a25ddc','#401694','#0073ea','#037f4c','#ff5ac4'];
const ASSISTANT_PET_DEFAULTS = {
  enabled:true,
  name:'TaskTrack',
  style:'retriever',
  color:'#d99a35',
  size:'medium',
  position:'bottom-right',
  motion:'float',
  autoBrief:false
};
const ASSISTANT_PET_STYLES = [
  {id:'retriever', label:'Golden retriever'},
  {id:'seal', label:'Seal'},
  {id:'cat', label:'Cat'},
  {id:'bunny', label:'Bunny'}
];
const ASSISTANT_PET_STYLE_COLORS = {
  retriever:'#d99a35',
  seal:'#aeb8c2',
  cat:'#7c8aa0',
  bunny:'#f2eef2'
};
const LANGUAGE_OPTIONS = {
  en:'English',
  ko:'한국어'
};
const TEXT = {
  en:{
    workspaceHome:'Workspace Home',
    updateFeed:'Update feed',
    monthlyClosing:'Monthly Closing',
    dashboard:'Dashboard',
    calendar:'Calendar',
    myWork:'My Work',
    meetings:'Meetings',
    integrations:'Integrations',
    automations:'Automations',
    docsHub:'Docs Hub',
    settings:'Settings',
    searchPlaceholder:'Search tasks, boards, docs, people',
    board:'Board',
    table:'Table',
    kanban:'Kanban',
    newTask:'New task',
    person:'Person',
    sort:'Sort',
    hide:'Hide'
  },
  ko:{
    workspaceHome:'워크스페이스 홈',
    updateFeed:'업데이트 피드',
    monthlyClosing:'월간 마감',
    dashboard:'대시보드',
    calendar:'캘린더',
    myWork:'내 작업',
    meetings:'회의록',
    integrations:'연동',
    automations:'자동화',
    docsHub:'문서 허브',
    settings:'설정',
    searchPlaceholder:'작업, 보드, 문서, 사람 검색',
    board:'보드',
    table:'테이블',
    kanban:'칸반',
    newTask:'새 작업',
    person:'사람',
    sort:'정렬',
    hide:'숨기기'
  }
};
const NOTIFICATION_DEFAULTS = {
  mentions:true,
  everyone:true,
  assigned:true,
  subscribed:true,
  authored:true,
  taskUpdates:true,
  dueReminders:true,
  desktop:true,
  push:true,
  email:true,
  emailDueReminders:true,
  dailyDigest:true
};
const ASSISTANT_ACTION_DEFAULTS = {
  gmailDrafts:true,
  gmailSend:false,
  calendarEvents:true,
  driveSearch:true,
  driveDocs:false,
  driveImport:true,
  driveManage:false,
  driveEditDocs:false,
  notionImport:true
};

/* ---------- default seed data ---------- */
function seed(){
  const names=['Alex Rivera','Jordan Vale','Sam Linden','Mira Chen','Theo Brooks','Nora Finch','Eli Rowan','Zoe Hart','Owen Reed','Isla Park','Leo Quinn','Ava Stone','Miles Avery','Cleo Lane','Remy Blake'];
  const colors=['#0073ea','#00c875','#fdab3d','#a25ddc','#ff6d3b'];
  const team=names.map((name,i)=>({id:'u'+(i+1),name,initials:name.split(' ').map(v=>v[0]).join(''),color:colors[i%5],email:name.toLowerCase().replace(' ','.')+'@example.invalid'}));
  const d=n=>{const date=new Date();date.setDate(date.getDate()+n);return date.toISOString().slice(0,10);};
  const departmentTasks=[
    ['Marketing','📣',['Review Lumen Grove creator brief','Approve Tidewell video captions','Prepare weekly creator report','Plan Fieldnote launch calendar','Review the next campaign budget','Archive completed content']],
    ['Ecommerce','🛍️',['Review storefront product copy','Check sample product images','Prepare weekly conversion report','Schedule seasonal page update','Review promotion dates','Close completed listing checks']],
    ['Sales','📊',['Prepare regional sales review','Verify warehouse forecast','Review account meeting notes','Plan next retailer presentation','Check open replenishment requests','Archive completed account review']],
    ['Operations','📦',['Review sample shipping queue','Confirm mock stock allocation','Prepare cross-team meeting notes','Update onboarding checklist','Review document permissions','Complete monthly process review']]
  ];
  const boards=departmentTasks.map(([name,icon,titles],b)=>({id:'b'+(b+1),name:name+' Tasks',icon,groups:[
    {id:'g'+b+'active',name:'Active Tasks',color:colors[b],open:true,tasks:titles.slice(0,3).map((name,j)=>({id:'t'+b+j,name,owner:'u'+((b*3+j)%15+1),status:j===1?'stuck':'working',priority:j===0?'high':'medium',due:d(j===1?-2:j+3),tlStart:d(-5),tlEnd:d(j+3),budget:300+j*250,notes:j===1?'Fictional sample: waiting for approval.':'Fictional sample for the local portfolio.'}))},
    {id:'g'+b+'plan',name:'Planning',color:'#fdab3d',open:true,tasks:titles.slice(3,5).map((name,j)=>({id:'t'+b+(j+3),name,owner:'u'+((b*3+j+3)%15+1),status:j?'on_hold':'not_started',priority:'medium',due:d(10+j*4),tlStart:d(5),tlEnd:d(10+j*4),budget:500,notes:'Draft plan; no external work has been scheduled.'}))},
    {id:'g'+b+'done',name:'Completed',color:'#00c875',open:true,tasks:[{id:'t'+b+'5',name:titles[5],owner:'u'+((b*3+5)%15+1),status:'done',priority:'low',due:d(-4),tlStart:d(-12),tlEnd:d(-4),budget:100,notes:'Completed in this fictional scenario.'}]}
  ]}));
  return {settings:{account:'TaskTrack',currentUser:'u1',currency:'$',language:'en',memberRoles:Object.fromEntries(team.map((p,i)=>[p.id,i?'member':'admin'])),notificationPrefs:{...NOTIFICATION_DEFAULTS},assistantActions:{...ASSISTANT_ACTION_DEFAULTS},guidedTourDone:true,assistantPet:{...ASSISTANT_PET_DEFAULTS}},team,boards,
    docs:[{id:'doc1',title:'Lumen Grove campaign brief',template:'Project Brief',content:'## Goal\n\nCoordinate the fictional Lumen Grove launch.\n\n## Owners\n\nMarketing prepares the brief; Ecommerce reviews the product page; Operations checks sample inventory.\n\n## Review\n\nAll email, calendar and AI actions in this portfolio are local simulations.',updated:Date.now()},{id:'doc2',title:'Weekly team meeting',template:'Meeting Notes',content:'## Agenda\n\nReview open tasks, blockers and next approvals.\n\n## Decisions\n\nKeep proposed campaign changes in review until the owner approves.',updated:Date.now()}],
    automations:[{id:'a1',name:'Move done tasks to Completed',active:true,boardId:'b1',trigger:'status_done',action:'move_group',targetGroupId:'g0done'}],integrations:[],assets:[],notifications:[],notificationPrefs:{},boardNotificationPrefs:{},updates:[{id:'upd1',authorId:'u1',boardId:'b1',groupId:'g0active',taskId:'t00',body:'@Jordan Vale, please review the fictional campaign brief before approving the next step.',mentions:['u2'],readBy:['u1'],bookmarkedBy:[],createdAt:Date.now()-3600000}],archive:[]};
}

/* ---------- state ---------- */
let DATA;
let OPEN_SUBTASK_TASKS = new Set();
function localDataKey(){
  return WORKSPACE_ID ? `tasktrack_data_${WORKSPACE_ID}` : 'tasktrack_data';
}
function load(){
  const raw = Store.get(localDataKey());
  if(raw){ try{ DATA = JSON.parse(raw); }catch(e){ DATA = seed(); } }
  else DATA = seed();
  if(!DATA.boards) DATA.boards = [];
  if(!DATA.team) DATA.team = [];
  if(!DATA.docs) DATA.docs = [];
  if(!Array.isArray(DATA.automations)) DATA.automations = [];
  if(!Array.isArray(DATA.integrations)) DATA.integrations = [];
  if(!Array.isArray(DATA.assets)) DATA.assets = [];
  if(!Array.isArray(DATA.updates)) DATA.updates = [];
  if(!Array.isArray(DATA.notifications)) DATA.notifications = [];
  if(!DATA.notificationPrefs || typeof DATA.notificationPrefs !== 'object') DATA.notificationPrefs = {};
  if(!DATA.boardNotificationPrefs || typeof DATA.boardNotificationPrefs !== 'object') DATA.boardNotificationPrefs = {};
  if(!Array.isArray(DATA.archive)) DATA.archive = [];
  normalizeData();
  if(applyRecurringTaskRollovers(DATA)) saveLocal();
}
function saveLocal(){ Store.set(localDataKey(), JSON.stringify(DATA)); }
function save(){ saveLocal(); queueCloudSave(); }

let VIEW = 'home';        // home | board | calendar | dashboard | mywork | integrations | docs | settings
let CURRENT_BOARD = null; // board id
let SEARCH = '';
let BOARD_VIEW = 'all';
let BOARD_LAYOUT = 'table';
let BOARD_PERSON_FILTERS = {};
let GLOBAL_SEARCH = '';
let UPDATE_FILTER = 'all';
let UPDATE_BOARD_FILTER = 'all';
let UPDATE_SHOW_UNREAD = false;
let FEED_COMPOSER_BOARD = '';
let FEED_COMPOSER_TASK = '';
let FEED_FILES = [];
let FEED_POSTING = false;
let CLOSING_MONTH = new Date().toISOString().slice(0,7);
let MY_WORK_COLLAPSED = new Set();
let MY_WORK_SORT = 'due';
let openDD = null;        // open dropdown reference
let CAL_MONTH = new Date().toISOString().slice(0,7);
let CALENDAR_GROUP_FILTER = 'all';
let draggingColumn = null;
let draggingTask = null;
let draggingSubtask = null;
let taskDragFinishedAt = 0;
let editableCellSelectionGuardUntil = 0;
let editableCellSelectionGesture = null;
let draggingGroup = null;
let draggingKanbanTask = null;
let SELECTED_TASKS = new Set();
let SESSION = null;
let AUTH_SUBSCRIPTION = null;
let CLOUD_USER_ID = null;
let WORKSPACE_ID = null;
let WORKSPACE_OWNER_ID = null;
let WORKSPACE_ROLE = 'member';
let WORKSPACE_UPDATED_AT = null;
let WORKSPACE_CHOICES = [];
let WORKSPACE_CHANNEL = null;
let WORKSPACE_POLL_TIMER = null;
let CURRENT_MEMBER_ID = null;
let CLOUD_STATUS = isSupabaseConfigured ? 'Connecting' : 'Local only';
let CLOUD_SAVE_TIMER = null;
let CLOUD_LOADING = false;
let APPLYING_REMOTE = false;
let AUTH_MODE = 'signin';
let INTEGRATION_CONNECTIONS = JSON.parse(localStorage.getItem('portfolio_integrations') || '[]');
let PRIVATE_GOOGLE_EVENTS = [];
let SHARED_GOOGLE_EVENTS = [];
let SHARED_GOOGLE_READS = new Set();
let PRIVATE_GOOGLE_SHARED_IDS = new Set();
let MY_GOOGLE_CALENDAR_ON = false;
let PRIVATE_GOOGLE_LOAD_KEY = '';
let PRIVATE_GOOGLE_LOADING = false;
let PRIVATE_GOOGLE_ERROR = '';
let SHARED_GOOGLE_LOAD_KEY = '';
let SHARED_GOOGLE_LOADING = false;
let NOTION_AI_PREVIEW = null;
let DOCS_FILTER = 'all';
let MEETING_ROWS = [];
let MEETING_SELECTED_ID = '';
let MEETING_DETAIL = null;
let MEETING_LIST_LOADED = false;
let MEETING_LOADING = false;
let MEETING_SEARCH = '';
let MEETING_SHOW_ARCHIVED = false;
let MEETING_TAB = 'overview';
let MEETING_REPOSITORY = null;
let MEETING_CHANNEL = null;
let MEETING_SYNC_TIMER = null;
let MEETING_SELECTED_ACTIONS = new Set();
let MEETING_ERROR = '';
let MEETING_DETAIL_REQUEST = 0;
let MEETING_PENDING_CAPTURES = [];
let MEETING_ACTIVE_CAPTURE = null;
// Deliberately process-local: reloads and duplicated tabs must prove the prior
// server lease is stale instead of inheriting control through shared storage.
const MEETING_CAPTURE_BROWSER_INSTANCE_ID = globalThis.crypto?.randomUUID?.()
  || `browser_${Date.now()}_${Math.random().toString(36).slice(2)}`;
let GOOGLE_MEET_SUBSCRIPTION = null;
let ASSISTANT_OPEN = false;
let ASSISTANT_STARTED = false;
let ASSISTANT_LOADING = false;
let ASSISTANT_ACTION_BUSY = false;
let ASSISTANT_MESSAGES = [];
let ASSISTANT_LAST_SCROLL_KEY = '';
let TASK_DETAIL = null;
let TASK_DETAIL_TAB = 'updates';
let TASK_EMAIL_SENDING = false;
let CALLBACK_NOTICE_SHOWN = false;
let LAST_SYNC_DATA = null;
let ROW_REFRESH_TIMER = null;
let TOUR_SHOWN_THIS_SESSION = false;
let PENDING_REMOTE_WORKSPACE = null;
let PENDING_REMOTE_REVIEW = null;
let PENDING_REMOTE_BANNER_HIDDEN = false;
let PENDING_REMOTE_PATCHES = [];
let REMOTE_RENDER_PENDING = false;
let REMOTE_RENDER_TIMER = null;
let PRESENCE_CHANNEL = null;
let REMOTE_PRESENCE = [];
let PRESENCE_TRACK_TIMER = null;
let PRESENCE_TRACK_DEBOUNCE = null;
let RECURRING_ROLLOVER_TIMER = null;
let SERVICE_WORKER_MESSAGE_BOUND = false;
const LOCAL_CLIENT_ID = `client_${Math.random().toString(36).slice(2)}_${Date.now().toString(36)}`;

const ROW_SYNC_TABLES = [
  'app_settings',
  'app_members',
  'app_boards',
  'app_groups',
  'app_columns',
  'app_tasks',
  'app_task_values',
  'app_docs',
  'app_assets',
  'app_automations',
  'app_integrations',
  'app_archive_items',
  'app_updates',
  'app_update_reads',
  'app_update_bookmarks',
  'app_notifications',
  'app_notification_preferences',
  'app_board_notification_preferences'
];
const ROW_UPSERT_ORDER = ROW_SYNC_TABLES.slice();
const ROW_DELETE_ORDER = [
  'app_update_reads',
  'app_update_bookmarks',
  'app_board_notification_preferences',
  'app_notification_preferences',
  'app_task_values',
  'app_updates',
  'app_archive_items',
  'app_automations',
  'app_integrations',
  'app_assets',
  'app_docs',
  'app_tasks',
  'app_columns',
  'app_groups',
  'app_boards',
  'app_members',
  'app_settings'
];
const ROW_CONFLICTS = {
  app_settings:'workspace_id',
  app_task_values:'workspace_id,task_id,column_id',
  app_update_reads:'workspace_id,update_id,member_id',
  app_update_bookmarks:'workspace_id,update_id,member_id',
  app_notification_preferences:'workspace_id,member_id',
  app_board_notification_preferences:'workspace_id,board_id,member_id'
};
const OPTIONAL_ROW_SYNC_TABLES = new Set(['app_assets', 'app_notifications', 'app_notification_preferences', 'app_board_notification_preferences']);
const PRIVATE_DELIVERY_TABLES = new Set(['app_notifications']);
const NOTIFICATION_STARTUP_LIMIT = 200;
const ROW_TABLE_AVAILABLE = Object.fromEntries(ROW_SYNC_TABLES.map(table=>[table, true]));
const EVERYONE_MENTION_ID = '__everyone';
const TASK_SUBSCRIBERS_KEY = '__subscribers';
const TASK_OWNERS_KEY = '__owners';
const TASK_DESCRIPTION_KEY = '__description';
const TASK_RECURRING_KEY = '__recurring';
const TASK_RECURRENCE_KEY = '__recurrence';
const TASK_SUBTASKS_OPEN_KEY = '__subtasksOpen';
const TASK_GOOGLE_CALENDAR_EVENT_KEY = '__googleCalendarEventId';
const TASK_GOOGLE_CALENDAR_LINK_KEY = '__googleCalendarLink';
const TASK_GOOGLE_CALENDAR_SYNCED_KEY = '__googleCalendarSyncedAt';
const TASK_GOOGLE_CALENDAR_EVENTS_KEY = '__googleCalendarEventsByOwner';
const TASK_DUE_TIME_KEY = '__dueTime';
const TASK_TIMELINE_START_TIME_KEY = '__timelineStartTime';
const TASK_TIMELINE_END_TIME_KEY = '__timelineEndTime';
const TASK_DATE_REMINDER_KEY = '__dateReminder';
const TASK_COLOR_KEY = '__taskColor';
const INTERNAL_TASK_VALUE_COLUMNS = [
  { id:TASK_SUBSCRIBERS_KEY, name:'Task subscribers' },
  { id:TASK_OWNERS_KEY, name:'Task owners' },
  { id:TASK_DESCRIPTION_KEY, name:'Task description' },
  { id:TASK_RECURRING_KEY, name:'Task recurring' },
  { id:TASK_RECURRENCE_KEY, name:'Task recurrence' },
  { id:TASK_GOOGLE_CALENDAR_EVENT_KEY, name:'Google Calendar event id' },
  { id:TASK_GOOGLE_CALENDAR_LINK_KEY, name:'Google Calendar event link' },
  { id:TASK_GOOGLE_CALENDAR_SYNCED_KEY, name:'Google Calendar synced at' },
  { id:TASK_GOOGLE_CALENDAR_EVENTS_KEY, name:'Google Calendar owner event map' },
  { id:TASK_DUE_TIME_KEY, name:'Task due time' },
  { id:TASK_TIMELINE_START_TIME_KEY, name:'Task timeline start time' },
  { id:TASK_TIMELINE_END_TIME_KEY, name:'Task timeline end time' },
  { id:TASK_DATE_REMINDER_KEY, name:'Task date reminder' },
  { id:TASK_COLOR_KEY, name:'Task color' }
];
const UPDATE_FILTERS = new Set(['all', 'mentions']);
const DEFAULT_COLUMN_WIDTHS = {
  name:260,
  owner:112,
  status:130,
  color:104,
  due:132,
  timeline:178,
  priority:126,
  budget:150,
  notes:240,
  assignedBy:150
};

/* ---------- helpers ---------- */
const $ = (sel,root=document)=>root.querySelector(sel);
const uid = (p='id')=> p+Math.random().toString(36).slice(2,9);
const esc = s => (s==null?'':String(s)).replace(/[&<>"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]));
const member = id => DATA.team.find(m=>m.id===id);
const board = id => DATA.boards.find(b=>b.id===id);
const fmtMoney = n => (DATA.settings.currency||'$') + (Number(n)||0).toLocaleString();
function normalizedUrlHref(value){
  const raw = String(value || '').trim();
  if(!raw) return '';
  const candidate = /^[a-z][a-z0-9+.-]*:/i.test(raw) ? raw : `https://${raw}`;
  try{
    const url = new URL(candidate);
    return ['http:', 'https:'].includes(url.protocol) ? url.href : '';
  }catch(e){
    return '';
  }
}
function canonicalUrlKey(value){
  const href = normalizedUrlHref(value);
  if(!href) return '';
  try{
    const url = new URL(href);
    const host = url.hostname.toLowerCase().replace(/^www\./,'');
    const path = url.pathname.replace(/\/+$/,'') || '/';
    const ytId = youtubeVideoId(url, host);
    if(ytId) return `video:youtube:${ytId}`;
    const vimeoId = vimeoVideoId(url, host);
    if(vimeoId) return `video:vimeo:${vimeoId}`;
    const tiktokId = matchPathId(path, /\/video\/([^/?#]+)/i);
    if(host.endsWith('tiktok.com') && tiktokId) return `video:tiktok:${tiktokId}`;
    const instagramMatch = host.endsWith('instagram.com') ? path.match(/^\/(reel|p|tv)\/([^/]+)/i) : null;
    if(instagramMatch) return `video:instagram:${instagramMatch[1].toLowerCase()}:${instagramMatch[2]}`;
    const driveId = googleDriveItemId(url, host);
    if(driveId) return `file:google:${driveId}`;
    url.hash = '';
    const keptParams = [...url.searchParams.entries()]
      .filter(([key])=>!isTrackingUrlParam(key))
      .sort(([a],[b])=>a.localeCompare(b));
    const query = keptParams.length ? `?${new URLSearchParams(keptParams).toString()}` : '';
    return `${url.protocol}//${host}${path}${query}`;
  }catch(e){
    return href.toLowerCase();
  }
}
function youtubeVideoId(url, host){
  if(host === 'youtu.be') return url.pathname.split('/').filter(Boolean)[0] || '';
  if(!host.endsWith('youtube.com')) return '';
  const byParam = url.searchParams.get('v');
  if(byParam) return byParam;
  return matchPathId(url.pathname, /^\/(?:shorts|embed|live)\/([^/?#]+)/i);
}
function vimeoVideoId(url, host){
  if(!host.endsWith('vimeo.com')) return '';
  const part = url.pathname.split('/').filter(Boolean).find(x=>/^\d+$/.test(x));
  return part || '';
}
function googleDriveItemId(url, host){
  if(host === 'drive.google.com'){
    const openId = url.searchParams.get('id');
    if(openId) return openId;
    return matchPathId(url.pathname, /^\/(?:file\/d|drive\/folders)\/([^/?#]+)/i);
  }
  if(host.endsWith('docs.google.com')){
    const match = url.pathname.match(/^\/(document|spreadsheets|presentation|forms)\/d\/([^/?#]+)/i);
    return match ? `${match[1].toLowerCase()}:${match[2]}` : '';
  }
  return '';
}
function matchPathId(path, re){
  const match = String(path || '').match(re);
  return match ? decodeURIComponent(match[1]) : '';
}
function isTrackingUrlParam(key){
  const k = String(key || '').toLowerCase();
  return k.startsWith('utm_') || ['fbclid','gclid','gbraid','wbraid','mc_cid','mc_eid','igshid','si','feature','ref','share','app'].includes(k);
}
const URL_LIKE_RE = /https?:\/\/[^\s<>"']+|www\.[^\s<>"']+|(?:[a-z0-9](?:[a-z0-9-]{0,61}[a-z0-9])?\.)+[a-z]{2,}(?:\/[^\s<>"']*)?/gi;
function cleanUrlMatch(raw){
  let text = String(raw || '');
  let trailing = '';
  while(text && /[.,!?;:)\]}]+$/.test(text)){
    trailing = text.slice(-1) + trailing;
    text = text.slice(0,-1);
  }
  return {text, trailing};
}
function extractLinksFromText(text){
  const str = String(text || '');
  const links = [];
  URL_LIKE_RE.lastIndex = 0;
  let match;
  while((match = URL_LIKE_RE.exec(str))){
    const previous = match.index > 0 ? str[match.index - 1] : '';
    if(previous === '@') continue;
    const cleaned = cleanUrlMatch(match[0]);
    if(!cleaned.text) continue;
    const href = normalizedUrlHref(cleaned.text);
    if(!href) continue;
    links.push({
      raw:cleaned.text,
      href,
      start:match.index,
      end:match.index + cleaned.text.length
    });
  }
  return links;
}
function renderLinkedText(text,{lineBreaks=false}={}){
  const str = String(text || '');
  const links = extractLinksFromText(str);
  if(!links.length) return lineBreaks ? esc(str).replace(/\n/g,'<br>') : esc(str);
  let html = '';
  let last = 0;
  links.forEach(link=>{
    html += esc(str.slice(last, link.start));
    html += `<a class="inline-link" href="${esc(link.href)}" target="_blank" rel="noopener noreferrer">${esc(link.raw)}</a>`;
    last = link.end;
  });
  html += esc(str.slice(last));
  return lineBreaks ? html.replace(/\n/g,'<br>') : html;
}
function firstLinkHref(text){
  return extractLinksFromText(text)[0]?.href || '';
}
function fieldLinkOpenHtml(value){
  const href = firstLinkHref(value);
  return href ? `<a class="field-link-open" href="${esc(href)}" target="_blank" rel="noopener noreferrer" title="${esc(href)}">Open</a>` : '';
}
function copyCellButton(g,t,field,value){
  if(!String(value || '').trim()) return '';
  return `<button type="button" class="cell-copy-btn" data-copy-cell="${esc(g.id)}|${esc(t.id)}|${esc(field)}" title="Copy cell content">${I.copy}</button>`;
}
const normalizeEmail = email => String(email||'').trim().toLowerCase();
const sessionEmail = () => normalizeEmail(SESSION?.user?.email || '');
function withTimeout(promise, message='Request timed out', ms=15000){
  let timer;
  const timeout = new Promise((_, reject)=>{
    timer = setTimeout(()=>reject(new Error(message)), ms);
  });
  return Promise.race([promise, timeout]).finally(()=>clearTimeout(timer));
}
function isEditingControl(active=document.activeElement){
  if(openDD?.dataset?.editType) return true;
  if(!active || active === document.body) return false;
  if(openDD && openDD.contains(active)) return true;
  if(active.closest?.('.modal-doc-editor,.task-drawer')) return active.matches?.('input,select,textarea,[contenteditable],.doc-page');
  return Boolean(active.matches?.('input,select,textarea,[contenteditable]'));
}
function isTextEditing(){
  const active = document.activeElement;
  return isEditingControl(active);
}
function currentEditingScope(){
  const target = currentEditingTarget();
  if(target) return {type:target.type, id:target.id};
  return null;
}
function currentEditingTarget(){
  const active = document.activeElement;
  const docEditor = active?.closest?.('[data-doc-editor-id]') || $('#modal-root [data-doc-editor-id]');
  if(docEditor?.dataset?.docEditorId && active?.closest?.('[data-doc-editor-id]')){
    return {type:'doc', id:docEditor.dataset.docEditorId, field:active?.id === 'doc-title' ? 'title' : 'content'};
  }
  if(TASK_DETAIL && active?.closest?.('.task-drawer')){
    if(active?.matches?.('[data-task-detail-name]')) return {type:'task', id:TASK_DETAIL.taskId, field:'name'};
    if(active?.id === 'task-detail-description') return {type:'task', id:TASK_DETAIL.taskId, field:`custom:${TASK_DESCRIPTION_KEY}`};
    return {type:'task', id:TASK_DETAIL.taskId, field:'detail'};
  }
  if(active?.dataset?.task) return {type:'task', id:active.dataset.task, field:active.dataset.edit || 'task'};
  if(openDD?.dataset?.editType && openDD.dataset.editId){
    return {type:openDD.dataset.editType, id:openDD.dataset.editId, field:openDD.dataset.editField || openDD.dataset.editType};
  }
  return null;
}
function currentViewingTarget(){
  if(TASK_DETAIL?.taskId) return {type:'task', id:TASK_DETAIL.taskId, field:'viewing'};
  const docEditor = $('#modal-root [data-doc-editor-id]');
  if(docEditor?.dataset?.docEditorId) return {type:'doc', id:docEditor.dataset.docEditorId, field:'viewing'};
  if(VIEW === 'board' && CURRENT_BOARD) return {type:'board', id:CURRENT_BOARD, field:'viewing'};
  return {type:VIEW || 'home', id:CURRENT_BOARD || WORKSPACE_ID || 'workspace', field:'viewing'};
}
function dropdownEditScope(anchor){
  if(!anchor) return null;
  const taskAttrs = {
    statusPick:'status',
    priorityPick:'priority',
    ownerPick:'owner',
    datePick:'timeline',
    tlPick:'timeline',
    customDatePick:'custom',
    customStatusPick:'custom',
    customRecurrencePick:'custom',
    fileCellPick:'custom',
    docCellPick:'custom'
  };
  for(const [attr, field] of Object.entries(taskAttrs)){
    const raw = anchor.dataset?.[attr];
    if(raw){
      const parts = String(raw).split('|');
      return {type:'task', id:parts[0], field:field === 'custom' ? `custom:${parts[2] || ''}` : field};
    }
  }
  if(TASK_DETAIL && anchor.closest?.('.task-drawer')) return {type:'task', id:TASK_DETAIL.taskId, field:'detail'};
  const docEditor = anchor.closest?.('[data-doc-editor-id]');
  if(docEditor?.dataset?.docEditorId) return {type:'doc', id:docEditor.dataset.docEditorId, field:'content'};
  return null;
}
function presenceForTarget(type,id){
  return REMOTE_PRESENCE.filter(p=>p?.target?.type === type && p?.target?.id === id);
}
function presenceBadgesHtml(type,id){
  const people = presenceForTarget(type,id).slice(0,3);
  if(!people.length) return '';
  const title = people.map(p=>`${p.name || 'Teammate'} ${p.target?.field === 'viewing' ? 'is viewing' : 'is editing'}`).join(', ');
  return `<span class="presence-stack" title="${esc(title)}">
    ${people.map(p=>`<span class="presence-dot" style="background:${esc(p.color || '#0073ea')}">${esc(p.initials || initialsFor(p.name))}</span>`).join('')}
    ${presenceForTarget(type,id).length > people.length ? `<span class="presence-more">+${presenceForTarget(type,id).length - people.length}</span>` : ''}
  </span>`;
}
function currentMemberId(){ return CURRENT_MEMBER_ID || DATA?.settings?.currentUser || DATA?.team?.[0]?.id || ''; }
function appLanguage(){
  const lang = DATA?.settings?.language || 'en';
  return LANGUAGE_OPTIONS[lang] ? lang : 'en';
}
function tr(key){
  const lang = appLanguage();
  return TEXT[lang]?.[key] || TEXT.en[key] || key;
}
function notificationPrefs(memberId=currentMemberId(), {useWorkspaceFallback=true} = {}){
  const base = {
    ...NOTIFICATION_DEFAULTS,
    ...(useWorkspaceFallback ? (DATA?.settings?.notificationPrefs || {}) : {})
  };
  const memberPrefs = memberId ? DATA?.notificationPrefs?.[memberId] : null;
  return {...base, ...(memberPrefs || {})};
}
function setNotificationPrefsForMember(memberId, prefs){
  if(!memberId) return;
  if(!DATA.notificationPrefs || typeof DATA.notificationPrefs !== 'object') DATA.notificationPrefs = {};
  DATA.notificationPrefs[memberId] = {...notificationPrefs(memberId), ...(prefs || {})};
}
function boardNotificationPrefs(boardId='', memberId=currentMemberId()){
  const boardMap = DATA?.boardNotificationPrefs?.[boardId] || {};
  const pref = memberId ? boardMap[memberId] : null;
  return {
    mode: pref?.mode || 'custom',
    prefs: pref?.prefs && typeof pref.prefs === 'object' ? pref.prefs : {}
  };
}
function setBoardNotificationPrefs(boardId, memberId, value){
  if(!boardId || !memberId) return;
  if(!DATA.boardNotificationPrefs || typeof DATA.boardNotificationPrefs !== 'object') DATA.boardNotificationPrefs = {};
  if(!DATA.boardNotificationPrefs[boardId]) DATA.boardNotificationPrefs[boardId] = {};
  DATA.boardNotificationPrefs[boardId][memberId] = {
    mode:['custom','receive_all','mute'].includes(value?.mode) ? value.mode : 'custom',
    prefs:value?.prefs && typeof value.prefs === 'object' ? value.prefs : {}
  };
}
function desktopNotificationsSupported(){
  return typeof window !== 'undefined' && 'Notification' in window;
}
function desktopNotificationsGranted(){
  return desktopNotificationsSupported() && Notification.permission === 'granted';
}
function desktopNotificationPrefKey(){
  return `friday_desktop_notifications_${WORKSPACE_ID || 'local'}_${currentMemberId() || 'user'}`;
}
function desktopNotificationsEnabled(){
  return notificationPrefs().desktop !== false && desktopNotificationsGranted() && Store.get(desktopNotificationPrefKey()) === '1';
}
function setDesktopNotificationsEnabled(enabled){
  Store.set(desktopNotificationPrefKey(), enabled ? '1' : '0');
}
function webPushSupported(){
  return Boolean(
    typeof window !== 'undefined'
    && 'serviceWorker' in navigator
    && 'PushManager' in window
    && desktopNotificationsSupported()
    && supabaseConfig.vapidPublicKey
    && supabase
    && SESSION?.user
    && WORKSPACE_ID
  );
}
function webPushConfigured(){
  return Boolean(supabaseConfig.vapidPublicKey);
}
function webPushPrefKey(){
  return `friday_web_push_${WORKSPACE_ID || 'local'}_${currentMemberId() || 'user'}`;
}
function webPushEnabled(){
  return webPushSupported() && notificationPrefs().push !== false && Store.get(webPushPrefKey()) === '1';
}
function setWebPushEnabled(enabled){
  Store.set(webPushPrefKey(), enabled ? '1' : '0');
}
function bindServiceWorkerMessages(){
  if(SERVICE_WORKER_MESSAGE_BOUND || !('serviceWorker' in navigator)) return;
  SERVICE_WORKER_MESSAGE_BOUND = true;
  navigator.serviceWorker.addEventListener('message', event=>{
    if(event.data?.type !== 'friday-notification-click') return;
    const notificationId = event.data.notificationId || '';
    if(notificationId && DATA?.notifications?.some(n=>n.id === notificationId)){
      openNotificationTarget(notificationId);
      return;
    }
    const taskId = event.data.taskId || '';
    if(taskId){
      const loc = findTaskLocationById(taskId);
      if(loc) openTaskDetail(loc.b.id, loc.g.id, loc.t.id, 'updates');
    }
  });
}
function urlBase64ToUint8Array(value){
  const padding = '='.repeat((4 - value.length % 4) % 4);
  const base64 = (value + padding).replace(/-/g, '+').replace(/_/g, '/');
  const raw = window.atob(base64);
  const out = new Uint8Array(raw.length);
  for(let i=0;i<raw.length;i++) out[i] = raw.charCodeAt(i);
  return out;
}
async function sha256Base64Url(value){
  const bytes = new TextEncoder().encode(String(value || ''));
  const hash = await crypto.subtle.digest('SHA-256', bytes);
  return btoa(String.fromCharCode(...new Uint8Array(hash))).replace(/\+/g,'-').replace(/\//g,'_').replace(/=+$/,'');
}
async function fridayServiceWorkerRegistration(){
  if(!('serviceWorker' in navigator)) return null;
  return navigator.serviceWorker.register('./friday-sw.js', {scope:'./'});
}
async function savePushSubscription(subscription){
  if(!subscription || !supabase || !WORKSPACE_ID || !SESSION?.user) return false;
  const json = subscription.toJSON();
  const endpoint = json.endpoint || subscription.endpoint || '';
  const p256dh = json.keys?.p256dh || '';
  const auth = json.keys?.auth || '';
  if(!endpoint || !p256dh || !auth) return false;
  const me = member(currentMemberId()) || {};
  const id = `push_${await sha256Base64Url(endpoint)}`;
  const { error } = await supabase
    .from('app_push_subscriptions')
    .upsert({
      workspace_id:WORKSPACE_ID,
      id,
      member_id:currentMemberId(),
      user_id:SESSION.user.id,
      email:normalizeEmail(me.email || SESSION.user.email),
      endpoint,
      p256dh,
      auth,
      user_agent:navigator.userAgent || ''
    }, {onConflict:'workspace_id,id'});
  if(error) throw error;
  return true;
}
async function registerWebPushSubscription(){
  if(!webPushSupported()) return false;
  const registration = await fridayServiceWorkerRegistration();
  if(!registration) return false;
  let subscription = await registration.pushManager.getSubscription();
  if(!subscription){
    subscription = await registration.pushManager.subscribe({
      userVisibleOnly:true,
      applicationServerKey:urlBase64ToUint8Array(supabaseConfig.vapidPublicKey)
    });
  }
  await savePushSubscription(subscription);
  setWebPushEnabled(true);
  return true;
}
async function removeWebPushSubscription(){
  if(!('serviceWorker' in navigator) || !supabase || !WORKSPACE_ID || !SESSION?.user) return;
  try{
    const registration = await navigator.serviceWorker.getRegistration('./');
    const subscription = await registration?.pushManager?.getSubscription?.();
    const endpoint = subscription?.endpoint || '';
    if(subscription) await subscription.unsubscribe();
    if(endpoint){
      await supabase
        .from('app_push_subscriptions')
        .delete()
        .eq('workspace_id', WORKSPACE_ID)
        .eq('endpoint', endpoint);
    }
  }catch(e){}
  setWebPushEnabled(false);
}
async function requestDesktopNotifications(){
  if(!desktopNotificationsSupported()){
    toast('Desktop notifications are not supported in this browser');
    return;
  }
  const permission = await Notification.requestPermission();
  setDesktopNotificationsEnabled(permission === 'granted');
  if(permission === 'granted' && notificationPrefs().push !== false && webPushSupported()){
    try{ await registerWebPushSubscription(); }catch(e){ console.warn('Web push subscription failed', e); }
  }
  renderContent();
  toast(permission === 'granted' ? 'Desktop notifications enabled' : 'Desktop notifications were not enabled');
}
function assistantActionPrefs(){
  return {...ASSISTANT_ACTION_DEFAULTS, ...(DATA?.settings?.assistantActions || {})};
}
function currentWorkspaceRole(){
  if(WORKSPACE_OWNER_ID && SESSION?.user?.id === WORKSPACE_OWNER_ID) return 'owner';
  return ['owner','admin','member'].includes(WORKSPACE_ROLE) ? WORKSPACE_ROLE : 'member';
}
function isWorkspaceAdmin(){
  return currentWorkspaceRole() === 'owner' || currentWorkspaceRole() === 'admin';
}
function memberRole(id){
  const roles = DATA?.settings?.memberRoles || {};
  return roles[id] === 'admin' ? 'admin' : 'member';
}
function currentMemberRole(){ return memberRole(currentMemberId()); }
function isCurrentAdmin(){ return isWorkspaceAdmin(); }
function setMemberRole(id, role){
  if(!DATA.settings.memberRoles || typeof DATA.settings.memberRoles !== 'object') DATA.settings.memberRoles = {};
  DATA.settings.memberRoles[id] = role === 'admin' ? 'admin' : 'member';
}
function setCurrentMemberId(id){
  CURRENT_MEMBER_ID = id || '';
  if(CLOUD_USER_ID) Store.set(`tasktrack_member_${CLOUD_USER_ID}`, CURRENT_MEMBER_ID);
}
function setSelectedWorkspaceId(id){
  if(CLOUD_USER_ID && id) Store.set(`tasktrack_workspace_${CLOUD_USER_ID}`, id);
}
function selectedWorkspaceId(){
  return CLOUD_USER_ID ? Store.get(`tasktrack_workspace_${CLOUD_USER_ID}`) : '';
}
const VIEW_ROUTES = {
  home:'home',
  updates:'updates',
  closing:'monthly-closing',
  dashboard:'dashboard',
  calendar:'calendar',
  mywork:'my-work',
  meetings:'meetings',
  integrations:'integrations',
  docs:'docs',
  settings:'settings'
};
const ROUTE_VIEWS = Object.fromEntries(Object.entries(VIEW_ROUTES).map(([view,route])=>[route,view]));
function routeHref(view){ return `#${VIEW_ROUTES[view] || 'home'}`; }
function updateRouteHash(view){
  const route = VIEW_ROUTES[view];
  if(!route) return;
  const url = new URL(window.location.href);
  url.hash = route;
  window.history.replaceState({}, '', url.toString());
}
function navigateToView(view){
  if(!VIEW_ROUTES[view]) return;
  VIEW = view;
  updateRouteHash(view);
  render();
  if(view === 'meetings') void ensureMeetingListLoaded();
}
function displayNameFromEmail(email){ return normalizeEmail(email).split('@')[0] || 'You'; }
function initialsFor(name){
  return String(name||'').split(/\s+/).filter(Boolean).map(w=>w[0]).join('').slice(0,2).toUpperCase() || 'YO';
}
function fmtDate(iso){ if(!iso) return ''; const d=new Date(iso+'T00:00'); if(isNaN(d)) return ''; 
  return d.toLocaleDateString(undefined,{month:'short',day:'numeric'}); }
function todayIso(){
  const d = new Date();
  d.setMinutes(d.getMinutes() - d.getTimezoneOffset());
  return d.toISOString().slice(0,10);
}
function validIsoDate(value){
  const raw = String(value || '').trim();
  if(!/^\d{4}-\d{2}-\d{2}$/.test(raw)) return '';
  const d = new Date(`${raw}T00:00:00`);
  return Number.isNaN(d.getTime()) ? '' : raw;
}
function isoDateAdd(iso, days){
  const d = new Date(`${iso}T00:00:00`);
  if(Number.isNaN(d.getTime())) return '';
  d.setDate(d.getDate() + Number(days || 0));
  return d.toISOString().slice(0,10);
}
function isoDateCompare(a,b){ return String(a || '').localeCompare(String(b || '')); }
function isoDateRange(startIso,endIso){
  let start = validIsoDate(startIso);
  let end = validIsoDate(endIso) || start;
  if(!start && end) start = end;
  if(!start) return [];
  if(isoDateCompare(start,end) > 0) [start,end] = [end,start];
  const out = [];
  for(let cur=start, guard=0; cur && isoDateCompare(cur,end) <= 0 && guard < 370; cur=isoDateAdd(cur,1), guard+=1){
    out.push(cur);
  }
  return out;
}
function isoDateDiffDays(startIso,endIso){
  const start = new Date(`${validIsoDate(startIso)}T00:00:00`);
  const end = new Date(`${validIsoDate(endIso)}T00:00:00`);
  if(Number.isNaN(start.getTime()) || Number.isNaN(end.getTime())) return 0;
  return Math.max(0, Math.round((end - start) / 86400000));
}
function localIsoDateFromValue(value){
  const d = value instanceof Date ? new Date(value) : new Date(value);
  if(Number.isNaN(d.getTime())) return '';
  d.setMinutes(d.getMinutes() - d.getTimezoneOffset());
  return d.toISOString().slice(0,10);
}
function validTimeValue(value){
  const raw = String(value || '').trim();
  return /^([01]\d|2[0-3]):[0-5]\d$/.test(raw) ? raw : '';
}
function fmtTimeValue(value){
  const raw = validTimeValue(value);
  if(!raw) return '';
  const [hh,mm] = raw.split(':').map(Number);
  const suffix = hh >= 12 ? 'PM' : 'AM';
  const hour = hh % 12 || 12;
  return `${hour}:${String(mm).padStart(2,'0')} ${suffix}`;
}
function fmtDateTime(iso,time){
  const date = fmtDate(iso);
  const formattedTime = fmtTimeValue(time);
  return [date, formattedTime].filter(Boolean).join(', ');
}
function taskMetaValue(t,key){ return (t?.custom || {})[key] || ''; }
function setTaskMetaValue(t,key,value){
  if(!t) return;
  if(!t.custom) t.custom = {};
  const clean = String(value || '').trim();
  if(clean) t.custom[key] = clean;
  else delete t.custom[key];
}
function taskDueLabel(t){
  return fmtDateTime(t?.due, taskMetaValue(t,TASK_DUE_TIME_KEY));
}
function taskTimelineLabel(t,start=t?.tlStart,end=t?.tlEnd){
  const first = validIsoDate(start);
  const last = validIsoDate(end) || first;
  if(!first) return '';
  const startLabel = fmtDateTime(first, taskMetaValue(t,TASK_TIMELINE_START_TIME_KEY) || taskMetaValue(t,TASK_DUE_TIME_KEY));
  const endLabel = last && last !== first ? fmtDateTime(last, taskMetaValue(t,TASK_TIMELINE_END_TIME_KEY) || taskMetaValue(t,TASK_DUE_TIME_KEY)) : '';
  return endLabel ? `${startLabel} - ${endLabel}` : startLabel;
}
function setTaskTimeline(t,{start='',end='',startTime='',endTime='',reminder='none'}={}){
  if(!t) return;
  let cleanStart = validIsoDate(start);
  let cleanEnd = validIsoDate(end) || cleanStart;
  if(cleanStart && cleanEnd && isoDateCompare(cleanStart,cleanEnd) > 0) [cleanStart,cleanEnd] = [cleanEnd,cleanStart];
  t.tlStart = cleanStart || '';
  t.tlEnd = cleanEnd || '';
  const cleanStartTime = validTimeValue(startTime);
  const cleanEndTime = validTimeValue(endTime) || cleanStartTime;
  setTaskMetaValue(t,TASK_TIMELINE_START_TIME_KEY,cleanStartTime);
  setTaskMetaValue(t,TASK_TIMELINE_END_TIME_KEY,cleanEndTime);
  setTaskMetaValue(t,TASK_DATE_REMINDER_KEY,reminder && reminder !== 'none' ? reminder : '');
}
function setTaskDueDate(t,{date='',time='',reminder='none'}={}){
  if(!t) return;
  t.due = validIsoDate(date) || '';
  setTaskMetaValue(t,TASK_DUE_TIME_KEY,validTimeValue(time));
  setTaskMetaValue(t,TASK_DATE_REMINDER_KEY,reminder && reminder !== 'none' ? reminder : '');
}
function clearTaskTimeline(t){
  if(!t) return;
  t.tlStart = '';
  t.tlEnd = '';
  setTaskMetaValue(t,TASK_TIMELINE_START_TIME_KEY,'');
  setTaskMetaValue(t,TASK_TIMELINE_END_TIME_KEY,'');
}
function clearTaskDueDate(t){
  if(!t) return;
  t.due = '';
  setTaskMetaValue(t,TASK_DUE_TIME_KEY,'');
}
function normalizeTaskSchedule(t){
  const start = validIsoDate(t?.tlStart);
  const end = validIsoDate(t?.tlEnd);
  const due = validIsoDate(t?.due);
  if(start || end){
    setTaskTimeline(t,{
      start:start || end,
      end:end || start,
      startTime:taskMetaValue(t,TASK_TIMELINE_START_TIME_KEY),
      endTime:taskMetaValue(t,TASK_TIMELINE_END_TIME_KEY),
      reminder:taskMetaValue(t,TASK_DATE_REMINDER_KEY) || 'none'
    });
  }
  setTaskDueDate(t,{
    date:due,
    time:taskMetaValue(t,TASK_DUE_TIME_KEY),
    reminder:taskMetaValue(t,TASK_DATE_REMINDER_KEY) || 'none'
  });
}
function calendarDateTimeValue(date,time,fallbackTime='09:00'){
  const cleanDate = validIsoDate(date);
  if(!cleanDate) return '';
  const cleanTime = validTimeValue(time) || fallbackTime;
  return `${cleanDate}T${cleanTime}:00`;
}
function calendarDateTimeAddMinutes(date,time,minutes){
  const value = calendarDateTimeValue(date,time);
  if(!value) return '';
  const d = new Date(value);
  if(Number.isNaN(d.getTime())) return value;
  d.setMinutes(d.getMinutes() + Number(minutes || 0));
  const offset = d.getTimezoneOffset();
  const local = new Date(d.getTime() - offset * 60000);
  return local.toISOString().slice(0,16) + ':00';
}
function fmtTimeAgo(ts){
  const n = Number(ts)||Date.now();
  const mins = Math.max(0, Math.round((Date.now()-n)/60000));
  if(mins < 1) return 'Just now';
  if(mins < 60) return `${mins}m ago`;
  const hours = Math.round(mins/60);
  if(hours < 24) return `${hours}h ago`;
  const days = Math.round(hours/24);
  if(days < 14) return `${days}d ago`;
  return new Date(n).toLocaleDateString(undefined,{month:'short',day:'numeric'});
}
function pushTaskWithChildren(out,b,g,t,parent=null){
  const row = {...t,_board:b,_group:g,_parent:parent};
  out.push(row);
  (t.subtasks||[]).forEach(st=>pushTaskWithChildren(out,b,g,st,row));
}
function allTasks(){
  const out=[];
  DATA.boards.forEach(b=>b.groups.forEach(g=>g.tasks.forEach(t=>pushTaskWithChildren(out,b,g,t))));
  return out;
}
function taskIsHidden(t){ return Boolean(t?.hiddenAt); }
function taskHasHiddenAncestor(t){
  let parent = t?._parent || null;
  while(parent){
    if(taskIsHidden(parent)) return true;
    parent = parent._parent || null;
  }
  return false;
}
function taskIsVisible(t){ return !taskIsHidden(t) && !taskHasHiddenAncestor(t); }
function allVisibleTasks(){ return allTasks().filter(taskIsVisible); }
function childTasks(t){ return Array.isArray(t?.subtasks) ? t.subtasks : []; }
function taskTree(t){ return [t, ...childTasks(t).flatMap(taskTree)]; }
function visibleTaskTree(t){
  if(taskIsHidden(t)) return [];
  return [t, ...childTasks(t).flatMap(visibleTaskTree)];
}
function boardTasks(b,{includeHidden=false}={}){
  return (b?.groups || []).flatMap(g=>(g.tasks || []).flatMap(includeHidden ? taskTree : visibleTaskTree));
}
function boardTaskCount(b){ return boardTasks(b).length; }
function boardDoneCount(b){ return boardTasks(b).filter(t=>t.status==='done').length; }
function boardUpdatedAt(b){
  const raw = b?.updated || b?.updatedAt || WORKSPACE_UPDATED_AT;
  const parsed = raw ? Date.parse(raw) : 0;
  return Number.isFinite(parsed) && parsed > 0 ? parsed : Date.now();
}
function taskIdsInGroup(g){
  return (g.tasks || []).flatMap(t=>taskTree(t).map(x=>x.id));
}
function selectedTaskIdsForBoard(boardId){
  pruneSelection();
  const b=board(boardId); if(!b)return [];
  const ids=new Set(b.groups.flatMap(taskIdsInGroup));
  return [...SELECTED_TASKS].filter(id=>ids.has(id));
}
function selectedTaskLocations(){
  pruneSelection();
  return [...SELECTED_TASKS].map(findTaskLocationById).filter(Boolean);
}
function pruneSelection(){
  const ids=new Set();
  DATA.boards.forEach(b=>b.groups.forEach(g=>taskIdsInGroup(g).forEach(id=>ids.add(id))));
  [...SELECTED_TASKS].forEach(id=>{ if(!ids.has(id)) SELECTED_TASKS.delete(id); });
}
function toast(msg){ const t=$('#toast'); t.textContent=msg; t.classList.add('show'); clearTimeout(t._tm); t._tm=setTimeout(()=>t.classList.remove('show'),1900); }

/* ---------- Supabase auth + workspace snapshot ---------- */
function renderLoading(message='Loading workspace'){
  $('#app').innerHTML = `
    <div class="auth-shell">
      <div class="auth-card">
        <div class="auth-mark"><img src="./favicon.png" alt="TaskTrack"></div>
        <h1>${esc(message)}</h1>
        <p>Connecting your task manager workspace.</p>
      </div>
    </div>`;
}

function renderAuth(message=''){
  const isSignup = AUTH_MODE === 'signup';
  $('#app').innerHTML = `
    <div class="auth-shell">
      <form class="auth-card" id="auth-form">
        <div class="auth-mark"><img src="./favicon.png" alt="TaskTrack"></div>
        <h1>${isSignup ? 'Create your workspace account' : 'Sign in to TaskTrack'}</h1>
        <p>Your tasks, boards, docs and automations will save to Supabase.</p>
        ${message ? `<div class="auth-message">${esc(message)}</div>` : ''}
        <div class="field"><label>Email</label><input id="auth-email" type="email" autocomplete="email" required></div>
        <div class="field"><label>Password</label><input id="auth-password" type="password" autocomplete="${isSignup ? 'new-password' : 'current-password'}" minlength="6" required></div>
        <button class="btn auth-primary" type="submit">${isSignup ? 'Create account' : 'Sign in'}</button>
        <button class="auth-switch" type="button" data-auth-mode="${isSignup ? 'signin' : 'signup'}">
          ${isSignup ? 'Already have an account? Sign in' : 'Need an account? Create one'}
        </button>
      </form>
    </div>`;
  wireAuth();
}

function renderWorkspacePicker(choices=[], message=''){
  const rows = choices.map(w=>`
    <button type="button" class="workspace-choice" data-select-workspace="${esc(w.id)}">
      <span class="workspace-choice-mark">${esc(String(w.name || 'TaskTrack').trim().slice(0,1).toUpperCase() || 'F')}</span>
      <span class="workspace-choice-copy">
        <strong>${esc(w.name || 'TaskTrack Workspace')}</strong>
        <small>${workspaceRoleLabel(w)}${w.id ? ` - ${esc(w.id)}` : ''}</small>
      </span>
    </button>`).join('');
  $('#app').innerHTML = `
    <div class="auth-shell">
      <div class="auth-card workspace-picker-card">
        <div class="auth-mark"><img src="./favicon.png" alt="TaskTrack"></div>
        <h1>Choose your workspace</h1>
        <p>${esc(message || 'Your account belongs to more than one TaskTrack workspace.')}</p>
        <div class="workspace-choice-list">${rows}</div>
        <button class="auth-switch" type="button" data-action="workspace-signout">Sign out</button>
      </div>
    </div>`;
  wireWorkspacePicker();
}

function workspaceRoleLabel(workspace){
  const role = workspace?.role || (workspace?.owner_id === SESSION?.user?.id ? 'owner' : 'member');
  if(role === 'owner') return 'Owner';
  if(role === 'admin') return 'Admin';
  return 'Member';
}

function wireWorkspacePicker(){
  document.querySelectorAll('[data-select-workspace]').forEach(el=>el.onclick=()=>{
    setSelectedWorkspaceId(el.dataset.selectWorkspace);
    void loadCloudWorkspace();
  });
  const signout = $('[data-action="workspace-signout"]');
  if(signout) signout.onclick=()=>void signOut();
}

function wireAuth(){
  const form = $('#auth-form');
  if(form) form.onsubmit = e => {
    e.preventDefault();
    void submitAuth();
  };
  document.querySelectorAll('[data-auth-mode]').forEach(el=>el.onclick=()=>{
    AUTH_MODE = el.dataset.authMode;
    renderAuth();
  });
}

async function submitAuth(){
  if(!supabase) return;
  const email = $('#auth-email').value.trim();
  const password = $('#auth-password').value;
  if(!email || !password){ renderAuth('Enter your email and password.'); return; }
  const submit = $('#auth-form button[type="submit"]');
  submit.disabled = true;
  submit.textContent = AUTH_MODE === 'signup' ? 'Creating account...' : 'Signing in...';
  const redirectTo = window.location.href.split('#')[0];
  const result = AUTH_MODE === 'signup'
    ? await supabase.auth.signUp({ email, password, options:{ emailRedirectTo: redirectTo } })
    : await supabase.auth.signInWithPassword({ email, password });
  if(result.error){
    renderAuth(result.error.message || 'Authentication failed.');
    return;
  }
  if(result.data?.session){
    await useSession(result.data.session);
    return;
  }
  renderAuth('Account created. Check your email, confirm the account, then come back here and sign in.');
}

async function bootApp(){
  bindServiceWorkerMessages();
  if(!isSupabaseConfigured || !supabase){
    load();
    applyHashRoute();
    render();
    startRecurringRolloverTimer();
    return;
  }
  renderLoading('Connecting to Supabase');
  const { data: listener } = supabase.auth.onAuthStateChange((_event, session)=>{
    void useSession(session);
  });
  AUTH_SUBSCRIPTION = listener?.subscription || null;
  const { data, error } = await supabase.auth.getSession();
  if(error){
    load();
    CLOUD_STATUS = 'Local only';
    render();
    toast(error.message || 'Could not connect to Supabase');
    return;
  }
  await useSession(data.session);
}

async function useSession(session){
  SESSION = session || null;
  const nextUserId = SESSION?.user?.id || null;
  if(nextUserId && nextUserId === CLOUD_USER_ID && DATA) return;
  CLOUD_USER_ID = nextUserId;
  INTEGRATION_CONNECTIONS = [];
  clearGoogleCalendarState();
  SELECTED_TASKS.clear();
  OPEN_SUBTASK_TASKS.clear();
  if(!SESSION){
    DATA = null;
    resetMeetingState();
    WORKSPACE_ID = null;
    WORKSPACE_OWNER_ID = null;
    WORKSPACE_ROLE = 'member';
    WORKSPACE_UPDATED_AT = null;
    WORKSPACE_CHOICES = [];
    CURRENT_MEMBER_ID = null;
    stopWorkspaceSync();
    stopRecurringRolloverTimer();
    CLOUD_STATUS = 'Signed out';
    renderAuth();
    return;
  }
  try{
    await loadCloudWorkspace();
  }catch(err){
    CLOUD_LOADING = false;
    CLOUD_STATUS = 'Load failed';
    renderAuth(err.message || 'Could not load your workspace. Please refresh and try again.');
  }
}

async function loadCloudWorkspace(){
  CLOUD_LOADING = true;
  CLOUD_STATUS = 'Loading';
  renderLoading('Loading your workspace');
  stopWorkspaceSync();
  stopRecurringRolloverTimer();
  LAST_SYNC_DATA = null;

  const workspaceResult = await resolveWorkspace();
  if(workspaceResult.error){
    CLOUD_LOADING = false;
    renderAuth(workspaceResult.error);
    return;
  }
  if(workspaceResult.choices?.length && !workspaceResult.workspace){
    CLOUD_LOADING = false;
    renderWorkspacePicker(workspaceResult.choices, workspaceResult.message || '');
    return;
  }

  if(WORKSPACE_ID !== workspaceResult.workspace.id){
    resetMeetingState();
    clearGoogleCalendarState();
  }
  WORKSPACE_ID = workspaceResult.workspace.id;
  WORKSPACE_OWNER_ID = workspaceResult.workspace.owner_id;
  WORKSPACE_ROLE = workspaceRole(workspaceResult.workspace, workspaceResult.memberships || []);
  WORKSPACE_CHOICES = workspaceResult.choices || [workspaceResult.workspace];
  setSelectedWorkspaceId(WORKSPACE_ID);

  const rowResult = await loadRowWorkspace(workspaceResult.workspace);
  if(rowResult.error){
    CLOUD_LOADING = false;
    renderAuth(rowResult.error || 'Could not load your shared workspace.');
    return;
  }

  if(rowResult.data?.boards){
    DATA = rowResult.data;
    LAST_SYNC_DATA = JSON.parse(JSON.stringify(DATA));
    WORKSPACE_UPDATED_AT = rowResult.updatedAt || new Date().toISOString();
  } else {
    const legacy = await loadLegacyWorkspaceSnapshot();
    if(legacy?.boards) DATA = legacy;
    else load();
  }

  const sessionMemberChanged = ensureSessionMember();
  normalizeData();
  const recurringChanged = applyRecurringTaskRollovers(DATA);
  saveLocal();
  if(CURRENT_BOARD && !board(CURRENT_BOARD)) CURRENT_BOARD = null;
  if(!CURRENT_BOARD && DATA.boards[0]) CURRENT_BOARD = DATA.boards[0].id;
  applyHashRoute();
  restorePersistedMeetingCapture();
  CLOUD_LOADING = false;
  const needsInitialSeed = !rowResult.data?.boards;
  CLOUD_STATUS = needsInitialSeed ? 'Preparing sync' : 'Synced';
  render();
  void loadWorkspaceSharedCalendar();
  void ensureMeetingListLoaded({ rerender:false });
  startWorkspaceSync();
  startRecurringRolloverTimer();
  void refreshIntegrationConnections({ rerender:true });
  if(needsInitialSeed && isCurrentAdmin()){
    setTimeout(()=>{ void saveCloudNow({ silent:false, full:true }); }, 100);
  } else if(sessionMemberChanged || recurringChanged){
    setTimeout(()=>{ void saveCloudNow({ silent:true }); }, 100);
  }
}

async function resolveWorkspace(){
  const email = sessionEmail();
  try{
    const memberships = await loadWorkspaceMemberships(email);
    const choices = await loadAccessibleWorkspaces(memberships);
    WORKSPACE_CHOICES = choices;

    const selected = selectedWorkspaceId();
    if(selected){
      const choice = choices.find(w=>w.id === selected);
      if(choice){
        await ensureWorkspaceMemberRowIfOwner(choice, choice.role || workspaceRole(choice, memberships));
        return { workspace:choice, choices, memberships };
      }
    }

    if(choices.length === 1){
      await ensureWorkspaceMemberRowIfOwner(choices[0], choices[0].role || workspaceRole(choices[0], memberships));
      setSelectedWorkspaceId(choices[0].id);
      return { workspace:choices[0], choices, memberships };
    }

    if(choices.length > 1){
      return {
        choices,
        message:'Choose which TaskTrack workspace you want to open.'
      };
    }

    if(REQUIRED_WORKSPACE_ID){
      const { data: workspace, error } = await supabase
        .from('workspaces')
        .select('id, owner_id, name, updated_at')
        .eq('id', REQUIRED_WORKSPACE_ID)
        .maybeSingle();
      if(error) throw error;
      if(workspace){
        await ensureWorkspaceMemberRowIfOwner(workspace, workspaceRole(workspace, memberships));
        setSelectedWorkspaceId(workspace.id);
        return { workspace, choices:[workspace], memberships };
      }
      return {
        error:`This deployment is pinned to a TaskTrack workspace that ${email || 'this account'} cannot access. Clear VITE_SHARED_WORKSPACE_ID or invite this email to that workspace.`
      };
    }

    const created = await createWorkspace('TaskTrack');
    return { workspace:created, choices:[created], memberships:[], created:true };
  }catch(err){
    return { error: err.message || 'Could not open a shared workspace.' };
  }
}

async function loadWorkspaceMemberships(email){
  if(!email) return [];
  const { data, error } = await supabase
    .from('workspace_members')
    .select('workspace_id, role, updated_at')
    .eq('email', email)
    .order('updated_at', { ascending:false });
  if(error) throw error;
  return data || [];
}

async function loadAccessibleWorkspaces(memberships=[]){
  const byId = new Map();
  const applyWorkspace = (workspace, source={})=>{
    if(!workspace?.id) return;
    const role = source.role || (workspace.owner_id === SESSION.user.id ? 'owner' : 'member');
    const current = byId.get(workspace.id) || {};
    byId.set(workspace.id, {
      ...current,
      ...workspace,
      role:current.role === 'owner' ? 'owner' : role,
      memberUpdatedAt:source.updated_at || current.memberUpdatedAt || workspace.updated_at || ''
    });
  };

  const membershipIds = [...new Set(memberships.map(m=>m.workspace_id).filter(Boolean))];
  if(membershipIds.length){
    const { data, error } = await supabase
      .from('workspaces')
      .select('id, owner_id, name, updated_at')
      .in('id', membershipIds);
    if(error) throw error;
    (data || []).forEach(workspace=>{
      const source = memberships.find(m=>m.workspace_id === workspace.id) || {};
      applyWorkspace(workspace, source);
    });
  }

  const { data: owned, error: ownedError } = await supabase
    .from('workspaces')
    .select('id, owner_id, name, updated_at')
    .eq('owner_id', SESSION.user.id)
    .order('updated_at', { ascending:false });
  if(ownedError) throw ownedError;
  (owned || []).forEach(workspace=>applyWorkspace(workspace, { role:'owner', updated_at:workspace.updated_at }));

  return [...byId.values()].sort((a,b)=>{
    const ap = Date.parse(a.memberUpdatedAt || a.updated_at || 0) || 0;
    const bp = Date.parse(b.memberUpdatedAt || b.updated_at || 0) || 0;
    return bp - ap || String(a.name || '').localeCompare(String(b.name || ''));
  });
}

function workspaceRole(workspace, memberships=[]){
  if(workspace?.owner_id === SESSION?.user?.id) return 'owner';
  return memberships.find(m=>m.workspace_id === workspace?.id)?.role || workspace?.role || 'member';
}

async function createWorkspace(name='TaskTrack'){
  const { data: created, error } = await supabase
    .from('workspaces')
    .insert({ owner_id: SESSION.user.id, name:String(name || 'TaskTrack').trim() || 'TaskTrack' })
    .select('id, owner_id, name, updated_at')
    .single();
  if(error) throw error;
  await ensureWorkspaceMemberRow(created, 'owner');
  setSelectedWorkspaceId(created.id);
  return {...created, role:'owner'};
}

async function ensureWorkspaceMemberRowIfOwner(workspace, role='member'){
  if(workspace?.owner_id !== SESSION?.user?.id) return;
  await ensureWorkspaceMemberRow(workspace, 'owner');
}

async function ensureWorkspaceMemberRow(workspace, role='member'){
  const email = sessionEmail();
  if(!workspace?.id || !email) return;
  await supabase.from('workspace_members').upsert({
    workspace_id: workspace.id,
    user_id: SESSION.user.id,
    email,
    name: displayNameFromEmail(email),
    role,
    updated_at: new Date().toISOString()
  }, { onConflict:'workspace_id,email' });
}

async function loadLegacyUserSnapshot(){
  const { data } = await supabase
    .from('app_state')
    .select('data')
    .eq('user_id', SESSION.user.id)
    .maybeSingle();
  return data?.data || null;
}

async function loadLegacyWorkspaceSnapshot(){
  if(!WORKSPACE_ID) return null;
  const { data } = await supabase
    .from('workspace_states')
    .select('data')
    .eq('workspace_id', WORKSPACE_ID)
    .maybeSingle();
  return data?.data || null;
}

async function loadRowWorkspace(workspace){
  const tableResult = await fetchRowTables();
  if(tableResult.error) return { error:tableResult.error };
  if(!hasRowWorkspaceData(tableResult.tables)) return { data:null };
  return {
    data:composeRowWorkspace(tableResult.tables, workspace),
    updatedAt:latestRowTimestamp(tableResult.tables)
  };
}

async function fetchRowTables(){
  const entries = await Promise.all(ROW_SYNC_TABLES.map(async table=>{
    try{
      const { data, error } = await withTimeout(
        rowLoadQuery(table),
        `Loading ${rowLoadLabel(table)} took too long.`,
        OPTIONAL_ROW_SYNC_TABLES.has(table) ? 8000 : 12000
      );
      if(error){
        if(OPTIONAL_ROW_SYNC_TABLES.has(table)){
          if(isMissingRowTableError(error)) ROW_TABLE_AVAILABLE[table] = false;
          return [table, [], null];
        }
        ROW_TABLE_AVAILABLE[table] = false;
        return [table, [], error];
      }
      ROW_TABLE_AVAILABLE[table] = true;
      return [table, data || [], null];
    }catch(error){
      if(OPTIONAL_ROW_SYNC_TABLES.has(table)){
        console.warn(`Skipping optional ${rowLoadLabel(table)} during startup`, error);
        if(isMissingRowTableError(error)) ROW_TABLE_AVAILABLE[table] = false;
        return [table, [], null];
      }
      ROW_TABLE_AVAILABLE[table] = false;
      return [table, [], error instanceof Error ? error : new Error(`Could not read ${table}.`)];
    }
  }));
  const failed = entries.find(([, , error])=>error);
  if(failed) return { error:failed[2].message || `Could not read ${failed[0]}.` };
  return { tables:Object.fromEntries(entries.map(([table, rows])=>[table, rows])) };
}

function rowLoadQuery(table){
  const query = supabase.from(table).select('*').eq('workspace_id', WORKSPACE_ID);
  if(table === 'app_notifications'){
    return query.order('created_at', { ascending:false }).limit(NOTIFICATION_STARTUP_LIMIT);
  }
  return query;
}

function rowLoadLabel(table){
  return table.replace(/^app_/, '').replaceAll('_', ' ');
}

function isMissingRowTableError(error){
  const code = String(error?.code || '');
  const message = String(error?.message || error?.details || '');
  return code === '42P01' || code === 'PGRST205' || /Could not find the table|does not exist|schema cache/i.test(message);
}

function hasRowWorkspaceData(tables){
  return [
    'app_members',
    'app_boards',
    'app_groups',
    'app_columns',
    'app_tasks',
    'app_docs',
    'app_assets',
    'app_updates'
  ].some(table=>(tables[table]||[]).length);
}

function latestRowTimestamp(tables){
  const stamps = Object.values(tables).flat().map(r=>Date.parse(r.updated_at || r.created_at || r.read_at || 0)).filter(Boolean);
  return stamps.length ? new Date(Math.max(...stamps)).toISOString() : new Date().toISOString();
}

function byPosition(a,b){
  const ap = Number.isFinite(Number(a.position)) ? Number(a.position) : 0;
  const bp = Number.isFinite(Number(b.position)) ? Number(b.position) : 0;
  return ap - bp || String(a.name||a.id||'').localeCompare(String(b.name||b.id||''));
}

function composeRowWorkspace(tables, workspace){
  const settingsRow = (tables.app_settings || [])[0] || {};
  const readMap = groupIdsByUpdate(tables.app_update_reads || []);
  const bookmarkMap = groupIdsByUpdate(tables.app_update_bookmarks || []);
  const valueMap = groupTaskValues(tables.app_task_values || []);
  const tasksByParent = groupTasksByParent(tables.app_tasks || []);
  const groupsByBoard = groupRows(tables.app_groups || [], 'board_id');
  const columnsByBoard = groupRows(tables.app_columns || [], 'board_id');

  const data = {
    settings:{
      account:settingsRow.account || workspace?.name || 'TaskTrack',
      currentUser:'',
      currency:settingsRow.currency || '$',
      language:settingsRow.data?.language || 'en',
      memberRoles:settingsRow.data?.memberRoles || {},
      notificationPrefs:settingsRow.data?.notificationPrefs || {},
      assistantActions:settingsRow.data?.assistantActions || {},
      guidedTourDone:settingsRow.data?.guidedTourDone === true,
      onboardingDismissed:settingsRow.data?.onboardingDismissed === true,
      assistantPet:settingsRow.data?.assistantPet || {}
    },
    team:(tables.app_members || []).slice().sort(byPosition).map(m=>({
      id:m.id,
      email:normalizeEmail(m.email),
      name:m.name || displayNameFromEmail(m.email),
      initials:m.initials || initialsFor(m.name || m.email),
      color:m.color || AVATAR_COLORS[0],
      role:(settingsRow.data?.memberRoles || {})[m.id] === 'admin' ? 'admin' : 'member'
    })),
    boards:(tables.app_boards || []).slice().sort(byPosition).map(b=>{
      const columnState = parseBoardColumnState(b.column_order);
      return {
        id:b.id,
        name:b.name || 'Untitled board',
        icon:b.icon || '📋',
        columnOrder:columnState.order,
        columnLabels:columnState.labels,
        hiddenColumns:columnState.hidden,
        sort:columnState.sort,
        personFilter:'',
        statusLabels:columnState.statusLabels,
        priorityLabels:columnState.priorityLabels,
        hiddenStatusLabels:columnState.hiddenStatusLabels,
        hiddenPriorityLabels:columnState.hiddenPriorityLabels,
        columnWidths:columnState.widths,
        calendarShowTimeline:columnState.calendarShowTimeline,
        columns:(columnsByBoard.get(b.id) || []).slice().filter(c=>!isInternalTaskValueColumn(c.id)).sort(byPosition).map(c=>({
          id:c.id,
          name:c.name || 'Untitled',
          type:CUSTOM_COLUMN_TYPES[c.type] ? c.type : 'text'
        })),
        groups:(groupsByBoard.get(b.id) || []).slice().sort(byPosition).map(g=>({
          id:g.id,
          name:g.name || 'Untitled group',
          color:g.color || '#0073ea',
          open:g.is_open !== false,
          tasks:(tasksByParent.get(`root:${g.id}`) || []).slice().sort(byPosition).map(t=>composeTask(t, valueMap, tasksByParent))
        }))
      };
    }),
    docs:(tables.app_docs || []).map(d=>({
      id:d.id,
      title:d.title || 'Untitled doc',
      template:d.template || '',
      content:d.content || '',
      updated:Number(d.updated_number) || Date.parse(d.updated_at) || Date.now()
    })),
    assets:(tables.app_assets || []).map(a=>({
      id:a.id,
      title:a.title || 'Untitled asset',
      kind:a.kind || 'file',
      mimeType:a.mime_type || '',
      size:Number(a.size_bytes) || 0,
      source:a.source || 'upload',
      storageBucket:a.storage_bucket || ASSET_BUCKET,
      storagePath:a.storage_path || '',
      externalUrl:a.external_url || '',
      description:a.description || '',
      createdBy:a.created_by || '',
      data:a.data && typeof a.data === 'object' ? a.data : {},
      updated:Number(a.updated_number) || Date.parse(a.updated_at) || Date.now()
    })),
    automations:(tables.app_automations || []).map(r=>({ ...(r.data || {}), id:r.id })),
    integrations:(tables.app_integrations || []).map(r=>({ ...(r.data || {}), id:r.id })),
    archive:(tables.app_archive_items || []).map(r=>({ ...(r.data || {}), id:r.id })),
    updates:(tables.app_updates || []).map(u=>({
      id:u.id,
      authorId:u.author_id || '',
      boardId:u.board_id || '',
      groupId:u.group_id || '',
      taskId:u.task_id || '',
      body:u.body || '',
      assetIds:Array.isArray(u.asset_ids) ? u.asset_ids : [],
      mentions:Array.isArray(u.mentions) ? u.mentions : [],
      readBy:readMap.get(u.id) || [],
      bookmarkedBy:bookmarkMap.get(u.id) || [],
      scheduledAt:u.scheduled_at ? Date.parse(u.scheduled_at) : null,
      createdAt:Date.parse(u.created_at) || Date.now()
    })).sort((a,b)=>(b.createdAt||0)-(a.createdAt||0)),
    notifications:(tables.app_notifications || []).map(n=>({
      id:n.id,
      recipientId:n.recipient_member_id || '',
      recipientUserId:n.recipient_user_id || '',
      recipientEmail:normalizeEmail(n.recipient_email),
      actorId:n.actor_member_id || '',
      boardId:n.board_id || '',
      groupId:n.group_id || '',
      taskId:n.task_id || '',
      updateId:n.update_id || '',
      type:n.type || 'general',
      title:n.title || '',
      body:n.body || '',
      channels:n.channels && typeof n.channels === 'object' ? n.channels : {},
      data:n.data && typeof n.data === 'object' ? n.data : {},
      dedupeKey:n.dedupe_key || '',
      readAt:n.read_at || '',
      emailedAt:n.emailed_at || '',
      emailStatus:n.email_status || '',
      emailAttempts:Number(n.email_attempts) || 0,
      desktopAt:n.desktop_at || '',
      pushedAt:n.pushed_at || '',
      pushStatus:n.push_status || '',
      pushAttempts:Number(n.push_attempts) || 0,
      createdAt:Date.parse(n.created_at) || Date.now()
    })).sort((a,b)=>(b.createdAt||0)-(a.createdAt||0)),
    notificationPrefs:Object.fromEntries((tables.app_notification_preferences || []).map(row=>[
      row.member_id,
      {...NOTIFICATION_DEFAULTS, ...(row.prefs && typeof row.prefs === 'object' ? row.prefs : {})}
    ])),
    boardNotificationPrefs:(tables.app_board_notification_preferences || []).reduce((out,row)=>{
      if(!out[row.board_id]) out[row.board_id] = {};
      out[row.board_id][row.member_id] = {
        mode:['custom','receive_all','mute'].includes(row.mode) ? row.mode : 'custom',
        prefs:row.prefs && typeof row.prefs === 'object' ? row.prefs : {}
      };
      return out;
    }, {})
  };
  return data;
}

function groupRows(rows, key){
  const out = new Map();
  rows.forEach(row=>{
    const id = row[key] || '';
    if(!out.has(id)) out.set(id, []);
    out.get(id).push(row);
  });
  return out;
}

function parseBoardColumnState(value){
  if(Array.isArray(value)) return { order:value, labels:{}, hidden:[], sort:{}, personFilter:'', statusLabels:{}, priorityLabels:{}, hiddenStatusLabels:[], hiddenPriorityLabels:[], widths:{}, calendarShowTimeline:false };
  if(value && typeof value === 'object'){
    return {
      order:Array.isArray(value.order) ? value.order : BASE_COLUMNS.map(c=>c.id),
      labels:value.labels && typeof value.labels === 'object' ? value.labels : {},
      hidden:Array.isArray(value.hidden) ? value.hidden : [],
      sort:value.sort && typeof value.sort === 'object' ? value.sort : {},
      personFilter:value.personFilter || '',
      statusLabels:value.statusLabels && typeof value.statusLabels === 'object' ? value.statusLabels : {},
      priorityLabels:value.priorityLabels && typeof value.priorityLabels === 'object' ? value.priorityLabels : {},
      hiddenStatusLabels:Array.isArray(value.hiddenStatusLabels) ? value.hiddenStatusLabels : [],
      hiddenPriorityLabels:Array.isArray(value.hiddenPriorityLabels) ? value.hiddenPriorityLabels : [],
      widths:value.widths && typeof value.widths === 'object' ? value.widths : {},
      calendarShowTimeline:value.calendarShowTimeline === true
    };
  }
  return { order:BASE_COLUMNS.map(c=>c.id), labels:{}, hidden:[], sort:{}, personFilter:'', statusLabels:{}, priorityLabels:{}, hiddenStatusLabels:[], hiddenPriorityLabels:[], widths:{}, calendarShowTimeline:false };
}

function serializeBoardColumnState(b, fallbackOrder){
  return {
    order:Array.isArray(b.columnOrder) ? b.columnOrder : fallbackOrder,
    labels:b.columnLabels || {},
    hidden:Array.isArray(b.hiddenColumns) ? b.hiddenColumns : [],
    sort:b.sort || {},
    personFilter:'',
    statusLabels:b.statusLabels || {},
    priorityLabels:b.priorityLabels || {},
    hiddenStatusLabels:Array.isArray(b.hiddenStatusLabels) ? b.hiddenStatusLabels : [],
    hiddenPriorityLabels:Array.isArray(b.hiddenPriorityLabels) ? b.hiddenPriorityLabels : [],
    widths:b.columnWidths || {},
    calendarShowTimeline:b.calendarShowTimeline === true
  };
}

function groupIdsByUpdate(rows){
  const out = new Map();
  rows.forEach(row=>{
    const id = row.update_id || '';
    if(!out.has(id)) out.set(id, []);
    out.get(id).push(row.member_id);
  });
  return out;
}

function groupTaskValues(rows){
  const out = new Map();
  rows.forEach(row=>{
    if(!out.has(row.task_id)) out.set(row.task_id, {});
    out.get(row.task_id)[row.column_id] = row.value ?? '';
  });
  return out;
}

function groupTasksByParent(rows){
  const out = new Map();
  rows.forEach(row=>{
    const key = row.parent_task_id ? `task:${row.parent_task_id}` : `root:${row.group_id}`;
    if(!out.has(key)) out.set(key, []);
    out.get(key).push(row);
  });
  return out;
}

function composeTask(row, valueMap, tasksByParent){
  return {
    id:row.id,
    name:row.name || 'New task',
    owner:row.owner_id || '',
    assignedBy:row.assigned_by || '',
    status:row.status || 'not_started',
    priority:row.priority || 'none',
    due:row.due || '',
    tlStart:row.timeline_start || '',
    tlEnd:row.timeline_end || '',
    budget:Number(row.budget) || 0,
    notes:row.notes || '',
    completedAt:row.completed_at || '',
    hiddenAt:row.archived_at || '',
    custom:valueMap.get(row.id) || {},
    subtasks:(tasksByParent.get(`task:${row.id}`) || []).slice().sort(byPosition).map(st=>composeTask(st, valueMap, tasksByParent))
  };
}

function ensureSessionMember(){
  if(!SESSION?.user || !DATA) return false;
  let changed = false;
  const email = sessionEmail();
  if(!DATA.settings){ DATA.settings = { account:'TaskTrack', currentUser:'u1', currency:'$' }; changed = true; }
  if(DATA.settings.authEmail !== email){ DATA.settings.authEmail = email; }
  if(!DATA.team?.length){ DATA.team = [{id:'u1',name:'You',initials:'YO',color:'#0073ea'}]; changed = true; }
  if(!DATA.settings.memberRoles || typeof DATA.settings.memberRoles !== 'object'){
    DATA.settings.memberRoles = {};
    changed = true;
  }
  DATA.team.forEach(m=>{
    if(m.email){
      const clean = normalizeEmail(m.email);
      if(m.email !== clean){ m.email = clean; changed = true; }
    }
  });

  let me = email ? DATA.team.find(m=>normalizeEmail(m.email) === email) : null;
  const savedMemberId = CLOUD_USER_ID ? Store.get(`tasktrack_member_${CLOUD_USER_ID}`) : '';
  if(!me && savedMemberId) me = DATA.team.find(m=>m.id === savedMemberId);
  if(!me && DATA.settings.currentUser) me = DATA.team.find(m=>m.id === DATA.settings.currentUser);
  if(!me){
    const name = displayNameFromEmail(email);
    me = { id:uid('u'), name, email, initials:initialsFor(name), color:AVATAR_COLORS[DATA.team.length%AVATAR_COLORS.length], role:'member' };
    DATA.team.unshift(me);
    changed = true;
  }
  if(email && !me.email){ me.email = email; changed = true; }
  if((!me.name || me.name === 'You') && email){ me.name = displayNameFromEmail(email); changed = true; }
  const initials = initialsFor(me.name);
  if(me.initials !== initials){ me.initials = initials; changed = true; }
  setCurrentMemberId(me.id);
  if(!DATA.settings.currentUser){
    DATA.settings.currentUser = me.id;
  }
  const hasAdmin = DATA.team.some(m=>memberRole(m.id) === 'admin');
  DATA.team.forEach((m,i)=>{
    const nextRole = hasAdmin ? memberRole(m.id) : (i === 0 ? 'admin' : 'member');
    if(m.role !== nextRole){ m.role = nextRole; changed = true; }
    if(DATA.settings.memberRoles[m.id] !== nextRole){ DATA.settings.memberRoles[m.id] = nextRole; changed = true; }
  });
  return changed;
}

function dataToRowSets(data){
  const rows = Object.fromEntries(ROW_SYNC_TABLES.map(table=>[table, []]));
  if(!WORKSPACE_ID || !data) return rows;
  const firstBoardId = (data.boards || []).find(b=>b?.id)?.id || '';

  rows.app_settings.push({
    workspace_id:WORKSPACE_ID,
    account:data.settings?.account || 'TaskTrack',
    currency:data.settings?.currency || '$',
    data:{
      language:LANGUAGE_OPTIONS[data.settings?.language] ? data.settings.language : 'en',
      memberRoles:data.settings?.memberRoles || {},
      notificationPrefs:{...NOTIFICATION_DEFAULTS, ...(data.settings?.notificationPrefs || {})},
      assistantActions:{...ASSISTANT_ACTION_DEFAULTS, ...(data.settings?.assistantActions || {})},
      guidedTourDone:data.settings?.guidedTourDone === true,
      onboardingDismissed:data.settings?.onboardingDismissed === true,
      assistantPet: assistantPetSettings(data.settings?.assistantPet)
    }
  });

  (data.team || []).forEach((m, position)=>{
    rows.app_members.push({
      workspace_id:WORKSPACE_ID,
      id:m.id,
      email:normalizeEmail(m.email),
      name:m.name || displayNameFromEmail(m.email),
      initials:m.initials || initialsFor(m.name || m.email),
      color:m.color || AVATAR_COLORS[position % AVATAR_COLORS.length],
      position
    });
  });

  (data.boards || []).forEach((b, boardPosition)=>{
    const builtInIds = BASE_COLUMNS.map(c=>c.id);
    const allowedTaskValueColumnIds = taskValueColumnIdsForBoard(b);
    rows.app_boards.push({
      workspace_id:WORKSPACE_ID,
      id:b.id,
      name:b.name || 'Untitled board',
      icon:b.icon || '',
      column_order:serializeBoardColumnState(b, builtInIds),
      position:boardPosition
    });
    (b.groups || []).forEach((g, groupPosition)=>{
      rows.app_groups.push({
        workspace_id:WORKSPACE_ID,
        id:g.id,
        board_id:b.id,
        name:g.name || 'Untitled group',
        color:g.color || '#0073ea',
        position:groupPosition,
        is_open:g.open !== false
      });
      (g.tasks || []).forEach((t, taskPosition)=>{
        pushTaskRows(rows, t, g.id, null, taskPosition, allowedTaskValueColumnIds);
      });
    });
    (b.columns || []).forEach((c, columnPosition)=>{
      rows.app_columns.push({
        workspace_id:WORKSPACE_ID,
        id:c.id,
        board_id:b.id,
        name:c.name || 'Untitled',
        type:CUSTOM_COLUMN_TYPES[c.type] ? c.type : 'text',
        position:columnPosition
      });
    });
  });
  if(firstBoardId){
    INTERNAL_TASK_VALUE_COLUMNS.forEach((c, position)=>{
      rows.app_columns.push({
        workspace_id:WORKSPACE_ID,
        id:c.id,
        board_id:firstBoardId,
        name:c.name,
        type:'text',
        position:900000 + position
      });
    });
  }

  (data.docs || []).forEach(d=>{
    rows.app_docs.push({
      workspace_id:WORKSPACE_ID,
      id:d.id,
      title:d.title || 'Untitled doc',
      template:d.template || '',
      content:d.content || '',
      updated_number:Number(d.updated) || null
    });
  });
  (data.assets || []).forEach(a=>{
    rows.app_assets.push({
      workspace_id:WORKSPACE_ID,
      id:a.id,
      title:a.title || 'Untitled asset',
      kind:a.kind || 'file',
      mime_type:a.mimeType || '',
      size_bytes:Number(a.size) || 0,
      source:a.source || 'upload',
      storage_bucket:a.storageBucket || ASSET_BUCKET,
      storage_path:a.storagePath || '',
      external_url:a.externalUrl || '',
      description:a.description || '',
      created_by:a.createdBy || '',
      updated_number:Number(a.updated) || null,
      data:a.data && typeof a.data === 'object' ? a.data : {}
    });
  });
  (data.automations || []).forEach(a=>{
    rows.app_automations.push({ workspace_id:WORKSPACE_ID, id:a.id, data:{...a} });
  });
  (data.integrations || []).forEach(i=>{
    rows.app_integrations.push({ workspace_id:WORKSPACE_ID, id:i.id, data:{...i} });
  });
  (data.archive || []).forEach(a=>{
    rows.app_archive_items.push({ workspace_id:WORKSPACE_ID, id:a.id, data:{...a} });
  });
  (data.updates || []).forEach(u=>{
    const created = isoFromTimestamp(u.createdAt) || new Date().toISOString();
    rows.app_updates.push({
      workspace_id:WORKSPACE_ID,
      id:u.id,
      author_id:u.authorId || '',
      board_id:u.boardId || '',
      group_id:u.groupId || '',
      task_id:u.taskId || '',
      body:u.body || '',
      asset_ids:Array.isArray(u.assetIds) ? u.assetIds : [],
      mentions:Array.isArray(u.mentions) ? u.mentions : [],
      scheduled_at:isoFromTimestamp(u.scheduledAt),
      created_at:created
    });
    (u.readBy || []).forEach(memberId=>{
      rows.app_update_reads.push({ workspace_id:WORKSPACE_ID, update_id:u.id, member_id:memberId });
    });
    (u.bookmarkedBy || []).forEach(memberId=>{
      rows.app_update_bookmarks.push({ workspace_id:WORKSPACE_ID, update_id:u.id, member_id:memberId });
    });
  });
  (data.notifications || []).forEach(n=>rows.app_notifications.push(notificationToRow(n)));
  Object.entries(data.notificationPrefs || {}).forEach(([memberId,prefs])=>{
    if(!memberId) return;
    const m = (data.team || []).find(x=>x.id === memberId) || {};
    rows.app_notification_preferences.push({
      workspace_id:WORKSPACE_ID,
      member_id:memberId,
      user_id:memberId === currentMemberId() ? (SESSION?.user?.id || null) : null,
      email:normalizeEmail(m.email),
      prefs:{...NOTIFICATION_DEFAULTS, ...(prefs || {})}
    });
  });
  Object.entries(data.boardNotificationPrefs || {}).forEach(([boardId,memberMap])=>{
    Object.entries(memberMap || {}).forEach(([memberId,value])=>{
      rows.app_board_notification_preferences.push({
        workspace_id:WORKSPACE_ID,
        board_id:boardId,
        member_id:memberId,
        mode:['custom','receive_all','mute'].includes(value?.mode) ? value.mode : 'custom',
        prefs:value?.prefs && typeof value.prefs === 'object' ? value.prefs : {}
      });
    });
  });
  return rows;
}

function notificationToRow(n){
  return {
    workspace_id:WORKSPACE_ID,
    id:n.id,
    recipient_member_id:n.recipientId || '',
    recipient_user_id:n.recipientUserId || null,
    recipient_email:normalizeEmail(n.recipientEmail),
    actor_member_id:n.actorId || '',
    board_id:n.boardId || '',
    group_id:n.groupId || '',
    task_id:n.taskId || '',
    update_id:n.updateId || '',
    type:n.type || 'general',
    title:n.title || '',
    body:n.body || '',
    channels:n.channels && typeof n.channels === 'object' ? n.channels : {},
    data:n.data && typeof n.data === 'object' ? n.data : {},
    dedupe_key:n.dedupeKey || '',
    read_at:n.readAt || null,
    emailed_at:n.emailedAt || null,
    email_status:n.emailStatus || 'pending',
    email_attempts:Number(n.emailAttempts) || 0,
    desktop_at:n.desktopAt || null,
    pushed_at:n.pushedAt || null,
    push_status:n.pushStatus || 'pending',
    push_attempts:Number(n.pushAttempts) || 0,
    created_at:isoFromTimestamp(n.createdAt) || new Date().toISOString()
  };
}

function pushTaskRows(rows, task, groupId, parentTaskId, position, allowedTaskValueColumnIds=new Set()){
  rows.app_tasks.push({
    workspace_id:WORKSPACE_ID,
    id:task.id,
    group_id:groupId,
    parent_task_id:parentTaskId,
    name:task.name || 'New task',
    owner_id:task.owner || '',
    assigned_by:task.assignedBy || '',
    status:task.status || 'not_started',
    priority:task.priority || 'none',
    due:task.due || null,
    timeline_start:task.tlStart || null,
    timeline_end:task.tlEnd || null,
    budget:Number(task.budget) || 0,
    notes:task.notes || '',
    position,
    completed_at:task.completedAt || null,
    archived_at:task.hiddenAt || null
  });
  Object.entries(task.custom || {}).forEach(([columnId, value])=>{
    if(!allowedTaskValueColumnIds.has(columnId)) return;
    rows.app_task_values.push({
      workspace_id:WORKSPACE_ID,
      task_id:task.id,
      column_id:columnId,
      value:value ?? ''
    });
  });
  (task.subtasks || []).forEach((subtask, subPosition)=>{
    pushTaskRows(rows, subtask, groupId, task.id, subPosition, allowedTaskValueColumnIds);
  });
}

function isoFromTimestamp(value){
  if(!value) return null;
  if(typeof value === 'string' && /\d{4}-\d{2}-\d{2}T/.test(value)) return value;
  const n = Number(value);
  if(!Number.isFinite(n)) return null;
  return new Date(n).toISOString();
}

function rowKey(table, row){
  if(table === 'app_settings') return row.workspace_id;
  if(table === 'app_task_values') return `${row.workspace_id}|${row.task_id}|${row.column_id}`;
  if(table === 'app_update_reads' || table === 'app_update_bookmarks') return `${row.workspace_id}|${row.update_id}|${row.member_id}`;
  if(table === 'app_notification_preferences') return `${row.workspace_id}|${row.member_id}`;
  if(table === 'app_board_notification_preferences') return `${row.workspace_id}|${row.board_id}|${row.member_id}`;
  return `${row.workspace_id}|${row.id}`;
}

function stableCopy(value){
  if(Array.isArray(value)) return value.map(stableCopy);
  if(value && typeof value === 'object'){
    return Object.keys(value).sort().reduce((out, key)=>{ out[key] = stableCopy(value[key]); return out; }, {});
  }
  return value;
}

function rowSignature(row){
  return JSON.stringify(stableCopy(row));
}

function keyedRows(rows, table){
  return new Map((rows || []).map(row=>[rowKey(table, row), row]));
}

function rowsWithTimestamp(table, rows){
  if(table === 'app_update_reads' || table === 'app_update_bookmarks') return rows;
  const now = new Date().toISOString();
  return rows.map(row=>({ ...row, updated_at:now }));
}

async function saveRowWorkspace({ full=false } = {}){
  if(full && !isCurrentAdmin()) throw new Error('Only workspace admins can replace the full workspace.');
  const nextRows = dataToRowSets(DATA);
  const prevRows = full || !LAST_SYNC_DATA ? Object.fromEntries(ROW_SYNC_TABLES.map(table=>[table, []])) : dataToRowSets(LAST_SYNC_DATA);

  if(full){
    for(const table of ROW_DELETE_ORDER){
      if(ROW_TABLE_AVAILABLE[table] === false) continue;
      const { error } = await supabase.from(table).delete().eq('workspace_id', WORKSPACE_ID);
      if(error) throw error;
    }
  } else {
    for(const table of ROW_DELETE_ORDER){
      if(ROW_TABLE_AVAILABLE[table] === false) continue;
      const previous = keyedRows(prevRows[table], table);
      const next = keyedRows(nextRows[table], table);
      for(const [key, row] of previous){
        if(next.has(key)) continue;
        const { error } = await deleteRow(table, row);
        if(error) throw error;
      }
    }
  }

  for(const table of ROW_UPSERT_ORDER){
    if(ROW_TABLE_AVAILABLE[table] === false) continue;
    if(PRIVATE_DELIVERY_TABLES.has(table)) continue;
    if(table === 'app_settings' && !isCurrentAdmin()) continue;
    const previous = keyedRows(prevRows[table], table);
    const changed = (nextRows[table] || []).filter(row=>{
      if(table === 'app_columns') return true;
      const old = previous.get(rowKey(table, row));
      return !old || rowSignature(old) !== rowSignature(row);
    });
    if(!changed.length) continue;
    const { error } = await supabase
      .from(table)
      .upsert(rowsWithTimestamp(table, changed), { onConflict:ROW_CONFLICTS[table] || 'workspace_id,id' });
    if(error) throw error;
  }
  LAST_SYNC_DATA = JSON.parse(JSON.stringify(DATA));
}

async function persistNotificationInsert(notification){
  if(!notification || !supabase || !SESSION?.user || !WORKSPACE_ID || ROW_TABLE_AVAILABLE.app_notifications === false) return;
  try{
    const [row] = rowsWithTimestamp('app_notifications', [notificationToRow(notification)]);
    const { error } = await supabase.from('app_notifications').insert(row);
    if(error){
      if(isMissingRowTableError(error)){ ROW_TABLE_AVAILABLE.app_notifications = false; return; }
      if(isDuplicateRowError(error)) return;
      console.warn('Notification delivery failed', error);
    }
  }catch(error){
    console.warn('Notification delivery failed', error);
  }
}

async function persistNotificationPatch(ids, patch){
  const cleanIds = [...new Set((ids || []).filter(Boolean))];
  if(!cleanIds.length || !supabase || !SESSION?.user || !WORKSPACE_ID || ROW_TABLE_AVAILABLE.app_notifications === false) return;
  try{
    const { error } = await supabase
      .from('app_notifications')
      .update({ ...patch, updated_at:new Date().toISOString() })
      .eq('workspace_id', WORKSPACE_ID)
      .in('id', cleanIds);
    if(error){
      if(isMissingRowTableError(error)){ ROW_TABLE_AVAILABLE.app_notifications = false; return; }
      console.warn('Notification status update failed', error);
    }
  }catch(error){
    console.warn('Notification status update failed', error);
  }
}

function isDuplicateRowError(error){
  const code = String(error?.code || '');
  const message = String(error?.message || error?.details || '');
  return code === '23505' || /duplicate key|already exists/i.test(message);
}

function deleteRow(table, row){
  let query = supabase.from(table).delete().eq('workspace_id', row.workspace_id);
  if(table === 'app_settings') return query;
  if(table === 'app_task_values') return query.eq('task_id', row.task_id).eq('column_id', row.column_id);
  if(table === 'app_update_reads' || table === 'app_update_bookmarks') return query.eq('update_id', row.update_id).eq('member_id', row.member_id);
  if(table === 'app_notification_preferences') return query.eq('member_id', row.member_id);
  if(table === 'app_board_notification_preferences') return query.eq('board_id', row.board_id).eq('member_id', row.member_id);
  return query.eq('id', row.id);
}

function queueCloudSave(){
  if(APPLYING_REMOTE || CLOUD_LOADING || !supabase || !SESSION?.user || !WORKSPACE_ID || !DATA) return;
  CLOUD_STATUS = 'Saving';
  updateCloudStatus();
  clearTimeout(CLOUD_SAVE_TIMER);
  CLOUD_SAVE_TIMER = setTimeout(()=>{ CLOUD_SAVE_TIMER = null; void saveCloudNow(); }, 700);
}

async function saveCloudNow({ silent=false, full=false } = {}){
  if(APPLYING_REMOTE || !supabase || !SESSION?.user || !WORKSPACE_ID || !DATA) return;
  clearTimeout(CLOUD_SAVE_TIMER);
  CLOUD_SAVE_TIMER = null;
  if(!silent){
    CLOUD_STATUS = 'Saving';
    updateCloudStatus();
  }
  try{
    await saveRowWorkspace({ full });
    WORKSPACE_UPDATED_AT = new Date().toISOString();
  }catch(error){
    CLOUD_STATUS = 'Save failed';
    updateCloudStatus();
    if(!silent) toast(error.message || 'Supabase save failed');
    return;
  }
  CLOUD_STATUS = 'Synced';
  updateCloudStatus();
}

function updateCloudStatus(){
  const el = $('#cloud-status');
  if(!el) return;
  el.textContent = CLOUD_STATUS;
  el.classList.toggle('error', CLOUD_STATUS === 'Save failed');
  el.classList.toggle('saving', CLOUD_STATUS === 'Saving' || CLOUD_STATUS === 'Loading');
  el.classList.toggle('queued', CLOUD_STATUS === 'Changes waiting' || CLOUD_STATUS === 'Review changes');
}

async function signOut(){
  await saveCloudNow({ silent:true });
  stopWorkspaceSync();
  if(supabase) await supabase.auth.signOut();
}

function startWorkspaceSync(){
  if(!supabase || !WORKSPACE_ID) return;
  stopWorkspaceSync();
  let channel = supabase.channel(`workspace-rows-${WORKSPACE_ID}`);
  ROW_SYNC_TABLES.forEach(table=>{
    channel = channel.on('postgres_changes', {
      event:'*',
      schema:'public',
      table,
      filter:`workspace_id=eq.${WORKSPACE_ID}`
    }, payload=>handleRealtimeRowChange(table, payload));
  });
  channel = channel.on('postgres_changes', {
    event:'*', schema:'public', table:'app_shared_calendar_events',
    filter:`workspace_id=eq.${WORKSPACE_ID}`
  }, ()=>{ void loadWorkspaceSharedCalendar({force:true}); });
  WORKSPACE_CHANNEL = channel.subscribe();
  startWorkspacePresence();
  WORKSPACE_POLL_TIMER = setInterval(()=>{
    if(!isEditingControl()) void fetchWorkspaceState();
    void loadWorkspaceSharedCalendar({force:true});
  }, 30000);
}

function stopWorkspaceSync(){
  if(WORKSPACE_CHANNEL && supabase){
    supabase.removeChannel(WORKSPACE_CHANNEL);
  }
  WORKSPACE_CHANNEL = null;
  clearTimeout(ROW_REFRESH_TIMER);
  ROW_REFRESH_TIMER = null;
  clearInterval(WORKSPACE_POLL_TIMER);
  WORKSPACE_POLL_TIMER = null;
  stopWorkspacePresence();
}

function handleRealtimeRowChange(table, payload){
  const row = payload?.eventType === 'DELETE' ? payload.old : payload.new;
  if(!row || row.workspace_id !== WORKSPACE_ID || APPLYING_REMOTE) return;
  const review = remotePatchReviewFor(table, payload);
  if(review?.conflict){
    queueRemoteRowPatch(table, payload, review);
    return;
  }
  const applied = applyRealtimeRowPatch(table, payload);
  if(applied){
    finishRealtimeRowPatch(row);
    if(table === 'app_notifications' && payload?.eventType !== 'DELETE'){
      showDesktopNotificationForId(row.id);
    }
  }
}

function finishRealtimeRowPatch(row){
  const stamp = row?.updated_at || row?.created_at || row?.read_at || new Date().toISOString();
  WORKSPACE_UPDATED_AT = stamp;
  LAST_SYNC_DATA = JSON.parse(JSON.stringify(DATA));
  CLOUD_STATUS = PENDING_REMOTE_PATCHES.length || PENDING_REMOTE_WORKSPACE ? CLOUD_STATUS : 'Synced';
  updateCloudStatus();
  saveLocal();
  if(isEditingControl()) scheduleDeferredRemoteRender();
  else render({preserveScroll:true});
}

function scheduleDeferredRemoteRender(){
  REMOTE_RENDER_PENDING = true;
  clearTimeout(REMOTE_RENDER_TIMER);
  REMOTE_RENDER_TIMER = setTimeout(()=>{
    if(isEditingControl()){
      scheduleDeferredRemoteRender();
      return;
    }
    REMOTE_RENDER_PENDING = false;
    render({preserveScroll:true});
  }, 800);
}

function startWorkspacePresence(){
  if(!supabase || !WORKSPACE_ID || PRESENCE_CHANNEL) return;
  PRESENCE_CHANNEL = supabase.channel(`workspace-presence-${WORKSPACE_ID}`, {
    config:{ presence:{ key:LOCAL_CLIENT_ID } }
  });
  PRESENCE_CHANNEL
    .on('presence', {event:'sync'}, ()=>syncPresenceState())
    .on('presence', {event:'join'}, ()=>syncPresenceState())
    .on('presence', {event:'leave'}, ()=>syncPresenceState())
    .subscribe(status=>{
      if(status === 'SUBSCRIBED') trackCurrentPresence();
    });
  clearInterval(PRESENCE_TRACK_TIMER);
  PRESENCE_TRACK_TIMER = setInterval(trackCurrentPresence, 5000);
}

function stopWorkspacePresence(){
  clearInterval(PRESENCE_TRACK_TIMER);
  PRESENCE_TRACK_TIMER = null;
  if(PRESENCE_CHANNEL && supabase) supabase.removeChannel(PRESENCE_CHANNEL);
  PRESENCE_CHANNEL = null;
  REMOTE_PRESENCE = [];
}

function schedulePresenceTrack(){
  clearTimeout(PRESENCE_TRACK_DEBOUNCE);
  PRESENCE_TRACK_DEBOUNCE = setTimeout(trackCurrentPresence, 120);
}

function trackCurrentPresence(){
  if(!PRESENCE_CHANNEL || !SESSION?.user) return;
  const me = member(currentMemberId()) || {};
  const target = currentEditingTarget() || currentViewingTarget();
  void PRESENCE_CHANNEL.track({
    userId:SESSION.user.id,
    memberId:currentMemberId(),
    name:me.name || displayNameFromEmail(SESSION.user.email),
    initials:me.initials || initialsFor(me.name || SESSION.user.email),
    color:me.color || '#0073ea',
    target,
    at:Date.now()
  });
}

function syncPresenceState(){
  if(!PRESENCE_CHANNEL) return;
  const state = PRESENCE_CHANNEL.presenceState?.() || {};
  REMOTE_PRESENCE = Object.entries(state).flatMap(([key, entries])=>
    (entries || []).filter(Boolean).map(meta=>({...meta, clientId:key}))
  ).filter(meta=>meta.clientId !== LOCAL_CLIENT_ID && meta.userId);
  if(!isEditingControl()) renderPresenceOnly();
}

function renderPresenceOnly(){
  document.querySelectorAll('[data-presence-target]').forEach(el=>{
    const [type,id] = el.dataset.presenceTarget.split('|');
    el.innerHTML = presenceBadgesHtml(type,id);
  });
}

function scheduleRemoteRefresh(){
  if(CLOUD_LOADING || CLOUD_STATUS === 'Saving' || CLOUD_SAVE_TIMER) return;
  clearTimeout(ROW_REFRESH_TIMER);
  ROW_REFRESH_TIMER = setTimeout(()=>{ void fetchWorkspaceState(); }, isTextEditing() ? 2000 : 250);
}

async function fetchWorkspaceState(){
  if(!supabase || !SESSION?.user || !WORKSPACE_ID || CLOUD_STATUS === 'Saving') return;
  const rowResult = await loadRowWorkspace({ id:WORKSPACE_ID, owner_id:WORKSPACE_OWNER_ID, name:DATA?.settings?.account || 'TaskTrack' });
  if(!rowResult.error && rowResult.data) applyRemoteWorkspace(rowResult);
}

function applyRemoteWorkspace(rowResult){
  if(!rowResult?.data?.boards || CLOUD_STATUS === 'Saving' || CLOUD_SAVE_TIMER) return;
  if(rowResult.updatedAt && rowResult.updatedAt === WORKSPACE_UPDATED_AT) return;
  if(isTextEditing()){
    queueRemoteWorkspace(rowResult);
    return;
  }
  commitRemoteWorkspace(rowResult);
}

function commitRemoteWorkspace(rowResult){
  if(!rowResult?.data?.boards) return;
  const previousData = DATA ? JSON.parse(JSON.stringify(DATA)) : null;
  APPLYING_REMOTE = true;
  DATA = rowResult.data;
  LAST_SYNC_DATA = JSON.parse(JSON.stringify(DATA));
  WORKSPACE_UPDATED_AT = rowResult.updatedAt || new Date().toISOString();
  const sessionMemberChanged = ensureSessionMember();
  normalizeData();
  const recurringChanged = applyRecurringTaskRollovers(DATA);
  saveLocal();
  CLOUD_STATUS = 'Synced';
  APPLYING_REMOTE = false;
  notifyNewRemoteNotifications(DATA, previousData);
  notifyNewRemoteUpdates(DATA, previousData);
  render();
  if(sessionMemberChanged || recurringChanged) save();
}

function findTaskInData(data, taskId){
  const search = (tasks=[], parent=null)=>{
    for(const t of tasks){
      if(t.id === taskId) return {t,parent};
      const found = search(t.subtasks || [], t);
      if(found) return found;
    }
    return null;
  };
  for(const b of data?.boards || []){
    for(const g of b.groups || []){
      const found = search(g.tasks || []);
      if(found) return {...found,b,g};
    }
  }
  return null;
}

function docInData(data, id){
  return (data?.docs || []).find(d=>d.id === id) || null;
}

function taskConflictSource(t){
  if(!t) return null;
  return {
    id:t.id,
    name:t.name || '',
    owner:t.owner || '',
    owners:parseIdList(t?.custom?.[TASK_OWNERS_KEY] ?? t?.owners ?? []),
    assignedBy:t.assignedBy || '',
    status:t.status || '',
    priority:t.priority || '',
    due:t.due || '',
    tlStart:t.tlStart || '',
    tlEnd:t.tlEnd || '',
    budget:Number(t.budget) || 0,
    notes:t.notes || '',
    custom:stableCopy(t.custom || {}),
    subtasks:(t.subtasks || []).map(taskConflictSource)
  };
}

function taskConflictFingerprint(t){
  return JSON.stringify(stableCopy(taskConflictSource(t)));
}

function docConflictFingerprint(d){
  return JSON.stringify(stableCopy({
    id:d?.id || '',
    title:d?.title || '',
    template:d?.template || '',
    content:d?.content || ''
  }));
}

function remoteReviewFor(rowResult){
  const base = {
    conflict:false,
    title:'Team changes are waiting',
    body:'Finish your current edit, then apply the latest workspace changes.'
  };
  const scope = currentEditingScope();
  if(!scope || !rowResult?.data) return base;
  if(scope.type === 'task'){
    const local = findTaskInData(DATA, scope.id);
    const incoming = findTaskInData(rowResult.data, scope.id);
    if(taskConflictFingerprint(local?.t) !== taskConflictFingerprint(incoming?.t)){
      const name = incoming?.t?.name || local?.t?.name || 'this task';
      return {
        conflict:true,
        scope,
        title:'This task changed in another session',
        body:`A teammate changed "${name}" while you were editing. Finish your edit, then apply latest to review the team version.`
      };
    }
  }
  if(scope.type === 'doc'){
    const local = docInData(DATA, scope.id);
    const incoming = docInData(rowResult.data, scope.id);
    if(docConflictFingerprint(local) !== docConflictFingerprint(incoming)){
      const title = incoming?.title || local?.title || 'this document';
      return {
        conflict:true,
        scope,
        title:'This document changed in another session',
        body:`A teammate changed "${title}" while you were editing. Save or finish your edit, then apply latest to review the team version.`
      };
    }
  }
  return base;
}

function queueRemoteWorkspace(rowResult){
  if(PENDING_REMOTE_WORKSPACE?.updatedAt && PENDING_REMOTE_WORKSPACE.updatedAt === rowResult.updatedAt) return;
  PENDING_REMOTE_WORKSPACE = rowResult;
  PENDING_REMOTE_REVIEW = remoteReviewFor(rowResult);
  PENDING_REMOTE_BANNER_HIDDEN = false;
  CLOUD_STATUS = PENDING_REMOTE_REVIEW?.conflict ? 'Review changes' : 'Changes waiting';
  updateCloudStatus();
  renderQueuedRemoteBanner();
}

function queueRemoteRowPatch(table, payload, review){
  const key = remotePatchKey(table, payload);
  PENDING_REMOTE_PATCHES = PENDING_REMOTE_PATCHES.filter(item=>remotePatchKey(item.table, item.payload) !== key);
  PENDING_REMOTE_PATCHES.push({table,payload,review});
  PENDING_REMOTE_REVIEW = review || {
    conflict:true,
    title:'A teammate changed the same field',
    body:'Finish your current edit, then apply latest to review the team version.'
  };
  PENDING_REMOTE_BANNER_HIDDEN = false;
  CLOUD_STATUS = 'Review changes';
  updateCloudStatus();
  renderQueuedRemoteBanner();
}

function remotePatchKey(table, payload){
  const row = payload?.eventType === 'DELETE' ? payload.old : payload.new;
  if(table === 'app_task_values') return `${table}|${row?.task_id || ''}|${row?.column_id || ''}`;
  if(table === 'app_update_reads' || table === 'app_update_bookmarks') return `${table}|${row?.update_id || ''}|${row?.member_id || ''}`;
  return `${table}|${row?.id || row?.workspace_id || ''}`;
}

function clearQueuedRemoteWorkspace(){
  PENDING_REMOTE_WORKSPACE = null;
  PENDING_REMOTE_PATCHES = [];
  PENDING_REMOTE_REVIEW = null;
  PENDING_REMOTE_BANNER_HIDDEN = false;
  removeQueuedRemoteBanner();
}

function finishQueuedRemoteApply(pending, patches=[]){
  clearQueuedRemoteWorkspace();
  if(pending) commitRemoteWorkspace(pending);
  else if(patches.length){
    patches.forEach(item=>applyRealtimeRowPatch(item.table, item.payload, {force:true}));
    LAST_SYNC_DATA = JSON.parse(JSON.stringify(DATA));
    CLOUD_STATUS = 'Synced';
    updateCloudStatus();
    saveLocal();
    render({preserveScroll:true});
  }
}

function applyQueuedRemoteWorkspace(){
  const pending = PENDING_REMOTE_WORKSPACE;
  const patches = PENDING_REMOTE_PATCHES.slice();
  if(!pending && !patches.length) return;
  if(PENDING_REMOTE_REVIEW?.conflict && $('#modal-root [data-doc-editor-id]')){
    toast('Save or close the document before applying team changes');
    return;
  }
  const wasEditing = isEditingControl();
  const active = document.activeElement;
  if(active && active !== document.body && active.blur) active.blur();
  if(wasEditing) setTimeout(()=>finishQueuedRemoteApply(pending, patches), 0);
  else finishQueuedRemoteApply(pending, patches);
}

function hideQueuedRemoteBanner(){
  PENDING_REMOTE_BANNER_HIDDEN = true;
  renderQueuedRemoteBanner();
}

function removeQueuedRemoteBanner(){
  $('#remote-review-banner')?.remove();
}

function renderQueuedRemoteBanner(){
  if((!PENDING_REMOTE_WORKSPACE && !PENDING_REMOTE_PATCHES.length) || PENDING_REMOTE_BANNER_HIDDEN){
    removeQueuedRemoteBanner();
    return;
  }
  const review = PENDING_REMOTE_REVIEW || {};
  let el = $('#remote-review-banner');
  if(!el){
    el = document.createElement('div');
    el.id = 'remote-review-banner';
    document.body.appendChild(el);
  }
  el.className = `remote-review-banner ${review.conflict ? 'conflict' : ''}`;
  el.innerHTML = `
    <div class="remote-review-copy">
      <strong>${esc(review.title || 'Team changes are waiting')}</strong>
      <span>${esc(review.body || 'Apply the latest workspace changes when you are ready.')}</span>
    </div>
    <div class="remote-review-actions">
      <button class="btn ghost sm" data-remote-hide>Later</button>
      <button class="btn sm" data-remote-apply>Apply latest</button>
    </div>`;
  el.querySelector('[data-remote-hide]')?.addEventListener('mousedown', e=>{
    e.preventDefault();
    hideQueuedRemoteBanner();
  });
  el.querySelector('[data-remote-apply]')?.addEventListener('mousedown', e=>{
    e.preventDefault();
    applyQueuedRemoteWorkspace();
  });
}

function remotePatchReviewFor(table, payload){
  const target = currentEditingTarget();
  if(!target || !isEditingControl()) return {conflict:false};
  const row = payload?.eventType === 'DELETE' ? payload.old : payload.new;
  if(!row) return {conflict:false};
  if(target.type === 'task'){
    if(table === 'app_tasks' && row.id === target.id){
      const loc = findTaskInData(DATA, target.id);
      const changed = payload.eventType === 'DELETE' ? ['task'] : changedTaskFieldsFromRow(row, loc?.t);
      if(fieldsOverlap(target.field, changed)){
        return taskFieldConflictReview(loc?.t, row, target.field, payload.eventType === 'DELETE');
      }
    }
    if(table === 'app_task_values' && row.task_id === target.id){
      const field = taskValueField(row.column_id);
      if(fieldsOverlap(target.field, [field])){
        const loc = findTaskInData(DATA, target.id);
        return taskFieldConflictReview(loc?.t, row, target.field, payload.eventType === 'DELETE');
      }
    }
  }
  if(target.type === 'doc' && table === 'app_docs' && row.id === target.id){
    const local = docInData(DATA, target.id);
    const changed = payload.eventType === 'DELETE' ? ['doc'] : changedDocFieldsFromRow(row, local);
    if(fieldsOverlap(target.field, changed)){
      return {
        conflict:true,
        scope:{type:'doc', id:target.id},
        title:'A teammate changed the same document field',
        body:`${fieldLabel(target.field)} changed on "${row.title || local?.title || 'this document'}" while you were editing. Save or close your edit, then apply latest to review it.`
      };
    }
  }
  return {conflict:false};
}

function taskFieldConflictReview(localTask, row, field, deleted=false){
  const name = row?.name || localTask?.name || 'this task';
  return {
    conflict:true,
    scope:{type:'task', id:localTask?.id || row?.id || row?.task_id || ''},
    title:deleted ? 'This task changed in another session' : 'A teammate changed the same task field',
    body:deleted
      ? `A teammate changed or removed "${name}" while you were editing it. Finish your edit, then apply latest to review the team version.`
      : `${fieldLabel(field)} changed on "${name}" while you were editing that field. Finish your edit, then apply latest to review it.`
  };
}

function fieldLabel(field=''){
  if(field === 'name') return 'Task name';
  if(field === 'owner') return 'Owners';
  if(field === 'status') return 'Status';
  if(field === 'priority') return 'Priority';
  if(field === 'due') return 'Due Date';
  if(field === 'timeline') return 'Timeline';
  if(field === 'budget') return 'Budget';
  if(field === 'notes') return 'Notes';
  if(field === 'hidden') return 'Visibility';
  if(field === 'title') return 'Title';
  if(field === 'content') return 'Content';
  if(field.startsWith('custom:')) return 'This custom column';
  return 'This field';
}

function taskValueField(columnId){
  if(columnId === TASK_OWNERS_KEY) return 'owner';
  if(columnId === TASK_DESCRIPTION_KEY) return `custom:${TASK_DESCRIPTION_KEY}`;
  return `custom:${columnId}`;
}

function fieldsOverlap(activeField, changedFields=[]){
  if(!activeField || activeField === 'detail') return false;
  const aliases = fieldAliases(activeField);
  return changedFields.some(field=>aliases.has(field) || fieldAliases(field).has(activeField));
}

function fieldAliases(field=''){
  const set = new Set([field]);
  if(field === 'timeline' || field === 'tlStart' || field === 'tlEnd'){
    set.add('timeline'); set.add('tlStart'); set.add('tlEnd'); set.add('due');
  }
  if(field === 'due'){
    set.add('timeline'); set.add('tlStart'); set.add('tlEnd'); set.add('due');
  }
  if(field === 'owner'){
    set.add('owner'); set.add(`custom:${TASK_OWNERS_KEY}`);
  }
  if(field === 'doc') set.add('title').add('content');
  if(field === 'task') set.add('name').add('owner').add('status').add('priority').add('due').add('timeline').add('budget').add('notes');
  return set;
}

function changedTaskFieldsFromRow(row, t){
  if(!t) return ['task'];
  const checks = [
    ['name', row.name || '', t.name || ''],
    ['owner', row.owner_id || '', t.owner || ''],
    ['status', row.status || '', t.status || ''],
    ['priority', row.priority || '', t.priority || ''],
    ['due', row.due || '', t.due || ''],
    ['tlStart', row.timeline_start || '', t.tlStart || ''],
    ['tlEnd', row.timeline_end || '', t.tlEnd || ''],
    ['budget', Number(row.budget) || 0, Number(t.budget) || 0],
    ['notes', row.notes || '', t.notes || ''],
    ['hidden', row.archived_at || '', t.hiddenAt || '']
  ];
  return checks.filter(([, incoming, local])=>String(incoming) !== String(local)).map(([field])=>field);
}

function changedDocFieldsFromRow(row, d){
  if(!d) return ['doc'];
  const out = [];
  if(String(row.title || '') !== String(d.title || '')) out.push('title');
  if(String(row.content || '') !== String(d.content || '')) out.push('content');
  return out;
}

function applyRealtimeRowPatch(table, payload){
  if(!DATA || !payload) return false;
  const eventType = payload.eventType || 'UPDATE';
  const row = eventType === 'DELETE' ? payload.old : payload.new;
  if(!row || row.workspace_id !== WORKSPACE_ID) return false;
  if(table === 'app_settings') return patchSettingsRow(row, eventType);
  if(table === 'app_members') return patchMemberRow(row, eventType);
  if(table === 'app_boards') return patchBoardRow(row, eventType);
  if(table === 'app_groups') return patchGroupRow(row, eventType);
  if(table === 'app_columns') return patchColumnRow(row, eventType);
  if(table === 'app_tasks') return patchTaskRow(row, eventType);
  if(table === 'app_task_values') return patchTaskValueRow(row, eventType);
  if(table === 'app_docs') return patchDocRow(row, eventType);
  if(table === 'app_assets') return patchAssetRow(row, eventType);
  if(table === 'app_automations') return patchJsonListRow('automations', row, eventType);
  if(table === 'app_integrations') return patchJsonListRow('integrations', row, eventType);
  if(table === 'app_archive_items') return patchJsonListRow('archive', row, eventType);
  if(table === 'app_updates') return patchUpdateRow(row, eventType);
  if(table === 'app_update_reads') return patchUpdateReadRow(row, eventType);
  if(table === 'app_update_bookmarks') return patchUpdateBookmarkRow(row, eventType);
  if(table === 'app_notifications') return patchNotificationRow(row, eventType);
  if(table === 'app_notification_preferences') return patchNotificationPreferenceRow(row, eventType);
  if(table === 'app_board_notification_preferences') return patchBoardNotificationPreferenceRow(row, eventType);
  return false;
}

function stampRowMeta(obj,row){
  if(!obj || !row) return obj;
  Object.defineProperties(obj, {
    _rowVersion:{value:Number(row.version) || 1, writable:true, configurable:true},
    _rowUpdatedAt:{value:row.updated_at || row.created_at || '', writable:true, configurable:true},
    _rowUpdatedBy:{value:row.updated_by || '', writable:true, configurable:true}
  });
  return obj;
}

function placeByPosition(list,item,position){
  const existing = list.indexOf(item);
  if(existing >= 0) list.splice(existing,1);
  const pos = Math.max(0, Math.min(Number.isFinite(Number(position)) ? Number(position) : list.length, list.length));
  list.splice(pos,0,item);
}

function patchSettingsRow(row,eventType){
  if(eventType === 'DELETE') return false;
  const data = row.data || {};
  DATA.settings = {
    ...(DATA.settings || {}),
    account:row.account || DATA.settings?.account || 'TaskTrack',
    currency:row.currency || DATA.settings?.currency || '$',
    language:data.language || DATA.settings?.language || 'en',
    memberRoles:data.memberRoles || DATA.settings?.memberRoles || {},
    notificationPrefs:data.notificationPrefs || DATA.settings?.notificationPrefs || {},
    assistantActions:data.assistantActions || DATA.settings?.assistantActions || {},
    guidedTourDone:data.guidedTourDone === true,
    onboardingDismissed:data.onboardingDismissed === true,
    assistantPet:data.assistantPet || DATA.settings?.assistantPet || {}
  };
  return true;
}

function patchMemberRow(row,eventType){
  DATA.team = DATA.team || [];
  const idx = DATA.team.findIndex(m=>m.id === row.id);
  if(eventType === 'DELETE'){
    if(idx >= 0) DATA.team.splice(idx,1);
    return idx >= 0;
  }
  const next = stampRowMeta({
    ...(idx >= 0 ? DATA.team[idx] : {}),
    id:row.id,
    email:normalizeEmail(row.email),
    name:row.name || displayNameFromEmail(row.email),
    initials:row.initials || initialsFor(row.name || row.email),
    color:row.color || AVATAR_COLORS[0],
    role:memberRole(row.id)
  }, row);
  if(idx >= 0) DATA.team[idx] = next;
  placeByPosition(DATA.team,next,row.position);
  return true;
}

function patchBoardRow(row,eventType){
  DATA.boards = DATA.boards || [];
  const idx = DATA.boards.findIndex(b=>b.id === row.id);
  if(eventType === 'DELETE'){
    if(idx >= 0) DATA.boards.splice(idx,1);
    if(CURRENT_BOARD === row.id) CURRENT_BOARD = DATA.boards[0]?.id || '';
    return idx >= 0;
  }
  const state = parseBoardColumnState(row.column_order);
  const existing = idx >= 0 ? DATA.boards[idx] : {id:row.id, groups:[], columns:[]};
  Object.assign(existing, {
    id:row.id,
    name:row.name || 'Untitled board',
    icon:row.icon || '📋',
    columnOrder:state.order,
    columnLabels:state.labels,
    hiddenColumns:state.hidden,
    sort:state.sort,
    personFilter:'',
    statusLabels:state.statusLabels,
    priorityLabels:state.priorityLabels,
    hiddenStatusLabels:state.hiddenStatusLabels,
    hiddenPriorityLabels:state.hiddenPriorityLabels,
    columnWidths:state.widths,
    calendarShowTimeline:state.calendarShowTimeline === true,
    groups:existing.groups || [],
    columns:existing.columns || []
  });
  stampRowMeta(existing,row);
  if(idx < 0) DATA.boards.push(existing);
  placeByPosition(DATA.boards,existing,row.position);
  return true;
}

function patchGroupRow(row,eventType){
  const b = board(row.board_id);
  const loc = findGroupLocation(row.id);
  if(eventType === 'DELETE'){
    if(loc) loc.b.groups = loc.b.groups.filter(g=>g.id !== row.id);
    return Boolean(loc);
  }
  if(!b) return false;
  const existing = loc?.g || {id:row.id, tasks:[]};
  if(loc && loc.b.id !== b.id) loc.b.groups = loc.b.groups.filter(g=>g.id !== row.id);
  Object.assign(existing, {
    id:row.id,
    name:row.name || 'Untitled group',
    color:row.color || '#0073ea',
    open:row.is_open !== false,
    tasks:existing.tasks || []
  });
  stampRowMeta(existing,row);
  if(!b.groups.includes(existing)) b.groups.push(existing);
  placeByPosition(b.groups,existing,row.position);
  return true;
}

function findGroupLocation(groupId){
  for(const b of DATA.boards || []){
    const g = (b.groups || []).find(group=>group.id === groupId);
    if(g) return {b,g};
  }
  return null;
}

function patchColumnRow(row,eventType){
  if(isInternalTaskValueColumn(row.id)) return false;
  const b = board(row.board_id);
  const existingLoc = findColumnLocation(row.id);
  if(eventType === 'DELETE'){
    if(existingLoc) existingLoc.b.columns = existingLoc.b.columns.filter(c=>c.id !== row.id);
    return Boolean(existingLoc);
  }
  if(!b) return false;
  const existing = existingLoc?.c || {id:row.id};
  if(existingLoc && existingLoc.b.id !== b.id) existingLoc.b.columns = existingLoc.b.columns.filter(c=>c.id !== row.id);
  Object.assign(existing, {
    id:row.id,
    name:row.name || 'Untitled',
    type:CUSTOM_COLUMN_TYPES[row.type] ? row.type : 'text'
  });
  stampRowMeta(existing,row);
  if(!b.columns.includes(existing)) b.columns.push(existing);
  placeByPosition(b.columns,existing,row.position);
  return true;
}

function findColumnLocation(columnId){
  for(const b of DATA.boards || []){
    const c = (b.columns || []).find(col=>col.id === columnId);
    if(c) return {b,c};
  }
  return null;
}

function patchTaskRow(row,eventType){
  const loc = findTaskInData(DATA, row.id);
  if(eventType === 'DELETE'){
    const removed = removeTaskFromData(row.id);
    return removed;
  }
  const targetList = targetTaskListForRow(row);
  if(!targetList) return false;
  const existing = loc?.t || {id:row.id, custom:{}, subtasks:[]};
  const currentParentId = loc?.parent?.id || '';
  const nextParentId = row.parent_task_id || '';
  const currentGroupId = loc?.g?.id || '';
  const nextGroupId = row.group_id || '';
  Object.assign(existing, {
    id:row.id,
    name:row.name || 'New task',
    owner:row.owner_id || '',
    assignedBy:row.assigned_by || '',
    status:row.status || 'not_started',
    priority:row.priority || 'none',
    due:row.due || '',
    tlStart:row.timeline_start || '',
    tlEnd:row.timeline_end || '',
    budget:Number(row.budget) || 0,
    notes:row.notes || '',
    completedAt:row.completed_at || '',
    hiddenAt:row.archived_at || '',
    custom:existing.custom || {},
    subtasks:existing.subtasks || []
  });
  stampRowMeta(existing,row);
  if(!loc || currentParentId !== nextParentId || currentGroupId !== nextGroupId){
    if(loc) removeTaskFromData(row.id);
    targetList.push(existing);
  }
  placeByPosition(targetList, existing, row.position);
  return true;
}

function targetTaskListForRow(row){
  const groupLoc = findGroupLocation(row.group_id);
  if(!groupLoc) return null;
  if(row.parent_task_id){
    const parent = findTaskInData(DATA, row.parent_task_id)?.t;
    if(parent){
      if(!Array.isArray(parent.subtasks)) parent.subtasks = [];
      return parent.subtasks;
    }
  }
  if(!Array.isArray(groupLoc.g.tasks)) groupLoc.g.tasks = [];
  return groupLoc.g.tasks;
}

function removeTaskFromData(taskId){
  const removeFrom = (tasks=[])=>{
    const idx = tasks.findIndex(t=>t.id === taskId);
    if(idx >= 0){
      tasks.splice(idx,1);
      return true;
    }
    return tasks.some(t=>removeFrom(t.subtasks || []));
  };
  return (DATA.boards || []).some(b=>(b.groups || []).some(g=>removeFrom(g.tasks || [])));
}

function patchTaskValueRow(row,eventType){
  const loc = findTaskInData(DATA, row.task_id);
  if(!loc?.t) return false;
  if(!loc.t.custom || typeof loc.t.custom !== 'object') loc.t.custom = {};
  if(eventType === 'DELETE') delete loc.t.custom[row.column_id];
  else loc.t.custom[row.column_id] = row.value ?? '';
  stampRowMeta(loc.t,row);
  return true;
}

function patchDocRow(row,eventType){
  DATA.docs = DATA.docs || [];
  const idx = DATA.docs.findIndex(d=>d.id === row.id);
  if(eventType === 'DELETE'){
    if(idx >= 0) DATA.docs.splice(idx,1);
    return idx >= 0;
  }
  const next = stampRowMeta({
    ...(idx >= 0 ? DATA.docs[idx] : {}),
    id:row.id,
    title:row.title || 'Untitled doc',
    template:row.template || '',
    content:row.content || '',
    updated:Number(row.updated_number) || Date.parse(row.updated_at) || Date.now()
  }, row);
  if(idx >= 0) DATA.docs[idx] = next;
  else DATA.docs.unshift(next);
  return true;
}

function patchAssetRow(row,eventType){
  DATA.assets = DATA.assets || [];
  const idx = DATA.assets.findIndex(a=>a.id === row.id);
  if(eventType === 'DELETE'){
    if(idx >= 0) DATA.assets.splice(idx,1);
    return idx >= 0;
  }
  const next = stampRowMeta({
    ...(idx >= 0 ? DATA.assets[idx] : {}),
    id:row.id,
    title:row.title || 'Untitled asset',
    kind:row.kind || 'file',
    mimeType:row.mime_type || '',
    size:Number(row.size_bytes) || 0,
    source:row.source || 'upload',
    storageBucket:row.storage_bucket || ASSET_BUCKET,
    storagePath:row.storage_path || '',
    externalUrl:row.external_url || '',
    description:row.description || '',
    createdBy:row.created_by || '',
    data:row.data && typeof row.data === 'object' ? row.data : {},
    updated:Number(row.updated_number) || Date.parse(row.updated_at) || Date.now()
  }, row);
  if(idx >= 0) DATA.assets[idx] = next;
  else DATA.assets.unshift(next);
  return true;
}

function patchJsonListRow(key,row,eventType){
  DATA[key] = DATA[key] || [];
  const idx = DATA[key].findIndex(item=>item.id === row.id);
  if(eventType === 'DELETE'){
    if(idx >= 0) DATA[key].splice(idx,1);
    return idx >= 0;
  }
  const next = stampRowMeta({...(row.data || {}), id:row.id}, row);
  if(idx >= 0) DATA[key][idx] = next;
  else DATA[key].unshift(next);
  return true;
}

function patchUpdateRow(row,eventType){
  DATA.updates = DATA.updates || [];
  const idx = DATA.updates.findIndex(u=>u.id === row.id);
  if(eventType === 'DELETE'){
    if(idx >= 0) DATA.updates.splice(idx,1);
    return idx >= 0;
  }
  const previous = idx >= 0 ? DATA.updates[idx] : {};
  const next = stampRowMeta({
    ...previous,
    id:row.id,
    authorId:row.author_id || '',
    boardId:row.board_id || '',
    groupId:row.group_id || '',
    taskId:row.task_id || '',
    body:row.body || '',
    assetIds:Array.isArray(row.asset_ids) ? row.asset_ids : [],
    mentions:Array.isArray(row.mentions) ? row.mentions : [],
    readBy:previous.readBy || [],
    bookmarkedBy:previous.bookmarkedBy || [],
    scheduledAt:row.scheduled_at ? Date.parse(row.scheduled_at) : null,
    createdAt:Date.parse(row.created_at) || Date.now()
  }, row);
  if(idx >= 0) DATA.updates[idx] = next;
  else DATA.updates.unshift(next);
  DATA.updates.sort((a,b)=>(b.createdAt||0)-(a.createdAt||0));
  return true;
}

function patchUpdateReadRow(row,eventType){
  const u = (DATA.updates || []).find(update=>update.id === row.update_id);
  if(!u) return false;
  if(!Array.isArray(u.readBy)) u.readBy = [];
  if(eventType === 'DELETE') u.readBy = u.readBy.filter(id=>id !== row.member_id);
  else if(!u.readBy.includes(row.member_id)) u.readBy.push(row.member_id);
  return true;
}

function patchUpdateBookmarkRow(row,eventType){
  const u = (DATA.updates || []).find(update=>update.id === row.update_id);
  if(!u) return false;
  if(!Array.isArray(u.bookmarkedBy)) u.bookmarkedBy = [];
  if(eventType === 'DELETE') u.bookmarkedBy = u.bookmarkedBy.filter(id=>id !== row.member_id);
  else if(!u.bookmarkedBy.includes(row.member_id)) u.bookmarkedBy.push(row.member_id);
  return true;
}

function notificationFromRow(row){
  return stampRowMeta({
    id:row.id,
    recipientId:row.recipient_member_id || '',
    recipientUserId:row.recipient_user_id || '',
    recipientEmail:normalizeEmail(row.recipient_email),
    actorId:row.actor_member_id || '',
    boardId:row.board_id || '',
    groupId:row.group_id || '',
    taskId:row.task_id || '',
    updateId:row.update_id || '',
    type:row.type || 'general',
    title:row.title || '',
    body:row.body || '',
    channels:row.channels && typeof row.channels === 'object' ? row.channels : {},
    data:row.data && typeof row.data === 'object' ? row.data : {},
    dedupeKey:row.dedupe_key || '',
    readAt:row.read_at || '',
    emailedAt:row.emailed_at || '',
    emailStatus:row.email_status || '',
    emailAttempts:Number(row.email_attempts) || 0,
    desktopAt:row.desktop_at || '',
    pushedAt:row.pushed_at || '',
    pushStatus:row.push_status || '',
    pushAttempts:Number(row.push_attempts) || 0,
    createdAt:Date.parse(row.created_at) || Date.now()
  }, row);
}

function patchNotificationRow(row,eventType){
  DATA.notifications = DATA.notifications || [];
  const idx = DATA.notifications.findIndex(n=>n.id === row.id);
  if(eventType === 'DELETE'){
    if(idx >= 0) DATA.notifications.splice(idx,1);
    return idx >= 0;
  }
  const next = notificationFromRow(row);
  if(idx >= 0) DATA.notifications[idx] = {...DATA.notifications[idx], ...next};
  else DATA.notifications.unshift(next);
  DATA.notifications.sort((a,b)=>(b.createdAt||0)-(a.createdAt||0));
  return true;
}

function patchNotificationPreferenceRow(row,eventType){
  if(!DATA.notificationPrefs || typeof DATA.notificationPrefs !== 'object') DATA.notificationPrefs = {};
  if(eventType === 'DELETE') delete DATA.notificationPrefs[row.member_id];
  else DATA.notificationPrefs[row.member_id] = {...NOTIFICATION_DEFAULTS, ...(row.prefs || {})};
  return true;
}

function patchBoardNotificationPreferenceRow(row,eventType){
  if(!DATA.boardNotificationPrefs || typeof DATA.boardNotificationPrefs !== 'object') DATA.boardNotificationPrefs = {};
  if(!DATA.boardNotificationPrefs[row.board_id]) DATA.boardNotificationPrefs[row.board_id] = {};
  if(eventType === 'DELETE') delete DATA.boardNotificationPrefs[row.board_id][row.member_id];
  else DATA.boardNotificationPrefs[row.board_id][row.member_id] = {
    mode:['custom','receive_all','mute'].includes(row.mode) ? row.mode : 'custom',
    prefs:row.prefs && typeof row.prefs === 'object' ? row.prefs : {}
  };
  return true;
}

async function refreshIntegrationConnections({ rerender=false } = {}){
  INTEGRATION_CONNECTIONS=JSON.parse(localStorage.getItem('portfolio_integrations')||'[]');
  if(rerender && DATA) render();
  return;

  if(!supabase || !SESSION?.user){
    INTEGRATION_CONNECTIONS = [];
    GOOGLE_MEET_SUBSCRIPTION = null;
    return;
  }
  const { data, error } = await supabase
    .from('integration_connections')
    .select('provider, external_account_id, account_label, scopes, metadata, connected_at')
    .eq('user_id', SESSION.user.id)
    .order('connected_at', { ascending:false });
  if(!error) INTEGRATION_CONNECTIONS = data || [];
  GOOGLE_MEET_SUBSCRIPTION = null;
  if(!error && INTEGRATION_CONNECTIONS.some(connection=>connection.provider === 'google_meet')){
    const {data:meetSubscription,error:meetError} = await supabase
      .from('app_google_meet_subscriptions')
      .select('id,state,expires_at,last_error,event_types,updated_at')
      .eq('workspace_id',WORKSPACE_ID)
      .eq('user_id',SESSION.user.id)
      .order('updated_at',{ascending:false})
      .limit(1)
      .maybeSingle();
    if(!meetError) GOOGLE_MEET_SUBSCRIPTION=meetSubscription || null;
  }
  if(rerender && (VIEW==='integrations' || VIEW==='meetings' || VIEW==='calendar')) renderContent();
}

function consumeIntegrationNotice(){
  if(CALLBACK_NOTICE_SHOWN) return;
  const url = new URL(window.location.href);
  const provider = url.searchParams.get('integration');
  const status = url.searchParams.get('status');
  if(!provider || !status) return;
  CALLBACK_NOTICE_SHOWN = true;
  url.searchParams.delete('integration');
  url.searchParams.delete('status');
  url.searchParams.delete('message');
  window.history.replaceState({}, '', url.toString());
  const label = integrationLabel(provider);
  setTimeout(()=>{
    toast(status === 'connected' ? `${label} connected` : `${label} connection needs setup`);
    void refreshIntegrationConnections({ rerender:true });
  }, 80);
}

/* ============================================================
   ICONS (inline svg)
   ============================================================ */
const I = {
  home:'<svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M3 9.5L12 3l9 6.5V20a1 1 0 01-1 1h-5v-6H9v6H4a1 1 0 01-1-1z"/></svg>',
  feed:'<svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="4" y="4" width="16" height="16" rx="2"/><path d="M8 9h8M8 13h5M8 17h7"/></svg>',
  board:'<svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="3" y="3" width="18" height="18" rx="2"/><path d="M3 9h18M9 21V9"/></svg>',
  dash:'<svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="3" y="3" width="8" height="10" rx="1"/><rect x="13" y="3" width="8" height="6" rx="1"/><rect x="13" y="11" width="8" height="10" rx="1"/><rect x="3" y="15" width="8" height="6" rx="1"/></svg>',
  work:'<svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="3" y="7" width="18" height="13" rx="2"/><path d="M8 7V5a2 2 0 012-2h4a2 2 0 012 2v2"/></svg>',
  meetings:'<svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M4 5h16v14H4z"/><path d="M8 3v4M16 3v4M4 9h16M8 13h3M8 16h6"/></svg>',
  integrations:'<svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M8 8h8v8H8z"/><path d="M3 12h5M16 12h5M12 3v5M12 16v5"/></svg>',
  docs:'<svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M14 3H6a2 2 0 00-2 2v14a2 2 0 002 2h12a2 2 0 002-2V9z"/><path d="M14 3v6h6M9 13h6M9 17h6"/></svg>',
  gear:'<svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="12" cy="12" r="3"/><path d="M19 12a7 7 0 00-.1-1l2-1.5-2-3.5-2.3 1a7 7 0 00-1.7-1l-.4-2.5h-4l-.4 2.5a7 7 0 00-1.7 1l-2.3-1-2 3.5L4.1 11a7 7 0 000 2l-2 1.5 2 3.5 2.3-1a7 7 0 001.7 1l.4 2.5h4l.4-2.5a7 7 0 001.7-1l2.3 1 2-3.5-2-1.5a7 7 0 00.1-1z"/></svg>',
  search:'<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="11" cy="11" r="7"/><path d="M21 21l-4-4"/></svg>',
  plus:'<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2"><path d="M12 5v14M5 12h14"/></svg>',
  filter:'<svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M3 5h18l-7 8v6l-4 2v-8z"/></svg>',
  sort:'<svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M8 7h10M8 12h7M8 17h4M4 7v10M2 15l2 2 2-2"/></svg>',
  hide:'<svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M3 3l18 18"/><path d="M10.6 10.6A2 2 0 0013.4 13.4"/><path d="M9.9 4.2A10.7 10.7 0 0112 4c5 0 9 4.5 10 8a11.8 11.8 0 01-2.2 3.8"/><path d="M6.1 6.1C4.1 7.5 2.7 9.6 2 12c1 3.5 5 8 10 8a10.7 10.7 0 004.1-.8"/></svg>',
  cal:'<svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="3" y="4" width="18" height="17" rx="2"/><path d="M3 9h18M8 2v4M16 2v4"/></svg>',
  columns:'<svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="3" y="4" width="18" height="16" rx="2"/><path d="M9 4v16M15 4v16"/></svg>',
  automation:'<svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M13 2L4 14h7l-1 8 10-13h-7l0-7z"/></svg>',
  copy:'<svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="8" y="8" width="12" height="12" rx="2"/><path d="M4 16V6a2 2 0 012-2h10"/></svg>',
  archive:'<svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="3" y="4" width="18" height="4" rx="1"/><path d="M5 8v12h14V8M10 12h4"/></svg>',
  trash:'<svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M3 6h18M8 6V4h8v2M6 6l1 14h10l1-14"/></svg>',
  convert:'<svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M7 7h10v10H7z"/><path d="M4 12h3M17 12h3M12 4v3M12 17v3"/></svg>',
  move:'<svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M5 12h14M13 6l6 6-6 6"/></svg>',
  close:'<svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M18 6L6 18M6 6l12 12"/></svg>',
  bookmark:'<svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M6 4h12v17l-6-4-6 4z"/></svg>',
  mention:'<svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="12" cy="12" r="4"/><path d="M16 8v5a3 3 0 006 0 10 10 0 10-4 8"/></svg>',
  dots:'<svg width="18" height="18" viewBox="0 0 24 24" fill="currentColor"><circle cx="5" cy="12" r="2"/><circle cx="12" cy="12" r="2"/><circle cx="19" cy="12" r="2"/></svg>',
  bell:'<svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M18 8a6 6 0 00-12 0c0 7-3 9-3 9h18s-3-2-3-9M13.7 21a2 2 0 01-3.4 0"/></svg>',
  download:'<svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M12 3v12m0 0l-4-4m4 4l4-4M5 21h14"/></svg>',
  link:'<svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M10 13a5 5 0 007 0l2-2a5 5 0 00-7-7l-1.2 1.2"/><path d="M14 11a5 5 0 00-7 0l-2 2a5 5 0 007 7l1.2-1.2"/></svg>',
  spark:'<svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M12 3l1.9 5.8L20 11l-6.1 2.2L12 19l-1.9-5.8L4 11l6.1-2.2L12 3z"/></svg>',
  assistant:'<svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="4" y="7" width="16" height="12" rx="3"/><path d="M9 7V4h6v3M9 13h.01M15 13h.01M10 17h4"/></svg>'
};

/* ---------- board column model ---------- */
const BASE_COLUMNS = [
  {id:'name', label:'Task', builtIn:true},
  {id:'owner', label:'Owner', builtIn:true},
  {id:'status', label:'Status', builtIn:true},
  {id:'color', label:'Color', builtIn:true},
  {id:'due', label:'Due Date', builtIn:true},
  {id:'timeline', label:'Timeline', builtIn:true},
  {id:'priority', label:'Priority', builtIn:true},
  {id:'budget', label:'Budget', builtIn:true},
  {id:'notes', label:'Notes', builtIn:true},
  {id:'assignedBy', label:'Assigned by', builtIn:true}
];
const CUSTOM_COLUMN_TYPES = {
  text:'Text',
  url:'URL / Link',
  recurrence:'Recurring',
  number:'Number',
  date:'Date',
  status:'Status',
  document:'Document',
  files:'Files / Media'
};
const RECURRENCE_OPTIONS = {
  off:{label:'Off', color:'#c4c4c4'},
  daily:{label:'Daily', color:'#0073ea'},
  weekly:{label:'Weekly', color:'#a25ddc'},
  monthly:{label:'Monthly', color:'#00a37a'}
};
const RECURRENCE_UNITS = {
  days:{label:'days', singular:'day'},
  weeks:{label:'weeks', singular:'week'},
  months:{label:'months', singular:'month'}
};
const TASK_PHASE_COLORS = ['#0073ea','#00c875','#fdab3d','#e2445c','#a25ddc','#579bfc','#401694','#00c2ff','#ff5ac4','#cab641','#037f4c','#784bd1'];
const BOARD_VIEWS = [
  {id:'all', label:'All'},
  {id:'mine', label:'My tasks'},
  {id:'due_week', label:'Due this week'},
  {id:'overdue', label:'Overdue'},
  {id:'blocked', label:'Blocked'},
  {id:'unassigned', label:'Unassigned'},
  {id:'review', label:'Needs review'},
  {id:'hidden', label:'Hidden'}
];
const BOARD_EMOJIS = ['📋','📣','📅','🎯','🚀','✨','💡','📈','🧾','✅','🔥','⭐','💬','🎨','🛒','📦','🏷️','🧪','🎬','📌'];

function normalizeData(){
  if(!DATA.settings) DATA.settings = { account:'TaskTrack', currentUser:'', currency:'$' };
  if(DATA.settings.currency === '₩') DATA.settings.currency = '$';
  if(!LANGUAGE_OPTIONS[DATA.settings.language]) DATA.settings.language = 'en';
  if(!DATA.settings.memberRoles || typeof DATA.settings.memberRoles !== 'object') DATA.settings.memberRoles = {};
  DATA.settings.notificationPrefs = notificationPrefs();
  DATA.settings.assistantActions = assistantActionPrefs();
  DATA.settings.guidedTourDone = DATA.settings.guidedTourDone === true;
  DATA.settings.onboardingDismissed = DATA.settings.onboardingDismissed === true;
  DATA.settings.assistantPet = assistantPetSettings(DATA.settings.assistantPet);
  if(!Array.isArray(DATA.team)) DATA.team = [];
  if(!Array.isArray(DATA.boards)) DATA.boards = [];
  if(!Array.isArray(DATA.docs)) DATA.docs = [];
  if(!Array.isArray(DATA.automations)) DATA.automations = [];
  DATA.team.forEach((m,i)=>{
    if(!m.id) m.id = uid('u');
    if(m.email) m.email = normalizeEmail(m.email);
    if(!m.name) m.name = displayNameFromEmail(m.email) || `Member ${i+1}`;
    if(!m.initials) m.initials = initialsFor(m.name);
    if(!m.color) m.color = AVATAR_COLORS[i % AVATAR_COLORS.length];
    m.role = DATA.settings.memberRoles[m.id] === 'admin' ? 'admin' : (m.role === 'admin' ? 'admin' : 'member');
    DATA.settings.memberRoles[m.id] = m.role;
  });
  if(DATA.team.length && !DATA.team.some(m=>m.role === 'admin')){
    DATA.team[0].role = 'admin';
    DATA.settings.memberRoles[DATA.team[0].id] = 'admin';
  }
  Object.keys(DATA.settings.memberRoles).forEach(id=>{
    if(!DATA.team.some(m=>m.id === id)) delete DATA.settings.memberRoles[id];
  });
  DATA.boards.forEach(normalizeBoard);
  DATA.automations = DATA.automations.filter(a=>a && a.boardId && DATA.boards.some(b=>b.id===a.boardId));
  if(!Array.isArray(DATA.integrations)) DATA.integrations = [];
  if(!Array.isArray(DATA.assets)) DATA.assets = [];
  if(!Array.isArray(DATA.updates)) DATA.updates = [];
  if(!Array.isArray(DATA.notifications)) DATA.notifications = [];
  if(!DATA.notificationPrefs || typeof DATA.notificationPrefs !== 'object') DATA.notificationPrefs = {};
  if(!DATA.boardNotificationPrefs || typeof DATA.boardNotificationPrefs !== 'object') DATA.boardNotificationPrefs = {};
  DATA.updates.forEach(u=>{
    if(!u.id) u.id = uid('upd');
    if(!Array.isArray(u.mentions)) u.mentions = [];
    if(!Array.isArray(u.readBy)) u.readBy = [];
    if(!Array.isArray(u.bookmarkedBy)) u.bookmarkedBy = [];
    if(!u.createdAt) u.createdAt = Date.now();
    if(isAssignmentUpdateBody(u.body)){
      u.kind = 'audit';
      u.visible = false;
    }
  });
  DATA.notifications.forEach(n=>{
    if(!n.id) n.id = uid('ntf');
    if(!n.recipientId) n.recipientId = currentMemberId();
    if(n.recipientEmail) n.recipientEmail = normalizeEmail(n.recipientEmail);
    if(!n.type) n.type = 'general';
    if(!n.title) n.title = 'TaskTrack notification';
    if(!n.body) n.body = '';
    if(!n.channels || typeof n.channels !== 'object') n.channels = {};
    if(!n.data || typeof n.data !== 'object') n.data = {};
    if(!n.pushStatus) n.pushStatus = n.channels.push === false ? 'skipped' : 'pending';
    if(!Number.isFinite(Number(n.pushAttempts))) n.pushAttempts = 0;
    if(!n.createdAt) n.createdAt = Date.now();
  });
  Object.keys(DATA.notificationPrefs).forEach(memberId=>{
    DATA.notificationPrefs[memberId] = {...NOTIFICATION_DEFAULTS, ...(DATA.notificationPrefs[memberId] || {})};
  });
  if(!Array.isArray(DATA.archive)) DATA.archive = [];
  DATA.docs.forEach(d=>{ if(!d.id) d.id = uid('doc'); if(!d.updated) d.updated = Date.now(); });
  DATA.assets.forEach(normalizeAsset);
  DATA.automations.forEach(a=>{ if(!a.id) a.id = uid('a'); });
  DATA.integrations.forEach(i=>{ if(!i.id) i.id = uid('int'); });
  DATA.archive.forEach(a=>{ if(!a.id) a.id = uid('arch'); });
}
function normalizeAsset(a){
  if(!a.id) a.id = uid('asset');
  if(!a.title) a.title = 'Untitled asset';
  if(!a.kind) a.kind = assetKindForMime(a.mimeType || '');
  if(!a.source) a.source = a.externalUrl ? 'link' : 'upload';
  if(!a.storageBucket) a.storageBucket = ASSET_BUCKET;
  if(!a.data || typeof a.data !== 'object') a.data = {};
  if(!a.updated) a.updated = Date.now();
  if(!a.createdBy) a.createdBy = currentMemberId();
  a.size = Number(a.size) || Number(a.sizeBytes) || Number(a.size_bytes) || 0;
}
function normalizeBoard(b){
  if(!b.id) b.id = uid('b');
  if(!b.name) b.name = 'Untitled board';
  if(!b.icon) b.icon = '📋';
  if(!Array.isArray(b.groups)) b.groups = [];
  if(!Array.isArray(b.columns)) b.columns = [];
  b.columns = b.columns.filter(c=>!isInternalTaskValueColumn(c?.id));
  b.columns.forEach(c=>{
    if(!c.id) c.id = uid('c');
    if(!c.name) c.name = 'Untitled';
    if(!CUSTOM_COLUMN_TYPES[c.type]) c.type = 'text';
  });
  const validIds = [...BASE_COLUMNS.map(c=>c.id), ...b.columns.map(c=>c.id)];
  if(!Array.isArray(b.columnOrder)) b.columnOrder = validIds.slice();
  b.columnOrder = b.columnOrder.filter(id=>validIds.includes(id));
  BASE_COLUMNS.map(c=>c.id).forEach((id,index)=>{
    if(b.columnOrder.includes(id)) return;
    const nextId = BASE_COLUMNS.slice(index + 1).map(c=>c.id).find(next=>b.columnOrder.includes(next));
    if(nextId) b.columnOrder.splice(b.columnOrder.indexOf(nextId), 0, id);
    else b.columnOrder.push(id);
  });
  b.columns.forEach(c=>{ if(!b.columnOrder.includes(c.id)) b.columnOrder.push(c.id); });
  if(!b.columnLabels || typeof b.columnLabels !== 'object') b.columnLabels = {};
  Object.keys(b.columnLabels).forEach(id=>{ if(!validIds.includes(id) || !b.columnLabels[id]) delete b.columnLabels[id]; });
  if(!Array.isArray(b.hiddenColumns)) b.hiddenColumns = [];
  b.hiddenColumns = [...new Set(b.hiddenColumns)].filter(id=>validIds.includes(id) && id !== 'name');
  if(!b.columnWidths || typeof b.columnWidths !== 'object') b.columnWidths = {};
  Object.keys(b.columnWidths).forEach(id=>{
    const width = Number(b.columnWidths[id]);
    if(!validIds.includes(id) || !Number.isFinite(width)) delete b.columnWidths[id];
    else b.columnWidths[id] = Math.min(720, Math.max(80, Math.round(width)));
  });
  if(!b.sort || typeof b.sort !== 'object') b.sort = {};
  if(!validIds.includes(b.sort.columnId)) b.sort = {};
  if(b.sort.direction !== 'desc') b.sort.direction = b.sort.columnId ? 'asc' : '';
  b.personFilter = '';
  b.calendarShowTimeline = b.calendarShowTimeline === true;
  if(!b.statusLabels || typeof b.statusLabels !== 'object') b.statusLabels = {};
  if(!b.priorityLabels || typeof b.priorityLabels !== 'object') b.priorityLabels = {};
  normalizeBoardLabelMap(b.statusLabels, STATUS);
  normalizeBoardLabelMap(b.priorityLabels, PRIORITY);
  if(!Array.isArray(b.hiddenStatusLabels)) b.hiddenStatusLabels = [];
  if(!Array.isArray(b.hiddenPriorityLabels)) b.hiddenPriorityLabels = [];
  b.hiddenStatusLabels = [...new Set(b.hiddenStatusLabels)].filter(key=>STATUS[key]);
  b.hiddenPriorityLabels = [...new Set(b.hiddenPriorityLabels)].filter(key=>PRIORITY[key]);
  const validTaskValueColumnIds = taskValueColumnIdsForBoard(b);
  b.groups.forEach(g=>{
    if(!g.id) g.id = uid('g');
    if(!g.name) g.name = 'Untitled group';
    if(!g.color) g.color = '#0073ea';
    if(!Array.isArray(g.tasks)) g.tasks = [];
    g.tasks.forEach(t=>{
      normalizeTask(t);
      pruneTaskCustomValues(t, validTaskValueColumnIds);
    });
  });
}
function normalizeTask(t){
  if(!t.id) t.id = uid('t');
  if(!t.name) t.name = 'New task';
  if(!t.custom || typeof t.custom!=='object') t.custom = {};
  if(!Array.isArray(t.subtasks)) t.subtasks = [];
  t.hiddenAt = t.hiddenAt || '';
  if(t.color) setTaskColor(t,t.color);
  normalizeTaskSchedule(t);
  if(t.recurring || t.custom[TASK_RECURRING_KEY] != null || t.custom[TASK_RECURRENCE_KEY] != null) setTaskRecurrence(t, taskRecurrence(t));
  if(t.status !== 'done' && t.completedAt) t.completedAt = '';
  const oldPersistedOpen = t.custom[TASK_SUBTASKS_OPEN_KEY] === true || t.custom[TASK_SUBTASKS_OPEN_KEY] === 'true';
  t.subtasksOpen = Boolean(t.subtasksOpen || oldPersistedOpen || OPEN_SUBTASK_TASKS.has(t.id));
  if(t.subtasksOpen) OPEN_SUBTASK_TASKS.add(t.id);
  delete t.custom[TASK_SUBTASKS_OPEN_KEY];
  t.subtasks.forEach(normalizeTask);
}
function allBoardColumns(b){
  normalizeBoard(b);
  const labels = b.columnLabels || {};
  const byId = new Map(BASE_COLUMNS.map(c=>[c.id,{...c,label:labels[c.id] || c.label,custom:false,builtIn:true}]));
  b.columns.forEach(c=>byId.set(c.id,{...c,label:labels[c.id] || c.name,custom:true,builtIn:false}));
  return b.columnOrder.map(id=>byId.get(id)).filter(Boolean);
}
function boardColumns(b){ return allBoardColumns(b).filter(c=>!b.hiddenColumns?.includes(c.id)); }
function defaultColumnWidth(id){ return DEFAULT_COLUMN_WIDTHS[id] || 160; }
function columnWidth(b,id){
  const n = Number(b?.columnWidths?.[id]);
  return Math.min(720, Math.max(80, Number.isFinite(n) ? n : defaultColumnWidth(id)));
}
function columnWidthStyle(b,id){
  const width = columnWidth(b,id);
  return `width:${width}px;min-width:${width}px`;
}
function findColumn(b,cid){ return b.columns.find(c=>c.id===cid); }
function normalizeBoardLabelMap(map, base){
  Object.keys(map).forEach((key,i)=>{
    const label = typeof map[key] === 'string' ? map[key] : map[key]?.label;
    const color = typeof map[key] === 'object' ? map[key]?.color : '';
    if(!label) delete map[key];
    else map[key] = { label:String(label), color:color || base[key]?.color || LABEL_COLORS[i % LABEL_COLORS.length] };
  });
}
function boardStatusOptions(b){
  normalizeBoard(b);
  const hidden = new Set(b.hiddenStatusLabels || []);
  return Object.fromEntries(Object.entries({...STATUS,...(b.statusLabels||{})}).filter(([key])=>!hidden.has(key)));
}
function boardPriorityOptions(b){
  normalizeBoard(b);
  const hidden = new Set(b.hiddenPriorityLabels || []);
  return Object.fromEntries(Object.entries({...PRIORITY,...(b.priorityLabels||{})}).filter(([key])=>!hidden.has(key)));
}
function statusOptionFor(b,key){ return boardStatusOptions(b)[key] || {label:String(key||'Not Started'),color:'#c4c4c4'}; }
function priorityOptionFor(b,key){ return boardPriorityOptions(b)[key] || {label:String(key||'None'),color:'#c4c4c4'}; }

/* ============================================================
   RENDER — shell
   ============================================================ */
function captureScrollState(){
  const c = $('#content');
  if(!c) return null;
  const tables = [...c.querySelectorAll('.group[data-group] .tbl-wrap')].map(wrap=>({
    groupId:wrap.closest('.group')?.dataset.group || '',
    top:wrap.scrollTop,
    left:wrap.scrollLeft
  })).filter(x=>x.groupId);
  return {
    view:c.dataset.renderView || '',
    board:c.dataset.renderBoard || '',
    top:c.scrollTop,
    left:c.scrollLeft,
    tables
  };
}

function restoreScrollState(state){
  if(!state) return;
  const c = $('#content');
  if(!c) return;
  if(state.view !== c.dataset.renderView || state.board !== c.dataset.renderBoard) return;
  const apply = ()=>{
    c.scrollTop = state.top || 0;
    c.scrollLeft = state.left || 0;
    (state.tables || []).forEach(item=>{
      const group = c.querySelector(`.group[data-group="${CSS.escape(item.groupId)}"]`);
      const wrap = group?.querySelector('.tbl-wrap');
      if(wrap){
        wrap.scrollTop = item.top || 0;
        wrap.scrollLeft = item.left || 0;
      }
    });
  };
  apply();
  requestAnimationFrame(apply);
}

function render({preserveScroll=true} = {}){
  const scrollState = preserveScroll ? captureScrollState() : null;
  const app = $('#app');
  const captureBanner = VIEW === 'meetings'
    ? ''
    : renderMeetingCaptureBanner(meetingCaptureViewModel(), { global:true });
  app.innerHTML = `${sidebar()}<div class="main">${topbar()}${captureBanner}<div class="content" id="content"></div></div>${opsAssistantWidget()}${taskDetailDrawer()}`;
  renderContent({scrollState});
  wireSidebar();
  wireTopbar();
  wireOpsAssistant();
  wireTaskDetailDrawer();
  consumeIntegrationNotice();
  maybeShowGuidedTour();
  renderQueuedRemoteBanner();
  trackCurrentPresence();
}

function sidebar(){
  const unreadUpdates = relevantFeedUpdates().filter(u=>!(u.readBy||[]).includes(currentMemberId())).length;
  const boardsHtml = DATA.boards.map(b=>`
    <div class="nav-item board-link ${VIEW==='board'&&CURRENT_BOARD===b.id?'active':''}">
      <a class="board-open" href="${boardHash(b.id)}" data-board="${b.id}" title="${esc(b.name)}">
        <span class="ic board-emoji">${boardIconHtml(b)}</span><span class="txt">${esc(b.name)}</span>
      </a>
      <button class="board-menu-trigger" data-board-menu="${b.id}" title="Board options">${I.dots}</button>
    </div>`).join('');
  return `
  <aside class="sidebar">
    <div class="sb-brand">
      <div class="sb-logo"><img src="./favicon.png" alt="TaskTrack"></div>
      <div><div class="brand-title">TaskTrack</div><div class="sub">Task Manager</div></div>
    </div>
    <div class="sb-scroll">
      <a class="nav-item ${VIEW==='home'?'active':''}" href="${routeHref('home')}" data-view="home"><span class="ic">${I.home}</span><span class="txt">${esc(tr('workspaceHome'))}</span></a>
      <a class="nav-item ${VIEW==='updates'?'active':''}" href="${routeHref('updates')}" data-view="updates"><span class="ic">${I.feed}</span><span class="txt">${esc(tr('updateFeed'))}</span>${unreadUpdates?`<span class="nav-badge">${unreadUpdates}</span>`:''}</a>
      <a class="nav-item ${VIEW==='dashboard'?'active':''}" href="${routeHref('dashboard')}" data-view="dashboard"><span class="ic">${I.dash}</span><span class="txt">${esc(tr('dashboard'))}</span></a>
      <a class="nav-item ${VIEW==='calendar'?'active':''}" href="${routeHref('calendar')}" data-view="calendar"><span class="ic">${I.cal}</span><span class="txt">${esc(tr('calendar'))}</span></a>
      <a class="nav-item ${VIEW==='closing'?'active':''}" href="${routeHref('closing')}" data-view="closing"><span class="ic">${I.work}</span><span class="txt">${esc(tr('monthlyClosing'))}</span></a>
      <a class="nav-item ${VIEW==='mywork'?'active':''}" href="${routeHref('mywork')}" data-view="mywork"><span class="ic">${I.work}</span><span class="txt">${esc(tr('myWork'))}</span></a>
      <a class="nav-item ${VIEW==='meetings'?'active':''}" href="${routeHref('meetings')}" data-view="meetings"><span class="ic">${I.meetings}</span><span class="txt">${esc(tr('meetings'))}</span></a>
      <a class="nav-item ${VIEW==='integrations'?'active':''}" href="${routeHref('integrations')}" data-view="integrations"><span class="ic">${I.integrations}</span><span class="txt">${esc(tr('integrations'))}</span></a>
      <a class="nav-item ${VIEW==='docs'?'active':''}" href="${routeHref('docs')}" data-view="docs"><span class="ic">${I.docs}</span><span class="txt">${esc(tr('docsHub'))}</span></a>
      <div class="sb-section">Boards</div>
      ${boardsHtml||'<div style="padding:6px 16px;color:var(--ink3);font-size:13px">No boards yet</div>'}
      <button class="add-board-btn" data-action="add-board"><span class="ic">${I.plus}</span> Add board</button>
    </div>
    <div class="sb-foot">
      <a class="nav-item ${VIEW==='settings'?'active':''}" href="${routeHref('settings')}" data-view="settings"><span class="ic">${I.gear}</span><span class="txt">${esc(tr('settings'))}</span></a>
    </div>
  </aside>`;
}

function boardIconHtml(b){
  const icon = String(b?.icon || '📋').trim() || '📋';
  return esc(icon);
}

function boardHash(boardId){
  return `#board=${encodeURIComponent(boardId)}`;
}

function boardUrl(boardId){
  const url = new URL(window.location.href);
  url.hash = boardHash(boardId).slice(1);
  return url.toString();
}

function applyHashRoute(){
  const raw = window.location.hash.replace(/^#/, '');
  const meetingMatch = raw.match(/^meeting=(.+)$/);
  if(meetingMatch){
    VIEW = 'meetings';
    MEETING_SELECTED_ID = decodeURIComponent(meetingMatch[1]);
    return;
  }
  const match = raw.match(/^board=(.+)$/);
  const id = match ? decodeURIComponent(match[1]) : '';
  if(id && board(id)){
    VIEW = 'board';
    CURRENT_BOARD = id;
    return;
  }
  const view = ROUTE_VIEWS[raw];
  if(view) VIEW = view;
  else if(!raw) VIEW = 'home';
  if(VIEW === 'board' && CURRENT_BOARD && !board(CURRENT_BOARD)) VIEW = 'home';
  if(VIEW !== 'board' && !CURRENT_BOARD && DATA.boards[0]) CURRENT_BOARD = DATA.boards[0].id;
  }

function topbar(){
  let title=tr('workspaceHome'), sub='';
  const currentBoard = VIEW==='board' ? board(CURRENT_BOARD) : null;
  if(VIEW==='board'){ title=currentBoard?currentBoard.name:'Board'; sub='Task board'; }
  else if(VIEW==='updates'){ title=tr('updateFeed'); sub='Mentions, task updates and team notes'; }
  else if(VIEW==='dashboard'){ title=tr('dashboard'); }
  else if(VIEW==='calendar'){ title=tr('calendar'); sub=CURRENT_BOARD && board(CURRENT_BOARD) ? board(CURRENT_BOARD).name : 'All task boards'; }
  else if(VIEW==='closing'){ title=tr('monthlyClosing'); sub='Review work across the workspace'; }
  else if(VIEW==='mywork'){ title=tr('myWork'); }
  else if(VIEW==='meetings'){ title=tr('meetings'); sub='Transcript-powered meeting notes'; }
  else if(VIEW==='integrations'){ title=tr('integrations'); sub='Connected apps and account access'; }
  else if(VIEW==='docs'){ title=tr('docsHub'); }
  else if(VIEW==='settings'){ title=tr('settings'); }
  const me = member(currentMemberId()) || DATA.team[0] || {initials:'?',color:'#999'};
  const unreadNoticeCount = unreadNotifications().length;
  return `
  <div class="topbar">
    <div class="tb-left">
      ${currentBoard?`<button class="tb-board-icon" data-change-board-icon="${currentBoard.id}" title="Change board emoji">${boardIconHtml(currentBoard)}</button>`:''}
      <div>
        <h1 class="tb-title ${VIEW==='board'?'editable-title':''}" ${VIEW==='board'&&CURRENT_BOARD?`contenteditable="true" spellcheck="false" data-rename-board="${CURRENT_BOARD}" title="Click to rename board"`:''}>${esc(title)}</h1>
        ${sub?`<div class="tb-sub">${esc(sub)}</div>`:''}
      </div>
    </div>
    <div class="global-search">
      <div class="global-search-box">
        ${I.search}
        <input id="global-search" value="${esc(GLOBAL_SEARCH)}" placeholder="${esc(tr('searchPlaceholder'))}" autocomplete="off">
        <div id="global-search-results" class="global-search-results ${GLOBAL_SEARCH.trim()?'open':''}">
          ${globalSearchMenuHtml(GLOBAL_SEARCH)}
        </div>
      </div>
    </div>
    <div class="tb-right">
      ${isSupabaseConfigured && SESSION ? `<span class="cloud-status" id="cloud-status">${esc(CLOUD_STATUS)}</span>` : ''}
      <button class="icon-btn" data-action="export" title="Export data">${I.download}</button>
      <button class="icon-btn notification-btn" data-action="notifications" title="Notifications">${I.bell}${unreadNoticeCount?`<span class="topbar-badge">${unreadNoticeCount > 99 ? '99+' : unreadNoticeCount}</span>`:''}</button>
      <div class="avatar" style="background:${me.color}" data-action="acct" title="Account">${esc(me.initials)}</div>
    </div>
  </div>`;
}

function globalSearchResults(query){
  const q = String(query || '').trim().toLowerCase();
  if(!q) return [];
  const out = [];
  DATA.boards.forEach(b=>{
    if(b.name.toLowerCase().includes(q)) out.push({type:'board', id:b.id, title:b.name, meta:'Board', icon:boardIconHtml(b)});
  });
  allVisibleTasks().forEach(t=>{
    const hay = [
      t.name,
      t.notes,
      t._board?.name,
      t._group?.name,
      taskOwnerNames(t),
      member(t.assignedBy)?.name
    ].filter(Boolean).join(' ').toLowerCase();
    if(hay.includes(q)) out.push({type:'task', id:t.id, boardId:t._board.id, title:t.name, meta:`${t._board.name} / ${t._group.name}`, icon:'TASK'});
  });
  (DATA.docs || []).forEach(d=>{
    const hay = `${d.title || ''} ${d.template || ''} ${d.content || ''}`.toLowerCase();
    if(hay.includes(q)) out.push({type:'doc', id:d.id, title:d.title || 'Untitled doc', meta:'Document', icon:'DOC'});
  });
  MEETING_ROWS.forEach(meeting=>{
    const loadedContent = MEETING_DETAIL?.id === meeting.id
      ? [
          MEETING_DETAIL.transcript,
          MEETING_DETAIL.sections?.summary,
          ...(MEETING_DETAIL.sections?.keyDiscussionPoints || []),
          ...(MEETING_DETAIL.sections?.decisions || []),
          ...(MEETING_DETAIL.sections?.openQuestions || []),
        ].join(' ')
      : '';
    const hay = [
      meeting.title,
      meeting.meetingDate,
      meeting.language,
      ...(meeting.attendees || []),
      loadedContent
    ].filter(Boolean).join(' ').toLowerCase();
    if(hay.includes(q)) out.push({
      type:'meeting',
      id:meeting.id,
      title:meeting.title || 'Untitled meeting',
      meta:`Meeting${meeting.meetingDate ? ` / ${meeting.meetingDate}` : ''}`,
      icon:'MTG'
    });
  });
  (DATA.assets || []).forEach(a=>{
    const hay = `${a.title || ''} ${a.description || ''} ${a.externalUrl || ''}`.toLowerCase();
    if(hay.includes(q)) out.push({type:'asset', id:a.id, title:a.title || 'Untitled asset', meta:a.kind === 'link' ? 'Link' : 'File / media', icon:assetIcon(a)});
  });
  DATA.team.forEach(m=>{
    const hay = `${m.name || ''} ${m.email || ''} ${m.initials || ''}`.toLowerCase();
    if(hay.includes(q)) out.push({type:'person', id:m.id, title:m.name || m.email || 'Team member', meta:m.email || `${memberRole(m.id)} member`, icon:m.initials || 'TM'});
  });
  const rank = {task:0, meeting:1, board:2, doc:3, asset:4, person:5};
  return out.sort((a,b)=>(rank[a.type] ?? 9)-(rank[b.type] ?? 9) || a.title.localeCompare(b.title)).slice(0,12);
}

function globalSearchMenuHtml(query){
  const results = globalSearchResults(query);
  if(!String(query || '').trim()) return '';
  if(!results.length) return `<div class="global-search-empty">No matches found</div>`;
  return results.map(r=>`
    <button class="global-search-result" data-global-result="${esc(r.type)}|${esc(r.id)}|${esc(r.boardId || '')}">
      <span class="global-result-icon">${Object.values(I).includes(r.icon) ? r.icon : esc(r.icon)}</span>
      <span><b>${esc(r.title)}</b><em>${esc(r.meta)}</em></span>
    </button>`).join('');
}

function updateGlobalSearchResults(){
  const root = $('#global-search-results');
  if(!root) return;
  root.innerHTML = globalSearchMenuHtml(GLOBAL_SEARCH);
  root.classList.toggle('open', Boolean(GLOBAL_SEARCH.trim()));
  wireGlobalSearchResults(root);
}

function wireGlobalSearchResults(root=document){
  root.querySelectorAll('[data-global-result]').forEach(btn=>btn.onclick=()=>{
    const [type,id,boardId] = btn.dataset.globalResult.split('|');
    GLOBAL_SEARCH = '';
    if(type === 'board') openBoard(id);
    else if(type === 'task') openTaskDetail(boardId,'',id,'updates');
    else if(type === 'meeting') openMeeting(id);
    else if(type === 'doc') openDoc(id);
    else if(type === 'asset') void openAsset(id);
    else if(type === 'person') navigateToView('settings');
    updateGlobalSearchResults();
  });
}

function renderContent({preserveScroll=true, scrollState=null} = {}){
  const state = scrollState || (preserveScroll ? captureScrollState() : null);
  const c = $('#content');
  if(!c) return;
  c.dataset.renderView = VIEW;
  c.dataset.renderBoard = CURRENT_BOARD || '';
  if(VIEW==='home') c.innerHTML = viewHome();
  else if(VIEW==='updates') c.innerHTML = viewUpdates();
  else if(VIEW==='board') c.innerHTML = viewBoard();
  else if(VIEW==='dashboard') c.innerHTML = viewDashboard();
  else if(VIEW==='calendar') c.innerHTML = viewCalendar();
  else if(VIEW==='closing') c.innerHTML = viewMonthlyClosing();
  else if(VIEW==='mywork') c.innerHTML = viewMyWork();
  else if(VIEW==='meetings') c.innerHTML = viewMeetings();
  else if(VIEW==='integrations') c.innerHTML = viewIntegrations();
  else if(VIEW==='docs') c.innerHTML = viewDocs();
  else if(VIEW==='settings') c.innerHTML = viewSettings();
  wireContent();
  restoreScrollState(state);
}

/* ============================================================
   FRIDAY INTELLIGENCE ASSISTANT
   ============================================================ */
function assistantPetSettings(value=DATA?.settings?.assistantPet){
  const raw = value && typeof value === 'object' ? value : {};
  const pet = {...ASSISTANT_PET_DEFAULTS, ...raw};
  pet.enabled = pet.enabled !== false;
  pet.name = String(pet.name || ASSISTANT_PET_DEFAULTS.name).trim().slice(0,32) || ASSISTANT_PET_DEFAULTS.name;
  pet.style = ASSISTANT_PET_STYLES.some(x=>x.id===pet.style) ? pet.style : ASSISTANT_PET_DEFAULTS.style;
  pet.color = /^#[0-9a-f]{6}$/i.test(String(pet.color||'')) ? pet.color : (ASSISTANT_PET_STYLE_COLORS[pet.style] || ASSISTANT_PET_DEFAULTS.color);
  pet.size = ['small','medium','large'].includes(pet.size) ? pet.size : ASSISTANT_PET_DEFAULTS.size;
  pet.position = ['bottom-right','bottom-left'].includes(pet.position) ? pet.position : ASSISTANT_PET_DEFAULTS.position;
  pet.motion = ['float','pulse','still'].includes(pet.motion) ? pet.motion : ASSISTANT_PET_DEFAULTS.motion;
  pet.autoBrief = pet.autoBrief === true;
  return pet;
}

function opsAssistantWidget(){
  const pet = assistantPetSettings();
  if(!pet.enabled) return '';
  const messages = ASSISTANT_MESSAGES.map((message,index)=>opsAssistantMessageHtml(message,index)).join('');
  const empty = !ASSISTANT_MESSAGES.length && !ASSISTANT_LOADING;
  return `
    <button class="ops-assistant-launcher pet-${pet.size} pos-${pet.position} ${ASSISTANT_OPEN?'open':''}" data-action="assistant-toggle" title="${esc(pet.name)} Intelligence">
      <span class="assistant-pet pet-${pet.style} pet-${pet.size} motion-${pet.motion}" style="--pet-color:${esc(pet.color)}">
        ${assistantPetFace(pet)}
        <span class="assistant-pet-label">${esc(pet.name)}</span>
      </span>
    </button>
    ${ASSISTANT_OPEN ? `<section class="ops-assistant-panel pos-${pet.position}" aria-label="${esc(pet.name)} Intelligence">
      <div class="ops-assistant-head">
        <button class="icon-btn" data-action="assistant-close" title="Close">${I.close}</button>
      </div>
      <div class="ops-assistant-body" id="ops-assistant-body">
        ${empty ? `<div class="ops-assistant-empty">
          <b>Ask about tasks, emails, files, or your calendar.</b>
        </div>` : messages}
        ${ASSISTANT_LOADING ? `<div class="ops-assistant-thinking">${I.spark}<span>Checking connected workspace data...</span></div>` : ''}
      </div>
      <div class="ops-assistant-composer">
        <textarea id="ops-assistant-input" rows="2" placeholder="Ask what needs attention, where a file is, or what happened with a stakeholder..."></textarea>
        <button class="btn sm" data-action="assistant-send">Send</button>
      </div>
    </section>` : ''}`;
}

function assistantPetFace(pet){
  return `<span class="pet-body" aria-hidden="true">
    <span class="pet-shadow"></span>
    <span class="pet-ear left"></span>
    <span class="pet-ear right"></span>
    <span class="pet-head"></span>
    <span class="pet-face"></span>
    <span class="pet-eye left"></span>
    <span class="pet-eye right"></span>
    <span class="pet-nose"></span>
    <span class="pet-smile"></span>
    <span class="pet-cheek left"></span>
    <span class="pet-cheek right"></span>
    <span class="pet-flipper left"></span>
    <span class="pet-flipper right"></span>
  </span>`;
}

function opsAssistantMessageHtml(message,index=0){
  const marker = ` data-ops-msg-index="${index}"`;
  if(message.role === 'user') return `<div class="ops-msg user"${marker}>${esc(message.text||'')}</div>`;
  if(message.result) return `<div class="ops-msg assistant"${marker}>${opsAssistantResultHtml(message.result)}</div>`;
  return `<div class="ops-msg assistant"${marker}>${formatAssistantText(message.text||'')}</div>`;
}

function opsAssistantResultHtml(result){
  const sources = Array.isArray(result.sources) ? result.sources.slice(0,6) : [];
  const actionLinks = sources.filter(s=>/^https:\/\/mail\.google\.com\/mail\//.test(String(s.url || '')));
  const pendingActions = Array.isArray(result.pendingActions) ? result.pendingActions.slice(0,6) : [];
  return `
    ${result.answer ? `<div class="ops-answer">${formatAssistantText(result.answer)}</div>` : ''}
    ${actionLinks.length ? `<div class="ops-action-links">${actionLinks.map(s=>`
      <a href="${esc(s.url||'#')}" target="_blank" rel="noreferrer">${esc(s.label||'Open draft')}</a>
      ${s.detail?`<span>${esc(s.detail)}</span>`:''}
    `).join('')}</div>` : ''}
    ${pendingActions.length ? `<div class="ops-pending-actions">${pendingActions.map(a=>assistantActionHtml(a)).join('')}</div>` : ''}
    ${result.recommendedNextAction ? `<div class="ops-next">${formatAssistantText(result.recommendedNextAction)}</div>` : ''}
  `;
}

function assistantActionHtml(action){
  const status = action.status || 'pending';
  if(status === 'done') return `<div class="ops-action-status done"><b>${esc(action.title||'Action completed')}</b><span>${esc(action.resultMessage||'Done')}</span></div>`;
  if(status === 'error') return `<div class="ops-action-status error"><b>${esc(action.title||'Action failed')}</b><span>${esc(action.error||'Could not complete this action.')}</span></div>`;
  if(status === 'running') return `<div class="ops-action-status"><b>${esc(action.title||'Running action')}</b><span>Working...</span></div>`;
  return `<div class="ops-action-card">
    <div><b>${esc(action.title||'Assistant action')}</b>${action.description?`<span>${esc(action.description)}</span>`:''}</div>
    <button class="btn sm" data-assistant-action="${esc(action.id||'')}">${esc(action.confirmLabel||'Review and run')}</button>
  </div>`;
}

function formatAssistantText(text){
  return esc(text).replace(/\n/g,'<br>');
}

function wireOpsAssistant(){
  document.querySelectorAll('[data-action="assistant-toggle"]').forEach(el=>el.onclick=()=>{
    ASSISTANT_OPEN = !ASSISTANT_OPEN;
    if(ASSISTANT_OPEN){
      ASSISTANT_STARTED = true;
      ASSISTANT_LAST_SCROLL_KEY = '';
    }
    render();
  });
  document.querySelectorAll('[data-action="assistant-close"]').forEach(el=>el.onclick=()=>{ ASSISTANT_OPEN=false; render(); });
  document.querySelectorAll('[data-action="assistant-send"]').forEach(el=>el.onclick=()=>sendOpsAssistantMessage());
  document.querySelectorAll('[data-assistant-action]').forEach(el=>el.onclick=()=>confirmAssistantAction(el.dataset.assistantAction));
  const input = $('#ops-assistant-input');
  if(input){
    input.onkeydown=e=>{
      if(e.key==='Enter' && !e.shiftKey){
        e.preventDefault();
        sendOpsAssistantMessage();
      }
    };
  }
  const body = $('#ops-assistant-body');
  if(body){
    const messageCount = ASSISTANT_MESSAGES.length;
    const scrollKey = assistantLastMessageScrollKey();
    if(messageCount && scrollKey !== ASSISTANT_LAST_SCROLL_KEY){
      const target = body.querySelector(`[data-ops-msg-index="${messageCount - 1}"]`);
      if(target){
        requestAnimationFrame(()=>scrollAssistantBodyToMessage(body, target));
      }
    }else if(!messageCount){
      body.scrollTop = 0;
    }
    ASSISTANT_LAST_SCROLL_KEY = scrollKey;
  }
}

function assistantLastMessageScrollKey(){
  const last = ASSISTANT_MESSAGES[ASSISTANT_MESSAGES.length - 1];
  if(!last) return 'empty';
  const answer = last.result?.answer || '';
  return `${ASSISTANT_MESSAGES.length}|${last.role || ''}|${last.text || ''}|${answer}`;
}

function scrollAssistantBodyToMessage(body,target){
  if(!body || !target) return;
  const bodyRect = body.getBoundingClientRect();
  const targetRect = target.getBoundingClientRect();
  const top = body.scrollTop + (targetRect.top - bodyRect.top) - 8;
  body.scrollTo({top:Math.max(0, top), behavior:'smooth'});
}

function sendOpsAssistantMessage(){
  const input = $('#ops-assistant-input');
  const text = input?.value.trim();
  if(!text || ASSISTANT_LOADING) return;
  if(input) input.value = '';
  void requestOpsAssistant(text, { showUser:true });
}

async function requestOpsAssistant(message,{showUser=true}={}){
if(!isAssignedTaskEmailRequest(message)){
    if(showUser)ASSISTANT_MESSAGES.push({role:'user',text:message});
    const open=allVisibleTasks().filter(t=>t.status!=='done');
    ASSISTANT_MESSAGES.push({role:'assistant',result:prepareAssistantResult({answer:'Local simulation: '+open.length+' open tasks are available in this sample. Review blocked tasks and their owners before changing a deadline. This is a fixed local response, not an AI-generated recommendation.',top3:[],sources:[],unknowns:['External calendars and messages are not read.'],pendingActions:[]})});ASSISTANT_OPEN=true;render();return;
  }

  if(isAssignedTaskEmailRequest(message)){
    if(showUser) ASSISTANT_MESSAGES.push({role:'user',text:message});
    ASSISTANT_MESSAGES.push({role:'assistant',result:prepareAssistantResult(buildAssignedTaskEmailDraftResult(message))});
    ASSISTANT_MESSAGES = ASSISTANT_MESSAGES.slice(-20);
    ASSISTANT_OPEN = true;
    render();
    return;
  }
  if(!supabase || !SESSION?.user){
    ASSISTANT_OPEN = true;
    ASSISTANT_MESSAGES.push({role:'assistant',text:'Sign in to TaskTrack before using the intelligence assistant.'});
    render();
    return;
  }
  if(showUser) ASSISTANT_MESSAGES.push({role:'user',text:message});
  ASSISTANT_LOADING = true;
  ASSISTANT_OPEN = true;
  render();

  const history = ASSISTANT_MESSAGES.map(m=>m.result
    ? { role:'assistant', content:m.result.answer || JSON.stringify(m.result.top3 || []) }
    : { role:m.role, content:m.text || '' });
  const { data, error } = await supabase.functions.invoke('ops-assistant', {
    body: { message, messages:history, projectSnapshot:buildOpsProjectSnapshot() },
  });

  ASSISTANT_LOADING = false;
  if(error || data?.error || !data?.result){
    ASSISTANT_MESSAGES.push({role:'assistant',text:await assistantFunctionErrorMessage(error, data) || 'I could not reach the assistant function yet. Make sure ops-assistant is deployed in Supabase.'});
  } else {
    ASSISTANT_MESSAGES.push({role:'assistant',result:prepareAssistantResult(data.result)});
  }
  ASSISTANT_MESSAGES = ASSISTANT_MESSAGES.slice(-20);
  render();
}

function prepareAssistantResult(result){
  const out = result && typeof result === 'object' ? result : { answer:String(result || '') };
  out.top3 = Array.isArray(out.top3) ? out.top3 : [];
  out.sources = Array.isArray(out.sources) ? out.sources : [];
  out.unknowns = Array.isArray(out.unknowns) ? out.unknowns : [];
  out.pendingActions = Array.isArray(out.pendingActions) ? out.pendingActions.map(action=>({
    ...action,
    id:action.id || uid('act'),
    payload:assistantActionPayload(action),
    status:action.status || 'pending'
  })) : [];
  return out;
}

function assistantActionPayload(action){
  if(action?.payload && typeof action.payload === 'object') return action.payload;
  if(typeof action?.payloadJson === 'string' && action.payloadJson.trim()){
    try{ return JSON.parse(action.payloadJson); }catch(_error){ return {}; }
  }
  return {};
}

function isAssignedTaskEmailRequest(message=''){
  const text = String(message || '').toLowerCase();
  return /(gmail|email|notify|notification|send)/.test(text)
    && /(assigned|assignee|owner|task|tasks)/.test(text)
    && /(send|draft|notify|notification|email|gmail)/.test(text);
}

function buildAssignedTaskEmailDraftResult(message=''){
  const openAssigned = allVisibleTasks().filter(t=>taskOwnerIds(t).length && t.status !== 'done');
  const grouped = new Map();
  const missingEmail = new Set();
  openAssigned.forEach(t=>{
    taskOwners(t).forEach(owner=>{
      if(!owner.email){
        missingEmail.add(owner.name || owner.id);
        return;
      }
      if(!grouped.has(owner.id)) grouped.set(owner.id, { owner, tasks:[] });
      grouped.get(owner.id).tasks.push(t);
    });
  });
  const groups = [...grouped.values()].sort((a,b)=>a.owner.name.localeCompare(b.owner.name));
  if(!groups.length){
    return {
      top3:[],
      answer:'I did not find any open assigned tasks with teammate email addresses. Add owner emails in Settings, then I can create Gmail draft notifications for them.',
      sources:[],
      unknowns:missingEmail.size ? [`Missing emails for: ${[...missingEmail].join(', ')}`] : [],
      pendingActions:[],
      recommendedNextAction:'Add teammate emails in Settings or assign tasks to teammates with emails.'
    };
  }
  const wantsSend = /\bsend|notify|notification|notifications\b/i.test(message);
  const actionPrefs = assistantActionPrefs();
  const messages = groups.map(({owner,tasks})=>({
    to:owner.email,
    subject:'Your open TaskTrack tasks',
    body:assignedTaskEmailBody(owner,tasks)
  }));
  const sources = groups.map(({owner,tasks})=>({
    label:`Draft Gmail to ${owner.name}`,
    url:gmailDraftUrl(owner,tasks),
    detail:`${tasks.length} open assigned task${tasks.length===1?'':'s'}`
  }));
  const pendingActions = [];
  if(wantsSend && actionPrefs.gmailSend){
    pendingActions.push({
      id:uid('act'),
      provider:'gmail',
      action:'gmail.send_emails',
      title:`Send ${groups.length} Gmail task notification${groups.length===1?'':'s'}`,
      description:'TaskTrack will send the task list emails through your connected Gmail account after you confirm.',
      confirmLabel:'Send emails',
      payload:{ messages }
    });
  } else if(actionPrefs.gmailDrafts){
    pendingActions.push({
      id:uid('act'),
      provider:'gmail',
      action:'gmail.create_drafts',
      title:`Create ${groups.length} Gmail draft${groups.length===1?'':'s'}`,
      description:'TaskTrack will create Gmail drafts through your connected Gmail account so you can review them before sending.',
      confirmLabel:'Create drafts',
      payload:{ messages }
    });
  }
  return {
    top3:groups.slice(0,3).map(({owner,tasks})=>({
      title:`${owner.name}: ${tasks.length} open assigned task${tasks.length===1?'':'s'}`,
      whyItMatters:tasks.slice(0,3).map(t=>`${t.name}${taskTimelineLabel(t) ? ` scheduled ${taskTimelineLabel(t)}` : ''}`).join('; '),
      recommendedAction:`Open the Gmail draft for ${owner.name}, review it, and send.`,
      sourceRefs:[`Draft Gmail to ${owner.name}`]
    })),
    answer:wantsSend && actionPrefs.gmailSend
      ? `I found ${groups.length} teammate${groups.length===1?'':'s'} with open assigned tasks. Confirm below and TaskTrack will send the Gmail notifications from your connected Gmail account.`
      : `I found ${groups.length} teammate${groups.length===1?'':'s'} with open assigned tasks. You can create Gmail drafts from TaskTrack, then review and send them manually.`,
    sources,
    unknowns:[
      ...(wantsSend && !actionPrefs.gmailSend ? ['Gmail send is disabled in Settings. An admin can enable it under Assistant action permissions.'] : []),
      ...(missingEmail.size ? [`Missing emails for: ${[...missingEmail].join(', ')}`] : [])
    ],
    pendingActions,
    recommendedNextAction:wantsSend && actionPrefs.gmailSend ? 'Review the action details, then confirm to send.' : 'Create the Gmail drafts below or use the manual compose links.'
  };
}

function gmailDraftUrl(owner,tasks){
  const subject = `Your open TaskTrack tasks`;
  const body = assignedTaskEmailBody(owner,tasks);
  const url = new URL('https://mail.google.com/mail/');
  url.searchParams.set('view','cm');
  url.searchParams.set('fs','1');
  url.searchParams.set('to',owner.email || '');
  url.searchParams.set('su',subject);
  url.searchParams.set('body',body);
  return url.toString();
}

function assignedTaskEmailBody(owner,tasks){
  const lines = [
    `Hi ${owner.name || 'there'},`,
    '',
    'Here are your open assigned tasks in TaskTrack:',
    '',
    ...tasks.slice(0,25).map((t,i)=>{
      const bits = [
        `${i+1}. ${t.name}`,
        `Board: ${t._board?.name || 'Unknown board'}`,
        `Group: ${t._group?.name || 'Unknown group'}`,
        `Status: ${statusOptionFor(t._board,t.status).label}`,
        taskTimelineLabel(t) ? `Timeline: ${taskTimelineLabel(t)}` : '',
        t.priority ? `Priority: ${priorityOptionFor(t._board,t.priority).label}` : '',
      ].filter(Boolean);
      return bits.join(' | ');
    }),
    tasks.length > 25 ? `\nAnd ${tasks.length - 25} more assigned tasks in TaskTrack.` : '',
    '',
    'Please update the status or notes in TaskTrack when you have a chance.',
    '',
    'Thanks,',
    'TaskTrack'
  ];
  return lines.filter(line=>line !== '').join('\n');
}

function findAssistantAction(actionId){
  for(const message of ASSISTANT_MESSAGES){
    const actions = message?.result?.pendingActions;
    if(!Array.isArray(actions)) continue;
    const action = actions.find(a=>a.id === actionId);
    if(action) return { message, action };
  }
  return null;
}

function confirmAssistantAction(actionId){
  const found = findAssistantAction(actionId);
  if(!found || ASSISTANT_ACTION_BUSY) return;
  const { action } = found;
  const detail = action.description || 'TaskTrack will run this action through the connected integration account.';
  modal('Confirm assistant action', detail,
    `<div class="integration-note">Action: ${esc(action.title || action.action || 'Assistant action')}</div>`,
    ()=>executeAssistantAction(actionId),
    action.confirmLabel || 'Run action');
}

async function executeAssistantAction(actionId){
const sampleFound=findAssistantAction(actionId);if(!sampleFound||ASSISTANT_ACTION_BUSY)return;
  const sampleAction=sampleFound.action;closeModal();sampleAction.status='done';sampleAction.resultMessage='Simulation completed locally. No email, calendar event, or document was sent to an external service.';
  recordPortfolioSimulation(sampleAction.action,sampleAction.payload);ASSISTANT_MESSAGES.push({role:'assistant',text:sampleAction.resultMessage});render();return;

  const found = findAssistantAction(actionId);
  if(!found || ASSISTANT_ACTION_BUSY) return;
  const { action } = found;
  closeModal();

  if(!supabase || !SESSION?.user || !WORKSPACE_ID){
    action.status = 'error';
    action.error = 'Sign in to the shared workspace before running assistant actions.';
    render();
    return;
  }

  ASSISTANT_ACTION_BUSY = true;
  action.status = 'running';
  render();
  const { data, error } = await supabase.functions.invoke('assistant-actions', {
    body:{
      workspaceId:WORKSPACE_ID,
      provider:action.provider,
      action:action.action,
      payload:action.payload || {},
      confirmed:true
    }
  });
  ASSISTANT_ACTION_BUSY = false;

  if(error || data?.error){
    action.status = 'error';
    action.error = await assistantFunctionErrorMessage(error, data);
    render();
    return;
  }

  action.status = 'done';
  action.resultMessage = assistantActionResultMessage(action, data?.result);
  if(action.provider === 'notion' && action.action === 'notion.import_preview' && data?.result?.preview){
    NOTION_AI_PREVIEW = data.result.preview;
    render();
    showNotionAiPreview(data.result.preview, action.payload?.prompt || action.description || 'Assistant Notion import');
    return;
  }
  if(action.provider === 'google_drive' && action.action === 'drive.import_files'){
    void fetchWorkspaceState();
  }
  ASSISTANT_MESSAGES.push({role:'assistant',text:action.resultMessage});
  ASSISTANT_MESSAGES = ASSISTANT_MESSAGES.slice(-20);
  render();
}

async function assistantFunctionErrorMessage(error,data){
  if(data?.error) return data.error;
  const response = error?.context;
  if(response?.clone){
    try{
      const body = await response.clone().json();
      if(body?.error) return body.error;
      if(body?.message) return body.message;
    }catch(e){}
  }
  return error?.message || 'The action failed.';
}

function assistantActionResultMessage(action,result){
  if(action.action === 'gmail.create_drafts') return `Created ${Number(result?.count || 0)} Gmail draft${Number(result?.count || 0)===1?'':'s'}.`;
  if(action.action === 'gmail.send_emails') return `Sent ${Number(result?.count || 0)} Gmail email${Number(result?.count || 0)===1?'':'s'}.`;
  if(action.action === 'calendar.create_event') return `Created the calendar event${result?.url ? `: ${result.url}` : '.'}`;
  if(action.action === 'drive.create_doc') return `Created the Drive document${result?.account ? ` using ${result.account}` : ''}${result?.url ? `: ${result.url}` : '.'}`;
  if(action.action === 'drive.search_files'){
    const files = Array.isArray(result?.files) ? result.files : [];
    const account = result?.account ? ` in ${result.account}` : '';
    if(!files.length) return `No matching Drive files were found${account}.`;
    return `Found ${files.length} Drive file${files.length===1?'':'s'}${account}:\n${files.slice(0,5).map(f=>`- ${f.name}${f.url ? `: ${f.url}` : ''}`).join('\n')}`;
  }
  if(action.action === 'drive.import_files'){
    const count = Number(result?.count || 0);
    const account = result?.account ? ` from ${result.account}` : '';
    if(!count) return `No Drive files were imported${account}.`;
    return `Imported ${count} Drive file${count===1?'':'s'}${account} into Docs Hub.`;
  }
  if(action.action === 'drive.rename_file') return `Renamed Drive file${result?.account ? ` using ${result.account}` : ''}${result?.name ? `: ${result.name}` : '.'}`;
  if(action.action === 'drive.move_file') return `Moved Drive file${result?.account ? ` using ${result.account}` : ''}${result?.name ? `: ${result.name}` : '.'}`;
  if(action.action === 'drive.trash_file') return `Moved Drive file to trash${result?.account ? ` using ${result.account}` : ''}${result?.name ? `: ${result.name}` : '.'}`;
  if(action.action === 'drive.share_file') return `Shared Drive file${result?.sharedWith ? ` with ${result.sharedWith}` : ''}${result?.role ? ` as ${result.role}` : ''}.`;
  if(action.action === 'drive.edit_doc') return `Updated Google Doc${result?.account ? ` using ${result.account}` : ''}${result?.url ? `: ${result.url}` : '.'}`;
  if(action.action === 'notion.import_preview') return 'Created the Notion import preview.';
  return 'Assistant action completed.';
}

function buildOpsProjectSnapshot(){
  const now = new Date().toISOString();
  return {
    generatedAt: now,
    workspaceId: WORKSPACE_ID || '',
    workspaceName: DATA.settings?.account || 'TaskTrack',
    currentMember: member(currentMemberId()) || null,
    team: (DATA.team||[]).map(m=>({id:m.id,name:m.name,email:m.email||''})),
    assistantActionPermissions:assistantActionPrefs(),
    integrations: INTEGRATION_CONNECTIONS.map(c=>({provider:c.provider,account:c.account_label||c.external_account_id||'',connectedAt:c.connected_at||''})),
    boards: DATA.boards.map(b=>({
      id:b.id,
      name:b.name,
      groups:b.groups.map(g=>({
        id:g.id,
        name:g.name,
        tasks:g.tasks.flatMap(t=>opsTaskSnapshot(t,b,g.name,0)).slice(0,80),
      })),
    })).slice(0,20),
    recentUpdates: (DATA.updates||[]).slice(0,20).map(u=>({
      body:u.body,
      createdAt:u.createdAt,
      board:board(u.boardId)?.name||'',
      taskId:u.taskId||'',
      mentions:u.mentions||[],
    })),
    docs: (DATA.docs||[]).slice(0,20).map(d=>({title:d.title,template:d.template||'',updated:d.updated||0})),
  };
}

function opsTaskSnapshot(t,b,groupName,depth){
  const assigner = member(t.assignedBy);
  const row = {
    id:t.id,
    name:t.name,
    group:groupName,
    depth,
    owner:taskOwnerNames(t),
    assignedBy:assigner?.name||'',
    status:statusOptionFor(b,t.status).label,
    priority:priorityOptionFor(b,t.priority).label,
    due:t.due||'',
    timeline:[t.tlStart,t.tlEnd].filter(Boolean).join(' - '),
    budget:Number(t.budget)||0,
    notes:t.notes||'',
    overdue:isOverdueTask(t),
  };
  return [row, ...(t.subtasks||[]).flatMap(st=>opsTaskSnapshot(st,b,groupName,depth+1))];
}

/* ============================================================
   VIEW: Workspace Home
   ============================================================ */
function viewHome(){
  const rows = DATA.boards.map(b=>{
    const tasks = boardTaskCount(b);
    const done = boardDoneCount(b);
    return `<div class="asset-row" data-open-board="${b.id}">
      <div class="asset-name"><span class="asset-ic">${boardIconHtml(b)}</span>${esc(b.name)}</div>
      <div style="color:var(--ink2)">${tasks} tasks</div>
      <div style="color:var(--ink2)">${done} done</div>
      <div style="color:var(--ink3)">You</div>
      <div style="color:var(--ink3)">${esc(fmtTimeAgo(boardUpdatedAt(b)))}</div>
    </div>`;
  }).join('');
  return `
    ${gettingStartedPanel()}
    <div class="home-hero">
      <div class="home-logo"><img src="./favicon.png" alt="TaskTrack"></div>
      <div>
        <h2>TaskTrack Workspace</h2>
        <div class="desc">All your task boards, dashboards and docs in one place. Board totals include subtasks.</div>
      </div>
      <div style="margin-left:auto"><button class="btn" data-action="add-board"><span>${I.plus}</span> New board</button></div>
    </div>
    <div class="asset-tbl">
      <div class="asset-row head">
        <div>Board name</div><div>Items</div><div>Completed</div><div>Owner</div><div>Last modified</div>
      </div>
      ${rows||'<div class="empty-state"><div class="big">📋</div>No boards yet — create your first task board.</div>'}
    </div>`;
}

function gettingStartedPanel(){
  if(DATA.settings?.onboardingDismissed) return '';
  const firstBoardId = DATA.boards[0]?.id || '';
  const steps = [
    {done:DATA.boards.length > 0, label:'Create or open a board', action:'add-board', button:'New board'},
    {done:allVisibleTasks().length > 0, label:'Add real tasks and owners', boardId:firstBoardId, button:'Open board'},
    {done:DATA.team.length > 1, label:'Invite your team', view:'settings', button:'Team settings'},
    {done:((DATA.docs || []).length + (DATA.assets || []).length) > 0, label:'Attach files or create docs', view:'docs', button:'Docs Hub'}
  ];
  return `<div class="getting-started">
    <div>
      <h2>Get TaskTrack ready for your team</h2>
      <p>Create a board, add real work, invite teammates, and attach files when you are ready.</p>
    </div>
    <div class="getting-steps">
      ${steps.map(s=>`<div class="getting-step ${s.done?'done':''}">
        <span>${s.done?'✓':'•'}</span><b>${esc(s.label)}</b>
        ${s.action?`<button class="btn ghost sm" data-action="${s.action}">${esc(s.button)}</button>`:''}
        ${s.boardId?`<button class="btn ghost sm" data-open-board="${s.boardId}">${esc(s.button)}</button>`:''}
        ${s.view?`<button class="btn ghost sm" data-go-view="${s.view}">${esc(s.button)}</button>`:''}
      </div>`).join('')}
    </div>
    <button class="icon-btn getting-dismiss" data-action="dismiss-onboarding" title="Hide checklist">${I.close}</button>
  </div>`;
}

function guidedTourSteps(){
  return [
    {title:'Welcome to TaskTrack', body:'TaskTrack keeps boards, tasks, updates, and docs in one shared workspace for your team.'},
    {title:'Start with boards', body:'Create a board for each active workflow. Use table view for details and Kanban view when you want to move work by status.'},
    {title:'Assign real ownership', body:'Add your team in Settings, then assign owners and timelines so everyone knows what needs attention.'},
    {title:'Use Docs Hub', body:'Create shared notes, upload files, and link docs or media back to the tasks they support.'},
    {title:'Tune notifications', body:'Choose whether your feed shows mentions, @Everyone posts, assigned tasks, followed tasks, and your own updates.'}
  ];
}

function maybeShowGuidedTour(){
  if(TOUR_SHOWN_THIS_SESSION || DATA.settings?.guidedTourDone || CLOUD_LOADING || $('#modal-root')?.innerHTML) return;
  TOUR_SHOWN_THIS_SESSION = true;
  openGuidedTour(0);
}

function openGuidedTour(step=0){
  const steps = guidedTourSteps();
  const safeStep = Math.min(Math.max(Number(step)||0, 0), steps.length - 1);
  const item = steps[safeStep];
  const root = $('#modal-root');
  if(!root) return;
  root.innerHTML = `<div class="overlay"><div class="modal tour-modal">
    <div class="tour-mark"><img src="./favicon.png" alt="TaskTrack"></div>
    <h3>${esc(item.title)}</h3>
    <div class="msub">${esc(item.body)}</div>
    <div class="tour-progress">
      ${steps.map((_,i)=>`<span class="${i===safeStep?'active':''}"></span>`).join('')}
    </div>
    <div class="modal-actions">
      <button class="btn ghost" data-tour-skip>Skip</button>
      ${safeStep>0?`<button class="btn ghost" data-tour-prev>Back</button>`:''}
      <button class="btn" data-tour-next>${safeStep===steps.length-1?'Finish':'Next'}</button>
    </div>
  </div></div>`;
  root.querySelector('[data-tour-skip]').onclick=finishGuidedTour;
  root.querySelector('[data-tour-next]').onclick=()=> safeStep===steps.length-1 ? finishGuidedTour() : openGuidedTour(safeStep + 1);
  root.querySelector('[data-tour-prev]')?.addEventListener('click',()=>openGuidedTour(safeStep - 1));
}

function finishGuidedTour(){
  DATA.settings.guidedTourDone = true;
  save();
  closeModal();
}

/* ============================================================
   VIEW: Update feed
   ============================================================ */
function viewUpdates(){
  if(!UPDATE_FILTERS.has(UPDATE_FILTER)) UPDATE_FILTER = 'all';
  if(FEED_COMPOSER_BOARD && !DATA.boards.some(b=>b.id===FEED_COMPOSER_BOARD)){
    FEED_COMPOSER_BOARD = '';
    FEED_COMPOSER_TASK = '';
  }
  const meId = currentMemberId();
  const updates = filteredUpdates();
  const all = relevantFeedUpdates();
  const mentionCount = all.filter(u=>updateMentionsMember(u, meId)).length;
  const noBoardCount = all.filter(u=>!u.boardId).length;
  const desktopActive = desktopNotificationsEnabled();
  const desktopLabel = desktopActive ? 'Desktop alerts on' : 'Enable desktop alerts';
  const boardRows = DATA.boards.map(b=>{
    const count = all.filter(u=>u.boardId===b.id).length;
    return `<button class="feed-board-filter ${UPDATE_BOARD_FILTER===b.id?'active':''}" data-feed-board="${b.id}">
      <span>${esc(b.name)}</span><span>${count}</span>
    </button>`;
  }).join('');
  return `
    <div class="feed-shell">
      <aside class="feed-side">
        <div class="feed-side-title">
          <div><h3>Update feed</h3></div>
        </div>
        <div class="feed-filter-head"><b>Filter by Board</b></div>
        <button class="feed-board-filter ${UPDATE_BOARD_FILTER==='all'?'active':''}" data-feed-board="all">
          <span>All boards in my feed</span><span>${all.length}</span>
        </button>
        <button class="feed-board-filter ${UPDATE_BOARD_FILTER==='none'?'active':''}" data-feed-board="none">
          <span>Updates without boards</span><span>${noBoardCount}</span>
        </button>
        ${boardRows}
      </aside>
      <main class="feed-main">
        <div class="feed-tabs">
          <button class="${UPDATE_FILTER==='all'?'active':''}" data-update-filter="all">All updates</button>
          <button class="${UPDATE_FILTER==='mentions'?'active':''}" data-update-filter="mentions">${I.mention} I was mentioned <span>${mentionCount}</span></button>
        </div>
        <div class="feed-controls">
          <div class="feed-control-left">
            <button class="feed-show ${UPDATE_SHOW_UNREAD?'active':''}" data-action="toggle-unread-feed">Show <b>${UPDATE_SHOW_UNREAD?'Unread only':'Unread updates'}</b></button>
            ${desktopNotificationsSupported()?`<button class="feed-desktop ${desktopActive?'active':''}" data-action="request-desktop-notifications">${I.bell} ${desktopLabel}</button>`:''}
          </div>
          <button class="feed-mark-read" data-action="mark-updates-read">Mark all as read</button>
        </div>
        ${feedComposer()}
        <div class="feed-list">
          ${updates.map(renderUpdateCard).join('') || `<div class="feed-empty">
            <h3>No updates here yet</h3>
            <p>Your feed shows tasks you own or follow, plus direct mentions and @Everyone updates.</p>
          </div>`}
        </div>
      </main>
    </div>`;
}

function feedTaskOptions(boardId='', selected=''){
  const tasks = allVisibleTasks().filter(t=>!boardId || t._board.id === boardId);
  return tasks.map(t=>{
    const value = `${t._board.id}|${t._group.id}|${t.id}`;
    const label = boardId ? t.name : `${t._board.name} / ${t.name}`;
    return `<option value="${esc(value)}" ${selected===value?'selected':''}>${esc(label)}</option>`;
  }).join('');
}

function feedComposer(){
  const selectedBoard = DATA.boards.some(b=>b.id === FEED_COMPOSER_BOARD) ? FEED_COMPOSER_BOARD : '';
  const taskOpts = feedTaskOptions(selectedBoard, FEED_COMPOSER_TASK);
  const taskStillValid = !FEED_COMPOSER_TASK || taskOpts.includes(`value="${esc(FEED_COMPOSER_TASK)}"`);
  if(!taskStillValid) FEED_COMPOSER_TASK = '';
  const boardOpts = DATA.boards.map(b=>`<option value="${esc(b.id)}" ${selectedBoard===b.id?'selected':''}>${esc(b.name)}</option>`).join('');
  const teamMentions = DATA.team.map(m=>`<button type="button" data-mention-insert="${m.id}">@${esc(m.name)}</button>`).join('');
  const mentions = `<button type="button" data-mention-insert="${EVERYONE_MENTION_ID}">@Everyone</button>${teamMentions}`;
  return `<div class="feed-composer">
    <div class="feed-composer-row">
      <textarea id="feed-body" rows="3" placeholder="Write an update. Use @Name or @Everyone to notify the right people."></textarea>
      <button class="btn" data-action="post-update">Post update</button>
    </div>
    <div class="mention-suggestions" id="mention-suggestions">${mentions}</div>
    <div class="feed-file-drop" id="feed-file-drop" role="button" tabindex="0">Drop files here or click to attach<input id="feed-file-input" type="file" multiple hidden></div>
    <div class="feed-file-list" id="feed-file-list">${feedFileListHtml()}</div>
    <div class="feed-composer-meta">
      <label>Board <select id="feed-board"><option value="">No board</option>${boardOpts}</select></label>
      <label>Task <select id="feed-task"><option value="">${selectedBoard?'No linked task':'All boards'}</option>${taskOpts}</select></label>
    </div>
  </div>`;
}

function feedFileListHtml(){
  return FEED_FILES.map((file,index)=>`<span class="feed-file-chip">📎 ${esc(file.name)} <button type="button" data-remove-feed-file="${index}" aria-label="Remove ${esc(file.name)}">×</button></span>`).join('');
}
function addFeedFiles(files){
  const existing = new Set(FEED_FILES.map(file=>`${file.name}:${file.size}:${file.lastModified}`));
  for(const file of files){
    const key = `${file.name}:${file.size}:${file.lastModified}`;
    if(!existing.has(key)){ FEED_FILES.push(file); existing.add(key); }
  }
  const list = $('#feed-file-list');
  if(list) list.innerHTML = feedFileListHtml();
  wireFeedFileRemoval();
}
function wireFeedFileRemoval(){
  document.querySelectorAll('[data-remove-feed-file]').forEach(el=>el.onclick=()=>{
    FEED_FILES.splice(Number(el.dataset.removeFeedFile),1);
    const list = $('#feed-file-list');
    if(list) list.innerHTML = feedFileListHtml();
    wireFeedFileRemoval();
  });
}

function isVisibleUpdate(u){
  if(!u) return false;
  if(u.visible === false || u.kind === 'audit') return false;
  const body = String(u.body || '').trim();
  if(/^Created task\s+["“]/i.test(body)) return false;
  if(isAssignmentUpdateBody(body)) return false;
  return true;
}

function isAssignmentUpdateBody(body){
  const text = String(body || '').trim().replace(/[\u201c\u201d]/g,'"');
  if(/^Assigned\s+"?.+"?\s+to\s+.+\.?$/i.test(text)) return true;
  return /^Assigned\s+["â€œ].+["â€]\s+to\s+.+\.?$/i.test(String(body || '').trim());
}

function feedUpdates(){
  return mergeSharedCalendarFeed((DATA.updates||[]).filter(isVisibleUpdate),
    SHARED_GOOGLE_EVENTS,SHARED_GOOGLE_READS,currentMemberId());
}

function taskSubscribers(t){
  const raw = t?.custom?.[TASK_SUBSCRIBERS_KEY] ?? t?.subscribers ?? [];
  if(Array.isArray(raw)) return raw.filter(Boolean);
  if(typeof raw === 'string'){
    const value = raw.trim();
    if(!value) return [];
    try{
      const parsed = JSON.parse(value);
      if(Array.isArray(parsed)) return parsed.filter(Boolean);
    }catch(e){}
    return value.split(',').map(x=>x.trim()).filter(Boolean);
  }
  return [];
}

function setTaskSubscribers(t, ids){
  if(!t) return;
  if(!t.custom || typeof t.custom !== 'object') t.custom = {};
  t.custom[TASK_SUBSCRIBERS_KEY] = [...new Set((ids || []).filter(Boolean))];
}

function visibleTaskCustom(t){
  return Object.fromEntries(Object.entries(t?.custom || {}).filter(([key])=>!String(key).startsWith('__')));
}
function parseIdList(value){
  if(Array.isArray(value)) return value.map(String).filter(Boolean);
  if(value == null || value === '') return [];
  if(typeof value === 'string'){
    const trimmed = value.trim();
    if(!trimmed) return [];
    try{
      const parsed = JSON.parse(trimmed);
      if(Array.isArray(parsed)) return parsed.map(String).filter(Boolean);
    }catch(e){}
    return trimmed.split(',').map(x=>x.trim()).filter(Boolean);
  }
  return [String(value)].filter(Boolean);
}
function validMemberIds(ids){
  const teamIds = new Set((DATA.team || []).map(m=>m.id));
  return [...new Set((ids || []).filter(Boolean))].filter(id=>teamIds.has(id));
}
function taskOwnerIds(t){
  const raw = parseIdList(t?.custom?.[TASK_OWNERS_KEY] ?? t?.owners ?? []);
  return validMemberIds([t?.owner, ...raw]);
}
function setTaskOwnerIds(t, ids){
  if(!t) return;
  if(!t.custom || typeof t.custom !== 'object') t.custom = {};
  const next = validMemberIds(ids);
  t.owner = next[0] || '';
  t.custom[TASK_OWNERS_KEY] = next;
}
function taskOwners(t){ return taskOwnerIds(t).map(id=>member(id)).filter(Boolean); }
function taskOwnerNames(t){ return taskOwners(t).map(m=>m.name).join(', '); }
function normalizeHexColor(value){
  const raw = String(value || '').trim();
  return /^#[0-9a-f]{6}$/i.test(raw) ? raw.toLowerCase() : '';
}
function taskColor(t){
  return normalizeHexColor(t?.custom?.[TASK_COLOR_KEY] || t?.color || '');
}
function setTaskColor(t, color){
  if(!t) return;
  if(!t.custom || typeof t.custom !== 'object') t.custom = {};
  const next = normalizeHexColor(color);
  if(next) t.custom[TASK_COLOR_KEY] = next;
  else delete t.custom[TASK_COLOR_KEY];
  delete t.color;
}
function taskColorLabel(t){
  const color = taskColor(t);
  if(!color) return 'No color';
  const names = {
    '#0073ea':'Blue',
    '#00c875':'Green',
    '#fdab3d':'Orange',
    '#e2445c':'Red',
    '#a25ddc':'Purple',
    '#579bfc':'Sky',
    '#401694':'Indigo',
    '#00c2ff':'Cyan',
    '#ff5ac4':'Pink',
    '#cab641':'Gold',
    '#037f4c':'Forest',
    '#784bd1':'Violet'
  };
  return names[color] || color.toUpperCase();
}
function ownerCalendarColor(t){
  const owner = taskOwners(t)[0];
  return normalizeHexColor(owner?.color) || '#8a94a6';
}
function calendarSyncOwnerKey(owner){
  if(!owner) return '';
  return String(owner.id || normalizeEmail(owner.email) || owner.name || '').trim();
}
function taskGoogleCalendarEvents(t){
  const raw = t?.custom?.[TASK_GOOGLE_CALENDAR_EVENTS_KEY];
  return raw && typeof raw === 'object' && !Array.isArray(raw) ? raw : {};
}
function taskGoogleCalendarEventForOwner(t, owner){
  const key = calendarSyncOwnerKey(owner);
  const map = taskGoogleCalendarEvents(t);
  const direct = key ? map[key] : null;
  if(direct && typeof direct === 'object') return direct;
  const emailKey = normalizeEmail(owner?.email);
  const byEmail = emailKey ? Object.values(map).find(entry=>normalizeEmail(entry?.email) === emailKey) : null;
  if(byEmail && typeof byEmail === 'object') return byEmail;
  const currentEmail = normalizeEmail(SESSION?.user?.email);
  if(owner?.id === currentMemberId() || (emailKey && emailKey === currentEmail)){
    return {
      id:t?.custom?.[TASK_GOOGLE_CALENDAR_EVENT_KEY] || '',
      url:t?.custom?.[TASK_GOOGLE_CALENDAR_LINK_KEY] || '',
      syncedAt:t?.custom?.[TASK_GOOGLE_CALENDAR_SYNCED_KEY] || '',
      email:emailKey,
      name:owner?.name || ''
    };
  }
  return {};
}
function setTaskGoogleCalendarEventForOwner(t, owner, result={}){
  if(!t || !owner) return;
  if(!t.custom || typeof t.custom !== 'object') t.custom = {};
  const key = calendarSyncOwnerKey(owner);
  if(!key) return;
  const previous = taskGoogleCalendarEventForOwner(t, owner);
  const next = {
    id:result.id || previous.id || '',
    url:result.url || previous.url || '',
    syncedAt:new Date().toISOString(),
    email:normalizeEmail(owner.email),
    name:owner.name || ''
  };
  t.custom[TASK_GOOGLE_CALENDAR_EVENTS_KEY] = {
    ...taskGoogleCalendarEvents(t),
    [key]:next
  };
  const currentEmail = normalizeEmail(SESSION?.user?.email);
  if(owner.id === currentMemberId() || (next.email && next.email === currentEmail)){
    t.custom[TASK_GOOGLE_CALENDAR_EVENT_KEY] = next.id;
    t.custom[TASK_GOOGLE_CALENDAR_LINK_KEY] = next.url;
    t.custom[TASK_GOOGLE_CALENDAR_SYNCED_KEY] = next.syncedAt;
  }
}
function calendarSyncOwners(t){
  const owners = taskOwners(t).filter(owner=>normalizeEmail(owner.email));
  const seen = new Set();
  return owners.filter(owner=>{
    const key = calendarSyncOwnerKey(owner);
    if(!key || seen.has(key)) return false;
    seen.add(key);
    return true;
  });
}
function calendarSyncCurrentOwner(){
  const me = member(currentMemberId()) || null;
  if(me?.id) return me;
  const email = sessionEmail();
  return email ? { id:currentMemberId() || SESSION?.user?.id || email, name:displayNameFromEmail(email), email } : null;
}
function taskDescription(t){ return String(t?.custom?.[TASK_DESCRIPTION_KEY] || t?.description || ''); }
function setTaskDescription(t, value){
  if(!t) return;
  if(!t.custom || typeof t.custom !== 'object') t.custom = {};
  t.custom[TASK_DESCRIPTION_KEY] = String(value || '');
}
function localDateTimeValue(value){
  const d = value instanceof Date ? value : new Date(value);
  if(Number.isNaN(d.getTime())) return '';
  const pad = n => String(n).padStart(2,'0');
  return `${d.getFullYear()}-${pad(d.getMonth()+1)}-${pad(d.getDate())}T${pad(d.getHours())}:${pad(d.getMinutes())}`;
}
function normalizeRecurrenceStart(value){
  const raw = String(value || '').trim();
  if(!raw) return '';
  const d = new Date(raw);
  return Number.isNaN(d.getTime()) ? '' : d.toISOString();
}
function splitRecurrenceValue(value){
  if(value && typeof value === 'object'){
    const base = value.schedule || value.frequency || value.recurrence || value;
    const startsAt = normalizeRecurrenceStart(value.startsAt || value.startAt || value.nextResetAt || value.nextAt || value.at);
    return { base, startsAt };
  }
  const raw = String(value || '').trim();
  const idx = raw.indexOf('@');
  if(idx < 0) return { base:raw, startsAt:'' };
  return { base:raw.slice(0,idx), startsAt:normalizeRecurrenceStart(raw.slice(idx+1)) };
}
function normalizeRecurrenceSchedule(value){
  if(value === true || value === 1 || value === '1' || value === 'true') return 'daily';
  if(value && typeof value === 'object'){
    const base = value.schedule || value.frequency || value.recurrence;
    if(base) return normalizeRecurrenceSchedule(base);
    return customRecurrenceValue(value.interval, value.unit);
  }
  const {base} = splitRecurrenceValue(value);
  const key = String(base || '').trim().toLowerCase();
  const custom = parseCustomRecurrence(key);
  if(custom) return customRecurrenceValue(custom.interval, custom.unit);
  return RECURRENCE_OPTIONS[key] ? key : 'off';
}
function normalizeRecurrenceValue(value){
  const {startsAt} = splitRecurrenceValue(value);
  const key = normalizeRecurrenceSchedule(value);
  if(key === 'off') return 'off';
  return startsAt ? `${key}@${startsAt}` : key;
}
function parseCustomRecurrence(value){
  const text = String(value || '').trim().toLowerCase();
  let match = text.match(/^(?:every:|custom:)(\d+):(day|days|week|weeks|month|months)$/);
  if(!match) match = text.match(/^(?:every\s*)?(\d+)\s*(day|days|week|weeks|month|months)$/);
  if(!match) return null;
  return { interval:Number(match[1]), unit:normalizeRecurrenceUnit(match[2]) };
}
function normalizeRecurrenceUnit(unit){
  const key = String(unit || '').trim().toLowerCase();
  if(key.startsWith('week')) return 'weeks';
  if(key.startsWith('month')) return 'months';
  return 'days';
}
function customRecurrenceValue(interval,unit){
  const n = Math.max(1, Math.min(365, Math.round(Number(interval) || 1)));
  return `every:${n}:${normalizeRecurrenceUnit(unit)}`;
}
function recurrenceValueWithStart(schedule, startsAt){
  const key = normalizeRecurrenceSchedule(schedule);
  if(key === 'off') return 'off';
  const start = normalizeRecurrenceStart(startsAt);
  return start ? `${key}@${start}` : key;
}
function defaultRecurrenceStart(){
  const d = new Date();
  d.setHours(d.getHours() + 1, 0, 0, 0);
  return localDateTimeValue(d);
}
function formatRecurrenceStart(value){
  const start = normalizeRecurrenceStart(value);
  if(!start) return '';
  const d = new Date(start);
  if(Number.isNaN(d.getTime())) return '';
  const opts = {month:'short', day:'numeric', hour:'numeric', minute:'2-digit'};
  if(d.getFullYear() !== new Date().getFullYear()) opts.year = 'numeric';
  return d.toLocaleString(undefined, opts);
}
function recurrenceMeta(value){
  const normalized = normalizeRecurrenceValue(value);
  const {base, startsAt} = splitRecurrenceValue(normalized);
  const key = normalizeRecurrenceSchedule(base);
  const custom = parseCustomRecurrence(key);
  if(custom){
    const unit = RECURRENCE_UNITS[custom.unit] || RECURRENCE_UNITS.days;
    const unitLabel = custom.interval === 1 ? unit.singular : unit.label;
    const label = `Every ${custom.interval} ${unitLabel}`;
    const startLabel = formatRecurrenceStart(startsAt);
    return {
      key,
      value:normalized,
      startsAt,
      label:startLabel ? `${label}, ${startLabel}` : label,
      color:'#0073ea',
      interval:custom.interval,
      unit:custom.unit,
      custom:true
    };
  }
  const opt = RECURRENCE_OPTIONS[key] || RECURRENCE_OPTIONS.off;
  const fixed = {
    off:{interval:0,unit:'days'},
    daily:{interval:1,unit:'days'},
    weekly:{interval:1,unit:'weeks'},
    monthly:{interval:1,unit:'months'}
  }[key] || {interval:0,unit:'days'};
  const startLabel = formatRecurrenceStart(startsAt);
  return {
    key,
    value:normalized,
    startsAt,
    ...opt,
    label:startLabel && key !== 'off' ? `${opt.label}, ${startLabel}` : opt.label,
    ...fixed,
    custom:false
  };
}
function recurrenceOptionFor(value){
  return recurrenceMeta(value);
}
function recurrenceLabel(value){ return recurrenceOptionFor(value).label; }
function taskColumnRecurrence(t,b){
  if(!b?.columns?.length) return 'off';
  for(const c of b.columns){
    if(c.type !== 'recurrence') continue;
    const value = normalizeRecurrenceValue(t?.custom?.[c.id]);
    if(value !== 'off') return value;
  }
  return 'off';
}
function taskRecurrence(t,b=null){
  const internal = normalizeRecurrenceValue(t?.custom?.[TASK_RECURRENCE_KEY]);
  if(internal !== 'off') return internal;
  const fromColumn = taskColumnRecurrence(t,b);
  if(fromColumn !== 'off') return fromColumn;
  const legacy = normalizeRecurrenceValue(t?.custom?.[TASK_RECURRING_KEY] ?? t?.recurring ?? false);
  return legacy;
}
function isRecurringTask(t,b=null){ return taskRecurrence(t,b) !== 'off'; }
function setTaskRecurrence(t, value){
  if(!t) return;
  if(!t.custom || typeof t.custom !== 'object') t.custom = {};
  const next = normalizeRecurrenceValue(value);
  if(next === 'off'){
    delete t.custom[TASK_RECURRENCE_KEY];
    delete t.custom[TASK_RECURRING_KEY];
    t.recurring = false;
  } else {
    t.custom[TASK_RECURRENCE_KEY] = next;
    t.custom[TASK_RECURRING_KEY] = 'true';
    t.recurring = true;
  }
}
function setTaskRecurring(t, value){ setTaskRecurrence(t, value ? recurrenceValueWithStart('daily', defaultRecurrenceStart()) : 'off'); }
function addRecurrenceInterval(date, meta){
  const next = new Date(date);
  if(meta.unit === 'weeks') next.setDate(next.getDate() + (meta.interval * 7));
  else if(meta.unit === 'months') next.setMonth(next.getMonth() + meta.interval);
  else next.setDate(next.getDate() + meta.interval);
  return next;
}
function nextRecurrenceStart(startsAt, meta){
  const start = new Date(startsAt);
  if(Number.isNaN(start.getTime()) || meta.key === 'off') return '';
  let next = addRecurrenceInterval(start, meta);
  let guard = 0;
  while(next.getTime() <= Date.now() && guard < 366){
    next = addRecurrenceInterval(next, meta);
    guard += 1;
  }
  return localDateTimeValue(next);
}
function advanceTaskRecurrenceStart(t,b,recurrence){
  const meta = recurrenceMeta(recurrence);
  if(!meta.startsAt) return;
  const nextStart = nextRecurrenceStart(meta.startsAt, meta);
  if(!nextStart) return;
  const nextValue = recurrenceValueWithStart(meta.key, nextStart);
  setTaskRecurrence(t, nextValue);
  (b?.columns || []).forEach(c=>{
    if(c.type !== 'recurrence') return;
    const existing = normalizeRecurrenceValue(t?.custom?.[c.id]);
    if(existing !== 'off') t.custom[c.id] = nextValue;
  });
}
function recurrenceResetDue(t, recurrence){
  const meta = recurrenceMeta(recurrence);
  if(meta.key === 'off') return false;
  if(meta.startsAt){
    const scheduled = new Date(meta.startsAt);
    return !Number.isNaN(scheduled.getTime()) && Date.now() >= scheduled.getTime();
  }
  if(!t?.completedAt) return false;
  const completed = new Date(t.completedAt);
  if(Number.isNaN(completed.getTime())) return false;
  const next = new Date(completed);
  const rolled = addRecurrenceInterval(next, meta);
  next.setTime(rolled.getTime());
  next.setHours(0,0,0,0);
  return Date.now() >= next.getTime();
}
function recurrenceResetStatus(b){
  const options = boardStatusOptions(b);
  if(options.not_started) return 'not_started';
  const firstNotDone = Object.keys(options).find(key=>key !== 'done');
  return firstNotDone || 'not_started';
}
function applyRecurringTaskRollovers(data=DATA){
  if(!data?.boards) return false;
  let changed = false;
  data.boards.forEach(b=>(b.groups || []).forEach(g=>taskTreeList(g.tasks).forEach(t=>{
    const recurrence = taskRecurrence(t,b);
    if(recurrence !== 'off' && t.status === 'done' && recurrenceResetDue(t, recurrence)){
      t.status = recurrenceResetStatus(b);
      t.completedAt = '';
      advanceTaskRecurrenceStart(t,b,recurrence);
      changed = true;
    }
  })));
  return changed;
}
function handleRecurringTaskRollovers({renderAfter=false, cloudSave=true} = {}){
  if(!DATA || isEditingControl()) return false;
  const changed = applyRecurringTaskRollovers(DATA);
  if(!changed) return false;
  saveLocal();
  if(cloudSave) queueCloudSave();
  if(renderAfter) render({preserveScroll:true});
  return true;
}
function startRecurringRolloverTimer(){
  clearInterval(RECURRING_ROLLOVER_TIMER);
  RECURRING_ROLLOVER_TIMER = setInterval(()=>{
    if(isEditingControl()) return;
    handleRecurringTaskRollovers({renderAfter:true});
  }, 60000);
}
function stopRecurringRolloverTimer(){
  clearInterval(RECURRING_ROLLOVER_TIMER);
  RECURRING_ROLLOVER_TIMER = null;
}
function ownerStackHtml(ids, {pickKey='', limit=3, emptyText='', small=false} = {}){
  const people = validMemberIds(ids).map(id=>member(id)).filter(Boolean);
  const attrs = pickKey ? ` data-owner-pick="${esc(pickKey)}"` : '';
  const cls = `owner-stack ${small?'small':''}`;
  if(!people.length){
    return `<button type="button" class="owner-empty ${small?'small':''}"${attrs} title="Assign owners">${emptyText ? esc(emptyText) : I.plus}</button>`;
  }
  const shown = people.slice(0, limit);
  const title = people.map(m=>m.name).join(', ');
  return `<button type="button" class="${cls}"${attrs} title="${esc(title)}">
    ${shown.map(m=>`<span class="owner-av ${small?'small':''}" style="background:${esc(m.color)}">${esc(m.initials)}</span>`).join('')}
    ${people.length > limit ? `<span class="owner-more">+${people.length - limit}</span>` : ''}
  </button>`;
}
function asset(id){ return (DATA.assets || []).find(a=>a.id===id); }
function doc(id){ return (DATA.docs || []).find(d=>d.id===id); }
function safeFileName(name){
  return String(name || 'file').replace(/[\\/:*?"<>|#%{}^~[\]`]+/g, '-').replace(/\s+/g, ' ').trim().slice(0,140) || 'file';
}
function fileExtension(name=''){
  const match = String(name || '').toLowerCase().match(/\.([a-z0-9]{1,12})$/);
  return match ? match[1] : '';
}
function isAllowedAssetFile(file){
  const ext = fileExtension(file?.name || '');
  const mime = String(file?.type || '').toLowerCase();
  if(BLOCKED_ASSET_EXTENSIONS.has(ext)) return false;
  if(mime === 'image/svg+xml' || mime === 'text/html' || mime === 'application/javascript' || mime === 'text/javascript') return false;
  if(mime.startsWith('image/') || mime.startsWith('video/') || mime.startsWith('audio/')) return true;
  if(ALLOWED_ASSET_MIME_EXACT.has(mime)) return true;
  return Boolean(ext && ALLOWED_ASSET_EXTENSIONS.has(ext));
}
function validateAssetFile(file, remote=false){
  if(!file) throw new Error('Choose a file to upload.');
  const limit = remote ? MAX_REMOTE_ASSET_BYTES : MAX_LOCAL_ASSET_BYTES;
  if(Number(file.size || 0) > limit){
    const mb = Math.round(limit / 1_000_000);
    throw new Error(`${file.name || 'This file'} is too large. Maximum upload size is ${mb} MB.`);
  }
  if(!isAllowedAssetFile(file)){
    throw new Error(`${file.name || 'This file'} is not an allowed upload type.`);
  }
}
function assetKindForMime(mime=''){
  if(mime.startsWith('image/')) return 'image';
  if(mime.startsWith('video/')) return 'video';
  if(mime === 'application/pdf') return 'pdf';
  if(mime.startsWith('audio/')) return 'audio';
  return 'file';
}
function assetKindForFile(file){ return assetKindForMime(file?.type || ''); }
function assetIcon(a){
  if(a?.kind === 'image') return 'IMG';
  if(a?.kind === 'video') return 'VID';
  if(a?.kind === 'pdf') return 'PDF';
  if(a?.kind === 'audio') return 'AUD';
  if(a?.kind === 'link') return 'URL';
  return 'FILE';
}
function assetUrlKey(a){
  if(a?.data?.canonicalUrl) return a.data.canonicalUrl;
  if(a?.data?.driveId) return `file:google:${a.data.driveId}`;
  return canonicalUrlKey(a?.externalUrl || a?.data?.sourceUrl || '');
}
function findDuplicateAssetUrl(url, {excludeId=''} = {}){
  const key = canonicalUrlKey(url);
  if(!key) return null;
  return (DATA.assets || []).find(a=>a.id !== excludeId && assetUrlKey(a) === key) || null;
}
function duplicateAssetMessage(existing, url=''){
  const key = canonicalUrlKey(url || existing?.externalUrl || '');
  const noun = key.startsWith('video:') ? 'video' : 'link';
  const title = existing?.title ? `: ${existing.title}` : '';
  return `This ${noun} had already been posted${title}.`;
}
function assetSizeLabel(bytes){
  const n = Number(bytes) || 0;
  if(!n) return '';
  if(n < 1024) return `${n} B`;
  if(n < 1024 * 1024) return `${Math.round(n / 1024)} KB`;
  return `${(n / 1024 / 1024).toFixed(n < 10 * 1024 * 1024 ? 1 : 0)} MB`;
}
function taskAssetIdsForColumn(t,cid){ return parseIdList(t?.custom?.[cid]); }
function taskDocIdForColumn(t,cid){ return parseIdList(t?.custom?.[cid])[0] || ''; }
function taskFileAssets(b,t){
  const ids = [];
  (b?.columns || []).filter(c=>c.type === 'files').forEach(c=>ids.push(...taskAssetIdsForColumn(t,c.id)));
  return [...new Set(ids)].map(id=>asset(id)).filter(Boolean);
}
function boardCustomColumnsByType(type){
  return DATA.boards.flatMap(b=>(b.columns||[]).filter(c=>c.type===type).map(c=>({b,c})));
}
function linkedTasksForAsset(assetId){
  const out = [];
  DATA.boards.forEach(b=>{
    const fileColumns = (b.columns||[]).filter(c=>c.type==='files');
    if(!fileColumns.length) return;
    b.groups.forEach(g=>taskTreeList(g.tasks).forEach(t=>{
      if(fileColumns.some(c=>taskAssetIdsForColumn(t,c.id).includes(assetId))) out.push({b,g,t});
    }));
  });
  return out;
}
function linkedTasksForDoc(docId){
  const out = [];
  DATA.boards.forEach(b=>{
    const docColumns = (b.columns||[]).filter(c=>c.type==='document');
    if(!docColumns.length) return;
    b.groups.forEach(g=>taskTreeList(g.tasks).forEach(t=>{
      if(docColumns.some(c=>taskDocIdForColumn(t,c.id) === docId)) out.push({b,g,t});
    }));
  });
  return out;
}
function taskTreeList(tasks=[]){
  return tasks.flatMap(t=>[t, ...taskTreeList(t.subtasks || [])]);
}
function isInternalTaskValueColumn(columnId){
  return INTERNAL_TASK_VALUE_COLUMNS.some(c=>c.id === columnId);
}
function taskValueColumnIdsForBoard(b){
  return new Set([...(b?.columns || []).map(c=>c.id).filter(Boolean), ...INTERNAL_TASK_VALUE_COLUMNS.map(c=>c.id)]);
}
function pruneTaskCustomValues(t, allowedIds){
  if(t?.custom && typeof t.custom === 'object'){
    Object.keys(t.custom).forEach(key=>{ if(!allowedIds.has(key)) delete t.custom[key]; });
  }
  (t?.subtasks || []).forEach(st=>pruneTaskCustomValues(st, allowedIds));
}
function setTaskCustomValueDeep(t, columnId, value){
  if(!t.custom || typeof t.custom !== 'object') t.custom = {};
  t.custom[columnId] = value;
  (t.subtasks || []).forEach(st=>setTaskCustomValueDeep(st, columnId, value));
}
function deleteTaskCustomValueDeep(t, columnId){
  if(t?.custom && typeof t.custom === 'object') delete t.custom[columnId];
  (t?.subtasks || []).forEach(st=>deleteTaskCustomValueDeep(st, columnId));
}
function unassignMemberFromTaskTree(t, memberId){
  if(!t) return;
  if(t.owner === memberId) t.owner = '';
  setTaskOwnerIds(t, taskOwnerIds(t).filter(id=>id !== memberId));
  if(t.assignedBy === memberId) t.assignedBy = '';
  setTaskSubscribers(t, taskSubscribers(t).filter(id=>id !== memberId));
  (t.subtasks || []).forEach(st=>unassignMemberFromTaskTree(st, memberId));
}

function areSubtasksOpen(t){
  return t?.subtasksOpen === true || OPEN_SUBTASK_TASKS.has(t?.id);
}

function setSubtasksOpen(t, open){
  if(!t) return;
  if(!t.custom || typeof t.custom !== 'object') t.custom = {};
  t.subtasksOpen = Boolean(open);
  if(open) OPEN_SUBTASK_TASKS.add(t.id);
  else OPEN_SUBTASK_TASKS.delete(t.id);
  delete t.custom[TASK_SUBTASKS_OPEN_KEY];
}

function isTaskSubscribed(t, memberId){
  return Boolean(t && memberId && taskSubscribers(t).includes(memberId));
}

function updateTaskLocation(u){
  if(!u?.taskId) return null;
  return (u.boardId ? findTaskLocation(u.boardId, u.taskId) : null) || findTaskLocationById(u.taskId);
}

function taskUpdates(t){
  if(!t?.id) return [];
  return feedUpdates().filter(u=>u.taskId === t.id);
}

function unreadTaskUpdates(t, memberId=currentMemberId()){
  if(!memberId) return [];
  return taskUpdates(t).filter(u=>!(u.readBy || []).includes(memberId));
}

function taskUpdateIndicatorHtml(b,g,t){
  const updates = taskUpdates(t);
  const unread = unreadTaskUpdates(t);
  const state = unread.length ? 'unread' : (updates.length ? 'has-updates' : '');
  const visibleCount = unread.length || updates.length;
  const count = visibleCount ? `<span class="task-update-count ${unread.length?'is-new':'is-read'}">${visibleCount > 99 ? '99+' : visibleCount}</span>` : '';
  const title = unread.length
    ? `${unread.length} new update${unread.length===1?'':'s'}`
    : (updates.length ? `${updates.length} task update${updates.length===1?'':'s'}` : 'No updates yet');
  return `<button type="button" class="task-update-indicator ${state}" data-open-task-updates="${esc(b.id)}|${esc(g.id)}|${esc(t.id)}" title="${esc(title)}" aria-label="${esc(title)}">${I.feed}${count}</button>`;
}

function markTaskUpdatesRead(taskId){
  const meId = currentMemberId();
  if(!taskId || !meId || !Array.isArray(DATA.updates)) return false;
  let changed = false;
  DATA.updates.forEach(u=>{
    if(u.taskId !== taskId) return;
    if(!Array.isArray(u.readBy)) u.readBy = [];
    if(!u.readBy.includes(meId)){
      u.readBy.push(meId);
      changed = true;
    }
  });
  return changed;
}

function updateMentionsMember(u, memberId){
  const mentions = u.mentions || [];
  return mentions.includes(memberId) || mentions.includes(EVERYONE_MENTION_ID);
}

function updateRelevantToMember(u, memberId){
  if(!memberId) return false;
  if(calendarFeedRelevantToMember(u)) return true;
  const prefs = notificationPrefs(memberId, {useWorkspaceFallback:false});
  if(prefs.authored && u.authorId === memberId) return true;
  const directMention = (u.mentions||[]).includes(memberId);
  const everyoneMention = (u.mentions||[]).includes(EVERYONE_MENTION_ID);
  if(prefs.mentions && directMention) return true;
  if(prefs.everyone && everyoneMention) return true;
  const loc = updateTaskLocation(u);
  if(!loc) return false;
  if(prefs.assigned && taskOwnerIds(loc.t).includes(memberId)) return true;
  if(prefs.subscribed && isTaskSubscribed(loc.t, memberId)) return true;
  return false;
}

function memberNotifications(memberId=currentMemberId()){
  const email = normalizeEmail(SESSION?.user?.email || member(memberId)?.email || '');
  return (DATA.notifications || []).filter(n=>{
    if(memberId && n.recipientId === memberId) return true;
    return email && normalizeEmail(n.recipientEmail) === email;
  }).sort((a,b)=>(b.createdAt||0)-(a.createdAt||0));
}

function unreadNotifications(memberId=currentMemberId()){
  return memberNotifications(memberId).filter(n=>!n.readAt);
}

function notificationBoardMuted(type, boardId, recipientId){
  const mode = boardNotificationPrefs(boardId, recipientId).mode;
  if(mode !== 'mute') return false;
  return !['assigned','mention','everyone','due_today','due_soon','due_1_day','due_3_day','overdue'].includes(type);
}

function shouldCreateNotification(type, boardId, recipientId, {task=null, update=null} = {}){
  if(!recipientId || recipientId === update?.authorId) return false;
  const prefs = notificationPrefs(recipientId, {useWorkspaceFallback:false});
  if(type === 'assigned') return prefs.assigned !== false;
  if(type === 'mention') return prefs.mentions !== false;
  if(type === 'everyone') return prefs.everyone !== false;
  if(['due_changed','due_today','due_soon','due_1_day','due_3_day','overdue'].includes(type)) return prefs.dueReminders !== false && !notificationBoardMuted(type, boardId, recipientId);
  if(type === 'task_update'){
    if(prefs.taskUpdates === false) return false;
    if(notificationBoardMuted(type, boardId, recipientId)) return false;
    if(boardNotificationPrefs(boardId, recipientId).mode === 'receive_all') return true;
    if(task && prefs.assigned !== false && taskOwnerIds(task).includes(recipientId)) return true;
    if(task && prefs.subscribed !== false && isTaskSubscribed(task, recipientId)) return true;
    return false;
  }
  if(type === 'automation') return !notificationBoardMuted(type, boardId, recipientId);
  return true;
}

function notificationChannelsFor(recipientId, type){
  const prefs = notificationPrefs(recipientId, {useWorkspaceFallback:false});
  const emailDueReminderEnabled = prefs.emailDueReminders !== false && prefs.email !== false;
  return {
    inApp:true,
    desktop:prefs.desktop !== false,
    email:emailDueReminderEnabled && ['due_1_day','due_3_day'].includes(type),
    push:prefs.push !== false && prefs.desktop !== false
  };
}

function createNotification({recipientId, type='general', title, body='', boardId='', groupId='', taskId='', updateId='', actorId=currentMemberId(), data={}, dedupeKey=''}) {
  if(!recipientId) return null;
  if(!Array.isArray(DATA.notifications)) DATA.notifications = [];
  if(dedupeKey && DATA.notifications.some(n=>n.dedupeKey === dedupeKey)) return null;
  const recipient = member(recipientId);
  const channels = notificationChannelsFor(recipientId, type);
  const item = {
    id:uid('ntf'),
    recipientId,
    recipientEmail:normalizeEmail(recipient?.email),
    recipientUserId:recipientId === currentMemberId() ? (SESSION?.user?.id || '') : '',
    actorId,
    boardId,
    groupId,
    taskId,
    updateId,
    type,
    title:title || 'TaskTrack notification',
    body,
    channels,
    data,
    dedupeKey,
    readAt:'',
    emailedAt:'',
    emailStatus:channels.email ? 'pending' : 'skipped',
    emailAttempts:0,
    desktopAt:'',
    pushedAt:'',
    pushStatus:channels.push ? 'pending' : 'skipped',
    pushAttempts:0,
    createdAt:Date.now()
  };
  DATA.notifications.unshift(item);
  DATA.notifications = DATA.notifications.slice(0,500);
  void persistNotificationInsert(item);
  return item;
}

function createTaskAssignmentNotifications({boardId, groupId, taskId, task, recipients=[]}){
  const actor = member(currentMemberId());
  recipients.forEach(recipientId=>{
    if(recipientId === currentMemberId()) return;
    if(!shouldCreateNotification('assigned', boardId, recipientId, {task})) return;
    createNotification({
      recipientId,
      type:'assigned',
      boardId,
      groupId,
      taskId,
      title:`You were assigned to ${task?.name || 'a task'}`,
      body:`${actor?.name || 'A teammate'} assigned you as Task Owner${task?.due ? `, due ${fmtDate(task.due)}` : ''}.`,
      data:{taskName:task?.name || '', due:task?.due || ''},
      dedupeKey:`assigned:${taskId}:${recipientId}:${Date.now()}`
    });
  });
}

function createDueChangedNotifications({boardId, groupId, taskId, task, previousDue='', nextDue=''}) {
  if(previousDue === nextDue || !nextDue) return;
  const recipients = taskOwnerIds(task);
  recipients.forEach(recipientId=>{
    if(recipientId === currentMemberId()) return;
    if(!shouldCreateNotification('due_changed', boardId, recipientId, {task})) return;
    createNotification({
      recipientId,
      type:'due_changed',
      boardId,
      groupId,
      taskId,
      title:`Due date set for ${task?.name || 'a task'}`,
      body:`${task?.name || 'This task'} is now due ${fmtDate(nextDue)}.`,
      data:{taskName:task?.name || '', previousDue, due:nextDue}
    });
  });
}

function createNotificationsForUpdate(u){
  if(!u || !isVisibleUpdate(u)) return;
  const loc = updateTaskLocation(u);
  const recipients = new Map();
  (u.mentions || []).forEach(id=>{
    if(id === EVERYONE_MENTION_ID){
      DATA.team.forEach(m=>{ if(m.id !== u.authorId) recipients.set(m.id, 'everyone'); });
    } else {
      recipients.set(id, 'mention');
    }
  });
  if(loc?.t){
    [...taskOwnerIds(loc.t), ...taskSubscribers(loc.t), ...DATA.team.map(m=>m.id)].forEach(id=>{
      if(!recipients.has(id)) recipients.set(id, 'task_update');
    });
  }
  recipients.forEach((type, recipientId)=>{
    if(recipientId === u.authorId) return;
    if(!shouldCreateNotification(type, u.boardId || loc?.b?.id || '', recipientId, {task:loc?.t || null, update:u})) return;
    const author = member(u.authorId);
    const taskName = loc?.t?.name || '';
    createNotification({
      recipientId,
      type,
      boardId:u.boardId || loc?.b?.id || '',
      groupId:u.groupId || loc?.g?.id || '',
      taskId:u.taskId || '',
      updateId:u.id,
      title:type === 'task_update' && taskName ? `Update on ${taskName}` : (type === 'everyone' ? 'Team update in TaskTrack' : 'You were mentioned in TaskTrack'),
      body:`${author?.name || 'A teammate'}: ${String(u.body || '').replace(/\s+/g,' ').trim().slice(0,180)}`,
      data:{taskName, updateBody:u.body || ''},
      dedupeKey:`update:${u.id}:${recipientId}`
    });
  });
}

function notificationTypeLabel(type){
  const labels = {
    assigned:'Assigned',
    mention:'Mention',
    everyone:'Team update',
    task_update:'Task update',
    due_changed:'Due date',
    due_today:'Due today',
    due_soon:'Due soon',
    due_1_day:'Due tomorrow',
    due_3_day:'Due in 3 days',
    overdue:'Overdue',
    automation:'Automation',
    meeting_shared:'Meeting shared',
    meeting_mention:'Meeting mention'
  };
  return labels[type] || 'Notification';
}

function notificationTargetLabel(n){
  const meetingId = n.data?.meetingId || '';
  const meeting = meetingId ? MEETING_ROWS.find(item=>item.id === meetingId) : null;
  if(meeting) return `${meeting.title}${meeting.meetingDate ? ` / ${meeting.meetingDate}` : ''}`;
  const loc = n.taskId ? findTaskLocationById(n.taskId) : null;
  if(loc?.t) return `${loc.b.name} / ${loc.t.name}`;
  const b = n.boardId ? board(n.boardId) : null;
  return b?.name || 'TaskTrack workspace';
}

function notificationRowHtml(n){
  const unread = !n.readAt;
  const actor = member(n.actorId);
  return `<button type="button" class="notification-row ${unread?'unread':''}" data-notification-open="${esc(n.id)}">
    <span class="notification-dot"></span>
    <span class="notification-main">
      <b>${esc(n.title || notificationTypeLabel(n.type))}</b>
      <span>${esc(n.body || notificationTargetLabel(n))}</span>
      <em>${esc(notificationTypeLabel(n.type))} · ${esc(notificationTargetLabel(n))}${actor?` · ${esc(actor.name)}`:''} · ${esc(fmtTimeAgo(n.createdAt))}</em>
    </span>
  </button>`;
}

function openNotificationsModal(){
  const notices = memberNotifications();
  const unreadCount = notices.filter(n=>!n.readAt).length;
  const desktopControl = desktopNotificationsSupported()
    ? `<button type="button" class="feed-desktop ${desktopNotificationsEnabled()?'active':''}" data-action="request-desktop-notifications">${I.bell} ${desktopNotificationsEnabled()?'Desktop alerts on':'Enable desktop alerts'}</button>`
    : '';
  const body = `<div class="notification-modal">
    <div class="notification-summary">${unreadCount ? `${unreadCount} unread notification${unreadCount===1?'':'s'}` : 'No unread notifications'}${desktopControl}</div>
    <div class="notification-list">
      ${notices.length ? notices.slice(0,60).map(notificationRowHtml).join('') : `<div class="feed-empty notification-empty"><h3>No notifications yet</h3><p>Assignments, mentions, automations, and due-date reminders will appear here.</p></div>`}
    </div>
  </div>`;
  modal('Notifications','Assignments, mentions, automations, and due reminders.',body,()=>{
    markNotificationsRead();
    closeModal();
    render();
  },'Mark all read');
  $('#modal-root .modal')?.classList.add('modal-notifications');
  $('#modal-root').querySelectorAll('[data-notification-open]').forEach(row=>{
    row.onclick=()=>openNotificationTarget(row.dataset.notificationOpen);
  });
  $('#modal-root').querySelector('[data-action="request-desktop-notifications"]')?.addEventListener('click', requestDesktopNotifications);
  const ok = $('#modal-root [data-mc="ok"]');
  if(ok && !unreadCount) ok.disabled = true;
}

function markNotificationsRead(ids=null){
  const selected = ids ? new Set(ids) : null;
  const now = new Date().toISOString();
  let changed = false;
  const changedIds = [];
  memberNotifications().forEach(n=>{
    if(selected && !selected.has(n.id)) return;
    if(!n.readAt){
      n.readAt = now;
      changed = true;
      changedIds.push(n.id);
    }
  });
  if(changed){
    saveLocal();
    void persistNotificationPatch(changedIds, {read_at:now});
  }
  return changed;
}

function openNotificationTarget(id){
  const n = (DATA.notifications || []).find(item=>item.id === id);
  if(!n) return;
  markNotificationsRead([id]);
  closeModal();
  if(n.data?.meetingId){
    openMeeting(n.data.meetingId);
    return;
  }
  if(n.taskId){
    const loc = findTaskLocationById(n.taskId);
    if(loc){
      openTaskDetail(loc.b.id, loc.g.id, loc.t.id, n.updateId ? 'updates' : 'activity');
      render();
      return;
    }
  }
  if(n.boardId && board(n.boardId)) openBoard(n.boardId);
  else navigateToView('updates');
}

function relevantFeedUpdates(){
  const meId = currentMemberId();
  return feedUpdates().filter(u=>updateRelevantToMember(u, meId));
}

function filteredUpdates(){
  if(!UPDATE_FILTERS.has(UPDATE_FILTER)) UPDATE_FILTER = 'all';
  const meId = currentMemberId();
  return relevantFeedUpdates().filter(u=>{
    if(UPDATE_SHOW_UNREAD && (u.readBy||[]).includes(meId)) return false;
    if(UPDATE_BOARD_FILTER==='none' && u.boardId) return false;
    if(UPDATE_BOARD_FILTER!=='all' && UPDATE_BOARD_FILTER!=='none' && u.boardId!==UPDATE_BOARD_FILTER) return false;
    if(UPDATE_FILTER==='mentions') return updateMentionsMember(u, meId);
    return true;
  });
}

function renderUpdateCard(u){
  if(u.kind === 'shared_calendar') return renderSharedCalendarUpdateCard(u);
  const meId = currentMemberId();
  const author = member(u.authorId) || {name:'TaskTrack',initials:'FR',color:'#0073ea'};
  const loc = updateTaskLocation(u);
  const b = u.boardId ? board(u.boardId) : null;
  const unread = !(u.readBy||[]).includes(meId);
  const directMentioned = (u.mentions||[]).includes(meId);
  const teamMentioned = (u.mentions||[]).includes(EVERYONE_MENTION_ID);
  const mentioned = directMentioned || teamMentioned;
  const subscribed = loc ? isTaskSubscribed(loc.t, meId) : false;
  return `<article class="feed-card ${unread?'unread':''} ${mentioned?'mentioned':''}" data-update-id="${u.id}">
    <button class="feed-read-toggle ${unread?'unread':''}" data-update-read="${u.id}" title="${unread?'Mark read':'Mark unread'}">${unread?'':'✓'}</button>
    <div class="feed-author">
      <div class="owner-av" style="background:${author.color}">${esc(author.initials)}</div>
      <div><b>${esc(author.name)}</b><span>${fmtTimeAgo(u.createdAt)}</span></div>
    </div>
    <div class="feed-body">${renderUpdateText(u.body||'')}</div>
    ${(u.assetIds||[]).length ? `<div class="feed-attachments">${u.assetIds.map(id=>{ const a=asset(id); return a ? `<button type="button" data-open-feed-asset="${esc(a.id)}">📎 ${esc(a.title)}</button>` : ''; }).join('')}</div>` : ''}
    ${(b || loc) ? `<div class="feed-card-meta">
      ${b?`<span>${esc(b.name)}</span>`:''}
      ${loc?`<span>${esc(loc.g.name)}</span><button data-open-feed-task="${esc((u.boardId || loc.b.id) + '|' + u.taskId)}">${esc(loc.t.name)}</button>`:''}
    </div>`:''}
    <div class="feed-card-actions">
      ${mentioned?`<span class="mention-note">${I.mention} ${teamMentioned && !directMentioned ? 'Team update' : 'Mentioned you'}</span>`:''}
      ${loc?`<button class="${subscribed?'active':''}" data-subscribe-task="${esc((u.boardId || loc.b.id) + '|' + u.taskId)}">${I.bell} ${subscribed?'Following task':'Follow task'}</button>`:''}
      ${canDeleteUpdate(u)?`<button class="danger-link" data-delete-update="${esc(u.id)}">${I.trash} Delete</button>`:''}
    </div>
  </article>`;
}

function renderUpdateText(text){
  let out = renderLinkedText(text, {lineBreaks:true});
  out = out.replace(/@everyone\b/gi, '<span class="mention-chip">@Everyone</span>');
  DATA.team.forEach(m=>{
    const labels = [m.name, m.email ? m.email.split('@')[0] : '', m.initials].filter(Boolean);
    labels.forEach(label=>{
      out = out.replace(new RegExp(rxEscape(`@${esc(label)}`),'gi'), () => `<span class="mention-chip">@${esc(m.name)}</span>`);
    });
  });
  return out;
}

function notifyNewRemoteUpdates(nextData, previousData){
  if(ROW_TABLE_AVAILABLE.app_notifications !== false) return;
  if(!nextData || !previousData || !desktopNotificationsEnabled()) return;
  const previousIds = new Set((previousData.updates || []).map(u=>u.id));
  const meId = currentMemberId();
  (nextData.updates || []).forEach(u=>{
    if(!u?.id || previousIds.has(u.id) || u.authorId === meId) return;
    if(!isVisibleUpdate(u)) return;
    if(!updateRelevantToMember(u, meId)) return;
    const author = member(u.authorId);
    const loc = updateTaskLocation(u);
    const title = loc ? `TaskTrack update: ${loc.t.name}` : 'TaskTrack workspace update';
    const body = `${author?.name || 'A teammate'}: ${String(u.body || '').replace(/\s+/g,' ').trim().slice(0,120)}`;
    try{
      new Notification(title, {
        body,
        tag:`friday-update-${u.id}`,
        icon:'./favicon.png'
      });
    }catch(e){}
  });
}

function notifyNewRemoteNotifications(nextData, previousData){
  if(!nextData || !previousData || !desktopNotificationsEnabled()) return;
  const previousIds = new Set((previousData.notifications || []).map(n=>n.id));
  memberNotifications().forEach(n=>{
    if(!n?.id || previousIds.has(n.id) || n.readAt || n.desktopAt) return;
    showDesktopNotificationForId(n.id);
  });
}

function showDesktopNotificationForId(id){
  const n = (DATA.notifications || []).find(item=>item.id === id);
  if(!n || n.readAt || n.desktopAt || n.channels?.desktop === false || !desktopNotificationsEnabled()) return false;
  const isMine = memberNotifications().some(item=>item.id === n.id);
  if(!isMine) return false;
  try{
    const notice = new Notification(n.title || 'TaskTrack notification', {
      body:n.body || notificationTargetLabel(n),
      tag:`friday-notification-${n.id}`,
      icon:'./favicon.png'
    });
    notice.onclick = ()=>{
      window.focus?.();
      openNotificationTarget(n.id);
    };
    n.desktopAt = new Date().toISOString();
    saveLocal();
    void persistNotificationPatch([n.id], {desktop_at:n.desktopAt});
    return true;
  }catch(e){
    return false;
  }
}

function rxEscape(s){ return String(s).replace(/[.*+?^${}()|[\]\\]/g,'\\$&'); }
function daysUntil(iso){
  if(!iso) return null;
  const d = new Date(iso+'T00:00');
  if(Number.isNaN(d.getTime())) return null;
  const today = new Date(); today.setHours(0,0,0,0);
  return Math.round((d - today) / 86400000);
}
function isDueWithin(iso, days){
  const diff = daysUntil(iso);
  return diff !== null && diff >= 0 && diff <= days;
}
function isOverdueTask(t){ return t.status !== 'done' && isPast(t.due); }

function matchesBoardView(t){
  if(BOARD_VIEW === 'hidden') return taskIsHidden(t);
  if(taskIsHidden(t)) return false;
  if(BOARD_VIEW === 'mine') return taskOwnerIds(t).includes(currentMemberId());
  if(BOARD_VIEW === 'due_week') return t.status !== 'done' && isDueWithin(t.due, 7);
  if(BOARD_VIEW === 'overdue') return isOverdueTask(t);
  if(BOARD_VIEW === 'blocked') return t.status === 'stuck';
  if(BOARD_VIEW === 'unassigned') return !taskOwnerIds(t).length;
  if(BOARD_VIEW === 'review') return t.status === 'review';
  return true;
}

function boardViewBar(){
  return `<div class="board-view-tabs">
    ${BOARD_VIEWS.map(v=>`<button class="${BOARD_VIEW===v.id?'active':''}" data-board-view="${v.id}">${esc(v.label)}</button>`).join('')}
  </div>`;
}

function boardPersonFilter(b){
  const id = BOARD_PERSON_FILTERS[b?.id] || '';
  return DATA.team.some(m=>m.id === id) ? id : '';
}

function setBoardPersonFilter(b,id){
  if(!b?.id) return;
  if(id && DATA.team.some(m=>m.id === id)) BOARD_PERSON_FILTERS[b.id] = id;
  else delete BOARD_PERSON_FILTERS[b.id];
}

function clearBoardFilters(){
  const b = board(CURRENT_BOARD);
  SEARCH = '';
  BOARD_VIEW = 'all';
  if(b) setBoardPersonFilter(b,'');
  renderContent();
}

function boardFilterSummary(b,q=''){
  const parts = [];
  if(q) parts.push(`search "${SEARCH.trim()}"`);
  if(BOARD_VIEW !== 'all') parts.push(BOARD_VIEWS.find(v=>v.id===BOARD_VIEW)?.label || 'view filter');
  const personId = boardPersonFilter(b);
  if(personId) parts.push(`person: ${member(personId)?.name || 'teammate'}`);
  return parts;
}

function mentionIdsFromText(text, extra=[]){
  const hay = String(text||'').toLowerCase();
  const ids = new Set(extra.filter(Boolean));
  if(hay.includes('@everyone')) ids.add(EVERYONE_MENTION_ID);
  DATA.team.forEach(m=>{
    const labels = [m.name, m.name?.split(/\s+/)[0], m.email ? m.email.split('@')[0] : '', m.initials].filter(Boolean);
    if(labels.some(label=>hay.includes(`@${String(label).toLowerCase()}`))) ids.add(m.id);
  });
  return [...ids];
}

function recordUpdate({body,boardId='',groupId='',taskId='',assetIds=[],mentions=[],kind='update',visible=true}){
  if(!body) return null;
  if(!Array.isArray(DATA.updates)) DATA.updates = [];
  const meId = currentMemberId();
  const loc = taskId ? ((boardId ? findTaskLocation(boardId, taskId) : null) || findTaskLocationById(taskId)) : null;
  if(loc && meId && !isTaskSubscribed(loc.t, meId)){
    setTaskSubscribers(loc.t, [...taskSubscribers(loc.t), meId]);
  }
  const item = {
    id:uid('upd'),
    authorId:meId,
    boardId,
    groupId,
    taskId,
    body,
    assetIds,
    kind,
    visible,
    mentions:mentionIdsFromText(body, mentions),
    readBy:[meId],
    bookmarkedBy:[],
    createdAt:Date.now()
  };
  DATA.updates.unshift(item);
  DATA.updates = DATA.updates.slice(0,500);
  createNotificationsForUpdate(item);
  return item;
}

async function postFeedUpdate(){
  if(FEED_POSTING) return;
  const body = $('#feed-body')?.value.trim();
  if(!body && !FEED_FILES.length){ toast('Write an update or attach a file first'); return; }
  let boardId = $('#feed-board')?.value || '';
  let groupId = '';
  let taskId = '';
  const taskVal = $('#feed-task')?.value || '';
  if(taskVal){
    const parts = taskVal.split('|');
    boardId = parts[0] || boardId;
    groupId = parts[1] || '';
    taskId = parts[2] || '';
  }
  const files = FEED_FILES.slice();
  try{
    files.forEach(file=>validateAssetFile(file,Boolean(supabase && SESSION?.user && WORKSPACE_ID)));
    FEED_POSTING = true;
    const button=document.querySelector('[data-action="post-update"]');
    if(button) button.disabled=true;
    const assets = files.length ? await uploadAssetFiles(files,{quiet:true}) : [];
    if(files.length && assets.length !== files.length){ toast('Some files could not be attached. Please retry.'); return; }
    recordUpdate({body:body || `Shared ${assets.length} file${assets.length===1?'':'s'}`,boardId,groupId,taskId,assetIds:assets.map(a=>a.id)});
    FEED_FILES = [];
    save();
    renderContent();
    toast('Update posted');
  }catch(error){ toast(error?.message || 'Could not post update'); }
  finally{ FEED_POSTING=false; const button=document.querySelector('[data-action="post-update"]'); if(button) button.disabled=false; }
}

function currentMentionQuery(text, pos){
  const before = String(text||'').slice(0,pos);
  const m = before.match(/(^|\s)@([A-Za-z0-9_. -]{0,40})$/);
  return m ? m[2].trim().toLowerCase() : '';
}

function filterMentionSuggestions(){
  const input = $('#feed-body');
  const wrap = $('#mention-suggestions');
  if(!input || !wrap) return;
  const q = currentMentionQuery(input.value, input.selectionStart || input.value.length);
  wrap.querySelectorAll('[data-mention-insert]').forEach(btn=>{
    if(btn.dataset.mentionInsert === EVERYONE_MENTION_ID){
      btn.style.display = !q || 'everyone team all'.includes(q) ? '' : 'none';
      return;
    }
    const m = member(btn.dataset.mentionInsert);
    const hay = `${m?.name||''} ${m?.email||''} ${m?.initials||''}`.toLowerCase();
    btn.style.display = !q || hay.includes(q) ? '' : 'none';
  });
}

function insertMention(memberId){
  const isEveryone = memberId === EVERYONE_MENTION_ID;
  const m = isEveryone ? {name:'Everyone'} : member(memberId);
  const input = $('#feed-body');
  if(!m || !input) return;
  const pos = input.selectionStart || input.value.length;
  const before = input.value.slice(0,pos);
  const after = input.value.slice(pos);
  const replaced = before.replace(/(^|\s)@([A-Za-z0-9_. -]{0,40})$/, `$1@${m.name} `);
  input.value = replaced === before ? `${before}${before && !/\s$/.test(before)?' ':''}@${m.name} ${after}` : `${replaced}${after}`;
  input.focus();
  input.selectionStart = input.selectionEnd = input.value.length - after.length;
  filterMentionSuggestions();
}

function toggleTaskSubscription(key){
  const [boardId, taskId] = String(key||'').split('|');
  const loc = (boardId ? findTaskLocation(boardId, taskId) : null) || findTaskLocationById(taskId);
  if(!loc){ toast('Task not found'); return; }
  const meId = currentMemberId();
  const ids = new Set(taskSubscribers(loc.t));
  const following = !ids.has(meId);
  if(following) ids.add(meId);
  else ids.delete(meId);
  setTaskSubscribers(loc.t, [...ids]);
  save();
  renderContent();
  toast(following ? 'Following task updates' : 'Stopped following task');
}

function toggleTaskRecurring(key){
  const [gid, tid] = String(key || '').split('|');
  const loc = findTaskLocationById(tid);
  if(!loc || (gid && loc.g.id !== gid)){ toast('Task not found'); return; }
  const next = taskRecurrence(loc.t, loc.b) === 'off' ? recurrenceValueWithStart('daily', defaultRecurrenceStart()) : 'off';
  setTaskRecurrenceForLocation(loc, next);
  save();
  render();
  toast(next === 'off' ? 'Recurring task disabled' : 'Daily recurring task enabled');
}
function setTaskRecurrenceForLocation(loc,value){
  if(!loc?.t) return;
  const next = normalizeRecurrenceValue(value);
  setTaskRecurrence(loc.t,next);
  if(!loc.t.custom || typeof loc.t.custom !== 'object') loc.t.custom = {};
  const recurrenceColumns = (loc.b?.columns || []).filter(c=>c.type === 'recurrence');
  if(next === 'off'){
    recurrenceColumns.forEach(c=>{
      if(normalizeRecurrenceValue(loc.t.custom[c.id]) !== 'off') loc.t.custom[c.id] = '';
    });
    return;
  }
  const target = recurrenceColumns.find(c=>normalizeRecurrenceValue(loc.t.custom[c.id]) !== 'off') || recurrenceColumns[0];
  if(target) loc.t.custom[target.id] = next;
}
function recurrenceStartParts(value){
  const local = localDateTimeValue(value) || defaultRecurrenceStart();
  const [rawDate, rawTime] = local.split('T');
  return {
    date:validIsoDate(rawDate) || todayIso(),
    time:validTimeValue(rawTime) || '09:00'
  };
}
function recurrenceStartValue(date,time){
  return `${validIsoDate(date) || todayIso()}T${validTimeValue(time) || '09:00'}`;
}
function recurrenceStartDisplay(value){
  const {date,time} = recurrenceStartParts(value);
  const d = new Date(`${date}T${time}:00`);
  if(Number.isNaN(d.getTime())) return `${date} ${time}`;
  return d.toLocaleString(undefined,{
    year:'numeric',
    month:'2-digit',
    day:'2-digit',
    hour:'numeric',
    minute:'2-digit'
  }).replace(',', '');
}
function recurrenceStartPanelHtml(state){
  const {date,time} = recurrenceStartParts(state.value);
  const month = /^\d{4}-\d{2}$/.test(String(state.month || '')) ? state.month : pickerMonthIso(date);
  const monthLabel = new Date(`${month}-01T00:00:00`).toLocaleDateString(undefined,{month:'long',year:'numeric'});
  return `<div class="recurrence-start-panel" data-recurrence-start-panel>
    <div class="date-picker-head">
      <b>${esc(monthLabel)}</b>
      <button type="button" data-recurrence-start-today>Today</button>
      <button type="button" data-recurrence-start-month="-1" aria-label="Previous month">&lt;</button>
      <button type="button" data-recurrence-start-month="1" aria-label="Next month">&gt;</button>
    </div>
    <div class="date-picker-grid">${taskDatePickerGridHtml({mode:'due',start:date,month})}</div>
    <label class="recurrence-start-time"><span>Time</span><input type="time" data-recurrence-start-time value="${esc(time)}"></label>
  </div>`;
}
function recurrenceStartControlHtml(value){
  const parts = recurrenceStartParts(value);
  return `<div class="recurrence-start-control" data-recurrence-start-control>
    <input type="hidden" id="recurrence-start-at" value="${esc(recurrenceStartValue(parts.date,parts.time))}">
    <button type="button" class="recurrence-start-field" data-recurrence-start-toggle aria-expanded="false">
      <span data-recurrence-start-label>${esc(recurrenceStartDisplay(value))}</span>
      <span class="recurrence-start-icon">${I.cal}</span>
    </button>
  </div>`;
}
function wireRecurrenceStartControl(dd){
  const control = dd?.querySelector('[data-recurrence-start-control]');
  if(!control) return;
  const hidden = control.querySelector('#recurrence-start-at');
  const toggle = control.querySelector('[data-recurrence-start-toggle]');
  const label = control.querySelector('[data-recurrence-start-label]');
  const state = {
    value:hidden?.value || defaultRecurrenceStart(),
    open:false,
    month:pickerMonthIso(recurrenceStartParts(hidden?.value || defaultRecurrenceStart()).date)
  };
  const setValue = value=>{
    state.value = value;
    if(hidden) hidden.value = value;
    if(label) label.textContent = recurrenceStartDisplay(value);
  };
  const renderPanel = ()=>{
    control.querySelector('[data-recurrence-start-panel]')?.remove();
    toggle?.setAttribute('aria-expanded', state.open ? 'true' : 'false');
    if(!state.open) return;
    control.insertAdjacentHTML('beforeend', recurrenceStartPanelHtml(state));
    const panel = control.querySelector('[data-recurrence-start-panel]');
    panel?.querySelectorAll('[data-picker-date]').forEach(btn=>btn.addEventListener('click',()=>{
      const nextDate = validIsoDate(btn.dataset.pickerDate);
      if(!nextDate) return;
      const {time} = recurrenceStartParts(state.value);
      setValue(recurrenceStartValue(nextDate,time));
      state.month = pickerMonthIso(nextDate);
      renderPanel();
    }));
    panel?.querySelector('[data-recurrence-start-time]')?.addEventListener('change',e=>{
      const {date} = recurrenceStartParts(state.value);
      setValue(recurrenceStartValue(date,e.target.value));
    });
    panel?.querySelector('[data-recurrence-start-today]')?.addEventListener('click',()=>{
      const {time} = recurrenceStartParts(state.value);
      const nextDate = todayIso();
      setValue(recurrenceStartValue(nextDate,time));
      state.month = pickerMonthIso(nextDate);
      renderPanel();
    });
    panel?.querySelectorAll('[data-recurrence-start-month]').forEach(btn=>btn.addEventListener('click',()=>{
      state.month = shiftPickerMonth(state.month, btn.dataset.recurrenceStartMonth);
      renderPanel();
    }));
  };
  toggle?.addEventListener('click',e=>{
    e.stopPropagation();
    state.open = !state.open;
    renderPanel();
  });
}
function pickTaskRecurrence(el){
  const [gid, tid] = String(el.dataset.taskRecurring || '').split('|');
  const loc = findTaskLocationById(tid);
  if(!loc || (gid && loc.g.id !== gid)){ toast('Task not found'); return; }
  const current = taskRecurrence(loc.t, loc.b);
  const currentMeta = recurrenceMeta(current);
  const opts = Object.entries(RECURRENCE_OPTIONS).map(([key,opt])=>`
    <button class="dd-opt recurrence-option ${currentMeta.key===key?'selected':''}" data-recurrence-k="${key}">
      <span class="recurrence-dot" style="background:${esc(opt.color)}"></span>
      <span>${esc(opt.label)}</span>
    </button>`).join('');
  const customInterval = currentMeta.custom ? currentMeta.interval : 2;
  const customUnit = currentMeta.custom ? currentMeta.unit : 'weeks';
  const startValue = currentMeta.startsAt ? localDateTimeValue(currentMeta.startsAt) : defaultRecurrenceStart();
  const unitOptions = Object.entries(RECURRENCE_UNITS).map(([unit,meta])=>
    `<option value="${unit}" ${customUnit===unit?'selected':''}>${esc(meta.label)}</option>`
  ).join('');
  const dd=buildDD(el, `<div class="recurrence-menu">
    <div class="control-title">Recurring schedule</div>
    ${opts}
    <div class="dd-divider"></div>
    <div class="recurrence-custom">
      <label>Next reset date and time</label>
      ${recurrenceStartControlHtml(startValue)}
      <button class="btn ghost sm" id="recurrence-date-apply">Apply date/time</button>
    </div>
    <div class="dd-divider"></div>
    <div class="recurrence-custom">
      <label>Custom interval</label>
      <div class="recurrence-custom-row">
        <input type="number" min="1" max="365" step="1" id="recurrence-custom-number" value="${esc(customInterval)}">
        <select id="recurrence-custom-unit">${unitOptions}</select>
      </div>
      <button class="btn sm" id="recurrence-custom-apply">Apply custom</button>
    </div>
  </div>`);
  wireRecurrenceStartControl(dd);
  const selectedStart = ()=>dd.querySelector('#recurrence-start-at')?.value || defaultRecurrenceStart();
  const apply = value=>{
    setTaskRecurrenceForLocation(loc,value);
    save();
    closeDD();
    render();
  };
  dd.querySelectorAll('[data-recurrence-k]').forEach(btn=>btn.onclick=()=>{
    const key = btn.dataset.recurrenceK;
    apply(key === 'off' ? 'off' : recurrenceValueWithStart(key, selectedStart()));
  });
  dd.querySelector('#recurrence-date-apply')?.addEventListener('click',()=>{
    const key = currentMeta.key === 'off' ? 'daily' : currentMeta.key;
    apply(recurrenceValueWithStart(key, selectedStart()));
  });
  dd.querySelector('#recurrence-custom-apply')?.addEventListener('click',()=>{
    const interval = dd.querySelector('#recurrence-custom-number')?.value || 1;
    const unit = dd.querySelector('#recurrence-custom-unit')?.value || 'days';
    apply(recurrenceValueWithStart(customRecurrenceValue(interval,unit), selectedStart()));
  });
}

function toggleUpdateBookmark(id){
  const u = DATA.updates.find(x=>x.id===id); if(!u)return;
  const meId = currentMemberId();
  if(!Array.isArray(u.bookmarkedBy)) u.bookmarkedBy = [];
  if(u.bookmarkedBy.includes(meId)) u.bookmarkedBy = u.bookmarkedBy.filter(x=>x!==meId);
  else u.bookmarkedBy.push(meId);
  save(); renderContent();
}

function toggleUpdateRead(id){
  if(id?.startsWith('shared-calendar:')){
    void toggleSharedCalendarRead(id.slice('shared-calendar:'.length));
    return;
  }
  const u = DATA.updates.find(x=>x.id===id); if(!u)return;
  const meId = currentMemberId();
  if(!Array.isArray(u.readBy)) u.readBy = [];
  if(u.readBy.includes(meId)) u.readBy = u.readBy.filter(x=>x!==meId);
  else u.readBy.push(meId);
  save(); render();
}

function canDeleteUpdate(u){
  if(!u?.id) return false;
  const meId = currentMemberId();
  return Boolean(meId && (u.authorId === meId || isCurrentAdmin()));
}

function deleteUpdate(id){
  const u = (DATA.updates || []).find(x=>x.id === id);
  if(!u){ toast('Update not found'); return; }
  if(!canDeleteUpdate(u)){ toast('Only the author or an admin can delete this update'); return; }
  const author = member(u.authorId);
  const preview = String(u.body || '').replace(/\s+/g,' ').trim().slice(0,120);
  confirmModal('Delete this update?', `${author?.name || 'A teammate'} posted: "${preview || 'No update text'}"`, ()=>{
    DATA.updates = (DATA.updates || []).filter(x=>x.id !== id);
    if(Array.isArray(DATA.notifications)) DATA.notifications = DATA.notifications.filter(n=>n.updateId !== id);
    save();
    void persistNotificationsForDeletedUpdate(id);
    closeModal();
    render();
    toast('Update deleted');
  });
}

async function persistNotificationsForDeletedUpdate(updateId){
  if(!updateId || !supabase || !SESSION?.user || !WORKSPACE_ID || ROW_TABLE_AVAILABLE.app_notifications === false) return;
  try{
    const { error } = await supabase
      .from('app_notifications')
      .delete()
      .eq('workspace_id', WORKSPACE_ID)
      .eq('update_id', updateId);
    if(error){
      if(isMissingRowTableError(error)){ ROW_TABLE_AVAILABLE.app_notifications = false; return; }
      console.warn('Notification cleanup failed', error);
    }
  }catch(error){
    console.warn('Notification cleanup failed', error);
  }
}

async function toggleSharedCalendarRead(eventId){
  const userId = SESSION?.user?.id;
  if(!supabase || !WORKSPACE_ID || !userId || !SHARED_GOOGLE_EVENTS.some(row=>row.id===eventId)) return;
  const wasRead = SHARED_GOOGLE_READS.has(eventId);
  const query = wasRead
    ? supabase.from('app_shared_calendar_event_reads').delete().eq('workspace_id',WORKSPACE_ID).eq('event_id',eventId).eq('user_id',userId)
    : supabase.from('app_shared_calendar_event_reads').insert({workspace_id:WORKSPACE_ID,event_id:eventId,user_id:userId});
  const {error} = await query;
  if(error){ toast('Could not update event read state'); return; }
  if(wasRead) SHARED_GOOGLE_READS.delete(eventId);
  else SHARED_GOOGLE_READS.add(eventId);
  render();
}

async function markUpdatesRead(){
  const meId = currentMemberId();
  DATA.updates.forEach(u=>{ if(!u.readBy) u.readBy=[]; if(!u.readBy.includes(meId)) u.readBy.push(meId); });
  save();
  const userId = SESSION?.user?.id;
  const unread = SHARED_GOOGLE_EVENTS.filter(row=>!SHARED_GOOGLE_READS.has(row.id));
  if(supabase && WORKSPACE_ID && userId && unread.length){
    const rows = unread.map(row=>({workspace_id:WORKSPACE_ID,event_id:row.id,user_id:userId}));
    const {error} = await supabase.from('app_shared_calendar_event_reads').insert(rows);
    if(error) toast('Could not mark shared events read');
    else unread.forEach(row=>SHARED_GOOGLE_READS.add(row.id));
  }
  render();
}

/* ============================================================
   TASK DETAIL DRAWER
   ============================================================ */
const TASK_DETAIL_TABS = [
  {id:'updates', label:'Updates', icon:I.feed},
  {id:'files', label:'Files', icon:I.docs},
  {id:'links', label:'Links', icon:I.link},
  {id:'activity', label:'Activity Log', icon:I.sort},
  {id:'description', label:'Item Description', icon:I.board},
  {id:'emails', label:'Emails & Activities', icon:I.mention}
];

function openTaskDetail(boardId,groupId,taskId,tab='updates'){
  const loc = (boardId ? findTaskLocation(boardId, taskId) : null) || findTaskLocationById(taskId);
  if(!loc){ toast('Task not found'); return; }
  TASK_DETAIL = {boardId:loc.b.id, groupId:loc.g.id, taskId:loc.t.id};
  TASK_DETAIL_TAB = TASK_DETAIL_TABS.some(t=>t.id===tab) ? tab : 'updates';
  const readChanged = markTaskUpdatesRead(loc.t.id);
  if(readChanged) save();
  render();
}

function closeTaskDetail(){
  TASK_DETAIL = null;
  TASK_DETAIL_TAB = 'updates';
  render();
}

function currentTaskDetailLocation(){
  if(!TASK_DETAIL) return null;
  const loc = (TASK_DETAIL.boardId ? findTaskLocation(TASK_DETAIL.boardId, TASK_DETAIL.taskId) : null) || findTaskLocationById(TASK_DETAIL.taskId);
  if(!loc){ TASK_DETAIL = null; return null; }
  TASK_DETAIL.boardId = loc.b.id;
  TASK_DETAIL.groupId = loc.g.id;
  TASK_DETAIL.taskId = loc.t.id;
  return loc;
}

function taskDetailDrawer(){
  const loc = currentTaskDetailLocation();
  if(!loc) return '';
  const {b,g,t} = loc;
  const st = statusOptionFor(b,t.status);
  const pr = priorityOptionFor(b,t.priority);
  const ownerIds = taskOwnerIds(t);
  const ownerKey = `${esc(t.id)}|${esc(g.id)}`;
  const recurrence = taskRecurrence(t,b);
  const recurring = recurrence !== 'off';
  const color = taskColor(t);
  return `<div class="task-drawer-layer" data-task-detail-layer>
    <div class="task-drawer-backdrop" data-action="close-task-detail"></div>
    <aside class="task-drawer" role="dialog" aria-label="Task details">
      <header class="task-drawer-head">
        <button class="icon-btn" data-action="close-task-detail" title="Close">${I.close}</button>
        <div class="task-drawer-title">
          <input class="task-detail-title-input" data-task-detail-name value="${esc(t.name)}" aria-label="Task name">
          <div class="task-drawer-meta">${esc(b.name)} / ${esc(g.name)}</div>
        </div>
        <span class="task-presence-slot" data-presence-target="task|${esc(t.id)}">${presenceBadgesHtml('task', t.id)}</span>
        <button class="icon-btn" data-owner-pick="${ownerKey}" title="Manage owners">${I.mention}</button>
      </header>
      <div class="task-detail-summary">
        <div>
          <span>Status</span>
          <b style="background:${esc(st.color)}">${esc(st.label)}</b>
        </div>
        <div>
          <span>Priority</span>
          <b style="background:${esc(pr.color)}">${esc(pr.label)}</b>
        </div>
        <div>
          <span>Timeline</span>
          <strong>${esc(taskTimelineLabel(t) || 'No timeline')}</strong>
        </div>
        <div>
          <span>Owners</span>
          ${ownerStackHtml(ownerIds, {pickKey:`${t.id}|${g.id}`, limit:4, small:true, emptyText:'+'})}
        </div>
        <div>
          <span>Calendar color</span>
          <button type="button" class="task-color-detail ${color?'active':''}" ${color?`style="--task-color:${esc(color)};--task-color-bg:${esc(rgbaFromHex(color,0.14))}"`:''} data-task-color-pick="${esc(t.id)}|${esc(g.id)}" title="${color?esc(color):'No calendar color'}">
            <i></i>${esc(taskColorLabel(t))}
          </button>
        </div>
        <div>
          <span>Recurring</span>
          <button type="button" class="task-recurring-toggle ${recurring?'active':''}" data-task-recurring="${esc(g.id)}|${esc(t.id)}">
            ${recurring ? recurrenceLabel(recurrence) : 'Off'}
          </button>
        </div>
      </div>
      <nav class="task-detail-tabs">
        ${TASK_DETAIL_TABS.map(tab=>`<button class="${TASK_DETAIL_TAB===tab.id?'active':''}" data-task-detail-tab="${tab.id}">${tab.icon}<span>${esc(tab.label)}</span></button>`).join('')}
      </nav>
      <section class="task-detail-body">
        ${taskDetailTabContent(loc)}
      </section>
    </aside>
  </div>`;
}

function taskDetailTabContent(loc){
  if(TASK_DETAIL_TAB === 'files') return taskDetailFilesHtml(loc);
  if(TASK_DETAIL_TAB === 'links') return taskDetailLinksHtml(loc);
  if(TASK_DETAIL_TAB === 'activity') return taskDetailActivityHtml(loc);
  if(TASK_DETAIL_TAB === 'description') return taskDetailDescriptionHtml(loc);
  if(TASK_DETAIL_TAB === 'emails') return taskDetailEmailsHtml(loc);
  return taskDetailUpdatesHtml(loc);
}

function taskDetailUpdatesHtml(loc){
  const {b,g,t} = loc;
  const updates = feedUpdates().filter(u=>u.taskId === t.id);
  const people = DATA.team.map(m=>`<button type="button" data-task-mention="${m.id}">@${esc(m.name)}</button>`).join('');
  return `<div class="task-updates-panel">
    <div class="task-update-composer">
      <textarea id="task-detail-update-body" rows="4" placeholder="Write an update. Use @Name or @Everyone to notify teammates."></textarea>
      <div class="task-mention-row">
        <button type="button" data-task-mention="${EVERYONE_MENTION_ID}">@Everyone</button>
        ${people}
      </div>
      <button class="btn" data-action="task-post-update">Post update</button>
    </div>
    <div class="task-update-list">
      ${updates.map(taskDetailUpdateCard).join('') || `<div class="task-detail-empty">
        <h3>No task updates yet</h3>
        <p>Post decisions, blockers, or next steps here so the team can follow the work.</p>
      </div>`}
    </div>
  </div>`;
}

function taskDetailUpdateCard(u){
  const author = member(u.authorId) || {name:'TaskTrack',initials:'FR',color:'#0073ea'};
  return `<article class="task-update-card">
    <div class="task-update-author">
      <span class="owner-av small" style="background:${esc(author.color)}">${esc(author.initials)}</span>
      <b>${esc(author.name)}</b>
      <em>${esc(fmtTimeAgo(u.createdAt))}</em>
      ${canDeleteUpdate(u)?`<button type="button" class="task-update-delete" data-delete-update="${esc(u.id)}" title="Delete update">${I.trash}</button>`:''}
    </div>
    <div class="task-update-text">${renderUpdateText(u.body || '')}</div>
  </article>`;
}

function taskDetailFilesHtml(loc){
  const files = taskFileAssets(loc.b, loc.t);
  return `<div class="task-files-panel">
    <div class="task-detail-actions">
      <button class="btn sm" data-action="task-upload-file">${I.plus}<span>Upload file</span></button>
      <button class="btn ghost sm" data-action="task-attach-existing-file">Attach from Docs Hub</button>
      <input type="file" id="task-detail-file-input" multiple hidden>
    </div>
    <div class="task-files-grid">
      ${files.map(taskDetailFileCard).join('') || `<div class="task-detail-empty wide">
        <h3>No files attached yet</h3>
        <p>Upload files here or attach files from Docs Hub. They will also appear in the board's Files column.</p>
      </div>`}
    </div>
  </div>`;
}

function taskDetailFileCard(a){
  return `<article class="task-file-card">
    <button class="task-file-thumb" data-open-asset="${esc(a.id)}">${assetPreviewThumb(a)}</button>
    <div class="task-file-meta">
      <b>${esc(a.title)}</b>
      <span>${esc(assetIcon(a))}${assetSizeLabel(a.size) ? ` / ${esc(assetSizeLabel(a.size))}` : ''}</span>
    </div>
  </article>`;
}

function assetPreviewThumb(a){
  if(a?.kind === 'image' && a.data?.localDataUrl) return `<img src="${esc(a.data.localDataUrl)}" alt="">`;
  return `<span>${esc(assetIcon(a))}</span>`;
}

function taskDetailLinksHtml(loc){
  const links = taskLinks(loc);
  return `<div class="task-links-panel">
    <div class="task-links-head">
      <div>
        <h3>URLs & Links</h3>
        <p>Links found across this task, notes, updates, custom columns, and attached assets.</p>
      </div>
      <span>${links.length} link${links.length===1?'':'s'}</span>
    </div>
    <div class="task-links-list">
      ${links.map(link=>`<a class="task-link-card" href="${esc(link.href)}" target="_blank" rel="noopener noreferrer">
        <span class="task-link-source">${esc(link.source)}</span>
        <strong>${esc(link.label || link.raw)}</strong>
        <em>${esc(link.host)}</em>
      </a>`).join('') || `<div class="task-detail-empty">
        <h3>No links found</h3>
        <p>Add a URL to notes, updates, the description, or a URL column and it will appear here.</p>
      </div>`}
    </div>
  </div>`;
}

function taskLinks(loc){
  const {b,t} = loc;
  const seen = new Set();
  const out = [];
  const add = (href, raw, source, label='')=>{
    if(!href) return;
    const key = href.toLowerCase();
    if(seen.has(key)) return;
    seen.add(key);
    let host = href;
    try{ host = new URL(href).hostname.replace(/^www\./,''); }catch(e){}
    out.push({href, raw, source, label:label || raw, host});
  };
  const addText = (text, source)=>{
    extractLinksFromText(text).forEach(link=>add(link.href, link.raw, source));
  };
  addText(t.name, 'Task name');
  addText(t.notes, 'Notes');
  addText(taskDescription(t), 'Description');
  (b.columns || []).forEach(c=>{
    const value = (t.custom || {})[c.id];
    if(c.type === 'url'){
      const href = normalizedUrlHref(value);
      if(href) add(href, String(value || href), c.name || 'URL column', c.name || String(value || href));
    } else if(c.type !== 'files' && c.type !== 'document'){
      addText(Array.isArray(value) ? value.join(' ') : String(value || ''), c.name || 'Custom column');
    }
  });
  feedUpdates().filter(u=>u.taskId === t.id).forEach(u=>{
    addText(u.body || '', `Update by ${member(u.authorId)?.name || 'TaskTrack'}`);
  });
  taskFileAssets(b,t).forEach(a=>{
    if(a.externalUrl) add(normalizedUrlHref(a.externalUrl), a.externalUrl, 'Attached asset', a.title);
  });
  return out;
}

function taskDetailActivityHtml(loc){
  const {b,g,t} = loc;
  const st = statusOptionFor(b,t.status);
  const pr = priorityOptionFor(b,t.priority);
  const files = taskFileAssets(b,t);
  const updates = feedUpdates().filter(u=>u.taskId === t.id).slice(0,6);
  const rows = [
    {kind:'Status', value:st.label, detail:'Current status', color:st.color},
    {kind:'Priority', value:pr.label, detail:'Current priority', color:pr.color},
    {kind:'Owners', value:taskOwnerNames(t) || 'Unassigned', detail:'Task responsibility'},
    {kind:'Timeline', value:taskTimelineLabel(t) || 'No timeline', detail:'Schedule'},
    {kind:'Files', value:`${files.length} attached`, detail:'Files gallery'},
    ...updates.map(u=>({kind:'Update', value:member(u.authorId)?.name || 'TaskTrack', detail:fmtTimeAgo(u.createdAt), body:u.body || ''}))
  ];
  return `<div class="task-activity-panel">
    ${rows.map(row=>`<div class="activity-row">
      <div class="activity-dot" style="${row.color?`background:${esc(row.color)}`:''}"></div>
      <div>
        <b>${esc(row.kind)}</b>
        <span>${esc(row.value)}</span>
        ${row.body ? `<p>${renderLinkedText(String(row.body).slice(0,160) + (String(row.body).length>160?'...':''))}</p>` : ''}
      </div>
      <em>${esc(row.detail)}</em>
    </div>`).join('')}
  </div>`;
}

function taskDetailDescriptionHtml(loc){
  const links = taskLinks(loc);
  return `<div class="task-description-panel">
    <div class="task-description-card">
      <div class="task-description-head">
        <h3>Item description</h3>
        <button class="btn sm" data-action="save-task-description">Save description</button>
      </div>
      <textarea id="task-detail-description" rows="12" placeholder="Describe the task, context, goals, audience, deliverables, and important links.">${esc(taskDescription(loc.t))}</textarea>
      ${links.length ? `<div class="description-link-preview">
        <span>Detected links</span>
        ${links.slice(0,4).map(link=>`<a href="${esc(link.href)}" target="_blank" rel="noopener noreferrer">${esc(link.label || link.raw)}</a>`).join('')}
      </div>` : ''}
    </div>
  </div>`;
}

function taskDetailEmailsHtml(loc){
  const owners = taskOwners(loc.t);
  const withEmail = owners.filter(o=>o.email);
  const missing = owners.filter(o=>!o.email);
  const subject = `TaskTrack task: ${loc.t.name || 'Task update'}`;
  const body = taskEmailBody(loc);
  const sendLabel = TASK_EMAIL_SENDING ? 'Sending...' : `Send email to all owners${withEmail.length ? ` (${withEmail.length})` : ''}`;
  return `<div class="task-email-panel">
    <div class="email-activity-card">
      <div>
        <h3>Task email</h3>
        <p>Emails send directly through Resend to each owner with an email address.</p>
      </div>
      <button class="btn sm" data-action="send-task-email-all" ${withEmail.length && !TASK_EMAIL_SENDING ? '' : 'disabled'}>${esc(sendLabel)}</button>
    </div>
    <div class="email-template-card">
      <label>Subject<input id="task-email-subject" value="${esc(subject)}"></label>
      <label>Message<textarea id="task-email-body" rows="8">${esc(body)}</textarea></label>
      <div class="email-recipient-list">
        ${withEmail.map(o=>`<span class="email-recipient-chip">${esc(o.name)}${o.email?` <em>${esc(o.email)}</em>`:''}</span>`).join('')}
        ${!withEmail.length ? '<div class="task-detail-empty">Add owner emails in Settings before sending task emails.</div>' : ''}
      </div>
      ${missing.length ? `<div class="integration-note">Missing emails for: ${esc(missing.map(o=>o.name).join(', '))}</div>` : ''}
    </div>
    <div class="email-activity-card subdued">
      <h3>Delivery note</h3>
      <p>One click sends the message to all task owners at once. Each owner receives their own email instead of a shared recipient list.</p>
    </div>
  </div>`;
}

function taskEmailBody(loc){
  const {b,g,t} = loc;
  const st = statusOptionFor(b,t.status).label;
  const pr = priorityOptionFor(b,t.priority).label;
  return [
    `Hi,`,
    '',
    `Here is the current TaskTrack task update for "${t.name}".`,
    '',
    `Board: ${b.name}`,
    `Group: ${g.name}`,
    `Status: ${st}`,
    `Priority: ${pr}`,
    taskTimelineLabel(t) ? `Timeline: ${taskTimelineLabel(t)}` : '',
    t.notes ? `Notes: ${t.notes}` : '',
    taskDescription(t) ? `Description: ${taskDescription(t)}` : '',
    '',
    'Please update TaskTrack with any changes or blockers.'
  ].filter(line=>line !== '').join('\n');
}

async function sendTaskEmailToOwners(){
const sampleLoc=currentTaskDetailLocation();if(!sampleLoc)return;
  const recipients=taskOwners(sampleLoc.t).filter(o=>normalizeEmail(o.email));if(!recipients.length){toast('Add owner emails in Settings first');return;}
  const sampleSubject=($('#task-email-subject')?.value||sampleLoc.t.name).trim(),sampleBody=($('#task-email-body')?.value||taskEmailBody(sampleLoc)).trim();
  if(!sampleSubject||!sampleBody){toast('Add a subject and message');return;}
  recordPortfolioSimulation('task-email',{recipients:recipients.map(o=>o.email),subject:sampleSubject,body:sampleBody});toast('Simulation recorded for '+recipients.length+' owners. No email was sent.');return;

  const loc = currentTaskDetailLocation();
  if(!loc || TASK_EMAIL_SENDING) return;
  const owners = taskOwners(loc.t);
  const withEmail = owners.filter(o=>normalizeEmail(o.email));
  if(!withEmail.length){ toast('Add owner emails in Settings before sending'); return; }
  if(!supabase || !SESSION?.user || !WORKSPACE_ID){ toast('Sign in before sending task emails'); return; }

  const subject = ($('#task-email-subject')?.value || `TaskTrack task: ${loc.t.name || 'Task update'}`).trim();
  const message = ($('#task-email-body')?.value || taskEmailBody(loc)).trim();
  if(!subject){ toast('Add a subject before sending'); return; }
  if(!message){ toast('Add a message before sending'); return; }

  TASK_EMAIL_SENDING = true;
  render();
  const { data, error } = await supabase.functions.invoke('task-email', {
    body:{
      workspaceId:WORKSPACE_ID,
      taskId:loc.t.id,
      subject,
      message
    }
  });
  TASK_EMAIL_SENDING = false;

  if(error || data?.error){
    toast(await assistantFunctionErrorMessage(error, data) || 'Could not send task email');
    render();
    return;
  }

  const sent = Number(data?.sent || 0);
  const failed = Number(data?.failed || 0);
  const skipped = Number(data?.skipped || 0);
  toast(failed ? `Sent ${sent}; ${failed} failed` : `Sent email to ${sent} owner${sent===1?'':'s'}${skipped?`; ${skipped} missing email${skipped===1?'':'s'}`:''}`);
  void fetchWorkspaceState();
  render();
}

function ensureTaskFilesColumn(loc){
  const col = ensureBoardColumnByType(loc.b,'files','Files');
  return col;
}

function openTaskAttachExistingFile(anchor){
  const loc = currentTaskDetailLocation();
  if(!loc) return;
  const attached = new Set(taskFileAssets(loc.b, loc.t).map(a=>a.id));
  const available = (DATA.assets || []).filter(a=>!attached.has(a.id));
  const html = `<div class="asset-picker">
    <div class="control-title">Attach from Docs Hub</div>
    ${available.length ? available.slice(0,24).map(a=>`<button class="picker-row-main full" data-task-attach-asset="${esc(a.id)}">
      <span class="asset-type">${esc(assetIcon(a))}</span><span>${esc(a.title)}</span><em>${esc(assetSizeLabel(a.size))}</em>
    </button>`).join('') : '<div class="picker-empty">No unattached files in Docs Hub.</div>'}
  </div>`;
  const dd = buildDD(anchor, html);
  dd.querySelectorAll('[data-task-attach-asset]').forEach(btn=>btn.onclick=()=>{
    const col = ensureTaskFilesColumn(loc);
    attachAssetsToTaskCell(loc.g.id, loc.t.id, col.id, [btn.dataset.taskAttachAsset], loc.b.id);
  });
}

function insertTaskMention(memberId){
  const input = $('#task-detail-update-body');
  if(!input) return;
  const label = memberId === EVERYONE_MENTION_ID ? 'Everyone' : member(memberId)?.name;
  if(!label) return;
  const prefix = input.value && !/\s$/.test(input.value) ? ' ' : '';
  input.value = `${input.value}${prefix}@${label} `;
  input.focus();
}

function postTaskDetailUpdate(){
  const loc = currentTaskDetailLocation();
  const body = $('#task-detail-update-body')?.value.trim();
  if(!loc || !body){ toast('Write an update first'); return; }
  recordUpdate({body,boardId:loc.b.id,groupId:loc.g.id,taskId:loc.t.id});
  save();
  TASK_DETAIL_TAB = 'updates';
  render();
  toast('Update posted');
}

function saveTaskDetailDescription(){
  const loc = currentTaskDetailLocation();
  if(!loc) return;
  setTaskDescription(loc.t, $('#task-detail-description')?.value || '');
  save();
  render();
  toast('Description saved');
}

function saveTaskDetailName(){
  const loc = currentTaskDetailLocation();
  const value = $('#app [data-task-detail-name]')?.value.trim();
  if(!loc || !value || value === loc.t.name) return;
  loc.t.name = value;
  save();
  render();
}

function wireTaskDetailDrawer(){
  if(!TASK_DETAIL) return;
  document.querySelectorAll('[data-action="close-task-detail"]').forEach(el=>el.onclick=closeTaskDetail);
  document.querySelectorAll('[data-task-detail-tab]').forEach(el=>el.onclick=()=>{ TASK_DETAIL_TAB=el.dataset.taskDetailTab; render(); });
  document.querySelectorAll('[data-task-mention]').forEach(el=>el.onclick=()=>insertTaskMention(el.dataset.taskMention));
  document.querySelectorAll('[data-action="task-post-update"]').forEach(el=>el.onclick=postTaskDetailUpdate);
  document.querySelectorAll('[data-action="save-task-description"]').forEach(el=>el.onclick=saveTaskDetailDescription);
  document.querySelectorAll('[data-action="send-task-email-all"]').forEach(el=>el.onclick=sendTaskEmailToOwners);
  document.querySelectorAll('[data-action="task-upload-file"]').forEach(el=>el.onclick=()=>$('#task-detail-file-input')?.click());
  document.querySelectorAll('[data-action="task-attach-existing-file"]').forEach(el=>el.onclick=e=>{ e.stopPropagation(); openTaskAttachExistingFile(e.currentTarget); });
  document.querySelectorAll('[data-open-asset]').forEach(el=>el.onclick=e=>{ e.stopPropagation(); void openAsset(el.dataset.openAsset); });
  document.querySelectorAll('[data-owner-pick]').forEach(el=>el.onclick=e=>{ e.stopPropagation(); pickOwner(el); });
  document.querySelectorAll('[data-task-color-pick]').forEach(el=>el.onclick=e=>{ e.stopPropagation(); pickTaskColor(el); });
  document.querySelectorAll('[data-task-recurring]').forEach(el=>el.onclick=e=>{ e.stopPropagation(); pickTaskRecurrence(el); });
  const fileInput = $('#task-detail-file-input');
  if(fileInput) fileInput.onchange=()=>{
    const loc = currentTaskDetailLocation();
    if(!loc) return;
    const col = ensureTaskFilesColumn(loc);
    void uploadAssetFiles(Array.from(fileInput.files || []), {attachTo:{boardId:loc.b.id,gid:loc.g.id,tid:loc.t.id,cid:col.id}});
  };
  const title = $('[data-task-detail-name]');
  if(title){
    title.onblur=saveTaskDetailName;
    title.onkeydown=e=>{
      if(e.key === 'Enter'){ e.preventDefault(); title.blur(); }
      if(e.key === 'Escape'){ e.preventDefault(); render(); }
    };
  }
}

/* ============================================================
   VIEW: Board (the core table)
   ============================================================ */
function viewBoard(){
  const b = board(CURRENT_BOARD);
  if(!b) return `<div class="empty-state"><div class="big">📋</div>Select or create a board from the left.</div>`;
  normalizeBoard(b);
  const q = SEARCH.trim().toLowerCase();
  const groupsHtml = BOARD_LAYOUT === 'kanban' ? viewKanbanBoard(b,q) : b.groups.map(g=>groupTable(b,g,q)).join('');
  const selectedLocs = selectedTaskLocations().filter(x=>x.b.id===b.id);
  const selected = selectedLocs.length;
  const selectedHidden = selectedLocs.length > 0 && selectedLocs.every(({t})=>taskIsHidden(t));
  const filterSummary = boardFilterSummary(b,q);
  return `
    <div class="toolbar board-toolbar">
      <div class="search-box">${I.search}<input id="board-search" placeholder="Search tasks" value="${esc(SEARCH)}"></div>
      <div class="layout-toggle">
        <button class="${BOARD_LAYOUT==='table'?'active':''}" data-board-layout="table">${esc(tr('table'))}</button>
        <button class="${BOARD_LAYOUT==='kanban'?'active':''}" data-board-layout="kanban">${esc(tr('kanban'))}</button>
      </div>
      <button class="tool-btn ${boardPersonFilter(b)?'active':''}" data-board-control="person">${I.mention}<span>${esc(tr('person'))}</span></button>
      <button class="tool-btn ${b.sort?.columnId?'active':''}" data-board-control="sort">${I.sort}<span>${esc(tr('sort'))}</span></button>
      <button class="tool-btn ${b.hiddenColumns?.length?'active':''}" data-board-control="hide">${I.hide}<span>${esc(tr('hide'))}</span></button>
      <button class="tool-btn ${b.calendarShowTimeline?'active':''}" data-board-control="calendar">${I.cal}<span>Calendar</span></button>
    </div>
    ${boardViewBar()}
    ${filterSummary.length ? `<div class="board-filter-banner">
      <span>Showing filtered results: ${esc(filterSummary.join(' / '))}</span>
      <button data-action="clear-board-filters">Clear filters</button>
    </div>` : ''}
    ${groupsHtml}
    ${BOARD_LAYOUT === 'table' ? `<button class="add-group-btn" data-action="add-group"><span>${I.plus}</span> Add new group</button>` : ''}
    ${selected?selectedActionBar(selected, selectedHidden):''}`;
}

function viewKanbanBoard(b,q=''){
  const statusEntries = Object.entries(boardStatusOptions(b));
  const tasks = b.groups.flatMap(g=>taskTreeList(g.tasks).map(t=>({b,g,t})))
    .filter(x=>taskMatchesActiveFilters(b,x.t,q));
  const byStatus = new Map(statusEntries.map(([key])=>[key, []]));
  tasks.forEach(x=>{
    const key = byStatus.has(x.t.status) ? x.t.status : 'not_started';
    byStatus.get(key).push(x);
  });
  return `<div class="kanban-board">
    ${statusEntries.map(([key,opt])=>`
      <section class="kanban-lane" data-kanban-status="${esc(key)}">
        <div class="kanban-lane-head">
          <span class="kanban-dot" style="background:${esc(opt.color)}"></span>
          <strong>${esc(opt.label)}</strong>
          <em>${(byStatus.get(key) || []).length}</em>
        </div>
        <div class="kanban-card-list">
          ${(byStatus.get(key) || []).map(({g,t})=>kanbanTaskCard(b,g,t)).join('') || '<div class="kanban-empty">Drop tasks here</div>'}
        </div>
      </section>`).join('')}
  </div>`;
}

function kanbanTaskCard(b,g,t){
  const pr = priorityOptionFor(b,t.priority);
  const childCount = childTasks(t).length;
  return `<article class="kanban-card" draggable="true" data-kanban-drag="${esc(t.id)}" data-open-kanban-task="${esc(b.id)}|${esc(t.id)}">
    <div class="kanban-card-title">${esc(t.name)} <span data-presence-target="task|${esc(t.id)}">${presenceBadgesHtml('task', t.id)}</span></div>
    <div class="kanban-card-meta">
      <span>${esc(g.name)}</span>
      ${taskTimelineLabel(t)?`<span>${esc(taskTimelineLabel(t))}</span>`:''}
      ${childCount?`<span>${childCount} subtask${childCount===1?'':'s'}</span>`:''}
    </div>
    <div class="kanban-card-foot">
      <span class="kanban-priority" style="background:${esc(pr.color)}">${esc(pr.label)}</span>
      ${ownerStackHtml(taskOwnerIds(t), {pickKey:`${t.id}|${g.id}`, limit:2, small:true})}
    </div>
  </article>`;
}

function taskSearchText(t,b=null){
  const custom = b ? (b.columns||[]).map(c=>customValueText(t,c)).join(' ') : Object.values(visibleTaskCustom(t)).join(' ');
  const owner = taskOwnerNames(t);
  const assigner = member(t.assignedBy)?.name || '';
  return [
    t.name,
    t.notes,
    owner,
    assigner,
    t.status,
    b ? statusOptionFor(b,t.status).label : STATUS[t.status]?.label,
    t.priority,
    b ? priorityOptionFor(b,t.priority).label : PRIORITY[t.priority]?.label,
    t.due,
    t.tlStart,
    t.tlEnd,
    t.budget,
    custom
  ].filter(v=>v!=null).join(' ').toLowerCase();
}
function taskMatchesSearch(t,q,b=null){ return !q || taskSearchText(t,b).includes(q); }
function customValueText(t,c){
  const val = t?.custom?.[c.id];
  if(c.type==='files') return parseIdList(val).map(id=>asset(id)?.title || id).join(' ');
  if(c.type==='document') return doc(parseIdList(val)[0])?.title || '';
  if(c.type==='recurrence') return recurrenceLabel(val);
  return Array.isArray(val) ? val.join(' ') : String(val || '');
}

function taskMatchesPersonFilter(b,t){
  const personFilter = boardPersonFilter(b);
  return !personFilter || taskOwnerIds(t).includes(personFilter) || t.assignedBy === personFilter;
}

function taskMatchesActiveFilters(b,t,q=''){
  return matchesBoardView(t) && taskMatchesPersonFilter(b,t) && (!q || taskMatchesSearch(t,q,b));
}

function taskTreeMatchesActiveFilters(b,t,q=''){
  if(BOARD_VIEW !== 'hidden' && taskIsHidden(t)) return false;
  return taskMatchesActiveFilters(b,t,q) || childTasks(t).some(st=>taskTreeMatchesActiveFilters(b,st,q));
}

function hasActiveTaskFilter(b,q=''){
  return Boolean(q || BOARD_VIEW !== 'all' || boardPersonFilter(b));
}

function visibleTaskCountForTree(b,t,q=''){
  if(!taskTreeMatchesActiveFilters(b,t,q)) return 0;
  const activeFilter = hasActiveTaskFilter(b,q);
  const children = sortTasksForBoard(b, childTasks(t)).filter(st=>!activeFilter || taskTreeMatchesActiveFilters(b,st,q));
  return 1 + children.reduce((sum,st)=>sum+visibleTaskCountForTree(b,st,q),0);
}

function renderedTaskIdsForTree(b,t,q=''){
  if(!taskTreeMatchesActiveFilters(b,t,q)) return [];
  const activeFilter = hasActiveTaskFilter(b,q);
  const children = sortTasksForBoard(b, childTasks(t)).filter(st=>!activeFilter || taskTreeMatchesActiveFilters(b,st,q));
  return [t.id, ...children.flatMap(st=>renderedTaskIdsForTree(b,st,q))];
}

function taskTreeBudget(t){
  return (Number(t.budget)||0) + childTasks(t).reduce((sum,st)=>sum+taskTreeBudget(st),0);
}

function sortTasksForBoard(b,tasks){
  const sort = b.sort || {};
  if(!sort.columnId) return tasks;
  const col = allBoardColumns(b).find(c=>c.id===sort.columnId);
  if(!col) return tasks;
  const dir = sort.direction === 'desc' ? -1 : 1;
  return tasks.map((task,index)=>({task,index})).sort((a,b2)=>{
    const cmp = compareTaskSortValues(taskSortValue(b,a.task,col), taskSortValue(b,b2.task,col));
    return cmp ? cmp * dir : a.index - b2.index;
  }).map(x=>x.task);
}

function compareTaskSortValues(a,b){
  if(typeof a === 'number' || typeof b === 'number') return (Number(a)||0) - (Number(b)||0);
  return String(a||'').localeCompare(String(b||''), undefined, {numeric:true,sensitivity:'base'});
}

function taskSortValue(b,t,c){
  if(c.id==='name') return t.name || '';
  if(c.id==='owner') return taskOwnerNames(t);
  if(c.id==='status') return statusOptionFor(b,t.status).label || '';
  if(c.id==='color') return taskColorLabel(t);
  if(c.id==='due') return validIsoDate(t.due) || '';
  if(c.id==='timeline') return validIsoDate(t.tlStart) || validIsoDate(t.tlEnd) || '';
  if(c.id==='priority') return priorityRank(b,t.priority);
  if(c.id==='budget') return Number(t.budget)||0;
  if(c.id==='notes') return t.notes || '';
  if(c.id==='assignedBy') return member(t.assignedBy)?.name || '';
  const val = (t.custom||{})[c.id];
  if(c.type==='number') return Number(String(val||'').replace(/[^0-9.\-]/g,'')) || 0;
  if(c.type==='date') return val || '';
  if(c.type==='status') return statusOptionFor(b,val).label || '';
  if(c.type==='recurrence') return recurrenceLabel(val || taskRecurrence(t,b));
  if(c.type==='files') return parseIdList(val).map(id=>asset(id)?.title || id).join(' ');
  if(c.type==='document') return doc(parseIdList(val)[0])?.title || '';
  return val || '';
}

function priorityRank(b,key){
  const builtIn = Object.keys(PRIORITY);
  const idx = builtIn.indexOf(key);
  if(idx >= 0) return idx;
  const custom = Object.keys(b.priorityLabels||{}).indexOf(key);
  return custom >= 0 ? builtIn.length + custom : -1;
}

function selectedActionBar(count, selectedHidden=false){
  const hideAction = selectedHidden
    ? `<button class="selection-action" data-bulk-action="unhide">${I.hide}<span>Unhide</span></button>`
    : `<button class="selection-action" data-bulk-action="hide">${I.hide}<span>Hide</span></button>`;
  return `<div class="selection-bar">
    <div class="selection-count"><span class="bubble">${count}</span><span>${count} task${count===1?'':'s'} selected</span></div>
    <button class="selection-action" data-bulk-action="duplicate">${I.copy}<span>Duplicate</span></button>
    <button class="selection-action" data-bulk-action="export">${I.download}<span>Export</span></button>
    ${hideAction}
    <button class="selection-action" data-bulk-action="archive">${I.archive}<span>Archive</span></button>
    <button class="selection-action" data-bulk-action="delete">${I.trash}<span>Delete</span></button>
    <button class="selection-action" data-bulk-action="convert">${I.convert}<span>Convert</span></button>
    <button class="selection-action" data-bulk-action="move">${I.move}<span>Move to</span></button>
    <button class="selection-close" data-action="clear-selection" title="Clear selection">${I.close}</button>
  </div>`;
}

function groupTable(b,g,q){
  let tasks = g.tasks;
  tasks = tasks.filter(t => taskTreeMatchesActiveFilters(b,t,q));
  tasks = sortTasksForBoard(b,tasks);
  const budgetSum = g.tasks.reduce((s,t)=>s+taskTreeBudget(t),0);
  const cols = boardColumns(b);
  const visibleIds = tasks.flatMap(t=>renderedTaskIdsForTree(b,t,q));
  const headers = cols.map(c=>`
          <th class="board-col ${c.custom?'custom-col':''}" draggable="true" data-col-id="${esc(c.id)}" data-col-drag="${c.id}" style="${columnWidthStyle(b,c.id)}" title="Drag to reorder">
            <span class="col-head"><span class="col-grip">::</span><span class="col-head-name">${esc(c.label)}</span>
              <button class="col-menu" data-col-menu="${c.id}" title="Column options">${I.dots}</button>
              <span class="col-resizer" data-col-resize="${esc(c.id)}" title="Resize column"></span>
            </span>
          </th>`).join('');
  const rows = tasks.map(t=>taskRows(b,g,t,q)).join('');
  const open = g.open!==false;
  const taskCount = taskIdsInGroup(g).length;
  const visibleTaskCount = tasks.reduce((sum,t)=>sum+visibleTaskCountForTree(b,t,q),0);
  const countLabel = hasActiveTaskFilter(b,q) && visibleTaskCount !== taskCount
    ? `${visibleTaskCount} of ${taskCount} shown`
    : `${taskCount} ${taskCount===1?'task':'tasks'}`;
  const allVisibleSelected = visibleIds.length > 0 && visibleIds.every(id=>SELECTED_TASKS.has(id));
  return `
  <div class="group" data-group="${g.id}">
    <div class="group-head">
      <button class="group-drag-handle" draggable="true" data-group-drag="${g.id}" title="Drag group">${I.move}</button>
      <button class="group-toggle" data-toggle-group="${g.id}" style="transform:rotate(${open?90:0}deg)">▸</button>
      <span class="group-title" style="color:${g.color}" contenteditable data-rename-group="${g.id}">${esc(g.name)}</span>
      <span class="group-count">${esc(countLabel)}</span>
      <span class="group-budget">Total budget: ${fmtMoney(budgetSum)}</span>
      <button class="group-menu" data-del-group="${g.id}" title="Delete group">${I.dots}</button>
    </div>
    <div class="tbl-wrap" ${open?'':'style="display:none"'}>
      <table class="board">
        <thead><tr>
          <th class="gcolor" style="background:${g.color}"></th>
          <th class="col-mini select-cell">
            <label class="row-check-target" title="Select all visible tasks">
              <input type="checkbox" class="row-select-all" data-select-group="${g.id}" ${allVisibleSelected?'checked':''}>
              <span class="checkmark"></span>
            </label>
          </th>
          ${headers}
          <th class="add-column-head"><button data-action="add-column" title="Add column">${I.plus}</button></th>
        </tr></thead>
        <tbody data-drop-group="${g.id}">
          ${rows}
          <tr><td class="gcolor" style="background:${g.color}"></td>
            <td colspan="${cols.length+2}"><div class="row-add" data-add-task="${g.id}">+ Add task</div></td></tr>
        </tbody>
      </table>
    </div>
  </div>`;
}

function taskRows(b,g,t,q=''){
  return taskRowsRecursive(b,g,t,q,0,'');
}

function taskRowsRecursive(b,g,t,q='',depth=0,parentId=''){
  const activeFilter = hasActiveTaskFilter(b,q);
  const children = sortTasksForBoard(b, childTasks(t)).filter(st=>!activeFilter || taskTreeMatchesActiveFilters(b,st,q));
  const open = areSubtasksOpen(t);
  const canShowChildren = children.length && (open || activeFilter);
  const main = taskRow(b,g,t,depth,parentId,children.length);
  const subtasks = canShowChildren ? children.map(st=>taskRowsRecursive(b,g,st,q,depth+1,t.id)).join('') : '';
  return main + subtasks;
}

function taskRow(b,g,t,depth=0,parentId='',childCount=0){
  const cells = boardColumns(b).map(c=>taskCell(b,g,t,c,depth,parentId,childCount)).join('');
  const isSubtask = depth > 0;
  const color = taskColor(t);
  const colorStyle = color
    ? `style="--task-color:${esc(color)};--task-color-bg:${esc(rgbaFromHex(color,0.16))};--task-color-hover-bg:${esc(rgbaFromHex(color,0.22))};--task-color-selected-bg:${esc(rgbaFromHex(color,0.28))}"`
    : '';
  return `
  <tr class="task-row ${isSubtask?'subtask-row':''} ${taskIsHidden(t)?'hidden-task-row':''} ${color?'has-task-color':''} ${SELECTED_TASKS.has(t.id)?'selected':''}" ${colorStyle} draggable="true" data-task="${t.id}" data-group="${g.id}" data-depth="${depth}" ${parentId?`data-parent="${parentId}"`:''}>
    <td class="gcolor" style="background:${g.color}"></td>
    <td class="col-mini select-cell">
      <label class="row-check-target" title="Select task">
        <input type="checkbox" class="row-check" data-check="${t.id}" ${SELECTED_TASKS.has(t.id)?'checked':''}>
        <span class="checkmark"></span>
      </label>
    </td>
    ${cells}
    <td class="add-column-cell"></td>
  </tr>`;
}

function taskCell(b,g,t,c,depth=0,parentId='',childCount=0){
  const assigner = member(t.assignedBy);
  const st = statusOptionFor(b,t.status);
  const pr = priorityOptionFor(b,t.priority);
  const tl = taskTimelineLabel(t);
  const style = columnWidthStyle(b,c.id);
  const isSubtask = depth > 0;
  if(c.id==='name') return `<td class="task-name" data-col-cell="name" style="${style}"><div class="cell task-name-cell ${isSubtask?'subtask-name-cell':''}" style="--depth:${depth}">
      ${childCount?`<button class="subtask-toggle ${areSubtasksOpen(t)?'open':''}" data-toggle-subtasks="${g.id}|${t.id}" title="${areSubtasksOpen(t)?'Hide subtasks':'Show subtasks'}">▸</button>`:'<span class="subtask-spacer"></span>'}
      ${taskUpdateIndicatorHtml(b,g,t)}
      <input class="cell-input task-name-input" data-edit="name" data-task="${esc(t.id)}" data-group="${esc(g.id)}" value="${esc(t.name)}" placeholder="Task name" title="${esc(t.name)}">
      <span data-presence-target="task|${esc(t.id)}">${presenceBadgesHtml('task', t.id)}</span>
      <button class="subtask-add" data-add-subtask="${g.id}|${t.id}" title="Add subtask">+ Subtask</button>
    </div></td>`;
  if(c.id==='owner') return `<td class="sel-cell" data-col-cell="owner" style="${style}">
      <div class="owner-cell">
        ${ownerStackHtml(taskOwnerIds(t), {pickKey:`${t.id}|${g.id}`, limit:3})}
      </div>
    </td>`;
  if(c.id==='status') return `<td class="sel-cell" data-col-cell="status" style="${style}"><button class="pill" style="background:${st.color}" data-status-pick="${t.id}|${g.id}">${esc(st.label)}</button></td>`;
  if(c.id==='color'){
    const color = taskColor(t);
    return `<td class="sel-cell" data-col-cell="color" style="${style}">
      <button type="button" class="task-color-cell ${color?'active':''}" ${color?`style="--task-color:${esc(color)};--task-color-bg:${esc(rgbaFromHex(color,0.14))}"`:''} data-task-color-pick="${esc(t.id)}|${esc(g.id)}" title="${color?esc(color):'Set task color'}">
        <i></i><span>${esc(color ? taskColorLabel(t) : 'Set color')}</span>
      </button>
    </td>`;
  }
  if(c.id==='due') return `<td class="sel-cell" data-col-cell="due" style="${style}"><div class="date-cell ${validIsoDate(t.due)?'':'empty'}" data-date-pick="${t.id}|${g.id}|due">${I.cal}<span>${taskDueLabel(t)||'Set due date'}</span></div></td>`;
  if(c.id==='timeline') return `<td class="sel-cell" data-col-cell="timeline" style="${style}">
      ${tl?`<div style="padding:8px 12px"><div class="timeline-bar" data-tl-pick="${t.id}|${g.id}">${tl}</div></div>`
          :`<div class="date-cell empty" data-tl-pick="${t.id}|${g.id}">${I.cal}<span>Set timeline</span></div>`}
    </td>`;
  if(c.id==='priority') return `<td class="sel-cell" data-col-cell="priority" style="${style}"><button class="pill" style="background:${pr.color}" data-priority-pick="${t.id}|${g.id}">${esc(pr.label)}</button></td>`;
  if(c.id==='budget') return `<td class="budget-cell" data-col-cell="budget" style="${style}"><div class="cell"><input class="cell-input" data-edit="budget" data-task="${t.id}" data-group="${g.id}" value="${esc(t.budget||'')}" placeholder="0" style="text-align:right"></div></td>`;
  if(c.id==='notes') return `<td data-col-cell="notes" style="${style}"><div class="cell link-aware-cell has-copy"><input class="cell-input" data-edit="notes" data-task="${t.id}" data-group="${g.id}" value="${esc(t.notes||'')}" placeholder="Add note">${fieldLinkOpenHtml(t.notes)}${copyCellButton(g,t,'notes',t.notes)}</div></td>`;
  if(c.id==='assignedBy') return `<td class="sel-cell" data-col-cell="assignedBy" style="${style}">
      <div class="assigned-by">${assigner?`<div class="owner-av" style="background:${assigner.color}">${esc(assigner.initials)}</div><span>${esc(assigner.name)}</span>`:'<span>Not set</span>'}</div>
    </td>`;
  return customTaskCell(b,g,t,c);
}

function customTaskCell(b,g,t,c){
  const val = (t.custom||{})[c.id] || '';
  const style = columnWidthStyle(b,c.id);
  if(c.type==='files'){
    const ids = taskAssetIdsForColumn(t,c.id);
    const chips = ids.map(id=>asset(id)).filter(Boolean).slice(0,3).map(a=>
      `<button class="asset-chip" data-open-asset="${a.id}" title="${esc(a.title)}"><span>${esc(assetIcon(a))}</span>${esc(a.title)}</button>`
    ).join('');
    const more = ids.length > 3 ? `<span class="asset-more">+${ids.length - 3}</span>` : '';
    return `<td class="custom-cell file-cell" data-col-cell="${esc(c.id)}" style="${style}">
      <div class="asset-cell">
        ${chips || `<button class="asset-empty" data-file-cell-pick="${t.id}|${g.id}|${c.id}">${I.plus}<span>Add file</span></button>`}
        ${chips ? `${more}<button class="asset-add-mini" data-file-cell-pick="${t.id}|${g.id}|${c.id}" title="Attach files">${I.plus}</button>` : ''}
      </div>
    </td>`;
  }
  if(c.type==='document'){
    const did = taskDocIdForColumn(t,c.id);
    const d = did ? doc(did) : null;
    return `<td class="custom-cell doc-cell" data-col-cell="${esc(c.id)}" style="${style}">
      <div class="asset-cell">
        ${d ? `<button class="doc-chip" data-open-doc="${d.id}" title="${esc(d.title)}">${I.docs}<span>${esc(d.title)}</span></button>
          <button class="asset-add-mini" data-doc-cell-pick="${t.id}|${g.id}|${c.id}" title="Change document">${I.dots}</button>`
        : `<button class="asset-empty" data-doc-cell-pick="${t.id}|${g.id}|${c.id}">${I.plus}<span>Add doc</span></button>`}
      </div>
    </td>`;
  }
  if(c.type==='date'){
    return `<td class="sel-cell custom-cell" data-col-cell="${esc(c.id)}" style="${style}"><div class="date-cell ${val?'':'empty'}" data-custom-date-pick="${t.id}|${g.id}|${c.id}">${I.cal}<span>${val?fmtDate(val):'Set date'}</span></div></td>`;
  }
  if(c.type==='status'){
    const st = val ? statusOptionFor(b,val) : null;
    return `<td class="sel-cell custom-cell" data-col-cell="${esc(c.id)}" style="${style}"><button class="pill ${st?'':'empty'}" style="background:${st?st.color:'#c4c4c4'}" data-custom-status-pick="${t.id}|${g.id}|${c.id}">${esc(st?st.label:'Set status')}</button></td>`;
  }
  if(c.type==='url'){
    const href = normalizedUrlHref(val);
    return `<td class="custom-cell url-custom-cell" data-col-cell="${esc(c.id)}" style="${style}">
      <div class="cell url-cell">
        <input class="cell-input url-input" data-edit="custom:${c.id}" data-task="${esc(t.id)}" data-group="${esc(g.id)}" value="${esc(val)}" placeholder="Paste link">
        ${href ? `<a class="url-open" href="${esc(href)}" target="_blank" rel="noopener noreferrer" title="${esc(href)}">Open</a>` : `<span class="url-open disabled">Open</span>`}
      </div>
    </td>`;
  }
  if(c.type==='recurrence'){
    const recurrence = normalizeRecurrenceValue(val || taskRecurrence(t,b));
    const opt = recurrenceOptionFor(recurrence);
    return `<td class="sel-cell custom-cell" data-col-cell="${esc(c.id)}" style="${style}">
      <div class="recurrence-cell-wrap">
        <button class="recurrence-pill ${recurrence !== 'off' ? 'active' : ''}" style="--recurrence-color:${esc(opt.color)}" title="${esc(opt.label)}" data-custom-recurrence-pick="${esc(t.id)}|${esc(g.id)}|${esc(c.id)}">
          ${esc(opt.label)}
        </button>
      </div>
    </td>`;
  }
  const align = c.type==='number' ? 'text-align:right' : '';
  const placeholder = c.type==='number' ? '0' : `Add ${c.name.toLowerCase()}`;
  return `<td class="custom-cell" data-col-cell="${esc(c.id)}" style="${style}"><div class="cell link-aware-cell ${c.type==='text'?'has-copy':''}"><input class="cell-input" data-edit="custom:${c.id}" data-task="${t.id}" data-group="${g.id}" value="${esc(val)}" placeholder="${esc(placeholder)}" style="${align}">${c.type==='text' ? `${fieldLinkOpenHtml(val)}${copyCellButton(g,t,`custom:${c.id}`,val)}` : ''}</div></td>`;
}

/* ============================================================
   VIEW: Calendar
   ============================================================ */
function clearGoogleCalendarState(){
  PRIVATE_GOOGLE_EVENTS = [];
  SHARED_GOOGLE_EVENTS = [];
  SHARED_GOOGLE_READS = new Set();
  PRIVATE_GOOGLE_SHARED_IDS = new Set();
  MY_GOOGLE_CALENDAR_ON = false;
  PRIVATE_GOOGLE_LOAD_KEY = '';
  PRIVATE_GOOGLE_LOADING = false;
  PRIVATE_GOOGLE_ERROR = '';
  SHARED_GOOGLE_LOAD_KEY = '';
  SHARED_GOOGLE_LOADING = false;
}

function renderSharedCalendarUpdateCard(u){
  const author = member(u.authorId) || {name:'A teammate',initials:'TM',color:'#15856d'};
  const unread = !(u.readBy||[]).includes(currentMemberId());
  const span = googleEventCalendarSpan(u.event);
  const when = span ? `${fmtDate(span.start)}${span.startTime ? ` ${span.startTime}` : ' · All day'}${span.end !== span.start ? ` – ${fmtDate(span.end)}` : ''}` : '';
  return `<article class="feed-card shared-calendar-feed ${unread?'unread':''}" data-update-id="${esc(u.id)}">
    <button class="feed-read-toggle ${unread?'unread':''}" data-update-read="${esc(u.id)}" title="${unread?'Mark read':'Mark unread'}">${unread?'':'✓'}</button>
    <div class="feed-author"><div class="owner-av" style="background:${esc(author.color)}">${esc(author.initials)}</div>
      <div><b>${esc(author.name)}</b><span>${fmtTimeAgo(u.createdAt)}</span></div></div>
    <div class="feed-body"><b>${u.canceled?'Canceled workspace event':'Shared a Google Calendar event'}</b><br>${esc(u.title)}</div>
    <div class="feed-card-meta"><span>${esc(when)}</span>${u.event.location?`<span>${esc(u.event.location)}</span>`:''}${u.event.sync_status==='stale'?`<span>${u.event.sync_error==='reconnect'?'Waiting for Google reconnect':'Google sync delayed'}</span>`:''}</div>
    ${u.event.description ? `<div class="shared-calendar-feed-description">${renderLinkedText(u.event.description,{lineBreaks:true})}</div>` : ''}
    ${u.canceled
      ? (u.authorId===currentMemberId()?`<div class="feed-card-actions"><button data-unshare-feed-google="${esc(u.sharedEventId)}">Unshare from workspace</button></div>`:'')
      : `<div class="feed-card-actions"><button data-open-shared-google="${esc(u.sharedEventId)}">View in workspace calendar</button></div>`}
  </article>`;
}

function openSharedGoogleEvent(eventId){
  const row = SHARED_GOOGLE_EVENTS.find(event=>event.id===eventId && event.sync_status!=='canceled');
  if(!row) return;
  CAL_MONTH = googleEventCalendarSpan(row).start.slice(0,7);
  navigateToView('calendar');
  requestAnimationFrame(()=>{
    const target = [...document.querySelectorAll('[data-cal-google]')]
      .find(el=>el.dataset.calGoogle===`shared:${eventId}`);
    if(target) showGoogleCalendarEventDetail(`shared:${eventId}`,target);
  });
}

async function loadWorkspaceSharedCalendar({force=false}={}){
  if(!supabase || !WORKSPACE_ID || !SESSION?.user || SHARED_GOOGLE_LOADING) return;
  const workspaceId = WORKSPACE_ID;
  if(!force && SHARED_GOOGLE_LOAD_KEY === workspaceId) return;
  SHARED_GOOGLE_LOADING = true;
  try{
    const userId = SESSION.user.id;
    const [{data,error},{data:reads,error:readsError}] = await Promise.all([
      supabase.from('app_shared_calendar_events').select('*').eq('workspace_id',workspaceId).order('shared_at',{ascending:false}).limit(2000),
      supabase.from('app_shared_calendar_event_reads').select('event_id').eq('workspace_id',workspaceId).eq('user_id',userId)
    ]);
    if(error || readsError) throw error || readsError;
    if(WORKSPACE_ID !== workspaceId) return;
    if(SESSION?.user?.id !== userId) return;
    const nextReads = new Set((reads||[]).map(row=>row.event_id));
    const changed = JSON.stringify(SHARED_GOOGLE_EVENTS) !== JSON.stringify(data || [])
      || JSON.stringify([...SHARED_GOOGLE_READS].sort()) !== JSON.stringify([...nextReads].sort());
    SHARED_GOOGLE_EVENTS = data || [];
    SHARED_GOOGLE_READS = nextReads;
    SHARED_GOOGLE_LOAD_KEY = workspaceId;
    if(changed && DATA) render();
  }catch(error){
    console.warn('Could not load shared Google events',error);
  }finally{
    SHARED_GOOGLE_LOADING = false;
  }
}

async function loadMyGoogleCalendar(windowStart,windowEnd){
  if(!MY_GOOGLE_CALENDAR_ON || !integrationConnection('google_calendar') || !supabase || !WORKSPACE_ID || !SESSION?.user || PRIVATE_GOOGLE_LOADING) return;
  const key = `${WORKSPACE_ID}|${SESSION.user.id}|${windowStart}|${windowEnd}`;
  if(PRIVATE_GOOGLE_LOAD_KEY === key) return;
  PRIVATE_GOOGLE_LOAD_KEY = key;
  PRIVATE_GOOGLE_LOADING = true;
  PRIVATE_GOOGLE_ERROR = '';
  try{
    const {data,error} = await supabase.functions.invoke('google-calendar-events',{body:{
      workspaceId:WORKSPACE_ID,
      timeMin:`${isoDateAdd(windowStart,-1)}T00:00:00Z`,
      timeMax:`${isoDateAdd(windowEnd,2)}T00:00:00Z`
    }});
    if(error || data?.error) throw new Error(await assistantFunctionErrorMessage(error,data));
    if(key !== PRIVATE_GOOGLE_LOAD_KEY) return;
    PRIVATE_GOOGLE_EVENTS = data?.events || [];
    PRIVATE_GOOGLE_SHARED_IDS = new Set(data?.sharedEventIds || []);
  }catch(error){
    if(key === PRIVATE_GOOGLE_LOAD_KEY) PRIVATE_GOOGLE_ERROR = error?.message || 'Could not load Google Calendar';
  }finally{
    PRIVATE_GOOGLE_LOADING = false;
    if(VIEW === 'calendar') renderContent();
  }
}

function viewCalendar(){
  const month = new Date(CAL_MONTH + '-01T00:00:00');
  const monthLabel = month.toLocaleDateString(undefined,{month:'long',year:'numeric'});
  const selectedBoard = CURRENT_BOARD ? board(CURRENT_BOARD) : null;
  const selectedGroup = selectedBoard && CALENDAR_GROUP_FILTER !== 'all'
    ? selectedBoard.groups.find(g=>g.id === CALENDAR_GROUP_FILTER)
    : null;
  const groupFilter = selectedGroup ? selectedGroup.id : 'all';
  if(!selectedGroup && CALENDAR_GROUP_FILTER !== 'all') CALENDAR_GROUP_FILTER = 'all';
  const first = new Date(month);
  const start = new Date(first);
  start.setDate(first.getDate() - first.getDay());
  const days = [];
  for(let i=0;i<42;i++){ const d=new Date(start); d.setDate(start.getDate()+i); days.push(d); }
  const windowStart = days[0].toISOString().slice(0,10);
  const windowEnd = days[days.length-1].toISOString().slice(0,10);
  const tasks = calendarVisibleTasks(windowStart, windowEnd);
  if(MY_GOOGLE_CALENDAR_ON) void loadMyGoogleCalendar(windowStart,windowEnd);
  const weeks = Array.from({length:6}, (_,i)=>days.slice(i*7, i*7+7));
  const groupButtons = selectedBoard ? `
    <div class="calendar-filter-row">
      <span class="calendar-filter-label">Groups</span>
      <button class="tool-btn ${groupFilter==='all'?'active':''}" data-calendar-group="all">All groups</button>
      ${selectedBoard.groups.map(g=>`<button class="tool-btn ${groupFilter===g.id?'active':''}" data-calendar-group="${esc(g.id)}">${esc(g.name)}</button>`).join('')}
    </div>` : '';
  return `
    <div class="toolbar calendar-filter-row">
      <span class="calendar-filter-label">Boards</span>
      <button class="btn ghost" data-action="calendar-all">All boards</button>
      ${DATA.boards.map(b=>`<button class="tool-btn ${selectedBoard&&selectedBoard.id===b.id?'active':''}" data-calendar-board="${b.id}">${esc(b.name)}</button>`).join('')}
    </div>
    ${groupButtons}
    <div class="toolbar calendar-filter-row google-calendar-toolbar">
      <span class="calendar-filter-label">Google Calendar</span>
      ${integrationConnection('google_calendar') ? `<button class="tool-btn ${MY_GOOGLE_CALENDAR_ON?'active':''}" data-action="toggle-my-google-calendar">${MY_GOOGLE_CALENDAR_ON?'Hide my Google Calendar':'Show my Google Calendar'}</button>` : `<button class="tool-btn" data-action="connect-calendar">Connect my Google Calendar</button>`}
      ${MY_GOOGLE_CALENDAR_ON && PRIVATE_GOOGLE_LOADING ? '<span class="calendar-load-note">Loading your events...</span>' : ''}
      ${MY_GOOGLE_CALENDAR_ON && PRIVATE_GOOGLE_ERROR ? `<span class="calendar-load-error">${esc(PRIVATE_GOOGLE_ERROR)}</span><button class="tool-btn" data-action="retry-my-google-calendar">Retry</button>` : ''}
    </div>
    <div class="calendar-shell">
      <div class="calendar-head">
        <div>
          <div class="calendar-title">${esc(monthLabel)}</div>
          <div style="font-size:12px;color:var(--ink3)">${esc(calendarFilterSubtitle(selectedBoard, selectedGroup))}</div>
        </div>
        <div class="calendar-actions">
          <button class="btn ghost sm" data-action="cal-prev">Previous</button>
          <button class="btn ghost sm" data-action="cal-today">Today</button>
          <button class="btn ghost sm" data-action="cal-next">Next</button>
        </div>
      </div>
      <div class="calendar-grid">
        <div class="calendar-dow-row">${['Sun','Mon','Tue','Wed','Thu','Fri','Sat'].map(d=>`<div class="cal-dow">${d}</div>`).join('')}</div>
        ${weeks.map(week=>calendarWeek(week, month, tasks)).join('')}
      </div>
      ${calendarItemsForWindow(tasks,windowStart,windowEnd).length ? '' : `<div class="calendar-empty">
        <h3>No scheduled tasks yet</h3>
        <p>Add a due date${selectedBoard && boardCalendarShowsTimeline(selectedBoard) ? ' or timeline' : ''} to a task and it will appear on this calendar.</p>
        ${DATA.boards[0]?.id ? `<button class="btn ghost sm" data-open-board="${DATA.boards[0].id}">Open a board</button>` : ''}
      </div>`}
    </div>`;
}
function calendarFilterSubtitle(selectedBoard, selectedGroup){
  if(selectedBoard && selectedGroup) return `${selectedBoard.name} / ${selectedGroup.name}`;
  if(selectedBoard) return `${selectedBoard.name} / All groups`;
  return 'All task boards';
}
function calendarWeek(days, month, tasks){
  const weekStart = days[0].toISOString().slice(0,10);
  const weekEnd = days[days.length - 1].toISOString().slice(0,10);
  const segments = calendarWeekSegments(calendarItemsForWindow(tasks, weekStart, weekEnd), weekStart, weekEnd);
  const eventRows = Math.max(1, segments.reduce((max,seg)=>Math.max(max, seg.slot + 1), 0));
  const dayCells = days.map((d,index)=>calendarDay(d,month,index,eventRows)).join('');
  const bars = segments.map(seg=>calendarItemButton(seg.item, seg)).join('');
  return `<div class="calendar-week" style="grid-template-rows:28px repeat(${eventRows}, minmax(44px,auto))">${dayCells}${bars}</div>`;
}
function calendarDay(d,month,index,eventRows){
  const iso = d.toISOString().slice(0,10);
  const muted = d.getMonth()!==month.getMonth();
  return `<div class="cal-day ${muted?'muted':''}" style="grid-column:${index + 1};grid-row:1 / span ${eventRows + 1}" data-cal-day="${iso}">
    <div class="cal-num">${d.getDate()}</div>
  </div>`;
}
function calendarVisibleTasks(windowStart='', windowEnd=''){
  const selectedBoard = CURRENT_BOARD ? board(CURRENT_BOARD) : null;
  const selectedGroupId = selectedBoard && selectedBoard.groups.some(g=>g.id === CALENDAR_GROUP_FILTER) ? CALENDAR_GROUP_FILTER : '';
  return allVisibleTasks().filter(t=>{
    if(selectedBoard && t._board.id !== selectedBoard.id) return false;
    if(selectedGroupId && t._group.id !== selectedGroupId) return false;
    if(windowStart && windowEnd) return taskCalendarOccurrences(t, windowStart, windowEnd).length > 0;
    return Boolean(taskCalendarSchedule(t) || recurrenceAnchorSchedule(t, null, recurrenceMeta(taskRecurrence(t,t._board))));
  });
}
function calendarItemsForWindow(tasks,windowStart,windowEnd){
  const googleItems = googleCalendarItems(
    MY_GOOGLE_CALENDAR_ON ? PRIVATE_GOOGLE_EVENTS : [],
    SHARED_GOOGLE_EVENTS,
    PRIVATE_GOOGLE_SHARED_IDS
  ).filter(item=>calendarItemIntersects(item,windowStart,windowEnd));
  return [...tasks.flatMap(t=>taskCalendarOccurrences(t, windowStart, windowEnd)),...googleItems]
    .sort((a,b)=>
      calendarItemTypeRank(a) - calendarItemTypeRank(b) ||
      isoDateCompare(a.start,b.start) ||
      isoDateCompare(b.end,a.end) ||
      String(a.task?.name || a.event?.title || '').localeCompare(String(b.task?.name || b.event?.title || ''))
    );
}
function calendarWeekSegments(items,weekStart,weekEnd){
  const rowEnds = [];
  return items
    .filter(item=>calendarItemIntersects(item, weekStart, weekEnd))
    .map(item=>{
      const start = isoDateCompare(item.start,weekStart) < 0 ? weekStart : item.start;
      const end = isoDateCompare(item.end,weekEnd) > 0 ? weekEnd : item.end;
      const colStart = isoDateDiffDays(weekStart, start) + 1;
      const colEnd = isoDateDiffDays(weekStart, end) + 2;
      let slot = rowEnds.findIndex(rowEnd=>rowEnd <= colStart);
      if(slot < 0){
        slot = rowEnds.length;
        rowEnds.push(0);
      }
      rowEnds[slot] = colEnd;
      return {
        item,
        colStart,
        colEnd,
        slot,
        continuesLeft:isoDateCompare(item.start, weekStart) < 0,
        continuesRight:isoDateCompare(item.end, weekEnd) > 0
      };
    });
}
function calendarItemTypeRank(item){
  if(item.type === 'due') return 0;
  if(item.type === 'recurring') return 1;
  if(item.type === 'timeline') return 2;
  if(item.type === 'google-shared') return 3;
  if(item.type === 'google-private') return 4;
  return 5;
}
function calendarItemIntersects(item,start,end){
  return isoDateCompare(item.end,start) >= 0 && isoDateCompare(item.start,end) <= 0;
}
function taskCalendarOccurrences(t,windowStart,windowEnd){
  const recurrence = taskRecurrence(t,t._board);
  const meta = recurrenceMeta(recurrence);
  if(meta.key !== 'off') return recurringCalendarOccurrences(t, meta, windowStart, windowEnd);
  const schedule = taskCalendarSchedule(t);
  if(!schedule || !calendarItemIntersects(schedule, windowStart, windowEnd)) return [];
  return [calendarItemFromSchedule(t, schedule, {
    type:schedule.source === 'Timeline' ? 'timeline' : 'due',
    label:taskCalendarScheduleLabel(t,schedule)
  })];
}
function recurringCalendarOccurrences(t,meta,windowStart,windowEnd){
  const seed = recurrenceAnchorSchedule(t, taskCalendarSchedule(t), meta);
  if(!seed || meta.interval < 1) return [];
  const duration = isoDateDiffDays(seed.start, seed.end);
  const recurrence = taskRecurrence(t,t._board);
  const out = [];
  let occurrenceStart = seed.start;
  let guard = 0;
  while(occurrenceStart && isoDateCompare(isoDateAdd(occurrenceStart, duration), windowStart) < 0 && guard < 520){
    const next = nextCalendarRecurrenceDate(occurrenceStart, meta);
    if(!next || next === occurrenceStart) return out;
    occurrenceStart = next;
    guard += 1;
  }
  while(occurrenceStart && isoDateCompare(occurrenceStart, windowEnd) <= 0 && guard < 1040){
    const occurrenceEnd = isoDateAdd(occurrenceStart, duration) || occurrenceStart;
    const item = calendarItemFromSchedule(t, {
      ...seed,
      start:occurrenceStart,
      end:occurrenceEnd,
      source:'Recurring'
    }, {
      type:'recurring',
      label:recurrenceLabel(recurrence),
      recurrence
    });
    if(calendarItemIntersects(item, windowStart, windowEnd)) out.push(item);
    const next = nextCalendarRecurrenceDate(occurrenceStart, meta);
    if(!next || next === occurrenceStart) break;
    occurrenceStart = next;
    guard += 1;
  }
  return out;
}
function recurrenceAnchorSchedule(t,baseSchedule,meta){
  if(meta.key === 'off') return null;
  const startFromRecurrence = meta.startsAt ? localIsoDateFromValue(meta.startsAt) : '';
  const recurrenceTime = meta.startsAt ? localDateTimeValue(meta.startsAt).slice(11,16) : '';
  if(baseSchedule){
    const start = startFromRecurrence || baseSchedule.start;
    const end = isoDateAdd(start, isoDateDiffDays(baseSchedule.start, baseSchedule.end)) || start;
    return {
      ...baseSchedule,
      start,
      end,
      startTime:recurrenceTime || baseSchedule.startTime,
      endTime:baseSchedule.endTime || recurrenceTime || baseSchedule.startTime,
      source:'Recurring'
    };
  }
  if(!startFromRecurrence) return null;
  return {
    start:startFromRecurrence,
    end:startFromRecurrence,
    startTime:recurrenceTime,
    endTime:recurrenceTime,
    source:'Recurring'
  };
}
function nextCalendarRecurrenceDate(startIso,meta){
  const d = new Date(`${startIso}T00:00:00`);
  if(Number.isNaN(d.getTime())) return '';
  return localIsoDateFromValue(addRecurrenceInterval(d, meta));
}
function calendarItemFromSchedule(t,schedule,{type,label='',recurrence=''}={}){
  return {
    id:`${t.id}-${type || schedule.source}-${schedule.start}`,
    task:t,
    type:type || 'due',
    start:schedule.start,
    end:schedule.end,
    startTime:schedule.startTime || '',
    endTime:schedule.endTime || '',
    label:label || taskCalendarScheduleLabel(t,schedule) || schedule.source,
    source:schedule.source,
    recurrence
  };
}
function calendarItemsForDay(iso){
  return calendarItemsForWindow(calendarVisibleTasks(iso, iso),iso,iso);
}
function calendarItemButton(item,segment={}){
  if(item.type === 'google-private' || item.type === 'google-shared'){
    const event = item.event;
    const shared = item.type === 'google-shared';
    const accent = shared ? '#15856d' : '#5a6aa6';
    const style = [
      segment.colStart ? `grid-column:${segment.colStart} / ${segment.colEnd}` : '',
      Number.isFinite(segment.slot) ? `grid-row:${segment.slot + 2}` : '',
      `--cal-accent:${accent};--cal-bg:${shared?'#e7f7f1':'#eff1fa'};--cal-border:${shared?'#a5d8c6':'#bec6e1'}`
    ].filter(Boolean).join(';');
    return `<button type="button" class="cal-task cal-bar ${item.type}" style="${style}" data-cal-google="${esc(item.id)}">
      <b>${esc(event.title || '(Untitled event)')}</b>
      <span>${shared?'Workspace event':'My Google Calendar'}${item.startTime ? ` · ${esc(item.startTime)}` : ''}</span>
    </button>`;
  }
  const t = item.task;
  const isRecurring = item.type === 'recurring';
  const cls = [
    'cal-task',
    'cal-bar',
    item.type || 'due',
    segment.continuesLeft ? 'continues-left' : '',
    segment.continuesRight ? 'continues-right' : ''
  ].filter(Boolean).join(' ');
  const label = item.label || (isRecurring ? 'Recurring' : 'Due date');
  const style = [
    segment.colStart ? `grid-column:${segment.colStart} / ${segment.colEnd}` : '',
    Number.isFinite(segment.slot) ? `grid-row:${segment.slot + 2}` : '',
    calendarItemStyle(item)
  ].filter(Boolean).join(';');
  return `<button type="button" class="${cls}" style="${style}" data-cal-task="${t._board.id}|${t.id}">
    <b>${esc(t.name)}</b>
    <span>${esc(label)} / ${esc(t._board.name)}</span>
  </button>`;
}
function calendarItemStyle(item){
  const accent = calendarItemAccent(item);
  return `--cal-accent:${esc(accent)};--cal-bg:${esc(rgbaFromHex(accent, item.type === 'recurring' ? 0.14 : 0.12))};--cal-border:${esc(rgbaFromHex(accent, 0.28))}`;
}
function calendarItemAccent(item){
  if(item.type === 'google-shared') return '#15856d';
  if(item.type === 'google-private') return '#5a6aa6';
  return taskColor(item.task) || ownerCalendarColor(item.task);
}
function rgbaFromHex(hex,alpha){
  const raw = String(hex || '').replace('#','').trim();
  if(!/^[0-9a-f]{6}$/i.test(raw)) return `rgba(0,115,234,${alpha})`;
  const n = parseInt(raw,16);
  const r = (n >> 16) & 255;
  const g = (n >> 8) & 255;
  const b = n & 255;
  return `rgba(${r},${g},${b},${alpha})`;
}
function showCalendarTaskDetail(key,anchor){
  const [, taskId] = String(key||'').split('|');
  const t = allTasks().find(x=>x.id === taskId);
  if(!t) return;
  const dd = buildDD(anchor, calendarTaskDetailHtml(t));
  dd.classList.add('calendar-popover');
  positionDropdown(dd, anchor);
  dd.querySelector('[data-open-cal-detail]')?.addEventListener('click',()=>{
    closeDD();
    openTaskDetail(t._board.id,t._group.id,t.id,'updates');
  });
  dd.querySelector('[data-save-cal-task]')?.addEventListener('click',()=>{
    const loc=findTaskLocationById(t.id);
    if(!loc) return;
    const summary=dd.querySelector('[data-cal-task-summary]')?.value || '';
    setTaskDescription(loc.t,summary);
    const done=dd.querySelector('[data-cal-task-done]')?.checked === true;
    if((loc.t.status==='done') !== done) setTaskStatusFromKanban(t.id,done?'done':'working');
    else { save(); renderContent(); }
    closeDD();
    toast('Calendar task saved');
  });
  dd.querySelector('[data-sync-google-calendar]')?.addEventListener('click',()=>syncTaskToGoogleCalendar(t.id));
}
function showGoogleCalendarEventDetail(key,anchor){
  const item = googleCalendarItems(PRIVATE_GOOGLE_EVENTS,SHARED_GOOGLE_EVENTS,PRIVATE_GOOGLE_SHARED_IDS).find(row=>row.id === key);
  if(!item) return;
  const event = item.event;
  const shared = item.type === 'google-shared';
  const owner = shared && event.owner_member_id === currentMemberId();
  const description = event.description || '';
  const displayEnd = googleEventDisplayEnd(event);
  const timeLabel = event.allDay ?? event.all_day
    ? `${fmtDate(item.start)}${item.end !== item.start ? ` – ${fmtDate(item.end)}` : ''} · All day`
    : `${fmtDate(item.start)} ${item.startTime} – ${fmtDate(displayEnd.date)} ${displayEnd.time}`;
  const syncLabel = shared && event.sync_status === 'stale'
    ? (event.sync_error === 'reconnect'
      ? (owner ? 'Google access expired. Reconnect Google Calendar to resume updates.' : 'Waiting for the owner to reconnect Google Calendar.')
      : `Google sync is delayed.${event.last_synced_at ? ` Last synced ${fmtTimeAgo(Date.parse(event.last_synced_at))}.` : ''}`)
    : (shared && event.last_synced_at ? `Last synced ${fmtTimeAgo(Date.parse(event.last_synced_at))}` : '');
  const dd = buildDD(anchor, `<div class="cal-pop-detail">
    <div class="cal-pop-head"><b>${esc(event.title || '(Untitled event)')}</b><span>${shared?'Shared with workspace':'Only you can see this Google event'}</span></div>
    <div class="cal-detail-grid">
      <span>When</span><b>${esc(timeLabel)}</b>
      ${event.location ? `<span>Location</span><b>${esc(event.location)}</b>` : ''}
      ${syncLabel ? `<span>Sync</span><b>${esc(syncLabel)}</b>` : ''}
    </div>
    ${description ? `<div class="cal-detail-notes">${renderLinkedText(description,{lineBreaks:true})}</div>` : ''}
    ${shared ? `<div class="cal-summary-editor"><label for="cal-event-summary">Workspace summary</label><textarea id="cal-event-summary" rows="8" placeholder="What happened, decisions, outcomes, and follow-up work...">${esc(event.workspace_summary || '')}</textarea><label class="cal-done-check"><input type="checkbox" id="cal-event-done" ${event.resolution_status==='done'?'checked':''}> Finished / done</label><button type="button" class="btn sm" data-save-shared-event>Save event</button></div>` : ''}
    <div class="cal-detail-actions">
      ${shared ? (owner ? `<button type="button" class="btn ghost sm" data-unshare-google="${esc(event.id)}">Unshare from workspace</button>` : '') : `<button type="button" class="btn sm" data-share-google="${esc(event.id)}">Share to the workspace</button>`}
    </div>
  </div>`);
  dd.classList.add('calendar-popover');
  positionDropdown(dd,anchor);
  dd.querySelector('[data-share-google]')?.addEventListener('click',()=>{ void sharePersonalGoogleEvent(event.id); });
  dd.querySelector('[data-unshare-google]')?.addEventListener('click',()=>{ void unshareWorkspaceGoogleEvent(event.id); });
  dd.querySelector('[data-save-shared-event]')?.addEventListener('click',async()=>{
    const summary=dd.querySelector('#cal-event-summary')?.value || '';
    const done=dd.querySelector('#cal-event-done')?.checked === true;
    if(summary.length>20000){ toast('Summary is too long'); return; }
    const patch={workspace_summary:summary,resolution_status:done?'done':'open',resolved_at:done?(event.resolved_at || new Date().toISOString()):null};
    const {data,error}=await supabase.from('app_shared_calendar_events').update(patch).eq('workspace_id',WORKSPACE_ID).eq('id',event.id).select().single();
    if(error){ toast(error.message || 'Could not save event'); return; }
    SHARED_GOOGLE_EVENTS=SHARED_GOOGLE_EVENTS.map(row=>row.id===event.id?data:row);
    closeDD(); renderContent(); toast('Event saved');
  });
}
async function sharePersonalGoogleEvent(eventId){
  if(!supabase || !WORKSPACE_ID) return;
  const {data,error} = await supabase.functions.invoke('google-calendar-share',{body:{
    action:'share',workspaceId:WORKSPACE_ID,eventId
  }});
  if(error || data?.error){ toast(await assistantFunctionErrorMessage(error,data)); return; }
  if(data?.event){
    SHARED_GOOGLE_EVENTS = [data.event,...SHARED_GOOGLE_EVENTS.filter(row=>row.id !== data.event.id)];
    PRIVATE_GOOGLE_SHARED_IDS.add(eventId);
  }
  closeDD();
  render();
  void loadWorkspaceSharedCalendar({force:true});
  toast(data?.alreadyShared ? 'Already shared with workspace' : 'Shared to workspace and announced in updates');
}
async function unshareWorkspaceGoogleEvent(sharedEventId){
  if(!supabase || !WORKSPACE_ID) return;
  const {data,error} = await supabase.functions.invoke('google-calendar-share',{body:{
    action:'unshare',workspaceId:WORKSPACE_ID,sharedEventId
  }});
  if(error || data?.error){ toast(await assistantFunctionErrorMessage(error,data)); return; }
  SHARED_GOOGLE_EVENTS = SHARED_GOOGLE_EVENTS.filter(row=>row.id !== sharedEventId);
  PRIVATE_GOOGLE_LOAD_KEY = '';
  closeDD();
  render();
  void loadWorkspaceSharedCalendar({force:true});
  toast('Removed from workspace');
}
function showCalendarMore(iso,anchor){
  const items = calendarItemsForDay(iso);
  const label = new Date(`${iso}T00:00:00`).toLocaleDateString(undefined,{weekday:'long',month:'short',day:'numeric'});
  const rows = items.map(item=>{
    if(item.event){
      return `<button type="button" class="cal-pop-row" data-cal-pop-google="${esc(item.id)}">
        <span class="cal-pop-accent" style="background:${esc(calendarItemAccent(item))}"></span>
        <span><b>${esc(item.event.title || '(Untitled event)')}</b><small>${item.type==='google-shared'?'Workspace event':'My Google Calendar'}</small></span>
      </button>`;
    }
    const t = item.task;
    return `<button type="button" class="cal-pop-row" data-cal-pop-task="${t._board.id}|${t.id}">
    <span class="cal-pop-accent" style="background:${esc(calendarItemAccent(item))}"></span>
    <span><b>${esc(t.name)}</b><small>${esc(t._board.name)} · ${esc(t._group.name)}</small></span>
  </button>`;
  }).join('').replace(/\s\u00c2\u00b7\s/g,' / ');
  const dd = buildDD(anchor, `<div class="cal-pop-head"><b>${esc(label)}</b><span>${items.length} item${items.length===1?'':'s'}</span></div><div class="cal-pop-list">${rows}</div>`);
  dd.classList.add('calendar-popover');
  positionDropdown(dd, anchor);
  dd.querySelectorAll('[data-cal-pop-task]').forEach(el=>el.onclick=()=>showCalendarTaskDetail(el.dataset.calPopTask, el));
  dd.querySelectorAll('[data-cal-pop-google]').forEach(el=>el.onclick=()=>showGoogleCalendarEventDetail(el.dataset.calPopGoogle, el));
}
function calendarTaskDetailHtml(t){
  const st = statusOptionFor(t._board,t.status);
  const pr = priorityOptionFor(t._board,t.priority);
  const recurrence = taskRecurrence(t,t._board);
  const recurring = recurrence !== 'off';
  const schedule = taskCalendarSchedule(t) || (recurring ? recurrenceAnchorSchedule(t, null, recurrenceMeta(recurrence)) : null);
  const scheduleLabel = taskCalendarScheduleLabel(t,schedule);
  const currentCalendarOwner = calendarSyncCurrentOwner();
  const ownerLinks = (currentCalendarOwner ? [currentCalendarOwner] : [])
    .map(owner=>({owner, event:taskGoogleCalendarEventForOwner(t, owner)}))
    .filter(row=>row.event?.url);
  const latestSyncedAt = Object.values(taskGoogleCalendarEvents(t))
    .map(event=>Date.parse(event?.syncedAt || ''))
    .filter(Boolean)
    .sort((a,b)=>b-a)[0] || Date.parse(t.custom?.[TASK_GOOGLE_CALENDAR_SYNCED_KEY] || '');
  const googleSynced = latestSyncedAt ? fmtTimeAgo(latestSyncedAt) : '';
  return `<div class="cal-pop-detail">
    <div class="cal-pop-head">
      <b>${esc(t.name)}</b>
      <span>${esc(t._board.name)} · ${esc(t._group.name)}</span>
    </div>
    <div class="cal-detail-grid">
      <span>${esc(schedule?.source || 'Schedule')}</span><b>${esc(scheduleLabel || 'No schedule')}</b>
      ${recurring ? `<span>Recurring</span><b>${esc(recurrenceLabel(recurrence))}</b>` : ''}
      <span>Owners</span><b>${esc(taskOwnerNames(t) || 'Unassigned')}</b>
      <span>Status</span><b><i style="background:${st.color}"></i>${esc(st.label)}</b>
      <span>Priority</span><b><i style="background:${pr.color}"></i>${esc(pr.label)}</b>
      ${googleSynced ? `<span>Google Cal</span><b>Synced ${esc(googleSynced)}</b>` : ''}
      ${Number(t.budget) ? `<span>Budget</span><b>${esc(fmtMoney(t.budget))}</b>` : ''}
    </div>
    ${t.notes ? `<div class="cal-detail-notes">${renderLinkedText(t.notes)}</div>` : ''}
    <div class="cal-summary-editor"><label for="cal-task-summary">Description / summary</label><textarea id="cal-task-summary" data-cal-task-summary rows="8" placeholder="What happened, decisions, outcomes, and follow-up work...">${esc(taskDescription(t))}</textarea><label class="cal-done-check"><input type="checkbox" data-cal-task-done ${t.status==='done'?'checked':''}> Finished / done</label><button type="button" class="btn sm" data-save-cal-task>Save task</button></div>
    <div class="cal-detail-actions">
      <button type="button" class="btn ghost sm" data-open-cal-detail>Open details</button>
      ${schedule ? `<button type="button" class="btn ghost sm" data-sync-google-calendar>${ownerLinks.length?'Update my calendar':'Sync my calendar'}</button>` : ''}
      ${ownerLinks.slice(0,1).map(row=>`<a class="btn ghost sm" href="${esc(row.event.url)}" target="_blank" rel="noopener noreferrer">Open my event</a>`).join('')}
    </div>
  </div>`;
}
function boardCalendarShowsTimeline(b){
  return b?.calendarShowTimeline === true;
}
function taskCalendarSchedule(t){
  if(boardCalendarShowsTimeline(t?._board)){
    let start = validIsoDate(t?.tlStart);
    let end = validIsoDate(t?.tlEnd) || start;
    if(start){
      if(end && isoDateCompare(start,end) > 0) [start,end] = [end,start];
      return {
        start,
        end:end || start,
        startTime:validTimeValue(taskMetaValue(t, TASK_TIMELINE_START_TIME_KEY) || taskMetaValue(t, TASK_DUE_TIME_KEY)),
        endTime:validTimeValue(taskMetaValue(t, TASK_TIMELINE_END_TIME_KEY) || taskMetaValue(t, TASK_DUE_TIME_KEY)),
        source:'Timeline'
      };
    }
  }
  const due = validIsoDate(t?.due);
  const days = isoDateRange(due, due);
  if(!days.length) return null;
  const startTime = validTimeValue(taskMetaValue(t, TASK_DUE_TIME_KEY));
  const endTime = startTime;
  return {
    start:days[0],
    end:days[days.length - 1],
    startTime,
    endTime,
    source:'Due date'
  };
}
function taskCalendarScheduleLabel(t,schedule=taskCalendarSchedule(t)){
  if(!schedule) return '';
  if(schedule.source === 'Recurring'){
    const startLabel = fmtDateTime(schedule.start, schedule.startTime || schedule.endTime);
    const endLabel = schedule.end && schedule.end !== schedule.start ? fmtDateTime(schedule.end, schedule.endTime || schedule.startTime) : '';
    return endLabel ? `${startLabel} - ${endLabel}` : startLabel;
  }
  if(schedule.source === 'Timeline'){
    const startLabel = fmtDateTime(schedule.start, schedule.startTime || schedule.endTime);
    const endLabel = schedule.end && schedule.end !== schedule.start ? fmtDateTime(schedule.end, schedule.endTime || schedule.startTime) : '';
    return endLabel ? `${startLabel} - ${endLabel}` : startLabel;
  }
  if(schedule.source === 'Due date') return fmtDateTime(schedule.start, schedule.startTime || schedule.endTime);
  return fmtDateTime(schedule.start, schedule.startTime || schedule.endTime);
}
function googleCalendarRecurrenceForTask(t){
  const meta = recurrenceMeta(taskRecurrence(t,t._board));
  if(meta.key === 'off' || !meta.interval) return [];
  const freq = meta.unit === 'months' ? 'MONTHLY' : (meta.unit === 'weeks' ? 'WEEKLY' : 'DAILY');
  const interval = Math.max(1, Math.min(365, Number(meta.interval) || 1));
  return [`RRULE:FREQ=${freq};INTERVAL=${interval}`];
}
function googleCalendarPayloadForTask(t, owner=null){
  const recurrenceValue = taskRecurrence(t,t._board);
  const recurrenceMetaValue = recurrenceMeta(recurrenceValue);
  const schedule = taskCalendarSchedule(t) || recurrenceAnchorSchedule(t, null, recurrenceMetaValue);
  if(!schedule) return null;
  const hasTime = Boolean(schedule.startTime || schedule.endTime);
  const startValue = hasTime ? calendarDateTimeValue(schedule.start, schedule.startTime || schedule.endTime, '09:00') : schedule.start;
  let endValue = hasTime ? calendarDateTimeValue(schedule.end, schedule.endTime || schedule.startTime, '17:00') : (isoDateAdd(schedule.end, 1) || schedule.end);
  if(hasTime && startValue && endValue && new Date(endValue).getTime() <= new Date(startValue).getTime()){
    endValue = calendarDateTimeAddMinutes(schedule.start, schedule.startTime || schedule.endTime, 60);
  }
  const status = statusOptionFor(t._board,t.status).label;
  const priority = priorityOptionFor(t._board,t.priority).label;
  const ownerEvent = owner ? taskGoogleCalendarEventForOwner(t, owner) : {};
  const recurrence = recurrenceMetaValue.key === 'off' ? [] : googleCalendarRecurrenceForTask(t);
  return {
    eventId:owner ? (ownerEvent.id || '') : (t.custom?.[TASK_GOOGLE_CALENDAR_EVENT_KEY] || ''),
    summary:t.name || 'TaskTrack task',
    start:startValue,
    end:endValue,
    recurrence,
    description:[
      `TaskTrack task: ${t.name || 'Untitled task'}`,
      `Board: ${t._board?.name || ''}`,
      `Group: ${t._group?.name || ''}`,
      `Schedule source: ${schedule.source}`,
      `Schedule: ${taskCalendarScheduleLabel(t,schedule) || `${fmtDate(schedule.start)} - ${fmtDate(schedule.end)}`}`,
      `Owners: ${taskOwnerNames(t) || 'Unassigned'}`,
      `Status: ${status}`,
      priority && t.priority !== 'none' ? `Priority: ${priority}` : '',
      t.notes ? `Notes: ${t.notes}` : '',
      `TaskTrack board: ${boardUrl(t._board.id)}`
    ].filter(Boolean).join('\n')
  };
}
async function syncTaskOwnerCalendar(t, owner, {silent=false}={}){
  const payload = googleCalendarPayloadForTask(t, owner);
  if(!payload) throw new Error('Add a due date first');
  const { data, error } = await supabase.functions.invoke('assistant-actions', {
    body:{
      workspaceId:WORKSPACE_ID,
      provider:'google_calendar',
      action:'calendar.create_event',
      payload,
      confirmed:true
    }
  });
  if(error || data?.error){
    const message = await assistantFunctionErrorMessage(error, data);
    if(!silent) toast(message);
    throw new Error(message);
  }
  const result = data?.result || {};
  setTaskGoogleCalendarEventForOwner(t, owner, result);
  return result;
}
async function syncAssignedTaskCalendars(t,{silent=false}={}){
  if(!supabase || !SESSION?.user || !WORKSPACE_ID){
    if(!silent) toast('Sign in before syncing Google Calendar');
    return {synced:0, failed:0};
  }
  const owner = calendarSyncCurrentOwner();
  if(!googleCalendarPayloadForTask(t, owner)){
    if(!silent) toast('Add a due date first');
    return {synced:0, failed:0};
  }
  if(!owner){
    if(!silent) toast('Add your email in Settings before syncing Google Calendar');
    return {synced:0, failed:0};
  }
  let synced = 0;
  let failed = 0;
  try{
    await syncTaskOwnerCalendar(t, owner, {silent});
    synced += 1;
  }catch(e){
    failed += 1;
    if(!silent) console.warn(e);
  }
  if(synced){
    save();
    if(!silent) renderContent();
  }
  if(!silent){
    if(synced && failed) toast('Synced my calendar; one update failed');
    else if(synced) toast('Synced my calendar');
    else toast('No calendar synced. Connect your Google Calendar first.');
  }
  return {synced, failed};
}
async function syncTaskToGoogleCalendar(taskId){
  const t = allTasks().find(x=>x.id === taskId);
  if(!t) return;
  if(!supabase || !SESSION?.user || !WORKSPACE_ID){
    toast('Sign in before syncing Google Calendar');
    return;
  }
  if(!googleCalendarPayloadForTask(t, calendarSyncCurrentOwner())){
    toast('Add a due date first');
    return;
  }
  toast('Syncing my Google Calendar...');
  await syncAssignedTaskCalendars(t);
}
function shiftCalendar(n){
  const d = new Date(CAL_MONTH + '-01T00:00:00');
  d.setMonth(d.getMonth()+n);
  CAL_MONTH = d.toISOString().slice(0,7);
  renderContent();
}

function viewMonthlyClosing(){
  const month = /^\d{4}-\d{2}$/.test(CLOSING_MONTH) ? CLOSING_MONTH : new Date().toISOString().slice(0,7);
  const label = new Date(`${month}-01T12:00:00`).toLocaleDateString(undefined,{month:'long',year:'numeric'});
  const tasks = allVisibleTasks().filter(t=>{
    const completed = String(t.completedAt || '').slice(0,7);
    return t.status==='done' ? completed===month || (!completed && String(t.due||'').slice(0,7)===month) : String(t.due||'').slice(0,7)===month;
  });
  const rank = {critical:0,urgent:0,high:1,medium:2,normal:2,low:3,none:4};
  tasks.sort((a,b)=>(rank[a.priority] ?? 3)-(rank[b.priority] ?? 3) || Number(b.status==='done')-Number(a.status==='done') || String(a.due||'').localeCompare(String(b.due||'')));
  const events = SHARED_GOOGLE_EVENTS.filter(row=>row.sync_status!=='canceled' && googleEventCalendarSpan(row)?.start?.slice(0,7)===month)
    .sort((a,b)=>String(a.start_date || a.start_at || '').localeCompare(String(b.start_date || b.start_at || '')));
  const done = tasks.filter(t=>t.status==='done').length;
  return `<div class="closing-shell">
    <div class="closing-head"><div><h2>${esc(label)}</h2><p>Completed work and the status of work scheduled this month, across everyone.</p></div><label>Month <input type="month" id="closing-month" value="${esc(month)}"></label></div>
    <div class="closing-stats"><div><b>${done}</b><span>Tasks done</span></div><div><b>${tasks.length-done}</b><span>Other tasks due</span></div><div><b>${events.filter(e=>e.resolution_status==='done').length}/${events.length}</b><span>Events finished</span></div></div>
    <h3>Tasks by priority</h3>
    <div class="closing-list">${tasks.map(t=>{
      const priority=priorityOptionFor(t._board,t.priority);
      const status=statusOptionFor(t._board,t.status);
      return `<button type="button" class="closing-row" data-closing-task="${esc(t._board.id)}|${esc(t.id)}"><span class="closing-priority" style="background:${esc(priority.color)}">${esc(priority.label)}</span><span class="closing-main"><b>${esc(t.name)}</b><small>${esc(t._board.name)} · ${esc(taskOwnerNames(t)||'Unassigned')}</small></span><span class="closing-status" style="color:${esc(status.color)}">${esc(status.label)}</span><time>${esc(t.status==='done' ? fmtDate(String(t.completedAt||t.due).slice(0,10)) : fmtDate(t.due))}</time></button>`;
    }).join('') || '<p class="closing-empty">No tasks completed or due this month.</p>'}</div>
    <h3>Calendar events</h3>
    <div class="closing-list">${events.map(event=>`<button type="button" class="closing-row" data-closing-event="${esc(event.id)}"><span class="closing-main"><b>${esc(event.title)}</b><small>${esc(googleEventCalendarSpan(event)?.start || '')}</small></span><span class="closing-status">${event.resolution_status==='done'?'Done':'Open'}</span>${event.workspace_summary?'<span>Summary added</span>':''}</button>`).join('') || '<p class="closing-empty">No shared events this month.</p>'}</div>
  </div>`;
}

/* ============================================================
   VIEW: Dashboard
   ============================================================ */
function viewDashboard(){
  const tasks = allVisibleTasks();
  const total = tasks.length;
  const inprog = tasks.filter(t=>t.status==='working').length;
  const stuck = tasks.filter(t=>t.status==='stuck').length;
  const done = tasks.filter(t=>t.status==='done').length;
  const overdue = tasks.filter(isOverdueTask);
  const dueSoon = tasks.filter(t=>t.status!=='done' && isDueWithin(t.due,7));
  const unassigned = tasks.filter(t=>!taskOwnerIds(t).length);
  const needsReview = tasks.filter(t=>t.status==='review');

  // status breakdown
  const byStatus = {}, statusPalette = {};
  tasks.forEach(t=>{
    const opt = statusOptionFor(t._board,t.status);
    byStatus[opt.label]=(byStatus[opt.label]||0)+1;
    statusPalette[opt.label]=opt;
  });
  // owner breakdown
  const byOwner = {};
  tasks.forEach(t=>{
    const owners = taskOwners(t);
    if(!owners.length){ byOwner.Unassigned=(byOwner.Unassigned||0)+1; return; }
    owners.forEach(o=>{ byOwner[o.name]=(byOwner[o.name]||0)+1; });
  });
  // priority
  const byPrio = {}, prioColors = {};
  tasks.forEach(t=>{
    if(t.priority==='none') return;
    const opt = priorityOptionFor(t._board,t.priority);
    byPrio[opt.label]=(byPrio[opt.label]||0)+1;
    prioColors[opt.label]=opt.color;
  });
  // budget by board
  const byBoard = DATA.boards.map(b=>({name:b.name, val:b.groups.reduce((s,g)=>s+g.tasks.reduce((x,t)=>x+taskTreeBudget(t),0),0)}));
  const hasBudget = byBoard.some(b=>b.val > 0);

  if(total===0) return `<div class="empty-state"><div class="big">📊</div>No tasks yet. Add some on a board to see reporting.</div>`;

  return `
    <div class="dashboard-note">Dashboard counts include visible tasks and subtasks across every board.</div>
    <div class="kpi-row">
      ${kpi('All Tasks', total, '#0073ea')}
      ${kpi('Overdue', overdue.length, '#e2445c')}
      ${kpi('Scheduled 7 Days', dueSoon.length, '#fdab3d')}
      ${kpi('Unassigned', unassigned.length, '#5e6170')}
      ${kpi('Needs Review', needsReview.length, '#579bfc')}
      ${kpi('Done', done, '#00c875')}
    </div>
    <div class="chart-row">
      <div class="card"><h3>Tasks by Status</h3>${donut(byStatus, statusPalette)}</div>
      <div class="card"><h3>Tasks by Owner</h3>${hbars(byOwner)}</div>
    </div>
    <div class="chart-row">
      <div class="card"><h3>By Priority</h3>${hbars(byPrio, Object.keys(byPrio).map(k=>prioColors[k]), false, 'No priorities set yet. Set a priority on task rows to populate this chart.')}</div>
      <div class="card"><h3>Budget by Board</h3>${hasBudget ? hbars(Object.fromEntries(byBoard.map(b=>[b.name,b.val])), null, true) : '<div class="mw-empty">Add budgets to task rows to see board totals here.</div>'}</div>
    </div>
    <div class="chart-row">
      <div class="card"><h3>Overdue Work</h3>${taskMiniList(overdue,'No overdue tasks')}</div>
      <div class="card"><h3>Scheduled This Week</h3>${taskMiniList(dueSoon,'Nothing scheduled in the next 7 days')}</div>
    </div>
    <div class="chart-row">
      <div class="card"><h3>Unassigned Tasks</h3>${taskMiniList(unassigned,'Every active task has an owner')}</div>
      <div class="card"><h3>Needs Review</h3>${taskMiniList(needsReview,'No tasks waiting for review')}</div>
    </div>`;
}
function kpi(label,num,color){
  return `<div class="kpi"><div class="lbl"><span class="dot" style="background:${color}"></span>${label}</div><div class="num">${num}</div></div>`;
}
function donut(counts, palette){
  const entries = Object.entries(counts).filter(([k,v])=>v>0);
  const total = entries.reduce((s,[,v])=>s+v,0)||1;
  let off=0; const R=52, C=2*Math.PI*R;
  const segs = entries.map(([k,v])=>{
    const frac=v/total, len=frac*C;
    const seg=`<circle r="${R}" cx="70" cy="70" fill="none" stroke="${palette[k].color}" stroke-width="22"
      stroke-dasharray="${len} ${C-len}" stroke-dashoffset="${-off}" transform="rotate(-90 70 70)"/>`;
    off+=len; return seg;
  }).join('');
  const legend = entries.map(([k,v])=>`<div class="legend-item"><span class="dot" style="background:${palette[k].color}"></span>${esc(palette[k].label)} · ${Math.round(v/total*100)}% (${v})</div>`).join('');
  return `<div class="donut-wrap">
    <svg width="140" height="140" viewBox="0 0 140 140">${segs}
      <text x="70" y="66" text-anchor="middle" font-size="26" font-weight="800" fill="#323338">${total}</text>
      <text x="70" y="86" text-anchor="middle" font-size="11" fill="#5e6170">total</text></svg>
    <div class="legend">${legend}</div></div>`;
}
function hbars(counts, colors, money, empty='No data'){
  const entries = Object.entries(counts);
  const max = Math.max(1, ...entries.map(([,v])=>v));
  const def = ['#0073ea','#00c875','#fdab3d','#e2445c','#a25ddc','#579bfc'];
  return `<div class="hbar">${entries.map(([k,v],i)=>{
    const col = (Array.isArray(colors)&&colors[i]) ? colors[i] : def[i%def.length];
    const w = Math.max(6,(v/max)*100);
    const lbl = money ? fmtMoney(v) : v;
    return `<div class="hbar-row"><div class="hbar-label" title="${esc(k)}">${esc(k)}</div>
      <div class="hbar-track"><div class="hbar-fill" style="width:${w}%;background:${col}">${lbl}</div></div></div>`;
  }).join('')||`<div class="mw-empty">${esc(empty)}</div>`}</div>`;
}
function taskMiniList(tasks,empty){
  const rows = tasks.slice(0,8).map(t=>{
    return `<div class="dash-task-row" data-dash-task="${t._board.id}|${t.id}">
      <div>
        <b>${esc(t.name)}</b>
        <span>${esc(t._board.name)} / ${esc(t._group.name)}</span>
      </div>
      <div>${esc(taskOwnerNames(t) || 'Unassigned')}</div>
      <div>${esc(taskDueLabel(t) || 'No due date')}</div>
    </div>`;
  }).join('');
  const more = tasks.length>8 ? `<div class="dash-more">+${tasks.length-8} more</div>` : '';
  return rows ? `<div class="dash-task-list">${rows}${more}</div>` : `<div class="mw-empty">${esc(empty)}</div>`;
}

/* ============================================================
   VIEW: My Work
   ============================================================ */
function viewMyWorkLegacy(){
  const me = currentMemberId();
  const mine = allVisibleTasks().filter(t=>taskOwnerIds(t).includes(me));
  const today = new Date(); today.setHours(0,0,0,0);
  const dd = iso => iso? new Date(iso+'T00:00') : null;
  const buckets = {
    'Past Dates':[], 'Today':[], 'This Week':[], 'Next Week':[], 'Later':[], 'Without a date':[]
  };
  mine.forEach(t=>{
    const d = dd(t.due);
    if(!d){ buckets['Without a date'].push(t); return; }
    const diff = Math.round((d-today)/86400000);
    if(diff<0) buckets['Past Dates'].push(t);
    else if(diff===0) buckets['Today'].push(t);
    else if(diff<=7) buckets['This Week'].push(t);
    else if(diff<=14) buckets['Next Week'].push(t);
    else buckets['Later'].push(t);
  });
  const colors={'Past Dates':'#e2445c','Today':'#00c875','This Week':'#0073ea','Next Week':'#579bfc','Later':'#fdab3d','Without a date':'#5e6170'};
  const sections = Object.entries(buckets).map(([name,items])=>`
    <div class="mw-bucket">
      <div class="mw-head"><span style="color:${colors[name]};font-weight:700">▾</span>
        <span class="ttl" style="color:${colors[name]}">${name}</span><span class="cnt">${items.length} item${items.length!==1?'s':''}</span></div>
      ${items.length? `<div class="mw-th"><div>Task</div><div>Group</div><div>Board</div><div></div><div>Due</div><div>Status</div><div>Priority</div></div>`+
        items.map(t=>{
        const st=statusOptionFor(t._board,t.status), pr=priorityOptionFor(t._board,t.priority);
        return `<div class="mw-row"><div class="accent" style="font-weight:600">${esc(t.name)}</div>
          <div style="color:var(--ink2)">${esc(t._group.name)}</div>
          <div style="color:var(--ink2)">${esc(t._board.name)}</div>
          <div></div>
          <div>${t.due?fmtDate(t.due):'—'}</div>
          <div><span style="background:${st.color};color:#fff;padding:3px 8px;border-radius:4px;font-size:12px;font-weight:600">${esc(st.label)}</span></div>
          <div><span style="background:${pr.color};color:#fff;padding:3px 8px;border-radius:4px;font-size:12px;font-weight:600">${esc(pr.label)}</span></div>
        </div>`;}).join('')
        : '<div class="mw-empty">No items</div>'}
    </div>`).join('');
  return `<div style="margin-bottom:14px;color:var(--ink2);font-size:13px">Showing tasks assigned to <b>${esc(member(me)?.name||'you')}</b>. Change the owner of a task to see it here.</div>${sections}`;
}

function myWorkBucketName(t,today){
  if(!t.due) return 'Without a date';
  const due = new Date(`${t.due}T00:00`);
  const diff = Math.round((due - today) / 86400000);
  if(diff < 0) return 'Past Dates';
  if(diff === 0) return 'Today';
  if(diff <= 7) return 'This Week';
  if(diff <= 14) return 'Next Week';
  return 'Later';
}

const MY_WORK_SORT_OPTIONS = [
  {id:'due',label:'Due date'},
  {id:'status',label:'Status'},
  {id:'priority',label:'Priority'}
];

function myWorkChildRow(parent,child){
  return {...child,_board:parent._board,_group:parent._group,_parent:parent};
}

function myWorkAllTasks(){
  return typeof allVisibleTasks === 'function' ? allVisibleTasks() : allTasks();
}

function myWorkTaskHidden(t){
  return typeof taskIsHidden === 'function' ? taskIsHidden(t) : Boolean(t?.hiddenAt);
}

function myWorkVisibleChildRows(t){
  return (t?.subtasks || [])
    .filter(child=>!myWorkTaskHidden(child))
    .map(child=>myWorkChildRow(t,child));
}

function myWorkHasAssignedDescendant(t,assignedIds){
  return myWorkVisibleChildRows(t).some(child=>assignedIds.has(child.id) || myWorkHasAssignedDescendant(child,assignedIds));
}

function myWorkHasSectionDescendant(t,sectionIds){
  return myWorkVisibleChildRows(t).some(child=>sectionIds.has(child.id) || myWorkHasSectionDescendant(child,sectionIds));
}

function myWorkPrimarySectionTask(t,sectionIds){
  if(sectionIds.has(t.id)) return t;
  for(const child of myWorkVisibleChildRows(t)){
    const found = myWorkPrimarySectionTask(child,sectionIds);
    if(found) return found;
  }
  return null;
}

function myWorkBucketForRoot(t,today,sectionIds){
  return myWorkBucketName(myWorkPrimarySectionTask(t,sectionIds) || t,today);
}

function myWorkRootTasks(sectionIds){
  return myWorkAllTasks()
    .filter(t=>!t._parent)
    .filter(t=>sectionIds.has(t.id) || myWorkHasSectionDescendant(t,sectionIds));
}

function myWorkRecurringDueOn(t,iso){
  const recurrence = taskRecurrence(t,t._board);
  const meta = recurrenceMeta(recurrence);
  if(meta.key === 'off' || meta.interval < 1) return false;
  const seed = recurrenceAnchorSchedule(t, taskCalendarSchedule(t), meta);
  if(!seed?.start) return false;
  let occurrenceStart = seed.start;
  let guard = 0;
  while(occurrenceStart && isoDateCompare(occurrenceStart, iso) < 0 && guard < 1040){
    const next = nextCalendarRecurrenceDate(occurrenceStart, meta);
    if(!next || next === occurrenceStart) return false;
    occurrenceStart = next;
    guard += 1;
  }
  return occurrenceStart === iso;
}

function myWorkRelevantChildren(t,sectionIds,{ancestorIncluded=false,mode='normal',today=todayIso()}={}){
  const currentIncluded = ancestorIncluded || sectionIds.has(t.id);
  return myWorkVisibleChildRows(t).filter(child=>{
    const childRecurring = isRecurringTask(child,child._board);
    if(mode === 'normal' && childRecurring) return myWorkHasSectionDescendant(child,sectionIds);
    if(mode === 'recurring') return currentIncluded || sectionIds.has(child.id) || myWorkHasSectionDescendant(child,sectionIds);
    return currentIncluded || sectionIds.has(child.id) || myWorkHasSectionDescendant(child,sectionIds);
  });
}

function myWorkStatusRank(t){
  const builtIn = Object.keys(STATUS);
  const idx = builtIn.indexOf(t.status);
  if(idx >= 0) return idx;
  const custom = Object.keys(t._board?.statusLabels || {}).indexOf(t.status);
  return custom >= 0 ? builtIn.length + custom : builtIn.length + 999;
}

function myWorkPriorityRank(t){
  const order = {critical:0,high:1,medium:2,low:3,none:4};
  if(order[t.priority] != null) return order[t.priority];
  const custom = Object.keys(t._board?.priorityLabels || {}).indexOf(t.priority);
  return custom >= 0 ? 5 + custom : 999;
}

function myWorkSortValue(t,sectionIds){
  const primary = myWorkPrimarySectionTask(t,sectionIds) || t;
  if(MY_WORK_SORT === 'status') return myWorkStatusRank(primary);
  if(MY_WORK_SORT === 'priority') return myWorkPriorityRank(primary);
  return validIsoDate(primary.due) || '9999-12-31';
}

function myWorkCompareTasks(a,b,sectionIds){
  const aVal = myWorkSortValue(a,sectionIds);
  const bVal = myWorkSortValue(b,sectionIds);
  let cmp = typeof aVal === 'number' || typeof bVal === 'number'
    ? (Number(aVal)||0) - (Number(bVal)||0)
    : String(aVal||'').localeCompare(String(bVal||''), undefined, {numeric:true,sensitivity:'base'});
  if(cmp) return cmp;
  cmp = String(validIsoDate(a.due) || '9999-12-31').localeCompare(String(validIsoDate(b.due) || '9999-12-31'));
  if(cmp) return cmp;
  return String(a.name||'').localeCompare(String(b.name||''), undefined, {numeric:true,sensitivity:'base'});
}

function myWorkSortedTasks(tasks,sectionIds){
  return tasks.map((task,index)=>({task,index}))
    .sort((a,b)=>{
      const cmp = myWorkCompareTasks(a.task,b.task,sectionIds);
      return cmp || a.index - b.index;
    })
    .map(x=>x.task);
}

function myWorkSectionRows(items,sectionIds,options={}){
  return myWorkSortedTasks(items,sectionIds).map(root=>{
    if(options.mode === 'normal' && isRecurringTask(root,root._board) && !sectionIds.has(root.id)){
      const children = myWorkSortedTasks(myWorkRelevantChildren(root,sectionIds,{ancestorIncluded:false,mode:'normal',today:options.today || todayIso()}),sectionIds);
      return children.map(child=>myWorkTaskRows(child,sectionIds,0,false,options)).join('');
    }
    return myWorkTaskRows(root,sectionIds,0,false,options);
  }).join('');
}

function myWorkTaskRows(t,sectionIds,depth=0,ancestorIncluded=false,options={}){
  const today = options.today || todayIso();
  const mode = options.mode || 'normal';
  const currentIncluded = ancestorIncluded || sectionIds.has(t.id);
  const children = myWorkSortedTasks(myWorkRelevantChildren(t,sectionIds,{ancestorIncluded,mode,today}),sectionIds);
  if(mode === 'normal' && isRecurringTask(t,t._board) && !sectionIds.has(t.id)){
    return children.map(child=>myWorkTaskRows(child,sectionIds,depth,currentIncluded,{mode,today})).join('');
  }
  const contextOnly = !currentIncluded && children.length > 0;
  const open = children.length > 0 && (contextOnly || areSubtasksOpen(t));
  const childRows = open ? children.map(child=>myWorkTaskRows(child,sectionIds,depth+1,currentIncluded,{mode,today})).join('') : '';
  return myWorkTaskRow(t,depth,children.length,open,contextOnly,{mode,today}) + childRows;
}

function myWorkDateLabel(t,{mode='normal',today=todayIso()}={}){
  if(mode === 'recurring'){
    if(myWorkRecurringDueOn(t,today)) return 'Recurring today';
    return recurrenceLabel(taskRecurrence(t,t._board));
  }
  return taskDueLabel(t);
}

function myWorkTaskRow(t,depth=0,childCount=0,open=false,contextOnly=false,options={}){
  const st = statusOptionFor(t._board,t.status);
  const pr = priorityOptionFor(t._board,t.priority);
  const dueLabel = myWorkDateLabel(t,options);
  const pickerKey = `${esc(t.id)}|${esc(t._group.id)}`;
  const indent = 12 + Math.max(0,Number(depth)||0) * 22;
  const toggleLabel = open ? 'Hide subtasks' : 'Show subtasks';
  const toggle = childCount && !contextOnly ? `<button type="button" class="subtask-toggle mw-subtask-toggle ${open?'open':''}" data-toggle-subtasks="${esc(t._group.id)}|${esc(t.id)}" title="${toggleLabel}" aria-label="${toggleLabel}" aria-expanded="${open?'true':'false'}">&#9656;</button>` : '<span class="subtask-spacer"></span>';
  return `<div class="mw-row ${depth?'mw-subtask-row':''} ${contextOnly?'mw-context-row':''}" data-my-work-task="${esc(t._board.id)}|${esc(t.id)}" role="button" tabindex="0" title="Open task details">
    <div class="accent mw-name-cell" style="padding-left:${indent}px;font-weight:600">
      ${toggle}
      <span class="mw-task-name">${esc(t.name)}</span>
    </div>
    <div style="color:var(--ink2)">${esc(t._group.name)}</div>
    <div style="color:var(--ink2)">${esc(t._board.name)}</div>
    <div class="mw-update-cell">${taskUpdateIndicatorHtml(t._board,t._group,t)}</div>
    <div><button type="button" class="date-cell mw-date-cell ${dueLabel?'':'empty'}" data-date-pick="${pickerKey}" title="Change due date">${I.cal}<span>${esc(dueLabel || 'Set date')}</span></button></div>
    <div><button type="button" class="pill mw-pill" style="background:${esc(st.color)}" data-status-pick="${pickerKey}" title="Change status">${esc(st.label)}</button></div>
    <div><button type="button" class="pill mw-pill" style="background:${esc(pr.color)}" data-priority-pick="${pickerKey}" title="Change priority">${esc(pr.label)}</button></div>
  </div>`;
}

function myWorkSortBar(){
  return `<div class="mw-toolbar">
    <div class="mw-help">Showing tasks assigned to <b>${esc(member(currentMemberId())?.name||'you')}</b>. Expand a task to view its subtasks.</div>
    <div class="mw-sort-tabs" aria-label="Sort My Work">
      <span>Sort by</span>
      ${MY_WORK_SORT_OPTIONS.map(opt=>`<button type="button" class="${MY_WORK_SORT===opt.id?'active':''}" data-my-work-sort="${esc(opt.id)}">${esc(opt.label)}</button>`).join('')}
    </div>
  </div>`;
}

function viewMyWork(){
  const me = currentMemberId();
  const mine = myWorkAllTasks().filter(t=>taskOwnerIds(t).includes(me));
  const todayKey = todayIso();
  const recurringAssignedIds = new Set(mine.filter(t=>isRecurringTask(t,t._board)).map(t=>t.id));
  const normalAssignedIds = new Set(mine.filter(t=>!isRecurringTask(t,t._board)).map(t=>t.id));
  const recurringRoots = myWorkRootTasks(recurringAssignedIds);
  const normalRoots = myWorkRootTasks(normalAssignedIds);
  const today = new Date();
  today.setHours(0,0,0,0);
  const buckets = {
    'Recurring':[],
    'Past Dates':[],
    'Today':[],
    'This Week':[],
    'Next Week':[],
    'Later':[],
    'Without a date':[]
  };
  recurringRoots.forEach(t=>buckets.Recurring.push(t));
  normalRoots.forEach(t=>buckets[myWorkBucketForRoot(t,today,normalAssignedIds)].push(t));
  const colors = {
    'Recurring':'#a25ddc',
    'Past Dates':'#e2445c',
    'Today':'#00c875',
    'This Week':'#0073ea',
    'Next Week':'#579bfc',
    'Later':'#fdab3d',
    'Without a date':'#5e6170'
  };
  const sections = Object.entries(buckets).map(([name,items])=>{
    const collapsed = MY_WORK_COLLAPSED.has(name);
    const mode = name === 'Recurring' ? 'recurring' : 'normal';
    const sectionIds = mode === 'recurring' ? recurringAssignedIds : normalAssignedIds;
    const body = collapsed ? '' : (items.length ? `
      <div class="mw-th"><div>Task</div><div>Group</div><div>Board</div><div>Updates</div><div>Due Date</div><div>Status</div><div>Priority</div></div>
      ${myWorkSectionRows(items,sectionIds,{mode,today:todayKey})}` : '<div class="mw-empty">No items</div>');
    return `<div class="mw-bucket">
      <button class="mw-head" data-my-work-toggle="${esc(name)}" aria-expanded="${collapsed?'false':'true'}">
        <span class="mw-caret" style="color:${colors[name]};font-weight:700;transform:rotate(${collapsed?-90:0}deg)">&#9662;</span>
        <span class="ttl" style="color:${colors[name]}">${esc(name)}</span>
        <span class="cnt">${items.length} item${items.length!==1?'s':''}</span>
      </button>
      ${body}
    </div>`;
  }).join('');
  return `${myWorkSortBar()}${sections}`;
}

/* ============================================================
   VIEW: Meetings
   ============================================================ */
function meetingCaptureStorageKey(){
  if(!SESSION?.user?.id || !WORKSPACE_ID) return '';
  return `friday_meeting_capture_${SESSION.user.id}_${WORKSPACE_ID}`;
}

function persistMeetingCapture(capture=MEETING_ACTIVE_CAPTURE,overrides={}){
  const key=meetingCaptureStorageKey();
  if(!key || !capture?.id) return;
  let existing=null;
  try{ existing=JSON.parse(Store.get(key) || 'null'); }catch(error){ existing=null; }
  if(overrides.forceLeaseOwnership !== true
    && existing?.browserInstanceId
    && existing.browserInstanceId !== MEETING_CAPTURE_BROWSER_INSTANCE_ID){
    return false;
  }
  Store.set(key,JSON.stringify({
    id:capture.id,
    sourceType:capture.sourceType,
    titleHint:capture.titleHint || '',
    meetingDateHint:capture.meetingDateHint || '',
    languageHint:capture.languageHint || '',
    startedAt:capture.startedAt || new Date().toISOString(),
    status:overrides.status || capture.status || 'recording',
    idempotencyKey:capture.idempotencyKey || '',
    clientSessionId:capture.clientSessionId || '',
    browserInstanceId:MEETING_CAPTURE_BROWSER_INSTANCE_ID,
    recoveryUploadComplete:overrides.recoveryUploadComplete ?? capture.recoveryUploadComplete === true,
    finalizationReady:overrides.finalizationReady ?? capture.finalizationReady === true,
    interruptedCoverageIncomplete:overrides.interruptedCoverageIncomplete ?? capture.interruptedCoverageIncomplete === true,
    realtimeUncertain:overrides.realtimeUncertain ?? capture.realtimeUncertain !== false,
    savedAt:new Date().toISOString()
  }));
  return true;
}

function clearPersistedMeetingCapture(){
  const key=meetingCaptureStorageKey();
  if(key) Store.set(key,'');
}

function restorePersistedMeetingCapture(){
  if(MEETING_ACTIVE_CAPTURE) return;
  const key=meetingCaptureStorageKey();
  if(!key) return;
  let saved=null;
  try{ saved=JSON.parse(Store.get(key) || 'null'); }catch(error){ saved=null; }
  const savedAt=Date.parse(saved?.savedAt || saved?.startedAt || 0);
  if(!saved?.id || !savedAt || Date.now()-savedAt > 30*60*60*1000){
    clearPersistedMeetingCapture();
    return;
  }
  const canRetryFinalization = saved.finalizationReady === true
    && saved.interruptedCoverageIncomplete !== true
    && saved.recoveryUploadComplete === true;
  if(!canRetryFinalization && saved.browserInstanceId && saved.browserInstanceId !== MEETING_CAPTURE_BROWSER_INSTANCE_ID){
    MEETING_ACTIVE_CAPTURE={
      ...saved,
      status:'external',
      segments:[],
      audioObjectPaths:[],
      pendingUploads:new Set(),
      uploadErrors:[],
      realtimeError:'This recording belongs to another TaskTrack tab or browser session. If that session was closed, wait about one minute, then clear the stopped recording safely.',
      realtimeUncertain:true,
      recoveryUploadComplete:false,
      finalizationReady:false,
      interruptedCoverageIncomplete:true,
      canRecover:false,
      mediaRecorder:null,
      peerConnection:null,
      dataChannel:null,
      stream:null,
      rawStreams:[],
      tickTimer:null,
      heartbeatTimer:null,
      realtimeCommitTimer:null,
      realtimeCompletionWaiters:new Set(),
      realtimeCompletionCount:0,
      realtimeManualCommitRequired:false,
      intentionalRealtimeClose:false,
    };
    return;
  }
  MEETING_ACTIVE_CAPTURE={
    ...saved,
    status:canRetryFinalization ? 'failed' : 'interrupted',
    segments:[],
    audioObjectPaths:[],
    pendingUploads:new Set(),
    uploadErrors:[],
    sequence:0,
    realtimeError:canRetryFinalization
      ? 'Finalization previously stopped. Retry to prepare notes from the saved audio.'
      : 'TaskTrack was closed while recording, so some audio may be missing. TaskTrack will not publish incomplete meeting notes; discard this capture and start a new recording.',
    realtimeUncertain:canRetryFinalization ? saved.realtimeUncertain !== false : true,
    recoveryUploadComplete:canRetryFinalization && saved.recoveryUploadComplete === true,
    finalizationReady:canRetryFinalization,
    interruptedCoverageIncomplete:!canRetryFinalization,
    canRecover:false,
    mediaRecorder:null,
    peerConnection:null,
    dataChannel:null,
    stream:null,
    rawStreams:[],
    tickTimer:null,
    heartbeatTimer:null,
    realtimeCommitTimer:null,
    realtimeCompletionWaiters:new Set(),
    realtimeCompletionCount:0,
    realtimeManualCommitRequired:false,
    intentionalRealtimeClose:false,
  };
}

function resetMeetingState(){
  stopMeetingSync();
  stopLocalMeetingCapture({keepDraft:true});
  MEETING_ROWS = [];
  MEETING_SELECTED_ID = '';
  MEETING_DETAIL = null;
  MEETING_LIST_LOADED = false;
  MEETING_LOADING = false;
  MEETING_SEARCH = '';
  MEETING_SHOW_ARCHIVED = false;
  MEETING_TAB = 'overview';
  MEETING_REPOSITORY = null;
  MEETING_SELECTED_ACTIONS = new Set();
  MEETING_ERROR = '';
  MEETING_PENDING_CAPTURES = [];
  MEETING_ACTIVE_CAPTURE = null;
  MEETING_DETAIL_REQUEST += 1;
}

function getMeetingRepository(){
  if(!MEETING_REPOSITORY){
    if(!supabase || !SESSION?.user?.id || !WORKSPACE_ID) throw new Error('Sign in to use Meetings.');
    MEETING_REPOSITORY = createMeetingRepository({
      supabase,
      workspaceId:WORKSPACE_ID,
      userId:SESSION.user.id
    });
  }
  return MEETING_REPOSITORY;
}

function meetingHash(id){ return `#meeting=${encodeURIComponent(id)}`; }
function meetingUrl(id){
  const url = new URL(window.location.href);
  url.hash = meetingHash(id).slice(1);
  return url.toString();
}
function syncMeetingRoute(){
  const url = new URL(window.location.href);
  url.hash = MEETING_SELECTED_ID ? meetingHash(MEETING_SELECTED_ID).slice(1) : VIEW_ROUTES.meetings;
  window.history.replaceState({},'',url.toString());
}

function openMeeting(id){
  if(!id) return;
  VIEW = 'meetings';
  MEETING_SELECTED_ID = id;
  MEETING_DETAIL = MEETING_DETAIL?.id === id ? MEETING_DETAIL : null;
  MEETING_SELECTED_ACTIONS.clear();
  const url = new URL(window.location.href);
  url.hash = meetingHash(id).slice(1);
  window.history.replaceState({}, '', url.toString());
  render();
}

async function ensureMeetingListLoaded({rerender=true,force=false}={}){
  if(!supabase || !SESSION?.user || !WORKSPACE_ID){
    MEETING_ERROR = 'Sign in to TaskTrack to use shared meeting notes.';
    MEETING_LIST_LOADED = true;
    if(rerender && VIEW === 'meetings') renderContent();
    return;
  }
  if(MEETING_LOADING || (MEETING_LIST_LOADED && !force)) return;
  MEETING_LOADING = true;
  MEETING_ERROR = '';
  if(rerender && VIEW === 'meetings') renderContent();
  try{
    const [meetingRows,pendingCaptures] = await Promise.all([
      getMeetingRepository().list({includeArchived:true}),
      loadPendingMeetingCaptures(),
    ]);
    MEETING_ROWS = meetingRows;
    MEETING_PENDING_CAPTURES = pendingCaptures;
    MEETING_LIST_LOADED = true;
    const selectedRow = MEETING_ROWS.find(item=>item.id === MEETING_SELECTED_ID);
    if(selectedRow){
      MEETING_SHOW_ARCHIVED = Boolean(selectedRow.archivedAt);
    }else{
      MEETING_SELECTED_ID = MEETING_ROWS.find(item=>
        Boolean(item.archivedAt) === MEETING_SHOW_ARCHIVED
      )?.id || '';
      MEETING_DETAIL = null;
    }
    startMeetingSync();
  }catch(error){
    MEETING_LIST_LOADED = true;
    MEETING_ERROR = isMissingRowTableError(error)
      ? 'The Meetings database migration has not been deployed yet.'
      : (error?.message || 'Could not load meeting notes.');
  }finally{
    MEETING_LOADING = false;
    if(rerender && VIEW === 'meetings') renderContent();
  }
}

async function refreshMeetingList({rerender=true}={}){
  MEETING_LIST_LOADED = false;
  return ensureMeetingListLoaded({rerender,force:true});
}

async function loadMeetingDetail(id,{rerender=true}={}){
  if(!id || !supabase || !SESSION?.user || !WORKSPACE_ID) return;
  const requestId = ++MEETING_DETAIL_REQUEST;
  MEETING_LOADING = true;
  if(rerender && VIEW === 'meetings') renderContent();
  try{
    const detail = await getMeetingRepository().load(id);
    if(requestId !== MEETING_DETAIL_REQUEST || MEETING_SELECTED_ID !== id) return;
    MEETING_DETAIL = detail;
    MEETING_ERROR = '';
  }catch(error){
    if(requestId !== MEETING_DETAIL_REQUEST || MEETING_SELECTED_ID !== id) return;
    MEETING_ERROR = error?.message || 'Could not load this meeting.';
    if(MEETING_DETAIL?.id === id) MEETING_DETAIL = null;
  }finally{
    if(requestId === MEETING_DETAIL_REQUEST){
      MEETING_LOADING = false;
      if(rerender && VIEW === 'meetings') renderContent();
    }
  }
}

function startMeetingSync(){
  if(MEETING_CHANNEL || !supabase || !WORKSPACE_ID) return;
  let channel = supabase.channel(`workspace-meetings-${WORKSPACE_ID}`);
  ['app_meetings','app_meeting_transcripts','app_meeting_sections','app_meeting_action_items','app_meeting_links','app_meeting_captures'].forEach(table=>{
    channel = channel.on('postgres_changes',{
      event:'*',
      schema:'public',
      table,
      filter:`workspace_id=eq.${WORKSPACE_ID}`
    },()=>scheduleMeetingRemoteRefresh());
  });
  MEETING_CHANNEL = channel.subscribe();
}

function stopMeetingSync(){
  if(MEETING_SYNC_TIMER) clearTimeout(MEETING_SYNC_TIMER);
  MEETING_SYNC_TIMER = null;
  if(MEETING_CHANNEL && supabase) supabase.removeChannel(MEETING_CHANNEL);
  MEETING_CHANNEL = null;
}

function scheduleMeetingRemoteRefresh(){
  if(MEETING_SYNC_TIMER) clearTimeout(MEETING_SYNC_TIMER);
  MEETING_SYNC_TIMER = setTimeout(async ()=>{
    MEETING_SYNC_TIMER = null;
    const active = document.activeElement;
    if(active?.closest?.('.meetings-page') && active.matches?.('input,textarea,select')){
      scheduleMeetingRemoteRefresh();
      return;
    }
    await refreshMeetingList({rerender:false});
    if(VIEW === 'meetings' && MEETING_SELECTED_ID) await loadMeetingDetail(MEETING_SELECTED_ID);
    else if(VIEW === 'meetings') renderContent();
  },350);
}

function viewMeetings(){
  if(!MEETING_LIST_LOADED && !MEETING_LOADING) queueMicrotask(()=>void ensureMeetingListLoaded());
  else if(
    MEETING_LIST_LOADED
    && MEETING_SELECTED_ID
    && MEETING_DETAIL?.id !== MEETING_SELECTED_ID
    && !MEETING_LOADING
  ){
    queueMicrotask(()=>void loadMeetingDetail(MEETING_SELECTED_ID));
  }
  const canDelete = Boolean(MEETING_DETAIL && (
    MEETING_DETAIL.uploaderUserId === SESSION?.user?.id || isCurrentAdmin()
  ));
  return renderMeetingsView({
    meetings:MEETING_ROWS,
    selectedId:MEETING_SELECTED_ID,
    detail:MEETING_DETAIL,
    loading:MEETING_LOADING,
    search:MEETING_SEARCH,
    showArchived:MEETING_SHOW_ARCHIVED,
    activeTab:MEETING_TAB,
    members:DATA.team || [],
    boards:DATA.boards || [],
    canPermanentlyDelete:canDelete,
    error:MEETING_ERROR,
    capture:meetingCaptureViewModel(),
    pendingCaptures:MEETING_PENDING_CAPTURES,
    googleMeetConnected:Boolean(integrationConnection('google_meet')),
    googleMeetImportEnabled:googleMeetImportIsEnabled()
  });
}

function googleMeetImportIsEnabled(){
  return ['active','pending','renewing'].includes(String(GOOGLE_MEET_SUBSCRIPTION?.state || '').toLowerCase());
}

async function loadPendingMeetingCaptures(){
  if(!supabase || !SESSION?.user || !WORKSPACE_ID) return [];
  const {data,error} = await supabase
    .from('app_meeting_captures')
    .select('id,source_type,status,meeting_id,title_hint,meeting_date_hint,language_hint,external_source_id,created_at,updated_at,heartbeat_at,source_metadata,error')
    .eq('workspace_id',WORKSPACE_ID)
    .in('status',['awaiting_confirmation','recording','paused','failed'])
    .order('updated_at',{ascending:false})
    .limit(25);
  if(error){
    if(isMissingRowTableError(error)) return [];
    throw error;
  }
  const staleBefore=Date.now()-90*1000;
  const drafts=(data || []).map(normalizeCaptureDraft).filter(draft=>{
    if(draft.status==='awaiting_confirmation') return true;
    const heartbeat=Date.parse(draft.heartbeatAt || draft.updatedAt || '');
    return ['recording','paused','failed'].includes(draft.status)
      && Number.isFinite(heartbeat)
      && heartbeat < staleBefore;
  }).map(draft=>({
    ...draft,
    canRecoverFromAudio:draft.status!=='awaiting_confirmation'
  }));
  if(MEETING_ACTIVE_CAPTURE && drafts.some(draft=>draft.id===MEETING_ACTIVE_CAPTURE.id)){
    MEETING_ACTIVE_CAPTURE=null;
    clearPersistedMeetingCapture();
  }
  return drafts;
}

async function completeSavedMeetingCapture(captureId,meetingId){
  const completionBody={action:'complete',captureId,meetingId};
  try{
    return await invokeMeetingCaptureSession(completionBody);
  }catch(firstError){
    return await invokeMeetingCaptureSession(completionBody).catch(()=>{ throw firstError; });
  }
}

async function loadMeetingCaptureDraft(captureId){
  const {data,error} = await supabase
    .from('app_meeting_captures')
    .select('*')
    .eq('workspace_id',WORKSPACE_ID)
    .eq('id',captureId)
    .single();
  if(error) throw error;
  return normalizeCaptureDraft(data);
}

function meetingCaptureViewModel(){
  const capture = MEETING_ACTIVE_CAPTURE;
  if(!capture) return null;
  const liveTranscript = canonicalCaptureTranscript((capture.segments || []).slice(-3));
  return {
    id:capture.id,
    sourceType:capture.sourceType,
    status:capture.status,
    canRecover:capture.canRecover === true && capture.interruptedCoverageIncomplete !== true,
    canRetryFinalization:capture.finalizationReady === true
      && capture.interruptedCoverageIncomplete !== true
      && (capture.recoveryUploadComplete === true || capture.realtimeUncertain === false),
    elapsedLabel:captureElapsedLabel(capture.startedAt),
    liveTranscript:liveTranscript || capture.realtimeError || ''
  };
}

function meetingCaptureBeforeUnload(event){
  if(!MEETING_ACTIVE_CAPTURE || !['recording','paused'].includes(MEETING_ACTIVE_CAPTURE.status)) return;
  event.preventDefault();
  event.returnValue = 'TaskTrack is still recording this meeting.';
}

function meetingCapturePageHidden(){
  if(MEETING_ACTIVE_CAPTURE && ['recording','paused','failed'].includes(MEETING_ACTIVE_CAPTURE.status)){
    if(['recording','paused'].includes(MEETING_ACTIVE_CAPTURE.status)){
      try{ MEETING_ACTIVE_CAPTURE.mediaRecorder?.requestData?.(); }catch(error){ /* best effort only */ }
      // The browser may terminate before the final dataavailable upload finishes.
      // Store a conservative reload marker without changing a recording that
      // continues normally after a temporary page-hide event.
      persistMeetingCapture(MEETING_ACTIVE_CAPTURE,{
        recoveryUploadComplete:false,
        finalizationReady:false,
        interruptedCoverageIncomplete:true,
        realtimeUncertain:true,
      });
      return;
    }
    persistMeetingCapture(MEETING_ACTIVE_CAPTURE);
  }
}

function meetingCaptureVisibilityChanged(){
  const capture=MEETING_ACTIVE_CAPTURE;
  if(document.visibilityState!=='visible' || !capture) return;
  if(['recording','paused'].includes(capture.status)){
    void heartbeatMeetingCapture(capture);
  }
  updateMeetingCaptureBanner();
}

function stopLocalMeetingCapture({keepDraft=true}={}){
  const capture = MEETING_ACTIVE_CAPTURE;
  if(!capture) return;
  clearInterval(capture.tickTimer);
  clearInterval(capture.heartbeatTimer);
  clearInterval(capture.realtimeCommitTimer);
  capture.tickTimer = null;
  capture.heartbeatTimer = null;
  capture.realtimeCommitTimer = null;
  capture.intentionalRealtimeClose=true;
  try{ capture.dataChannel?.close(); }catch(error){ /* no-op */ }
  try{ capture.peerConnection?.close(); }catch(error){ /* no-op */ }
  if(capture.mediaRecorder && capture.mediaRecorder.state !== 'inactive'){
    try{ capture.mediaRecorder.stop(); }catch(error){ /* no-op */ }
  }
  (capture.rawStreams || []).forEach(stream=>stream?.getTracks?.().forEach(track=>track.stop()));
  capture.stream?.getTracks?.().forEach(track=>track.stop());
  try{ void capture.audioContext?.close?.(); }catch(error){ /* no-op */ }
  window.removeEventListener('beforeunload',meetingCaptureBeforeUnload);
  if(keepDraft){
    capture.status='interrupted';
    capture.realtimeUncertain=true;
    capture.recoveryUploadComplete=false;
    capture.finalizationReady=false;
    capture.interruptedCoverageIncomplete=true;
    capture.canRecover=false;
    capture.realtimeError='Recording stopped before TaskTrack could verify complete coverage. TaskTrack will not publish incomplete meeting notes.';
    persistMeetingCapture(capture);
  }
  if(!keepDraft) MEETING_ACTIVE_CAPTURE = null;
}

function openMeetingRecordModal(){
  if(!supabase || !SESSION?.user){
    toast('Sign in to record meetings in TaskTrack');
    return;
  }
  if(MEETING_ACTIVE_CAPTURE){
    toast('TaskTrack is already recording a meeting');
    return;
  }
  if(!navigator.mediaDevices?.getUserMedia || typeof MediaRecorder === 'undefined'){
    modal('Recording is not supported','Use a current version of Chrome, Edge, or Safari with microphone access.','',closeModal,'OK');
    return;
  }
  const today = new Date().toISOString().slice(0,10);
  const body = `<div class="meeting-record-form">
    <div class="field"><label>Meeting title (optional)</label><input id="meeting-record-title" placeholder="TaskTrack will suggest one from the transcript"></div>
    <div class="meeting-confirm-grid">
      <div class="field"><label>Meeting date</label><input id="meeting-record-date" type="date" value="${today}"></div>
      <div class="field"><label>Capture source</label><select id="meeting-record-source">
        <option value="friday_mic">Room microphone</option>
        <option value="browser_tab">Google Meet tab + my microphone (desktop)</option>
      </select></div>
      <div class="field meeting-confirm-title"><label>Language hint (optional)</label><input id="meeting-record-language" placeholder="Example: English, Korean, or mixed"></div>
    </div>
    <label class="meeting-record-consent"><input id="meeting-record-consent" type="checkbox"> <span><strong>Everyone has been informed that TaskTrack will record and transcribe this meeting.</strong><small>TaskTrack will show a persistent recording indicator. Temporary audio is deleted after processing and is never kept as a recording library.</small></span></label>
    <div class="meeting-privacy-note">Keep TaskTrack open during phone recordings. For remote Google Meets, the Google Meet transcript connection is more reliable than recording the call on the same phone.</div>
    <div id="meeting-record-status" class="meeting-upload-status"></div>
  </div>`;
  modal('Record in TaskTrack','Capture an in-person meeting and prepare editable notes.',body,()=>void startMeetingCaptureFromModal(),'Start recording');
  $('#modal-root .modal')?.classList.add('modal-wide');
}

async function requestMeetingCaptureStream(sourceType){
  if(sourceType === 'browser_tab'){
    if(!navigator.mediaDevices?.getDisplayMedia) throw new Error('Browser tab capture is not available in this browser.');
    const display = await navigator.mediaDevices.getDisplayMedia({video:true,audio:true});
    const tabAudioTrack = display.getAudioTracks()[0];
    if(!tabAudioTrack){
      display.getTracks().forEach(track=>track.stop());
      throw new Error('No tab audio was shared. Choose a browser tab and enable “Share tab audio.”');
    }
    let microphone=null;
    let audioContext=null;
    try{
      microphone=await navigator.mediaDevices.getUserMedia({
        audio:{echoCancellation:true,noiseSuppression:true,autoGainControl:true},
        video:false
      });
      const AudioContextClass=window.AudioContext || window.webkitAudioContext;
      if(!AudioContextClass) throw new Error('This browser cannot mix the Meet tab with your microphone.');
      audioContext=new AudioContextClass();
      await audioContext.resume?.();
      const destination=audioContext.createMediaStreamDestination();
      const tabSource=audioContext.createMediaStreamSource(new MediaStream([tabAudioTrack]));
      const microphoneSource=audioContext.createMediaStreamSource(microphone);
      tabSource.connect(destination);
      microphoneSource.connect(destination);
      return {
        stream:destination.stream,
        rawStreams:[display,microphone],
        audioContext,
        audioNodes:[tabSource,microphoneSource,destination],
        criticalTracks:[tabAudioTrack,...microphone.getAudioTracks()]
      };
    }catch(error){
      microphone?.getTracks?.().forEach(track=>track.stop());
      display.getTracks().forEach(track=>track.stop());
      try{ await audioContext?.close?.(); }catch(closeError){ /* no-op */ }
      throw error;
    }
  }
  const microphone = await navigator.mediaDevices.getUserMedia({
    audio:{echoCancellation:true,noiseSuppression:true,autoGainControl:true},
    video:false
  });
  return {stream:microphone,rawStreams:[microphone],audioContext:null,audioNodes:[],criticalTracks:microphone.getAudioTracks()};
}

function bindMeetingCaptureSourceGuards(capture){
  (capture?.criticalTracks || []).forEach(track=>{
    const handleEnded=()=>{
      if(MEETING_ACTIVE_CAPTURE!==capture || !['recording','paused'].includes(capture.status) || capture.sourceLossHandled) return;
      capture.sourceLossHandled=true;
      stopLocalMeetingCapture({keepDraft:true});
      capture.status='interrupted';
      capture.realtimeUncertain=true;
      capture.recoveryUploadComplete=false;
      capture.finalizationReady=false;
      capture.interruptedCoverageIncomplete=true;
      capture.canRecover=false;
      const sourceLabel=capture.sourceType === 'browser_tab' ? 'Google Meet tab sharing or microphone access' : 'Microphone access';
      capture.realtimeError=`${sourceLabel} stopped, so part of the conversation may be missing. TaskTrack will not publish incomplete notes; discard this capture and start a new recording.`;
      persistMeetingCapture(capture);
      render();
      toast(`${sourceLabel} stopped; this capture cannot produce complete notes`);
      void invokeMeetingCaptureSession({
        action:'pause',
        captureId:capture.id,
        clientSessionId:capture.clientSessionId,
      }).catch(()=>{});
    };
    track.addEventListener?.('ended',handleEnded,{once:true});
    if(track.readyState==='ended') queueMicrotask(handleEnded);
  });
}

async function invokeMeetingCaptureSession(body){
  const {data,error} = await supabase.functions.invoke('meeting-capture-session',{
    body:{workspaceId:WORKSPACE_ID,memberId:currentMemberId(),...body}
  });
  if(error || data?.error) throw new Error(await assistantFunctionErrorMessage(error,data) || 'Could not start the meeting capture.');
  return data?.result || data || {};
}

async function clearStoppedExternalMeetingCapture(){
  const capture=MEETING_ACTIVE_CAPTURE;
  if(!capture || capture.status!=='external') return;
  const clientSessionId=globalThis.crypto?.randomUUID?.() || uid('client_session_');
  try{
    await invokeMeetingCaptureSession({
      action:'start',
      captureId:capture.id,
      idempotencyKey:capture.idempotencyKey || '',
      sourceType:capture.sourceType,
      clientSessionId,
    });
    capture.clientSessionId=clientSessionId;
    try{
      await invokeMeetingCaptureSession({action:'cancel',captureId:capture.id,clientSessionId});
    }catch(error){
      capture.status='interrupted';
      capture.realtimeError='TaskTrack reclaimed the stopped recording but could not confirm cleanup. Use Cancel again.';
      persistMeetingCapture(capture,{forceLeaseOwnership:true});
      render();
      throw error;
    }
    MEETING_ACTIVE_CAPTURE=null;
    clearPersistedMeetingCapture();
    render();
    toast('Stopped recording cleared');
  }catch(error){
    if(capture.status==='external' && meetingCaptureLeaseWasLost(error)){
      toast('That recording is still active in another TaskTrack tab. Try again after it stops or after about one minute.');
      return;
    }
    toast(error?.message || 'TaskTrack could not clear the stopped recording yet');
  }
}

function meetingCaptureLeaseWasLost(error){
  return /active in another browser session|capture_active_elsewhere/i.test(String(error?.message || error || ''));
}

function handleMeetingCaptureLeaseLoss(capture){
  if(!capture || MEETING_ACTIVE_CAPTURE!==capture) return;
  stopLocalMeetingCapture({keepDraft:true});
  capture.status='external';
  capture.realtimeUncertain=true;
  capture.realtimeError='This recording continued in another TaskTrack tab, so this tab stopped using the microphone and temporary audio. Clear it here only after the other tab has stopped.';
  persistMeetingCapture(capture);
  render();
  toast('Recording controls moved to another TaskTrack tab');
}

async function heartbeatMeetingCapture(capture){
  if(!capture?.id || !capture.clientSessionId || MEETING_ACTIVE_CAPTURE!==capture) return;
  try{
    await invokeMeetingCaptureSession({
      action:'heartbeat',
      captureId:capture.id,
      clientSessionId:capture.clientSessionId,
    });
    capture.heartbeatFailures=0;
  }catch(error){
    capture.heartbeatFailures=Number(capture.heartbeatFailures || 0)+1;
    if(meetingCaptureLeaseWasLost(error)) handleMeetingCaptureLeaseLoss(capture);
  }
}

async function startMeetingCaptureFromModal(){
  if(!$('#meeting-record-consent')?.checked){
    toast('Confirm that everyone has been informed before recording');
    return;
  }
  const sourceType = $('#meeting-record-source')?.value === 'browser_tab' ? 'browser_tab' : 'friday_mic';
  const titleHint = $('#meeting-record-title')?.value.trim() || '';
  const meetingDateHint = $('#meeting-record-date')?.value || new Date().toISOString().slice(0,10);
  const languageHint = $('#meeting-record-language')?.value.trim() || '';
  const ok = $('#modal-root [data-mc="ok"]');
  const cancel = $('#modal-root [data-mc="cancel"]');
  const status = $('#meeting-record-status');
  const idempotencyKey=`recording:${globalThis.crypto?.randomUUID?.() || uid('capture_')}`;
  const clientSessionId=globalThis.crypto?.randomUUID?.() || uid('client_session_');
  if(ok){ ok.disabled=true; ok.textContent='Requesting microphone…'; }
  if(cancel) cancel.disabled=true;
  let media=null;
  try{
    media = await requestMeetingCaptureStream(sourceType);
    if(ok) ok.textContent='Starting transcription…';
    let session;
    try{
      const startBody={action:'start',sourceType,titleHint,meetingDateHint,languageHint,consentConfirmed:true,idempotencyKey,clientSessionId};
      try{
        session = await invokeMeetingCaptureSession(startBody);
      }catch(firstError){
        // Retrying the same idempotency key safely recovers a response that was
        // lost after the server had already created the capture.
        session = await invokeMeetingCaptureSession(startBody).catch(()=>{ throw firstError; });
      }
    }catch(error){
      media.rawStreams.forEach(stream=>stream.getTracks().forEach(track=>track.stop()));
      try{ await media.audioContext?.close?.(); }catch(closeError){ /* no-op */ }
      throw error;
    }
    const captureRow = session.capture || session.captureRow || session;
    const captureId = session.captureId || captureRow.id;
    if(!captureId) throw new Error('TaskTrack did not receive a capture identifier.');
    MEETING_ACTIVE_CAPTURE = {
      id:captureId,
      sourceType,
      status:'recording',
      titleHint,
      meetingDateHint,
      languageHint,
      startedAt:captureRow.started_at || captureRow.startedAt || new Date().toISOString(),
      idempotencyKey,
      clientSessionId,
      chunkSessionId:`${Date.now()}-${clientSessionId}`,
      stream:media.stream,
      rawStreams:media.rawStreams,
      audioContext:media.audioContext || null,
      audioNodes:media.audioNodes || [],
      criticalTracks:media.criticalTracks || [],
      segments:[],
      audioObjectPaths:[],
      pendingUploads:new Set(),
      uploadErrors:[],
      recoveryUploadComplete:true,
      finalizationReady:false,
      interruptedCoverageIncomplete:false,
      canRecover:false,
      sequence:0,
      realtimeError:'',
      realtimeUncertain:true,
      hasInterruptedCoverage:false,
      mediaRecorder:null,
      peerConnection:null,
      dataChannel:null,
      tickTimer:null,
      heartbeatTimer:null,
      realtimeCommitTimer:null,
      realtimeCompletionWaiters:new Set(),
      realtimeCompletionCount:0,
      realtimeManualCommitRequired:false,
      intentionalRealtimeClose:false,
    };
    persistMeetingCapture(MEETING_ACTIVE_CAPTURE,{forceLeaseOwnership:true});
    startMeetingMediaRecorder(MEETING_ACTIVE_CAPTURE);
    bindMeetingCaptureSourceGuards(MEETING_ACTIVE_CAPTURE);
    const realtime = session.realtime || session.transcriptionSession || session.session || null;
    if(realtime) void connectMeetingRealtime(MEETING_ACTIVE_CAPTURE,realtime).catch(error=>{
      if(!MEETING_ACTIVE_CAPTURE || MEETING_ACTIVE_CAPTURE.id !== captureId) return;
      MEETING_ACTIVE_CAPTURE.realtimeUncertain=true;
      MEETING_ACTIVE_CAPTURE.realtimeError = `Live captions unavailable: ${error?.message || 'connection failed'}. Audio recovery is still recording.`;
      updateMeetingCaptureBanner();
    });
    MEETING_ACTIVE_CAPTURE.tickTimer = setInterval(updateMeetingCaptureBanner,1000);
    MEETING_ACTIVE_CAPTURE.heartbeatTimer = setInterval(()=>{
      if(!MEETING_ACTIVE_CAPTURE || MEETING_ACTIVE_CAPTURE.id !== captureId) return;
      void heartbeatMeetingCapture(MEETING_ACTIVE_CAPTURE);
    },30000);
    window.addEventListener('beforeunload',meetingCaptureBeforeUnload);
    closeModal();
    render();
    toast('TaskTrack is recording');
  }catch(error){
    const activeCapture=MEETING_ACTIVE_CAPTURE;
    if(activeCapture){
      stopLocalMeetingCapture({keepDraft:true});
      try{
        await invokeMeetingCaptureSession({
          action:'cancel',
          captureId:activeCapture.id,
          clientSessionId:activeCapture.clientSessionId,
        });
        MEETING_ACTIVE_CAPTURE=null;
        clearPersistedMeetingCapture();
      }catch(cleanupError){
        activeCapture.status='interrupted';
        activeCapture.realtimeError='TaskTrack could not finish starting or confirm server cleanup. The capture remains recoverable instead of being hidden.';
        MEETING_ACTIVE_CAPTURE=activeCapture;
        persistMeetingCapture(activeCapture);
      }
    }else if(media){
      media.rawStreams?.forEach(stream=>stream.getTracks().forEach(track=>track.stop()));
      media.stream?.getTracks?.().forEach(track=>track.stop());
      try{ await media.audioContext?.close?.(); }catch(closeError){ /* no-op */ }
    }
    if(ok){ ok.disabled=false; ok.textContent='Start recording'; }
    if(cancel) cancel.disabled=false;
    if(status) status.textContent=error?.message || 'Could not start recording.';
  }
}

function preferredMeetingRecordingMimeType(){
  return ['audio/webm;codecs=opus','audio/webm','audio/mp4']
    .find(type=>MediaRecorder.isTypeSupported?.(type)) || '';
}

function startMeetingMediaRecorder(capture){
  const mimeType = preferredMeetingRecordingMimeType();
  // Keep the encrypted recovery recording small enough to reassemble into one
  // Audio API request after a long meeting (roughly 22 MB for 90 minutes).
  const recorderOptions = {audioBitsPerSecond:32000};
  if(mimeType) recorderOptions.mimeType=mimeType;
  const recorder = new MediaRecorder(capture.stream,recorderOptions);
  capture.mediaRecorder = recorder;
  capture.finalizationReady=false;
  recorder.addEventListener('dataavailable',event=>{
    if(!event.data?.size) return;
    const upload = uploadMeetingCaptureChunk(capture,event.data).catch(error=>{
      capture.recoveryUploadComplete=false;
      capture.uploadErrors.push(error?.message || 'Audio recovery chunk upload failed.');
    }).finally(()=>capture.pendingUploads.delete(upload));
    capture.pendingUploads.add(upload);
  });
  recorder.addEventListener('error',event=>{
    capture.recoveryUploadComplete=false;
    capture.uploadErrors.push(event.error?.message || 'The browser recorder reported an error.');
  });
  recorder.addEventListener('stop',()=>{
    if(MEETING_ACTIVE_CAPTURE!==capture || !['recording','paused'].includes(capture.status)) return;
    capture.status='failed';
    capture.realtimeUncertain=true;
    capture.recoveryUploadComplete=false;
    capture.finalizationReady=false;
    capture.interruptedCoverageIncomplete=true;
    capture.canRecover=false;
    capture.realtimeError='The browser stopped recording unexpectedly, so TaskTrack cannot verify complete coverage and will not publish incomplete meeting notes.';
    persistMeetingCapture(capture);
    render();
  });
  recorder.start(15000);
}

async function uploadMeetingCaptureChunk(capture,blob){
  const sequence = capture.sequence++;
  const extension = blob.type.includes('mp4') ? 'm4a' : 'webm';
  const sessionId=String(capture.chunkSessionId || `${Date.now()}-${uid('session_')}`).replace(/[^a-zA-Z0-9_-]/g,'');
  const path = `${WORKSPACE_ID}/${capture.id}/${sessionId}-${String(sequence).padStart(6,'0')}.${extension}`;
  const contentType=String(blob.type || 'audio/webm').split(';')[0].trim() || 'audio/webm';
  const {error} = await supabase.storage.from(MEETING_CAPTURE_BUCKET).upload(path,blob,{
    contentType,
    upsert:false
  });
  if(error) throw error;
  capture.audioObjectPaths.push(path);
}

function realtimeClientSecret(realtime){
  return realtime?.clientSecret?.value
    || realtime?.client_secret?.value
    || realtime?.clientSecret
    || realtime?.client_secret
    || realtime?.ephemeralKey
    || realtime?.value
    || '';
}

function markMeetingRealtimeUncertain(capture,message='Live transcription was interrupted. TaskTrack will verify the transcript from saved audio.'){
  if(!capture || capture.intentionalRealtimeClose) return;
  capture.realtimeUncertain=true;
  capture.realtimeError=message;
  updateMeetingCaptureBanner();
}

function resolveMeetingRealtimeCompletion(capture){
  capture.realtimeCompletionCount=Number(capture.realtimeCompletionCount || 0)+1;
  for(const waiter of capture.realtimeCompletionWaiters || []) waiter(true);
  capture.realtimeCompletionWaiters?.clear?.();
}

function waitForMeetingRealtimeCompletion(capture,previousCount,timeoutMs=8000){
  if(Number(capture?.realtimeCompletionCount || 0)>previousCount) return Promise.resolve(true);
  return new Promise(resolve=>{
    let settled=false;
    const finish=value=>{
      if(settled) return;
      settled=true;
      clearTimeout(timer);
      capture?.realtimeCompletionWaiters?.delete?.(finish);
      resolve(value);
    };
    const timer=setTimeout(()=>finish(false),timeoutMs);
    capture?.realtimeCompletionWaiters?.add?.(finish);
  });
}

async function connectMeetingRealtime(capture,realtime){
  const secret = realtimeClientSecret(realtime);
  if(!secret) throw new Error('No temporary transcription token was returned.');
  const pc = new RTCPeerConnection();
  capture.peerConnection = pc;
  capture.stream.getAudioTracks().forEach(track=>pc.addTrack(track,capture.stream));
  const channel = pc.createDataChannel('oai-events');
  capture.dataChannel = channel;
  capture.realtimeManualCommitRequired=Boolean(realtime.manualCommitRequired);
  pc.addEventListener('connectionstatechange',()=>{
    if(['failed','disconnected','closed'].includes(pc.connectionState)){
      markMeetingRealtimeUncertain(capture,'Live transcription disconnected. TaskTrack will verify the transcript from saved audio.');
    }
  });
  channel.addEventListener('open',()=>{
    if(!capture.hasInterruptedCoverage) capture.realtimeUncertain=false;
    if(!capture.realtimeManualCommitRequired) return;
    clearInterval(capture.realtimeCommitTimer);
    capture.realtimeCommitTimer=setInterval(()=>sendMeetingRealtimeCommit(capture),8000);
  });
  channel.addEventListener('message',event=>{
    let payload;
    try{ payload=JSON.parse(event.data); }catch(error){ return; }
    const segment = realtimeTranscriptEventToSegment(payload,capture.segments.length);
    if(!segment) return;
    capture.segments = appendCaptureSegment(capture.segments,segment);
    resolveMeetingRealtimeCompletion(capture);
    updateMeetingCaptureBanner();
  });
  channel.addEventListener('error',()=>markMeetingRealtimeUncertain(capture));
  channel.addEventListener('close',()=>markMeetingRealtimeUncertain(capture));
  const offer = await pc.createOffer();
  await pc.setLocalDescription(offer);
  const endpoint = realtime.endpoint || realtime.url || 'https://api.openai.com/v1/realtime/calls';
  const response = await fetch(endpoint,{
    method:'POST',
    headers:{Authorization:`Bearer ${secret}`,'Content-Type':'application/sdp'},
    body:offer.sdp
  });
  if(!response.ok) throw new Error((await response.text().catch(()=>'')) || 'OpenAI Realtime rejected the connection.');
  await pc.setRemoteDescription({type:'answer',sdp:await response.text()});
}

function sendMeetingRealtimeCommit(capture,{force=false}={}){
  if(!capture?.realtimeManualCommitRequired || capture.dataChannel?.readyState !== 'open') return false;
  if(!force && capture.status !== 'recording') return false;
  try{
    capture.dataChannel.send(JSON.stringify({type:'input_audio_buffer.commit'}));
    return true;
  }catch(error){
    markMeetingRealtimeUncertain(capture,'TaskTrack could not confirm the latest live transcript. Saved audio will be used for verification.');
    return false;
  }
}

function updateMeetingCaptureBanner(){
  const capture = MEETING_ACTIVE_CAPTURE;
  if(!capture) return;
  const elapsed = document.querySelector('.meeting-capture-copy>div>span');
  const transcript = document.querySelector('.meeting-capture-copy>p');
  if(elapsed) elapsed.textContent=captureElapsedLabel(capture.startedAt);
  if(transcript){
    transcript.textContent=canonicalCaptureTranscript((capture.segments || []).slice(-3))
      || capture.realtimeError
      || (capture.status === 'paused' ? 'Recording paused.' : 'Listening for speech…');
  }
}

async function pauseMeetingCapture(paused=true){
  const capture = MEETING_ACTIVE_CAPTURE;
  if(!capture || !['recording','paused'].includes(capture.status)) return;
  try{
    if(paused) sendMeetingRealtimeCommit(capture,{force:true});
    if(paused && capture.mediaRecorder?.state === 'recording') capture.mediaRecorder.pause();
    if(!paused && capture.mediaRecorder?.state === 'paused') capture.mediaRecorder.resume();
    capture.stream.getAudioTracks().forEach(track=>{ track.enabled=!paused; });
    capture.status=paused?'paused':'recording';
    await invokeMeetingCaptureSession({action:paused?'pause':'resume',captureId:capture.id,clientSessionId:capture.clientSessionId});
    persistMeetingCapture(capture);
    render();
  }catch(error){
    if(meetingCaptureLeaseWasLost(error)){
      handleMeetingCaptureLeaseLoss(capture);
      return;
    }
    toast(error?.message || `Could not ${paused?'pause':'resume'} recording`);
  }
}

async function stopMeetingMediaRecorder(capture){
  const recorder = capture.mediaRecorder;
  if(!recorder || recorder.state === 'inactive') return;
  await new Promise(resolve=>{
    recorder.addEventListener('stop',resolve,{once:true});
    try{ recorder.requestData(); }catch(error){ /* no-op */ }
    recorder.stop();
  });
}

async function stopAndFinalizeMeetingCapture(){
  const capture = MEETING_ACTIVE_CAPTURE;
  if(!capture || !['recording','paused','failed'].includes(capture.status)) return;
  capture.status='finalizing';
  capture.finalizationReady=false;
  clearInterval(capture.tickTimer);
  clearInterval(capture.heartbeatTimer);
  clearInterval(capture.realtimeCommitTimer);
  capture.realtimeCommitTimer=null;
  render();
  try{
    await stopMeetingMediaRecorder(capture);
    if(capture.realtimeManualCommitRequired){
      const previousCompletionCount=Number(capture.realtimeCompletionCount || 0);
      const finalCommitSent=sendMeetingRealtimeCommit(capture,{force:true});
      const finalCompletionReceived=finalCommitSent
        ? await waitForMeetingRealtimeCompletion(capture,previousCompletionCount)
        : false;
      if(!finalCompletionReceived){
        capture.realtimeUncertain=true;
        capture.realtimeError='The final live transcript was not confirmed in time. TaskTrack will verify it from saved audio.';
      }
    }
    capture.intentionalRealtimeClose=true;
    try{ capture.dataChannel?.close(); }catch(error){ /* no-op */ }
    try{ capture.peerConnection?.close(); }catch(error){ /* no-op */ }
    (capture.rawStreams || []).forEach(stream=>stream.getTracks().forEach(track=>track.stop()));
    capture.stream?.getTracks?.().forEach(track=>track.stop());
    try{ await capture.audioContext?.close?.(); }catch(error){ /* no-op */ }
    window.removeEventListener('beforeunload',meetingCaptureBeforeUnload);
    await Promise.all([...capture.pendingUploads]);
    capture.recoveryUploadComplete=capture.interruptedCoverageIncomplete !== true && capture.uploadErrors.length===0;
    capture.finalizationReady=true;
    persistMeetingCapture(capture);
    if(capture.realtimeUncertain && !capture.recoveryUploadComplete){
      throw new Error('Live transcription and the temporary audio backup were both interrupted. TaskTrack kept the draft for review instead of creating incomplete notes.');
    }
    const transcript = canonicalCaptureTranscript(capture.segments);
    const {data,error} = await supabase.functions.invoke('meeting-capture-finalize',{
      body:{
        workspaceId:WORKSPACE_ID,
        captureId:capture.id,
        clientSessionId:capture.clientSessionId || '',
        transcript,
        segments:capture.segments,
        audioObjectPaths:capture.audioObjectPaths,
        forceAudioRecovery:capture.realtimeUncertain===true,
        recoveryUploadComplete:capture.recoveryUploadComplete===true,
        titleHint:capture.titleHint,
        meetingDateHint:capture.meetingDateHint,
        languageHint:capture.languageHint
      }
    });
    if(error || data?.error) throw new Error(await assistantFunctionErrorMessage(error,data) || 'Could not finalize this recording.');
    let result = data?.result || data || {};
    let captureDraft = normalizeCaptureDraft(result.capture || result.captureRow || result);
    if(!captureDraft.analysis?.summary && !captureDraft.transcript){
      captureDraft = await loadMeetingCaptureDraft(capture.id);
    }
    const analysis = await ensureGeneratedMeetingAnalysis({
      ...(captureDraft.analysis || result.analysis || {}),
      transcript:captureDraft.transcript || result.transcript || transcript
    },{
      title:captureDraft.analysis?.title || result.analysis?.title || capture.titleHint,
      meetingDate:captureDraft.analysis?.meetingDate || result.analysis?.meetingDate || capture.meetingDateHint,
      language:captureDraft.analysis?.language || result.analysis?.language || capture.languageHint || 'Unknown'
    });
    MEETING_ACTIVE_CAPTURE=null;
    clearPersistedMeetingCapture();
    await refreshMeetingList({rerender:false});
    render();
    openMeetingConfirmationModal(analysis,'TaskTrack recording',{
      sourceType:capture.sourceType,
      captureId:capture.id,
      externalSourceId:''
    });
  }catch(error){
    if(meetingCaptureLeaseWasLost(error)){
      handleMeetingCaptureLeaseLoss(capture);
      return;
    }
    capture.status='failed';
    capture.realtimeError=error?.message || 'Could not finalize the recording.';
    persistMeetingCapture(capture);
    render();
    const canRetry = capture.finalizationReady === true
      && capture.interruptedCoverageIncomplete !== true
      && (capture.recoveryUploadComplete === true || capture.realtimeUncertain === false);
    if(canRetry){
      modal('Recording needs attention',`${capture.realtimeError} TaskTrack will keep encrypted recovery audio for up to 24 hours so you can retry.`,'',()=>{ closeModal(); void stopAndFinalizeMeetingCapture(); },'Retry finalization');
    }else{
      modal('Recording incomplete',`${capture.realtimeError} Cancel this capture and start a new recording so TaskTrack does not publish incomplete notes.`,'',closeModal,'Close');
    }
  }
}

function confirmCancelMeetingCapture(){
  if(!MEETING_ACTIVE_CAPTURE) return;
  confirmModal('Cancel this recording?','The capture will be discarded and temporary audio will be queued for deletion.',async ()=>{
    const capture = MEETING_ACTIVE_CAPTURE;
    stopLocalMeetingCapture({keepDraft:true});
    try{
      await invokeMeetingCaptureSession({action:'cancel',captureId:capture.id,clientSessionId:capture.clientSessionId});
    }catch(error){
      closeModal();
      if(meetingCaptureLeaseWasLost(error)){
        handleMeetingCaptureLeaseLoss(capture);
        return;
      }
      capture.status='interrupted';
      capture.realtimeError='The browser stopped recording, but TaskTrack could not confirm server cleanup. Try Cancel again; automatic cleanup remains the fallback.';
      MEETING_ACTIVE_CAPTURE=capture;
      persistMeetingCapture(capture);
      render();
      toast(error?.message || 'Server cleanup was not confirmed');
      return;
    }
    MEETING_ACTIVE_CAPTURE=null;
    clearPersistedMeetingCapture();
    closeModal();
    render();
    toast('Recording cancelled');
  });
}

async function reviewPendingMeetingCapture(captureId){
  try{
    let capture = await loadMeetingCaptureDraft(captureId);
    if(['recording','paused','failed'].includes(capture.status) && !capture.transcript){
      toast('Recovering the saved audio and preparing meeting notes…');
      const {data,error}=await supabase.functions.invoke('meeting-capture-finalize',{
        body:{
          workspaceId:WORKSPACE_ID,
          captureId:capture.id,
          clientSessionId:capture.clientSessionId,
          transcript:'',
          segments:[],
          audioObjectPaths:capture.audioObjectPaths,
          forceAudioRecovery:true,
          recoveryUploadComplete:true,
          titleHint:capture.titleHint,
          meetingDateHint:capture.meetingDateHint,
          languageHint:capture.languageHint
        }
      });
      if(error || data?.error) throw new Error(await assistantFunctionErrorMessage(error,data) || 'Could not recover this recording.');
      capture=normalizeCaptureDraft(data?.result?.capture || data?.result?.captureRow || data?.capture || data);
      if(!capture.transcript) capture=await loadMeetingCaptureDraft(captureId);
      MEETING_PENDING_CAPTURES=await loadPendingMeetingCaptures();
    }
    if(capture.meetingId){
      const savedMeeting=await getMeetingRepository().load(capture.meetingId).catch(()=>null);
      if(savedMeeting?.contentSavedAt){
        await completeSavedMeetingCapture(capture.id,capture.meetingId);
        MEETING_PENDING_CAPTURES=MEETING_PENDING_CAPTURES.filter(item=>item.id!==capture.id);
        await refreshMeetingList({rerender:false});
        openMeeting(capture.meetingId);
        toast('Meeting save confirmed and temporary transcript copy deleted');
        return;
      }
    }
    let analysis = {
      ...(capture.analysis || {}),
      transcript:capture.transcript,
      title:capture.analysis?.title || capture.titleHint,
      meetingDate:capture.analysis?.meetingDate || capture.meetingDateHint,
      language:capture.analysis?.language || capture.languageHint || 'Unknown'
    };
    if(!analysis.transcript) throw new Error('This meeting draft does not have a finalized transcript yet.');
    if(!validateGeneratedMeetingAnalysis(analysis).valid){
      toast('Transcript ready. Generating meeting notes…');
      analysis = await ensureGeneratedMeetingAnalysis(analysis);
    }
    openMeetingConfirmationModal(analysis,capture.sourceType === 'google_meet_transcript' ? 'Google Meet' : 'TaskTrack recording',{
      sourceType:capture.sourceType,
      captureId:capture.id,
      meetingId:capture.meetingId,
      externalSourceId:capture.externalSourceId
    });
  }catch(error){
    toast(error?.message || 'Could not open this meeting draft');
  }
}

function discardPendingMeetingCapture(captureId){
  confirmModal('Discard this meeting draft?','The transcript draft and any meeting-specific action proposals will be permanently removed.',async ()=>{
    try{
      await invokeMeetingCaptureSession({action:'discard',captureId});
      MEETING_PENDING_CAPTURES=MEETING_PENDING_CAPTURES.filter(item=>item.id !== captureId);
      closeModal();
      renderContent();
      toast('Meeting draft discarded');
    }catch(error){
      toast(error?.message || 'Could not discard this meeting draft');
    }
  });
}

function openGoogleMeetSetupModal(){
  if(!integrationConnection('google_meet')){
    startIntegrationOauth('google_meet');
    return;
  }
  const enabled=googleMeetImportIsEnabled();
  const body = `<div class="meeting-google-setup">
    <div class="integration-note">TaskTrack will automatically import every completed transcript for Google Meets this connected account owns or organizes. Transcript text is sent to OpenAI to prepare the draft. Until you confirm it, only you and workspace administrators can see it.</div>
    <ul><li>No Google Drive recording is downloaded or retained.</li><li>Duplicate Google events cannot create duplicate meetings.</li><li>Google transcription must be available and enabled for the meeting.</li><li>You can stop future imports at any time.</li></ul>
    ${enabled?'<button type="button" class="btn ghost danger" id="meeting-google-disable">Stop automatic imports</button>':''}
    <div id="meeting-google-status" class="meeting-upload-status"></div>
  </div>`;
  modal('Automatic Google Meet imports',enabled?'Imports are currently enabled for this account.':'Subscribe this Google account to TaskTrack.',body,()=>void enableGoogleMeetImport(),enabled?'Renew subscription':'Enable automatic imports');
  const disable=$('#meeting-google-disable');
  if(disable) disable.onclick=()=>void disableGoogleMeetImport();
}

async function enableGoogleMeetImport(){
  const ok = $('#modal-root [data-mc="ok"]');
  const status = $('#meeting-google-status');
  if(ok){ ok.disabled=true; ok.textContent='Connecting…'; }
  try{
    const {data,error} = await supabase.functions.invoke('google-meet-subscribe',{body:{workspaceId:WORKSPACE_ID}});
    if(error || data?.error) throw new Error(await assistantFunctionErrorMessage(error,data) || 'Could not enable Google Meet imports.');
    await refreshIntegrationConnections({rerender:false});
    closeModal();
    toast('Automatic Google Meet transcript imports enabled');
  }catch(error){
    if(ok){ ok.disabled=false; ok.textContent='Enable automatic imports'; }
    if(status) status.textContent=error?.message || 'Google Meet setup is incomplete.';
  }
}

async function disableGoogleMeetImport(){
  const button=$('#meeting-google-disable');
  const status=$('#meeting-google-status');
  if(button){ button.disabled=true; button.textContent='Stopping imports…'; }
  try{
    const {data,error}=await supabase.functions.invoke('google-meet-subscribe',{
      body:{workspaceId:WORKSPACE_ID,action:'disable'}
    });
    if(error || data?.error) throw new Error(await assistantFunctionErrorMessage(error,data) || 'Could not stop Google Meet imports.');
    await refreshIntegrationConnections({rerender:false});
    closeModal();
    renderContent();
    toast('Automatic Google Meet imports stopped');
  }catch(error){
    if(button){ button.disabled=false; button.textContent='Stop automatic imports'; }
    if(status) status.textContent=error?.message || 'Could not stop automatic imports.';
  }
}

function openMeetingUploadModal(){
  if(false){
    toast('Sign in to upload meeting transcripts');
    return;
  }
  const body = `
    <div class="meeting-upload-form">
      <label class="meeting-file-drop">
        <strong>Choose a transcript file</strong>
        <span>TXT, VTT, SRT, DOC, DOCX, RTF, ODT, or PDF · 10 MB maximum</span>
        <input id="meeting-upload-file" type="file" accept=".txt,.text,.md,.vtt,.srt,.doc,.docx,.rtf,.odt,.pdf">
      </label>
      <div class="meeting-upload-or"><span>or paste the transcript</span></div>
      <textarea id="meeting-upload-text" rows="12" placeholder="Paste the complete transcript here…"></textarea>
      <div class="meeting-upload-status" id="meeting-upload-status"></div>
      <p class="meeting-privacy-note">Simulation: pasted text stays on this computer. The sample produces an editable draft using fixed local rules. No AI service receives the transcript.</p>
    </div>`;
  modal('Upload meeting transcript','TaskTrack will immediately generate editable meeting notes from the transcript.',body,()=>void processMeetingUpload(),'Upload & generate notes');
  $('#modal-root .modal')?.classList.add('modal-wide');
}

async function meetingFilePayload(file){
  if(!file) return null;
  const extension = String(file.name || '').toLowerCase().split('.').pop() || '';
  const mimeTypes = {
    txt:'text/plain',
    text:'text/plain',
    md:'text/markdown',
    vtt:'text/vtt',
    srt:'application/x-subrip',
    doc:'application/msword',
    docx:'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    rtf:'application/rtf',
    odt:'application/vnd.oasis.opendocument.text',
    pdf:'application/pdf'
  };
  const mimeType = file.type || mimeTypes[extension] || 'application/octet-stream';
  const dataUrl = await new Promise((resolve,reject)=>{
    const reader = new FileReader();
    reader.onload=()=>resolve(String(reader.result || '').replace(/^data:;base64,/i,`data:${mimeType};base64,`));
    reader.onerror=()=>reject(reader.error || new Error('Could not read the transcript file.'));
    reader.readAsDataURL(file);
  });
  return {name:file.name || 'transcript.txt',mimeType,dataUrl};
}

async function invokeMeetingTranscript({transcript='',file=null,mode='analyze',outputLanguage='original'}={}){
if(file && !transcript){if(!/text|vtt|subrip|markdown/i.test(file.mimeType||''))throw new Error('Local simulation supports pasted text and TXT/VTT/SRT files. PDF and Office extraction need the external service.');transcript=decodeURIComponent(escape(atob(String(file.dataUrl).split(',')[1]||'')));}
  if(!transcript.trim())throw new Error('Paste a transcript first.');
  return {transcript,title:'Simulated meeting draft',meetingDate:new Date().toISOString().slice(0,10),language:'English',summary:'Local simulation: review and edit this draft against the transcript.',keyDiscussionPoints:transcript.split(/\n/).filter(Boolean).slice(0,5),decisions:[],actionItems:[],openQuestions:['Review the transcript and add confirmed decisions and owners.'],attendees:[]};

  const {data,error} = await supabase.functions.invoke('meeting-transcript',{
    body:{
      workspaceId:WORKSPACE_ID,
      transcript,
      file,
      mode,
      outputLanguage
    }
  });
  if(error || data?.error || !data?.result){
    throw new Error(await assistantFunctionErrorMessage(error,data) || 'Could not process this transcript.');
  }
  const validation = validateGeneratedMeetingAnalysis(data.result);
  if(!validation.valid){
    throw new Error(`TaskTrack received the transcript, but the AI draft was incomplete (${validation.missing.join(', ')}). Nothing was saved. Please try again.`);
  }
  return data.result;
}

async function ensureGeneratedMeetingAnalysis(analysis={},fallbacks={}){
  const candidate = {
    ...analysis,
    transcript:String(analysis.transcript || fallbacks.transcript || '').trim(),
    title:analysis.title || fallbacks.title || '',
    meetingDate:analysis.meetingDate || fallbacks.meetingDate || '',
    language:analysis.language || fallbacks.language || 'Unknown'
  };
  if(validateGeneratedMeetingAnalysis(candidate).valid) return candidate;
  const generated = await invokeMeetingTranscript({transcript:candidate.transcript});
  return {
    ...generated,
    title:generated.title || candidate.title,
    meetingDate:generated.meetingDate || candidate.meetingDate,
    language:generated.language || candidate.language,
    attendees:Array.isArray(generated.attendees) && generated.attendees.length
      ? generated.attendees
      : (Array.isArray(candidate.attendees) ? candidate.attendees : [])
  };
}

async function processMeetingUpload(){
  const file = $('#meeting-upload-file')?.files?.[0] || null;
  const transcript = $('#meeting-upload-text')?.value.trim() || '';
  if(file && transcript){
    toast('Choose a file or paste text, not both');
    return;
  }
  if(!file && !transcript){
    toast('Choose a transcript file or paste the transcript');
    return;
  }
  if(file && file.size > 10_000_000){
    toast('Transcript files must be 10 MB or smaller');
    return;
  }
  const ok = $('#modal-root [data-mc="ok"]');
  const cancel = $('#modal-root [data-mc="cancel"]');
  const status = $('#meeting-upload-status');
  if(ok){ ok.disabled = true; ok.textContent = 'Preparing notes…'; }
  if(cancel) cancel.disabled = true;
  if(status) status.textContent = 'Extracting the transcript and preparing the meeting draft. This may take a moment.';
  try{
    const extension=String(file?.name || '').toLowerCase().split('.').pop() || '';
    const browserReadableTextFile=file && ['txt','text','md','vtt','srt'].includes(extension);
    const submittedTranscript=browserReadableTextFile ? String(await file.text()).trim() : transcript;
    const filePayload=browserReadableTextFile ? null : await meetingFilePayload(file);
    const analysis = await invokeMeetingTranscript({transcript:submittedTranscript,file:filePayload});
    closeModal();
    openMeetingConfirmationModal(analysis,file?.name || '');
  }catch(error){
    if(ok){ ok.disabled = false; ok.textContent = 'Upload & generate notes'; }
    if(cancel) cancel.disabled = false;
    if(status) status.textContent = error?.message || 'Could not prepare meeting notes.';
  }
}

function meetingMemberChecks(selectedIds=[]){
  const selected = new Set(selectedIds);
  return (DATA.team || []).filter(memberRow=>memberRow.id !== currentMemberId()).map(memberRow=>`
    <label class="meeting-access-member">
      <input type="checkbox" data-meeting-access-member="${esc(memberRow.id)}" ${selected.has(memberRow.id)?'checked':''}>
      <span class="owner-av" style="background:${esc(memberRow.color || '#0073ea')}">${esc(memberRow.initials || 'TM')}</span>
      <span><strong>${esc(memberRow.name || 'Team member')}</strong><small>${esc(memberRow.email || '')}</small></span>
    </label>`).join('');
}

function meetingBoardChecks(selectedIds=[]){
  const selected = new Set(selectedIds);
  return (DATA.boards || []).map(boardRow=>`
    <label class="check-line">
      <input type="checkbox" data-meeting-board="${esc(boardRow.id)}" ${selected.has(boardRow.id)?'checked':''}>
      ${boardIconHtml(boardRow)} ${esc(boardRow.name)}
    </label>`).join('') || '<div class="integration-note">No boards are available yet.</div>';
}

function openMeetingConfirmationModal(analysis,sourceName='',sourceContext={}){
  const suggestedTitle = analysis.title || sourceName.replace(/\.[^.]+$/,'') || 'Untitled meeting';
  const body = `
    <div class="meeting-confirm-grid">
      <div class="field meeting-confirm-title"><label>Meeting title</label><input id="meeting-confirm-title" value="${esc(suggestedTitle)}"></div>
      <div class="field"><label>Meeting date</label><input id="meeting-confirm-date" type="date" value="${esc(analysis.meetingDate || '')}"></div>
      <div class="field"><label>Detected language</label><input value="${esc(analysis.language || 'Unknown')}" disabled></div>
      <div class="field meeting-confirm-title"><label>Attendees</label><input id="meeting-confirm-attendees" value="${esc((analysis.attendees || []).join(', '))}" placeholder="Names separated by commas"></div>
      <div class="field meeting-confirm-title"><label>Visibility</label>
        <select id="meeting-confirm-visibility">
          <option value="workspace">Everyone in this workspace</option>
          <option value="restricted">Restricted to selected members</option>
        </select>
      </div>
      <div class="meeting-confirm-title meeting-access-panel" id="meeting-confirm-access" hidden>
        <strong>Selected members</strong>
        <span>The uploader and workspace administrators can always access restricted meetings.</span>
        <div class="meeting-access-list">${meetingMemberChecks()}</div>
      </div>
      <div class="meeting-confirm-title meeting-related-boards">
        <strong>Related boards</strong>
        <span>Optional. This helps search and task creation.</span>
        <div class="meeting-board-list">${meetingBoardChecks()}</div>
      </div>
      <div class="meeting-confirm-title">${renderMeetingAnalysisPreview(analysis)}</div>
    </div>`;
  modal('Confirm meeting details','The AI notes are ready. Review the meeting details before saving.',body,()=>void confirmMeetingDraft(analysis,sourceContext),'Save meeting');
  $('#modal-root .modal')?.classList.add('modal-wide');
  const visibility = $('#meeting-confirm-visibility');
  if(visibility) visibility.onchange=()=>{
    const panel = $('#meeting-confirm-access');
    if(panel) panel.hidden = visibility.value !== 'restricted';
  };
}

async function confirmMeetingDraft(analysis,sourceContext={}){
  if(!SESSION?.user){ toast("Meeting analysis is a local preview. Cloud meeting saving is disabled in this sample."); return; }
  const ok = $('#modal-root [data-mc="ok"]');
  if(ok){ ok.disabled = true; ok.textContent = 'Saving…'; }
  const draft = createMeetingDraft({
    id:sourceContext.meetingId || (sourceContext.captureId ? `meeting_${sourceContext.captureId}` : ''),
    workspaceId:WORKSPACE_ID,
    uploaderId:currentMemberId(),
    sourceType:sourceContext.sourceType || 'upload',
    externalSourceId:sourceContext.externalSourceId || '',
    captureId:sourceContext.captureId || '',
    transcript:analysis.transcript || '',
    analysis:{
      ...analysis,
      title:$('#meeting-confirm-title')?.value.trim() || analysis.title,
      meetingDate:$('#meeting-confirm-date')?.value || '',
      attendees:String($('#meeting-confirm-attendees')?.value || '').split(',').map(value=>value.trim()).filter(Boolean)
    }
  });
  draft.uploaderUserId = SESSION.user.id;
  draft.visibility = $('#meeting-confirm-visibility')?.value === 'restricted' ? 'restricted' : 'workspace';
  draft.allowedMemberIds = [...document.querySelectorAll('[data-meeting-access-member]:checked')].map(input=>input.dataset.meetingAccessMember);
  draft.relatedBoardIds = [...document.querySelectorAll('[data-meeting-board]:checked')].map(input=>input.dataset.meetingBoard);
  try{
    let captureCompletionWarning='';
    MEETING_DETAIL = await getMeetingRepository().save(draft);
    if(sourceContext.captureId){
      try{
        await completeSavedMeetingCapture(sourceContext.captureId,MEETING_DETAIL.id);
        MEETING_PENDING_CAPTURES = MEETING_PENDING_CAPTURES.filter(item=>item.id !== sourceContext.captureId);
      }catch(captureError){
        console.warn('Meeting saved, but capture completion could not be recorded',captureError);
        captureCompletionWarning='Meeting notes were saved, but TaskTrack could not finish deleting the temporary capture copy. The draft remains visible so cleanup can be retried.';
      }
    }
    notifyMeetingShared(MEETING_DETAIL,[]);
    const savedMeetingId = MEETING_DETAIL.id;
    MEETING_SELECTED_ID = savedMeetingId;
    MEETING_TAB = 'overview';
    closeModal();
    openMeeting(savedMeetingId);
    await refreshMeetingList({rerender:false});
    if(VIEW === 'meetings' && MEETING_SELECTED_ID === savedMeetingId) renderContent();
    toast(captureCompletionWarning || 'Meeting notes created');
  }catch(error){
    if(ok){ ok.disabled = false; ok.textContent = 'Save meeting'; }
    toast(error?.message || 'Could not save the meeting');
  }
}

function meetingNotificationHash(value){
  let hash = 0;
  for(const char of String(value || '')) hash = ((hash << 5) - hash + char.charCodeAt(0)) | 0;
  return Math.abs(hash).toString(36);
}

function notifyMeetingShared(meeting,previousAllowedIds=[]){
  if(meeting.visibility !== 'restricted') return;
  const previous = new Set(previousAllowedIds || []);
  (meeting.allowedMemberIds || []).filter(id=>!previous.has(id) && id !== currentMemberId()).forEach(recipientId=>{
    createNotification({
      recipientId,
      type:'meeting_shared',
      title:`Meeting shared: ${meeting.title}`,
      body:'You can view and edit this restricted meeting in TaskTrack.',
      data:{meetingId:meeting.id},
      dedupeKey:`meeting-shared:${meeting.id}:${recipientId}`
    });
  });
  saveLocal();
}

function notifyMeetingMentions(text,sectionKey){
  if(!MEETING_DETAIL) return;
  mentionIdsFromText(text).filter(id=>id !== EVERYONE_MENTION_ID && id !== currentMemberId()).forEach(recipientId=>{
    createNotification({
      recipientId,
      type:'meeting_mention',
      title:`Mentioned in ${MEETING_DETAIL.title}`,
      body:`You were mentioned in the meeting ${sectionKey === 'summary' ? 'summary' : 'notes'}.`,
      data:{meetingId:MEETING_DETAIL.id},
      dedupeKey:`meeting-mention:${MEETING_DETAIL.id}:${sectionKey}:${recipientId}:${meetingNotificationHash(text)}`
    });
  });
  saveLocal();
}

async function saveCurrentMeeting({mentionText='',sectionKey=''}={}){
  if(!MEETING_DETAIL) return false;
  try{
    MEETING_DETAIL = await getMeetingRepository().save(MEETING_DETAIL);
    const index = MEETING_ROWS.findIndex(item=>item.id === MEETING_DETAIL.id);
    const row = {
      ...MEETING_DETAIL,
      transcript:undefined,
      sections:undefined,
      actionItems:undefined,
      links:undefined
    };
    if(index >= 0) MEETING_ROWS[index] = row;
    else MEETING_ROWS.unshift(row);
    if(mentionText) notifyMeetingMentions(mentionText,sectionKey);
    MEETING_ERROR = '';
    return true;
  }catch(error){
    MEETING_ERROR = error?.message || 'Could not save meeting changes.';
    toast(MEETING_ERROR);
    return false;
  }
}

function openMeetingAccessModal(){
  if(!MEETING_DETAIL) return;
  const previousAllowed = [...(MEETING_DETAIL.allowedMemberIds || [])];
  const previousVisibility = MEETING_DETAIL.visibility;
  const previousBoards = [...(MEETING_DETAIL.relatedBoardIds || [])];
  const body = `
    <div class="meeting-access-settings">
      <div class="field"><label>Visibility</label>
        <select id="meeting-access-visibility">
          <option value="workspace" ${MEETING_DETAIL.visibility!=='restricted'?'selected':''}>Everyone in this workspace</option>
          <option value="restricted" ${MEETING_DETAIL.visibility==='restricted'?'selected':''}>Restricted to selected members</option>
        </select>
      </div>
      <div id="meeting-access-members" ${MEETING_DETAIL.visibility==='restricted'?'':'hidden'}>
        <div class="meeting-privacy-note">The uploader and workspace administrators can always access restricted meetings.</div>
        <div class="meeting-access-list">${meetingMemberChecks(MEETING_DETAIL.allowedMemberIds)}</div>
      </div>
      <div class="meeting-related-boards">
        <strong>Related boards</strong>
        <div class="meeting-board-list">${meetingBoardChecks(MEETING_DETAIL.relatedBoardIds)}</div>
      </div>
    </div>`;
  modal('Meeting access','Everyone who can view this meeting can also edit it.',body,async ()=>{
    const visibility = $('#meeting-access-visibility')?.value === 'restricted' ? 'restricted' : 'workspace';
    const allowed = [...document.querySelectorAll('[data-meeting-access-member]:checked')].map(input=>input.dataset.meetingAccessMember);
    if(
      visibility === 'restricted'
      && currentMemberId() !== MEETING_DETAIL.uploaderId
      && !isCurrentAdmin()
      && !allowed.includes(currentMemberId())
    ){
      allowed.push(currentMemberId());
    }
    const boards = [...document.querySelectorAll('[data-meeting-board]:checked')].map(input=>input.dataset.meetingBoard);
    MEETING_DETAIL.visibility = visibility;
    MEETING_DETAIL.allowedMemberIds = visibility === 'restricted' ? allowed : [];
    MEETING_DETAIL.relatedBoardIds = boards;
    const saved = await saveCurrentMeeting();
    if(!saved){
      MEETING_DETAIL.visibility = previousVisibility;
      MEETING_DETAIL.allowedMemberIds = previousAllowed;
      MEETING_DETAIL.relatedBoardIds = previousBoards;
      return;
    }
    notifyMeetingShared(MEETING_DETAIL,previousAllowed);
    closeModal();
    render();
  },'Save access');
  $('#modal-root .modal')?.classList.add('modal-wide');
  const select = $('#meeting-access-visibility');
  if(select) select.onchange=()=>{
    const members = $('#meeting-access-members');
    if(members) members.hidden = select.value !== 'restricted';
  };
}

async function copyMeetingNotes(){
  if(!MEETING_DETAIL) return;
  await writeClipboardText(formatMeetingNotes(MEETING_DETAIL));
  toast('Formatted meeting notes copied');
}

function downloadMeetingPdf(){
  if(!MEETING_DETAIL) return;
  const notes = formatMeetingNotes(MEETING_DETAIL);
  const popup = window.open('','_blank');
  if(!popup){
    toast('Allow pop-ups to create the PDF');
    return;
  }
  popup.opener = null;
  const [documentTitle,...noteLines] = notes.split('\n');
  const body = esc(noteLines.join('\n')).replace(/\n/g,'<br>');
  popup.document.write(`<!doctype html><html><head><meta charset="utf-8"><title>${esc(MEETING_DETAIL.title)}</title>
    <style>body{font-family:Arial,"Noto Sans KR",sans-serif;color:#1f2233;max-width:820px;margin:48px auto;line-height:1.55;padding:0 28px}h1{font-size:25px;margin-bottom:24px}.notes{font-size:14px}@media print{body{margin:24px auto}.print-help{display:none}}</style>
    </head><body><h1>${esc(documentTitle || MEETING_DETAIL.title)}</h1><div class="notes">${body}</div><p class="print-help">Choose “Save as PDF” in the print window.</p><script>window.onload=()=>setTimeout(()=>window.print(),150);<\/script></body></html>`);
  popup.document.close();
}

function archiveCurrentMeeting(archived=true){
  if(!MEETING_DETAIL) return;
  const label = archived ? 'Archive' : 'Restore';
  const action = async ()=>{
    try{
      await getMeetingRepository().archive(MEETING_DETAIL.id,archived);
      MEETING_DETAIL.archivedAt = archived ? new Date().toISOString() : '';
      MEETING_SELECTED_ID = '';
      MEETING_DETAIL = null;
      await refreshMeetingList({rerender:false});
      syncMeetingRoute();
      render();
      toast(archived ? 'Meeting archived' : 'Meeting restored');
    }catch(error){
      toast(error?.message || `Could not ${label.toLowerCase()} the meeting`);
    }
  };
  if(archived) confirmModal('Archive this meeting?','It will move to the archived Meetings view and can be restored later.',()=>{ closeModal(); void action(); });
  else void action();
}

function permanentlyDeleteCurrentMeeting(){
  if(!MEETING_DETAIL) return;
  const id = MEETING_DETAIL.id;
  const captureId=MEETING_DETAIL.captureId || '';
  confirmModal('Permanently delete this meeting?','The transcript, notes, and proposed actions will be removed. TaskTrack tasks already created from it will remain.',async ()=>{
    try{
      if(captureId){
        await invokeMeetingCaptureSession({action:'purge',captureId,meetingId:id});
      }
      await getMeetingRepository().remove(id);
      closeModal();
      MEETING_SELECTED_ID = '';
      MEETING_DETAIL = null;
      const url = new URL(window.location.href);
      url.hash = VIEW_ROUTES.meetings;
      window.history.replaceState({},'',url.toString());
      await refreshMeetingList({rerender:false});
      render();
      toast('Meeting permanently deleted');
    }catch(error){
      toast(error?.message || 'Could not delete the meeting');
    }
  });
}

function regenerationPreviewList(title,value){
  const rows = Array.isArray(value) ? value : [value].filter(Boolean);
  return `<div class="meeting-regenerate-preview"><strong>${esc(title)}</strong>${rows.length ? `<ul>${rows.map(row=>`<li>${esc(row)}</li>`).join('')}</ul>` : '<span>No items found</span>'}</div>`;
}

function openMeetingRegenerateModal(){
  if(!MEETING_DETAIL?.transcript.trim()){
    toast('Add transcript text before regenerating notes');
    return;
  }
  const body = `<div class="meeting-regenerate-form">
    <div class="field"><label>Generated notes language</label>
      <select id="meeting-regenerate-language">
        <option value="original">Transcript’s original language</option>
        <option value="en">English</option>
        <option value="ko">Korean</option>
      </select>
    </div>
    <div id="meeting-regenerate-status" class="meeting-upload-status"></div>
    <p class="meeting-privacy-note">TaskTrack will create a preview. Your current notes will not change until you choose which sections to accept.</p>
  </div>`;
  modal('Regenerate meeting notes','Use the current edited transcript as the source.',body,()=>void prepareMeetingRegeneration(),'Create preview');
}

async function prepareMeetingRegeneration(){
  const ok = $('#modal-root [data-mc="ok"]');
  const status = $('#meeting-regenerate-status');
  if(ok){ ok.disabled = true; ok.textContent = 'Generating…'; }
  if(status) status.textContent = 'Preparing a new draft from the current transcript.';
  try{
    const analysis = await invokeMeetingTranscript({
      transcript:MEETING_DETAIL.transcript,
      mode:'regenerate',
      outputLanguage:$('#meeting-regenerate-language')?.value || 'original'
    });
    closeModal();
    openMeetingRegenerationPreview(createRegenerationPreview(analysis));
  }catch(error){
    if(ok){ ok.disabled = false; ok.textContent = 'Create preview'; }
    if(status) status.textContent = error?.message || 'Could not regenerate the notes.';
  }
}

function openMeetingRegenerationPreview(preview){
  const body = `<div class="meeting-regeneration-review">
    <label class="check-line"><input type="checkbox" data-meeting-accept-section="summary" checked> Replace meeting summary</label>
    ${regenerationPreviewList('New summary',preview.sections.summary)}
    <label class="check-line"><input type="checkbox" data-meeting-accept-section="keyDiscussionPoints" checked> Replace key discussion points</label>
    ${regenerationPreviewList('New key points',preview.sections.keyDiscussionPoints)}
    <label class="check-line"><input type="checkbox" data-meeting-accept-section="decisions" checked> Replace decisions</label>
    ${regenerationPreviewList('New decisions',preview.sections.decisions)}
    <label class="check-line"><input type="checkbox" data-meeting-accept-section="actionItems"> Replace proposed action items</label>
    ${regenerationPreviewList('New action items',preview.actionItems.map(item=>item.text))}
    <label class="check-line"><input type="checkbox" data-meeting-accept-section="openQuestions" checked> Replace open questions</label>
    ${regenerationPreviewList('New open questions',preview.sections.openQuestions)}
  </div>`;
  modal('Review regenerated notes','Choose exactly which sections TaskTrack should replace.',body,async ()=>{
    const accepted = [...document.querySelectorAll('[data-meeting-accept-section]:checked')].map(input=>input.dataset.meetingAcceptSection);
    const previous = MEETING_DETAIL;
    MEETING_DETAIL = applyRegenerationPreview(MEETING_DETAIL,preview,accepted);
    const saved = await saveCurrentMeeting();
    if(!saved){
      MEETING_DETAIL = previous;
      return;
    }
    closeModal();
    render();
    toast('Selected meeting sections updated');
  },'Apply selected sections');
  $('#modal-root .modal')?.classList.add('modal-wide');
}

function addMeetingActionItem(){
  if(!MEETING_DETAIL) return;
  const action = {
    id:uid('meeting_action'),
    text:'New action item',
    suggestedOwnerName:'',
    suggestedDueDate:'',
    confirmedOwnerId:'',
    confirmedDueDate:'',
    linkedTaskId:'',
    linkedBoardId:'',
    linkedGroupId:'',
    status:'proposed'
  };
  if(!Array.isArray(MEETING_DETAIL.actionItems)) MEETING_DETAIL.actionItems = [];
  MEETING_DETAIL.actionItems.push(action);
  MEETING_SELECTED_ACTIONS.add(action.id);
  void saveCurrentMeeting();
  renderContent();
}

function meetingTaskBoardOptions(selectedId=''){
  return (DATA.boards || []).map(boardRow=>
    `<option value="${esc(boardRow.id)}" ${boardRow.id===selectedId?'selected':''}>${esc(boardRow.name)}</option>`
  ).join('');
}

function meetingTaskGroupOptions(boardId,selectedId=''){
  const boardRow = board(boardId);
  return (boardRow?.groups || []).map(groupRow=>
    `<option value="${esc(groupRow.id)}" ${groupRow.id===selectedId?'selected':''}>${esc(groupRow.name)}</option>`
  ).join('');
}

function meetingTaskOwnerOptions(selectedId=''){
  return `<option value="">Unassigned</option>${(DATA.team || []).map(memberRow=>
    `<option value="${esc(memberRow.id)}" ${memberRow.id===selectedId?'selected':''}>${esc(memberRow.name)}</option>`
  ).join('')}`;
}

function openMeetingTaskReviewModal(){
  if(!MEETING_DETAIL) return;
  const selected = (MEETING_DETAIL.actionItems || []).filter(action=>
    action.status === 'proposed' && MEETING_SELECTED_ACTIONS.has(action.id)
  );
  if(!selected.length){
    toast('Select at least one proposed action item');
    return;
  }
  if(!(DATA.boards || []).length){
    toast('Create a board before converting action items into tasks');
    return;
  }
  const preferredBoardId = (MEETING_DETAIL.relatedBoardIds || []).find(id=>board(id))
    || (CURRENT_BOARD && board(CURRENT_BOARD) ? CURRENT_BOARD : DATA.boards[0].id);
  const body = `<div class="meeting-task-review">
    <p class="meeting-privacy-note">Nothing is created until you confirm. TaskTrack may suggest a matching member, but you control the final assignee.</p>
    ${selected.map(action=>{
      const suggestedMember = suggestMeetingActionOwner(action.suggestedOwnerName,DATA.team || []);
      const selectedOwnerId = action.confirmedOwnerId || suggestedMember?.id || '';
      const selectedBoardId = action.linkedBoardId && board(action.linkedBoardId)
        ? action.linkedBoardId
        : preferredBoardId;
      const selectedGroupId = action.linkedGroupId
        || board(selectedBoardId)?.groups?.[0]?.id
        || '';
      return `<div class="meeting-task-review-row" data-meeting-task-review="${esc(action.id)}">
        <div class="field meeting-task-review-name"><label>Task name</label><input data-review-field="name" value="${esc(action.text)}"></div>
        <div class="field"><label>Task owner</label><select data-review-field="owner">${meetingTaskOwnerOptions(selectedOwnerId)}</select></div>
        <div class="field"><label>Deadline</label><input data-review-field="due" type="date" value="${esc(action.confirmedDueDate || action.suggestedDueDate || '')}"></div>
        <div class="field"><label>Board</label><select data-review-field="board">${meetingTaskBoardOptions(selectedBoardId)}</select></div>
        <div class="field"><label>Group</label><select data-review-field="group">${meetingTaskGroupOptions(selectedBoardId,selectedGroupId)}</select></div>
        ${action.suggestedOwnerName && !suggestedMember
          ? `<div class="meeting-task-owner-note">"${esc(action.suggestedOwnerName)}" did not uniquely match a TaskTrack member, so the task is unassigned and the name will be kept in its notes.</div>`
          : suggestedMember
            ? `<div class="meeting-task-owner-note">Suggested match: ${esc(suggestedMember.name)}. Confirm or change it before creating the task.</div>`
            : ''}
      </div>`;
    }).join('')}
  </div>`;
  modal('Review proposed tasks','Confirm the task details before adding them to TaskTrack.',body,()=>void createTasksFromMeetingReview(),'Create selected tasks');
  $('#modal-root .modal')?.classList.add('modal-wide');
  document.querySelectorAll('[data-meeting-task-review]').forEach(row=>{
    const boardSelect = row.querySelector('[data-review-field="board"]');
    const groupSelect = row.querySelector('[data-review-field="group"]');
    if(boardSelect && groupSelect) boardSelect.onchange=()=>{
      groupSelect.innerHTML = meetingTaskGroupOptions(boardSelect.value);
    };
  });
}

async function refreshCanonicalMeetingTasks(results=[]){
  const ids=[...new Set(results.map(result=>result.taskId).filter(Boolean))];
  if(!ids.length) throw new Error('TaskTrack did not return the created task links.');
  const [{data:taskRows,error:taskError},{data:valueRows,error:valueError}]=await Promise.all([
    supabase.from('app_tasks').select('*').eq('workspace_id',WORKSPACE_ID).in('id',ids),
    supabase.from('app_task_values').select('*').eq('workspace_id',WORKSPACE_ID).in('task_id',ids),
  ]);
  if(taskError) throw taskError;
  if(valueError) throw valueError;
  if((taskRows || []).length!==ids.length) throw new Error('One or more canonical tasks could not be reloaded.');
  for(const row of taskRows || []){
    if(!patchTaskRow(row,'UPDATE')) throw new Error('A canonical task points to a board or group that is not loaded.');
    const loc=findTaskInData(DATA,row.id);
    if(loc?.t) loc.t.custom={};
  }
  for(const row of valueRows || []) patchTaskValueRow(row,'UPDATE');
  return new Map(ids.map(id=>[id,findTaskInData(DATA,id)]));
}

async function createTasksFromMeetingReview(){
  if(!MEETING_DETAIL) return;
  const rows = [...document.querySelectorAll('[data-meeting-task-review]')];
  const pending = rows.map(row=>{
    const action = (MEETING_DETAIL.actionItems || []).find(item=>item.id === row.dataset.meetingTaskReview);
    const boardId = row.querySelector('[data-review-field="board"]')?.value || '';
    const groupId = row.querySelector('[data-review-field="group"]')?.value || '';
    const boardRow = board(boardId);
    const groupRow = boardRow?.groups?.find(group=>group.id === groupId);
    return {
      action,
      boardRow,
      groupRow,
      name:row.querySelector('[data-review-field="name"]')?.value.trim() || '',
      ownerId:row.querySelector('[data-review-field="owner"]')?.value || '',
      due:row.querySelector('[data-review-field="due"]')?.value || ''
    };
  });
  if(pending.some(item=>!item.action || !item.boardRow || !item.groupRow || !item.name)){
    toast('Each task needs a name, board, and group');
    return;
  }
  const ok = $('#modal-root [data-mc="ok"]');
  if(ok){ ok.disabled = true; ok.textContent = 'Creating tasks...'; }
  pending.forEach(item=>{
    item.action.text = item.name;
    item.action.confirmedOwnerId = item.ownerId;
    item.action.confirmedDueDate = item.due;
    const taskDraft = buildTaskDraftFromMeetingAction(item.action,{meetingTitle:MEETING_DETAIL.title});
    item.taskDraft=taskDraft;
    item.task={
      id:uid('t'),
      name:taskDraft.name,
      owner:'',
      assignedBy:currentMemberId(),
      status:'not_started',
      priority:'none',
      due:taskDraft.due,
      tlStart:'',
      tlEnd:'',
      budget:0,
      notes:[taskDraft.notes,`Meeting notes: ${meetingUrl(MEETING_DETAIL.id)}`].filter(Boolean).join('\n'),
      custom:{},
      subtasks:[]
    };
    setTaskOwnerIds(item.task,taskDraft.ownerIds);
  });
  const rpcActions = pending.map(item=>({
    actionId:item.action.id,
    taskId:item.task.id,
    boardId:item.boardRow.id,
    groupId:item.groupRow.id,
    name:item.task.name,
    ownerId:item.ownerId,
    due:item.task.due,
    notes:item.task.notes,
  }));
  const {data:taskResults,error:taskError} = await supabase.rpc('confirm_meeting_actions',{
    p_workspace_id:WORKSPACE_ID,
    p_meeting_id:MEETING_DETAIL.id,
    p_actions:rpcActions,
  });
  if(taskError){
    if(ok){ ok.disabled=false; ok.textContent='Create selected tasks'; }
    toast(taskError.message || 'TaskTrack could not create these tasks safely');
    return;
  }
  const meetingId=MEETING_DETAIL.id;
  const meetingTitle=MEETING_DETAIL.title;
  try{
    const canonicalResults=(taskResults || []).map(canonicalMeetingTaskResult).filter(Boolean);
    const resultByAction=new Map(canonicalResults.map(result=>[result.actionId,result]));
    if(canonicalResults.length!==pending.length || pending.some(item=>!resultByAction.has(item.action.id))){
      throw new Error('TaskTrack received an incomplete task confirmation response.');
    }
    const taskLocations=await refreshCanonicalMeetingTasks(canonicalResults);
    let createdCount=0;
    for(const item of pending){
      const result=resultByAction.get(item.action.id);
      const loc=taskLocations.get(result.taskId);
      if(!loc?.t || loc.b.id!==result.boardId || loc.g.id!==result.groupId){
        throw new Error('A confirmed task did not match its canonical board and group.');
      }
      if(!result.shouldEmitCreationSideEffects) continue;
      createdCount+=1;
      recordUpdate({
        boardId:result.boardId,
        groupId:result.groupId,
        taskId:result.taskId,
        body:`Created from meeting "${meetingTitle}".`,
        kind:'audit',
        visible:false
      });
      createTaskAssignmentNotifications({
        boardId:result.boardId,
        groupId:result.groupId,
        taskId:result.taskId,
        task:loc.t,
        recipients:taskOwnerIds(loc.t)
      });
    }
    if(createdCount) save();
    MEETING_SELECTED_ACTIONS.clear();
    await loadMeetingDetail(meetingId,{rerender:false});
    closeModal();
    render();
    const restoredCount=pending.length-createdCount;
    toast(createdCount && restoredCount
      ? `${createdCount} task${createdCount===1?'':'s'} created; ${restoredCount} existing link${restoredCount===1?'':'s'} restored`
      : createdCount
        ? `${createdCount} task${createdCount===1?'':'s'} created`
        : 'These tasks were already created; TaskTrack restored their original links');
  }catch(error){
    await fetchWorkspaceState().catch(()=>{});
    await loadMeetingDetail(meetingId,{rerender:false});
    closeModal();
    render();
    toast(error?.message || 'Tasks were saved, but TaskTrack had to refresh their links.');
  }
}

/* ============================================================
   VIEW: Integrations
   ============================================================ */
const INTEGRATION_PROVIDERS = [
  { id:'google_drive', name:'Google Drive', tag:'Files', text:'Attach and find docs, sheets, briefs, and assets.', accountHint:'Connect your main marketing Drive or the Google account that can access it.', scopeLabel:'Read Drive files for search/import by default; write actions require explicit backend write-scope configuration' },
  { id:'google_calendar', name:'Google Calendar', tag:'Dates', text:'Sync task timelines and launch milestones.', accountHint:'Connect the calendar account that should receive task timelines.', scopeLabel:'View and edit calendar events' },
  { id:'google_meet', name:'Google Meet', tag:'Meetings', text:'Import completed Meet transcripts into TaskTrack for review and task conversion.', accountHint:'Connect the Google account that organizes your team meetings.', scopeLabel:'Read Meet metadata and completed transcript entries; no Drive recording access by default' },
  { id:'gmail', name:'Gmail', tag:'Email', text:'Link relevant threads to tasks and follow-ups.', accountHint:'Connect either your personal Gmail or a shared marketing inbox.', scopeLabel:'Read Gmail and create/send confirmed messages' }
];

function viewIntegrations(){
  const tiles = INTEGRATION_PROVIDERS.map(integrationTile).join('');
  return `
    <div class="integrations-grid">${tiles}</div>
    <div class="integration-panel">
      <div class="integration-panel-head">
        <div>
          <h3>Connection privacy</h3>
          <p>Each tile uses its own OAuth login, so Drive, Calendar, Meet, and Gmail can all use different external accounts.</p>
        </div>
        <button class="btn ghost sm" data-action="refresh-integrations">Refresh status</button>
      </div>
      <div class="integration-note">Connected accounts are private to the signed-in app user. You can switch or disconnect each service independently without affecting the other connections.</div>
    </div>`;
}

function integrationTile(p){
  const connection = integrationConnection(p.id);
  const connected = Boolean(connection);
  const connectedAt = connection?.connected_at ? new Date(connection.connected_at).toLocaleDateString() : '';
  return `<div class="integration-tile">
    <div class="integration-tile-top">
      <div class="integration-icon">${p.name.slice(0,1)}</div>
      <span class="integration-tag">${esc(p.tag)}</span>
    </div>
    <h3>${esc(p.name)}</h3>
    <p>${esc(p.text)}</p>
    <div class="integration-hint">${esc(p.accountHint)}</div>
    <div class="integration-status ${connected?'connected':''}">
      ${connected ? `Connected as ${esc(connection.account_label||connection.external_account_id||p.name)}` : 'Not connected'}
    </div>
    ${connected ? `<div class="integration-meta">${esc(p.scopeLabel)}${connectedAt?` · Connected ${esc(connectedAt)}`:''}</div>` : `<div class="integration-meta">${esc(p.scopeLabel)}</div>`}
    <div class="integration-actions">
      <button class="btn ghost sm" data-oauth-provider="${p.id}">${connected?'Switch account':'Connect account'}</button>
      ${connected ? `<button class="btn sm danger" data-disconnect-provider="${p.id}">Disconnect</button>` : ''}
    </div>
  </div>`;
}

function integrationConnection(provider){
  return INTEGRATION_CONNECTIONS.find(x=>x.provider === provider);
}

function integrationLabel(provider){
  return INTEGRATION_PROVIDERS.find(p=>p.id===provider)?.name || provider.replace(/_/g,' ');
}

function integrationCallbackUrl(kind){
  const base = supabaseConfig.url || 'https://PROJECT_REF.supabase.co';
  const fn = kind === 'notion' ? 'notion-oauth-callback' : 'google-oauth-callback';
  return `${base.replace(/\/+$/,'')}/functions/v1/${fn}`;
}

/* ============================================================
   VIEW: Automations
   ============================================================ */
function viewAutomations(){
  const rows = DATA.automations.map(a=>{
    const b = board(a.boardId);
    return `<div class="automation-row">
      <div>
        <div class="automation-title">${esc(a.name)}</div>
        <div class="automation-meta">${esc(automationTriggerLabel(a))} -> ${esc(automationActionLabel(a))}</div>
      </div>
      <div style="color:var(--ink2);font-weight:600">${esc(b?.name||'Missing board')}</div>
      <div><span class="auto-badge">${a.active===false?'Paused':'Active'}</span></div>
      <div style="display:flex;gap:4px;justify-content:flex-end">
        <button class="icon-btn" data-toggle-automation="${a.id}" title="${a.active===false?'Activate':'Pause'}">${a.active===false?'▶':'Ⅱ'}</button>
        <button class="icon-btn" data-del-automation="${a.id}" title="Delete">&times;</button>
      </div>
    </div>`;
  }).join('');
  return `
    <div class="toolbar">
      <button class="btn" data-action="add-automation"><span>${I.plus}</span> New automation</button>
      <button class="btn ghost" data-action="run-automations" ${DATA.automations.length?'':'disabled title="Create an automation before running rules manually"'}><span>${I.automation}</span> Run automations now</button>
    </div>
    <div class="automation-layout">
      <div class="automation-list">
        ${rows||'<div class="empty-state"><div class="big">⚡</div>No automations yet. Create a rule to move, assign, or update tasks automatically.</div>'}
      </div>
      <div class="auto-panel">
        <h3>Available triggers</h3>
        <p>Status changes to Done, owner gets assigned, or a due date is past.</p>
        <h3>Available actions</h3>
        <p>Move to a group, set status, set priority, or notify task owners.</p>
      </div>
    </div>`;
}
function automationTriggerLabel(a){
  if(a.trigger==='status_done') return 'When status changes to Done';
  if(a.trigger==='owner_assigned') return 'When an owner is assigned';
  if(a.trigger==='due_past') return 'When due date is past';
  return 'When rule matches';
}
function automationActionLabel(a){
  const b=board(a.boardId);
  if(a.action==='move_group') return `Move to ${b?.groups.find(g=>g.id===a.targetGroupId)?.name||'selected group'}`;
  if(a.action==='set_status') return `Set status to ${b ? statusOptionFor(b,a.targetStatus).label : 'selected status'}`;
  if(a.action==='set_priority') return `Set priority to ${b ? priorityOptionFor(b,a.targetPriority).label : 'selected priority'}`;
  if(a.action==='notify') return 'Notify task owners';
  return 'Run action';
}

/* ============================================================
   VIEW: Docs Hub
   ============================================================ */
const DOC_TEMPLATES = [
  {name:'Meeting Notes', color:'#0073ea', body:'## Attendees\n\n## Agenda\n\n## Notes\n\n## Action Items\n'},
  {name:'Project Brief', color:'#e2445c', body:'## Goal\n\n## Audience\n\n## Channels\n\n## Timeline\n\n## Budget\n'},
  {name:'Task Plan', color:'#fdab3d', body:'## Objective\n\n## Key Messages\n\n## Deliverables\n\n## KPIs\n'},
  {name:'Weekly Status', color:'#00c875', body:'## Wins\n\n## In Progress\n\n## Blockers\n\n## Next Week\n'},
  {name:'Retrospective', color:'#a25ddc', body:'## What went well\n\n## What didn\'t\n\n## What we\'ll change\n'}
];
function viewDocsLegacy(){
  const tmpls = DOC_TEMPLATES.map((t,i)=>`
    <div class="tmpl" data-tmpl="${i}">
      <div class="top" style="background:${t.color}">${I.docs}</div>
      <div class="bot"><div class="t">${esc(t.name)}</div><div class="by">TaskTrack template</div></div>
    </div>`).join('');
  const myDocs = DATA.docs.map(d=>`
    <div class="doc-list-item" data-open-doc="${d.id}">
      <span class="asset-ic">${I.docs}</span>
      <div style="flex:1"><div style="font-weight:600">${esc(d.title)}</div>
        <div style="font-size:12px;color:var(--ink3)">${d.template?esc(d.template)+' · ':''}${new Date(d.updated).toLocaleDateString()}</div></div>
      <button class="icon-btn" data-del-doc="${d.id}" title="Delete">✕</button>
    </div>`).join('');
  return `
    <div class="toolbar" style="justify-content:flex-end"><button class="btn" data-action="new-doc"><span>${I.plus}</span> Create new doc</button></div>
    <div class="docs-banner">
      <div style="font-size:12px;color:var(--ink3);font-weight:600">NEW TO DOCS?</div>
      <h2>Start from a template</h2>
      <p>Pick a structure below to skip the blank page.</p>
    </div>
    <div style="font-weight:700;margin-bottom:12px">Start with a template</div>
    <div class="tmpl-grid" style="margin-bottom:30px">${tmpls}</div>
    <div style="font-weight:700;margin-bottom:12px">Your docs</div>
    ${myDocs||'<div class="mw-empty">No docs yet — create one above.</div>'}`;
}

function viewDocs(){
  const tmpls = DOC_TEMPLATES.map((t,i)=>`
    <div class="tmpl" data-tmpl="${i}">
      <div class="top" style="background:${t.color}">${I.docs}</div>
      <div class="bot"><div class="t">${esc(t.name)}</div><div class="by">TaskTrack template</div></div>
    </div>`).join('');
  const docs = DATA.docs || [];
  const assets = DATA.assets || [];
  const tabs = [
    {id:'all', label:'All', count:docs.length + assets.length},
    {id:'docs', label:'Documents', count:docs.length},
    {id:'files', label:'Files / Media', count:assets.filter(a=>a.kind!=='link').length},
    {id:'links', label:'Links', count:assets.filter(a=>a.kind==='link').length},
    {id:'recent', label:'Recently used', count:docs.length + assets.length}
  ];
  const filteredAssets = assets.filter(a=>{
    if(DOCS_FILTER==='files') return a.kind !== 'link';
    if(DOCS_FILTER==='links') return a.kind === 'link';
    if(DOCS_FILTER==='docs') return false;
    return true;
  }).slice().sort((a,b)=>(Number(b.updated)||0)-(Number(a.updated)||0));
  const filteredDocs = DOCS_FILTER==='files' || DOCS_FILTER==='links' ? [] : docs.slice().sort((a,b)=>(Number(b.updated)||0)-(Number(a.updated)||0));
  const recentItems = [...docs.map(d=>({type:'doc',updated:d.updated,item:d})), ...assets.map(a=>({type:'asset',updated:a.updated,item:a}))]
    .sort((a,b)=>(Number(b.updated)||0)-(Number(a.updated)||0));
  const docRows = filteredDocs.map(d=>{
    const links = linkedTasksForDoc(d.id);
    return `<div class="doc-hub-row" data-open-doc="${d.id}">
      <span class="hub-type doc">DOC</span>
      <div class="hub-main"><strong>${esc(d.title)}</strong><span>${esc(d.template || 'TaskTrack document')} - ${new Date(d.updated).toLocaleDateString()}</span></div>
      <div class="hub-linked">${links.length ? `${links.length} linked task${links.length===1?'':'s'}` : 'No task links'}</div>
      <button class="btn ghost sm hub-link-btn" data-link-doc-task="${d.id}">Link</button>
      <button class="icon-btn" data-del-doc="${d.id}" title="Delete">${I.trash}</button>
    </div>`;
  }).join('');
  const assetRows = filteredAssets.map(a=>{
    const links = linkedTasksForAsset(a.id);
    return `<div class="doc-hub-row" data-open-asset="${a.id}">
      <span class="hub-type">${esc(assetIcon(a))}</span>
      <div class="hub-main"><strong>${esc(a.title)}</strong><span>${esc([a.source || 'upload', assetSizeLabel(a.size), a.description].filter(Boolean).join(' - ')) || 'File asset'}</span></div>
      <div class="hub-linked">${links.length ? `${links.length} linked task${links.length===1?'':'s'}` : 'No task links'}</div>
      <button class="btn ghost sm hub-link-btn" data-link-asset-task="${a.id}">Link</button>
      <button class="icon-btn" data-del-asset="${a.id}" title="Delete">${I.trash}</button>
    </div>`;
  }).join('');
  const recentRows = DOCS_FILTER==='recent' ? recentItems.slice(0,20).map(x=>{
    if(x.type==='doc') return `<div class="doc-hub-row" data-open-doc="${x.item.id}">
      <span class="hub-type doc">DOC</span><div class="hub-main"><strong>${esc(x.item.title)}</strong><span>${new Date(x.item.updated).toLocaleDateString()}</span></div><div class="hub-linked">Document</div>
    </div>`;
    return `<div class="doc-hub-row" data-open-asset="${x.item.id}">
      <span class="hub-type">${esc(assetIcon(x.item))}</span><div class="hub-main"><strong>${esc(x.item.title)}</strong><span>${new Date(x.item.updated).toLocaleDateString()}</span></div><div class="hub-linked">Asset</div>
    </div>`;
  }).join('') : '';
  const rows = DOCS_FILTER==='recent' ? recentRows : [docRows, assetRows].filter(Boolean).join('');
  return `
    <div class="docs-hub-head">
      <div>
        <h2>Docs Hub</h2>
        <p>Shared documents, uploaded files, and task attachments for this workspace.</p>
      </div>
      <div class="docs-actions">
        <button class="btn ghost" data-action="add-link-asset">${I.plus}<span>Add link</span></button>
        <button class="btn ghost" data-action="upload-assets">${I.plus}<span>Upload file</span></button>
        <button class="btn" data-action="new-doc"><span>${I.plus}</span> Create doc</button>
        <input id="docs-file-upload" type="file" multiple hidden>
      </div>
    </div>
    <div class="docs-tabs">
      ${tabs.map(t=>`<button class="${DOCS_FILTER===t.id?'active':''}" data-docs-filter="${t.id}">${esc(t.label)} <span>${t.count}</span></button>`).join('')}
    </div>
    ${DOCS_FILTER==='docs' || DOCS_FILTER==='all' ? `<div class="docs-template-panel">
      <div><strong>Start from a template</strong><span>Use a lightweight TaskTrack document structure.</span></div>
      <div class="tmpl-grid compact">${tmpls}</div>
    </div>` : ''}
    <div class="doc-hub-list">
      ${rows || '<div class="mw-empty">No items here yet. Upload a file or create a document above.</div>'}
    </div>`;
}

/* ============================================================
   VIEW: Settings
   ============================================================ */
function viewSettings(){
  const pet = assistantPetSettings();
  const admin = isCurrentAdmin();
  const adminDisabled = admin ? '' : 'disabled';
  const prefs = notificationPrefs();
  const actionPrefs = assistantActionPrefs();
  const team = DATA.team.map(m=>`
    <div class="team-row">
      <div class="owner-av" style="background:${m.color}">${esc(m.initials)}</div>
      <div style="flex:1"><div style="font-weight:600">${esc(m.name)} ${m.id===currentMemberId()?'<span style="color:var(--ink3);font-weight:500">(you)</span>':''}</div>${m.email?`<div style="font-size:12px;color:var(--ink3)">${esc(m.email)}</div>`:''}</div>
      <span class="role-badge ${memberRole(m.id)}">${memberRole(m.id)==='admin'?'Admin':'Member'}</span>
      ${admin?`<select class="role-select" data-member-role="${m.id}" ${DATA.team.filter(x=>memberRole(x.id)==='admin').length===1&&memberRole(m.id)==='admin'?'disabled':''}>
        <option value="admin" ${memberRole(m.id)==='admin'?'selected':''}>Admin</option>
        <option value="member" ${memberRole(m.id)==='member'?'selected':''}>Member</option>
      </select>`:''}
      ${m.id===currentMemberId()?'':`<button class="btn ghost sm" data-make-me="${m.id}" title="Link this teammate to your signed-in account on this device">This is me</button>`}
      ${admin&&DATA.team.length>1?`<button class="icon-btn" data-del-member="${m.id}" title="Remove">✕</button>`:''}
    </div>`).join('');
  const currencyOptions = ['$','€','¥'].map(sym=>`<option value="${esc(sym)}" ${DATA.settings.currency===sym?'selected':''}>${esc(sym)}</option>`).join('');
  const languageOptions = Object.entries(LANGUAGE_OPTIONS).map(([key,label])=>`<option value="${key}" ${appLanguage()===key?'selected':''}>${esc(label)}</option>`).join('');
  const workspaceOptions = WORKSPACE_CHOICES.map(w=>`<option value="${esc(w.id)}" ${w.id===WORKSPACE_ID?'selected':''}>${esc(w.name || 'TaskTrack')} (${workspaceRoleLabel(w)})</option>`).join('');
  const boardNotificationRows = DATA.boards.map(b=>{
    const pref = boardNotificationPrefs(b.id);
    return `<label class="board-notify-row">
      <span>${boardIconHtml(b)} ${esc(b.name)}</span>
      <select data-board-notify-mode="${esc(b.id)}">
        <option value="custom" ${pref.mode==='custom'?'selected':''}>Custom</option>
        <option value="receive_all" ${pref.mode==='receive_all'?'selected':''}>Receive all relevant</option>
        <option value="mute" ${pref.mode==='mute'?'selected':''}>Mute board</option>
      </select>
    </label>`;
  }).join('');
  return `
    <div class="settings-sec">
      <h3>Workspace</h3>
      <div class="hint">Basic info shown across the app. This workspace is shared with signed-in members by email.</div>
      <div class="field"><label>Workspace / Company name</label><input id="set-account" value="${esc(DATA.settings.account||'')}" ${adminDisabled}></div>
      ${WORKSPACE_CHOICES.length>1?`<div class="field"><label>Active workspace</label><select id="set-active-workspace">${workspaceOptions}</select></div>`:''}
      <div class="field"><label>Budget currency</label><select id="set-currency" style="max-width:160px" ${adminDisabled}>${currencyOptions}</select></div>
      <div class="field"><label>Language</label><select id="set-language" style="max-width:220px" ${adminDisabled}>${languageOptions}</select></div>
      ${admin ? '' : '<div class="integration-note" style="margin-bottom:14px">Only admins can change workspace name, currency, language, assistant appearance, or assistant action permissions.</div>'}
      ${WORKSPACE_ID?`<div class="integration-note" style="margin-bottom:14px">Shared workspace sync is on for invited members.<br>Workspace ID: ${esc(WORKSPACE_ID)}</div>`:''}
      <div style="display:flex;gap:10px;flex-wrap:wrap">
        ${admin ? '<button class="btn sm" data-action="save-workspace">Save changes</button>' : ''}
        ${isSupabaseConfigured&&SESSION&&admin?`<button class="btn ghost sm" data-action="create-workspace">Create separate workspace</button>`:''}
        <button class="btn ghost sm" data-action="restart-tour">Restart guide</button>
      </div>
    </div>

    <div class="settings-sec">
      <h3>TaskTrack Intelligence pet</h3>
      <div class="hint">Choose how the assistant appears in the bottom corner. Clicking the pet opens the natural-language chatbox.</div>
      <div class="pet-settings-layout">
        <div class="pet-settings-preview">
          <div class="assistant-pet pet-${pet.style} pet-${pet.size} motion-${pet.motion}" style="--pet-color:${esc(pet.color)}">
            ${assistantPetFace(pet)}
            <span class="assistant-pet-label">${esc(pet.name)}</span>
          </div>
        </div>
        <div class="pet-settings-fields">
          <label class="check-line"><input type="checkbox" id="set-assistant-enabled" ${pet.enabled?'checked':''} ${adminDisabled}> Show assistant pet</label>
          <div class="field"><label>Pet name</label><input id="set-assistant-name" value="${esc(pet.name)}" maxlength="32" ${adminDisabled}></div>
          <div class="settings-two">
            <div class="field"><label>Pet style</label><select id="set-assistant-style" ${adminDisabled}>
              ${ASSISTANT_PET_STYLES.map(x=>`<option value="${x.id}" ${x.id===pet.style?'selected':''}>${esc(x.label)}</option>`).join('')}
            </select></div>
            <div class="field"><label>Accent color</label><input type="color" id="set-assistant-color" value="${esc(pet.color)}" ${adminDisabled}></div>
          </div>
          <div class="settings-two">
            <div class="field"><label>Size</label><select id="set-assistant-size" ${adminDisabled}>
              ${['small','medium','large'].map(x=>`<option value="${x}" ${x===pet.size?'selected':''}>${esc(x[0].toUpperCase()+x.slice(1))}</option>`).join('')}
            </select></div>
            <div class="field"><label>Corner</label><select id="set-assistant-position" ${adminDisabled}>
              <option value="bottom-right" ${pet.position==='bottom-right'?'selected':''}>Bottom right</option>
              <option value="bottom-left" ${pet.position==='bottom-left'?'selected':''}>Bottom left</option>
            </select></div>
          </div>
          <div class="field"><label>Idle motion</label><select id="set-assistant-motion" ${adminDisabled}>
            <option value="float" ${pet.motion==='float'?'selected':''}>Float</option>
            <option value="pulse" ${pet.motion==='pulse'?'selected':''}>Pulse</option>
            <option value="still" ${pet.motion==='still'?'selected':''}>Still</option>
          </select></div>
        </div>
      </div>
      <div style="display:flex;gap:10px;flex-wrap:wrap;margin-top:14px">
        ${admin ? '<button class="btn sm" data-action="save-workspace">Save pet settings</button>' : ''}
        <button class="btn ghost sm" data-action="assistant-open-settings">${I.assistant} Open assistant</button>
      </div>
    </div>

    <div class="settings-sec">
      <h3>Assistant action permissions</h3>
      <div class="hint">These settings control what the assistant can offer after a user confirms. External actions still run through the signed-in user's connected account.</div>
      <div class="notify-grid">
        <label class="check-line ${admin?'':'muted'}"><input type="checkbox" id="assist-act-gmail-drafts" ${actionPrefs.gmailDrafts?'checked':''} ${admin?'':'disabled'}> Create Gmail drafts</label>
        <label class="check-line ${admin?'':'muted'}"><input type="checkbox" id="assist-act-gmail-send" ${actionPrefs.gmailSend?'checked':''} ${admin?'':'disabled'}> Send Gmail emails</label>
        <label class="check-line ${admin?'':'muted'}"><input type="checkbox" id="assist-act-calendar-events" ${actionPrefs.calendarEvents?'checked':''} ${admin?'':'disabled'}> Create Google Calendar events</label>
        <label class="check-line ${admin?'':'muted'}"><input type="checkbox" id="assist-act-drive-search" ${actionPrefs.driveSearch?'checked':''} ${admin?'':'disabled'}> Search Google Drive</label>
        <label class="check-line ${admin?'':'muted'}"><input type="checkbox" id="assist-act-drive-docs" ${actionPrefs.driveDocs?'checked':''} ${admin?'':'disabled'}> Create Google Drive docs</label>
        <label class="check-line ${admin?'':'muted'}"><input type="checkbox" id="assist-act-drive-import" ${actionPrefs.driveImport?'checked':''} ${admin?'':'disabled'}> Import Drive files to Docs Hub</label>
        <label class="check-line ${admin?'':'muted'}"><input type="checkbox" id="assist-act-drive-manage" ${actionPrefs.driveManage?'checked':''} ${admin?'':'disabled'}> Manage Drive files</label>
        <label class="check-line ${admin?'':'muted'}"><input type="checkbox" id="assist-act-drive-edit-docs" ${actionPrefs.driveEditDocs?'checked':''} ${admin?'':'disabled'}> Edit Google Docs text</label>
        <label class="check-line ${admin?'':'muted'}"><input type="checkbox" id="assist-act-notion-import" ${actionPrefs.notionImport?'checked':''} ${admin?'':'disabled'}> Create Notion import previews</label>
      </div>
      ${admin ? `<button class="btn sm" data-action="save-workspace">Save assistant permissions</button>` : `<div class="integration-note">Only admins can change assistant action permissions.</div>`}
    </div>

    <div class="settings-sec">
      <h3>Notification preferences</h3>
      <div class="hint">Choose when TaskTrack should create notifications and email reminders.</div>
      <div class="notify-grid">
        <label class="check-line"><input type="checkbox" id="notify-mentions" ${prefs.mentions?'checked':''}> Direct @mentions</label>
        <label class="check-line"><input type="checkbox" id="notify-everyone" ${prefs.everyone?'checked':''}> @Everyone updates</label>
        <label class="check-line"><input type="checkbox" id="notify-assigned" ${prefs.assigned?'checked':''}> Tasks assigned to me</label>
        <label class="check-line"><input type="checkbox" id="notify-subscribed" ${prefs.subscribed?'checked':''}> Tasks I follow</label>
        <label class="check-line"><input type="checkbox" id="notify-task-updates" ${prefs.taskUpdates?'checked':''}> Task update notifications</label>
        <label class="check-line"><input type="checkbox" id="notify-due-reminders" ${prefs.dueReminders?'checked':''}> Due date reminders</label>
        <label class="check-line"><input type="checkbox" id="notify-authored" ${prefs.authored?'checked':''}> Updates I posted</label>
        <label class="check-line ${desktopNotificationsSupported()?'':'muted'}"><input type="checkbox" id="notify-desktop" ${prefs.desktop!==false?'checked':''} ${desktopNotificationsSupported()?'':'disabled'}> Desktop notifications</label>
        <label class="check-line ${webPushConfigured()?'':'muted'}"><input type="checkbox" id="notify-push" ${prefs.push!==false?'checked':''} ${webPushConfigured()?'':'disabled'}> Background browser push</label>
        <label class="check-line"><input type="checkbox" id="notify-email" ${(prefs.emailDueReminders !== false && prefs.email !== false)?'checked':''}> Email 3-day and 1-day due reminders</label>
        <label class="check-line"><input type="checkbox" id="notify-daily-digest" ${prefs.dailyDigest !== false && prefs.email !== false?'checked':''}> Weekday task digest email</label>
      </div>
      <div class="field">
        <label>Board notification modes</label>
        <div class="board-notify-list">${boardNotificationRows || '<div class="integration-note">Create a board before setting board notification modes.</div>'}</div>
      </div>
      <button class="btn sm" data-action="save-workspace">Save notification preferences</button>
    </div>

    <div class="settings-sec">
      <h3>Team members</h3>
      <div class="hint">Add the people who run tasks. Use the same email they use to create their app account so they can access this shared workspace.</div>
      ${team}
      ${admin ? `<div style="display:grid;grid-template-columns:1fr 1.2fr auto;gap:8px;margin-top:14px">
        <input id="new-member-name" placeholder="Add team member name…" style="flex:1;padding:9px 12px;border:1px solid var(--line2);border-radius:6px">
        <input id="new-member-email" type="email" placeholder="Email used for login" style="padding:9px 12px;border:1px solid var(--line2);border-radius:6px">
        <button class="btn sm" data-action="add-member"><span>${I.plus}</span> Add</button>
      </div>` : `<div class="integration-note">Only admins can add, remove, or change roles for team members.</div>`}
    </div>

    <div class="settings-sec">
      <h3>Your data</h3>
      <div class="hint">${Store.persistent
        ? 'Your data is saved automatically to this shared workspace when signed in. Use Export for backups. Import accepts a TaskTrack JSON export and asks before replacing current data.'
        : '⚠ This preview can\'t save to disk. Download the file and open it in your browser for data to persist. Export still works.'}</div>
      <div style="display:flex;gap:10px;flex-wrap:wrap">
        <button class="btn ghost sm" data-action="export"><span>${I.download}</span> Export data (.json)</button>
        ${admin ? `<label class="btn ghost sm" style="cursor:pointer"><span>${I.download}</span> Import data
          <input type="file" id="import-file" accept="application/json" style="display:none"></label>
        <button class="btn danger sm" data-action="reset">Reset to sample data</button>` : ''}
      </div>
    </div>`;
}

/* ============================================================
   WIRING
   ============================================================ */
function wireSidebar(){
  document.querySelectorAll('[data-view]').forEach(el=>el.onclick=e=>{ e.preventDefault(); navigateToView(el.dataset.view); });
  document.querySelectorAll('[data-board]').forEach(el=>el.onclick=e=>{ e.preventDefault(); openBoard(el.dataset.board); });
  document.querySelectorAll('[data-board-menu]').forEach(el=>el.onclick=e=>{ e.stopPropagation(); openBoardMenu(el, el.dataset.boardMenu); });
  document.querySelectorAll('[data-action="add-board"]').forEach(el=>el.onclick=addBoard);
}
function wireTopbar(){
  const ex=$('[data-action="export"]'); if(ex) ex.onclick=exportData;
  const notifications=$('[data-action="notifications"]'); if(notifications) notifications.onclick=openNotificationsModal;
  const ac=$('[data-action="acct"]'); if(ac) ac.onclick=toggleAcct;
  const gs=$('#global-search');
  if(gs){
    gs.oninput=()=>{ GLOBAL_SEARCH=gs.value; updateGlobalSearchResults(); };
    gs.onfocus=updateGlobalSearchResults;
    gs.onkeydown=e=>{
      if(e.key==='Escape'){
        GLOBAL_SEARCH='';
        gs.value='';
        updateGlobalSearchResults();
        gs.blur();
      }
    };
  }
  wireGlobalSearchResults();
  document.querySelectorAll('[data-change-board-icon]').forEach(el=>el.onclick=()=>changeBoardIcon(el.dataset.changeBoardIcon));
  const title=$('[data-rename-board]');
  if(title){
    const original = title.textContent;
    title.onkeydown=e=>{
      if(e.key==='Enter'){ e.preventDefault(); title.blur(); }
      if(e.key==='Escape'){ e.preventDefault(); title.textContent=original; title.blur(); }
    };
    title.onblur=()=>renameBoard(title.dataset.renameBoard, title.textContent);
  }
}

function wireContent(){
  // global: close dropdowns on outside click handled separately
  document.querySelectorAll('[data-go-view]').forEach(el=>el.onclick=()=>navigateToView(el.dataset.goView));
  document.querySelectorAll('[data-action="dismiss-onboarding"]').forEach(el=>el.onclick=()=>{
    DATA.settings.onboardingDismissed = true;
    save(); renderContent(); toast('Checklist hidden');
  });
  document.querySelectorAll('[data-action="assistant-open"]').forEach(el=>el.onclick=()=>{ ASSISTANT_OPEN=true; render(); });
  // HOME
  document.querySelectorAll('[data-open-board]').forEach(el=>el.onclick=()=>openBoard(el.dataset.openBoard));
  document.querySelectorAll('[data-action="add-board"]').forEach(el=>el.onclick=addBoard);

  // UPDATES
  document.querySelectorAll('[data-update-filter]').forEach(el=>el.onclick=()=>{ UPDATE_FILTER=el.dataset.updateFilter; renderContent(); });
  document.querySelectorAll('[data-feed-board]').forEach(el=>el.onclick=()=>{
    UPDATE_BOARD_FILTER=el.dataset.feedBoard;
    if(DATA.boards.some(b=>b.id===UPDATE_BOARD_FILTER)){
      FEED_COMPOSER_BOARD = UPDATE_BOARD_FILTER;
      FEED_COMPOSER_TASK = '';
    }
    renderContent();
  });
  document.querySelectorAll('[data-action="toggle-unread-feed"]').forEach(el=>el.onclick=()=>{ UPDATE_SHOW_UNREAD=!UPDATE_SHOW_UNREAD; renderContent(); });
  document.querySelectorAll('[data-action="mark-updates-read"]').forEach(el=>el.onclick=markUpdatesRead);
  document.querySelectorAll('[data-action="request-desktop-notifications"]').forEach(el=>el.onclick=requestDesktopNotifications);
  document.querySelectorAll('[data-action="post-update"]').forEach(el=>el.onclick=postFeedUpdate);
  const feedDrop=$('#feed-file-drop');
  const feedInput=$('#feed-file-input');
  if(feedDrop && feedInput){
    feedDrop.onclick=()=>feedInput.click();
    feedInput.onclick=e=>e.stopPropagation();
    feedDrop.onkeydown=e=>{ if(e.key==='Enter' || e.key===' '){ e.preventDefault(); feedInput.click(); } };
    feedDrop.ondragover=e=>{ if(e.dataTransfer?.types?.includes('Files')){ e.preventDefault(); feedDrop.classList.add('over'); } };
    feedDrop.ondragleave=()=>feedDrop.classList.remove('over');
    feedDrop.ondrop=e=>{ e.preventDefault(); feedDrop.classList.remove('over'); addFeedFiles(Array.from(e.dataTransfer?.files || [])); };
    feedInput.onchange=e=>{ addFeedFiles(Array.from(e.target.files || [])); feedInput.value=''; };
    wireFeedFileRemoval();
  }
  document.querySelectorAll('[data-open-feed-asset]').forEach(el=>el.onclick=()=>openAsset(el.dataset.openFeedAsset));
  const feedBody=$('#feed-body'); if(feedBody) feedBody.oninput=filterMentionSuggestions;
  const feedBoard=$('#feed-board');
  if(feedBoard) feedBoard.onchange=()=>{
    FEED_COMPOSER_BOARD = feedBoard.value || '';
    FEED_COMPOSER_TASK = '';
    const taskSelect = $('#feed-task');
    if(taskSelect) taskSelect.innerHTML = `<option value="">${FEED_COMPOSER_BOARD?'No linked task':'All boards'}</option>${feedTaskOptions(FEED_COMPOSER_BOARD)}`;
  };
  const feedTask=$('#feed-task');
  if(feedTask) feedTask.onchange=()=>{
    FEED_COMPOSER_TASK = feedTask.value || '';
    if(FEED_COMPOSER_TASK){
      FEED_COMPOSER_BOARD = FEED_COMPOSER_TASK.split('|')[0] || FEED_COMPOSER_BOARD;
      if(feedBoard) feedBoard.value = FEED_COMPOSER_BOARD;
    }
  };
  document.querySelectorAll('[data-mention-insert]').forEach(el=>el.onclick=()=>insertMention(el.dataset.mentionInsert));
  document.querySelectorAll('[data-subscribe-task]').forEach(el=>el.onclick=()=>toggleTaskSubscription(el.dataset.subscribeTask));
  document.querySelectorAll('[data-update-read]').forEach(el=>el.onclick=()=>toggleUpdateRead(el.dataset.updateRead));
  document.querySelectorAll('[data-delete-update]').forEach(el=>el.onclick=e=>{
    e.preventDefault();
    e.stopPropagation();
    deleteUpdate(el.dataset.deleteUpdate);
  });
  document.querySelectorAll('[data-open-feed-task]').forEach(el=>el.onclick=()=>{ const [bid,tid]=el.dataset.openFeedTask.split('|'); openTaskDetail(bid,'',tid,'updates'); });
  document.querySelectorAll('[data-open-shared-google]').forEach(el=>el.onclick=()=>openSharedGoogleEvent(el.dataset.openSharedGoogle));
  document.querySelectorAll('[data-unshare-feed-google]').forEach(el=>el.onclick=()=>{ void unshareWorkspaceGoogleEvent(el.dataset.unshareFeedGoogle); });

  // BOARD
  const bs=$('#board-search'); if(bs){ bs.oninput=()=>{ SEARCH=bs.value; const pos=bs.selectionStart; renderContent(); const n=$('#board-search'); if(n){n.focus();n.setSelectionRange(pos,pos);} }; }
  document.querySelectorAll('[data-board-control]').forEach(el=>el.onclick=e=>{
    e.preventDefault();
    e.stopPropagation();
    openBoardControl(el, el.dataset.boardControl);
  });
  document.querySelectorAll('[data-action="clear-board-filters"]').forEach(el=>el.onclick=clearBoardFilters);
  document.querySelectorAll('[data-action="new-task"],[data-add-task]').forEach(el=>el.onclick=()=>{
    const gid = el.dataset.addTask || board(CURRENT_BOARD).groups[0]?.id; if(gid) addTask(gid);
  });
  document.querySelectorAll('[data-add-subtask]').forEach(el=>el.onclick=e=>{ e.stopPropagation(); const [gid,tid]=el.dataset.addSubtask.split('|'); addSubtask(gid,tid); });
  document.querySelectorAll('[data-toggle-subtasks]').forEach(el=>el.onclick=e=>{ e.stopPropagation(); const [gid,tid]=el.dataset.toggleSubtasks.split('|'); toggleSubtasks(gid,tid); });
  document.querySelectorAll('[data-open-task-updates]').forEach(el=>el.onclick=e=>{ e.stopPropagation(); const [bid,gid,tid]=el.dataset.openTaskUpdates.split('|'); openTaskDetail(bid,gid,tid,'updates'); });
  document.querySelectorAll('[data-open-task-row]').forEach(el=>el.onclick=e=>{ e.stopPropagation(); const [bid,gid,tid]=el.dataset.openTaskRow.split('|'); openTaskDetail(bid,gid,tid,'updates'); });
  document.querySelectorAll('tr.task-row').forEach(row=>row.onclick=e=>{
    if(!shouldOpenTaskDetailFromRowClick(e)) return;
    openTaskDetail(CURRENT_BOARD,row.dataset.group,row.dataset.task,'updates');
  });
  document.querySelectorAll('[data-action="add-column"]').forEach(el=>el.onclick=addColumn);
  document.querySelectorAll('[data-action="calendar-board"]').forEach(el=>el.onclick=()=>navigateToView('calendar'));
  document.querySelectorAll('[data-board-layout]').forEach(el=>el.onclick=()=>{
    BOARD_LAYOUT = el.dataset.boardLayout === 'kanban' ? 'kanban' : 'table';
    renderContent();
  });
  document.querySelectorAll('[data-board-view]').forEach(el=>el.onclick=()=>{ BOARD_VIEW=el.dataset.boardView; renderContent(); });
  document.querySelectorAll('[data-action="add-group"]').forEach(el=>el.onclick=addGroup);
  document.querySelectorAll('[data-toggle-group]').forEach(el=>el.onclick=()=>toggleGroup(el.dataset.toggleGroup));
  document.querySelectorAll('[data-del-group]').forEach(el=>el.onclick=()=>delGroup(el.dataset.delGroup));
  document.querySelectorAll('[data-rename-group]').forEach(el=>el.onblur=()=>renameGroup(el.dataset.renameGroup, el.textContent));
  document.querySelectorAll('[data-del-column]').forEach(el=>el.onclick=e=>{e.stopPropagation();delColumn(el.dataset.delColumn);});
  document.querySelectorAll('[data-col-menu]').forEach(el=>el.onclick=e=>{e.stopPropagation();openColumnMenu(el, el.dataset.colMenu);});
  wireColumnDrag();
  wireColumnResize();
  wireGroupDrag();
  wireTaskDrag();
  wireKanbanDrag();
  document.querySelectorAll('.row-check').forEach(el=>{
    el.onclick=e=>e.stopPropagation();
    el.onchange=e=>{ e.stopPropagation(); toggleTaskSelection(el.dataset.check, el.checked); };
  });
  document.querySelectorAll('[data-select-group]').forEach(el=>{
    el.onclick=e=>e.stopPropagation();
    el.onchange=e=>{ e.stopPropagation(); toggleGroupSelection(el.dataset.selectGroup, el.checked); };
  });
  document.querySelectorAll('.row-check-target,.select-cell').forEach(el=>{
    el.onclick=e=>e.stopPropagation();
    el.ondblclick=e=>e.stopPropagation();
  });
  document.querySelectorAll('.selection-bar').forEach(el=>{
    el.onpointerdown=e=>e.stopPropagation();
    el.onclick=e=>e.stopPropagation();
  });
  document.querySelectorAll('[data-bulk-action]').forEach(el=>el.onclick=e=>{
    e.preventDefault();
    e.stopPropagation();
    handleBulkAction(el.dataset.bulkAction);
  });
  document.querySelectorAll('[data-action="clear-selection"]').forEach(el=>el.onclick=e=>{ e.preventDefault(); e.stopPropagation(); SELECTED_TASKS.clear(); renderContent(); });
  // editable cells
  document.querySelectorAll('.cell-input').forEach(el=>{
    el.dataset.initialValue = el.value;
    el.oninput=()=>{ el.dataset.draftValue = el.value; };
    el.onblur=()=>{
      const next = el.dataset.draftValue ?? el.value;
      if(next !== el.dataset.initialValue) editTask(el.dataset.group, el.dataset.task, el.dataset.edit, next);
    };
    el.onkeydown=e=>{
      if(e.key==='Enter'){ e.preventDefault(); el.blur(); }
      if(e.key==='Escape'){
        e.preventDefault();
        el.value = el.dataset.initialValue || '';
        el.dataset.draftValue = el.value;
        el.blur();
      }
    };
    el.addEventListener('pointerdown', e=>beginEditableCellSelectionGesture(el, e));
    el.addEventListener('mousedown', e=>e.stopPropagation());
    el.addEventListener('dragstart', e=>{ e.preventDefault(); e.stopPropagation(); suppressEditableCellRowOpen(); });
    el.addEventListener('selectstart', e=>{ e.stopPropagation(); suppressEditableCellRowOpen(); });
    el.addEventListener('focus', ()=>setTaskRowDragEditing(el, true));
    el.addEventListener('blur', ()=>setTaskRowDragEditing(el, false));
  });
  document.querySelectorAll('[data-copy-cell]').forEach(el=>el.onclick=e=>{
    e.preventDefault();
    e.stopPropagation();
    void copyCellContent(el.dataset.copyCell);
  });
  // pickers
  document.querySelectorAll('[data-status-pick]').forEach(el=>el.onclick=e=>{e.stopPropagation();pickStatus(el);});
  document.querySelectorAll('[data-priority-pick]').forEach(el=>el.onclick=e=>{e.stopPropagation();pickPriority(el);});
  document.querySelectorAll('[data-owner-pick]').forEach(el=>el.onclick=e=>{e.stopPropagation();pickOwner(el);});
  document.querySelectorAll('[data-task-color-pick]').forEach(el=>el.onclick=e=>{e.stopPropagation();pickTaskColor(el);});
  document.querySelectorAll('[data-date-pick]').forEach(el=>el.onclick=e=>{e.stopPropagation();pickDate(el);});
  document.querySelectorAll('[data-tl-pick]').forEach(el=>el.onclick=e=>{e.stopPropagation();pickTimeline(el);});
  document.querySelectorAll('[data-custom-date-pick]').forEach(el=>el.onclick=e=>{e.stopPropagation();pickCustomDate(el);});
  document.querySelectorAll('[data-custom-status-pick]').forEach(el=>el.onclick=e=>{e.stopPropagation();pickCustomStatus(el);});
  document.querySelectorAll('[data-custom-recurrence-pick]').forEach(el=>el.onclick=e=>{e.stopPropagation();pickCustomRecurrence(el);});
  document.querySelectorAll('[data-file-cell-pick]').forEach(el=>el.onclick=e=>{e.stopPropagation();pickFilesCell(el);});
  document.querySelectorAll('[data-doc-cell-pick]').forEach(el=>el.onclick=e=>{e.stopPropagation();pickDocumentCell(el);});
  document.querySelectorAll('[data-open-asset]').forEach(el=>el.onclick=e=>{e.stopPropagation();void openAsset(el.dataset.openAsset);});

  // CALENDAR
  document.querySelectorAll('[data-action="cal-prev"]').forEach(el=>el.onclick=()=>shiftCalendar(-1));
  document.querySelectorAll('[data-action="cal-next"]').forEach(el=>el.onclick=()=>shiftCalendar(1));
  document.querySelectorAll('[data-action="cal-today"]').forEach(el=>el.onclick=()=>{ CAL_MONTH=new Date().toISOString().slice(0,7); renderContent(); });
  document.querySelectorAll('[data-action="toggle-my-google-calendar"]').forEach(el=>el.onclick=()=>{
    MY_GOOGLE_CALENDAR_ON = !MY_GOOGLE_CALENDAR_ON;
    PRIVATE_GOOGLE_ERROR = '';
    renderContent();
  });
  document.querySelectorAll('[data-action="retry-my-google-calendar"]').forEach(el=>el.onclick=()=>{
    PRIVATE_GOOGLE_LOAD_KEY = '';
    PRIVATE_GOOGLE_ERROR = '';
    renderContent();
  });
  document.querySelectorAll('[data-action="connect-calendar"]').forEach(el=>el.onclick=()=>navigateToView('integrations'));
  document.querySelectorAll('[data-action="calendar-all"]').forEach(el=>el.onclick=()=>{ CURRENT_BOARD=null; CALENDAR_GROUP_FILTER='all'; navigateToView('calendar'); });
  document.querySelectorAll('[data-calendar-board]').forEach(el=>el.onclick=()=>{ CURRENT_BOARD=el.dataset.calendarBoard; CALENDAR_GROUP_FILTER='all'; navigateToView('calendar'); });
  document.querySelectorAll('[data-calendar-group]').forEach(el=>el.onclick=()=>{ CALENDAR_GROUP_FILTER=el.dataset.calendarGroup || 'all'; renderContent(); });
  document.querySelectorAll('[data-cal-task]').forEach(el=>el.onclick=e=>{ e.stopPropagation(); showCalendarTaskDetail(el.dataset.calTask, el); });
  document.querySelectorAll('[data-cal-google]').forEach(el=>el.onclick=e=>{ e.stopPropagation(); showGoogleCalendarEventDetail(el.dataset.calGoogle, el); });
  document.querySelectorAll('[data-cal-more]').forEach(el=>el.onclick=e=>{ e.stopPropagation(); showCalendarMore(el.dataset.calMore, el); });
  const closingMonth=$('#closing-month');
  if(closingMonth) closingMonth.onchange=()=>{ CLOSING_MONTH=closingMonth.value; renderContent(); };
  document.querySelectorAll('[data-closing-task]').forEach(el=>el.onclick=()=>{ const [bid,tid]=el.dataset.closingTask.split('|'); openTaskDetail(bid,'',tid,'activity'); });
  document.querySelectorAll('[data-closing-event]').forEach(el=>el.onclick=()=>{
    const event=SHARED_GOOGLE_EVENTS.find(row=>row.id===el.dataset.closingEvent);
    if(!event) return;
    CAL_MONTH=googleEventCalendarSpan(event).start.slice(0,7);
    navigateToView('calendar');
    requestAnimationFrame(()=>{ const anchor=[...document.querySelectorAll('[data-cal-google]')].find(node=>node.dataset.calGoogle===`shared:${event.id}`); if(anchor) showGoogleCalendarEventDetail(`shared:${event.id}`,anchor); });
  });
  document.querySelectorAll('[data-dash-task]').forEach(el=>el.onclick=()=>{ const [bid,tid]=el.dataset.dashTask.split('|'); openTaskDetail(bid,'',tid,'updates'); });
  document.querySelectorAll('[data-open-kanban-task]').forEach(el=>el.onclick=e=>{ e.stopPropagation(); const [bid,tid]=el.dataset.openKanbanTask.split('|'); openTaskDetail(bid,'',tid,'updates'); });
  document.querySelectorAll('[data-my-work-task]').forEach(el=>{
    const open = ()=>{ const [bid,tid]=el.dataset.myWorkTask.split('|'); openTaskDetail(bid,'',tid,'updates'); };
    el.onclick=open;
    el.onkeydown=e=>{ if(e.key==='Enter' || e.key===' '){ e.preventDefault(); open(); } };
  });
  document.querySelectorAll('[data-my-work-toggle]').forEach(el=>el.onclick=()=>{
    const key = el.dataset.myWorkToggle;
    if(MY_WORK_COLLAPSED.has(key)) MY_WORK_COLLAPSED.delete(key);
    else MY_WORK_COLLAPSED.add(key);
    renderContent();
  });
  document.querySelectorAll('[data-my-work-sort]').forEach(el=>el.onclick=()=>{
    const next = el.dataset.myWorkSort;
    if(MY_WORK_SORT_OPTIONS.some(opt=>opt.id === next)){
      MY_WORK_SORT = next;
      renderContent();
    }
  });

  // MEETINGS
  document.querySelectorAll('[data-action="meeting-upload"]').forEach(el=>el.onclick=openMeetingUploadModal);
  document.querySelectorAll('[data-action="meeting-record"]').forEach(el=>el.onclick=openMeetingRecordModal);
  document.querySelectorAll('[data-action="meeting-google-meet"]').forEach(el=>el.onclick=openGoogleMeetSetupModal);
  document.querySelectorAll('[data-action="meeting-capture-pause"]').forEach(el=>el.onclick=()=>void pauseMeetingCapture(true));
  document.querySelectorAll('[data-action="meeting-capture-resume"]').forEach(el=>el.onclick=()=>void pauseMeetingCapture(false));
  document.querySelectorAll('[data-action="meeting-capture-stop"]').forEach(el=>el.onclick=()=>void stopAndFinalizeMeetingCapture());
  document.querySelectorAll('[data-action="meeting-capture-cancel"]').forEach(el=>el.onclick=confirmCancelMeetingCapture);
  document.querySelectorAll('[data-action="meeting-capture-clear-external"]').forEach(el=>el.onclick=()=>void clearStoppedExternalMeetingCapture());
  document.querySelectorAll('[data-review-meeting-capture]').forEach(el=>el.onclick=()=>void reviewPendingMeetingCapture(el.dataset.reviewMeetingCapture));
  document.querySelectorAll('[data-discard-meeting-capture]').forEach(el=>el.onclick=()=>discardPendingMeetingCapture(el.dataset.discardMeetingCapture));
  document.querySelectorAll('[data-open-meeting]').forEach(el=>el.onclick=()=>openMeeting(el.dataset.openMeeting));
  const meetingSearch=$('#meeting-search');
  if(meetingSearch) meetingSearch.oninput=()=>{
    MEETING_SEARCH = meetingSearch.value;
    const pos = meetingSearch.selectionStart;
    renderContent();
    const next = $('#meeting-search');
    if(next){
      next.focus();
      next.setSelectionRange(pos,pos);
    }
  };
  const showArchived=$('#meeting-show-archived');
  if(showArchived) showArchived.onchange=()=>{
    MEETING_SHOW_ARCHIVED = showArchived.checked;
    const visible = MEETING_ROWS.filter(item=>Boolean(item.archivedAt) === MEETING_SHOW_ARCHIVED);
    if(!visible.some(item=>item.id === MEETING_SELECTED_ID)){
      MEETING_SELECTED_ID = visible[0]?.id || '';
      MEETING_DETAIL = null;
      if(MEETING_SELECTED_ID) void loadMeetingDetail(MEETING_SELECTED_ID);
    }
    syncMeetingRoute();
    renderContent();
  };
  document.querySelectorAll('[data-meeting-tab]').forEach(el=>el.onclick=()=>{
    MEETING_TAB = el.dataset.meetingTab;
    renderContent();
  });
  document.querySelectorAll('[data-meeting-field]').forEach(el=>el.onblur=()=>{
    if(!MEETING_DETAIL) return;
    const field = el.dataset.meetingField;
    const next = el.value.trim();
    if(field === 'title') MEETING_DETAIL.title = next || 'Untitled meeting';
    if(field === 'meetingDate') MEETING_DETAIL.meetingDate = next;
    void saveCurrentMeeting();
  });
  const meetingAttendees=$('[data-meeting-attendees]');
  if(meetingAttendees) meetingAttendees.onblur=()=>{
    if(!MEETING_DETAIL) return;
    MEETING_DETAIL.attendees = meetingAttendees.value.split(',').map(value=>value.trim()).filter(Boolean);
    void saveCurrentMeeting();
  };
  document.querySelectorAll('[data-meeting-section]').forEach(el=>el.onblur=()=>{
    if(!MEETING_DETAIL) return;
    const key = el.dataset.meetingSection;
    const value = el.value.trim();
    MEETING_DETAIL.sections[key] = key === 'summary'
      ? value
      : value.split(/\r?\n/).map(row=>row.trim()).filter(Boolean);
    void saveCurrentMeeting({mentionText:value,sectionKey:key});
  });
  const meetingTranscript=$('[data-meeting-transcript]');
  if(meetingTranscript) meetingTranscript.onblur=()=>{
    if(!MEETING_DETAIL) return;
    MEETING_DETAIL.transcript = meetingTranscript.value.trim();
    void saveCurrentMeeting();
  };
  document.querySelectorAll('[data-meeting-action-select]').forEach(el=>{
    el.checked = MEETING_SELECTED_ACTIONS.has(el.dataset.meetingActionSelect);
    el.onchange=()=>{
      if(el.checked) MEETING_SELECTED_ACTIONS.add(el.dataset.meetingActionSelect);
      else MEETING_SELECTED_ACTIONS.delete(el.dataset.meetingActionSelect);
    };
  });
  document.querySelectorAll('[data-meeting-action-field]').forEach(el=>el.onblur=()=>{
    if(!MEETING_DETAIL) return;
    const [id,field] = el.dataset.meetingActionField.split('|');
    const action = MEETING_DETAIL.actionItems.find(item=>item.id === id);
    if(!action || action.status !== 'proposed') return;
    action[field] = el.value.trim();
    void saveCurrentMeeting();
  });
  document.querySelectorAll('[data-action="meeting-access"]').forEach(el=>el.onclick=openMeetingAccessModal);
  document.querySelectorAll('[data-action="meeting-copy"]').forEach(el=>el.onclick=()=>void copyMeetingNotes());
  document.querySelectorAll('[data-action="meeting-pdf"]').forEach(el=>el.onclick=downloadMeetingPdf);
  document.querySelectorAll('[data-action="meeting-regenerate"]').forEach(el=>el.onclick=openMeetingRegenerateModal);
  document.querySelectorAll('[data-action="meeting-archive"]').forEach(el=>el.onclick=()=>archiveCurrentMeeting(true));
  document.querySelectorAll('[data-action="meeting-restore"]').forEach(el=>el.onclick=()=>archiveCurrentMeeting(false));
  document.querySelectorAll('[data-action="meeting-delete"]').forEach(el=>el.onclick=permanentlyDeleteCurrentMeeting);
  document.querySelectorAll('[data-action="meeting-add-action"]').forEach(el=>el.onclick=addMeetingActionItem);
  document.querySelectorAll('[data-action="meeting-create-tasks"]').forEach(el=>el.onclick=openMeetingTaskReviewModal);
  document.querySelectorAll('[data-open-meeting-task]').forEach(el=>el.onclick=()=>{
    const [boardId,taskId] = el.dataset.openMeetingTask.split('|');
    openTaskDetail(boardId,'',taskId,'updates');
  });

  // INTEGRATIONS
  document.querySelectorAll('[data-oauth-provider]').forEach(el=>el.onclick=()=>startIntegrationOauth(el.dataset.oauthProvider));
  document.querySelectorAll('[data-disconnect-provider]').forEach(el=>el.onclick=()=>disconnectIntegration(el.dataset.disconnectProvider));
  document.querySelectorAll('[data-action="refresh-integrations"]').forEach(el=>el.onclick=()=>refreshIntegrationConnections({ rerender:true }));

  // DOCS
  document.querySelectorAll('[data-docs-filter]').forEach(el=>el.onclick=()=>{ DOCS_FILTER=el.dataset.docsFilter || 'all'; renderContent(); });
  document.querySelectorAll('[data-action="upload-assets"]').forEach(el=>el.onclick=triggerDocsUpload);
  document.querySelectorAll('[data-action="add-link-asset"]').forEach(el=>el.onclick=createLinkAsset);
  const docsUpload=$('#docs-file-upload'); if(docsUpload) docsUpload.onchange=e=>uploadAssetFiles(Array.from(e.target.files || []));
  document.querySelectorAll('[data-tmpl]').forEach(el=>el.onclick=()=>newDocFromTemplate(+el.dataset.tmpl));
  document.querySelectorAll('[data-action="new-doc"]').forEach(el=>el.onclick=()=>newDocFromTemplate(-1));
  document.querySelectorAll('[data-open-doc]').forEach(el=>el.onclick=()=>openDoc(el.dataset.openDoc));
  document.querySelectorAll('[data-link-doc-task]').forEach(el=>el.onclick=e=>{ e.stopPropagation(); linkDocToTask(el.dataset.linkDocTask); });
  document.querySelectorAll('[data-link-asset-task]').forEach(el=>el.onclick=e=>{ e.stopPropagation(); linkAssetToTask(el.dataset.linkAssetTask); });
  document.querySelectorAll('[data-del-doc]').forEach(el=>el.onclick=e=>{e.stopPropagation();delDoc(el.dataset.delDoc);});
  document.querySelectorAll('[data-del-asset]').forEach(el=>el.onclick=e=>{e.stopPropagation();delAsset(el.dataset.delAsset);});

  // SETTINGS
  document.querySelectorAll('[data-action="save-workspace"]').forEach(el=>el.onclick=saveWorkspace);
  const activeWorkspace = $('#set-active-workspace');
  if(activeWorkspace) activeWorkspace.onchange=()=>{ void switchWorkspace(activeWorkspace.value); };
  document.querySelectorAll('[data-action="create-workspace"]').forEach(el=>el.onclick=createWorkspaceModal);
  document.querySelectorAll('[data-action="restart-tour"]').forEach(el=>el.onclick=()=>{
    DATA.settings.guidedTourDone = false;
    TOUR_SHOWN_THIS_SESSION = true;
    openGuidedTour(0);
  });
  document.querySelectorAll('[data-action="assistant-open-settings"]').forEach(el=>el.onclick=()=>{ ASSISTANT_OPEN=true; render(); });
  document.querySelectorAll('[data-member-role]').forEach(el=>el.onchange=()=>changeMemberRole(el.dataset.memberRole, el.value));
  const am=$('[data-action="add-member"]'); if(am) am.onclick=addMember;
  document.querySelectorAll('[data-make-me]').forEach(el=>el.onclick=()=>{ setCurrentMemberId(el.dataset.makeMe); render(); toast('Switched current user'); });
  document.querySelectorAll('[data-del-member]').forEach(el=>el.onclick=()=>delMember(el.dataset.delMember));
  document.querySelectorAll('[data-action="export"]').forEach(el=>el.onclick=exportData);
  document.querySelectorAll('[data-action="reset"]').forEach(el=>el.onclick=resetData);
  const imp=$('#import-file'); if(imp) imp.onchange=importData;
}

function shouldOpenTaskDetailFromRowClick(e){
  if(!e || e.defaultPrevented) return false;
  if(Date.now() - taskDragFinishedAt < 350) return false;
  if(Date.now() < editableCellSelectionGuardUntil || hasActiveEditableCellSelection()) return false;
  const interactive = e.target.closest(`
    input,button,select,textarea,[contenteditable],
    a,label,.pill,.date-cell,.timeline-bar,.select-cell,.row-check-target,
    .owner-cell,.asset-cell,.doc-cell,.custom-cell,.budget-cell,.subtask-toggle,
    .subtask-add,.asset-chip,.doc-chip,.asset-empty,.asset-add-mini
  `);
  return !interactive;
}

function suppressEditableCellRowOpen(ms=700){
  editableCellSelectionGuardUntil = Math.max(editableCellSelectionGuardUntil, Date.now() + ms);
}

function hasInputTextSelection(input){
  if(!input || typeof input.selectionStart !== 'number' || typeof input.selectionEnd !== 'number') return false;
  return input.selectionStart !== input.selectionEnd;
}

function hasActiveEditableCellSelection(){
  const active = document.activeElement;
  return active?.classList?.contains('cell-input') && hasInputTextSelection(active);
}

function beginEditableCellSelectionGesture(input, e){
  e.stopPropagation();
  setTaskRowDragEditing(input, true);
  editableCellSelectionGesture = {
    input,
    x:e.clientX ?? 0,
    y:e.clientY ?? 0,
    moved:false
  };
  document.addEventListener('pointermove', trackEditableCellSelectionGesture, true);
  document.addEventListener('pointerup', finishEditableCellSelectionGesture, true);
}

function trackEditableCellSelectionGesture(e){
  if(!editableCellSelectionGesture) return;
  const dx = Math.abs((e.clientX ?? 0) - editableCellSelectionGesture.x);
  const dy = Math.abs((e.clientY ?? 0) - editableCellSelectionGesture.y);
  if(dx > 3 || dy > 3){
    editableCellSelectionGesture.moved = true;
    suppressEditableCellRowOpen();
  }
}

function finishEditableCellSelectionGesture(e){
  const gesture = editableCellSelectionGesture;
  document.removeEventListener('pointermove', trackEditableCellSelectionGesture, true);
  document.removeEventListener('pointerup', finishEditableCellSelectionGesture, true);
  editableCellSelectionGesture = null;
  if(!gesture) return;
  const endedInsideInput = e.target === gesture.input;
  if(gesture.moved || !endedInsideInput || hasInputTextSelection(gesture.input)){
    suppressEditableCellRowOpen();
  }
}

function cellCopyValue(gid,tid,field){
  const t = findTask(gid,tid) || findTaskLocationById(tid)?.t;
  if(!t) return '';
  if(field?.startsWith('custom:')){
    return String(t.custom?.[field.slice(7)] || '');
  }
  if(field === 'notes') return String(t.notes || '');
  if(field === 'name') return String(t.name || '');
  return String(t[field] || '');
}

async function copyCellContent(key){
  const [gid, tid, field] = String(key || '').split('|');
  const text = cellCopyValue(gid,tid,field);
  if(!text.trim()){
    toast('Nothing to copy');
    return;
  }
  await writeClipboardText(text);
  toast('Copied');
}

async function writeClipboardText(text){
  if(navigator.clipboard?.writeText){
    try{
      await navigator.clipboard.writeText(text);
      return;
    }catch(e){}
  }
  const textarea = document.createElement('textarea');
  textarea.value = text;
  textarea.setAttribute('readonly','');
  textarea.style.position = 'fixed';
  textarea.style.left = '-9999px';
  document.body.appendChild(textarea);
  textarea.select();
  try{ document.execCommand('copy'); }catch(e){}
  textarea.remove();
}

/* ---------- close any open dropdown / popovers on outside click ---------- */
document.addEventListener('click', (e)=>{
  if(Date.now() < editableCellSelectionGuardUntil && e.target.closest?.('tr.task-row')){
    e.preventDefault();
    e.stopPropagation();
  }
}, true);
document.addEventListener('click', (e)=>{
  if(openDD && !openDD.contains(e.target)) closeDD();
  const gs=$('#global-search-results');
  if(gs && !e.target.closest('.global-search-box')) gs.classList.remove('open');
  const am=$('#acct-menu'); if(am && !am.contains(e.target) && !e.target.closest('[data-action="acct"]')) am.remove();
});
window.addEventListener('hashchange', ()=>{
  if(!DATA) return;
  applyHashRoute();
  render();
});
window.addEventListener('pagehide',meetingCapturePageHidden);
document.addEventListener('visibilitychange',meetingCaptureVisibilityChanged);
window.addEventListener('resize', ()=>{ if(!isEditingControl()) closeDD(); });
document.addEventListener('scroll', (e)=>{
  if(!openDD || openDD.contains(e.target) || isEditingControl()) return;
  closeDD();
}, true);
document.addEventListener('focusin', schedulePresenceTrack, true);
document.addEventListener('input', schedulePresenceTrack, true);
document.addEventListener('click', schedulePresenceTrack, true);

/* ============================================================
   ACTIONS — boards / groups / tasks
   ============================================================ */
function openBoard(boardId, { updateHash=true } = {}){
  if(!board(boardId)) return;
  VIEW = 'board';
  CURRENT_BOARD = boardId;
  if(updateHash) window.history.replaceState({}, '', boardUrl(boardId));
  render();
}

function openBoardMenu(anchor,bid){
  const b = board(bid);
  if(!b) return;
  const idx = DATA.boards.findIndex(x=>x.id===bid);
  const canMoveUp = idx > 0;
  const canMoveDown = idx >= 0 && idx < DATA.boards.length - 1;
  const html = `
    <div class="board-menu-list">
      <button class="board-menu-item" data-board-act="open">${I.board}<span>Open board</span></button>
      <button class="board-menu-item" data-board-act="new-tab">${I.move}<span>Open in new tab</span></button>
      <div class="board-menu-divider"></div>
      <button class="board-menu-item" data-board-act="rename">${I.convert}<span>Rename</span></button>
      <button class="board-menu-item" data-board-act="icon">${I.board}<span>Change emoji</span></button>
      <button class="board-menu-item" data-board-act="move-up" ${canMoveUp?'':'disabled'}>${I.move}<span>Move up</span></button>
      <button class="board-menu-item" data-board-act="move-down" ${canMoveDown?'':'disabled'}>${I.move}<span>Move down</span></button>
      <button class="board-menu-item" data-board-act="duplicate">${I.copy}<span>Duplicate</span></button>
      <button class="board-menu-item" data-board-act="template">${I.docs}<span>Save as template</span></button>
      <div class="board-menu-divider"></div>
      <button class="board-menu-item danger" data-board-act="delete">${I.trash}<span>Delete</span></button>
      <button class="board-menu-item" data-board-act="archive">${I.archive}<span>Archive</span></button>
    </div>`;
  const dd = buildDD(anchor, html);
  dd.classList.add('board-menu-popover');
  dd.querySelectorAll('[data-board-act]').forEach(btn=>btn.onclick=()=>{
    const action = btn.dataset.boardAct;
    closeDD();
    if(action==='open') return openBoard(bid);
    if(action==='new-tab') return window.open(boardUrl(bid), '_blank', 'noopener');
    if(action==='rename') return renameBoardModal(bid);
    if(action==='icon') return changeBoardIcon(bid);
    if(action==='move-up') return moveBoard(bid, -1);
    if(action==='move-down') return moveBoard(bid, 1);
    if(action==='duplicate') return duplicateBoard(bid);
    if(action==='template') return saveBoardAsTemplate(bid);
    if(action==='delete') return deleteBoard(bid);
    if(action==='archive') return archiveBoard(bid);
  });
}

function boardIconFieldHtml(value='📋'){
  return `<div class="field"><label>Board emoji</label><input id="m-board-icon" value="${esc(value)}" maxlength="4" placeholder="📋"></div>
    <div class="emoji-grid">
      ${BOARD_EMOJIS.map(e=>`<button type="button" data-emoji-choice="${esc(e)}">${esc(e)}</button>`).join('')}
    </div>`;
}
function wireEmojiPicker(root=document){
  root.querySelectorAll('[data-emoji-choice]').forEach(btn=>btn.onclick=()=>{
    const input = root.querySelector('#m-board-icon');
    if(input) input.value = btn.dataset.emojiChoice || '📋';
  });
}

function addBoard(){
  modal('New board','Give your task board a name.',
    `<div class="field"><label>Board name</label><input id="m-board-name" placeholder="e.g. Q4 Tasks" autofocus></div>
     ${boardIconFieldHtml('📋')}`,
    ()=>{
      const name=$('#m-board-name').value.trim()||'Untitled board';
      const icon=$('#m-board-icon').value.trim()||'📋';
      const b={id:uid('b'),name,icon,columns:[],columnOrder:BASE_COLUMNS.map(c=>c.id),groups:[
        {id:uid('g'),name:'To-Do',color:'#0073ea',open:true,tasks:[]},
        {id:uid('g'),name:'Completed',color:'#00c875',open:true,tasks:[]}
      ]};
      DATA.boards.push(b); save(); VIEW='board'; CURRENT_BOARD=b.id; closeModal(); render(); toast('Board created');
    });
  wireEmojiPicker($('#modal-root'));
}
function addGroup(){
  const b=board(CURRENT_BOARD); if(!b)return;
  const color=GROUP_COLORS[b.groups.length%GROUP_COLORS.length];
  b.groups.push({id:uid('g'),name:'New Group',color,open:true,tasks:[]}); save(); render();
}
function toggleGroup(gid){ const b=board(CURRENT_BOARD); const g=b.groups.find(x=>x.id===gid); g.open=g.open===false?true:false; save(); renderContent(); wireContent(); }
function delGroup(gid){
  const b=board(CURRENT_BOARD); const g=b.groups.find(x=>x.id===gid);
  confirmModal(`Delete group “${g.name}”?`, `This removes the group and its ${g.tasks.length} task(s).`,()=>{
    b.groups=b.groups.filter(x=>x.id!==gid); save(); closeModal(); render(); toast('Group deleted');
  });
}
function renameGroup(gid,name){ const b=board(CURRENT_BOARD); const g=b.groups.find(x=>x.id===gid); if(g){ g.name=name.trim()||g.name; save(); } }
function renameBoard(boardId,name){
  const b=board(boardId); if(!b)return;
  const next = String(name||'').trim() || b.name || 'Untitled board';
  if(next === b.name){ render(); return; }
  b.name = next;
  recordUpdate({boardId:b.id,body:`Renamed board to "${b.name}".`});
  save();
  render();
  toast('Board renamed');
}
function renameBoardModal(boardId){
  const b=board(boardId); if(!b)return;
  modal('Rename board','Update the board name shown in the sidebar and page header.',
    `<div class="field"><label>Board name</label><input id="m-board-rename" value="${esc(b.name)}" autofocus></div>`,
    ()=>{
      renameBoard(boardId, $('#m-board-rename').value);
      closeModal();
    }, 'Rename');
}
function changeBoardIcon(boardId){
  const b=board(boardId); if(!b)return;
  const current = String(b.icon || '📋').startsWith('<svg') ? '📋' : (b.icon || '📋');
  modal('Change board emoji','Choose an emoji that helps the team recognize this board.',
    boardIconFieldHtml(current),
    ()=>{
      b.icon = $('#m-board-icon').value.trim() || '📋';
      save();
      closeModal();
      render();
      toast('Board emoji updated');
    }, 'Save');
  wireEmojiPicker($('#modal-root'));
}
function moveBoard(boardId,dir){
  const idx=DATA.boards.findIndex(b=>b.id===boardId);
  const to=idx+dir;
  if(idx<0 || to<0 || to>=DATA.boards.length)return;
  DATA.boards.splice(to,0,DATA.boards.splice(idx,1)[0]);
  save(); render(); toast('Board moved');
}
function duplicateBoard(boardId){
  const source=board(boardId); if(!source)return;
  const colMap={};
  const copy=JSON.parse(JSON.stringify(source));
  copy.id=uid('b');
  copy.name=`Copy of ${source.name || 'Untitled board'}`;
  copy.icon=source.icon || '📋';
  copy.columns=(source.columns||[]).map(c=>{
    const id=uid('c');
    colMap[c.id]=id;
    return {...c,id};
  });
  copy.columnOrder=(source.columnOrder||BASE_COLUMNS.map(c=>c.id)).map(id=>colMap[id]||id);
  copy.columnWidths=Object.fromEntries(Object.entries(source.columnWidths||{}).map(([id,width])=>[colMap[id]||id,width]));
  copy.groups=(source.groups||[]).map(g=>({
    ...g,
    id:uid('g'),
    tasks:(g.tasks||[]).map(t=>cloneBoardTask(t,colMap))
  }));
  const idx=DATA.boards.findIndex(b=>b.id===boardId);
  DATA.boards.splice(idx+1,0,copy);
  save();
  VIEW='board';
  CURRENT_BOARD=copy.id;
  render();
  toast('Board duplicated');
}
function cloneBoardTask(t,colMap){
  const custom={};
  Object.entries(visibleTaskCustom(t)).forEach(([cid,val])=>{ custom[colMap[cid]||cid]=val; });
  return {
    ...JSON.parse(JSON.stringify(t)),
    id:uid('t'),
    custom,
    subtasks:(t.subtasks||[]).map(st=>cloneBoardTask(st,colMap))
  };
}
function saveBoardAsTemplate(boardId){
  const b=board(boardId); if(!b)return;
  const columns=boardColumns(b).map(c=>c.label||c.name).join(', ');
  const groups=(b.groups||[]).map(g=>`### ${g.name}\n${(g.tasks||[]).map(t=>`- ${t.name}`).join('\n') || '- No tasks yet'}`).join('\n\n');
  DATA.docs.unshift({
    id:uid('doc'),
    title:`${b.name} board template`,
    template:'Board template',
    content:`## Board\n\n${b.name}\n\n## Columns\n\n${columns}\n\n## Groups and tasks\n\n${groups}\n`,
    updated:Date.now()
  });
  save(); render(); toast('Template saved in Docs Hub');
}
function removeBoard(boardId){
  const idx=DATA.boards.findIndex(b=>b.id===boardId);
  if(idx<0)return null;
  const [removed]=DATA.boards.splice(idx,1);
  DATA.automations=DATA.automations.filter(a=>a.boardId!==boardId);
  if(CURRENT_BOARD===boardId){
    CURRENT_BOARD=DATA.boards[Math.min(idx,DATA.boards.length-1)]?.id || null;
    VIEW=CURRENT_BOARD ? 'board' : 'home';
  }
  return removed;
}
function deleteBoard(boardId){
  const b=board(boardId); if(!b)return;
  confirmModal(`Delete "${b.name}"?`,'This permanently removes the board, its groups, and all tasks on it.',()=>{
    removeBoard(boardId);
    save(); closeModal(); render(); toast('Board deleted');
  });
}
function archiveBoard(boardId){
  const b=board(boardId); if(!b)return;
  confirmModal(`Archive "${b.name}"?`,'This removes the board from the sidebar and keeps a copy in your archived data.',()=>{
    const removed=removeBoard(boardId);
    if(removed) DATA.archive.unshift({id:uid('arch'),type:'board',archivedAt:Date.now(),board:removed});
    save(); closeModal(); render(); toast('Board archived');
  });
}
function addTask(gid){
  const b=board(CURRENT_BOARD); const g=b.groups.find(x=>x.id===gid);
  const t = {id:uid('t'),name:'New task',owner:'',assignedBy:currentMemberId(),status:'not_started',priority:'none',due:'',tlStart:'',tlEnd:'',budget:0,notes:'',custom:{},subtasks:[]};
  g.tasks.push(t);
  recordUpdate({boardId:b.id,groupId:g.id,taskId:t.id,body:`Created task "${t.name}".`,kind:'audit',visible:false});
  save(); renderContent(); wireContent();
}
function editTask(gid,tid,field,val){
  const b=board(CURRENT_BOARD); const g=b.groups.find(x=>x.id===gid); const t=findTaskInGroup(g,tid)?.t;
  if(!t)return;
  if(field && field.startsWith('custom:')){
    const cid = field.slice(7);
    const col = findColumn(b,cid);
    if(!t.custom) t.custom = {};
    if(col?.type === 'recurrence'){
      const next = normalizeRecurrenceValue(val);
      t.custom[cid] = next === 'off' ? '' : next;
      setTaskRecurrence(t,next);
    } else {
      const next = col?.type==='number' ? (Number(String(val).replace(/[^0-9.\-]/g,''))||0) : val;
      if(col?.type === 'url'){
        const duplicate = findDuplicateAssetUrl(next);
        if(duplicate) toast(duplicateAssetMessage(duplicate, next));
      }
      t.custom[cid] = next;
    }
    save();
    if(['number','url','recurrence'].includes(col?.type)){ renderContent(); wireContent(); }
    return;
  }
  if(field==='budget') t.budget=Number(String(val).replace(/[^0-9.\-]/g,''))||0;
  else if(field==='name') t.name=String(val || '').trim() || 'New task';
  else t[field]=val;
  if(field==='status'){
    if(t.status==='done' && !t.completedAt) t.completedAt = new Date().toISOString();
    if(t.status!=='done') t.completedAt = '';
  }
  const changed = applyAutomations({type:'field_change',boardId:b.id,groupId:gid,taskId:tid,field,value:t[field]});
  save();
  if(field==='budget' || changed){ renderContent(); wireContent(); } // refresh sums and moved rows
}

function addSubtask(gid,parentId){
  const b=board(CURRENT_BOARD); const g=b.groups.find(x=>x.id===gid);
  const parent=findTaskInGroup(g,parentId)?.t;
  if(!parent)return;
  if(!Array.isArray(parent.subtasks)) parent.subtasks=[];
  setSubtasksOpen(parent, true);
  const st={id:uid('t'),name:'New subtask',owner:'',assignedBy:currentMemberId(),status:'not_started',priority:'none',due:'',tlStart:'',tlEnd:'',budget:0,notes:'',custom:{},subtasks:[]};
  setTaskOwnerIds(st, taskOwnerIds(parent));
  parent.subtasks.push(st);
  recordUpdate({
    boardId:b.id,
    groupId:g.id,
    taskId:st.id,
    mentions:taskOwnerIds(st),
    body:`Added subtask "${st.name}" under "${parent.name}".`
  });
  save(); renderContent(); wireContent();
}

function toggleSubtasks(gid,tid){
  const t=findTask(gid,tid) || findTaskLocationById(tid)?.t;
  if(!t || !childTasks(t).length)return;
  setSubtasksOpen(t, !areSubtasksOpen(t));
  save();
  renderContent();
}

function addColumn(){
  const b=board(CURRENT_BOARD); if(!b)return;
  normalizeBoard(b);
  modal('New column','Create a new field for this board.',
    `<div class="field"><label>Column name</label><input id="m-col-name" placeholder="e.g. Channel, Asset Link, CPA" autofocus></div>
     <div class="field"><label>Column type</label><select id="m-col-type">
       ${Object.entries(CUSTOM_COLUMN_TYPES).map(([k,v])=>`<option value="${k}">${v}</option>`).join('')}
     </select></div>`,
    ()=>{
      const name=$('#m-col-name').value.trim()||'New column';
      const type=$('#m-col-type').value;
      const col={id:uid('c'),name,type};
      b.columns.push(col);
      b.columnOrder.push(col.id);
      b.groups.forEach(g=>g.tasks.forEach(t=>setTaskCustomValueDeep(t, col.id, '')));
      save(); closeModal(); renderContent(); toast('Column added');
    }, 'Add column');
}
function delColumn(cid){
  const b=board(CURRENT_BOARD); const col=findColumn(b,cid); if(!b||!col)return;
  confirmModal(`Delete column "${col.name}"?`,'This removes the column values from every task on this board.',()=>{
    b.columns=b.columns.filter(c=>c.id!==cid);
    b.columnOrder=b.columnOrder.filter(id=>id!==cid);
    b.groups.forEach(g=>g.tasks.forEach(t=>deleteTaskCustomValueDeep(t, cid)));
    save(); closeModal(); renderContent(); toast('Column deleted');
  });
}

function openBoardControl(anchor,type){
  if(type==='person') return openPersonFilter(anchor);
  if(type==='hide') return openHideColumns(anchor);
  if(type==='sort') return openSortMenu(anchor);
  if(type==='calendar') return openBoardCalendarOptions(anchor);
}

function openPersonFilter(anchor){
  const b=board(CURRENT_BOARD); if(!b)return;
  normalizeBoard(b);
  const personFilter = boardPersonFilter(b);
  const html = `<div class="board-control-pop">
    <div class="control-title">Filter this board by person</div>
    <button class="person-filter-option ${!personFilter?'selected':''}" data-person-filter="">
      <span class="owner-empty">${I.plus}</span><span>All people</span>
    </button>
    ${DATA.team.map(m=>`<button class="person-filter-option ${personFilter===m.id?'selected':''}" data-person-filter="${m.id}">
      <span class="owner-av" style="background:${m.color}">${esc(m.initials)}</span><span>${esc(m.name)}</span>
    </button>`).join('')}
  </div>`;
  const dd=buildDD(anchor, html);
  dd.querySelectorAll('[data-person-filter]').forEach(btn=>btn.onclick=()=>{
    setBoardPersonFilter(b, btn.dataset.personFilter || '');
    closeDD();
    renderContent();
  });
}

function openHideColumns(anchor){
  const b=board(CURRENT_BOARD); if(!b)return;
  normalizeBoard(b);
  const cols=allBoardColumns(b);
  const selectedCount=cols.filter(c=>!b.hiddenColumns.includes(c.id)).length;
  const html = `<div class="board-control-pop wide">
    <div class="control-title">Display columns</div>
    <input class="control-search" data-column-search placeholder="Find columns to show/hide">
    <label class="column-toggle all"><input type="checkbox" data-col-toggle-all ${selectedCount===cols.length?'checked':''}> <span>All columns</span><em>${selectedCount} selected</em></label>
    <div class="column-toggle-list">
      ${cols.map(c=>`<label class="column-toggle" data-column-toggle-row>
        <input type="checkbox" data-col-toggle="${c.id}" ${b.hiddenColumns.includes(c.id)?'':'checked'} ${c.id==='name'?'disabled':''}>
        <span>${esc(c.label)}</span>
      </label>`).join('')}
    </div>
  </div>`;
  const dd=buildDD(anchor, html);
  const search=dd.querySelector('[data-column-search]');
  if(search) search.oninput=()=>{
    const q=search.value.trim().toLowerCase();
    dd.querySelectorAll('[data-column-toggle-row]').forEach(row=>{
      row.style.display = row.textContent.toLowerCase().includes(q) ? '' : 'none';
    });
  };
  const all=dd.querySelector('[data-col-toggle-all]');
  if(all) all.onchange=()=>{
    b.hiddenColumns = all.checked ? [] : cols.filter(c=>c.id!=='name').map(c=>c.id);
    save(); renderContent();
  };
  dd.querySelectorAll('[data-col-toggle]').forEach(box=>box.onchange=()=>{
    const cid=box.dataset.colToggle;
    if(cid==='name') return;
    b.hiddenColumns = box.checked
      ? b.hiddenColumns.filter(id=>id!==cid)
      : [...new Set([...b.hiddenColumns,cid])];
    save(); renderContent();
  });
}

function openSortMenu(anchor){
  const b=board(CURRENT_BOARD); if(!b)return;
  normalizeBoard(b);
  const cols=allBoardColumns(b);
  const html = `<div class="board-control-pop sort-pop">
    <div class="control-title">Sort by</div>
    <div class="sort-row">
      <select data-sort-column>
        <option value="">Choose column</option>
        ${cols.map(c=>`<option value="${c.id}" ${b.sort?.columnId===c.id?'selected':''}>${esc(c.label)}</option>`).join('')}
      </select>
      <select data-sort-direction>
        <option value="asc" ${b.sort?.direction!=='desc'?'selected':''}>Ascending</option>
        <option value="desc" ${b.sort?.direction==='desc'?'selected':''}>Descending</option>
      </select>
    </div>
    <button class="btn ghost sm control-clear" data-clear-sort>Clear sort</button>
  </div>`;
  const dd=buildDD(anchor, html);
  const apply=()=>{
    const columnId=dd.querySelector('[data-sort-column]').value;
    const direction=dd.querySelector('[data-sort-direction]').value;
    b.sort = columnId ? {columnId,direction} : {};
    save(); renderContent();
  };
  dd.querySelector('[data-sort-column]').onchange=apply;
  dd.querySelector('[data-sort-direction]').onchange=apply;
  dd.querySelector('[data-clear-sort]').onclick=()=>{ b.sort={}; save(); closeDD(); renderContent(); };
}

function openBoardCalendarOptions(anchor){
  const b=board(CURRENT_BOARD); if(!b)return;
  normalizeBoard(b);
  const html = `<div class="board-control-pop wide">
    <div class="control-title">Calendar display</div>
    <label class="column-toggle calendar-option">
      <input type="checkbox" data-calendar-timeline-toggle ${b.calendarShowTimeline?'checked':''}>
      <span>Show full timeline ranges</span>
    </label>
    <div class="control-note">When this is on, tasks from this board use their full timeline span on the calendar. When it is off, they use only due dates.</div>
  </div>`;
  const dd=buildDD(anchor, html);
  dd.querySelector('[data-calendar-timeline-toggle]').onchange=e=>{
    b.calendarShowTimeline = e.target.checked;
    save();
    renderContent();
  };
}

function openColumnMenu(anchor,cid){
  const b=board(CURRENT_BOARD); if(!b)return;
  normalizeBoard(b);
  const c=allBoardColumns(b).find(col=>col.id===cid);
  if(!c)return;
  const canDelete = cid !== 'name';
  const labelEditor = c.id==='status' || c.id==='priority' || c.type==='status';
  const html = `<div class="board-menu-list">
    <button class="board-menu-item" data-col-act="rename">${I.convert}<span>Rename</span></button>
    ${labelEditor?`<button class="board-menu-item" data-col-act="labels">${I.plus}<span>Edit labels</span></button>`:''}
    <div class="board-menu-divider"></div>
    <button class="board-menu-item danger" data-col-act="delete" ${canDelete?'':'disabled'}>${I.trash}<span>${c.custom?'Delete':'Hide'} column</span></button>
  </div>`;
  const dd=buildDD(anchor, html);
  dd.classList.add('board-menu-popover');
  dd.querySelectorAll('[data-col-act]').forEach(btn=>btn.onclick=()=>{
    const act=btn.dataset.colAct;
    closeDD();
    if(act==='rename') return renameColumnLabel(cid);
    if(act==='labels') return editBoardLabels(c.id==='priority' ? 'priority' : 'status');
    if(act==='delete') return deleteOrHideColumn(cid);
  });
}

function renameColumnLabel(cid){
  const b=board(CURRENT_BOARD); if(!b)return;
  const c=allBoardColumns(b).find(col=>col.id===cid);
  if(!c)return;
  modal('Rename column','Change how this column appears on this board.',
    `<div class="field"><label>Column name</label><input id="m-col-label" value="${esc(c.label)}" autofocus></div>`,
    ()=>{
      const next=$('#m-col-label').value.trim() || c.label;
      if(c.custom){
        const custom=findColumn(b,cid);
        if(custom) custom.name=next;
        delete b.columnLabels[cid];
      } else {
        b.columnLabels[cid]=next;
      }
      save(); closeModal(); renderContent(); toast('Column renamed');
    }, 'Rename');
}

function deleteOrHideColumn(cid){
  const b=board(CURRENT_BOARD); if(!b || cid==='name')return;
  const c=allBoardColumns(b).find(col=>col.id===cid);
  if(!c)return;
  if(c.custom) return delColumn(cid);
  b.hiddenColumns = [...new Set([...b.hiddenColumns,cid])];
  save(); renderContent(); toast('Column hidden');
}

function editBoardLabels(type){
  const b=board(CURRENT_BOARD); if(!b)return;
  const isPriority = type === 'priority';
  const base = isPriority ? PRIORITY : STATUS;
  const options = isPriority ? boardPriorityOptions(b) : boardStatusOptions(b);
  const hiddenKey = isPriority ? 'hiddenPriorityLabels' : 'hiddenStatusLabels';
  modal(isPriority?'Edit priority labels':'Edit status labels','Rename labels or add new ones for this board.',
    `<div class="label-editor" data-label-editor="${type}">
      ${Object.entries(options).map(([key,opt])=>labelEditRow(key,opt,true,Boolean(base[key]))).join('')}
      <button type="button" class="add-label-row" data-add-label-row>+ New label</button>
    </div>`,
    ()=>{
      const next={};
      const validKeys=new Set(Object.keys(base));
      const visibleKeys=new Set();
      $('#modal-root').querySelectorAll('[data-label-row]').forEach(row=>{
        const originalKey=row.dataset.labelKey || '';
        const key=originalKey || uid(isPriority?'p':'s');
        const label=row.querySelector('[data-label-name]').value.trim();
        const color=row.querySelector('[data-label-color]').value || LABEL_COLORS[0];
        if(!label)return;
        validKeys.add(key);
        visibleKeys.add(key);
        if(!base[key] || base[key].label !== label || base[key].color !== color) next[key]={label,color};
      });
      if(!visibleKeys.size){
        toast('Keep at least one label');
        return;
      }
      b[hiddenKey] = Object.keys(base).filter(key=>!visibleKeys.has(key));
      if(isPriority) b.priorityLabels=next;
      else b.statusLabels=next;
      replaceDeletedLabelValues(b,type,visibleKeys);
      save(); closeModal(); renderContent(); toast('Labels updated');
    }, 'Apply');
  wireLabelEditor($('#modal-root'), type);
}

function labelEditRow(key,opt,canDelete=true,builtIn=false){
  return `<div class="label-edit-row ${builtIn?'built-in':''}" data-label-row data-label-key="${esc(key||'')}">
    <input type="color" data-label-color value="${esc(opt.color||LABEL_COLORS[0])}">
    <input data-label-name value="${esc(opt.label||'New label')}">
    <button type="button" data-remove-label ${canDelete?'':'disabled'}>${I.trash}</button>
  </div>`;
}

function wireLabelEditor(root,type){
  const list=root.querySelector('[data-label-editor]');
  root.querySelectorAll('[data-remove-label]').forEach(btn=>btn.onclick=()=>{
    if(btn.disabled)return;
    btn.closest('[data-label-row]')?.remove();
  });
  const add=root.querySelector('[data-add-label-row]');
  if(add) add.onclick=()=>{
    const color=LABEL_COLORS[list.querySelectorAll('[data-label-row]').length % LABEL_COLORS.length];
    add.insertAdjacentHTML('beforebegin', labelEditRow('', {label:'New label',color}, true));
    const row=add.previousElementSibling;
    row.querySelector('[data-remove-label]').onclick=()=>row.remove();
    row.querySelector('[data-label-name]').focus();
  };
}

function replaceDeletedLabelValues(b,type,validKeys){
  const field = type === 'priority' ? 'priority' : 'status';
  const fallback = [...validKeys][0] || (type === 'priority' ? 'none' : 'not_started');
  b.groups.forEach(g=>g.tasks.forEach(t=>replaceDeletedLabelOnTask(t,field,validKeys,fallback)));
}
function replaceDeletedLabelOnTask(t,field,validKeys,fallback){
  if(t[field] && !validKeys.has(t[field])) t[field]=fallback;
  (t.subtasks||[]).forEach(st=>replaceDeletedLabelOnTask(st,field,validKeys,fallback));
}

function wireColumnDrag(){
  document.querySelectorAll('[data-col-drag]').forEach(th=>{
    th.ondragstart=e=>{
      if(e.target.closest('.col-resizer')){ e.preventDefault(); return; }
      draggingColumn=th.dataset.colDrag;
      th.classList.add('dragging');
      e.dataTransfer.effectAllowed='move';
      e.dataTransfer.setData('text/plain',draggingColumn);
    };
    th.ondragover=e=>{ e.preventDefault(); th.classList.add('drag-over'); };
    th.ondragleave=()=>th.classList.remove('drag-over');
    th.ondrop=e=>{
      e.preventDefault();
      th.classList.remove('drag-over');
      moveColumn(draggingColumn, th.dataset.colDrag);
    };
    th.ondragend=()=>{ th.classList.remove('dragging'); document.querySelectorAll('.drag-over').forEach(x=>x.classList.remove('drag-over')); draggingColumn=null; };
  });
}
function wireColumnResize(){
  document.querySelectorAll('[data-col-resize]').forEach(handle=>{
    handle.onmousedown=e=>{
      e.preventDefault();
      e.stopPropagation();
      const b=board(CURRENT_BOARD); if(!b)return;
      normalizeBoard(b);
      const cid=handle.dataset.colResize;
      const startX=e.clientX;
      const startWidth=columnWidth(b,cid);
      document.body.classList.add('resizing-column');
      const applyWidth=width=>{
        const next = Math.min(720, Math.max(80, Math.round(width)));
        b.columnWidths[cid] = next;
        document.querySelectorAll('[data-col-id],[data-col-cell]').forEach(el=>{
          if(el.dataset.colId===cid || el.dataset.colCell===cid){
            el.style.width = `${next}px`;
            el.style.minWidth = `${next}px`;
          }
        });
      };
      const onMove=ev=>applyWidth(startWidth + ev.clientX - startX);
      const onUp=()=>{
        document.removeEventListener('mousemove', onMove);
        document.removeEventListener('mouseup', onUp);
        document.body.classList.remove('resizing-column');
        save();
      };
      document.addEventListener('mousemove', onMove);
      document.addEventListener('mouseup', onUp);
    };
  });
}
function moveColumn(fromId,toId){
  if(!fromId || !toId || fromId===toId)return;
  const b=board(CURRENT_BOARD); if(!b)return;
  normalizeBoard(b);
  const order=b.columnOrder;
  const from=order.indexOf(fromId), to=order.indexOf(toId);
  if(from<0 || to<0)return;
  order.splice(to,0,order.splice(from,1)[0]);
  save(); renderContent(); toast('Column moved');
}

function wireGroupDrag(){
  document.querySelectorAll('[data-group-drag]').forEach(handle=>{
    const groupEl = handle.closest('.group');
    handle.ondragstart=e=>{
      draggingGroup = handle.dataset.groupDrag;
      groupEl?.classList.add('dragging');
      e.dataTransfer.effectAllowed='move';
      e.dataTransfer.setData('text/plain',draggingGroup);
    };
    handle.ondragend=()=>{
      groupEl?.classList.remove('dragging');
      document.querySelectorAll('.group-over').forEach(x=>x.classList.remove('group-over'));
      draggingGroup = null;
    };
  });
  document.querySelectorAll('.group').forEach(groupEl=>{
    groupEl.ondragover=e=>{
      if(draggingGroup && groupEl.dataset.group!==draggingGroup){
        e.preventDefault();
        groupEl.classList.add('group-over');
      }
    };
    groupEl.ondragleave=()=>groupEl.classList.remove('group-over');
    groupEl.ondrop=e=>{
      if(!draggingGroup)return;
      e.preventDefault();
      groupEl.classList.remove('group-over');
      moveGroup(draggingGroup, groupEl.dataset.group);
    };
  });
}
function moveGroup(fromId,toId){
  if(!fromId || !toId || fromId===toId)return;
  const b=board(CURRENT_BOARD); if(!b)return;
  const from=b.groups.findIndex(g=>g.id===fromId);
  const to=b.groups.findIndex(g=>g.id===toId);
  if(from<0 || to<0)return;
  b.groups.splice(to,0,b.groups.splice(from,1)[0]);
  save(); renderContent(); toast('Group moved');
}

function setTaskRowDragEditing(input, editing){
  const row = input?.closest?.('tr.task-row');
  const cell = input?.closest?.('.cell');
  if(cell){
    cell.classList.toggle('editing-cell', Boolean(editing));
    cell.classList.toggle('editing-task-name', Boolean(editing) && cell.classList.contains('task-name-cell'));
  }
  if(!row) return;
  if(editing){
    row.dataset.dragDisabled = '1';
    row.removeAttribute('draggable');
  } else {
    delete row.dataset.dragDisabled;
    row.setAttribute('draggable', 'true');
  }
}

function wireTaskDrag(){
  document.querySelectorAll('tr.task-row').forEach(row=>{
    row.ondragstart=e=>{
      if(row.dataset.dragDisabled === '1'){
        e.preventDefault(); return;
      }
      if(e.target.closest('input,button,select,textarea,[contenteditable],.cell,.pill,.date-cell,.timeline-bar,.select-cell,.row-check-target')){
        e.preventDefault(); return;
      }
      if(row.dataset.parent){
        draggingSubtask={taskId:row.dataset.task, parentId:row.dataset.parent, groupId:row.dataset.group};
        draggingTask=null;
      } else {
        draggingTask={taskId:row.dataset.task, fromGroupId:row.dataset.group};
        draggingSubtask=null;
      }
      row.classList.add('dragging');
      e.dataTransfer.effectAllowed='move';
      e.dataTransfer.setData('text/plain',draggingSubtask?.taskId || draggingTask.taskId);
    };
    row.ondragover=e=>{
      if(draggingSubtask){
        const siblingTarget = row.dataset.parent === draggingSubtask.parentId && row.dataset.task !== draggingSubtask.taskId;
        const parentTarget = !row.dataset.parent && row.dataset.task === draggingSubtask.parentId;
        if(siblingTarget || parentTarget){
          e.preventDefault();
          row.classList.add('task-over');
        }
        return;
      }
      if(draggingTask && !row.dataset.parent){
        e.preventDefault();
        row.classList.add('task-over');
      }
    };
    row.ondragleave=()=>row.classList.remove('task-over');
    row.ondrop=e=>{
      if(!draggingTask && !draggingSubtask)return;
      e.preventDefault(); e.stopPropagation();
      row.classList.remove('task-over');
      if(draggingSubtask){
        const beforeTaskId = row.dataset.parent === draggingSubtask.parentId ? row.dataset.task : null;
        moveSubtaskWithinParent(draggingSubtask.taskId, draggingSubtask.parentId, beforeTaskId, draggingSubtask.parentId);
        return;
      }
      if(row.dataset.parent) return;
      moveTask(draggingTask.taskId, draggingTask.fromGroupId, row.dataset.group, row.dataset.task);
    };
    row.ondragend=()=>{ row.classList.remove('dragging'); cleanupTaskDrag(); };
  });
  document.querySelectorAll('[data-drop-group]').forEach(tbody=>{
    tbody.ondragover=e=>{ if(draggingTask){ e.preventDefault(); } };
    tbody.ondrop=e=>{
      if(!draggingTask || e.target.closest('tr.task-row'))return;
      e.preventDefault();
      moveTask(draggingTask.taskId, draggingTask.fromGroupId, tbody.dataset.dropGroup, null);
    };
  });
}
function cleanupTaskDrag(){
  document.querySelectorAll('.task-over,.task-row.dragging').forEach(el=>el.classList.remove('task-over','dragging'));
  if(draggingTask || draggingSubtask) taskDragFinishedAt = Date.now();
  draggingTask=null;
  draggingSubtask=null;
}
function moveTask(taskId,fromGroupId,toGroupId,beforeTaskId){
  if(taskId===beforeTaskId)return false;
  const b=board(CURRENT_BOARD); if(!b)return false;
  const fromGroup=b.groups.find(g=>g.id===fromGroupId);
  const toGroup=b.groups.find(g=>g.id===toGroupId);
  if(!fromGroup || !toGroup)return false;
  const idx=fromGroup.tasks.findIndex(t=>t.id===taskId);
  if(idx<0)return false;
  const [task]=fromGroup.tasks.splice(idx,1);
  let insertAt = beforeTaskId ? toGroup.tasks.findIndex(t=>t.id===beforeTaskId) : -1;
  if(insertAt<0) insertAt=toGroup.tasks.length;
  toGroup.tasks.splice(insertAt,0,task);
  save(); cleanupTaskDrag(); renderContent(); toast('Task moved');
  return true;
}

function moveSubtaskWithinParent(taskId,parentId,beforeTaskId,beforeParentId){
  if(!taskId || !parentId || taskId===beforeTaskId || beforeParentId!==parentId)return false;
  const parent = findTaskLocationById(parentId)?.t;
  if(!parent || !Array.isArray(parent.subtasks))return false;
  const idx = parent.subtasks.findIndex(t=>t.id===taskId);
  if(idx<0)return false;
  const [task] = parent.subtasks.splice(idx,1);
  let insertAt = beforeTaskId ? parent.subtasks.findIndex(t=>t.id===beforeTaskId) : -1;
  if(insertAt<0) insertAt = parent.subtasks.length;
  parent.subtasks.splice(insertAt,0,task);
  setSubtasksOpen(parent, true);
  save(); cleanupTaskDrag(); renderContent(); toast('Subtask moved');
  return true;
}

function wireKanbanDrag(){
  document.querySelectorAll('[data-kanban-drag]').forEach(card=>{
    card.ondragstart=e=>{
      draggingKanbanTask = card.dataset.kanbanDrag;
      card.classList.add('dragging');
      e.dataTransfer.effectAllowed='move';
      e.dataTransfer.setData('text/plain', draggingKanbanTask);
    };
    card.ondragend=()=>{
      card.classList.remove('dragging');
      document.querySelectorAll('.kanban-lane.over').forEach(x=>x.classList.remove('over'));
      draggingKanbanTask = null;
    };
  });
  document.querySelectorAll('[data-kanban-status]').forEach(lane=>{
    lane.ondragover=e=>{
      if(!draggingKanbanTask) return;
      e.preventDefault();
      lane.classList.add('over');
    };
    lane.ondragleave=()=>lane.classList.remove('over');
    lane.ondrop=e=>{
      if(!draggingKanbanTask) return;
      e.preventDefault();
      lane.classList.remove('over');
      setTaskStatusFromKanban(draggingKanbanTask, lane.dataset.kanbanStatus);
    };
  });
}

function setTaskStatusFromKanban(taskId,status){
  const loc = findTaskLocationById(taskId);
  if(!loc || !status) return;
  if(loc.t.status === status){ draggingKanbanTask = null; return; }
  loc.t.status = status;
  if(status === 'done' && !loc.t.completedAt) loc.t.completedAt = new Date().toISOString();
  if(status !== 'done') loc.t.completedAt = '';
  const opt = statusOptionFor(loc.b,status);
  recordUpdate({boardId:loc.b.id,groupId:loc.g.id,taskId:loc.t.id,body:`Moved "${loc.t.name}" to ${opt.label}.`});
  applyAutomations({type:'field_change',boardId:loc.b.id,groupId:loc.g.id,taskId:loc.t.id,field:'status',value:status});
  save();
  draggingKanbanTask = null;
  renderContent();
  toast('Task status updated');
}

function toggleTaskSelection(taskId,checked){
  checked ? SELECTED_TASKS.add(taskId) : SELECTED_TASKS.delete(taskId);
  renderContent();
}
function toggleGroupSelection(gid,checked){
  const b=board(CURRENT_BOARD); const g=b?.groups.find(x=>x.id===gid); if(!g)return;
  taskIdsInGroup(g).forEach(id=>checked ? SELECTED_TASKS.add(id) : SELECTED_TASKS.delete(id));
  renderContent();
}
function handleBulkAction(action){
  const locs = normalizeBulkLocations(selectedTaskLocations().filter(x=>x.b.id===CURRENT_BOARD));
  if(!locs.length){ SELECTED_TASKS.clear(); renderContent(); return; }
  if(action==='duplicate') return duplicateSelectedTasks(locs);
  if(action==='export') return exportSelectedTasks(locs);
  if(action==='hide') return setSelectedTasksHidden(locs, true);
  if(action==='unhide') return setSelectedTasksHidden(locs, false);
  if(action==='archive') return archiveSelectedTasks(locs);
  if(action==='delete') return deleteSelectedTasks(locs);
  if(action==='convert') return convertSelectedTasks(locs);
  if(action==='move') return moveSelectedTasks(locs);
}
function duplicateSelectedTasks(locs){
  locs.forEach(loc=>{
    const {t}=loc;
    const copy=cloneTaskWithNewIds(t);
    copy.name=`Copy of ${t.name||'Untitled task'}`;
    const siblings = taskSiblingList(loc);
    const idx=siblings.findIndex(x=>x.id===t.id);
    siblings.splice(idx+1,0,copy);
  });
  SELECTED_TASKS.clear(); save(); renderContent(); toast(`${locs.length} task${locs.length===1?'':'s'} duplicated`);
}
function exportSelectedTasks(locs){
  const rows = locs.map(({b,g,t})=>({board:b.name, group:g.name, name:t.name, owner:taskOwnerNames(t), status:statusOptionFor(b,t.status).label||'', due:t.due||'', timeline_start:t.tlStart||'', timeline_end:t.tlEnd||'', priority:priorityOptionFor(b,t.priority).label||'', budget:t.budget||0, notes:t.notes||'', custom:visibleTaskCustom(t)}));
  const blob=new Blob([JSON.stringify(rows,null,2)],{type:'application/json'});
  const url=URL.createObjectURL(blob); const a=document.createElement('a');
  a.href=url; a.download=`tasktrack-selected-tasks-${new Date().toISOString().slice(0,10)}.json`; a.click();
  URL.revokeObjectURL(url); toast('Selected tasks exported');
}
function setTaskHidden(t, hidden){
  if(!t) return;
  t.hiddenAt = hidden ? (t.hiddenAt || new Date().toISOString()) : '';
}
function setSelectedTasksHidden(locs, hidden){
  locs.forEach(({t})=>setTaskHidden(t, hidden));
  SELECTED_TASKS.clear();
  save();
  renderContent();
  toast(`${locs.length} task${locs.length===1?'':'s'} ${hidden?'hidden':'unhidden'}`);
}
function archiveSelectedTasks(locs){
  confirmModal(`Archive ${locs.length} task${locs.length===1?'':'s'}?`,'Archived tasks are removed from the board and kept in your exported data.',()=>{
    DATA.archive.push(...locs.map(({b,g,t})=>({id:uid('arch'),archivedAt:Date.now(),boardId:b.id,boardName:b.name,groupId:g.id,groupName:g.name,task:cloneTask(t)})));
    removeTaskIds(locs.map(x=>x.t.id));
    SELECTED_TASKS.clear(); save(); closeModal(); renderContent(); toast('Task archived');
  });
}
function deleteSelectedTasks(locs){
  confirmModal(`Delete ${locs.length} task${locs.length===1?'':'s'}?`,'This permanently removes the selected tasks from the board.',()=>{
    removeTaskIds(locs.map(x=>x.t.id));
    SELECTED_TASKS.clear(); save(); closeModal(); renderContent(); toast('Task deleted');
  });
}
function convertSelectedTasks(locs){
  locs.forEach(({b,g,t})=>{
    DATA.docs.unshift({id:uid('doc'),title:`${t.name||'Untitled task'} - Task brief`,template:'Converted task',
      content:`## Task\n\n${t.name||'Untitled task'}\n\n## Board\n\n${b.name}\n\n## Group\n\n${g.name}\n\n## Status\n\n${statusOptionFor(b,t.status).label||''}\n\n## Timeline\n\n${taskTimelineLabel(t)||''}\n\n## Notes\n\n${t.notes||''}\n`,
      updated:Date.now()});
  });
  SELECTED_TASKS.clear(); save(); renderContent(); toast('Task converted to doc');
}
function moveSelectedTasks(locs){
  const b=board(CURRENT_BOARD);
  modal('Move tasks','Choose the group where the selected tasks should go.',
    `<div class="field"><label>Move to group</label><select id="m-move-group">
      ${b.groups.map(g=>`<option value="${g.id}">${esc(g.name)}</option>`).join('')}
    </select></div>`,
    ()=>{
      const gid=$('#m-move-group').value; const target=b.groups.find(g=>g.id===gid);
      if(target){
        locs.forEach(({g,t})=>{
          if(g.id===target.id && g.tasks.some(x=>x.id===t.id)) return;
          const moved = detachTaskById(g,t.id);
          if(moved) target.tasks.push(moved);
        });
      }
      SELECTED_TASKS.clear(); save(); closeModal(); renderContent(); toast('Tasks moved');
    }, 'Move tasks');
}
function removeTaskIds(ids){
  const set=new Set(ids);
  DATA.boards.forEach(b=>b.groups.forEach(g=>{ g.tasks = filterRemovedTasks(g.tasks, set); }));
}
function filterRemovedTasks(tasks,set){
  return (tasks || []).filter(t=>!set.has(t.id)).map(t=>{
    if(Array.isArray(t.subtasks)) t.subtasks = filterRemovedTasks(t.subtasks, set);
    return t;
  });
}
function taskSiblingList(loc){ return loc.parent ? loc.parent.subtasks : loc.g.tasks; }
function detachTaskById(g,tid){ return detachTaskFromList(g.tasks,tid); }
function detachTaskFromList(tasks,tid){
  const idx=(tasks || []).findIndex(t=>t.id===tid);
  if(idx>=0) return tasks.splice(idx,1)[0];
  for(const t of tasks || []){
    const found = detachTaskFromList(t.subtasks || [], tid);
    if(found) return found;
  }
  return null;
}
function normalizeBulkLocations(locs){
  const selectedIds = new Set(locs.map(({t})=>t.id));
  return locs.filter(loc=>{
    let parent = loc.parent;
    while(parent){
      if(selectedIds.has(parent.id)) return false;
      parent = findTaskLocationById(parent.id)?.parent || null;
    }
    return true;
  });
}
function cloneTask(t){ return JSON.parse(JSON.stringify(t)); }
function cloneTaskWithNewIds(t){
  const copy = cloneTask(t);
  refreshTaskIds(copy);
  return copy;
}
function refreshTaskIds(t){
  t.id = uid('t');
  (t.subtasks || []).forEach(refreshTaskIds);
}

/* ---------- pickers (dropdowns) ---------- */
function buildDD(anchor, html){
  if(openDD) closeDD();
  const dd=document.createElement('div'); dd.className='dd-menu'; dd.innerHTML=html;
  dd.__anchor = anchor;
  const scope = dropdownEditScope(anchor);
  if(scope?.type && scope.id){
    dd.dataset.editType = scope.type;
    dd.dataset.editId = scope.id;
    dd.dataset.editField = scope.field || scope.type;
  }
  dd.addEventListener('click', e=>e.stopPropagation());
  document.body.appendChild(dd);
  positionDropdown(dd, anchor);
  openDD=dd; return dd;
}
function positionDropdown(dd, anchor=dd?.__anchor){
  if(!dd || !anchor?.getBoundingClientRect) return;
  const rect = anchor.getBoundingClientRect();
  const pad = 10;
  dd.style.maxHeight = '';
  const ddRect = dd.getBoundingClientRect();
  const width = Math.min(ddRect.width || 180, window.innerWidth - pad * 2);
  const height = Math.min(ddRect.height || 0, window.innerHeight - pad * 2);
  let left = Math.min(Math.max(rect.left, pad), window.innerWidth - width - pad);
  const belowTop = rect.bottom + 6;
  const aboveTop = rect.top - height - 6;
  const belowSpace = window.innerHeight - belowTop - pad;
  const aboveSpace = rect.top - pad - 6;
  let top = belowTop;
  let maxHeight = Math.max(140, belowSpace);
  if(height > belowSpace && aboveSpace > belowSpace){
    top = Math.max(pad, aboveTop);
    maxHeight = Math.max(140, aboveSpace);
  }
  if(top < pad) top = pad;
  if(top + height > window.innerHeight - pad) maxHeight = Math.max(140, window.innerHeight - top - pad);
  dd.style.left = `${left}px`;
  dd.style.top = `${top}px`;
  dd.style.maxHeight = `${Math.min(maxHeight, window.innerHeight - pad * 2)}px`;
}
function parseKey(el,attr){ return el.dataset[attr].split('|'); }
function taskLocationFromKey(tid,gid='',boardId=''){
  return (boardId ? findTaskLocation(boardId, tid) : null)
    || (CURRENT_BOARD ? findTaskLocation(CURRENT_BOARD, tid) : null)
    || findTaskLocationById(tid)
    || null;
}

function pickTaskColor(el){
  const [tid,rawGid]=parseKey(el,'taskColorPick');
  const loc = (CURRENT_BOARD ? findTaskLocation(CURRENT_BOARD, tid) : null) || findTaskLocationById(tid);
  const t=loc?.t || findTask(rawGid,tid);
  if(!t) return;
  const selected = taskColor(t);
  const html = `<div class="task-color-picker">
    <div class="control-title">Task calendar color</div>
    <div class="task-color-grid">
      ${TASK_PHASE_COLORS.map(color=>`<button type="button" class="${selected===color?'selected':''}" data-task-color="${esc(color)}" style="--task-color:${esc(color)}" title="${esc(color)}"><span></span></button>`).join('')}
    </div>
    <button type="button" class="dd-opt task-color-clear" data-task-color="">Clear color</button>
  </div>`;
  const dd=buildDD(el, html);
  dd.querySelectorAll('[data-task-color]').forEach(btn=>btn.onclick=()=>{
    setTaskColor(t, btn.dataset.taskColor || '');
    save();
    closeDD();
    if(TASK_DETAIL) render();
    else { renderContent(); wireContent(); }
  });
}

function pickStatus(el){
  const [tid,gid]=parseKey(el,'statusPick');
  const loc = taskLocationFromKey(tid,gid);
  const b=loc?.b || board(CURRENT_BOARD);
  if(!b) return;
  const dd=buildDD(el, Object.entries(boardStatusOptions(b)).map(([k,s])=>
    `<div class="dd-opt" data-k="${k}"><div class="swatch" style="background:${s.color}">${esc(s.label)}</div></div>`).join('')+
    `<div class="dd-divider"></div><div class="dd-opt" data-edit-labels="status">${I.plus} Edit labels</div>`);
  dd.querySelectorAll('[data-k]').forEach(o=>o.onclick=()=>{ setField(loc?.g?.id || gid,tid,'status',o.dataset.k,{boardId:b.id}); });
  dd.querySelector('[data-edit-labels]')?.addEventListener('click',()=>{ closeDD(); editBoardLabels('status'); });
}
function pickPriority(el){
  const [tid,gid]=parseKey(el,'priorityPick');
  const loc = taskLocationFromKey(tid,gid);
  const b=loc?.b || board(CURRENT_BOARD);
  if(!b) return;
  const dd=buildDD(el, Object.entries(boardPriorityOptions(b)).map(([k,p])=>
    `<div class="dd-opt" data-k="${k}"><div class="swatch" style="background:${p.color}">${esc(p.label)}</div></div>`).join('')+
    `<div class="dd-divider"></div><div class="dd-opt" data-edit-labels="priority">${I.plus} Edit labels</div>`);
  dd.querySelectorAll('[data-k]').forEach(o=>o.onclick=()=>{ setField(loc?.g?.id || gid,tid,'priority',o.dataset.k,{boardId:b.id}); });
  dd.querySelector('[data-edit-labels]')?.addEventListener('click',()=>{ closeDD(); editBoardLabels('priority'); });
}
function pickOwner(el){
  const [tid,rawGid]=parseKey(el,'ownerPick');
  const loc = taskLocationFromKey(tid,rawGid);
  const b=loc?.b || board(CURRENT_BOARD);
  const t=loc?.t || findTask(rawGid,tid);
  const gid=loc?.g?.id || rawGid;
  if(!b || !t) return;
  const selected = new Set(taskOwnerIds(t));
  const html = `<div class="owner-picker">
    <div class="owner-picker-head">
      <div>
        <div class="control-title">Task owners</div>
        <div class="owner-picker-hint">Checks save automatically.</div>
      </div>
      <button class="btn ghost sm" data-owner-clear>Clear</button>
    </div>
    ${DATA.team.map(m=>`<label class="owner-picker-row">
      <input type="checkbox" data-owner-check="${m.id}" ${selected.has(m.id)?'checked':''}>
      <span class="owner-av" style="background:${esc(m.color)}">${esc(m.initials)}</span>
      <span>${esc(m.name)}</span>
      ${m.email?`<em>${esc(m.email)}</em>`:''}
    </label>`).join('')}
  </div>`;
  const dd=buildDD(el, html);
  const applyOwners = (ids,{close=false, rerender=false}={})=>{
    const before = taskOwnerIds(t);
    const beforePrimary = t.owner || '';
    const next = validMemberIds(ids);
    if(before.join('|') === next.join('|')) return;
    setTaskOwnerIds(t, next);
    if(t.owner && t.owner !== beforePrimary) t.assignedBy = currentMemberId();
    const added = taskOwnerIds(t).filter(id=>!before.includes(id));
    if(added.length){
      recordUpdate({
        boardId:b.id,
        groupId:gid,
        taskId:tid,
        mentions:added,
        kind:'audit',
        visible:false,
        body:`Assigned "${t.name}" to ${taskOwnerNames(t)}.`
      });
      createTaskAssignmentNotifications({boardId:b.id, groupId:gid, taskId:tid, task:t, recipients:added});
    }
    if(t.owner && t.owner !== beforePrimary){
      applyAutomations({type:'field_change',boardId:b.id,groupId:gid,taskId:tid,field:'owner',value:t.owner});
    }
    save();
    refreshOwnerPickerAnchors(tid, gid, t);
    if(close) closeDD();
    if(rerender) render();
  };
  dd.querySelectorAll('[data-owner-check]').forEach(box=>box.addEventListener('change',()=>{
    const ids = [...dd.querySelectorAll('[data-owner-check]:checked')].map(box=>box.dataset.ownerCheck);
    applyOwners(ids);
  }));
  dd.querySelector('[data-owner-clear]')?.addEventListener('click',()=>{
    dd.querySelectorAll('[data-owner-check]').forEach(box=>{ box.checked = false; });
    applyOwners([]);
  });
}

function refreshOwnerPickerAnchors(tid,gid,t){
  const key = `${tid}|${gid}`;
  document.querySelectorAll(`[data-owner-pick="${CSS.escape(key)}"]`).forEach(anchor=>{
    if(!anchor.classList.contains('owner-stack') && !anchor.classList.contains('owner-empty')) return;
    const small = anchor.classList.contains('small');
    const limit = anchor.closest('.task-detail-summary') ? 4 : 3;
    const html = ownerStackHtml(taskOwnerIds(t), {pickKey:key, limit, small, emptyText:anchor.textContent?.trim() === '+' ? '+' : ''});
    const temp = document.createElement('span');
    temp.innerHTML = html;
    const next = temp.firstElementChild;
    if(next){
      next.onclick = e=>{ e.stopPropagation(); pickOwner(next); };
      anchor.replaceWith(next);
    }
  });
}
function pickerMonthIso(value){
  const date = validIsoDate(value) || todayIso();
  return date.slice(0,7);
}
function shiftPickerMonth(monthIso, amount){
  const base = /^\d{4}-\d{2}$/.test(String(monthIso || '')) ? monthIso : todayIso().slice(0,7);
  const d = new Date(`${base}-01T00:00:00`);
  d.setMonth(d.getMonth() + Number(amount || 0));
  return d.toISOString().slice(0,7);
}
function datePickerReminderOptions(selected){
  return [
    ['none','None'],
    ['at_time','At time'],
    ['day_before','1 day before'],
    ['week_before','1 week before']
  ].map(([value,label])=>`<option value="${value}" ${selected===value?'selected':''}>${label}</option>`).join('');
}
function taskDatePickerGridHtml(state){
  const baseMonth = /^\d{4}-\d{2}$/.test(String(state.month || '')) ? state.month : pickerMonthIso(state.start || state.end);
  const first = new Date(`${baseMonth}-01T00:00:00`);
  const month = first.getMonth();
  const start = new Date(first);
  start.setDate(first.getDate() - first.getDay());
  const selectedStart = validIsoDate(state.start);
  const selectedEnd = state.includeEnd ? (validIsoDate(state.end) || selectedStart) : selectedStart;
  let rangeStart = selectedStart;
  let rangeEnd = selectedEnd;
  if(rangeStart && rangeEnd && isoDateCompare(rangeStart,rangeEnd) > 0) [rangeStart,rangeEnd] = [rangeEnd,rangeStart];
  const today = todayIso();
  const days = [];
  for(let i=0;i<42;i++){
    const d = new Date(start);
    d.setDate(start.getDate() + i);
    const iso = d.toISOString().slice(0,10);
    const inRange = state.mode === 'timeline' && state.includeEnd && rangeStart && rangeEnd && isoDateCompare(iso,rangeStart) >= 0 && isoDateCompare(iso,rangeEnd) <= 0;
    const classes = [
      'date-picker-day',
      d.getMonth() !== month ? 'muted' : '',
      iso === selectedStart ? 'selected start' : '',
      iso === selectedEnd ? 'selected end' : '',
      inRange ? 'in-range' : '',
      iso === today ? 'today' : ''
    ].filter(Boolean).join(' ');
    days.push(`<button type="button" class="${classes}" data-picker-date="${iso}">${d.getDate()}</button>`);
  }
  return `${['Su','Mo','Tu','We','Th','Fr','Sa'].map(day=>`<span class="date-picker-dow">${day}</span>`).join('')}${days.join('')}`;
}
function taskDatePickerHtml(state){
  const month = /^\d{4}-\d{2}$/.test(String(state.month || '')) ? state.month : pickerMonthIso(state.start || state.end);
  const monthLabel = new Date(`${month}-01T00:00:00`).toLocaleDateString(undefined,{month:'long',year:'numeric'});
  const isTimeline = state.mode === 'timeline';
  const endInput = isTimeline && state.includeEnd ? `
    <input type="date" class="${state.active==='end'?'active':''}" data-picker-input="end" value="${esc(state.end || state.start || '')}" aria-label="End date">` : '';
  const timeInputs = state.includeTime ? `
    <div class="date-picker-time-row">
      <label><span>${isTimeline?'Start time':'Time'}</span><input type="time" data-picker-time="start" value="${esc(state.startTime || '')}"></label>
      ${isTimeline && state.includeEnd ? `<label><span>End time</span><input type="time" data-picker-time="end" value="${esc(state.endTime || '')}"></label>` : ''}
    </div>` : '';
  return `<div class="task-date-picker" data-picker-mode="${esc(state.mode)}">
    <div class="date-picker-fields">
      <input type="date" class="${state.active==='start'?'active':''}" data-picker-input="start" value="${esc(state.start || '')}" aria-label="Start date">
      ${endInput}
    </div>
    ${timeInputs}
    <div class="date-picker-head">
      <b>${esc(monthLabel)}</b>
      <button type="button" data-picker-today>Today</button>
      <button type="button" data-picker-month="-1" aria-label="Previous month">&lt;</button>
      <button type="button" data-picker-month="1" aria-label="Next month">&gt;</button>
    </div>
    <div class="date-picker-grid">${taskDatePickerGridHtml({...state, month})}</div>
    <div class="date-picker-options">
      ${isTimeline ? `<button type="button" class="date-picker-option" data-picker-toggle="end"><span>End date</span><i class="${state.includeEnd?'on':''}"></i></button>` : ''}
      <button type="button" class="date-picker-option" data-picker-toggle="time"><span>Include time</span><i class="${state.includeTime?'on':''}"></i></button>
      <label class="date-picker-select"><span>Remind</span><select data-picker-reminder>${datePickerReminderOptions(state.reminder || 'none')}</select></label>
    </div>
    <div class="date-picker-footer">
      <button type="button" class="btn ghost sm" data-picker-clear>Clear</button>
      <button type="button" class="btn sm" data-picker-apply>Done</button>
    </div>
  </div>`;
}
function openTaskDatePicker(anchor,{mode,tid,gid}){
  const initialLoc = taskLocationFromKey(tid,gid);
  const t=initialLoc?.t || findTask(gid,tid);
  if(!t) return;
  const taskLoc = ()=>taskLocationFromKey(tid,gid,initialLoc?.b?.id || '');
  const isTimeline = mode === 'timeline';
  const scheduleStart = isTimeline ? validIsoDate(t.tlStart) : validIsoDate(t.due);
  const scheduleEnd = isTimeline ? (validIsoDate(t.tlEnd) || scheduleStart) : scheduleStart;
  const initialStart = scheduleStart || todayIso();
  const initialEnd = isTimeline ? (scheduleEnd || initialStart) : '';
  const state = {
    mode:isTimeline ? 'timeline' : 'due',
    active:'start',
    start:initialStart,
    end:initialEnd,
    includeEnd:isTimeline,
    includeTime:isTimeline
      ? Boolean(validTimeValue(taskMetaValue(t,TASK_TIMELINE_START_TIME_KEY)) || validTimeValue(taskMetaValue(t,TASK_TIMELINE_END_TIME_KEY)))
      : Boolean(validTimeValue(taskMetaValue(t,TASK_DUE_TIME_KEY))),
    startTime:isTimeline ? validTimeValue(taskMetaValue(t,TASK_TIMELINE_START_TIME_KEY)) : validTimeValue(taskMetaValue(t,TASK_DUE_TIME_KEY)),
    endTime:isTimeline ? validTimeValue(taskMetaValue(t,TASK_TIMELINE_END_TIME_KEY)) : '',
    reminder:taskMetaValue(t,TASK_DATE_REMINDER_KEY) || 'none',
    month:pickerMonthIso(scheduleStart || scheduleEnd)
  };
  if(isTimeline && state.includeEnd && state.start && state.end && isoDateCompare(state.start,state.end) > 0){
    [state.start,state.end] = [state.end,state.start];
  }
  const dd=buildDD(anchor,taskDatePickerHtml(state));
  dd.classList.add('task-date-picker-popover');
  positionDropdown(dd, anchor);
  let dirty = false;
  let suppressCloseCommit = false;
  const normalizeRange = ()=>{
    if(!isTimeline || !state.includeEnd || !state.start || !state.end) return;
    if(isoDateCompare(state.start,state.end) > 0) [state.start,state.end] = [state.end,state.start];
  };
  const markDirty = ()=>{ dirty = true; };
  const commitSchedule = ({rerender=false} = {})=>{
    if(!dirty) return false;
    const loc = taskLoc();
    const b = loc?.b;
    const g = loc?.g;
    const t2 = loc?.t;
    if(!b || !g || !t2) return false;
    normalizeRange();
    const start = validIsoDate(state.start);
    const end = state.includeEnd ? (validIsoDate(state.end) || start) : start;
    if(!start){
      dirty = false;
      return false;
    }
    const previousDue = t2.due || '';
    if(isTimeline){
      setTaskTimeline(t2,{
        start,
        end,
        startTime:state.includeTime ? validTimeValue(state.startTime) : '',
        endTime:state.includeTime && state.includeEnd ? validTimeValue(state.endTime) : (state.includeTime ? validTimeValue(state.startTime) : ''),
        reminder:state.reminder || 'none'
      });
      applyAutomations({type:'field_change',boardId:b.id,groupId:g.id,taskId:tid,field:'timeline',value:taskTimelineLabel(t2)});
    } else {
      setTaskDueDate(t2,{
        date:start,
        time:state.includeTime ? validTimeValue(state.startTime) : '',
        reminder:state.reminder || 'none'
      });
      applyAutomations({type:'field_change',boardId:b.id,groupId:g.id,taskId:tid,field:'due',value:taskDueLabel(t2)});
      createDueChangedNotifications({boardId:b.id, groupId:g.id, taskId:tid, task:t2, previousDue, nextDue:t2.due});
    }
    save();
    dirty = false;
    if(rerender){
      renderContent();
      wireContent();
    }
    return true;
  };
  const apply = ()=>{
    markDirty();
    commitSchedule();
    suppressCloseCommit = true;
    closeDD({commit:false});
    renderContent();
    wireContent();
  };
  const clear = ()=>{
    const loc = taskLoc();
    const t2 = loc?.t;
    if(!loc?.b || !loc?.g || !t2) return;
    if(isTimeline) clearTaskTimeline(t2);
    else clearTaskDueDate(t2);
    save();
    dirty = false;
    suppressCloseCommit = true;
    closeDD({commit:false});
    renderContent();
    wireContent();
  };
  dd.__beforeClose = ()=>{
    if(suppressCloseCommit) return;
    const saved = commitSchedule();
    if(saved) setTimeout(()=>{ renderContent(); wireContent(); }, 0);
  };
  const renderPicker = ()=>{
    normalizeRange();
    dd.innerHTML = taskDatePickerHtml(state);
    wirePicker();
    positionDropdown(dd, anchor);
  };
  function wirePicker(){
    dd.querySelectorAll('[data-picker-input]').forEach(input=>{
      input.addEventListener('focus',()=>{
        state.active = input.dataset.pickerInput || 'start';
        dd.querySelectorAll('[data-picker-input]').forEach(next=>{
          next.classList.toggle('active', next.dataset.pickerInput === state.active);
        });
      });
      input.addEventListener('change',()=>{
        const key = input.dataset.pickerInput === 'end' ? 'end' : 'start';
        state[key] = validIsoDate(input.value);
        if(key === 'start' && state.includeEnd && (!dirty || !state.end || isoDateCompare(state.end,state.start) < 0)) state.end = state.start;
        if(state[key]) state.month = pickerMonthIso(state[key]);
        markDirty();
        renderPicker();
      });
    });
    dd.querySelectorAll('[data-picker-time]').forEach(input=>{
      input.addEventListener('change',()=>{
        if(input.dataset.pickerTime === 'end') state.endTime = validTimeValue(input.value);
        else state.startTime = validTimeValue(input.value);
        markDirty();
      });
    });
    dd.querySelectorAll('[data-picker-date]').forEach(btn=>btn.addEventListener('click',()=>{
      const iso = validIsoDate(btn.dataset.pickerDate);
      if(!iso) return;
      if(isTimeline && state.includeEnd && state.active === 'end'){
        state.end = iso;
      } else {
        state.start = iso;
        if(isTimeline && state.includeEnd){
          if(!dirty || !state.end || isoDateCompare(state.end,state.start) < 0) state.end = iso;
          state.active = 'end';
        }
      }
      state.month = pickerMonthIso(iso);
      markDirty();
      renderPicker();
    }));
    dd.querySelector('[data-picker-today]')?.addEventListener('click',()=>{
      const iso = todayIso();
      if(isTimeline && state.includeEnd && state.active === 'end') state.end = iso;
      else state.start = iso;
      state.month = pickerMonthIso(iso);
      markDirty();
      renderPicker();
    });
    dd.querySelectorAll('[data-picker-month]').forEach(btn=>btn.addEventListener('click',()=>{
      state.month = shiftPickerMonth(state.month, btn.dataset.pickerMonth);
      renderPicker();
    }));
    dd.querySelector('[data-picker-toggle="end"]')?.addEventListener('click',()=>{
      state.includeEnd = !state.includeEnd;
      if(state.includeEnd){
        state.end = state.end || state.start || todayIso();
        state.active = 'end';
      } else {
        state.end = '';
        state.endTime = '';
        state.active = 'start';
      }
      markDirty();
      renderPicker();
    });
    dd.querySelector('[data-picker-toggle="time"]')?.addEventListener('click',()=>{
      state.includeTime = !state.includeTime;
      if(state.includeTime && !state.startTime) state.startTime = '09:00';
      if(state.includeTime && isTimeline && state.includeEnd && !state.endTime) state.endTime = '17:00';
      markDirty();
      renderPicker();
    });
    dd.querySelector('[data-picker-reminder]')?.addEventListener('change',e=>{ state.reminder = e.target.value || 'none'; markDirty(); });
    dd.querySelector('[data-picker-clear]')?.addEventListener('click',clear);
    dd.querySelector('[data-picker-apply]')?.addEventListener('click',apply);
  }
  wirePicker();
}
function pickDate(el){
  const [tid,gid]=parseKey(el,'datePick');
  openTaskDatePicker(el,{mode:'due',tid,gid});
}
function pickTimeline(el){
  const [tid,gid]=parseKey(el,'tlPick');
  openTaskDatePicker(el,{mode:'timeline',tid,gid});
}
function pickCustomDate(el){
  const [tid,gid,cid]=parseKey(el,'customDatePick');
  const t=findTask(gid,tid); const val=(t.custom||{})[cid]||'';
  const dd=buildDD(el,`<div style="padding:4px"><input type="date" id="dd-custom-date" value="${val}" style="padding:7px;border:1px solid var(--line2);border-radius:6px">
    <div style="display:flex;gap:6px;margin-top:6px"><button class="btn ghost sm" id="dd-custom-clear">Clear</button><button class="btn sm" id="dd-custom-done">Done</button></div></div>`);
  $('#dd-custom-date').onchange=()=>{
    const value = $('#dd-custom-date').value;
    setCustomField(gid,tid,cid,value,{close:false,renderView:false});
    updateDateCellPreview(el,value);
  };
  $('#dd-custom-clear').onclick=()=>{ setCustomField(gid,tid,cid,'',{close:false,renderView:false}); updateDateCellPreview(el,''); $('#dd-custom-date').value=''; };
  $('#dd-custom-done').onclick=()=>{ closeDD(); renderContent(); };
}
function pickCustomStatus(el){
  const [tid,gid,cid]=parseKey(el,'customStatusPick');
  const b=board(CURRENT_BOARD);
  const opts = Object.entries(boardStatusOptions(b)).map(([k,s])=>`<div class="dd-opt" data-k="${k}"><div class="swatch" style="background:${s.color}">${esc(s.label)}</div></div>`).join('')+
    `<div class="dd-opt" data-k=""><div class="swatch" style="background:#c4c4c4">Clear</div></div>
     <div class="dd-divider"></div><div class="dd-opt" data-edit-labels="status">${I.plus} Edit labels</div>`;
  const dd=buildDD(el, opts);
  dd.querySelectorAll('[data-k]').forEach(o=>o.onclick=()=>setCustomField(gid,tid,cid,o.dataset.k));
  dd.querySelector('[data-edit-labels]')?.addEventListener('click',()=>{ closeDD(); editBoardLabels('status'); });
}
function pickCustomRecurrence(el){
  const [tid,gid,cid]=parseKey(el,'customRecurrencePick');
  const t=findTask(gid,tid);
  const b=board(CURRENT_BOARD);
  const current = normalizeRecurrenceValue((t?.custom || {})[cid] || taskRecurrence(t,b));
  const currentMeta = recurrenceMeta(current);
  const opts = Object.entries(RECURRENCE_OPTIONS).map(([key,opt])=>`
    <button class="dd-opt recurrence-option ${currentMeta.key===key?'selected':''}" data-recurrence-k="${key}">
      <span class="recurrence-dot" style="background:${esc(opt.color)}"></span>
      <span>${esc(opt.label)}</span>
    </button>`).join('');
  const customInterval = currentMeta.custom ? currentMeta.interval : 2;
  const customUnit = currentMeta.custom ? currentMeta.unit : 'weeks';
  const startValue = currentMeta.startsAt ? localDateTimeValue(currentMeta.startsAt) : defaultRecurrenceStart();
  const unitOptions = Object.entries(RECURRENCE_UNITS).map(([unit,meta])=>
    `<option value="${unit}" ${customUnit===unit?'selected':''}>${esc(meta.label)}</option>`
  ).join('');
  const dd=buildDD(el, `<div class="recurrence-menu">
    <div class="control-title">Recurring schedule</div>
    ${opts}
    <div class="dd-divider"></div>
    <div class="recurrence-custom">
      <label>Next reset date and time</label>
      ${recurrenceStartControlHtml(startValue)}
      <button class="btn ghost sm" id="recurrence-date-apply">Apply date/time</button>
    </div>
    <div class="dd-divider"></div>
    <div class="recurrence-custom">
      <label>Custom interval</label>
      <div class="recurrence-custom-row">
        <input type="number" min="1" max="365" step="1" id="recurrence-custom-number" value="${esc(customInterval)}">
        <select id="recurrence-custom-unit">${unitOptions}</select>
      </div>
      <button class="btn sm" id="recurrence-custom-apply">Apply custom</button>
    </div>
  </div>`);
  wireRecurrenceStartControl(dd);
  const selectedStart = ()=>dd.querySelector('#recurrence-start-at')?.value || defaultRecurrenceStart();
  dd.querySelectorAll('[data-recurrence-k]').forEach(btn=>btn.onclick=()=>{
    const key = btn.dataset.recurrenceK;
    setCustomField(gid,tid,cid,key === 'off' ? 'off' : recurrenceValueWithStart(key, selectedStart()));
  });
  dd.querySelector('#recurrence-date-apply')?.addEventListener('click',()=>{
    const key = currentMeta.key === 'off' ? 'daily' : currentMeta.key;
    setCustomField(gid,tid,cid,recurrenceValueWithStart(key, selectedStart()));
  });
  dd.querySelector('#recurrence-custom-apply')?.addEventListener('click',()=>{
    const interval = dd.querySelector('#recurrence-custom-number')?.value || 1;
    const unit = dd.querySelector('#recurrence-custom-unit')?.value || 'days';
    setCustomField(gid,tid,cid,recurrenceValueWithStart(customRecurrenceValue(interval,unit), selectedStart()));
  });
}
function pickFilesCell(el){
  const [tid,gid,cid]=parseKey(el,'fileCellPick');
  const t=findTask(gid,tid);
  if(!t)return;
  const selectedIds = taskAssetIdsForColumn(t,cid);
  const selected = selectedIds.map(id=>asset(id)).filter(Boolean);
  const available = (DATA.assets || []).filter(a=>!selectedIds.includes(a.id));
  const html = `<div class="asset-picker">
    <div class="control-title">Files / Media</div>
    <button class="btn sm asset-picker-primary" data-upload-cell>${I.plus}<span>Upload file</span></button>
    <input type="file" multiple data-file-cell-input hidden>
    ${selected.length ? `<div class="picker-section-title">Attached</div>
      ${selected.map(a=>`<div class="picker-row">
        <button class="picker-row-main" data-open-asset="${a.id}"><span class="asset-type">${esc(assetIcon(a))}</span><span>${esc(a.title)}</span></button>
        <button class="icon-btn sm-icon" data-detach-asset="${a.id}" title="Remove">${I.close}</button>
      </div>`).join('')}` : '<div class="picker-empty">No files attached yet.</div>'}
    <div class="picker-section-title">From Docs Hub</div>
    ${available.length ? available.slice(0,20).map(a=>`<button class="picker-row-main full" data-attach-asset="${a.id}">
      <span class="asset-type">${esc(assetIcon(a))}</span><span>${esc(a.title)}</span><em>${esc(assetSizeLabel(a.size))}</em>
    </button>`).join('') : '<div class="picker-empty">No files in Docs Hub yet.</div>'}
  </div>`;
  const dd=buildDD(el, html);
  dd.querySelector('[data-upload-cell]')?.addEventListener('click',()=>dd.querySelector('[data-file-cell-input]')?.click());
  dd.querySelector('[data-file-cell-input]')?.addEventListener('change',e=>uploadAssetFiles(Array.from(e.target.files || []), {attachTo:{gid,tid,cid}}));
  dd.querySelectorAll('[data-attach-asset]').forEach(btn=>btn.onclick=()=>attachAssetsToTaskCell(gid,tid,cid,[btn.dataset.attachAsset]));
  dd.querySelectorAll('[data-detach-asset]').forEach(btn=>btn.onclick=e=>{ e.stopPropagation(); detachAssetFromTaskCell(gid,tid,cid,btn.dataset.detachAsset); });
  dd.querySelectorAll('[data-open-asset]').forEach(btn=>btn.onclick=e=>{ e.stopPropagation(); void openAsset(btn.dataset.openAsset); });
}
function pickDocumentCell(el){
  const [tid,gid,cid]=parseKey(el,'docCellPick');
  const t=findTask(gid,tid);
  if(!t)return;
  const did = taskDocIdForColumn(t,cid);
  const docs = DATA.docs || [];
  const html = `<div class="asset-picker">
    <div class="control-title">Document</div>
    <button class="btn sm asset-picker-primary" data-create-cell-doc>${I.plus}<span>Create document</span></button>
    ${did ? `<button class="btn ghost sm asset-picker-primary" data-clear-cell-doc>Clear document</button>` : ''}
    <div class="picker-section-title">From Docs Hub</div>
    ${docs.length ? docs.slice(0,24).map(d=>`<button class="picker-row-main full ${d.id===did?'selected':''}" data-attach-doc="${d.id}">
      <span class="asset-type">DOC</span><span>${esc(d.title)}</span><em>${new Date(d.updated).toLocaleDateString()}</em>
    </button>`).join('') : '<div class="picker-empty">No documents yet.</div>'}
  </div>`;
  const dd=buildDD(el, html);
  dd.querySelector('[data-create-cell-doc]')?.addEventListener('click',()=>createDocForTaskCell(gid,tid,cid));
  dd.querySelector('[data-clear-cell-doc]')?.addEventListener('click',()=>attachDocToTaskCell(gid,tid,cid,''));
  dd.querySelectorAll('[data-attach-doc]').forEach(btn=>btn.onclick=()=>attachDocToTaskCell(gid,tid,cid,btn.dataset.attachDoc));
}
function findTaskForWrite(gid,tid,boardId=''){
  return (boardId ? findTaskLocation(boardId, tid)?.t : null) || findTaskLocationById(tid)?.t || findTask(gid,tid);
}
function attachAssetsToTaskCell(gid,tid,cid,assetIds,boardId=''){
  const t=findTaskForWrite(gid,tid,boardId);
  if(!t)return;
  if(!t.custom) t.custom = {};
  const next = [...new Set([...taskAssetIdsForColumn(t,cid), ...(assetIds || []).filter(Boolean)])];
  t.custom[cid] = next;
  save();
  closeDD();
  if(TASK_DETAIL) render();
  else { renderContent(); wireContent(); }
  toast('File attached');
}
function detachAssetFromTaskCell(gid,tid,cid,assetId){
  const t=findTaskForWrite(gid,tid);
  if(!t)return;
  if(!t.custom) t.custom = {};
  t.custom[cid] = taskAssetIdsForColumn(t,cid).filter(id=>id!==assetId);
  save();
  closeDD();
  if(TASK_DETAIL) render();
  else { renderContent(); wireContent(); }
  toast('File removed');
}
function attachDocToTaskCell(gid,tid,cid,docId){
  const t=findTaskForWrite(gid,tid);
  if(!t)return;
  if(!t.custom) t.custom = {};
  t.custom[cid] = docId || '';
  save();
  closeDD();
  renderContent();
  wireContent();
  toast(docId ? 'Document attached' : 'Document cleared');
}
function createDocForTaskCell(gid,tid,cid){
  const t=findTask(gid,tid);
  const title = `${t?.name || 'Task'} notes`;
  const d={id:uid('doc'),title,template:'Task document',
    content:`## ${title}\n\n## Summary\n\n## Assets\n\n## Next steps\n`,updated:Date.now()};
  DATA.docs.unshift(d);
  attachDocToTaskCell(gid,tid,cid,d.id);
  openDoc(d.id);
}
function ensureBoardColumnByType(b,type,name){
  normalizeBoard(b);
  let col = (b.columns || []).find(c=>c.type === type);
  if(col) return col;
  col = {id:uid('c'), name, type};
  b.columns.push(col);
  b.columnOrder.push(col.id);
  return col;
}
function taskPickerOptions(){
  return allVisibleTasks().map(t=>`<option value="${esc(t._board.id)}|${esc(t._group.id)}|${esc(t.id)}">${esc(t._board.name)} / ${esc(t._group.name)} / ${esc(t.name)}</option>`).join('');
}

function parseTaskPickerValue(value){
  const parts = String(value || '').split('|');
  if(parts.length >= 3) return [parts[0] || '', parts[1] || '', parts[2] || ''];
  return [parts[0] || '', '', parts[1] || ''];
}

function findTaskLocationFromPicker(value){
  const [boardId, groupId, taskId] = parseTaskPickerValue(value);
  const b = boardId ? board(boardId) : null;
  const g = b && groupId ? b.groups.find(x=>x.id === groupId) : null;
  const found = g ? findTaskInGroup(g, taskId) : null;
  return found ? {b,g,t:found.t,parent:found.parent} : ((boardId ? findTaskLocation(boardId, taskId) : null) || findTaskLocationById(taskId));
}

function linkDocToTask(docId){
  const d = doc(docId);
  if(!d) return;
  const options = taskPickerOptions();
  if(!options){ toast('Create a task before linking docs'); return; }
  modal('Link document to task',`Choose where "${d.title}" should appear on a board.`,
    `<div class="field"><label>Task</label><select id="link-task-target">${options}</select></div>`,
    ()=>{
      const loc = findTaskLocationFromPicker($('#link-task-target').value);
      if(!loc){ toast('Task not found'); return; }
      const col = ensureBoardColumnByType(loc.b,'document','Document');
      if(!loc.t.custom) loc.t.custom = {};
      loc.t.custom[col.id] = docId;
      recordUpdate({boardId:loc.b.id,groupId:loc.g.id,taskId:loc.t.id,body:`Linked document "${d.title}" to "${loc.t.name}".`});
      save(); closeModal(); render(); toast('Document linked to task');
    },
    'Link document');
}
function linkAssetToTask(assetId){
  const a = asset(assetId);
  if(!a) return;
  const options = taskPickerOptions();
  if(!options){ toast('Create a task before linking files'); return; }
  modal('Link file to task',`Choose where "${a.title}" should appear on a board.`,
    `<div class="field"><label>Task</label><select id="link-task-target">${options}</select></div>`,
    ()=>{
      const loc = findTaskLocationFromPicker($('#link-task-target').value);
      if(!loc){ toast('Task not found'); return; }
      const col = ensureBoardColumnByType(loc.b,'files','Files');
      if(!loc.t.custom) loc.t.custom = {};
      loc.t.custom[col.id] = [...new Set([...taskAssetIdsForColumn(loc.t,col.id), assetId])];
      recordUpdate({boardId:loc.b.id,groupId:loc.g.id,taskId:loc.t.id,body:`Linked file "${a.title}" to "${loc.t.name}".`});
      save(); closeModal(); render(); toast('File linked to task');
    },
    'Link file');
}
function findTaskInGroup(g,tid){ return findTaskInList(g?.tasks || [], tid, null); }
function findTaskInList(tasks,tid,parent=null){
  for(const t of tasks || []){
    if(t.id===tid) return {t,parent};
    const found = findTaskInList(t.subtasks || [], tid, t);
    if(found) return found;
  }
  return null;
}
function findTask(gid,tid){ const b=board(CURRENT_BOARD); const g=b?.groups?.find(x=>x.id===gid); return g ? findTaskInGroup(g,tid)?.t : null; }
function closeDD({commit=true}={}){
  if(!openDD) return;
  const dd = openDD;
  openDD = null;
  if(commit && typeof dd.__beforeClose === 'function'){
    try{ dd.__beforeClose(); }catch(e){ console.warn('Dropdown close handler failed', e); }
  }
  dd.remove();
}
function updateDateCellPreview(el,value){
  if(!el) return;
  el.classList.toggle('empty', !value);
  el.innerHTML = `${I.cal}<span>${value ? fmtDate(value) : 'Set date'}</span>`;
}
function setField(gid,tid,field,val,{close=true,renderView=true,boardId=''}={}){
  const loc = taskLocationFromKey(tid,gid,boardId);
  const b=loc?.b || board(CURRENT_BOARD);
  const g=loc?.g;
  const t=loc?.t || findTask(gid,tid);
  if(b && g && t){
    const prevOwner = t.owner;
    if(field === 'owner') setTaskOwnerIds(t, val ? [val] : []);
    else t[field]=val;
    if(field==='status'){
      if(t.status==='done' && !t.completedAt) t.completedAt = new Date().toISOString();
      if(t.status!=='done') t.completedAt = '';
    }
    if(field==='owner' && val && val!==prevOwner){
      t.assignedBy = currentMemberId();
      const assignee = member(val);
      recordUpdate({
        boardId:b.id,
        groupId:g?.id || gid,
        taskId:tid,
        mentions:[val],
        kind:'audit',
        visible:false,
        body:`Assigned "${t.name}" to ${assignee ? '@'+assignee.name : 'a teammate'}.`
      });
      createTaskAssignmentNotifications({boardId:b.id, groupId:g.id, taskId:tid, task:t, recipients:[val]});
    }
    applyAutomations({type:'field_change',boardId:b.id,groupId:g.id,taskId:tid,field,value:val});
    save();
  }
  if(close) closeDD();
  if(renderView){ renderContent(); wireContent(); }
}
function setCustomField(gid,tid,cid,val,{close=true,renderView=true}={}){
  const t=findTask(gid,tid);
  const b=board(CURRENT_BOARD);
  const col=b ? findColumn(b,cid) : null;
  if(t){
    if(!t.custom) t.custom={};
    if(col?.type === 'recurrence'){
      const next = normalizeRecurrenceValue(val);
      t.custom[cid] = next === 'off' ? '' : next;
      setTaskRecurrence(t,next);
    } else {
      if(col?.type === 'url'){
        const duplicate = findDuplicateAssetUrl(val);
        if(duplicate) toast(duplicateAssetMessage(duplicate, val));
      }
      t.custom[cid]=val;
    }
    save();
  }
  if(close) closeDD();
  if(renderView){ renderContent(); wireContent(); }
}

/* ---------- automations ---------- */
function addAutomation(){
  const defaultBoard = board(CURRENT_BOARD) || DATA.boards[0];
  if(!defaultBoard){ toast('Create a board first'); return; }
  const boardOpts = DATA.boards.map(b=>`<option value="${b.id}" ${b.id===defaultBoard.id?'selected':''}>${esc(b.name)}</option>`).join('');
  const groupOpts = defaultBoard.groups.map(g=>`<option value="${g.id}">${esc(g.name)}</option>`).join('');
  const statusOpts = Object.entries(boardStatusOptions(defaultBoard)).map(([k,s])=>`<option value="${k}">${esc(s.label)}</option>`).join('');
  const priorityOpts = Object.entries(boardPriorityOptions(defaultBoard)).filter(([k])=>k!=='none').map(([k,p])=>`<option value="${k}">${esc(p.label)}</option>`).join('');
  modal('New automation','Create a board rule that runs when tasks change.',
    `<div class="field"><label>Rule name</label><input id="m-auto-name" value="Move done tasks"></div>
     <div class="field"><label>Board</label><select id="m-auto-board">${boardOpts}</select></div>
     <div class="field"><label>Trigger</label><select id="m-auto-trigger">
       <option value="status_done">Status changes to Done</option>
       <option value="owner_assigned">Owner gets assigned</option>
       <option value="due_past">Due date is past</option>
     </select></div>
     <div class="field"><label>Action</label><select id="m-auto-action">
       <option value="move_group">Move to group</option>
       <option value="set_status">Set status</option>
       <option value="set_priority">Set priority</option>
       <option value="notify">Notify task owners</option>
     </select></div>
     <div class="field"><label>Target group</label><select id="m-auto-group">${groupOpts}</select></div>
     <div class="field"><label>Target status</label><select id="m-auto-status">${statusOpts}</select></div>
     <div class="field"><label>Target priority</label><select id="m-auto-priority">${priorityOpts}</select></div>`,
    ()=>{
      const b=board($('#m-auto-board').value);
      const auto={id:uid('a'),name:$('#m-auto-name').value.trim()||'Untitled automation',active:true,
        boardId:b.id,trigger:$('#m-auto-trigger').value,action:$('#m-auto-action').value,
        targetGroupId:$('#m-auto-group').value,targetStatus:$('#m-auto-status').value,targetPriority:$('#m-auto-priority').value};
      DATA.automations.push(auto); save(); closeModal(); render(); toast('Automation created');
    }, 'Create automation');
  $('#m-auto-board').onchange=()=>{
    const b=board($('#m-auto-board').value);
    $('#m-auto-group').innerHTML=b.groups.map(g=>`<option value="${g.id}">${esc(g.name)}</option>`).join('');
  };
}
function toggleAutomation(id){
  const a=DATA.automations.find(x=>x.id===id); if(!a)return;
  a.active = a.active===false;
  save(); renderContent(); toast(a.active?'Automation active':'Automation paused');
}
function delAutomation(id){
  const a=DATA.automations.find(x=>x.id===id); if(!a)return;
  confirmModal(`Delete "${a.name}"?`,'This permanently removes the automation rule.',()=>{
    DATA.automations=DATA.automations.filter(x=>x.id!==id);
    save(); closeModal(); renderContent(); toast('Automation deleted');
  });
}
function applyAutomations(event){
  let changed=false;
  DATA.automations.filter(a=>a.active!==false && a.boardId===event.boardId).forEach(a=>{
    if(automationMatches(a,event)) changed = runAutomationAction(a,event) || changed;
  });
  return changed;
}
function automationMatches(a,event){
  const loc=findTaskLocation(event.boardId,event.taskId); if(!loc)return false;
  if(a.trigger==='status_done') return event.type==='field_change' && event.field==='status' && event.value==='done';
  if(a.trigger==='owner_assigned') return event.type==='field_change' && event.field==='owner' && !!event.value;
  if(a.trigger==='due_past') return isPast(loc.t.due) && (event.type==='manual' || event.field==='due');
  return false;
}
function runAutomationAction(a,event){
  const loc=findTaskLocation(event.boardId,event.taskId); if(!loc)return false;
  const {b,g,t}=loc;
  if(a.action==='move_group'){
    const target=b.groups.find(x=>x.id===a.targetGroupId);
    if(target && target.id!==g.id){ g.tasks=g.tasks.filter(x=>x.id!==t.id); target.tasks.push(t); return true; }
  }
  if(a.action==='set_status' && a.targetStatus && t.status!==a.targetStatus){ t.status=a.targetStatus; return true; }
  if(a.action==='set_priority' && a.targetPriority && t.priority!==a.targetPriority){ t.priority=a.targetPriority; return true; }
  if(a.action==='notify'){
    const recipients = taskOwnerIds(t).filter(id=>shouldCreateNotification('automation', b.id, id, {task:t}));
    recipients.forEach(recipientId=>{
      createNotification({
        recipientId,
        type:'automation',
        boardId:b.id,
        groupId:g.id,
        taskId:t.id,
        title:`Automation alert: ${t.name}`,
        body:`${a.name || 'A TaskTrack automation'} matched this task.`,
        data:{automationId:a.id, automationName:a.name || ''}
      });
    });
    return recipients.length > 0;
  }
  return false;
}
function runDueAutomationsNow(){
  let changed=false, matches=0;
  DATA.automations.filter(a=>a.active!==false && a.trigger==='due_past').forEach(a=>{
    const b=board(a.boardId); if(!b)return;
    b.groups.forEach(g=>g.tasks.slice().forEach(t=>{
      if(isPast(t.due)){ matches++; changed = runAutomationAction(a,{type:'manual',boardId:b.id,groupId:g.id,taskId:t.id}) || changed; }
    }));
  });
  if(changed) save();
  render();
  toast(matches ? 'Automations ran' : 'No matching timeline automations');
}
function findTaskLocation(boardId,taskId){
  const b=board(boardId); if(!b)return null;
  for(const g of b.groups){
    const found=findTaskInGroup(g,taskId);
    if(found)return {b,g,t:found.t,parent:found.parent};
  }
  return null;
}
function findTaskLocationById(taskId){
  for(const b of DATA.boards){
    for(const g of b.groups){
      const found=findTaskInGroup(g,taskId);
      if(found)return {b,g,t:found.t,parent:found.parent};
    }
  }
  return null;
}
function isPast(iso){
  if(!iso)return false;
  const d=new Date(iso+'T00:00'); const today=new Date(); today.setHours(0,0,0,0);
  return !isNaN(d) && d<today;
}

/* ============================================================
   ACTIONS — docs
   ============================================================ */
function triggerDocsUpload(){
  const input=$('#docs-file-upload');
  if(input) input.click();
}
async function uploadAssetFiles(files, {attachTo=null,quiet=false} = {}){
  const list = (files || []).filter(Boolean);
  if(!list.length) return [];
  const created = [];
  if(!quiet) toast(`Uploading ${list.length} file${list.length===1?'':'s'}...`);
  for(const file of list){
    try{
      const remoteUpload = Boolean(supabase && SESSION?.user && WORKSPACE_ID);
      validateAssetFile(file, remoteUpload);
      const id = uid('asset');
      const cleanName = safeFileName(file.name);
      const baseAsset = {
        id,
        title:file.name || cleanName,
        kind:assetKindForFile(file),
        mimeType:file.type || '',
        size:file.size || 0,
        source:'upload',
        storageBucket:ASSET_BUCKET,
        storagePath:'',
        externalUrl:'',
        description:'',
        createdBy:currentMemberId(),
        updated:Date.now(),
        data:{}
      };
      if(remoteUpload){
        const path = `${WORKSPACE_ID}/${id}/${cleanName}`;
        const { error } = await supabase.storage.from(ASSET_BUCKET).upload(path, file, {
          contentType:file.type || undefined,
          cacheControl:'3600',
          metadata:{
            workspaceId:WORKSPACE_ID,
            assetId:id,
            uploadedBy:SESSION.user.id,
            originalName:file.name || cleanName
          },
          upsert:false
        });
        if(error) throw error;
        baseAsset.storagePath = path;
      } else {
        baseAsset.source = 'local';
        baseAsset.data.localDataUrl = await readFileAsDataUrl(file);
      }
      DATA.assets.unshift(baseAsset);
      created.push(baseAsset);
    }catch(error){
      toast(error?.message || `Could not upload ${file.name || 'file'}`);
    }
  }
  if(created.length && attachTo){
    const t=findTaskForWrite(attachTo.gid, attachTo.tid, attachTo.boardId || '');
    if(t){
      if(!t.custom) t.custom = {};
      t.custom[attachTo.cid] = [...new Set([...taskAssetIdsForColumn(t,attachTo.cid), ...created.map(a=>a.id)])];
    }
  }
  if(created.length){
    save();
    if(!quiet){
      closeDD();
      if(TASK_DETAIL) render();
      else { renderContent(); wireContent(); }
      toast(`${created.length} file${created.length===1?'':'s'} added to Docs Hub`);
    }
  }
  return created;
}
function readFileAsDataUrl(file){
  return new Promise((resolve,reject)=>{
    const reader = new FileReader();
    reader.onload=()=>resolve(String(reader.result || ''));
    reader.onerror=()=>reject(reader.error || new Error('Could not read file'));
    reader.readAsDataURL(file);
  });
}
function createLinkAsset(){
  modal('Add external file link','Use this for Google Drive, Dropbox, Figma, or any file that already lives somewhere else.',
    `<div class="field"><label>Display name</label><input id="asset-link-title" placeholder="e.g. Q3 creative folder"></div>
     <div class="field"><label>URL</label><input id="asset-link-url" placeholder="https://"></div>
     <div class="field"><label>Description</label><textarea id="asset-link-desc" rows="3" placeholder="Optional context for the team"></textarea></div>`,
    ()=>{
      const url=normalizedUrlHref($('#asset-link-url').value);
      if(!url){ toast('Enter a valid link'); return; }
      const duplicate = findDuplicateAssetUrl(url);
      if(duplicate){ toast(duplicateAssetMessage(duplicate, url)); return; }
      const canonicalUrl = canonicalUrlKey(url);
      DATA.assets.unshift({
        id:uid('asset'),
        title:$('#asset-link-title').value.trim() || url,
        kind:'link',
        mimeType:'',
        size:0,
        source:'link',
        storageBucket:'',
        storagePath:'',
        externalUrl:url,
        description:$('#asset-link-desc').value.trim(),
        createdBy:currentMemberId(),
        updated:Date.now(),
        data:{canonicalUrl}
      });
      save(); closeModal(); render(); toast('Link added to Docs Hub');
    }, 'Add link');
}
async function assetAccessUrl(a){
  if(!a) return '';
  if(a.externalUrl) return a.externalUrl;
  if(a.data?.localDataUrl) return a.data.localDataUrl;
  if(a.storagePath && supabase){
    const { data, error } = await supabase.storage
      .from(a.storageBucket || ASSET_BUCKET)
      .createSignedUrl(a.storagePath, 60 * 60);
    if(error) throw error;
    return data?.signedUrl || '';
  }
  return '';
}
async function openAsset(id){
  const a=asset(id);
  if(!a)return;
  try{
    const url = await assetAccessUrl(a);
    if(!url){ toast('No file URL available'); return; }
    const preview = assetPreviewHtml(a,url);
    const root=$('#modal-root');
    root.innerHTML=`<div class="overlay"><div class="modal modal-wide asset-preview-modal">
      <div class="asset-preview-head">
        <div><h3>${esc(a.title)}</h3><div class="msub">${esc([assetIcon(a), assetSizeLabel(a.size), a.description].filter(Boolean).join(' - '))}</div></div>
        <button class="icon-btn" data-asset-close title="Close">${I.close}</button>
      </div>
      ${preview}
      <div class="modal-actions"><button class="btn ghost" data-asset-close>Close</button><button class="btn" data-asset-open>Open in new tab</button></div>
    </div></div>`;
    root.querySelectorAll('[data-asset-close]').forEach(btn=>btn.onclick=closeModal);
    root.querySelector('[data-asset-open]').onclick=()=>window.open(url,'_blank','noopener');
    const ov=root.querySelector('.overlay');
    ov.onclick=e=>{ if(e.target===ov) closeModal(); };
  }catch(error){
    toast(error?.message || 'Could not open file');
  }
}
function assetPreviewHtml(a,url){
  if(a.kind === 'image') return `<div class="asset-preview-body"><img src="${esc(url)}" alt="${esc(a.title)}"></div>`;
  if(a.kind === 'video') return `<div class="asset-preview-body"><video src="${esc(url)}" controls></video></div>`;
  if(a.kind === 'audio') return `<div class="asset-preview-body compact"><audio src="${esc(url)}" controls></audio></div>`;
  if(a.kind === 'pdf') return `<div class="asset-preview-body"><iframe src="${esc(url)}" title="${esc(a.title)}"></iframe></div>`;
  return `<div class="asset-preview-body compact"><div class="asset-preview-file"><span>${esc(assetIcon(a))}</span><strong>${esc(a.title)}</strong><p>Preview is not available for this file type. Open it in a new tab to view or download.</p></div></div>`;
}
function newDocFromTemplate(i){
  const tmpl = i>=0 ? DOC_TEMPLATES[i] : {name:'', body:''};
  const d={id:uid('doc'),title:tmpl.name?tmpl.name+' — '+new Date().toLocaleDateString():'Untitled doc',
    template:tmpl.name||'', content:tmpl.body, updated:Date.now()};
  DATA.docs.unshift(d); save(); openDoc(d.id);
}
function openDocLegacy(id){
  const d=DATA.docs.find(x=>x.id===id); if(!d)return;
  modal('Edit doc','Simple rich-text-free notes. Saves automatically when you click Save.',
    `<div class="field"><label>Title</label><input id="doc-title" value="${esc(d.title)}"></div>
     <div class="field"><label>Content</label><textarea id="doc-content" rows="12" style="width:100%;padding:10px;border:1px solid var(--line2);border-radius:6px;font-family:ui-monospace,monospace;font-size:13px;resize:vertical">${esc(d.content)}</textarea></div>`,
    ()=>{ d.title=$('#doc-title').value.trim()||'Untitled'; d.content=$('#doc-content').value; d.updated=Date.now(); save(); closeModal(); render(); toast('Doc saved'); },
    'Save doc');
}

function docContentToHtml(content){
  const raw = String(content || '').trim();
  if(!raw) return '<p><br></p>';
  if(/<\/?[a-z][\s\S]*>/i.test(raw)) return sanitizeDocHtml(raw);
  const lines = raw.replace(/\r\n/g,'\n').split('\n');
  const html = [];
  let listOpen = false;
  lines.forEach(line=>{
    const trimmed = line.trim();
    if(!trimmed){
      if(listOpen){ html.push('</ul>'); listOpen=false; }
      return;
    }
    const bullet = trimmed.match(/^[-*]\s+(.+)$/);
    if(bullet){
      if(!listOpen){ html.push('<ul>'); listOpen=true; }
      html.push(`<li>${esc(bullet[1])}</li>`);
      return;
    }
    if(listOpen){ html.push('</ul>'); listOpen=false; }
    if(trimmed.startsWith('### ')) html.push(`<h3>${esc(trimmed.slice(4))}</h3>`);
    else if(trimmed.startsWith('## ')) html.push(`<h2>${esc(trimmed.slice(3))}</h2>`);
    else if(trimmed.startsWith('# ')) html.push(`<h1>${esc(trimmed.slice(2))}</h1>`);
    else html.push(`<p>${esc(trimmed)}</p>`);
  });
  if(listOpen) html.push('</ul>');
  return sanitizeDocHtml(html.join('') || '<p><br></p>');
}

function sanitizeDocHtml(html){
  const template = document.createElement('template');
  template.innerHTML = String(html || '');
  const allowed = new Set(['P','DIV','BR','H1','H2','H3','STRONG','B','EM','I','U','UL','OL','LI','BLOCKQUOTE','A','SPAN','FONT']);
  const cleanNode = node=>{
    [...node.childNodes].forEach(child=>{
      if(child.nodeType === Node.TEXT_NODE) return;
      if(child.nodeType !== Node.ELEMENT_NODE){ child.remove(); return; }
      if(!allowed.has(child.tagName)){
        cleanNode(child);
        child.replaceWith(...child.childNodes);
        return;
      }
      [...child.attributes].forEach(attr=>{
        const name = attr.name.toLowerCase();
        if(child.tagName === 'A' && name === 'href' && /^(https?:|mailto:)/i.test(attr.value)) return;
        if(child.tagName === 'FONT' && name === 'size') return;
        child.removeAttribute(attr.name);
      });
      if(child.tagName === 'A'){
        child.setAttribute('target','_blank');
        child.setAttribute('rel','noopener noreferrer');
      }
      cleanNode(child);
    });
  };
  cleanNode(template.content);
  return template.innerHTML || '<p><br></p>';
}

function saveDocFromEditor(id){
  const d = doc(id);
  const editor = $('#doc-editor');
  if(!d || !editor) return;
  d.title = $('#doc-title')?.value.trim() || 'Untitled doc';
  d.content = sanitizeDocHtml(editor.innerHTML);
  d.updated = Date.now();
  save();
  closeModal();
  render();
  toast('Doc saved');
}

function formatDocEditor(command, value=null){
  const editor = $('#doc-editor');
  if(!editor) return;
  editor.focus();
  document.execCommand(command, false, value);
}

function wireDocEditor(id){
  const root = $('#modal-root');
  root.querySelectorAll('[data-doc-format]').forEach(btn=>btn.onclick=()=>formatDocEditor(btn.dataset.docFormat));
  root.querySelectorAll('[data-doc-block]').forEach(btn=>btn.onclick=()=>formatDocEditor('formatBlock', btn.dataset.docBlock));
  root.querySelector('[data-doc-size]')?.addEventListener('change',e=>formatDocEditor('fontSize', e.target.value));
  root.querySelector('[data-doc-clear]')?.addEventListener('click',()=>formatDocEditor('removeFormat'));
  root.querySelector('[data-doc-save]')?.addEventListener('click',()=>saveDocFromEditor(id));
  root.querySelectorAll('[data-doc-close]').forEach(btn=>btn.onclick=closeModal);
  const editor = $('#doc-editor');
  if(editor) setTimeout(()=>editor.focus(),30);
}

function openDoc(id){
  const d=DATA.docs.find(x=>x.id===id); if(!d)return;
  const root=$('#modal-root');
  root.innerHTML=`<div class="overlay"><div class="modal modal-doc-editor" data-doc-editor-id="${esc(id)}">
    <div class="doc-editor-head">
      <div>
        <input id="doc-title" class="doc-title-input" value="${esc(d.title)}" aria-label="Document title">
        <div class="msub">${esc(d.template || 'TaskTrack document')} - Last updated ${new Date(d.updated || Date.now()).toLocaleDateString()}</div>
      </div>
      <button class="icon-btn" data-doc-close title="Close">${I.close}</button>
    </div>
    <div class="doc-toolbar" aria-label="Document formatting">
      <button type="button" data-doc-block="p">Paragraph</button>
      <button type="button" data-doc-block="h1">H1</button>
      <button type="button" data-doc-block="h2">H2</button>
      <button type="button" data-doc-block="h3">H3</button>
      <span></span>
      <select data-doc-size title="Font size">
        <option value="3">Normal</option>
        <option value="2">Small</option>
        <option value="4">Large</option>
        <option value="5">Title</option>
      </select>
      <button type="button" data-doc-format="bold"><b>B</b></button>
      <button type="button" data-doc-format="italic"><i>I</i></button>
      <button type="button" data-doc-format="underline"><u>U</u></button>
      <button type="button" data-doc-format="insertUnorderedList">Bullets</button>
      <button type="button" data-doc-format="insertOrderedList">Numbers</button>
      <button type="button" data-doc-clear>Clear</button>
    </div>
    <div class="doc-page-wrap">
      <div id="doc-editor" class="doc-page" contenteditable="true" spellcheck="true">${docContentToHtml(d.content)}</div>
    </div>
    <div class="modal-actions"><button class="btn ghost" data-doc-close>Cancel</button><button class="btn" data-doc-save>Save doc</button></div>
  </div></div>`;
  const ov=root.querySelector('.overlay');
  ov.onclick=e=>{ if(e.target===ov) closeModal(); };
  wireDocEditor(id);
}
function delDoc(id){
  DATA.docs=DATA.docs.filter(x=>x.id!==id);
  DATA.boards.forEach(b=>(b.groups||[]).forEach(g=>taskTreeList(g.tasks).forEach(t=>{
    (b.columns||[]).filter(c=>c.type==='document').forEach(c=>{
      if(taskDocIdForColumn(t,c.id) === id){ if(!t.custom) t.custom = {}; t.custom[c.id] = ''; }
    });
  })));
  save(); render(); toast('Doc deleted');
}
function delAsset(id){
  const a=asset(id); if(!a)return;
  confirmModal(`Delete "${a.title}"?`,'This removes the asset from Docs Hub and from any task file cells. Uploaded storage files are also removed when possible.',async ()=>{
    try{
      DATA.boards.forEach(b=>(b.groups||[]).forEach(g=>taskTreeList(g.tasks).forEach(t=>{
        (b.columns||[]).filter(c=>c.type==='files').forEach(c=>{
          const next = taskAssetIdsForColumn(t,c.id).filter(assetId=>assetId!==id);
          if((t.custom||{})[c.id]) t.custom[c.id] = next;
        });
      })));
      DATA.assets = (DATA.assets || []).filter(x=>x.id!==id);
      if(a.storagePath && supabase){
        await supabase.storage.from(a.storageBucket || ASSET_BUCKET).remove([a.storagePath]);
      }
      save(); closeModal(); render(); toast('Asset deleted');
    }catch(error){
      toast(error?.message || 'Could not delete asset');
    }
  });
}

/* ============================================================
   ACTIONS - integrations and imports
   ============================================================ */
async function startIntegrationOauth(provider){
INTEGRATION_CONNECTIONS=INTEGRATION_CONNECTIONS.filter(c=>c.provider!==provider).concat({provider,external_account_id:'fictional-local-account',account_label:'Simulation · alex.rivera@example.invalid',scopes:[],metadata:{simulation:true},connected_at:new Date().toISOString()});
  localStorage.setItem('portfolio_integrations',JSON.stringify(INTEGRATION_CONNECTIONS));closeModal();render();toast('Simulation connected. No external account was contacted.');return;

  if(!supabase || !SESSION?.user){
    modal('Supabase sign-in required','Sign in before connecting external apps.','',closeModal,'OK');
    return;
  }
  const button = document.querySelector(`[data-oauth-provider="${provider}"]`);
  if(button){ button.disabled = true; button.textContent = 'Opening...'; }
  const functionName = provider === 'notion' ? 'notion-oauth-start' : 'google-oauth-start';
  const redirectTo = `${window.location.origin}${window.location.pathname}`;
  const body = provider === 'notion' ? { redirectTo } : { provider, redirectTo };
  const { data, error } = await supabase.functions.invoke(functionName, { body });
  if(error || !data?.authorizationUrl){
    if(button){ button.disabled = false; button.textContent = integrationConnection(provider) ? 'Switch account' : 'Connect account'; }
    showIntegrationSetupModal(provider, error?.message || data?.error || 'The connection function is not deployed yet.');
    return;
  }
  window.location.href = data.authorizationUrl;
}

async function disconnectIntegration(provider){
confirmModal('Disconnect simulated account?','Only this local sample connection will be removed.',()=>{INTEGRATION_CONNECTIONS=INTEGRATION_CONNECTIONS.filter(c=>c.provider!==provider);localStorage.setItem('portfolio_integrations',JSON.stringify(INTEGRATION_CONNECTIONS));closeModal();render();});return;

  if(!supabase || !SESSION?.user){
    modal('Supabase sign-in required','Sign in before disconnecting external apps.','',closeModal,'OK');
    return;
  }
  const label = integrationLabel(provider);
  confirmModal(`Disconnect ${label}?`,'TaskTrack will stop using this external account until someone connects it again.',async ()=>{
    await disconnectIntegrationConfirmed(provider, label);
  });
}

async function disconnectIntegrationConfirmed(provider,label=integrationLabel(provider)){
  if(provider === 'google_meet' && googleMeetImportIsEnabled()){
    const {data:disableData,error:disableError}=await supabase.functions.invoke('google-meet-subscribe',{
      body:{workspaceId:WORKSPACE_ID,action:'disable'}
    });
    if(disableError || disableData?.error){
      toast('Stop automatic Meet imports before disconnecting this account');
      return;
    }
  }
  const { error } = await supabase
    .from('integration_connections')
    .delete()
    .eq('user_id', SESSION.user.id)
    .eq('provider', provider);
  if(error){
    toast(`Could not disconnect ${label}`);
    return;
  }
  closeModal();
  await refreshIntegrationConnections({ rerender:true });
  toast(`${label} disconnected`);
}

function showIntegrationSetupModal(provider,message){
  const label = integrationLabel(provider);
  modal(`${label} is not ready yet`, message || 'This connection is not available yet.',
    `<div class="integration-setup-list">
      <div>Ask a workspace admin to finish the secure connection setup for ${esc(label)}.</div>
      <div>No tasks or files were changed.</div>
    </div>`,
    closeModal,
    'OK');
}

function openNotionAiImport(){
  if(false){
    modal('Supabase sign-in required','Sign in before importing from Notion.','',closeModal,'OK');
    return;
  }
  if(!integrationConnection('notion')){
    modal('Connect Notion first','The AI import assistant needs access to the Notion workspace you want to import from.',
      '<div class="integration-note">After Notion is connected, come back here and describe the exact tasks or database you want to import.</div>',
      ()=>startIntegrationOauth('notion'),
      'Connect Notion');
    return;
  }

  modal('Notion AI Import Assistant','Describe what you want to import. Nothing will be added until you review the preview.',
    `<div class="field">
      <label>What should TaskTrack import from Notion?</label>
      <textarea id="notion-ai-prompt" class="ai-prompt" rows="6" placeholder="Example: Import the active content calendar tasks from Notion, group them by status, keep owner, due date, priority, channel, and notes. Skip completed items."></textarea>
    </div>
    <div id="notion-ai-error" class="integration-note hidden"></div>
    <div class="integration-note">The assistant can only read Notion pages and databases selected when you connected Notion.</div>`,
    previewNotionAiImport,
    'Create preview');
  const dialog = document.querySelector('#modal-root .modal');
  if(dialog) dialog.classList.add('modal-wide');
}

async function previewNotionAiImport(){
const samplePrompt=$('#notion-ai-prompt')?.value.trim();if(!samplePrompt)return;
  const samplePreview={boardName:'Simulated imported work',customColumns:[],groups:[{name:'For review',tasks:[{name:'Review fictional launch checklist',status:'not_started',priority:'medium',notes:'Local simulation; no Notion workspace was read.'},{name:'Assign an owner after reviewing the brief',status:'not_started',priority:'low'}]}],skipped:[]};
  NOTION_AI_PREVIEW=samplePreview;showNotionAiPreview(samplePreview,samplePrompt);return;

  const prompt = $('#notion-ai-prompt')?.value.trim();
  const err = $('#notion-ai-error');
  const ok = $('#modal-root [data-mc="ok"]');
  if(!prompt){
    if(err){ err.textContent='Write a short prompt first.'; err.classList.remove('hidden'); }
    return;
  }
  if(ok){ ok.disabled = true; ok.textContent = 'Reading Notion...'; }
  if(err) err.classList.add('hidden');

  const { data, error } = await supabase.functions.invoke('notion-ai-import', {
    body: { prompt, limit: 120 },
  });

  if(error || data?.error || !data?.preview){
    if(ok){ ok.disabled = false; ok.textContent = 'Create preview'; }
    const message = error?.message || data?.error || 'Could not build a Notion import preview.';
    if(err){ err.textContent = message; err.classList.remove('hidden'); }
    return;
  }

  NOTION_AI_PREVIEW = data.preview;
  showNotionAiPreview(data.preview, prompt);
}

function showNotionAiPreview(preview,prompt){
  const groups = Array.isArray(preview?.groups) ? preview.groups : [];
  const taskCount = groups.reduce((sum,g)=>sum+(Array.isArray(g.tasks)?g.tasks.length:0),0);
  const customCount = Array.isArray(preview?.customColumns) ? preview.customColumns.length : 0;
  const groupHtml = groups.map(g=>`
    <div class="ai-preview-group">
      <div><strong>${esc(g.name || 'Imported tasks')}</strong><span>${Array.isArray(g.tasks)?g.tasks.length:0} tasks</span></div>
      <ul>${(Array.isArray(g.tasks)?g.tasks:[]).slice(0,5).map(t=>`<li>${esc(t.name || 'Untitled task')}</li>`).join('')}</ul>
    </div>`).join('');
  const skipped = Array.isArray(preview?.skipped) && preview.skipped.length
    ? `<div class="ai-preview-skip"><strong>Skipped</strong>${preview.skipped.slice(0,4).map(x=>`<div>${esc(x.source || 'Notion row')}: ${esc(x.reason || 'Skipped')}</div>`).join('')}</div>`
    : '';

  modal('Review Notion AI Preview', `${taskCount} task${taskCount===1?'':'s'} ready to import.`,
    `<div class="ai-preview-summary">
      <div><b>Board</b><span>${esc(preview?.boardName || 'Notion Import')}</span></div>
      <div><b>Custom columns</b><span>${customCount}</span></div>
      <div><b>Request</b><span>${esc(prompt)}</span></div>
    </div>
    ${preview?.summary ? `<div class="integration-note">${esc(preview.summary)}</div>` : ''}
    <div class="ai-preview-list">${groupHtml || '<div class="empty-state">No tasks matched that request.</div>'}</div>
    ${skipped}`,
    ()=>importNotionAiPreview(NOTION_AI_PREVIEW),
    'Import tasks');
  const dialog = document.querySelector('#modal-root .modal');
  if(dialog) dialog.classList.add('modal-wide');
}

function importNotionAiPreview(preview){
  const groups = Array.isArray(preview?.groups) ? preview.groups : [];
  const taskCount = groups.reduce((sum,g)=>sum+(Array.isArray(g.tasks)?g.tasks.length:0),0);
  if(!taskCount){ toast('No tasks to import'); return; }

  const b={id:uid('b'),name:String(preview.boardName||'Notion Import').trim()||'Notion Import',icon:'N',columns:[],columnOrder:BASE_COLUMNS.map(c=>c.id),groups:[]};
  const customDefs = dedupeImportColumns(preview);
  const customIds = {};
  customDefs.forEach(col=>{
    const id = uid('c');
    b.columns.push({ id, name:col.name, type:col.type });
    b.columnOrder.push(id);
    customIds[normImportHeader(col.name)] = id;
  });

  groups.forEach(group=>{
    const g={id:uid('g'),name:String(group.name||'Imported from Notion').trim()||'Imported from Notion',color:GROUP_COLORS[b.groups.length % GROUP_COLORS.length],open:true,tasks:[]};
    (Array.isArray(group.tasks)?group.tasks:[]).forEach(row=>{
      const t = notionAiTaskToTask(row, customIds);
      if(t) g.tasks.push(t);
    });
    if(g.tasks.length) b.groups.push(g);
  });
  if(!b.groups.length){ toast('No tasks to import'); return; }

  DATA.boards.push(b);
  DATA.integrations.unshift({id:uid('int'),provider:'notion_ai',name:`Notion AI: ${b.name}`,connectedAt:Date.now(),metadata:{rows:taskCount,boardId:b.id}});
  recordUpdate({boardId:b.id,groupId:b.groups[0]?.id||'',body:`Imported ${taskCount} task${taskCount===1?'':'s'} from Notion with AI Import.`});
  normalizeData();
  save();
  closeModal();
  VIEW='board';
  CURRENT_BOARD=b.id;
  render();
  toast(`${taskCount} task${taskCount===1?'':'s'} imported`);
}

function dedupeImportColumns(preview){
  const columns = new Map();
  const reserved = new Set(['task','name','owner','assignedto','status','duedate','due','timeline','priority','budget','notes']);
  (Array.isArray(preview?.customColumns)?preview.customColumns:[]).forEach(col=>{
    const name = String(col?.name||'').trim();
    if(!name || reserved.has(normImportHeader(name))) return;
    columns.set(normImportHeader(name), { name, type:safeImportColumnType(col?.type) });
  });
  (Array.isArray(preview?.groups)?preview.groups:[]).forEach(g=>{
    (Array.isArray(g.tasks)?g.tasks:[]).forEach(t=>{
      (Array.isArray(t.customValues)?t.customValues:[]).forEach(cv=>{
        const name=String(cv?.name||'').trim();
        if(!name || reserved.has(normImportHeader(name))) return;
        const key=normImportHeader(name);
        if(!columns.has(key)) columns.set(key, { name, type:'text' });
      });
    });
  });
  return Array.from(columns.values()).slice(0,12);
}

function safeImportColumnType(type){
  return ['text','number','date','status'].includes(type) ? type : 'text';
}

function notionAiTaskToTask(row, customIds){
  const name = String(row?.name||'').trim();
  if(!name) return null;
  const t={id:uid('t'),name,owner:findImportOwner(row.ownerText),status:safeImportStatus(row.status),
    priority:safeImportPriority(row.priority),due:parseImportDate(row.due),tlStart:parseImportDate(row.tlStart),
    tlEnd:parseImportDate(row.tlEnd),budget:Number(row.budget)||0,notes:String(row.notes||''),assignedBy:currentMemberId(),custom:{},subtasks:[]};
  (Array.isArray(row.customValues)?row.customValues:[]).forEach(cv=>{
    const id = customIds[normImportHeader(cv?.name)];
    if(id) t.custom[id]=String(cv?.value||'');
  });
  return t;
}

function safeImportStatus(value){
  const key=String(value||'').trim();
  return STATUS[key] ? key : mapImportStatus(key);
}

function safeImportPriority(value){
  const key=String(value||'').trim();
  return PRIORITY[key] ? key : mapImportPriority(key);
}

function readNotionCsvFile(e){
  const f=e.target.files[0]; if(!f)return;
  const r=new FileReader();
  r.onload=()=>{
    try{
      const rows=parseCsv(String(r.result||''));
      if(!rows.length) throw new Error('empty');
      showNotionImportModal(f.name, rows);
    }catch(err){
      toast('Could not read that CSV');
    }finally{
      e.target.value='';
    }
  };
  r.readAsText(f);
}

function showNotionImportModal(fileName, rows){
  const cleanName = fileName.replace(/\.[^.]+$/,'').trim() || 'Notion import';
  const boardOpts = DATA.boards.map(b=>`<option value="${b.id}">${esc(b.name)}</option>`).join('');
  modal('Import Notion CSV', `${rows.length} row${rows.length===1?'':'s'} found in ${fileName}.`,
    `<div class="field"><label>Import destination</label><select id="m-import-mode">
      <option value="new">Create a new board</option>
      ${DATA.boards.length ? '<option value="existing">Add to an existing board</option>' : ''}
    </select></div>
    <div class="field" id="m-import-new-wrap"><label>New board name</label><input id="m-import-board-name" value="${esc(cleanName)}"></div>
    <div id="m-import-existing-wrap" style="display:none">
      <div class="field"><label>Board</label><select id="m-import-board">${boardOpts}</select></div>
      <div class="field"><label>Group</label><select id="m-import-group"></select></div>
    </div>
    <div class="integration-note">Mapped fields: task name, owner, status, due date, timeline, priority, budget, and notes. Extra Notion properties become custom columns.</div>`,
    ()=>{
      const mode=$('#m-import-mode').value;
      const opts = mode==='new'
        ? { mode, boardName:$('#m-import-board-name').value.trim() || cleanName }
        : { mode, boardId:$('#m-import-board').value, groupId:$('#m-import-group').value };
      importNotionRows(rows, fileName, opts);
    },
    'Import tasks');
  const mode=$('#m-import-mode');
  const boardSelect=$('#m-import-board');
  const groupSelect=$('#m-import-group');
  const updateGroupOptions=()=>{
    const b=board(boardSelect?.value);
    if(groupSelect) groupSelect.innerHTML=(b?.groups||[]).map(g=>`<option value="${g.id}">${esc(g.name)}</option>`).join('');
  };
  const updateMode=()=>{
    const existing = mode.value === 'existing';
    $('#m-import-new-wrap').style.display = existing ? 'none' : '';
    $('#m-import-existing-wrap').style.display = existing ? '' : 'none';
    if(existing) updateGroupOptions();
  };
  if(boardSelect) boardSelect.onchange=updateGroupOptions;
  mode.onchange=updateMode;
  updateMode();
}

function importNotionRows(rows,fileName,opts){
  const headers = Object.keys(rows[0]||{});
  const map = buildNotionHeaderMap(headers);
  const mappedHeaders = new Set(Object.values(map).filter(Boolean));
  let b, g;
  if(opts.mode==='new'){
    b={id:uid('b'),name:opts.boardName,icon:'N',columns:[],columnOrder:BASE_COLUMNS.map(c=>c.id),groups:[
      {id:uid('g'),name:'Imported from Notion',color:'#0073ea',open:true,tasks:[]}
    ]};
    DATA.boards.push(b);
    g=b.groups[0];
  } else {
    b=board(opts.boardId);
    g=b?.groups.find(x=>x.id===opts.groupId);
    if(!b || !g){ toast('Choose a board and group'); return; }
  }
  const extraHeaders = headers.filter(h=>!mappedHeaders.has(h) && rows.some(r=>String(r[h]||'').trim()));
  const customIds = ensureImportColumns(b, rows, extraHeaders);
  const tasks = rows.map(r=>notionRowToTask(r, map, customIds)).filter(Boolean);
  g.tasks.push(...tasks);
  DATA.integrations.unshift({id:uid('int'),provider:'notion_csv',name:`Notion CSV: ${fileName}`,connectedAt:Date.now(),metadata:{rows:tasks.length,boardId:b.id}});
  recordUpdate({boardId:b.id,groupId:g.id,body:`Imported ${tasks.length} task${tasks.length===1?'':'s'} from Notion CSV "${fileName}".`});
  normalizeData();
  save();
  closeModal();
  VIEW='board';
  CURRENT_BOARD=b.id;
  render();
  toast(`${tasks.length} task${tasks.length===1?'':'s'} imported`);
}

function parseCsv(text){
  const rows=[]; let row=[], cur='', inQuotes=false;
  for(let i=0;i<text.length;i++){
    const ch=text[i], next=text[i+1];
    if(ch==='"' && inQuotes && next==='"'){ cur+='"'; i++; continue; }
    if(ch==='"'){ inQuotes=!inQuotes; continue; }
    if(ch===',' && !inQuotes){ row.push(cur); cur=''; continue; }
    if((ch==='\n' || ch==='\r') && !inQuotes){
      if(ch==='\r' && next==='\n') i++;
      row.push(cur); cur='';
      if(row.some(v=>String(v).trim())) rows.push(row);
      row=[]; continue;
    }
    cur+=ch;
  }
  row.push(cur);
  if(row.some(v=>String(v).trim())) rows.push(row);
  if(rows.length<2) return [];
  const headers=rows[0].map((h,i)=>String(h||`Column ${i+1}`).trim() || `Column ${i+1}`);
  return rows.slice(1).map(values=>{
    const obj={};
    headers.forEach((h,i)=>{ obj[h]=String(values[i] ?? '').trim(); });
    return obj;
  }).filter(obj=>Object.values(obj).some(v=>String(v).trim()));
}

function buildNotionHeaderMap(headers){
  const aliases = {
    name:['task','name','title','item','campaign'],
    owner:['owner','assignee','assigned to','person','people'],
    status:['status','state','phase'],
    due:['due','due date','deadline','date'],
    tlStart:['timeline start','start date','start'],
    tlEnd:['timeline end','end date','end'],
    priority:['priority'],
    budget:['budget','cost','spend'],
    notes:['notes','note','description','summary']
  };
  return Object.fromEntries(Object.entries(aliases).map(([key,list])=>[key,findImportHeader(headers,list)]));
}
function normImportHeader(s){ return String(s||'').toLowerCase().replace(/[^a-z0-9]/g,''); }
function findImportHeader(headers,aliases){
  const byNorm = new Map(headers.map(h=>[normImportHeader(h),h]));
  for(const a of aliases){ const hit=byNorm.get(normImportHeader(a)); if(hit) return hit; }
  for(const h of headers){
    const n=normImportHeader(h);
    if(aliases.some(a=>n.includes(normImportHeader(a)))) return h;
  }
  return '';
}

function ensureImportColumns(b, rows, headers){
  const out={};
  headers.forEach(h=>{
    const existing=b.columns.find(c=>c.name.toLowerCase()===h.toLowerCase());
    if(existing){ out[h]=existing.id; return; }
    const type=inferImportColumnType(rows.map(r=>r[h]).filter(Boolean));
    const col={id:uid('c'),name:h,type};
    b.columns.push(col); b.columnOrder.push(col.id); out[h]=col.id;
  });
  return out;
}
function inferImportColumnType(values){
  const sample=values.slice(0,12).map(v=>String(v||'').trim()).filter(Boolean);
  if(sample.length && sample.every(v=>!Number.isNaN(Number(String(v).replace(/[$,%\s,]/g,''))))) return 'number';
  if(sample.length && sample.every(v=>Boolean(parseImportDate(v)))) return 'date';
  return 'text';
}

function notionRowToTask(row,map,customIds){
  const name = row[map.name] || row[Object.keys(row)[0]] || 'Imported task';
  const t={id:uid('t'),name,owner:findImportOwner(row[map.owner]),status:mapImportStatus(row[map.status]),
    priority:mapImportPriority(row[map.priority]),due:parseImportDate(row[map.due]),tlStart:parseImportDate(row[map.tlStart]),
    tlEnd:parseImportDate(row[map.tlEnd]),budget:parseImportNumber(row[map.budget]),notes:row[map.notes]||'',assignedBy:currentMemberId(),custom:{},subtasks:[]};
  Object.entries(customIds).forEach(([header,cid])=>{ t.custom[cid]=row[header]||''; });
  return t;
}
function parseImportNumber(v){ return Number(String(v||'').replace(/[$,%\s,]/g,'')) || 0; }
function parseImportDate(v){
  const s=String(v||'').trim(); if(!s) return '';
  const d = /^\d{4}-\d{2}-\d{2}$/.test(s) ? new Date(s+'T00:00:00') : new Date(s);
  return Number.isNaN(d.getTime()) ? '' : d.toISOString().slice(0,10);
}
function mapImportStatus(v){
  const s=String(v||'').toLowerCase();
  if(/done|complete|completed|shipped|published|approved/.test(s)) return 'done';
  if(/working|progress|active|doing|started/.test(s)) return 'working';
  if(/stuck|blocked|issue|problem/.test(s)) return 'stuck';
  if(/hold|paused|waiting/.test(s)) return 'on_hold';
  return 'not_started';
}
function mapImportPriority(v){
  const s=String(v||'').toLowerCase();
  if(/critical|urgent|p0/.test(s)) return 'critical';
  if(/high|p1/.test(s)) return 'high';
  if(/medium|med|p2/.test(s)) return 'medium';
  if(/low|p3/.test(s)) return 'low';
  return 'none';
}
function findImportOwner(v){
  const s=String(v||'').toLowerCase().trim(); if(!s) return '';
  return DATA.team.find(m=>m.name.toLowerCase()===s || s.includes(m.name.toLowerCase()))?.id || '';
}

/* ============================================================
   ACTIONS — settings / data
   ============================================================ */
async function switchWorkspace(workspaceId){
  if(!workspaceId || workspaceId === WORKSPACE_ID) return;
  await saveCloudNow({ silent:true });
  setSelectedWorkspaceId(workspaceId);
  VIEW = 'home';
  CURRENT_BOARD = null;
  await loadCloudWorkspace();
}

function createWorkspaceModal(){
  if(!supabase || !SESSION?.user){ toast('Sign in before creating a workspace'); return; }
  modal('Create separate workspace','Use this for another department that should have its own boards, tasks, docs, and members.',
    `<div class="field"><label>Workspace name</label><input id="m-workspace-name" value="TaskTrack" maxlength="80" autofocus></div>`,
    async ()=>{
      const ok = $('#modal-root [data-mc="ok"]');
      if(ok){ ok.disabled = true; ok.textContent = 'Creating...'; }
      try{
        const workspace = await createWorkspace($('#m-workspace-name')?.value || 'TaskTrack');
        closeModal();
        setSelectedWorkspaceId(workspace.id);
        await loadCloudWorkspace();
        toast('Workspace created');
      }catch(err){
        if(ok){ ok.disabled = false; ok.textContent = 'Create'; }
        toast(err.message || 'Could not create workspace');
      }
    }, 'Create');
}

function saveWorkspace(){
  const admin = isCurrentAdmin();
  const previousPet = assistantPetSettings(DATA.settings.assistantPet);
  const nextPetStyle = $('#set-assistant-style')?.value || 'retriever';
  let nextPetColor = $('#set-assistant-color')?.value || ASSISTANT_PET_STYLE_COLORS[nextPetStyle] || '#d99a35';
  const previousDefaultColor = ASSISTANT_PET_STYLE_COLORS[previousPet.style] || ASSISTANT_PET_DEFAULTS.color;
  if(nextPetStyle !== previousPet.style && nextPetColor.toLowerCase() === previousDefaultColor.toLowerCase()){
    nextPetColor = ASSISTANT_PET_STYLE_COLORS[nextPetStyle] || nextPetColor;
  }
  if(admin){
    DATA.settings.account=$('#set-account')?.value.trim()||'TaskTrack';
    DATA.settings.currency=$('#set-currency')?.value.trim()||'$';
    DATA.settings.language = LANGUAGE_OPTIONS[$('#set-language')?.value] ? $('#set-language').value : 'en';
  }
  const currentPrefs = notificationPrefs();
  const checkedOrCurrent = (id, fallback) => {
    const input = $(`#${id}`);
    return input ? input.checked : fallback;
  };
  const emailDueRemindersEnabled = checkedOrCurrent('notify-email', currentPrefs.emailDueReminders !== false && currentPrefs.email !== false);
  const dailyDigestEnabled = checkedOrCurrent('notify-daily-digest', currentPrefs.dailyDigest !== false && currentPrefs.email !== false);
  const nextNotificationPrefs = {
    mentions:checkedOrCurrent('notify-mentions', currentPrefs.mentions !== false),
    everyone:checkedOrCurrent('notify-everyone', currentPrefs.everyone !== false),
    assigned:checkedOrCurrent('notify-assigned', currentPrefs.assigned !== false),
    subscribed:checkedOrCurrent('notify-subscribed', currentPrefs.subscribed !== false),
    taskUpdates:checkedOrCurrent('notify-task-updates', currentPrefs.taskUpdates !== false),
    dueReminders:checkedOrCurrent('notify-due-reminders', currentPrefs.dueReminders !== false),
    authored:checkedOrCurrent('notify-authored', currentPrefs.authored !== false),
    desktop:checkedOrCurrent('notify-desktop', currentPrefs.desktop !== false),
    push:checkedOrCurrent('notify-push', currentPrefs.push !== false),
    email:emailDueRemindersEnabled || dailyDigestEnabled,
    emailDueReminders:emailDueRemindersEnabled,
    dailyDigest:dailyDigestEnabled
  };
  setNotificationPrefsForMember(currentMemberId(), nextNotificationPrefs);
  DATA.settings.notificationPrefs = nextNotificationPrefs;
  document.querySelectorAll('[data-board-notify-mode]').forEach(select=>{
    setBoardNotificationPrefs(select.dataset.boardNotifyMode, currentMemberId(), {mode:select.value});
  });
  if($('#notify-desktop')?.checked === false) setDesktopNotificationsEnabled(false);
  if($('#notify-desktop')?.checked === true && desktopNotificationsGranted()) setDesktopNotificationsEnabled(true);
  if($('#notify-desktop')?.checked === true && !desktopNotificationsGranted()){
    toast('Enable desktop notifications from the notification bell first');
  }
  if($('#notify-push')?.checked === false){
    void removeWebPushSubscription();
  } else if($('#notify-push')?.checked === true && desktopNotificationsGranted()){
    void registerWebPushSubscription().catch(error=>console.warn('Web push subscription failed', error));
  }
  const previousActions = assistantActionPrefs();
  DATA.settings.assistantActions = admin ? {
    gmailDrafts:$('#assist-act-gmail-drafts')?.checked !== false,
    gmailSend:$('#assist-act-gmail-send')?.checked === true,
    calendarEvents:$('#assist-act-calendar-events')?.checked !== false,
    driveSearch:$('#assist-act-drive-search')?.checked !== false,
    driveDocs:$('#assist-act-drive-docs')?.checked !== false,
    driveImport:$('#assist-act-drive-import')?.checked !== false,
    driveManage:$('#assist-act-drive-manage')?.checked === true,
    driveEditDocs:$('#assist-act-drive-edit-docs')?.checked === true,
    notionImport:$('#assist-act-notion-import')?.checked !== false
  } : previousActions;
  if(admin){
    DATA.settings.assistantPet = assistantPetSettings({
      enabled:$('#set-assistant-enabled')?.checked !== false,
      name:$('#set-assistant-name')?.value || 'TaskTrack',
      style:nextPetStyle,
      color:nextPetColor,
      size:$('#set-assistant-size')?.value || 'medium',
      position:$('#set-assistant-position')?.value || 'bottom-right',
      motion:$('#set-assistant-motion')?.value || 'float',
      autoBrief:false
    });
  }
  if(!DATA.settings.assistantPet.enabled) ASSISTANT_OPEN = false;
  if(admin && WORKSPACE_ID && supabase && SESSION?.user){
    void supabase.from('workspaces').update({ name:DATA.settings.account, updated_at:new Date().toISOString() }).eq('id', WORKSPACE_ID);
  }
  save(); render(); toast('Workspace saved');
}
async function addMember(){
  if(!isCurrentAdmin()){ toast('Only admins can add team members'); return; }
  const email=normalizeEmail($('#new-member-email')?.value || '');
  const rawName=$('#new-member-name').value.trim();
  const name=rawName || displayNameFromEmail(email);
  if(!name){ toast('Enter a name or email'); return; }
  if(email && DATA.team.some(m=>normalizeEmail(m.email)===email)){ toast('That email is already on the team'); return; }
  const memberRecord={id:uid('u'),name,email,initials:initialsFor(name),color:AVATAR_COLORS[DATA.team.length%AVATAR_COLORS.length],role:'member'};
  setMemberRole(memberRecord.id, 'member');
  DATA.team.push(memberRecord);
  let inviteOk = true;
  if(email && WORKSPACE_ID && supabase && SESSION?.user){
    const { error } = await supabase.from('workspace_members').upsert({
      workspace_id: WORKSPACE_ID,
      email,
      name,
      role:'member',
      updated_at: new Date().toISOString()
    }, { onConflict:'workspace_id,email' });
    inviteOk = !error;
  }
  save(); render(); toast(inviteOk ? 'Member added' : 'Member added, but workspace access needs the owner');
}
async function delMember(id){
  if(!isCurrentAdmin()){ toast('Only admins can remove team members'); return; }
  const removed = DATA.team.find(m=>m.id===id);
  if(!removed) return;
  if(memberRole(id)==='admin' && DATA.team.filter(m=>memberRole(m.id)==='admin').length <= 1){
    toast('Keep at least one admin');
    return;
  }
  confirmModal(`Remove ${removed.name}?`,'This removes the teammate from TaskTrack and unassigns their tasks in this workspace.',async ()=>{
    DATA.boards.forEach(b=>b.groups.forEach(g=>g.tasks.forEach(t=>unassignMemberFromTaskTree(t, id))));
    if(currentMemberId()===id) setCurrentMemberId(DATA.team.find(m=>m.id!==id)?.id||'');
    DATA.team=DATA.team.filter(m=>m.id!==id);
    if(DATA.settings.memberRoles) delete DATA.settings.memberRoles[id];
    save(); closeModal(); render(); toast('Member removed');
    if(removed?.email && WORKSPACE_ID && supabase){
      await supabase.from('workspace_members').delete().eq('workspace_id', WORKSPACE_ID).eq('email', normalizeEmail(removed.email));
    }
  });
}
async function changeMemberRole(id, role){
  if(!isCurrentAdmin()){ toast('Only admins can change roles'); renderContent(); return; }
  const m = member(id);
  if(!m) return;
  const nextRole = role === 'admin' ? 'admin' : 'member';
  if(memberRole(id)==='admin' && nextRole==='member' && DATA.team.filter(x=>memberRole(x.id)==='admin').length <= 1){
    toast('Keep at least one admin');
    renderContent();
    return;
  }
  const previousRole = memberRole(id);
  if(WORKSPACE_ID && supabase && SESSION?.user && normalizeEmail(m.email)){
    const { error } = await supabase
      .from('workspace_members')
      .update({ role:nextRole, updated_at:new Date().toISOString() })
      .eq('workspace_id', WORKSPACE_ID)
      .eq('email', normalizeEmail(m.email));
    if(error){
      toast(error.message || 'Could not update the real workspace role');
      renderContent();
      return;
    }
  }
  m.role = nextRole;
  setMemberRole(id, nextRole);
  if(id === currentMemberId()) WORKSPACE_ROLE = nextRole;
  save();
  renderContent();
  toast(previousRole === nextRole ? 'Role unchanged' : 'Role updated');
}
function exportData(){
  const blob=new Blob([JSON.stringify(DATA,null,2)],{type:'application/json'});
  const url=URL.createObjectURL(blob); const a=document.createElement('a');
  a.href=url; a.download=`tasktrack-data-${new Date().toISOString().slice(0,10)}.json`; a.click();
  URL.revokeObjectURL(url); toast('Data exported');
}
function importData(e){
  if(!isCurrentAdmin()){ toast('Only admins can import over the shared workspace'); if(e?.target) e.target.value = ''; return; }
  const f=e.target.files[0]; if(!f)return;
  const r=new FileReader();
  r.onload=()=>{ try{
      const obj=JSON.parse(r.result);
      if(!obj.boards) throw new Error('bad');
      confirmModal('Import TaskTrack data?','This replaces the current workspace data with the selected JSON file. Export a backup first if you need one.',()=>{
        DATA=obj; if(!DATA.automations) DATA.automations=[]; normalizeData(); save(); closeModal(); render(); toast('Data imported');
      });
    }catch(err){ toast('Invalid file'); } };
  r.readAsText(f);
  e.target.value = '';
}
function resetData(){
  if(!isCurrentAdmin()){ toast('Only admins can reset the shared workspace'); return; }
  confirmModal('Reset to sample data?','This permanently replaces all boards, tasks, docs, and settings in this workspace with demo content. Export first if you want a backup.',()=>{
    DATA=seed(); normalizeData(); save(); closeModal(); VIEW='home'; render(); toast('Reset complete');
  });
}

/* ============================================================
   ACCOUNT MENU
   ============================================================ */
function toggleAcct(){
  const ex=$('#acct-menu'); if(ex){ ex.remove(); return; }
  const me=member(currentMemberId())||DATA.team[0]||{name:'You',initials:'YO',color:'#0073ea'};
  const m=document.createElement('div'); m.id='acct-menu'; m.className='acct-menu';
  m.innerHTML=`
    <div class="acct-head"><div class="owner-av" style="background:${me.color}">${esc(me.initials)}</div>
      <div><div style="font-weight:700">${esc(me.name)}</div><div style="font-size:12px;color:var(--ink3)">${esc(DATA.settings.account||'TaskTrack')}</div></div></div>
    <div class="acct-grid">
      <div class="acct-item" data-go="settings">${I.gear} Workspace settings</div>
      <div class="acct-item" data-go="mywork">${I.work} My Work</div>
      <div class="acct-item" data-act="export">${I.download} Export data</div>
      ${isSupabaseConfigured && SESSION ? `<div class="acct-item" data-act="signout">${I.close} Sign out</div>` : ''}
    </div>`;
  document.body.appendChild(m);
  m.querySelectorAll('[data-go]').forEach(el=>el.onclick=()=>{ m.remove(); navigateToView(el.dataset.go); });
  const exp = m.querySelector('[data-act="export"]');
  if(exp) exp.onclick=()=>{ m.remove(); exportData(); };
  const out = m.querySelector('[data-act="signout"]');
  if(out) out.onclick=()=>{ m.remove(); void signOut(); };
}

/* ============================================================
   MODALS
   ============================================================ */
function modal(title,sub,bodyHtml,onConfirm,confirmLabel='Create'){
  const root=$('#modal-root');
  root.innerHTML=`<div class="overlay"><div class="modal"><h3>${esc(title)}</h3><div class="msub">${esc(sub)}</div>
    ${bodyHtml}<div class="modal-actions"><button class="btn ghost" data-mc="cancel">Cancel</button>
    <button class="btn" data-mc="ok">${esc(confirmLabel)}</button></div></div></div>`;
  const ov=root.querySelector('.overlay');
  ov.onclick=e=>{ if(e.target===ov) closeModal(); };
  root.querySelector('[data-mc="cancel"]').onclick=closeModal;
  root.querySelector('[data-mc="ok"]').onclick=onConfirm;
  const inp=root.querySelector('input,textarea'); if(inp) setTimeout(()=>inp.focus(),30);
}
function confirmModal(title,sub,onConfirm){
  modal(title,sub,'',onConfirm,'Confirm');
  const ok=$('#modal-root [data-mc="ok"]'); if(ok) ok.classList.add('danger');
}
function closeModal(){ $('#modal-root').innerHTML=''; }

/* ============================================================
   BOOT
   ============================================================ */
export function mountTaskManager(){
  void bootApp();
  return () => {
    closeDD();
    closeModal();
    stopWorkspaceSync();
    stopMeetingSync();
    stopRecurringRolloverTimer();
    clearTimeout(CLOUD_SAVE_TIMER);
    CLOUD_SAVE_TIMER = null;
    AUTH_SUBSCRIPTION?.unsubscribe?.();
  };
}

function recordPortfolioSimulation(action,payload){const key='portfolio_simulation_log';const rows=JSON.parse(localStorage.getItem(key)||'[]');rows.push({action,payload,simulation:true,at:new Date().toISOString()});localStorage.setItem(key,JSON.stringify(rows));}

