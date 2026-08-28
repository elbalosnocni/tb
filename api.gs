/**
 * COMPANY INTERNAL NOTICE PORTAL v2
 * Backend: Google Apps Script Web App + Google Sheets + Drive
 *
 * Deploy as Web App:
 *   Execute as: Me
 *   Who has access: your company/domain or anyone with the URL as appropriate.
 *
 * Then put the Web App /exec URL into index.html:
 *   const API_URL = 'https://script.google.com/macros/s/XXXX/exec';
 *
 * Run setupSystem() once from the Apps Script editor.
 */

const CONFIG = {
  appName: 'Cổng thông báo nội bộ',
  timezone: 'Asia/Ho_Chi_Minh',
  sessionSeconds: 8 * 60 * 60,
  maxAttachmentBytes: 25 * 1024 * 1024,
  sheets: {
    employees: 'Employees',
    announcements: 'Announcements',
    targets: 'AnnouncementTargets',
    reads: 'AnnouncementReads',
    imports: 'ImportHistory',
    logins: 'LoginHistory',
    audit: 'AuditLog'
  }
};

const HEADERS = {
  Employees: ['EmployeeID','FullName','Department','Position','Email','Phone','PasswordHash','PasswordSalt','Role','Status','CreatedAt','UpdatedAt','LastLoginAt'],
  Announcements: ['ID','Title','Content','Type','PublishDate','Author','Priority','Pinned','Status','Attachment','Keywords','CreatedAt','UpdatedAt','ScheduledAt'],
  AnnouncementTargets: ['AnnouncementID','TargetType','TargetValue'],
  AnnouncementReads: ['AnnouncementID','EmployeeID','ReadAt'],
  ImportHistory: ['ImportID','FileName','ImportedBy','ImportedAt','TotalRows','Added','Updated','Skipped','MarkedInactive','MissingAction','Status','Notes'],
  LoginHistory: ['LoginID','EmployeeID','LoginAt','Success','IP','UserAgent'],
  AuditLog: ['AuditID','At','ActorID','Action','Entity','EntityID','Details']
};

const TYPES = [
  '📢 Thông báo chung','🕐 Thời gian làm việc','💰 Lương & thưởng','🎁 Phúc lợi',
  '🏥 Sức khỏe','🦺 An toàn / HSE','🏭 Sản xuất','📋 Nhân sự',
  '🎉 Hoạt động công ty','⚠️ Khẩn cấp'
];

/* ---------- HTTP API ---------- */

function doGet(e) {
  return json_({
    ok: true,
    service: CONFIG.appName,
    version: '2.0',
    time: Utilities.formatDate(new Date(), CONFIG.timezone, "yyyy-MM-dd'T'HH:mm:ss"),
    message: 'API is running. Use POST for commands.'
  });
}

function doPost(e) {
  try {
    const body = parseBody_(e);
    const action = String(body.action || '').trim();
    const result = route_(action, body);
    return json_(result);
  } catch (err) {
    return json_({ok:false, error:String(err && err.message || err)});
  }
}

function route_(action, b) {
  switch(action) {
    case 'login': return apiLogin_(b);
    case 'logout': return apiLogout_(b);
    case 'me': return apiMe_(b);
    case 'dashboard': return apiDashboard_(b);
    case 'announcement': return apiAnnouncement_(b);
    case 'markRead': return apiMarkRead_(b);
    case 'readStats': return apiReadStats_(b);
    case 'adminDashboard': return apiAdminDashboard_(b);
    case 'saveAnnouncement': return apiSaveAnnouncement_(b);
    case 'deleteAnnouncement': return apiDeleteAnnouncement_(b);
    case 'togglePinned': return apiTogglePinned_(b);
    case 'changePassword': return apiChangePassword_(b);
    case 'previewImport': return apiPreviewImport_(b);
    case 'commitImport': return apiCommitImport_(b);
    case 'employees': return apiEmployees_(b);
    case 'uploadAttachment': return apiUploadAttachment_(b);
    default: throw new Error('Action không hợp lệ: ' + action);
  }
}

/* ---------- SETUP ---------- */

function setupSystem() {
  const ss = SpreadsheetApp.getActive();
  PropertiesService.getScriptProperties().setProperty('SPREADSHEET_ID', ss.getId());

  Object.keys(HEADERS).forEach(name => {
    let sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    const headers = HEADERS[name];
    sh.getRange(1,1,1,headers.length).setValues([headers]);
    sh.setFrozenRows(1);
    sh.getRange(1,1,1,headers.length).setFontWeight('bold');
  });

  seedAdmin_();
  installScheduleTrigger_();

  return 'Đã setup hệ thống. ADMIN001 / ChangeMe@123';
}

function seedAdmin_() {
  const sh = sheet_(CONFIG.sheets.employees);
  if (sh.getLastRow() >= 2) return;

  const salt = token_(16);
  sh.appendRow([
    'ADMIN001','System Admin','HR','Administrator','','',
    hash_('ChangeMe@123', salt),salt,'ADMIN','ACTIVE',new Date(),new Date(),''
  ]);
}

function installScheduleTrigger_() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'processScheduledAnnouncements') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('processScheduledAnnouncements').timeBased().everyMinutes(5).create();
}

function processScheduledAnnouncements() {
  const sh = sheet_(CONFIG.sheets.announcements);
  const rows = data_(sh);
  const ix = idx_(HEADERS.Announcements);
  const now = new Date();

  rows.forEach((r,i) => {
    const status = String(r[ix.Status] || '').toUpperCase();
    const scheduled = r[ix.ScheduledAt];
    if (status === 'SCHEDULED' && scheduled && new Date(scheduled) <= now) {
      sh.getRange(i+2, ix.Status+1).setValue('PUBLISHED');
      sh.getRange(i+2, ix.UpdatedAt+1).setValue(now);
    }
  });
}

/* ---------- AUTH ---------- */

function apiLogin_(b) {
  const employeeId = String(b.employeeId || '').trim().toUpperCase();
  const password = String(b.password || '');
  if (!employeeId || !password) throw new Error('Vui lòng nhập mã nhân viên và mật khẩu.');

  const emp = findEmployee_(employeeId);
  const ok = !!emp && emp.Status === 'ACTIVE' &&
    hash_(password, emp.PasswordSalt) === emp.PasswordHash;

  logLogin_(employeeId, ok, b.ip || '', b.userAgent || '');
  if (!ok) throw new Error('Mã nhân viên hoặc mật khẩu không đúng.');

  const token = token_(32);
  const session = {
    employeeId,
    fullName: emp.FullName,
    department: emp.Department,
    role: emp.Role,
    createdAt: Date.now()
  };

  CacheService.getScriptCache().put('SESSION_' + token, JSON.stringify(session), CONFIG.sessionSeconds);
  updateLastLogin_(emp);
  audit_(employeeId,'LOGIN','Employee',employeeId,'Đăng nhập thành công');

  return {ok:true, token, user:session};
}

function apiLogout_(b) {
  if (b.token) CacheService.getScriptCache().remove('SESSION_' + b.token);
  return {ok:true};
}

function apiMe_(b) {
  return {ok:true,user:session_(b.token)};
}

/* ---------- EMPLOYEE ---------- */

function apiDashboard_(b) {
  const s = session_(b.token);
  const all = visibleAnnouncements_(s.employeeId);
  const unread = unreadIds_(s.employeeId, all.map(a=>a.ID));
  const f = b.filters || {};

  const filtered = all.filter(a => {
    if (f.month && !sameMonth_(a.PublishDate,f.month)) return false;
    if (f.type && a.Type !== f.type) return false;
    if (f.department && !normalize_(targetText_(a.ID)).includes(normalize_(f.department))) return false;
    if (f.unreadOnly && !unread.has(a.ID)) return false;

    if (f.search) {
      const q=normalize_(f.search);
      const hay=normalize_([a.ID,a.Title,a.Content,a.Type,a.Keywords,targetText_(a.ID)].join(' '));
      if (!hay.includes(q)) return false;
    }
    return true;
  });

  return {
    ok:true,
    user:s,
    announcements:filtered.map(a=>publicAnnouncement_(a,unread.has(a.ID))),
    unreadCount:unread.size,
    types:TYPES,
    departments:departments_(),
    months:months_(all)
  };
}

function apiAnnouncement_(b) {
  const s=session_(b.token);
  const a=findAnnouncement_(b.id);
  if (!a) throw new Error('Không tìm thấy thông báo.');

  const admin=isAdmin_(s);
  if (!admin && !canView_(a.ID,s.employeeId)) throw new Error('Bạn không có quyền xem.');

  if (!admin) markRead_(a.ID,s.employeeId);

  return {
    ok:true,
    announcement:publicAnnouncement_(a,!admin),
    attachment:parseAttachment_(a.Attachment),
    readAt:findRead_(a.ID,s.employeeId)
  };
}

function apiMarkRead_(b) {
  const s=session_(b.token);
  const a=findAnnouncement_(b.id);
  if (!a || (!isAdmin_(s) && !canView_(a.ID,s.employeeId))) throw new Error('Không có quyền.');
  markRead_(a.ID,s.employeeId);
  return {ok:true};
}

function apiReadStats_(b) {
  const s=session_(b.token);
  requireAdmin_(s);

  const a=findAnnouncement_(b.id);
  if (!a) throw new Error('Không tìm thấy thông báo.');

  const audience=targetEmployeeIds_(a.ID);
  const readSet=new Set(readRows_(a.ID).map(x=>x.EmployeeID));
  let read=0;
  audience.forEach(id=>{if(readSet.has(id))read++;});

  return {
    ok:true,
    announcementId:a.ID,
    title:a.Title,
    audience:audience.length,
    read,
    unread:Math.max(0,audience.length-read),
    rate:audience.length ? read/audience.length*100 : 0
  };
}

/* ---------- ADMIN ---------- */

function apiAdminDashboard_(b) {
  const s=session_(b.token);
  requireAdmin_(s);

  const anns=announcementObjects_();
  const employees=employeeObjects_();
  const reads=data_(sheet_(CONFIG.sheets.reads));

  return {
    ok:true,
    stats:{
      employees:employees.filter(x=>x.Status==='ACTIVE').length,
      announcements:anns.length,
      published:anns.filter(x=>x.Status==='PUBLISHED').length,
      scheduled:anns.filter(x=>x.Status==='SCHEDULED').length,
      totalReads:reads.length
    },
    announcements:anns.sort((a,b)=>new Date(b.PublishDate||0)-new Date(a.PublishDate||0))
      .slice(0,500).map(publicAdminAnnouncement_)
  };
}

function apiSaveAnnouncement_(b) {
  const s=session_(b.token);
  requireAdmin_(s);
  const d=b.data || {};

  if (!d.title || !d.content || !d.type) throw new Error('Thiếu tiêu đề, nội dung hoặc loại.');

  const sh=sheet_(CONFIG.sheets.announcements);
  const now=new Date();
  const id=d.id || nextAnnouncementId_();
  const old=findAnnouncement_(id);
  const status=d.scheduledAt ? 'SCHEDULED' : (d.status || 'PUBLISHED');

  const row=[
    id,String(d.title).trim(),String(d.content),String(d.type),
    d.publishDate ? new Date(d.publishDate) : now,
    s.fullName,String(d.priority||'NORMAL'),!!d.pinned,status,
    d.attachmentJson ? JSON.stringify(d.attachmentJson) : (old ? old.Attachment : ''),
    String(d.keywords||''),old ? old.CreatedAt : now,now,
    d.scheduledAt ? new Date(d.scheduledAt) : ''
  ];

  if(old) sh.getRange(old._row,1,1,row.length).setValues([row]);
  else sh.appendRow(row);

  replaceTargets_(id,d.targets || [{type:'ALL',value:'ALL'}]);
  audit_(s.employeeId,old?'UPDATE_ANNOUNCEMENT':'CREATE_ANNOUNCEMENT','Announcement',id,d.title);

  return {ok:true,id};
}

function apiDeleteAnnouncement_(b) {
  const s=session_(b.token);
  requireAdmin_(s);
  const a=findAnnouncement_(b.id);
  if (!a) throw new Error('Không tìm thấy.');
  sheet_(CONFIG.sheets.announcements).deleteRow(a._row);
  audit_(s.employeeId,'DELETE_ANNOUNCEMENT','Announcement',a.ID,a.Title);
  return {ok:true};
}

function apiTogglePinned_(b) {
  const s=session_(b.token);
  requireAdmin_(s);
  const a=findAnnouncement_(b.id);
  if (!a) throw new Error('Không tìm thấy.');

  const sh=sheet_(CONFIG.sheets.announcements);
  const ix=idx_(HEADERS.Announcements);
  sh.getRange(a._row,ix.Pinned+1).setValue(!a.Pinned);
  sh.getRange(a._row,ix.UpdatedAt+1).setValue(new Date());
  audit_(s.employeeId,'TOGGLE_PIN','Announcement',a.ID,String(!a.Pinned));
  return {ok:true,pinned:!a.Pinned};
}

function apiChangePassword_(b) {
  const s=session_(b.token);
  if (!b.newPassword || String(b.newPassword).length<8) throw new Error('Mật khẩu mới tối thiểu 8 ký tự.');

  const e=findEmployee_(s.employeeId);
  if (!e || hash_(String(b.oldPassword||''),e.PasswordSalt)!==e.PasswordHash) {
    throw new Error('Mật khẩu cũ không đúng.');
  }

  const salt=token_(16);
  const sh=sheet_(CONFIG.sheets.employees);
  const ix=idx_(HEADERS.Employees);
  sh.getRange(e._row,ix.PasswordHash+1).setValue(hash_(b.newPassword,salt));
  sh.getRange(e._row,ix.PasswordSalt+1).setValue(salt);
  sh.getRange(e._row,ix.UpdatedAt+1).setValue(new Date());

  audit_(s.employeeId,'CHANGE_PASSWORD','Employee',s.employeeId,'');
  return {ok:true};
}

/* ---------- IMPORT ---------- */

function apiPreviewImport_(b) {
  const s=session_(b.token);
  requireAdmin_(s);

  const rows=normalizeImportRows_(b.rows || []);
  const existing={};
  employeeObjects_().forEach(e=>existing[e.EmployeeID]=e);

  const seen=new Set(), duplicates=[];
  const preview=rows.map(r=>{
    if(seen.has(r.EmployeeID)) duplicates.push(r.EmployeeID);
    seen.add(r.EmployeeID);
    const old=existing[r.EmployeeID];
    return {...r,action:old?'UPDATE':'ADD'};
  });

  return {
    ok:true,
    total:preview.length,
    duplicates:[...new Set(duplicates)],
    added:preview.filter(x=>x.action==='ADD').length,
    updated:preview.filter(x=>x.action==='UPDATE').length,
    rows:preview.slice(0,500)
  };
}

function apiCommitImport_(b) {
  const s=session_(b.token);
  requireAdmin_(s);

  const rows=normalizeImportRows_(b.rows || []);
  const missingAction=b.missingAction==='INACTIVE'?'INACTIVE':'NO_CHANGE';
  const fileName=String(b.fileName||'employee_import.xlsx');

  const sh=sheet_(CONFIG.sheets.employees);
  const map={};
  employeeObjects_().forEach(e=>map[e.EmployeeID]=e);

  const incoming=new Set();
  let added=0,updated=0,skipped=0,markedInactive=0;

  rows.forEach(r=>{
    if(!r.EmployeeID){skipped++;return;}
    incoming.add(r.EmployeeID);
    const old=map[r.EmployeeID];

    if(!old){
      const salt=token_(16);
      const initial=r.Password || r.EmployeeID;
      sh.appendRow([
        r.EmployeeID,r.FullName,r.Department,r.Position,r.Email,r.Phone,
        hash_(initial,salt),salt,r.Role||'EMPLOYEE','ACTIVE',new Date(),new Date(),''
      ]);
      added++;
    } else {
      const ix=idx_(HEADERS.Employees);
      sh.getRange(old._row,ix.FullName+1).setValue(r.FullName || old.FullName);
      sh.getRange(old._row,ix.Department+1).setValue(r.Department || old.Department);
      sh.getRange(old._row,ix.Position+1).setValue(r.Position || old.Position);
      sh.getRange(old._row,ix.Email+1).setValue(r.Email || old.Email);
      sh.getRange(old._row,ix.Phone+1).setValue(r.Phone || old.Phone);
      sh.getRange(old._row,ix.Role+1).setValue(r.Role || old.Role || 'EMPLOYEE');
      sh.getRange(old._row,ix.Status+1).setValue('ACTIVE');
      sh.getRange(old._row,ix.UpdatedAt+1).setValue(new Date());
      updated++;
    }
  });

  if(missingAction==='INACTIVE'){
    employeeObjects_().forEach(e=>{
      if(e.EmployeeID && !incoming.has(e.EmployeeID) && e.Status==='ACTIVE'){
        const ix=idx_(HEADERS.Employees);
        sh.getRange(e._row,ix.Status+1).setValue('INACTIVE');
        sh.getRange(e._row,ix.UpdatedAt+1).setValue(new Date());
        markedInactive++;
      }
    });
  }

  const importId='IMP-'+Utilities.getUuid().slice(0,8).toUpperCase();
  sheet_(CONFIG.sheets.imports).appendRow([
    importId,fileName,s.employeeId,new Date(),rows.length,added,updated,skipped,
    markedInactive,missingAction,'SUCCESS',''
  ]);

  audit_(s.employeeId,'IMPORT_EMPLOYEES','EmployeeImport',importId,
    JSON.stringify({fileName,total:rows.length,added,updated,skipped,markedInactive,missingAction}));

  return {ok:true,importId,added,updated,skipped,markedInactive};
}

function apiEmployees_(b) {
  const s=session_(b.token);
  requireAdmin_(s);
  const q=normalize_(b.query||'');
  let rows=employeeObjects_();
  if(q) rows=rows.filter(e=>normalize_([e.EmployeeID,e.FullName,e.Department,e.Position,e.Email].join(' ')).includes(q));
  return {ok:true,employees:rows.slice(0,1000)};
}

/* ---------- ATTACHMENTS ---------- */

function apiUploadAttachment_(b) {
  const s=session_(b.token);
  requireAdmin_(s);
  if(!b.file || !b.file.base64 || !b.file.name) throw new Error('Thiếu file.');

  const bytes=Utilities.base64Decode(b.file.base64);
  if(bytes.length>CONFIG.maxAttachmentBytes) throw new Error('File tối đa 25MB.');

  const folder=getAttachmentFolder_();
  const safe=String(b.file.name).replace(/[\\/:*?"<>|#%]/g,'_');
  const blob=Utilities.newBlob(bytes,b.file.mimeType||'application/octet-stream',safe);
  const file=folder.createFile(blob);

  try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK,DriveApp.Permission.VIEW); } catch(e) {}

  const id=file.getId();
  const attachment={
    id,
    name:safe,
    mimeType:file.getMimeType(),
    size:bytes.length,
    url:'https://drive.google.com/uc?export=download&id='+id,
    previewUrl:'https://drive.google.com/file/d/'+id+'/preview'
  };

  audit_(s.employeeId,'UPLOAD_ATTACHMENT','DriveFile',id,safe);
  return {ok:true,attachment};
}

function getAttachmentFolder_() {
  const props=PropertiesService.getScriptProperties();
  const id=props.getProperty('ATTACHMENT_FOLDER_ID');
  if(id) {
    try { return DriveApp.getFolderById(id); } catch(e) {}
  }

  const it=DriveApp.getFoldersByName('CompanyNoticeAttachments');
  const folder=it.hasNext()?it.next():DriveApp.createFolder('CompanyNoticeAttachments');
  props.setProperty('ATTACHMENT_FOLDER_ID',folder.getId());
  return folder;
}

/* ---------- DATA HELPERS ---------- */

function ss_() {
  const id=PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  return SpreadsheetApp.openById(id || SpreadsheetApp.getActive().getId());
}
function sheet_(name) {
  const s=ss_().getSheetByName(name);
  if(!s) throw new Error('Thiếu Sheet '+name+'. Chạy setupSystem().');
  return s;
}
function data_(sh) {
  return sh.getLastRow()<2 ? [] : sh.getRange(2,1,sh.getLastRow()-1,sh.getLastColumn()).getValues();
}
function idx_(headers) {
  const o={};headers.forEach((x,i)=>o[x]=i);return o;
}
function obj_(headers,row) {
  const o={};headers.forEach((x,i)=>o[x]=row[i]);return o;
}
function employeeObjects_() {
  return data_(sheet_(CONFIG.sheets.employees)).map((r,i)=>{const o=obj_(HEADERS.Employees,r);o._row=i+2;return o;});
}
function announcementObjects_() {
  return data_(sheet_(CONFIG.sheets.announcements)).map((r,i)=>{
    const o=obj_(HEADERS.Announcements,r);o._row=i+2;
    o.Pinned=o.Pinned===true || String(o.Pinned).toUpperCase()==='TRUE';
    return o;
  });
}
function findEmployee_(id) {
  return employeeObjects_().find(e=>String(e.EmployeeID).toUpperCase()===String(id).toUpperCase()) || null;
}
function findAnnouncement_(id) {
  return announcementObjects_().find(a=>String(a.ID)===String(id)) || null;
}
function session_(token) {
  if(!token) throw new Error('Phiên đăng nhập đã hết hạn.');
  const raw=CacheService.getScriptCache().get('SESSION_'+token);
  if(!raw) throw new Error('Phiên đăng nhập đã hết hạn.');
  return JSON.parse(raw);
}
function isAdmin_(s) { return s.role==='ADMIN' || s.role==='HR'; }
function requireAdmin_(s) { if(!isAdmin_(s)) throw new Error('Không có quyền quản trị.'); }

function visibleAnnouncements_(employeeId) {
  return announcementObjects_()
    .filter(a=>a.Status==='PUBLISHED' && (!a.PublishDate || new Date(a.PublishDate)<=new Date()))
    .filter(a=>canView_(a.ID,employeeId))
    .sort((a,b)=>{
      if(a.Pinned!==b.Pinned)return a.Pinned?-1:1;
      return new Date(b.PublishDate||0)-new Date(a.PublishDate||0);
    });
}

function canView_(announcementId,employeeId) {
  const targets=data_(sheet_(CONFIG.sheets.targets)).filter(r=>String(r[0])===String(announcementId));
  if(!targets.length)return true;

  const e=findEmployee_(employeeId);
  if(!e)return false;

  return targets.some(t=>{
    const type=String(t[1]||'ALL').toUpperCase();
    const value=String(t[2]||'ALL');
    if(type==='ALL')return true;
    if(type==='EMPLOYEE')return value.toUpperCase()===employeeId.toUpperCase();
    if(type==='DEPARTMENT')return normalize_(e.Department)===normalize_(value);
    if(type==='TYPE')return normalize_(e.Role)===normalize_(value) || normalize_(e.Position)===normalize_(value);
    return false;
  });
}

function replaceTargets_(id,targets) {
  const sh=sheet_(CONFIG.sheets.targets);
  const rows=data_(sh);
  for(let i=rows.length-1;i>=0;i--) if(String(rows[i][0])===String(id)) sh.deleteRow(i+2);
  (targets.length?targets:[{type:'ALL',value:'ALL'}]).forEach(t=>{
    sh.appendRow([id,String(t.type||'ALL').toUpperCase(),String(t.value||'ALL')]);
  });
}

function targetText_(id) {
  const rows=data_(sheet_(CONFIG.sheets.targets)).filter(r=>String(r[0])===String(id));
  if(!rows.length)return 'Tất cả';
  return rows.map(r=>`${r[1]}: ${r[2]}`).join(', ');
}

function targetEmployeeIds_(id) {
  return employeeObjects_().filter(e=>e.Status==='ACTIVE')
    .filter(e=>canView_(id,e.EmployeeID)).map(e=>e.EmployeeID);
}

function readRows_(id) {
  return data_(sheet_(CONFIG.sheets.reads)).filter(r=>String(r[0])===String(id))
    .map(r=>({AnnouncementID:r[0],EmployeeID:r[1],ReadAt:r[2]}));
}

function findRead_(id,employeeId) {
  const x=readRows_(id).find(r=>String(r.EmployeeID)===String(employeeId));
  return x?dateOut_(x.ReadAt):null;
}
function markRead_(id,employeeId) {
  if(findRead_(id,employeeId))return;
  sheet_(CONFIG.sheets.reads).appendRow([id,employeeId,new Date()]);
}
function unreadIds_(employeeId,ids) {
  const read=new Set(data_(sheet_(CONFIG.sheets.reads))
    .filter(r=>String(r[1])===String(employeeId)).map(r=>String(r[0])));
  return new Set(ids.filter(id=>!read.has(String(id))));
}

function publicAnnouncement_(a,unread) {
  return {
    id:a.ID,title:a.Title,content:a.Content,type:a.Type,publishDate:dateOut_(a.PublishDate),
    author:a.Author,priority:a.Priority,pinned:a.Pinned,status:a.Status,
    attachment:parseAttachment_(a.Attachment),keywords:a.Keywords,
    unread:!!unread,targetText:targetText_(a.ID)
  };
}
function publicAdminAnnouncement_(a) {
  const x=publicAnnouncement_(a,false);
  x.scheduledAt=dateOut_(a.ScheduledAt);
  x.createdAt=dateOut_(a.CreatedAt);
  return x;
}
function parseAttachment_(x) {
  if(!x)return null;
  try{return typeof x==='string'?JSON.parse(x):x}catch(e){return {name:String(x)}}
}
function departments_() {
  return [...new Set(employeeObjects_().map(e=>String(e.Department||'').trim()).filter(Boolean))].sort();
}
function months_(rows) {
  const o={};
  rows.forEach(a=>{
    if(a.PublishDate){
      const k=Utilities.formatDate(new Date(a.PublishDate),CONFIG.timezone,'MM/yyyy');
      o[k]=(o[k]||0)+1;
    }
  });
  return Object.keys(o).sort((a,b)=>monthKey_(b)-monthKey_(a)).map(k=>({month:k,count:o[k]}));
}
function monthKey_(x) { const p=x.split('/');return Number(p[1])*100+Number(p[0]); }
function sameMonth_(date,month) {
  return !date || Utilities.formatDate(new Date(date),CONFIG.timezone,'MM/yyyy')===month;
}
function nextAnnouncementId_() {
  return 'TB-'+Utilities.formatDate(new Date(),CONFIG.timezone,'yyyyMMdd')+'-'+Utilities.getUuid().slice(0,6).toUpperCase();
}
function normalize_(v) {
  return String(v||'').toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g,'').trim();
}
function dateOut_(d) {
  if(!d)return '';
  const x=new Date(d);if(isNaN(x))return '';
  return Utilities.formatDate(x,CONFIG.timezone,"yyyy-MM-dd'T'HH:mm:ss");
}
function token_(bytes) {
  const raw=Utilities.getUuid().replace(/-/g,'')+Utilities.getUuid().replace(/-/g,'');
  return raw.slice(0,bytes*2);
}
function hash_(password,salt) {
  const bytes=Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,String(salt)+'|'+String(password),Utilities.Charset.UTF_8);
  return bytes.map(b=>(b<0?b+256:b).toString(16).padStart(2,'0')).join('');
}
function logLogin_(id,ok,ip,ua) {
  sheet_(CONFIG.sheets.logins).appendRow(['LOGIN-'+token_(8),id,new Date(),ok,ip,ua]);
}
function updateLastLogin_(e) {
  sheet_(CONFIG.sheets.employees).getRange(e._row,idx_(HEADERS.Employees).LastLoginAt+1).setValue(new Date());
}
function audit_(actor,action,entity,id,details) {
  sheet_(CONFIG.sheets.audit).appendRow(['AUD-'+token_(8),new Date(),actor,action,entity,id,String(details||'')]);
}
function normalizeImportRows_(rows) {
  if(!Array.isArray(rows))throw new Error('Dữ liệu Excel không hợp lệ.');
  return rows.map(r=>{
    const keys=Object.keys(r);
    const pick=names=>{
      for(const n of names){
        const k=keys.find(x=>normalize_(x)===normalize_(n));
        if(k!==undefined)return r[k];
      }
      return '';
    };
    return {
      EmployeeID:String(pick(['EmployeeID','Mã nhân viên','Ma nhan vien','Mã NV','Ma NV'])).trim().toUpperCase(),
      FullName:String(pick(['FullName','Họ tên','Ho ten','Tên nhân viên','Ten nhan vien'])).trim(),
      Department:String(pick(['Department','Phòng ban','Phong ban','Bộ phận','Bo phan'])).trim(),
      Position:String(pick(['Position','Chức vụ','Chuc vu','Vị trí','Vi tri'])).trim(),
      Email:String(pick(['Email','E-mail'])).trim(),
      Phone:String(pick(['Phone','Điện thoại','Dien thoai','SĐT','SDT'])).trim(),
      Role:String(pick(['Role','Vai trò','Vai tro'])).trim().toUpperCase() || 'EMPLOYEE',
      Password:String(pick(['Password','Mật khẩu','Mat khau'])).trim()
    };
  }).filter(x=>x.EmployeeID);
}
function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o))
    .setMimeType(ContentService.MimeType.JSON);
}
function parseBody_(e) {
  if(!e || !e.postData || !e.postData.contents) return {};
  const raw=e.postData.contents;
  try{return JSON.parse(raw)}catch(err){}
  const p=e.parameter||{};
  return p;
}
