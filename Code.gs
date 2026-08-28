/**
 * CỔNG THÔNG BÁO NỘI BỘ - Google Apps Script
 * Version 1.0
 * Mô hình: Google Sheets + GAS + Google Drive
 *
 * CÀI ĐẶT:
 * 1) Tạo Google Sheet -> Extensions > Apps Script.
 * 2) Tạo file HTML tên index và dán Index.html.
 * 3) Dán toàn bộ Code.gs này.
 * 4) Save.
 * 5) Deploy > New deployment > Web app:
 *      Execute as: Me
 *      Who has access: Anyone with Google account / Anyone (tùy chính sách công ty)
 * 6) Mở URL /exec. Hệ thống tự tạo Sheet + Drive.
 *
 * Tài khoản mặc định lần đầu:
 *      ADMIN
 *      Admin@123456
 * Sau khi đăng nhập hãy đổi mật khẩu ngay.
 */

const APP = {
  NAME: 'Cổng thông báo nội bộ',
  VERSION: '1.0.0',
  SESSION_TTL: 6 * 60 * 60,
  MAX_FILE_BYTES: 20 * 1024 * 1024,
  DEFAULT_ADMIN: 'ADMIN',
  DEFAULT_ADMIN_PASSWORD: 'Admin@123456',
  ROOT_FOLDER: 'COMPANY_INTERNAL_PORTAL',
  SHEETS: {
    CONFIG: 'Config',
    USERS: 'Users',
    EMPLOYEES: 'Employees',
    ANNOUNCEMENTS: 'Announcements',
    TARGETS: 'AnnouncementTargets',
    ATTACHMENTS: 'AnnouncementAttachments',
    READS: 'AnnouncementReads',
    IMPORTS: 'ImportHistory',
    LOGINS: 'LoginHistory',
    AUDIT: 'AuditLog'
  },
  TYPES: [
    '📢 Thông báo chung',
    '🕐 Thời gian làm việc',
    '💰 Lương & thưởng',
    '🎁 Phúc lợi',
    '🏥 Sức khỏe',
    '🦺 An toàn / HSE',
    '🏭 Sản xuất',
    '📋 Nhân sự',
    '🎉 Hoạt động công ty',
    '⚠️ Khẩn cấp'
  ]
};

const HEADERS = {
  Config: ['Key','Value','UpdatedAt'],
  Users: ['Username','PasswordHash','Salt','Role','EmployeeID','Status','CreatedAt','UpdatedAt','MustChangePassword'],
  Employees: ['EmployeeID','EmployeeCode','FullName','Department','Factory','Position','Email','Phone','Status','CreatedAt','UpdatedAt'],
  Announcements: ['ID','Title','Content','Type','PublishDate','ScheduledAt','Author','Priority','Pinned','Status','Keywords','CreatedAt','UpdatedAt'],
  AnnouncementTargets: ['ID','AnnouncementID','TargetType','TargetValue'],
  AnnouncementAttachments: ['ID','AnnouncementID','FileName','FileId','MimeType','Size','DriveUrl','PreviewUrl','CreatedAt'],
  AnnouncementReads: ['AnnouncementID','EmployeeID','ReadAt'],
  ImportHistory: ['ImportID','FileName','ImportedBy','ImportedAt','TotalRows','NewEmployees','UpdatedEmployees','UnchangedEmployees','InactiveEmployees','ErrorRows','MissingPolicy','Details'],
  LoginHistory: ['LoginAt','Username','EmployeeID','Success','Ip','UserAgent'],
  AuditLog: ['Timestamp','Actor','Action','Entity','EntityID','Details']
};

function doGet() {
  ensureSystem_();
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle(APP.NAME)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function api(action, token, data) {
  try {
    action = String(action || '').trim().toUpperCase();
    data = data || {};
    if (action === 'BOOTSTRAP') return bootstrap_();
    if (action === 'INITIALIZE') return initializeSystem_(data.password);
    if (action === 'LOGIN') return login_(data.username, data.password);
    if (action === 'SETUP_STATUS') return setupStatus_();

    ensureSystem_();
    const user = requireSession_(token);

    switch (action) {
      case 'LOGOUT': return logout_(token, user);
      case 'GET_ME': return {status:'SUCCESS', user:safeUser_(user)};
      case 'GET_ANNOUNCEMENTS': return getAnnouncements_(user, data);
      case 'GET_ANNOUNCEMENT': return getAnnouncement_(user, data.id);
      case 'MARK_READ': return markRead_(user, data.id);
      case 'GET_FILTERS': return getFilters_(user);
      case 'GET_READ_STATS': requireRole_(user,['ADMIN','HR ADMIN','HR']); return getReadStats_(data.id);
      case 'SAVE_ANNOUNCEMENT': requireRole_(user,['ADMIN','HR ADMIN','HR']); return saveAnnouncement_(user,data);
      case 'DELETE_ANNOUNCEMENT': requireRole_(user,['ADMIN','HR ADMIN','HR']); return deleteAnnouncement_(user,data.id);
      case 'TOGGLE_PIN': requireRole_(user,['ADMIN','HR ADMIN','HR']); return togglePin_(user,data.id);
      case 'GET_ADMIN_ANNOUNCEMENTS': requireRole_(user,['ADMIN','HR ADMIN','HR']); return getAdminAnnouncements_(data);
      case 'GET_EMPLOYEES': requireRole_(user,['ADMIN','HR ADMIN','HR']); return getEmployees_(data);
      case 'IMPORT_PREVIEW': requireRole_(user,['ADMIN','HR ADMIN']); return importPreview_(user,data);
      case 'IMPORT_CONFIRM': requireRole_(user,['ADMIN','HR ADMIN']); return importConfirm_(user,data);
      case 'GET_IMPORT_HISTORY': requireRole_(user,['ADMIN','HR ADMIN']); return getImportHistory_();
      case 'CHANGE_PASSWORD': return changePassword_(user,data.oldPassword,data.newPassword);
      case 'UPLOAD_ATTACHMENT': requireRole_(user,['ADMIN','HR ADMIN','HR']); return uploadAttachment_(user,data);
      case 'DELETE_ATTACHMENT': requireRole_(user,['ADMIN','HR ADMIN','HR']); return deleteAttachment_(user,data.fileId,data.attachmentId);
      default: return {status:'ERROR',message:'Hành động không hợp lệ.'};
    }
  } catch (e) {
    console.error(e);
    return {status:'ERROR',message:e && e.message ? e.message : String(e)};
  }
}

/* =========================
   SYSTEM / DATABASE
========================= */

function getSS_() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

function ensureSystem_() {
  const ss = getSS_();
  Object.keys(HEADERS).forEach(name => {
    let sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    const h = HEADERS[name];
    if (sh.getLastRow() === 0) sh.getRange(1,1,1,h.length).setValues([h]);
    else {
      const current = sh.getRange(1,1,1,Math.max(sh.getLastColumn(),h.length)).getValues()[0];
      if (current.slice(0,h.length).join('|') !== h.join('|')) {
        sh.getRange(1,1,1,h.length).setValues([h]);
      }
    }
    sh.setFrozenRows(1);
    sh.getRange(1,1,1,h.length).setFontWeight('bold');
  });
  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty('ROOT_FOLDER_ID')) {
    const folder = getOrCreateFolder_(APP.ROOT_FOLDER);
    props.setProperty('ROOT_FOLDER_ID', folder.getId());
    props.setProperty('ANNOUNCEMENT_FOLDER_ID', getOrCreateChildFolder_(folder,'Announcements').getId());
    props.setProperty('IMPORT_FOLDER_ID', getOrCreateChildFolder_(folder,'EmployeeImports').getId());
    props.setProperty('TEMP_FOLDER_ID', getOrCreateChildFolder_(folder,'Temp').getId());
  }
  if (!getConfig_('SYSTEM_INITIALIZED')) {
    setConfig_('SYSTEM_INITIALIZED','false');
  }
}

function bootstrap_() {
  ensureSystem_();
  return {
    status:'SUCCESS',
    initialized:getConfig_('SYSTEM_INITIALIZED') === 'true',
    app:{name:APP.NAME,version:APP.VERSION,types:APP.TYPES},
    setup: setupStatus_()
  };
}

function setupStatus_() {
  ensureSystem_();
  return {status:'SUCCESS',initialized:getConfig_('SYSTEM_INITIALIZED') === 'true'};
}

/**
 * Có thể gọi trực tiếp từ editor nếu muốn, nhưng KHÔNG bắt buộc.
 * Không còn lỗi "Hãy chạy setupSystem(...)".
 */
function setupSystem() {
  ensureSystem_();
  const password = APP.DEFAULT_ADMIN_PASSWORD;
  return initializeSystem_(password);
}

function initializeSystem_(password) {
  ensureSystem_();
  if (getConfig_('SYSTEM_INITIALIZED') === 'true') {
    return {status:'SUCCESS',alreadyInitialized:true,message:'Hệ thống đã được khởi tạo.'};
  }
  password = String(password || '').trim();
  if (password.length < 8) throw new Error('Mật khẩu Admin phải có ít nhất 8 ký tự.');

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    if (getConfig_('SYSTEM_INITIALIZED') === 'true') {
      return {status:'SUCCESS',alreadyInitialized:true};
    }
    const now = new Date();
    const salt = Utilities.getUuid();
    const hash = hashPassword_(password,salt);
    appendRow_('Users',[
      APP.DEFAULT_ADMIN,hash,salt,'ADMIN','', 'ACTIVE',now,now,'TRUE'
    ]);
    setConfig_('SYSTEM_INITIALIZED','true');
    setConfig_('INITIALIZED_AT',now.toISOString());
    setConfig_('DEFAULT_ADMIN','ADMIN');
    audit_('SYSTEM','INITIALIZE','SYSTEM','','Khởi tạo hệ thống');
    return {status:'SUCCESS',message:'Khởi tạo thành công. Tài khoản ADMIN đã được tạo.'};
  } finally {
    lock.releaseLock();
  }
}

function getConfig_(key) {
  const rows = getRows_('Config');
  const r = rows.find(x => String(x.Key) === String(key));
  return r ? r.Value : '';
}
function setConfig_(key,value) {
  const sh = sheet_('Config'), values = sh.getDataRange().getValues();
  for (let i=1;i<values.length;i++) {
    if (String(values[i][0]) === String(key)) {
      sh.getRange(i+1,2,1,2).setValues([[value,new Date()]]);
      return;
    }
  }
  sh.appendRow([key,value,new Date()]);
}

function sheet_(name) {
  const sh = getSS_().getSheetByName(name);
  if (!sh) throw new Error('Sheet chưa tồn tại: '+name);
  return sh;
}

function getRows_(name) {
  const sh = sheet_(name), data = sh.getDataRange().getValues();
  if (data.length < 2) return [];
  const headers = data[0];
  return data.slice(1).map((r,i)=>{
    const o={_row:i+2};
    headers.forEach((h,j)=>o[h]=r[j]);
    return o;
  });
}

function appendRow_(name,row) { sheet_(name).appendRow(row); }

function updateObjectRow_(name,rowNumber,obj) {
  const sh=sheet_(name), headers=HEADERS[name];
  const vals=headers.map(h=>obj[h] !== undefined ? obj[h] : '');
  sh.getRange(rowNumber,1,1,vals.length).setValues([vals]);
}

function deleteRow_(name,rowNumber) { sheet_(name).deleteRow(rowNumber); }

function now_(){return new Date();}
function iso_(v){ if(!v) return ''; const d=v instanceof Date?v:new Date(v); return isNaN(d)?String(v):d.toISOString(); }
function displayDate_(v){ if(!v)return ''; const d=v instanceof Date?v:new Date(v); if(isNaN(d))return String(v); return Utilities.formatDate(d,Session.getScriptTimeZone()||'Asia/Ho_Chi_Minh','dd/MM/yyyy HH:mm'); }
function dateOnly_(v){ if(!v)return ''; const d=v instanceof Date?v:new Date(v); if(isNaN(d))return ''; return Utilities.formatDate(d,Session.getScriptTimeZone()||'Asia/Ho_Chi_Minh','yyyy-MM-dd'); }
function esc_(s){return String(s??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));}

/* =========================
   AUTH
========================= */

function randomToken_(){return Utilities.getUuid()+'-'+Utilities.getUuid();}
function hashPassword_(password,salt) {
  const bytes=Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,String(salt)+'|'+String(password));
  return bytes.map(b=>((b+256)%256).toString(16).padStart(2,'0')).join('');
}
function verifyPassword_(password,salt,hash){return hashPassword_(password,salt)===hash;}

function login_(username,password) {
  ensureSystem_();
  if (getConfig_('SYSTEM_INITIALIZED') !== 'true') return {status:'SETUP_REQUIRED'};
  username=String(username||'').trim();
  const user=getRows_('Users').find(x=>String(x.Username).toLowerCase()===username.toLowerCase());
  const ok=user && String(user.Status).toUpperCase()==='ACTIVE' && verifyPassword_(String(password||''),String(user.Salt),String(user.PasswordHash));
  appendRow_('LoginHistory',[now_(),username,user?user.EmployeeID:'',ok?'TRUE':'FALSE','','']);
  if(!ok) return {status:'ERROR',message:'Mã nhân viên hoặc mật khẩu không đúng.'};

  let emp=null;
  if(user.EmployeeID) emp=getRows_('Employees').find(x=>String(x.EmployeeID)===String(user.EmployeeID)||String(x.EmployeeCode)===String(user.EmployeeID));
  const token=randomToken_();
  CacheService.getScriptCache().put('SESSION_'+token,JSON.stringify({
    username:user.Username,role:user.Role,employeeId:user.EmployeeID||'',name:emp?emp.FullName:user.Username,
    department:emp?emp.Department:'',factory:emp?emp.Factory:'',mustChangePassword:String(user.MustChangePassword)==='TRUE'
  }),APP.SESSION_TTL);
  return {status:'SUCCESS',token,user:safeUser_({
    username:user.Username,role:user.Role,employeeId:user.EmployeeID||'',name:emp?emp.FullName:user.Username,
    department:emp?emp.Department:'',factory:emp?emp.Factory:'',mustChangePassword:String(user.MustChangePassword)==='TRUE'
  })};
}

function requireSession_(token) {
  if(!token) throw new Error('Phiên đăng nhập không hợp lệ.');
  const raw=CacheService.getScriptCache().get('SESSION_'+token);
  if(!raw) throw new Error('Phiên đăng nhập đã hết hạn. Vui lòng đăng nhập lại.');
  const u=JSON.parse(raw);
  CacheService.getScriptCache().put('SESSION_'+token,JSON.stringify(u),APP.SESSION_TTL);
  return u;
}
function safeUser_(u){const x=Object.assign({},u); delete x.password; delete x.PasswordHash; delete x.Salt; return x;}
function requireRole_(u,roles){if(!roles.includes(u.role))throw new Error('Bạn không có quyền thực hiện thao tác này.');}
function logout_(token,user){CacheService.getScriptCache().remove('SESSION_'+token);audit_(user.username,'LOGOUT','SESSION','', 'Đăng xuất');return {status:'SUCCESS'};}

function changePassword_(user,oldPassword,newPassword) {
  if(String(newPassword||'').length<8) throw new Error('Mật khẩu mới phải có ít nhất 8 ký tự.');
  const rows=getRows_('Users'), r=rows.find(x=>String(x.Username)===String(user.username));
  if(!r || !verifyPassword_(oldPassword,r.Salt,r.PasswordHash)) throw new Error('Mật khẩu hiện tại không đúng.');
  const salt=Utilities.getUuid(), hash=hashPassword_(newPassword,salt);
  const sh=sheet_('Users');
  sh.getRange(r._row,2,1,2).setValues([[hash,salt]]);
  sh.getRange(r._row,9).setValue('FALSE');
  audit_(user.username,'CHANGE_PASSWORD','USER',user.username,'Đổi mật khẩu');
  return {status:'SUCCESS',message:'Đổi mật khẩu thành công.'};
}

/* =========================
   EMPLOYEES / VISIBILITY
========================= */

function getEmployee_(id){
  if(!id)return null;
  return getRows_('Employees').find(x=>String(x.EmployeeID)===String(id)||String(x.EmployeeCode)===String(id));
}
function visibleToUser_(announcement,user){
  if(['ADMIN','HR ADMIN','HR'].includes(user.role)) return true;
  const targets=getRows_('AnnouncementTargets').filter(x=>String(x.AnnouncementID)===String(announcement.ID));
  if(!targets.length)return true;
  return targets.some(t=>{
    const type=String(t.TargetType).toUpperCase(), value=String(t.TargetValue).trim().toLowerCase();
    if(type==='ALL')return true;
    if(type==='EMPLOYEE')return value===String(user.employeeId||'').trim().toLowerCase();
    if(type==='DEPARTMENT')return value===String(user.department||'').trim().toLowerCase();
    if(type==='FACTORY')return value===String(user.factory||'').trim().toLowerCase();
    return false;
  });
}

function getPublishedAnnouncements_(){
  const now=new Date();
  return getRows_('Announcements').filter(a=>{
    const status=String(a.Status);
    if(status!=='Published')return false;
    const publish=a.ScheduledAt ? new Date(a.ScheduledAt) : new Date(a.PublishDate);
    return isNaN(publish) || publish<=now;
  });
}

function getAnnouncements_(user,filter){
  filter=filter||{};
  let list=getPublishedAnnouncements_().filter(a=>visibleToUser_(a,user));
  const q=String(filter.q||'').trim().toLowerCase();
  const month=String(filter.month||'');
  const type=String(filter.type||'');
  const dept=String(filter.department||'');
  const unreadOnly=!!filter.unreadOnly;
  const reads=new Set(getRows_('AnnouncementReads').filter(r=>String(r.EmployeeID)===String(user.employeeId)).map(r=>String(r.AnnouncementID)));

  if(q)list=list.filter(a=>[a.ID,a.Title,a.Content,a.Type,a.Keywords].join(' ').toLowerCase().includes(q));
  if(type)list=list.filter(a=>String(a.Type)===type);
  if(month)list=list.filter(a=>dateOnly_(a.PublishDate).slice(0,7)===month);
  if(dept) {
    const targets=getRows_('AnnouncementTargets');
    list=list.filter(a=>targets.some(t=>String(t.AnnouncementID)===String(a.ID)&&String(t.TargetType).toUpperCase()==='DEPARTMENT'&&String(t.TargetValue)===dept));
  }
  if(unreadOnly)list=list.filter(a=>!reads.has(String(a.ID)));

  list.sort((a,b)=>{
    if(String(a.Pinned)==='TRUE' && String(b.Pinned)!=='TRUE')return -1;
    if(String(b.Pinned)==='TRUE' && String(a.Pinned)!=='TRUE')return 1;
    const pa={URGENT:0,HIGH:1,NORMAL:2,LOW:3};
    const pp=(pa[String(a.Priority)]??2)-(pa[String(b.Priority)]??2);
    if(pp)return pp;
    return new Date(b.PublishDate)-new Date(a.PublishDate);
  });

  return {status:'SUCCESS',items:list.map(a=>announcementDto_(a,user,reads.has(String(a.ID))))};
}

function announcementDto_(a,user,isRead){
  const targets=getRows_('AnnouncementTargets').filter(x=>String(x.AnnouncementID)===String(a.ID));
  const atts=getRows_('AnnouncementAttachments').filter(x=>String(x.AnnouncementID)===String(a.ID));
  return {
    id:a.ID,title:a.Title,content:a.Content,type:a.Type,publishDate:iso_(a.PublishDate),
    publishDateText:displayDate_(a.PublishDate),scheduledAt:iso_(a.ScheduledAt),author:a.Author,
    priority:a.Priority,pinned:String(a.Pinned)==='TRUE',status:a.Status,keywords:a.Keywords,
    read:!!isRead,attachments:atts.map(attachmentDto_),targets:targets.map(t=>({type:t.TargetType,value:t.TargetValue}))
  };
}

function attachmentDto_(a){return {id:a.ID,announcementId:a.AnnouncementID,fileName:a.FileName,fileId:a.FileId,mimeType:a.MimeType,size:a.Size,driveUrl:a.DriveUrl,previewUrl:a.PreviewUrl};}

function getAnnouncement_(user,id){
  const a=getRows_('Announcements').find(x=>String(x.ID)===String(id));
  if(!a||!visibleToUser_(a,user))throw new Error('Không tìm thấy thông báo hoặc bạn không có quyền xem.');
  const reads=getRows_('AnnouncementReads').some(r=>String(r.AnnouncementID)===String(id)&&String(r.EmployeeID)===String(user.employeeId));
  return {status:'SUCCESS',item:announcementDto_(a,user,reads)};
}

function markRead_(user,id){
  if(!user.employeeId)throw new Error('Tài khoản chưa gắn mã nhân viên.');
  const a=getRows_('Announcements').find(x=>String(x.ID)===String(id));
  if(!a||!visibleToUser_(a,user))throw new Error('Không có quyền.');
  const sh=sheet_('AnnouncementReads'), exists=getRows_('AnnouncementReads').some(r=>String(r.AnnouncementID)===String(id)&&String(r.EmployeeID)===String(user.employeeId));
  if(!exists){sh.appendRow([id,user.employeeId,now_()]);audit_(user.username,'READ','ANNOUNCEMENT',id,'Đã đọc thông báo');}
  return {status:'SUCCESS',readAt:now_().toISOString()};
}

function getFilters_(user){
  const list=getPublishedAnnouncements_().filter(a=>visibleToUser_(a,user));
  const months=[...new Set(list.map(a=>dateOnly_(a.PublishDate).slice(0,7)).filter(Boolean))].sort().reverse();
  const departments=[...new Set(getRows_('Employees').map(e=>String(e.Department||'').trim()).filter(Boolean))].sort((a,b)=>a.localeCompare(b,'vi'));
  return {status:'SUCCESS',months,departments,types:APP.TYPES};
}

/* =========================
   ADMIN ANNOUNCEMENTS
========================= */

function nextAnnouncementId_(){
  const d=Utilities.formatDate(new Date(),Session.getScriptTimeZone()||'Asia/Ho_Chi_Minh','yyyyMMdd');
  const prefix='TB-'+d+'-';
  const ids=getRows_('Announcements').map(x=>String(x.ID)).filter(x=>x.indexOf(prefix)===0);
  return prefix+String(ids.length+1).padStart(3,'0');
}

function saveAnnouncement_(user,data){
  if(!data.title||!String(data.title).trim())throw new Error('Vui lòng nhập tiêu đề.');
  if(!data.content||!String(data.content).trim())throw new Error('Vui lòng nhập nội dung.');
  const sh=sheet_('Announcements'), rows=getRows_('Announcements');
  const id=data.id || nextAnnouncementId_(), existing=rows.find(x=>String(x.ID)===String(id));
  const now=new Date();
  const publish=data.publishDate ? new Date(data.publishDate) : now;
  const scheduled=data.scheduledAt ? new Date(data.scheduledAt) : '';
  const status=data.status || (scheduled && scheduled>now ? 'Scheduled':'Published');
  const obj={
    ID:id,Title:String(data.title).trim(),Content:String(data.content),
    Type:data.type||APP.TYPES[0],PublishDate:publish,ScheduledAt:scheduled,
    Author:existing?existing.Author:user.username,Priority:data.priority||'NORMAL',
    Pinned:data.pinned?'TRUE':'FALSE',Status:status,Keywords:String(data.keywords||''),
    CreatedAt:existing?existing.CreatedAt:now,UpdatedAt:now
  };
  if(existing)updateObjectRow_('Announcements',existing._row,obj);
  else appendRow_('Announcements',HEADERS.Announcements.map(h=>obj[h]));
  // targets replace
  getRows_('AnnouncementTargets').filter(x=>String(x.AnnouncementID)===String(id)).sort((a,b)=>b._row-a._row).forEach(x=>deleteRow_('AnnouncementTargets',x._row));
  (data.targets||[{type:'ALL',value:'ALL'}]).forEach(t=>appendRow_('AnnouncementTargets',[Utilities.getUuid(),id,String(t.type||'ALL').toUpperCase(),String(t.value||'ALL')]));
  audit_(user.username,existing?'UPDATE':'CREATE','ANNOUNCEMENT',id,obj.Title);
  return {status:'SUCCESS',id,message:existing?'Đã cập nhật thông báo.':'Đã tạo thông báo.'};
}

function getAdminAnnouncements_(filter){
  filter=filter||{};
  let list=getRows_('Announcements');
  const q=String(filter.q||'').toLowerCase();
  if(q)list=list.filter(a=>[a.ID,a.Title,a.Type,a.Author,a.Keywords].join(' ').toLowerCase().includes(q));
  list.sort((a,b)=>new Date(b.PublishDate)-new Date(a.PublishDate));
  return {status:'SUCCESS',items:list.map(a=>announcementDto_(a,{},false)),stats:{
    total:list.length,published:list.filter(a=>a.Status==='Published').length,scheduled:list.filter(a=>a.Status==='Scheduled').length
  }};
}

function deleteAnnouncement_(user,id){
  const a=getRows_('Announcements').find(x=>String(x.ID)===String(id));
  if(!a)throw new Error('Thông báo không tồn tại.');
  // Soft delete để giữ lịch sử đọc/audit.
  const sh=sheet_('Announcements'); sh.getRange(a._row,10).setValue('Deleted'); sh.getRange(a._row,13).setValue(now_());
  audit_(user.username,'DELETE','ANNOUNCEMENT',id,a.Title);
  return {status:'SUCCESS',message:'Đã xóa thông báo.'};
}
function togglePin_(user,id){
  const a=getRows_('Announcements').find(x=>String(x.ID)===String(id)); if(!a)throw new Error('Không tìm thấy.');
  const value=String(a.Pinned)==='TRUE'?'FALSE':'TRUE'; sheet_('Announcements').getRange(a._row,8+1).setValue(value);
  audit_(user.username,'PIN','ANNOUNCEMENT',id,value==='TRUE'?'Ghim':'Bỏ ghim');
  return {status:'SUCCESS',pinned:value==='TRUE'};
}

function getReadStats_(id){
  const a=getRows_('Announcements').find(x=>String(x.ID)===String(id)); if(!a)throw new Error('Không tìm thấy thông báo.');
  const employees=getRows_('Employees').filter(e=>String(e.Status||'ACTIVE').toUpperCase()!=='INACTIVE');
  const targets=getRows_('AnnouncementTargets').filter(t=>String(t.AnnouncementID)===String(id));
  const eligible=employees.filter(e=>targets.length===0||targets.some(t=>{
    const ty=String(t.TargetType).toUpperCase(), v=String(t.TargetValue).toLowerCase();
    if(ty==='ALL')return true;
    if(ty==='EMPLOYEE')return v===String(e.EmployeeID).toLowerCase()||v===String(e.EmployeeCode).toLowerCase();
    if(ty==='DEPARTMENT')return v===String(e.Department).toLowerCase();
    if(ty==='FACTORY')return v===String(e.Factory).toLowerCase();
    return false;
  }));
  const readMap={}; getRows_('AnnouncementReads').filter(r=>String(r.AnnouncementID)===String(id)).forEach(r=>readMap[String(r.EmployeeID)]=r.ReadAt);
  const read=eligible.filter(e=>readMap[String(e.EmployeeID)]).length, total=eligible.length;
  return {status:'SUCCESS',id,title:a.Title,total,read,unread:Math.max(0,total-read),rate:total?Math.round(read/total*10000)/100:0,
    unreadEmployees:eligible.filter(e=>!readMap[String(e.EmployeeID)]).map(e=>({employeeId:e.EmployeeID,code:e.EmployeeCode,name:e.FullName,department:e.Department}))};
}

/* =========================
   FILES / DRIVE
========================= */

function getOrCreateFolder_(name){
  const it=DriveApp.getFoldersByName(name); return it.hasNext()?it.next():DriveApp.createFolder(name);
}
function getOrCreateChildFolder_(parent,name){
  const it=parent.getFoldersByName(name); return it.hasNext()?it.next():parent.createFolder(name);
}
function uploadAttachment_(user,data){
  const bytes=Utilities.base64Decode(String(data.base64||''));
  if(!bytes.length)throw new Error('File rỗng.');
  if(bytes.length>APP.MAX_FILE_BYTES)throw new Error('File vượt quá 20 MB.');
  const mime=String(data.mimeType||'application/octet-stream');
  const allowed=/^(application\/pdf|image\/(jpeg|png|webp|gif)|video\/(mp4|webm|quicktime)|application\/(msword|vnd\.openxmlformats-officedocument\.wordprocessingml\.document|vnd\.ms-excel|vnd\.openxmlformats-officedocument\.spreadsheetml\.sheet|vnd\.ms-powerpoint|vnd\.openxmlformats-officedocument\.presentationml\.presentation))$/i;
  if(!allowed.test(mime))throw new Error('Định dạng file chưa được hỗ trợ.');
  const folder=DriveApp.getFolderById(PropertiesService.getScriptProperties().getProperty('ANNOUNCEMENT_FOLDER_ID'));
  const blob=Utilities.newBlob(bytes,mime,String(data.fileName||'attachment'));
  const file=folder.createFile(blob);
  const id='ATT-'+Utilities.getUuid();
  appendRow_('AnnouncementAttachments',[id,data.announcementId||'',file.getName(),file.getId(),mime,file.getSize(),file.getUrl(),'https://drive.google.com/file/d/'+file.getId()+'/preview',now_()]);
  audit_(user.username,'UPLOAD','ATTACHMENT',id,file.getName());
  return {status:'SUCCESS',attachment:{id,announcementId:data.announcementId||'',fileName:file.getName(),fileId:file.getId(),mimeType:mime,size:file.getSize(),driveUrl:file.getUrl(),previewUrl:'https://drive.google.com/file/d/'+file.getId()+'/preview'}};
}
function deleteAttachment_(user,fileId,attachmentId){
  if(fileId){try{DriveApp.getFileById(fileId).setTrashed(true);}catch(e){}}
  const a=getRows_('AnnouncementAttachments').find(x=>String(x.ID)===String(attachmentId));
  if(a)deleteRow_('AnnouncementAttachments',a._row);
  audit_(user.username,'DELETE','ATTACHMENT',attachmentId,fileId||'');
  return {status:'SUCCESS'};
}

/* =========================
   IMPORT EMPLOYEES
========================= */

function getEmployees_(filter){
  let list=getRows_('Employees');
  const q=String((filter||{}).q||'').toLowerCase();
  if(q)list=list.filter(e=>[e.EmployeeID,e.EmployeeCode,e.FullName,e.Department,e.Factory,e.Position].join(' ').toLowerCase().includes(q));
  return {status:'SUCCESS',items:list.map(e=>({employeeId:e.EmployeeID,employeeCode:e.EmployeeCode,name:e.FullName,department:e.Department,factory:e.Factory,position:e.Position,status:e.Status}))};
}

function parseImportRows_(rows){
  if(!Array.isArray(rows)||rows.length<2)throw new Error('Excel không có dữ liệu.');
  const headers=rows[0].map(x=>String(x||'').trim());
  const aliases={
    employeeid:['employeeid','employee id','mã nhân viên','ma nhan vien','mã nv','ma nv','code','employee code'],
    name:['fullname','full name','họ tên','ho ten','họ và tên','ho va ten','name'],
    department:['department','phòng ban','phong ban','bộ phận','bo phan','dept'],
    factory:['factory','xưởng','xuong','nhà máy','nha may'],
    position:['position','chức vụ','chuc vu','vị trí','vi tri'],
    email:['email'],phone:['phone','điện thoại','dien thoai']
  };
  const idx={};
  Object.keys(aliases).forEach(k=>idx[k]=headers.findIndex(h=>aliases[k].includes(h.toLowerCase())));
  if(idx.employeeid<0)throw new Error('Không tìm thấy cột Mã nhân viên.');
  return rows.slice(1).map((r,i)=>({
    row:i+2,
    employeeId:String(r[idx.employeeid]??'').trim(),
    name:idx.name>=0?String(r[idx.name]??'').trim():'',
    department:idx.department>=0?String(r[idx.department]??'').trim():'',
    factory:idx.factory>=0?String(r[idx.factory]??'').trim():'',
    position:idx.position>=0?String(r[idx.position]??'').trim():'',
    email:idx.email>=0?String(r[idx.email]??'').trim():'',
    phone:idx.phone>=0?String(r[idx.phone]??'').trim():''
  })).filter(x=>x.employeeId);
}

function importPreview_(user,data){
  const rows=parseImportRows_(data.rows);
  const existingMap={};getRows_('Employees').forEach(e=>existingMap[String(e.EmployeeID).toLowerCase()]=e);
  const seen={}; let added=0,updated=0,unchanged=0,errors=0;
  const items=rows.map(x=>{
    const key=x.employeeId.toLowerCase();
    if(seen[key]){errors++;return {...x,action:'DUPLICATE_IN_FILE'};}
    seen[key]=true;
    const old=existingMap[key];
    if(!old){added++;return {...x,action:'NEW'};}
    const changed=['name','department','factory','position','email','phone'].some(k=>String(old[{name:'FullName',department:'Department',factory:'Factory',position:'Position',email:'Email',phone:'Phone'}[k]]||'')!==String(x[k]||''));
    if(changed)updated++;else unchanged++;
    return {...x,action:changed?'UPDATE':'UNCHANGED'};
  });
  return {status:'SUCCESS',items,summary:{total:rows.length,added,updated,unchanged,errors}};
}

function importConfirm_(user,data){
  const rows=parseImportRows_(data.rows), policy=data.missingPolicy==='INACTIVE'?'INACTIVE':'NO_CHANGE';
  const sh=sheet_('Employees'), existingRows=getRows_('Employees'), map={};existingRows.forEach(e=>map[String(e.EmployeeID).toLowerCase()]=e);
  const seen={};let added=0,updated=0,unchanged=0,errors=0;
  rows.forEach(x=>{
    const key=x.employeeId.toLowerCase(); if(seen[key]){errors++;return;} seen[key]=true;
    const old=map[key], now=new Date();
    const obj={EmployeeID:old?old.EmployeeID:x.employeeId,EmployeeCode:x.employeeId,FullName:x.name,Department:x.department,Factory:x.factory,Position:x.position,Email:x.email,Phone:x.phone,Status:'ACTIVE',CreatedAt:old?old.CreatedAt:now,UpdatedAt:now};
    if(old){updateObjectRow_('Employees',old._row,obj);updated++;}else{appendRow_('Employees',HEADERS.Employees.map(h=>obj[h]));added++;}
  });
  let inactive=0;
  if(policy==='INACTIVE'){
    existingRows.forEach(e=>{if(!seen[String(e.EmployeeID).toLowerCase()] && String(e.Status).toUpperCase()!=='INACTIVE'){sheet_('Employees').getRange(e._row,9).setValue('INACTIVE');sheet_('Employees').getRange(e._row,11).setValue(new Date());inactive++;}});
  }
  const id='IMP-'+Utilities.formatDate(new Date(),Session.getScriptTimeZone()||'Asia/Ho_Chi_Minh','yyyyMMdd-HHmmss')+'-'+Utilities.getUuid().slice(0,6);
  appendRow_('ImportHistory',[id,data.fileName||'',user.username,now_(),rows.length,added,updated,rows.length-added-updated,inactive,errors,policy,'']);
  audit_(user.username,'IMPORT','EMPLOYEES',id,JSON.stringify({fileName:data.fileName,added,updated,inactive,errors}));
  return {status:'SUCCESS',summary:{total:rows.length,added,updated,unchanged:rows.length-added-updated,errors,inactive},importId:id};
}

function getImportHistory_(){return {status:'SUCCESS',items:getRows_('ImportHistory').sort((a,b)=>b._row-a._row).slice(0,100).map(x=>({importId:x.ImportID,fileName:x.FileName,importedBy:x.ImportedBy,importedAt:iso_(x.ImportedAt),total:x.TotalRows,added:x.NewEmployees,updated:x.UpdatedEmployees,inactive:x.InactiveEmployees,errors:x.ErrorRows,policy:x.MissingPolicy}))};}

/* =========================
   AUDIT
========================= */

function audit_(actor,action,entity,id,details){appendRow_('AuditLog',[now_(),actor,action,entity,id,String(details||'').slice(0,5000)]);}
