/**
 * COMPANY INTERNAL NOTICE PORTAL
 * Google Apps Script + Google Sheets
 *
 * Run setupSystem() once before deploying as Web App.
 */

const CONFIG = {
  appName: 'Cổng thông báo nội bộ',
  timezone: Session.getScriptTimeZone() || 'Asia/Ho_Chi_Minh',
  sessionHours: 8,
  maxAttachmentBytes: 25 * 1024 * 1024,
  sheets: {
    employees: 'Employees',
    announcements: 'Announcements',
    reads: 'AnnouncementReads',
    targets: 'AnnouncementTargets',
    imports: 'ImportHistory',
    logins: 'LoginHistory',
    audit: 'AuditLog'
  }
};

const HEADERS = {
  Employees: [
    'EmployeeID','FullName','Department','Position','Email','Phone',
    'PasswordHash','PasswordSalt','Role','Status','CreatedAt','UpdatedAt','LastLoginAt'
  ],
  Announcements: [
    'ID','Title','Content','Type','PublishDate','Author','Priority','Pinned',
    'Status','Attachment','Keywords','CreatedAt','UpdatedAt','ScheduledAt'
  ],
  AnnouncementTargets: [
    'AnnouncementID','TargetType','TargetValue'
  ],
  AnnouncementReads: [
    'AnnouncementID','EmployeeID','ReadAt'
  ],
  ImportHistory: [
    'ImportID','FileName','ImportedBy','ImportedAt','TotalRows','Added',
    'Updated','Skipped','MarkedInactive','MissingAction','Status','Notes'
  ],
  LoginHistory: [
    'LoginID','EmployeeID','LoginAt','Success','IP','UserAgent'
  ],
  AuditLog: [
    'AuditID','At','ActorID','Action','Entity','EntityID','Details'
  ]
};

function doPost(e) {
  try {
    const body = e && e.postData && e.postData.contents
      ? JSON.parse(e.postData.contents) : {};
    const action = String(body.action || '').trim();
    const data = body.data || {};

    const map = {
      login: function(){ return login(body.employeeId, body.password, body.mode, body.userAgent); },
      logout: function(){ return logout(body.token); },
      me: function(){ return me(body.token); },
      dashboard: function(){ return getDashboard(body.token, body.filters); },
      announcement: function(){ return getAnnouncement(body.token, body.id); },
      important: function(){ return getImportant(body.token); },
      markRead: function(){ return markRead(body.token, body.announcementId); },
      readStats: function(){ return getReadStats(body.token, body.id); },
      adminDashboard: function(){ return getAdminDashboard(body.token); },
      employees: function(){ return getEmployees(body.token, body.query); },
      saveEmployee: function(){ return saveEmployee(body.token, data); },
      toggleEmployeeStatus: function(){ return toggleEmployeeStatus(body.token, body.employeeId); },
      deleteEmployee: function(){ return deleteEmployee(body.token, body.employeeId); },
      previewEmployeeImport: function(){ return previewEmployeeImport(body.token, {rows:body.rows || []}); },
      commitEmployeeImport: function(){ return commitEmployeeImport(body.token, {rows:body.rows || [], fileName:body.fileName, missingAction:body.missingAction}); },
      saveAnnouncement: function(){ return saveAnnouncement(body.token, data); },
      deleteAnnouncement: function(){ return deleteAnnouncement(body.token, body.id); },
      togglePinned: function(){ return togglePinned(body.token, body.id); },
      uploadAttachment: function(){ return uploadAttachment(body.token, body.file); }
    };

    if (!map[action]) throw new Error('Action không được hỗ trợ: ' + action);
    const result = map[action]();
    return json_(result);
  } catch (err) {
    return json_({ok:false, error:String(err && err.message ? err.message : err)});
  }
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj || {ok:true}))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet() {
  return HtmlService.createTemplateFromFile('app')
    .evaluate()
    .setTitle(CONFIG.appName)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(name) {
  return HtmlService.createHtmlOutputFromFile(name).getContent();
}

/** Run once. */
function setupSystem() {
  const ss = SpreadsheetApp.getActive();
  PropertiesService.getScriptProperties().setProperty('SPREADSHEET_ID', ss.getId());

  Object.keys(HEADERS).forEach(name => {
    let sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    const headers = HEADERS[name];
    if (sh.getLastRow() === 0) {
      sh.getRange(1, 1, 1, headers.length).setValues([headers]);
      sh.setFrozenRows(1);
    } else {
      const existing = sh.getRange(1,1,1,Math.max(sh.getLastColumn(), headers.length)).getValues()[0];
      headers.forEach((h,i) => {
        if (existing[i] !== h) sh.getRange(1,i+1).setValue(h);
      });
    }
  });

  seedAdminIfNeeded_();
  installTriggers_();

  return {
    ok: true,
    spreadsheetId: ss.getId(),
    message: 'Đã tạo/cập nhật toàn bộ cấu trúc Google Sheets.'
  };
}

function seedAdminIfNeeded_() {
  const sh = getSheet_(CONFIG.sheets.employees);
  const rows = getDataRows_(sh);
  if (rows.length) return;

  const salt = randomToken_(16);
  const hash = hashPassword_('ChangeMe@123', salt);
  sh.appendRow([
    'ADMIN001','System Admin','HR','Administrator','','',
    hash,salt,'ADMIN','ACTIVE',now_(),now_(),''
  ]);
}

function installTriggers_() {
  const ss = SpreadsheetApp.getActive();
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'processScheduledAnnouncements') {
      ScriptApp.deleteTrigger(t);
    }
  });
  ScriptApp.newTrigger('processScheduledAnnouncements')
    .timeBased().everyMinutes(5).create();

  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'onOpen') return;
  });
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📢 Cổng thông báo')
    .addItem('⚙️ Thiết lập hệ thống', 'setupSystem')
    .addItem('⏰ Xử lý thông báo hẹn giờ', 'processScheduledAnnouncements')
    .addToUi();
}

function processScheduledAnnouncements() {
  const sh = getSheet_(CONFIG.sheets.announcements);
  const values = sh.getDataRange().getValues();
  if (values.length < 2) return;
  const h = values[0];
  const idx = indexMap_(h);
  const now = new Date();

  for (let r=1; r<values.length; r++) {
    const status = String(values[r][idx.Status] || '').toUpperCase();
    const scheduled = values[r][idx.ScheduledAt];
    if (status === 'SCHEDULED' && scheduled && new Date(scheduled) <= now) {
      sh.getRange(r+1, idx.Status+1).setValue('PUBLISHED');
      sh.getRange(r+1, idx.UpdatedAt+1).setValue(now);
    }
  }
}

function getPublicConfig() {
  return {appName: CONFIG.appName, timezone: CONFIG.timezone};
}

function login(employeeId, password, mode, userAgent) {
  employeeId = String(employeeId || '').trim().toUpperCase();
  password = String(password || '');
  mode = String(mode || 'employee').toLowerCase();
  if (!employeeId) return {ok:false, message:'Vui lòng nhập EmployeeID / số căn cước.'};

  const emp = findEmployee_(employeeId);
  let success = !!emp && String(emp.Status).toUpperCase() === 'ACTIVE';

  if (success && mode === 'admin') {
    success = ['ADMIN','HR'].includes(String(emp.Role || '').toUpperCase()) &&
      !!emp.PasswordHash &&
      hashPassword_(password, emp.PasswordSalt) === emp.PasswordHash;
  } else if (success && mode !== 'admin') {
    // Nhân viên thường đăng nhập chỉ bằng EmployeeID/CCCD.
    // Tài khoản quản trị không được tự động vào khu vực quản trị bằng chế độ nhân viên.
    if (['ADMIN','HR'].includes(String(emp.Role || '').toUpperCase())) {
      success = false;
    }
  }

  writeLoginHistory_(employeeId, success, userAgent);

  if (!success) {
    return {ok:false, message: mode === 'admin'
      ? 'Tài khoản quản trị không hợp lệ, sai mật khẩu hoặc không hoạt động.'
      : 'EmployeeID / số căn cước không đúng, không tồn tại hoặc tài khoản không hoạt động.'};
  }

  const token = randomToken_(32);
  const payload = {
    employeeId: employeeId,
    role: String(emp.Role || 'EMPLOYEE').toUpperCase(),
    fullName: emp.FullName,
    department: emp.Department,
    position: emp.Position,
    mode: mode === 'admin' ? 'admin' : 'employee',
    createdAt: Date.now()
  };

  CacheService.getScriptCache().put(
    'SESSION_' + token,
    JSON.stringify(payload),
    CONFIG.sessionHours * 3600
  );

  updateEmployeeLastLogin_(employeeId);
  audit_(employeeId, 'LOGIN', 'Employee', employeeId,
    mode === 'admin' ? 'Đăng nhập quản trị thành công' : 'Đăng nhập nhân viên thành công');

  return {
    ok:true,
    token:token,
    user:{
      employeeId:employeeId,
      fullName:emp.FullName,
      department:emp.Department,
      position:emp.Position,
      role:String(emp.Role || 'EMPLOYEE').toUpperCase(),
      mode:payload.mode
    }
  };
}

function logout(token) {
  if (token) CacheService.getScriptCache().remove('SESSION_' + token);
  return {ok:true};
}

function me(token) {
  const session = requireSession_(token);
  return {
    ok:true,
    user: session
  };
}

function getDashboard(token, filters) {
  const session = requireSession_(token);
  filters = filters || {};

  const all = getVisibleAnnouncements_(session.employeeId);
  const unreadIds = getUnreadIds_(session.employeeId, all.map(a=>a.ID));

  const filtered = all.filter(a => {
    if (filters.dateFrom) {
      const from = new Date(String(filters.dateFrom) + 'T00:00:00');
      if (new Date(a.PublishDate || 0) < from) return false;
    }
    if (filters.dateTo) {
      const to = new Date(String(filters.dateTo) + 'T23:59:59');
      if (new Date(a.PublishDate || 0) > to) return false;
    }
    if (filters.type && a.Type !== filters.type) return false;
    if (filters.department && normalize_(getTargetText_(a.ID)).indexOf(normalize_(filters.department)) < 0) return false;
    if (filters.unreadOnly && !unreadIds.has(a.ID)) return false;
    if (filters.search) {
      const q = normalize_(filters.search);
      const hay = normalize_([
        a.ID,a.Title,a.Content,a.Type,a.Keywords,a.TargetText
      ].join(' '));
      if (hay.indexOf(q) < 0) return false;
    }
    return true;
  });

  return {
    ok:true,
    user: session,
    announcements: filtered.map(a => publicAnnouncement_(a, unreadIds.has(a.ID))),
    unreadCount: unreadIds.size,
    countsByMonth: countMonths_(all),
    types: getAnnouncementTypes_(),
    departments: getDepartments_()
  };
}

function getAnnouncement(token, id) {
  const session = requireSession_(token);
  const a = findAnnouncement_(id);
  const isAdmin = session.mode === 'admin' && (session.role === 'ADMIN' || session.role === 'HR');
  if (!a || (!isAdmin && !canViewAnnouncement_(a, session.employeeId))) {
    throw new Error('Bạn không có quyền xem thông báo này.');
  }
  const reads = getReadsForAnnouncement_(id);
  const alreadyRead = reads.some(x => x.EmployeeID === session.employeeId);
  if (!alreadyRead) {
    markReadInternal_(id, session.employeeId);
  }
  return {
    ok:true,
    announcement: publicAnnouncement_(a, true),
    attachment: parseAttachment_(a.Attachment),
    targets: getTargets_(id),
    readAt: findRead_(id, session.employeeId)
  };
}

function getImportant(token) {
  const session = requireSession_(token);
  const all = getVisibleAnnouncements_(session.employeeId)
    .filter(a => a.Pinned || String(a.Priority).toUpperCase() === 'HIGH' || String(a.Priority).toUpperCase() === 'URGENT');
  const unreadIds = getUnreadIds_(session.employeeId, all.map(a=>a.ID));
  return {ok:true, announcements:all.map(a=>publicAnnouncement_(a, unreadIds.has(a.ID)))};
}

function markRead(token, announcementId) {
  const session = requireSession_(token);
  const a = findAnnouncement_(announcementId);
  if (!a || !canViewAnnouncement_(a, session.employeeId)) throw new Error('Không có quyền.');
  markReadInternal_(announcementId, session.employeeId);
  return {ok:true};
}

function getReadStats(token, announcementId) {
  const session = requireSession_(token);
  requireAdmin_(session);
  const a = findAnnouncement_(announcementId);
  if (!a) throw new Error('Không tìm thấy thông báo.');

  const targetEmployees = getTargetEmployeeIds_(announcementId);
  const readIds = new Set(getReadsForAnnouncement_(announcementId).map(x=>x.EmployeeID));
  let read = 0;
  targetEmployees.forEach(id => { if (readIds.has(id)) read++; });

  const total = targetEmployees.length;
  return {
    ok:true,
    announcementId,
    title:a.Title,
    audience:total,
    read,
    unread:Math.max(0,total-read),
    rate:total ? (read/total*100) : 0
  };
}

/* ---------------- ADMIN ---------------- */

function adminListAnnouncements(token, filters) {
  const session = requireSession_(token);
  requireAdmin_(session);
  filters = filters || {};

  let rows = getAnnouncementObjects_();
  if (filters.search) {
    const q = normalize_(filters.search);
    rows = rows.filter(a => normalize_([a.ID,a.Title,a.Content,a.Type,a.Keywords,a.Author].join(' ')).includes(q));
  }
  if (filters.status) rows = rows.filter(a => a.Status === filters.status);
  return {ok:true, announcements:rows.map(publicAdminAnnouncement_)};
}

function saveAnnouncement(token, data) {
  const session = requireSession_(token);
  requireAdmin_(session);
  data = data || {};

  if (!data.title || !data.content || !data.type) throw new Error('Thiếu tiêu đề, nội dung hoặc loại.');

  const sh = getSheet_(CONFIG.sheets.announcements);
  const now = now_();
  const id = data.id || nextAnnouncementId_();
  const existing = findAnnouncement_(id);

  let status = data.status || 'PUBLISHED';
  if (data.scheduledAt) status = 'SCHEDULED';

  const row = [
    id,
    String(data.title).trim(),
    String(data.content),
    String(data.type),
    data.publishDate ? new Date(data.publishDate) : now,
    session.fullName,
    String(data.priority || 'NORMAL'),
    !!data.pinned,
    status,
    data.attachmentJson ? JSON.stringify(data.attachmentJson) : (existing ? String(existing.Attachment || '') : ''),
    String(data.keywords || ''),
    existing ? existing.CreatedAt : now,
    now,
    data.scheduledAt ? new Date(data.scheduledAt) : ''
  ];

  if (existing) {
    sh.getRange(existing._row,1,1,row.length).setValues([row]);
  } else {
    sh.appendRow(row);
  }

  replaceTargets_(id, data.targets || [{type:'ALL',value:'ALL'}]);
  audit_(session.employeeId, existing ? 'UPDATE_ANNOUNCEMENT' : 'CREATE_ANNOUNCEMENT', 'Announcement', id, data.title);

  return {ok:true,id};
}

function deleteAnnouncement(token, id) {
  const session = requireSession_(token);
  requireAdmin_(session);
  const a = findAnnouncement_(id);
  if (!a) throw new Error('Không tìm thấy.');
  getSheet_(CONFIG.sheets.announcements).deleteRow(a._row);
  audit_(session.employeeId, 'DELETE_ANNOUNCEMENT', 'Announcement', id, a.Title);
  return {ok:true};
}

function togglePinned(token, id) {
  const session = requireSession_(token);
  requireAdmin_(session);
  const a = findAnnouncement_(id);
  if (!a) throw new Error('Không tìm thấy.');
  const sh = getSheet_(CONFIG.sheets.announcements);
  sh.getRange(a._row, indexMap_(HEADERS.Announcements).Pinned+1).setValue(!a.Pinned);
  sh.getRange(a._row, indexMap_(HEADERS.Announcements).UpdatedAt+1).setValue(now_());
  audit_(session.employeeId, 'TOGGLE_PIN', 'Announcement', id, String(!a.Pinned));
  return {ok:true};
}

function getAdminDashboard(token) {
  const session = requireSession_(token);
  requireAdmin_(session);

  const announcements = getAnnouncementObjects_();
  const emp = getEmployeeObjects_();
  const reads = getDataRows_(getSheet_(CONFIG.sheets.reads));

  return {
    ok:true,
    stats:{
      employees: emp.filter(e=>e.Status==='ACTIVE').length,
      announcements: announcements.length,
      published: announcements.filter(a=>a.Status==='PUBLISHED').length,
      scheduled: announcements.filter(a=>a.Status==='SCHEDULED').length,
      totalReads: reads.length
    },
    announcements: announcements.slice(0,100).map(publicAdminAnnouncement_)
  };
}

function changeOwnPassword(token, oldPassword, newPassword) {
  const session = requireSession_(token);
  if (!newPassword || String(newPassword).length < 8) throw new Error('Mật khẩu mới tối thiểu 8 ký tự.');

  const emp = findEmployee_(session.employeeId);
  if (!emp || hashPassword_(oldPassword, emp.PasswordSalt) !== emp.PasswordHash) {
    throw new Error('Mật khẩu cũ không đúng.');
  }
  const salt = randomToken_(16);
  const hash = hashPassword_(newPassword, salt);
  const sh = getSheet_(CONFIG.sheets.employees);
  sh.getRange(emp._row, indexMap_(HEADERS.Employees).PasswordHash+1).setValue(hash);
  sh.getRange(emp._row, indexMap_(HEADERS.Employees).PasswordSalt+1).setValue(salt);
  sh.getRange(emp._row, indexMap_(HEADERS.Employees).UpdatedAt+1).setValue(now_());
  audit_(session.employeeId, 'CHANGE_PASSWORD', 'Employee', session.employeeId, '');
  return {ok:true};
}

/* ---------------- EMPLOYEE ADMIN ---------------- */
function saveEmployee(token, data) {
  const session=requireSession_(token);
  requireAdmin_(session);
  data=data||{};
  const id=String(data.employeeId||'').trim().toUpperCase();
  const name=String(data.fullName||'').trim();
  if(!id || !name) throw new Error('EmployeeID và họ tên là bắt buộc.');
  if(!/^\d+$/.test(id)) throw new Error('EmployeeID / số căn cước chỉ được chứa chữ số.');

  const existing=findEmployee_(id);
  const role=['EMPLOYEE','HR','ADMIN'].includes(String(data.role||'EMPLOYEE').toUpperCase())
    ? String(data.role||'EMPLOYEE').toUpperCase() : 'EMPLOYEE';
  const status=String(data.status||'ACTIVE').toUpperCase()==='INACTIVE'?'INACTIVE':'ACTIVE';
  const sh=getSheet_(CONFIG.sheets.employees);
  const now=now_();

  let passwordHash=existing ? existing.PasswordHash : '';
  let passwordSalt=existing ? existing.PasswordSalt : '';
  if ((role==='ADMIN' || role==='HR') && String(data.adminPassword||'').trim()) {
    if (String(data.adminPassword).length < 8) throw new Error('Mật khẩu quản trị tối thiểu 8 ký tự.');
    passwordSalt=randomToken_(16);
    passwordHash=hashPassword_(String(data.adminPassword),passwordSalt);
  }
  if(existing) {
    const row=[
      id,name,String(data.department||''),String(data.position||''),String(data.email||''),
      String(data.phone||''),passwordHash,passwordSalt,role,status,
      existing.CreatedAt||now,now,existing.LastLoginAt||''
    ];
    sh.getRange(existing._row,1,1,row.length).setValues([row]);
    audit_(session.employeeId,'UPDATE_EMPLOYEE','Employee',id,JSON.stringify({role,status}));
    return {ok:true,message:'Đã cập nhật nhân viên.'};
  }

  // Nhân viên mới không cần mật khẩu. Nếu là ADMIN/HR, tạo mật khẩu quản trị mặc định
  // để admin có thể đăng nhập chế độ quản trị và sau đó cập nhật mật khẩu.
  let salt='',hash='';
  if(role==='ADMIN' || role==='HR'){
    salt=randomToken_(16);hash=hashPassword_('ChangeMe@123',salt);
  }
  sh.appendRow([id,name,String(data.department||''),String(data.position||''),String(data.email||''),
    String(data.phone||''),hash,salt,role,status,now,now,'']);
  audit_(session.employeeId,'CREATE_EMPLOYEE','Employee',id,JSON.stringify({role,status}));
  return {ok:true,message:'Đã thêm nhân viên.'};
}

function toggleEmployeeStatus(token, employeeId) {
  const session=requireSession_(token);requireAdmin_(session);
  const id=String(employeeId||'').trim().toUpperCase();
  if(id===session.employeeId) throw new Error('Không thể tự khóa tài khoản đang đăng nhập.');
  const e=findEmployee_(id);if(!e)throw new Error('Không tìm thấy nhân viên.');
  const next=String(e.Status).toUpperCase()==='ACTIVE'?'INACTIVE':'ACTIVE';
  const sh=getSheet_(CONFIG.sheets.employees),m=indexMap_(HEADERS.Employees);
  sh.getRange(e._row,m.Status+1).setValue(next);sh.getRange(e._row,m.UpdatedAt+1).setValue(now_());
  audit_(session.employeeId,'TOGGLE_EMPLOYEE_STATUS','Employee',id,next);
  return {ok:true,status:next};
}

function deleteEmployee(token, employeeId) {
  const session=requireSession_(token);requireAdmin_(session);
  const id=String(employeeId||'').trim().toUpperCase();
  if(id===session.employeeId) throw new Error('Không thể tự xóa tài khoản đang đăng nhập.');
  const e=findEmployee_(id);if(!e)throw new Error('Không tìm thấy nhân viên.');
  if(['ADMIN','HR'].includes(String(e.Role).toUpperCase())) {
    const admins=getEmployeeObjects_().filter(x=>String(x.Status).toUpperCase()==='ACTIVE' && ['ADMIN','HR'].includes(String(x.Role).toUpperCase()));
    if(admins.length<=1) throw new Error('Không thể xóa quản trị viên cuối cùng.');
  }
  getSheet_(CONFIG.sheets.employees).deleteRow(e._row);
  audit_(session.employeeId,'DELETE_EMPLOYEE','Employee',id,e.FullName);
  return {ok:true};
}

/* ---------------- IMPORT EXCEL ----------------
 * Client reads XLSX using SheetJS and sends rows as JSON.
 * Required column: EmployeeID (or common Vietnamese variants).
 */

function previewEmployeeImport(token, payload) {
  const session = requireSession_(token);
  requireAdmin_(session);
  const rows = normalizeEmployeeImportRows_(payload.rows || []);
  const existing = {};
  getEmployeeObjects_().forEach(e => existing[e.EmployeeID] = e);

  const preview = rows.map(r => {
    const old = existing[r.EmployeeID];
    return {
      ...r,
      action: old ? 'UPDATE' : 'ADD',
      oldName: old ? old.FullName : '',
      oldDepartment: old ? old.Department : ''
    };
  });

  const ids = new Set();
  const duplicates = [];
  preview.forEach(r=>{
    if(ids.has(r.EmployeeID)) duplicates.push(r.EmployeeID);
    ids.add(r.EmployeeID);
  });

  return {
    ok:true,
    total:preview.length,
    duplicates:[...new Set(duplicates)],
    rows:preview,
    existingCount:Object.keys(existing).length
  };
}

function commitEmployeeImport(token, payload) {
  const session = requireSession_(token);
  requireAdmin_(session);

  const rows = normalizeEmployeeImportRows_(payload.rows || []);
  const missingAction = payload.missingAction === 'INACTIVE' ? 'INACTIVE' : 'NO_CHANGE';
  const fileName = String(payload.fileName || 'employee_import.xlsx');

  const sh = getSheet_(CONFIG.sheets.employees);
  const data = getDataRows_(sh);
  const map = {};
  data.forEach((r,i)=>{
    const obj = rowToObject_(HEADERS.Employees,r);
    obj._row = i+2;
    map[obj.EmployeeID] = obj;
  });

  let added=0, updated=0, skipped=0, markedInactive=0;
  const incoming = new Set();

  rows.forEach(r=>{
    if (!r.EmployeeID) { skipped++; return; }
    incoming.add(r.EmployeeID);
    const old = map[r.EmployeeID];

    if (!old) {
      const salt = randomToken_(16);
      const initialPassword = r.Password || r.EmployeeID;
      const hash = hashPassword_(initialPassword, salt);
      sh.appendRow([
        r.EmployeeID,r.FullName,r.Department,r.Position,r.Email,r.Phone,
        hash,salt,r.Role || 'EMPLOYEE','ACTIVE',now_(),now_(),''
      ]);
      added++;
    } else {
      const row = [
        r.EmployeeID,
        r.FullName || old.FullName,
        r.Department || old.Department,
        r.Position || old.Position,
        r.Email || old.Email,
        r.Phone || old.Phone,
        old.PasswordHash,
        old.PasswordSalt,
        r.Role || old.Role || 'EMPLOYEE',
        'ACTIVE',
        old.CreatedAt || now_(),
        now_(),
        old.LastLoginAt || ''
      ];
      sh.getRange(old._row,1,1,row.length).setValues([row]);
      updated++;
    }
  });

  if (missingAction === 'INACTIVE') {
    Object.keys(map).forEach(id=>{
      if (!incoming.has(id) && map[id].Status === 'ACTIVE') {
        sh.getRange(map[id]._row, indexMap_(HEADERS.Employees).Status+1).setValue('INACTIVE');
        sh.getRange(map[id]._row, indexMap_(HEADERS.Employees).UpdatedAt+1).setValue(now_());
        markedInactive++;
      }
    });
  }

  const importId = 'IMP-' + Utilities.getUuid().slice(0,8).toUpperCase();
  getSheet_(CONFIG.sheets.imports).appendRow([
    importId,fileName,session.employeeId,now_(),rows.length,added,updated,skipped,
    markedInactive,missingAction,'SUCCESS',''
  ]);

  audit_(session.employeeId,'IMPORT_EMPLOYEES','EmployeeImport',importId,
    JSON.stringify({fileName,total:rows.length,added,updated,skipped,markedInactive,missingAction}));

  return {ok:true,importId,added,updated,skipped,markedInactive};
}

function getEmployees(token, query) {
  const session = requireSession_(token);
  requireAdmin_(session);
  const q = normalize_(query || '');
  let rows = getEmployeeObjects_();
  if (q) rows = rows.filter(e => normalize_([
    e.EmployeeID,e.FullName,e.Department,e.Position,e.Email
  ].join(' ')).includes(q));
  return {ok:true, employees:rows.slice(0,500)};
}

function uploadAttachment(token, file) {
  const session=requireSession_(token);requireAdmin_(session);
  file=file||{};
  const name=String(file.name||'').trim();
  const mime=String(file.mimeType||'application/octet-stream');
  const b64=String(file.base64||'');
  if(!name || !b64) throw new Error('Tệp không hợp lệ.');
  const bytes=Utilities.base64Decode(b64);
  if(bytes.length>CONFIG.maxAttachmentBytes) throw new Error('File tối đa 25MB.');

  const folderId=PropertiesService.getScriptProperties().getProperty('ATTACHMENT_FOLDER_ID');
  let folder;
  if(folderId){
    folder=DriveApp.getFolderById(folderId);
  }else{
    folder=DriveApp.createFolder('Internal Notice Attachments');
    PropertiesService.getScriptProperties().setProperty('ATTACHMENT_FOLDER_ID',folder.getId());
  }
  const blob=Utilities.newBlob(bytes,mime,name);
  const created=folder.createFile(blob);
  const id=created.getId();
  const url='https://drive.google.com/file/d/'+id+'/view';
  audit_(session.employeeId,'UPLOAD_ATTACHMENT','File',id,name);
  return {ok:true,attachment:{id,name,mimeType:mime,size:bytes.length,url:url,previewUrl:url}};
}

/* ---------------- HELPERS ---------------- */

function getSS_() {
  const id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  return SpreadsheetApp.openById(id || SpreadsheetApp.getActive().getId());
}

function getSheet_(name) {
  const sh = getSS_().getSheetByName(name);
  if (!sh) throw new Error('Thiếu Sheet ' + name + '. Hãy chạy setupSystem().');
  return sh;
}

function getDataRows_(sh) {
  if (sh.getLastRow() < 2) return [];
  return sh.getRange(2,1,sh.getLastRow()-1,sh.getLastColumn()).getValues();
}

function indexMap_(headers) {
  const m={}; headers.forEach((h,i)=>m[h]=i); return m;
}

function rowToObject_(headers,row) {
  const o={}; headers.forEach((h,i)=>o[h]=row[i]); return o;
}

function getEmployeeObjects_() {
  return getDataRows_(getSheet_(CONFIG.sheets.employees))
    .map((r,i)=>{const o=rowToObject_(HEADERS.Employees,r);o._row=i+2;return o;});
}

function findEmployee_(id) {
  return getEmployeeObjects_().find(e=>String(e.EmployeeID).toUpperCase()===id) || null;
}

function updateEmployeeLastLogin_(id) {
  const e=findEmployee_(id); if(!e) return;
  const sh=getSheet_(CONFIG.sheets.employees);
  sh.getRange(e._row,indexMap_(HEADERS.Employees).LastLoginAt+1).setValue(now_());
}

function getAnnouncementObjects_() {
  const sh=getSheet_(CONFIG.sheets.announcements);
  return getDataRows_(sh).map((r,i)=>{
    const o=rowToObject_(HEADERS.Announcements,r); o._row=i+2;
    o.Pinned = o.Pinned === true || String(o.Pinned).toUpperCase()==='TRUE';
    return o;
  });
}

function findAnnouncement_(id) {
  return getAnnouncementObjects_().find(a=>String(a.ID)===String(id)) || null;
}

function nextAnnouncementId_() {
  return 'TB-' + Utilities.formatDate(new Date(), CONFIG.timezone, 'yyyyMMdd') + '-' +
    Utilities.getUuid().slice(0,6).toUpperCase();
}

function publicAnnouncement_(a, unread) {
  return {
    id:a.ID,title:a.Title,content:a.Content,type:a.Type,publishDate:dateOut_(a.PublishDate),
    author:a.Author,priority:a.Priority,pinned:a.Pinned,status:a.Status,
    attachment:parseAttachment_(a.Attachment),keywords:a.Keywords,
    unread:!!unread,targetText:getTargetText_(a.ID)
  };
}

function publicAdminAnnouncement_(a) {
  return {...publicAnnouncement_(a,false), scheduledAt:dateOut_(a.ScheduledAt),createdAt:dateOut_(a.CreatedAt),targets:getTargets_(a.ID)};
}

function getVisibleAnnouncements_(employeeId) {
  return getAnnouncementObjects_()
    .filter(a => a.Status === 'PUBLISHED' && new Date(a.PublishDate || 0) <= new Date())
    .filter(a => canViewAnnouncement_(a, employeeId))
    .sort((a,b)=>{
      if (a.Pinned !== b.Pinned) return a.Pinned ? -1 : 1;
      const da=new Date(a.PublishDate||0).getTime(), db=new Date(b.PublishDate||0).getTime();
      return db-da;
    });
}

function canViewAnnouncement_(a, employeeId) {
  const rows=getDataRows_(getSheet_(CONFIG.sheets.targets));
  const targets=rows.filter(r=>String(r[0])===String(a.ID));
  if (!targets.length) return true;
  return targets.some(t=>{
    const type=String(t[1]||'ALL').toUpperCase();
    const value=String(t[2]||'ALL');
    if(type==='ALL') return true;
    const e=findEmployee_(employeeId);
    if(!e) return false;
    if(type==='EMPLOYEE') return value.toUpperCase()===employeeId.toUpperCase();
    if(type==='DEPARTMENT') return normalize_(e.Department)===normalize_(value);
    if(type==='TYPE') return normalize_(e.Role)===normalize_(value) || normalize_(e.Position)===normalize_(value);
    return false;
  });
}

function isTargetVisible_(announcementId, employeeId, department) {
  const a=findAnnouncement_(announcementId);
  if(!a) return false;
  const e=findEmployee_(employeeId);
  if(!e) return false;
  if(normalize_(e.Department)!==normalize_(department)) return false;
  return canViewAnnouncement_(a, employeeId);
}

function replaceTargets_(id, targets) {
  const sh=getSheet_(CONFIG.sheets.targets);
  const rows=getDataRows_(sh);
  for(let i=rows.length-1;i>=0;i--) {
    if(String(rows[i][0])===String(id)) sh.deleteRow(i+2);
  }
  (targets.length ? targets : [{type:'ALL',value:'ALL'}]).forEach(t=>{
    sh.appendRow([id,String(t.type||'ALL').toUpperCase(),String(t.value||'ALL')]);
  });
}

function getTargets_(id) {
  return getDataRows_(getSheet_(CONFIG.sheets.targets))
    .filter(r=>String(r[0])===String(id))
    .map(r=>({type:String(r[1]||'ALL').toUpperCase(),value:String(r[2]||'ALL')}));
}

function getTargetText_(id) {
  const rows=getDataRows_(getSheet_(CONFIG.sheets.targets)).filter(r=>String(r[0])===String(id));
  if(!rows.length) return 'Tất cả';
  return rows.map(r=>`${r[1]}: ${r[2]}`).join(', ');
}

function getTargetEmployeeIds_(id) {
  const emps=getEmployeeObjects_().filter(e=>e.Status==='ACTIVE');
  return emps.filter(e=>canViewAnnouncement_(findAnnouncement_(id),e.EmployeeID)).map(e=>e.EmployeeID);
}

function getReadsForAnnouncement_(id) {
  return getDataRows_(getSheet_(CONFIG.sheets.reads))
    .filter(r=>String(r[0])===String(id))
    .map(r=>({AnnouncementID:r[0],EmployeeID:r[1],ReadAt:r[2]}));
}

function findRead_(id, employeeId) {
  const x=getReadsForAnnouncement_(id).find(r=>String(r.EmployeeID)===String(employeeId));
  return x ? dateOut_(x.ReadAt) : null;
}

function markReadInternal_(id, employeeId) {
  if(findRead_(id,employeeId)) return;
  getSheet_(CONFIG.sheets.reads).appendRow([id,employeeId,now_()]);
}

function getUnreadIds_(employeeId, ids) {
  const set=new Set();
  getDataRows_(getSheet_(CONFIG.sheets.reads)).forEach(r=>{
    if(String(r[1])===String(employeeId)) set.add(String(r[0]));
  });
  return new Set(ids.filter(id=>!set.has(String(id))));
}

function getAnnouncementTypes_() {
  return [
    '📢 Thông báo chung','🕐 Thời gian làm việc','💰 Lương & thưởng','🎁 Phúc lợi',
    '🏥 Sức khỏe','🦺 An toàn / HSE','🏭 Sản xuất','📋 Nhân sự',
    '🎉 Hoạt động công ty','⚠️ Khẩn cấp'
  ];
}

function getDepartments_() {
  return [...new Set(getEmployeeObjects_().map(e=>String(e.Department||'').trim()).filter(Boolean))].sort();
}

function countMonths_(rows) {
  const o={};
  rows.forEach(a=>{
    if(!a.PublishDate) return;
    const k=Utilities.formatDate(new Date(a.PublishDate),CONFIG.timezone,'MM/yyyy');
    o[k]=(o[k]||0)+1;
  });
  return o;
}

function sameMonth_(date, month) {
  if(!date || !month) return true;
  const d=Utilities.formatDate(new Date(date),CONFIG.timezone,'MM/yyyy');
  return d===month;
}

function parseAttachment_(v) {
  if(!v) return null;
  try { return typeof v === 'string' ? JSON.parse(v) : v; } catch(e) { return {name:String(v)}; }
}

function requireSession_(token) {
  if(!token) throw new Error('Phiên đăng nhập đã hết hạn.');
  const raw=CacheService.getScriptCache().get('SESSION_'+token);
  if(!raw) throw new Error('Phiên đăng nhập đã hết hạn.');
  const session=JSON.parse(raw);
  const emp=findEmployee_(String(session.employeeId || '').toUpperCase());
  if(!emp || String(emp.Status).toUpperCase() !== 'ACTIVE') {
    CacheService.getScriptCache().remove('SESSION_'+token);
    throw new Error('Tài khoản không còn hoạt động.');
  }
  // Quyền luôn lấy lại từ Sheets để không tin dữ liệu cũ trong trình duyệt/cache.
  session.role=String(emp.Role || 'EMPLOYEE').toUpperCase();
  session.fullName=emp.FullName;
  session.department=emp.Department;
  session.position=emp.Position;
  return session;
}

function requireAdmin_(session) {
  if(session.mode !== 'admin' || !['ADMIN','HR'].includes(String(session.role).toUpperCase())) {
    throw new Error('Không có quyền quản trị.');
  }
}

function hashPassword_(password,salt) {
  const bytes=Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    String(salt)+'|'+String(password),
    Utilities.Charset.UTF_8
  );
  return bytes.map(b=>(b<0?b+256:b).toString(16).padStart(2,'0')).join('');
}

function randomToken_(bytes) {
  return Utilities.getUuid().replace(/-/g,'') + Utilities.getUuid().replace(/-/g,'').slice(0,Math.max(0,bytes*2-32));
}

function normalize_(v) {
  return String(v||'').toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g,'').trim();
}

function dateOut_(d) {
  if(!d) return '';
  const dt=new Date(d);
  if(isNaN(dt)) return '';
  return Utilities.formatDate(dt,CONFIG.timezone,"yyyy-MM-dd'T'HH:mm:ss");
}

function now_(){ return new Date(); }

function writeLoginHistory_(employeeId, success, userAgent) {
  const sh=getSheet_(CONFIG.sheets.logins);
  sh.appendRow([
    'LOGIN-'+Utilities.getUuid().slice(0,8),employeeId,now_(),success,'',String(userAgent || '')
  ]);
}

function audit_(actor,action,entity,entityId,details) {
  getSheet_(CONFIG.sheets.audit).appendRow([
    'AUD-'+Utilities.getUuid().slice(0,8),now_(),actor,action,entity,entityId,String(details||'')
  ]);
}

function normalizeEmployeeImportRows_(rows) {
  if(!Array.isArray(rows)) throw new Error('Dữ liệu Excel không hợp lệ.');
  return rows.map(r=>{
    const keys=Object.keys(r);
    const pick=(names)=>{
      for(const n of names){
        const k=keys.find(x=>normalize_(x)===normalize_(n));
        if(k!==undefined) return r[k];
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
  }).filter(r=>r.EmployeeID);
}