/*******************************************************
 * COMPANY INTERNAL PORTAL
 * Google Apps Script Backend
 *
 * Sheets:
 * Employees
 * Users
 * Announcements
 * AnnouncementReads
 * ImportHistory
 * AuditLog
 *******************************************************/

const CONFIG = {
  APP_NAME: 'COMPANY - CỔNG THÔNG BÁO NỘI BỘ',

  SESSION_TTL: 21600, // 6 hours

  MAX_ATTACHMENT_BYTES: 25 * 1024 * 1024,

  DRIVE_FOLDER_NAME: 'COMPANY_INTERNAL_PORTAL_ATTACHMENTS',

  PASSWORD_MIN_LENGTH: 6,

  ANNOUNCEMENT_TYPES: {
    GENERAL: {
      label: 'Thông báo chung',
      icon: '📢'
    },
    WORK_TIME: {
      label: 'Thời gian làm việc',
      icon: '🕐'
    },
    SALARY: {
      label: 'Lương & thưởng',
      icon: '💰'
    },
    BENEFIT: {
      label: 'Phúc lợi',
      icon: '🎁'
    },
    HEALTH: {
      label: 'Sức khỏe',
      icon: '🏥'
    },
    HSE: {
      label: 'An toàn / HSE',
      icon: '🦺'
    },
    PRODUCTION: {
      label: 'Sản xuất',
      icon: '🏭'
    },
    HR: {
      label: 'Nhân sự',
      icon: '📋'
    },
    EVENT: {
      label: 'Hoạt động công ty',
      icon: '🎉'
    },
    URGENT: {
      label: 'Khẩn cấp',
      icon: '⚠️'
    }
  }
};


/* =====================================================
 * SHEET DEFINITIONS
 * ===================================================== */

const SHEETS = {

  Employees: [
    'EmployeeID',
    'FullName',
    'Department',
    'Position',
    'Email',
    'Phone',
    'Status',
    'CreatedAt',
    'UpdatedAt',
    'LastLogin'
  ],

  Users: [
    'EmployeeID',
    'PasswordHash',
    'Role',
    'Status',
    'CreatedAt',
    'UpdatedAt'
  ],

  Announcements: [
    'ID',
    'Title',
    'Content',
    'Type',
    'PublishDate',
    'Author',
    'Priority',
    'Pinned',
    'Status',
    'Attachment',
    'Keywords',

    // Targeting
    'TargetType',
    'TargetValue',

    // Schedule
    'ScheduledAt',

    // Audit
    'CreatedAt',
    'UpdatedAt'
  ],

  AnnouncementReads: [
    'AnnouncementID',
    'EmployeeID',
    'ReadAt'
  ],

  ImportHistory: [
    'ImportID',
    'FileName',
    'ImportedAt',
    'ImportedBy',
    'TotalRows',
    'Added',
    'Updated',
    'Skipped',
    'InactiveMarked',
    'Mode',
    'Status'
  ],

  AuditLog: [
    'ID',
    'Time',
    'EmployeeID',
    'Action',
    'Entity',
    'EntityID',
    'Details',
    'IPAddress'
  ]
};


/* =====================================================
 * WEB APP
 * ===================================================== */

function doGet() {

  return HtmlService
    .createTemplateFromFile('Index')
    .evaluate()
    .setTitle(CONFIG.APP_NAME)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}


function include(filename) {
  return HtmlService
    .createHtmlOutputFromFile(filename)
    .getContent();
}


/* =====================================================
 * INITIAL SETUP
 * ===================================================== */

function setupSystem(adminPassword) {

  if (!adminPassword) {
    throw new Error(
      'Hãy chạy setupSystem("MatKhauAdminMoi")'
    );
  }

  if (String(adminPassword).length <
      CONFIG.PASSWORD_MIN_LENGTH) {

    throw new Error(
      'Mật khẩu Admin quá ngắn.'
    );
  }

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  Object.keys(SHEETS).forEach(name => {

    let sheet = ss.getSheetByName(name);

    if (!sheet) {
      sheet = ss.insertSheet(name);
    }

    const headers = SHEETS[name];

    if (sheet.getLastRow() === 0) {

      sheet
        .getRange(
          1,
          1,
          1,
          headers.length
        )
        .setValues([headers]);

      sheet.setFrozenRows(1);

      sheet
        .getRange(
          1,
          1,
          1,
          headers.length
        )
        .setFontWeight('bold');
    }
  });


  // Tạo admin nếu chưa có
  const users =
    getSheet_('Users');

  const values =
    users.getDataRange().getValues();

  const adminExists =
    values.slice(1).some(
      r => String(r[0]).toUpperCase() === 'ADMIN'
    );

  if (!adminExists) {

    users.appendRow([
      'ADMIN',
      hashPassword(adminPassword),
      'ADMIN',
      'ACTIVE',
      new Date(),
      new Date()
    ]);
  }


  // Drive folder
  getAttachmentFolder_();

  return {
    success: true,
    message:
      'Khởi tạo hệ thống thành công.'
  };
}


/* =====================================================
 * GENERIC SHEET FUNCTIONS
 * ===================================================== */

function getSheet_(name) {

  const sheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName(name);

  if (!sheet) {
    throw new Error(
      'Không tìm thấy Sheet: ' + name
    );
  }

  return sheet;
}


function getRows_(sheetName) {

  const sheet =
    getSheet_(sheetName);

  const values =
    sheet.getDataRange().getValues();

  if (!values.length) {
    return [];
  }

  const headers =
    values[0].map(String);

  return values.slice(1).map(row => {

    const obj = {};

    headers.forEach((h, i) => {
      obj[h] = row[i];
    });

    return obj;
  });
}


function rowToObject_(headers, row) {

  const obj = {};

  headers.forEach((h, i) => {
    obj[h] = row[i];
  });

  return obj;
}


function objectToRow_(headers, obj) {

  return headers.map(
    h => obj[h] !== undefined
      ? obj[h]
      : ''
  );
}


/* =====================================================
 * PASSWORD
 * ===================================================== */

function hashPassword(password) {

  const bytes =
    Utilities.computeDigest(
      Utilities.DigestAlgorithm.SHA_256,
      String(password)
    );

  return bytes
    .map(b =>
      ('0' + (b & 0xFF).toString(16))
        .slice(-2)
    )
    .join('');
}


/* =====================================================
 * SESSION
 * ===================================================== */

function createSession_(employeeId, role) {

  const token =
    Utilities.getUuid() +
    Utilities.getUuid();

  const data = {
    employeeId: employeeId,
    role: role,
    createdAt: Date.now()
  };

  CacheService
    .getScriptCache()
    .put(
      'SESSION_' + token,
      JSON.stringify(data),
      CONFIG.SESSION_TTL
    );

  return token;
}


function getSession_(token) {

  if (!token) {
    throw new Error('Phiên đăng nhập không hợp lệ.');
  }

  const cache =
    CacheService.getScriptCache();

  const raw =
    cache.get('SESSION_' + token);

  if (!raw) {
    throw new Error(
      'Phiên đăng nhập đã hết hạn. Vui lòng đăng nhập lại.'
    );
  }

  return JSON.parse(raw);
}


function requireLogin_(token) {
  return getSession_(token);
}


function requireRole_(token, roles) {

  const session =
    requireLogin_(token);

  if (
    !roles.includes(
      String(session.role).toUpperCase()
    )
  ) {
    throw new Error(
      'Bạn không có quyền thực hiện thao tác này.'
    );
  }

  return session;
}


/* =====================================================
 * LOGIN
 * ===================================================== */

function login(employeeId, password) {

  employeeId =
    String(employeeId || '')
      .trim()
      .toUpperCase();

  password =
    String(password || '');

  if (!employeeId || !password) {
    throw new Error(
      'Vui lòng nhập mã nhân viên và mật khẩu.'
    );
  }

  const users =
    getRows_('Users');

  const user =
    users.find(
      u =>
        String(u.EmployeeID)
          .toUpperCase() === employeeId
    );

  if (!user) {
    throw new Error(
      'Mã nhân viên hoặc mật khẩu không đúng.'
    );
  }

  if (
    String(user.Status).toUpperCase()
    !== 'ACTIVE'
  ) {
    throw new Error(
      'Tài khoản đã bị khóa.'
    );
  }

  if (
    String(user.PasswordHash)
      !== hashPassword(password)
  ) {
    throw new Error(
      'Mã nhân viên hoặc mật khẩu không đúng.'
    );
  }

  const employees =
    getRows_('Employees');

  const employee =
    employees.find(
      e =>
        String(e.EmployeeID)
          .toUpperCase() === employeeId
    );

  if (!employee) {
    throw new Error(
      'Không tìm thấy thông tin nhân viên.'
    );
  }

  if (
    String(employee.Status).toUpperCase()
    !== 'ACTIVE'
  ) {
    throw new Error(
      'Nhân viên đang ở trạng thái INACTIVE.'
    );
  }

  const role =
    String(user.Role || 'USER')
      .toUpperCase();

  const token =
    createSession_(
      employeeId,
      role
    );

  updateLastLogin_(employeeId);

  audit_(
    employeeId,
    'LOGIN',
    'USER',
    employeeId,
    'Đăng nhập hệ thống'
  );

  return {
    success: true,
    token: token,
    user: {
      employeeId: employeeId,
      fullName: employee.FullName,
      department: employee.Department,
      position: employee.Position,
      role: role
    }
  };
}


function logout(token) {

  const session =
    getSession_(token);

  CacheService
    .getScriptCache()
    .remove('SESSION_' + token);

  audit_(
    session.employeeId,
    'LOGOUT',
    'USER',
    session.employeeId,
    'Đăng xuất'
  );

  return {
    success: true
  };
}


function updateLastLogin_(employeeId) {

  const sheet =
    getSheet_('Employees');

  const data =
    sheet.getDataRange().getValues();

  const headers =
    data[0];

  const idCol =
    headers.indexOf('EmployeeID');

  const loginCol =
    headers.indexOf('LastLogin');

  if (
    idCol < 0 ||
    loginCol < 0
  ) return;

  for (let i = 1; i < data.length; i++) {

    if (
      String(data[i][idCol])
        .toUpperCase()
        === employeeId
    ) {

      sheet
        .getRange(
          i + 1,
          loginCol + 1
        )
        .setValue(new Date());

      break;
    }
  }
}


/* =====================================================
 * EMPLOYEE PROFILE
 * ===================================================== */

function getCurrentUser(token) {

  const session =
    requireLogin_(token);

  const employees =
    getRows_('Employees');

  const employee =
    employees.find(
      e =>
        String(e.EmployeeID)
          .toUpperCase()
        === String(session.employeeId)
          .toUpperCase()
    );

  if (!employee) {
    throw new Error(
      'Không tìm thấy nhân viên.'
    );
  }

  return {
    employee: employee,
    role: session.role
  };
}


/* =====================================================
 * ANNOUNCEMENTS
 * ===================================================== */

function getAnnouncements(token, filters) {

  const session =
    requireLogin_(token);

  filters =
    filters || {};

  const employeeId =
    String(session.employeeId)
      .toUpperCase();

  const employees =
    getRows_('Employees');

  const employee =
    employees.find(
      e =>
        String(e.EmployeeID)
          .toUpperCase()
        === employeeId
    );

  if (!employee) {
    throw new Error(
      'Không tìm thấy nhân viên.'
    );
  }

  const department =
    normalize_(employee.Department);

  const announcements =
    getRows_('Announcements');

  const reads =
    getRows_('AnnouncementReads');


  const readSet =
    new Set(
      reads
        .filter(
          r =>
            String(r.EmployeeID)
              .toUpperCase()
            === employeeId
        )
        .map(r =>
          String(r.AnnouncementID)
        )
    );


  const now =
    new Date();

  let result =
    announcements
      .filter(a => {

        const status =
          String(a.Status)
            .toUpperCase();

        if (status !== 'PUBLISHED') {
          return false;
        }

        const publishDate =
          parseDate_(a.PublishDate);

        if (
          publishDate &&
          publishDate > now
        ) {
          return false;
        }

        const scheduled =
          parseDate_(a.ScheduledAt);

        if (
          scheduled &&
          scheduled > now
        ) {
          return false;
        }

        return canEmployeeSeeAnnouncement_(
          a,
          employee,
          department
        );
      })
      .map(a => {

        const id =
          String(a.ID);

        return {
          id: id,
          title: a.Title,
          content: a.Content,
          type: a.Type,
          typeLabel:
            getTypeLabel_(a.Type),
          typeIcon:
            getTypeIcon_(a.Type),
          publishDate:
            formatDateTime_(a.PublishDate),
          author: a.Author,
          priority: a.Priority,
          pinned:
            String(a.Pinned)
              .toUpperCase() === 'TRUE',
          attachment: a.Attachment,
          keywords: a.Keywords,
          read: readSet.has(id)
        };
      });


  // Search
  const search =
    normalize_(filters.search);

  if (search) {

    result =
      result.filter(a => {

        const text =
          normalize_([
            a.title,
            a.content,
            a.keywords,
            a.type,
            a.typeLabel,
            employee.Department,
            a.id
          ].join(' '));

        return text.includes(search);
      });
  }


  // Type
  if (filters.type) {

    result =
      result.filter(
        a =>
          String(a.type)
            === String(filters.type)
      );
  }


  // Month
  if (filters.month) {

    const month =
      String(filters.month);

    result =
      result.filter(a => {

        const d =
          parseDate_(a.publishDate);

        if (!d) return false;

        const m =
          String(
            d.getMonth() + 1
          ).padStart(2, '0');

        const y =
          String(d.getFullYear());

        return (
          y + '-' + m === month
        );
      });
  }


  // Unread only
  if (
    filters.unread === true ||
    String(filters.unread) === 'true'
  ) {

    result =
      result.filter(
        a => !a.read
      );
  }


  // Latest first
  result.sort(
    (a, b) => {

      if (
        a.pinned !== b.pinned
      ) {
        return a.pinned
          ? -1
          : 1;
      }

      const da =
        parseDate_(a.publishDate)
          || new Date(0);

      const db =
        parseDate_(b.publishDate)
          || new Date(0);

      return db - da;
    }
  );


  return {
    success: true,
    announcements: result,
    unreadCount:
      result.filter(a => !a.read).length
  };
}


/* =====================================================
 * TARGETING
 * ===================================================== */

function canEmployeeSeeAnnouncement_(
  announcement,
  employee,
  department
) {

  const targetType =
    String(
      announcement.TargetType || 'ALL'
    ).toUpperCase();

  if (targetType === 'ALL') {
    return true;
  }


  if (
    targetType === 'DEPARTMENT'
  ) {

    const target =
      normalize_(
        announcement.TargetValue
      );

    return (
      target === department
    );
  }


  if (
    targetType === 'EMPLOYEE'
  ) {

    const ids =
      String(
        announcement.TargetValue || ''
      )
      .split(',')
      .map(x =>
        x.trim().toUpperCase()
      );

    return ids.includes(
      String(employee.EmployeeID)
        .toUpperCase()
    );
  }


  return false;
}


/* =====================================================
 * READ ANNOUNCEMENT
 * ===================================================== */

function markAnnouncementRead(
  token,
  announcementId
) {

  const session =
    requireLogin_(token);

  announcementId =
    String(announcementId);

  const reads =
    getRows_('AnnouncementReads');

  const exists =
    reads.some(
      r =>
        String(r.AnnouncementID)
          === announcementId &&
        String(r.EmployeeID)
          .toUpperCase()
          === String(session.employeeId)
            .toUpperCase()
    );

  if (!exists) {

    getSheet_('AnnouncementReads')
      .appendRow([
        announcementId,
        session.employeeId,
        new Date()
      ]);

    audit_(
      session.employeeId,
      'READ',
      'ANNOUNCEMENT',
      announcementId,
      'Đã đọc thông báo'
    );
  }

  return {
    success: true
  };
}


/* =====================================================
 * ANNOUNCEMENT DETAIL
 * ===================================================== */

function getAnnouncementDetail(
  token,
  announcementId
) {

  const session =
    requireLogin_(token);

  const data =
    getAnnouncements(
      token,
      {}
    );

  const item =
    data.announcements.find(
      a =>
        String(a.id)
        === String(announcementId)
    );

  if (!item) {
    throw new Error(
      'Không tìm thấy thông báo.'
    );
  }

  markAnnouncementRead(
    token,
    announcementId
  );

  return {
    success: true,
    announcement: item
  };
}


/* =====================================================
 * ADMIN: CREATE ANNOUNCEMENT
 * ===================================================== */

function createAnnouncement(
  token,
  data
) {

  const session =
    requireRole_(
      token,
      ['ADMIN', 'HR ADMIN', 'HR']
    );

  data =
    data || {};

  if (!data.title) {
    throw new Error(
      'Tiêu đề không được để trống.'
    );
  }

  if (!data.content) {
    throw new Error(
      'Nội dung không được để trống.'
    );
  }


  const id =
    generateAnnouncementId_();


  let status =
    String(
      data.status || 'PUBLISHED'
    ).toUpperCase();


  const publishDate =
    data.publishDate
      ? new Date(data.publishDate)
      : new Date();


  const scheduledAt =
    data.scheduledAt
      ? new Date(data.scheduledAt)
      : '';


  if (
    scheduledAt &&
    scheduledAt instanceof Date &&
    scheduledAt > new Date()
  ) {
    status = 'SCHEDULED';
  }


  getSheet_('Announcements')
    .appendRow([

      id,

      data.title,

      data.content,

      data.type || 'GENERAL',

      publishDate,

      session.employeeId,

      data.priority || 'NORMAL',

      data.pinned === true
        ? true
        : false,

      status,

      data.attachment || '',

      data.keywords || '',

      data.targetType || 'ALL',

      data.targetValue || '',

      scheduledAt,

      new Date(),

      new Date()
    ]);


  audit_(
    session.employeeId,
    'CREATE',
    'ANNOUNCEMENT',
    id,
    JSON.stringify(data)
  );


  return {
    success: true,
    id: id
  };
}


/* =====================================================
 * ADMIN: UPDATE ANNOUNCEMENT
 * ===================================================== */

function updateAnnouncement(
  token,
  announcementId,
  data
) {

  const session =
    requireRole_(
      token,
      ['ADMIN', 'HR ADMIN', 'HR']
    );

  const sheet =
    getSheet_('Announcements');

  const values =
    sheet.getDataRange().getValues();

  const headers =
    values[0];

  const idCol =
    headers.indexOf('ID');

  let rowIndex = -1;

  for (
    let i = 1;
    i < values.length;
    i++
  ) {

    if (
      String(values[i][idCol])
        === String(announcementId)
    ) {

      rowIndex = i + 1;
      break;
    }
  }

  if (rowIndex < 0) {
    throw new Error(
      'Không tìm thấy thông báo.'
    );
  }


  const current =
    rowToObject_(
      headers,
      values[rowIndex - 1]
    );


  Object.keys(data || {})
    .forEach(key => {

      if (
        headers.includes(key)
      ) {
        current[key] =
          data[key];
      }
    });


  current.UpdatedAt =
    new Date();


  sheet
    .getRange(
      rowIndex,
      1,
      1,
      headers.length
    )
    .setValues([
      objectToRow_(
        headers,
        current
      )
    ]);


  audit_(
    session.employeeId,
    'UPDATE',
    'ANNOUNCEMENT',
    announcementId,
    JSON.stringify(data)
  );


  return {
    success: true
  };
}


/* =====================================================
 * ADMIN: DELETE
 * ===================================================== */

function deleteAnnouncement(
  token,
  announcementId
) {

  const session =
    requireRole_(
      token,
      ['ADMIN', 'HR ADMIN']
    );

  const sheet =
    getSheet_('Announcements');

  const data =
    sheet.getDataRange().getValues();

  const headers =
    data[0];

  const idCol =
    headers.indexOf('ID');

  for (
    let i = 1;
    i < data.length;
    i++
  ) {

    if (
      String(data[i][idCol])
        === String(announcementId)
    ) {

      sheet.deleteRow(i + 1);

      audit_(
        session.employeeId,
        'DELETE',
        'ANNOUNCEMENT',
        announcementId,
        'Xóa thông báo'
      );

      return {
        success: true
      };
    }
  }

  throw new Error(
    'Không tìm thấy thông báo.'
  );
}


/* =====================================================
 * PIN
 * ===================================================== */

function toggleAnnouncementPin(
  token,
  announcementId
) {

  const session =
    requireRole_(
      token,
      ['ADMIN', 'HR ADMIN', 'HR']
    );

  const sheet =
    getSheet_('Announcements');

  const data =
    sheet.getDataRange().getValues();

  const headers =
    data[0];

  const idCol =
    headers.indexOf('ID');

  const pinCol =
    headers.indexOf('Pinned');

  for (
    let i = 1;
    i < data.length;
    i++
  ) {

    if (
      String(data[i][idCol])
        === String(announcementId)
    ) {

      const current =
        String(
          data[i][pinCol]
        ).toUpperCase() === 'TRUE';

      sheet
        .getRange(
          i + 1,
          pinCol + 1
        )
        .setValue(!current);

      audit_(
        session.employeeId,
        'PIN',
        'ANNOUNCEMENT',
        announcementId,
        String(!current)
      );

      return {
        success: true,
        pinned: !current
      };
    }
  }

  throw new Error(
    'Không tìm thấy thông báo.'
  );
}


/* =====================================================
 * ADMIN LIST
 * ===================================================== */

function adminGetAnnouncements(
  token
) {

  requireRole_(
    token,
    ['ADMIN', 'HR ADMIN', 'HR']
  );

  return {
    success: true,
    announcements:
      getRows_('Announcements')
        .reverse()
  };
}


/* =====================================================
 * READ STATISTICS
 * ===================================================== */

function getAnnouncementStatistics(
  token,
  announcementId
) {

  requireRole_(
    token,
    ['ADMIN', 'HR ADMIN', 'HR']
  );

  const announcements =
    getRows_('Announcements');

  const announcement =
    announcements.find(
      a =>
        String(a.ID)
        === String(announcementId)
    );

  if (!announcement) {
    throw new Error(
      'Không tìm thấy thông báo.'
    );
  }


  const employees =
    getRows_('Employees')
      .filter(
        e =>
          String(e.Status)
            .toUpperCase()
          === 'ACTIVE'
      );


  const eligible =
    employees.filter(
      e =>
        canEmployeeSeeAnnouncement_(
          announcement,
          e,
          normalize_(e.Department)
        )
    );


  const reads =
    getRows_('AnnouncementReads');

  const readIds =
    new Set(
      reads
        .filter(
          r =>
            String(r.AnnouncementID)
            === String(announcementId)
        )
        .map(
          r =>
            String(r.EmployeeID)
              .toUpperCase()
        )
    );


  const readCount =
    eligible.filter(
      e =>
        readIds.has(
          String(e.EmployeeID)
            .toUpperCase()
        )
    ).length;


  const total =
    eligible.length;

  const unread =
    Math.max(
      0,
      total - readCount
    );

  const rate =
    total
      ? (readCount / total) * 100
      : 0;


  return {
    success: true,

    announcementId:
      announcementId,

    title:
      announcement.Title,

    targetCount:
      total,

    readCount:
      readCount,

    unreadCount:
      unread,

    readRate:
      Number(rate.toFixed(2))
  };
}


/* =====================================================
 * ATTACHMENT UPLOAD
 * ===================================================== */

function uploadAttachment(
  token,
  file
) {

  const session =
    requireRole_(
      token,
      ['ADMIN', 'HR ADMIN', 'HR']
    );

  if (!file) {
    throw new Error(
      'Không có file.'
    );
  }

  const bytes =
    Utilities.base64Decode(
      file.base64
    );

  if (
    bytes.length >
    CONFIG.MAX_ATTACHMENT_BYTES
  ) {
    throw new Error(
      'File vượt quá 25 MB.'
    );
  }


  const blob =
    Utilities.newBlob(
      bytes,
      file.mimeType,
      file.name
    );


  const folder =
    getAttachmentFolder_();


  const driveFile =
    folder.createFile(blob);


  audit_(
    session.employeeId,
    'UPLOAD',
    'ATTACHMENT',
    driveFile.getId(),
    file.name
  );


  return {
    success: true,

    id:
      driveFile.getId(),

    name:
      driveFile.getName(),

    url:
      driveFile.getUrl(),

    downloadUrl:
      'https://drive.google.com/uc?export=download&id='
      + driveFile.getId()
  };
}


function getAttachmentFolder_() {

  const props =
    PropertiesService
      .getScriptProperties();

  const savedId =
    props.getProperty(
      'ATTACHMENT_FOLDER_ID'
    );

  if (savedId) {

    try {
      return DriveApp
        .getFolderById(savedId);
    } catch (e) {}
  }


  const folders =
    DriveApp
      .getFoldersByName(
        CONFIG.DRIVE_FOLDER_NAME
      );

  let folder;

  if (folders.hasNext()) {
    folder = folders.next();
  } else {
    folder =
      DriveApp.createFolder(
        CONFIG.DRIVE_FOLDER_NAME
      );
  }


  props.setProperty(
    'ATTACHMENT_FOLDER_ID',
    folder.getId()
  );

  return folder;
}


/* =====================================================
 * EXCEL IMPORT
 *
 * Excel is parsed on browser with SheetJS.
 * Server receives normalized rows.
 * ===================================================== */

function previewEmployeeImport(
  token,
  rows,
  fileName
) {

  requireRole_(
    token,
    ['ADMIN', 'HR ADMIN']
  );

  const normalized =
    normalizeEmployeeRows_(rows);

  const employees =
    getRows_('Employees');

  const existing =
    new Set(
      employees.map(
        e =>
          String(e.EmployeeID)
            .toUpperCase()
      )
    );


  let added = 0;
  let updated = 0;
  let skipped = 0;


  normalized.forEach(r => {

    if (!r.EmployeeID) {
      skipped++;
      return;
    }

    if (
      existing.has(
        String(r.EmployeeID)
          .toUpperCase()
      )
    ) {
      updated++;
    } else {
      added++;
    }
  });


  return {
    success: true,

    fileName:
      fileName || '',

    totalRows:
      normalized.length,

    added:
      added,

    updated:
      updated,

    skipped:
      skipped,

    rows:
      normalized
  };
}


function confirmEmployeeImport(
  token,
  payload
) {

  const session =
    requireRole_(
      token,
      ['ADMIN', 'HR ADMIN']
    );

  if (!payload) {
    throw new Error(
      'Dữ liệu import không hợp lệ.'
    );
  }

  const rows =
    payload.rows || [];

  const fileName =
    payload.fileName || '';

  const mode =
    payload.mode === 'INACTIVE'
      ? 'INACTIVE'
      : 'NO_CHANGE';


  const sheet =
    getSheet_('Employees');

  const values =
    sheet.getDataRange().getValues();

  const headers =
    values[0];


  const idCol =
    headers.indexOf(
      'EmployeeID'
    );


  const rowMap = {};

  for (
    let i = 1;
    i < values.length;
    i++
  ) {

    const id =
      String(
        values[i][idCol]
      )
      .trim()
      .toUpperCase();

    if (id) {
      rowMap[id] = i + 1;
    }
  }


  const importedIds =
    new Set();


  let added = 0;
  let updated = 0;
  let skipped = 0;
  let inactiveMarked = 0;


  rows.forEach(row => {

    const employeeId =
      String(
        row.EmployeeID || ''
      )
      .trim()
      .toUpperCase();

    if (!employeeId) {
      skipped++;
      return;
    }

    importedIds.add(employeeId);


    if (
      rowMap[employeeId]
    ) {

      const rowNumber =
        rowMap[employeeId];

      const current =
        sheet
          .getRange(
            rowNumber,
            1,
            1,
            headers.length
          )
          .getValues()[0];


      headers.forEach(
        (header, index) => {

          if (
            row[header] !== undefined &&
            header !== 'CreatedAt' &&
            header !== 'LastLogin'
          ) {

            if (
              row[header] !== ''
            ) {
              current[index] =
                row[header];
            }
          }
        }
      );


      const statusIndex =
        headers.indexOf('Status');

      if (statusIndex >= 0) {
        current[statusIndex] =
          'ACTIVE';
      }

      const updatedIndex =
        headers.indexOf('UpdatedAt');

      if (updatedIndex >= 0) {
        current[updatedIndex] =
          new Date();
      }


      sheet
        .getRange(
          rowNumber,
          1,
          1,
          headers.length
        )
        .setValues([current]);


      updated++;

    } else {

      const newRow =
        headers.map(
          header => {

            if (
              header === 'CreatedAt'
            ) {
              return new Date();
            }

            if (
              header === 'UpdatedAt'
            ) {
              return new Date();
            }

            if (
              header === 'Status'
            ) {
              return 'ACTIVE';
            }

            return row[header] || '';
          }
        );


      sheet.appendRow(
        newRow
      );

      added++;
    }
  });


  /*
   * IMPORTANT:
   * Không bao giờ xóa nhân viên.
   *
   * Chỉ đánh dấu INACTIVE nếu Admin
   * chọn chế độ này.
   */

  if (mode === 'INACTIVE') {

    const allValues =
      sheet
        .getDataRange()
        .getValues();

    const statusCol =
      headers.indexOf(
        'Status'
      );

    for (
      let i = 1;
      i < allValues.length;
      i++
    ) {

      const id =
        String(
          allValues[i][idCol]
        )
        .trim()
        .toUpperCase();

      if (
        id &&
        !importedIds.has(id) &&
        String(
          allValues[i][statusCol]
        ).toUpperCase()
        === 'ACTIVE'
      ) {

        sheet
          .getRange(
            i + 1,
            statusCol + 1
          )
          .setValue(
            'INACTIVE'
          );

        inactiveMarked++;
      }
    }
  }


  const importId =
    'IMP-' +
    Utilities.formatDate(
      new Date(),
      Session.getScriptTimeZone(),
      'yyyyMMdd-HHmmss'
    );


  getSheet_('ImportHistory')
    .appendRow([

      importId,

      fileName,

      new Date(),

      session.employeeId,

      rows.length,

      added,

      updated,

      skipped,

      inactiveMarked,

      mode,

      'SUCCESS'
    ]);


  audit_(
    session.employeeId,
    'IMPORT_EMPLOYEES',
    'EMPLOYEES',
    importId,
    JSON.stringify({
      fileName,
      added,
      updated,
      skipped,
      inactiveMarked,
      mode
    })
  );


  return {
    success: true,

    importId,

    fileName,

    totalRows:
      rows.length,

    added,

    updated,

    skipped,

    inactiveMarked
  };
}


/* =====================================================
 * EMPLOYEE NORMALIZATION
 * ===================================================== */

function normalizeEmployeeRows_(rows) {

  if (!Array.isArray(rows)) {
    return [];
  }

  return rows.map(row => {

    const r = {};

    Object.keys(row || {})
      .forEach(key => {

        const normalized =
          normalizeHeader_(key);

        r[normalized] =
          row[key];
      });


    return {

      EmployeeID:
        firstValue_(
          r,
          [
            'EmployeeID',
            'MaNhanVien',
            'MaNV',
            'EmployeeCode',
            'Code',
            'ID'
          ]
        ),

      FullName:
        firstValue_(
          r,
          [
            'FullName',
            'HoTen',
            'HoVaTen',
            'Name',
            'Ten'
          ]
        ),

      Department:
        firstValue_(
          r,
          [
            'Department',
            'PhongBan',
            'BoPhan',
            'DonVi',
            'Xuong'
          ]
        ),

      Position:
        firstValue_(
          r,
          [
            'Position',
            'ChucVu',
            'ViTri'
          ]
        ),

      Email:
        firstValue_(
          r,
          [
            'Email',
            'EmailNhanVien'
          ]
        ),

      Phone:
        firstValue_(
          r,
          [
            'Phone',
            'SoDienThoai',
            'DienThoai'
          ]
        ),

      Status: 'ACTIVE',

      CreatedAt: '',

      UpdatedAt: '',

      LastLogin: ''
    };
  });
}


function normalizeHeader_(value) {

  return String(value || '')
    .trim()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/đ/g, 'd')
    .replace(/Đ/g, 'D')
    .replace(/[^a-zA-Z0-9]/g, '')
    .replace(
      /^(MaNhanVien|Manv)$/i,
      'EmployeeID'
    )
    .replace(
      /^HoTen$/i,
      'FullName'
    )
    .replace(
      /^PhongBan$/i,
      'Department'
    )
    .replace(
      /^ChucVu$/i,
      'Position'
    )
    .replace(
      /^SoDienThoai$/i,
      'Phone'
    );
}


function firstValue_(obj, keys) {

  for (const key of keys) {

    if (
      obj[key] !== undefined &&
      obj[key] !== null &&
      String(obj[key]).trim() !== ''
    ) {
      return obj[key];
    }
  }

  return '';
}


/* =====================================================
 * IMPORT HISTORY
 * ===================================================== */

function getImportHistory(token) {

  requireRole_(
    token,
    ['ADMIN', 'HR ADMIN']
  );

  return {
    success: true,
    history:
      getRows_('ImportHistory')
        .reverse()
  };
}


/* =====================================================
 * ADMIN USER MANAGEMENT
 * ===================================================== */

function createOrResetUser(
  token,
  employeeId,
  password,
  role
) {

  const session =
    requireRole_(
      token,
      ['ADMIN']
    );

  employeeId =
    String(employeeId)
      .trim()
      .toUpperCase();

  role =
    String(role || 'USER')
      .toUpperCase();


  if (
    password.length <
    CONFIG.PASSWORD_MIN_LENGTH
  ) {
    throw new Error(
      'Mật khẩu quá ngắn.'
    );
  }


  const sheet =
    getSheet_('Users');

  const data =
    sheet.getDataRange().getValues();

  const headers =
    data[0];

  const idCol =
    headers.indexOf(
      'EmployeeID'
    );


  for (
    let i = 1;
    i < data.length;
    i++
  ) {

    if (
      String(data[i][idCol])
        .toUpperCase()
      === employeeId
    ) {

      sheet
        .getRange(
          i + 1,
          2
        )
        .setValue(
          hashPassword(password)
        );

      sheet
        .getRange(
          i + 1,
          3
        )
        .setValue(role);

      sheet
        .getRange(
          i + 1,
          4
        )
        .setValue('ACTIVE');

      sheet
        .getRange(
          i + 1,
          6
        )
        .setValue(new Date());

      audit_(
        session.employeeId,
        'RESET_PASSWORD',
        'USER',
        employeeId,
        role
      );

      return {
        success: true,
        mode: 'UPDATED'
      };
    }
  }


  sheet.appendRow([

    employeeId,

    hashPassword(password),

    role,

    'ACTIVE',

    new Date(),

    new Date()
  ]);


  audit_(
    session.employeeId,
    'CREATE_USER',
    'USER',
    employeeId,
    role
  );


  return {
    success: true,
    mode: 'CREATED'
  };
}


/* =====================================================
 * DASHBOARD
 * ===================================================== */

function getDashboard(
  token
) {

  const session =
    requireLogin_(token);

  const current =
    getCurrentUser(token);

  const announcements =
    getAnnouncements(
      token,
      {}
    );


  const allVisible =
    announcements.announcements;


  const unread =
    allVisible.filter(
      a => !a.read
    ).length;


  return {

    success: true,

    user:
      current.employee,

    role:
      session.role,

    unreadCount:
      unread,

    totalAnnouncements:
      allVisible.length,

    urgent:
      allVisible.filter(
        a =>
          a.type === 'URGENT'
      ).slice(0, 5),

    pinned:
      allVisible
        .filter(a => a.pinned)
        .slice(0, 5),

    latest:
      allVisible.slice(0, 10)
  };
}


/* =====================================================
 * TYPES
 * ===================================================== */

function getAnnouncementTypes() {

  return CONFIG.ANNOUNCEMENT_TYPES;
}


function getTypeLabel_(type) {

  const item =
    CONFIG.ANNOUNCEMENT_TYPES[
      String(type)
    ];

  return item
    ? item.label
    : type;
}


function getTypeIcon_(type) {

  const item =
    CONFIG.ANNOUNCEMENT_TYPES[
      String(type)
    ];

  return item
    ? item.icon
    : '📢';
}


/* =====================================================
 * HELPERS
 * ===================================================== */

function normalize_(value) {

  return String(value || '')
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/đ/g, 'd')
    .trim();
}


function parseDate_(value) {

  if (!value) {
    return null;
  }

  if (
    Object.prototype.toString
      .call(value)
    === '[object Date]'
  ) {
    return value;
  }

  const d =
    new Date(value);

  return isNaN(d.getTime())
    ? null
    : d;
}


function formatDateTime_(value) {

  const d =
    parseDate_(value);

  if (!d) {
    return '';
  }

  return Utilities.formatDate(
    d,
    Session.getScriptTimeZone(),
    'dd/MM/yyyy HH:mm'
  );
}


function generateAnnouncementId_() {

  const now =
    new Date();

  return (
    'TB-' +
    Utilities.formatDate(
      now,
      Session.getScriptTimeZone(),
      'yyyyMMdd-HHmmss'
    ) +
    '-' +
    Math.floor(
      Math.random() * 1000
    )
  );
}


/* =====================================================
 * AUDIT
 * ===================================================== */

function audit_(
  employeeId,
  action,
  entity,
  entityId,
  details
) {

  try {

    getSheet_('AuditLog')
      .appendRow([

        Utilities.getUuid(),

        new Date(),

        employeeId || '',

        action || '',

        entity || '',

        entityId || '',

        details || '',

        ''
      ]);

  } catch (e) {

    console.error(
      'Audit error:',
      e
    );
  }
}


/* =====================================================
 * ADMIN STATS
 * ===================================================== */

function getAdminDashboard(
  token
) {

  requireRole_(
    token,
    ['ADMIN', 'HR ADMIN', 'HR']
  );

  const employees =
    getRows_('Employees');

  const announcements =
    getRows_('Announcements');

  const reads =
    getRows_('AnnouncementReads');


  const activeEmployees =
    employees.filter(
      e =>
        String(e.Status)
          .toUpperCase()
        === 'ACTIVE'
    ).length;


  const published =
    announcements.filter(
      a =>
        String(a.Status)
          .toUpperCase()
        === 'PUBLISHED'
    ).length;


  return {

    success: true,

    employees:
      employees.length,

    activeEmployees,

    inactiveEmployees:
      employees.length -
      activeEmployees,

    announcements:
      announcements.length,

    publishedAnnouncements:
      published,

    totalReads:
      reads.length
  };
}
