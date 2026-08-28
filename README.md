# Company Internal Notice Portal v2

## Mục tiêu

Bản v2 tách hoàn toàn frontend HTML khỏi Google Apps Script UI.

- `index.html` = frontend độc lập.
- `api.gs` = REST-like JSON API.
- Google Sheets = database.
- Google Drive = file storage.
- Frontend gọi GAS bằng `fetch()`.
- Không dùng `google.script.run`.
- Không phụ thuộc `doGet()` để render giao diện.

## Cài đặt

### 1. Google Sheet

Tạo Google Spreadsheet.

### 2. Apps Script

Extensions → Apps Script.

Tạo `api.gs` và dán code trong file `api.gs`.

Chạy:

```text
setupSystem()
```

Sau đó cấp quyền.

Hệ thống tạo:

```text
Employees
Announcements
AnnouncementTargets
AnnouncementReads
ImportHistory
LoginHistory
AuditLog
```

### 3. Deploy API

Deploy → New deployment → Web app.

Khuyến nghị:

```text
Execute as: Me
Who has access: người dùng/domain phù hợp
```

Copy URL có dạng:

```text
https://script.google.com/macros/s/XXXXXXXX/exec
```

### 4. Frontend

Mở `index.html`.

Tìm:

```javascript
const API_URL = 'PASTE_YOUR_GAS_WEB_APP_EXEC_URL_HERE';
```

Đổi thành URL `/exec` vừa copy.

Ví dụ:

```javascript
const API_URL = 'https://script.google.com/macros/s/XXXXXXXX/exec';
```

Sau đó upload `index.html` lên:

- GitHub Pages
- server nội bộ
- hosting HTML
- hoặc mở trực tiếp nếu trình duyệt/CORS policy cho phép.

## Tài khoản admin ban đầu

```text
EmployeeID: ADMIN001
Password: ChangeMe@123
```

Đăng nhập và đổi mật khẩu ngay.

## Excel Import

Header có thể là:

```text
EmployeeID
FullName
Department
Position
Email
Phone
Role
Password
```

hoặc tiếng Việt:

```text
Mã nhân viên
Họ tên
Phòng ban
Bộ phận
Chức vụ
Điện thoại
Vai trò
Mật khẩu
```

Quy tắc:

1. EmployeeID là khóa chính.
2. ID mới → ADD.
3. ID đã tồn tại → UPDATE thông tin nhân viên.
4. Không tạo bản ghi trùng.
5. Không xóa nhân viên cũ.
6. Nếu file mới không có một nhân viên:
   - NO_CHANGE: giữ nguyên.
   - INACTIVE: đánh dấu INACTIVE.
7. Không thay đổi PasswordHash/PasswordSalt của nhân viên cũ khi import.
8. Lưu ImportHistory.
9. Lưu AuditLog.
10. Lịch sử đọc/login không bị xóa.

## Phân quyền thông báo

`AnnouncementTargets`:

```text
AnnouncementID | TargetType  | TargetValue
TB-001         | ALL         | ALL
TB-002         | DEPARTMENT  | Xưởng Bánh
TB-003         | EMPLOYEE    | NV000123
TB-004         | TYPE        | SUPERVISOR
```

## Attachment

HR upload:

- PDF
- Word
- Excel
- PowerPoint
- PNG/JPG/GIF/WebP
- MP4

File được lưu trong Drive folder:

```text
CompanyNoticeAttachments
```

Metadata được lưu vào `Announcements.Attachment`.

## Lưu ý CORS / mở file

Frontend dùng `fetch()` tới GAS Web App.

Nếu trình duyệt/server của bạn chặn request khi mở `index.html` bằng `file://`, hãy chạy frontend qua:

- GitHub Pages
- web server nội bộ
- hosting HTTPS

Đây là cách ổn định hơn.

## Production

Bản v2 đã bỏ `google.script.run` khỏi frontend. Tuy nhiên nếu triển khai cho hàng nghìn công nhân, nên nâng cấp tiếp:

- rate limiting login
- session store bền vững hơn CacheService
- batch writes cho Reads
- index/cache cho Employees và AnnouncementTargets
- phân trang API
- tìm kiếm server-side
- upload resumable/chunk nếu file lớn
- Google Workspace SSO/OIDC nếu công ty cần bảo mật cao
- CSP/security headers ở frontend hosting


## v2.1
- Không lưu token đăng nhập giữa các lần mở/reload trang.
- Loading overlay khi đăng nhập, import, lưu thông báo.
- Tab Quan trọng dùng API riêng.
- Mỗi thông báo có “Bấm vào xem chi tiết →”.
- API báo rõ khi GAS trả HTML thay vì JSON.
- Khi sửa deployment, tạo **New version** và dùng URL `/exec`, không dùng `/dev`.
