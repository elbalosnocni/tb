# Company Internal Portal - Google Apps Script

## 1. Cài đặt
- Tạo Google Sheet.
- Extensions > Apps Script.
- Tạo `Code.gs`, dán Code.gs.
- Tạo HTML file tên `index`, dán index.html.
- Deploy > New deployment > Web app.
- Execute as: Me.
- Who has access: theo chính sách công ty.
- Mở URL /exec.

Không cần chạy `setupSystem("MatKhauAdminMoi")`. Lần đầu Web App tự hiện màn hình khởi tạo.

## 2. Tài khoản lần đầu
Nếu không đổi mật khẩu ở màn hình setup thì có thể chạy `setupSystem()` từ editor để tạo ADMIN mặc định:
- Username: ADMIN
- Password: Admin@123456

Nên đổi ngay sau đăng nhập.

## 3. Google Sheets tự tạo
Config, Users, Employees, Announcements, AnnouncementTargets, AnnouncementAttachments, AnnouncementReads, ImportHistory, LoginHistory, AuditLog.

## 4. Google Drive
Tự tạo COMPANY_INTERNAL_PORTAL/Announcements, EmployeeImports, Temp.

## 5. Import Excel
Sheet đầu tiên của file Excel được đọc. Cột mã nhân viên hỗ trợ các tên: EmployeeID, Employee Code, Mã nhân viên, Mã NV...
Mã nhân viên là khóa chính. Không xóa nhân viên cũ. Có lựa chọn đánh dấu INACTIVE.

## 6. Lưu ý file
File đính kèm tối đa 20 MB/file. Apps Script có giới hạn runtime/quota; với video lớn hoặc lưu lượng rất cao nên chuyển storage sang Cloud Storage.
