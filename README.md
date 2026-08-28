# Cổng thông báo nội bộ — Google Apps Script

Web App nội bộ cho công nhân/nhân viên đăng nhập để xem thông báo. Dữ liệu chính nằm trong Google Sheets.

## Chức năng

- Đăng nhập bằng Mã nhân viên + mật khẩu.
- Password được lưu dạng SHA-256 + salt, không lưu plaintext.
- Thông báo mới nhất ở trên; thông báo ghim ở trên cùng.
- Lọc theo tháng, loại, bộ phận, chưa đọc.
- Tìm kiếm trong mã thông báo, tiêu đề, nội dung, loại, từ khóa.
- Trạng thái Đã đọc / Chưa đọc.
- Theo dõi số người đã đọc, chưa đọc và tỷ lệ đọc.
- HR/Admin tạo, sửa, xem, ghim, xóa, hẹn giờ đăng.
- Phân quyền theo ALL / DEPARTMENT / EMPLOYEE / TYPE.
- Đính kèm PDF, Word, Excel, PowerPoint, ảnh, video qua Google Drive.
- Import Excel nhân viên:
  - EmployeeID là khóa chính.
  - Không tạo bản ghi trùng.
  - Không xóa nhân viên cũ.
  - Có tùy chọn đánh dấu INACTIVE cho nhân viên không xuất hiện trong file.
  - Có ImportHistory + AuditLog.
- Có LoginHistory và AnnouncementReads.

## Cấu trúc Sheets

### Employees

`EmployeeID, FullName, Department, Position, Email, Phone, PasswordHash, PasswordSalt, Role, Status, CreatedAt, UpdatedAt, LastLoginAt`

### Announcements

`ID, Title, Content, Type, PublishDate, Author, Priority, Pinned, Status, Attachment, Keywords, CreatedAt, UpdatedAt, ScheduledAt`

### AnnouncementTargets

`AnnouncementID, TargetType, TargetValue`

Ví dụ:

- `TB-001 | ALL | ALL`
- `TB-002 | DEPARTMENT | Xưởng Bánh`
- `TB-003 | EMPLOYEE | NV000123`
- `TB-004 | TYPE | SUPERVISOR`

### AnnouncementReads

`AnnouncementID, EmployeeID, ReadAt`

### ImportHistory

Lưu lịch sử import, file, người import, thời gian, số thêm/cập nhật/bỏ qua/INACTIVE.

### LoginHistory

Lưu lịch sử đăng nhập thành công/thất bại.

### AuditLog

Lưu hành động quản trị.

## Cài đặt

1. Tạo một Google Spreadsheet mới.
2. Extensions → Apps Script.
3. Tạo các file:
   - `Code.gs`
   - `Attachment.gs`
   - `app.html`
   - `style.html`
   - `client.html`
4. Copy code tương ứng.
5. Trong Apps Script, chạy `setupSystem()` một lần và cấp quyền.
6. Sau setup, tài khoản admin mẫu:
   - Mã: `ADMIN001`
   - Mật khẩu: `ChangeMe@123`
7. Đăng nhập và đổi mật khẩu ngay.
8. Deploy → New deployment → Web app.
9. Execute as: Me.
10. Who has access: chọn phạm vi phù hợp với Google Workspace của công ty.

## Import Excel

File Excel nên có dòng đầu là tên cột. Hệ thống nhận các tên phổ biến:

- EmployeeID / Mã nhân viên / Mã NV
- FullName / Họ tên
- Department / Phòng ban / Bộ phận
- Position / Chức vụ / Vị trí
- Email
- Phone / Điện thoại / SĐT
- Role / Vai trò
- Password / Mật khẩu (tùy chọn)

Nếu Password bỏ trống khi thêm nhân viên mới, mật khẩu ban đầu = EmployeeID.

**Lưu ý:** file Excel được đọc ở trình duyệt bằng SheetJS CDN. Nếu môi trường công ty không cho CDN, tải SheetJS về và đổi URL trong `client.html`.

## Hẹn giờ

`setupSystem()` tạo trigger 5 phút/lần. Khi `Status=SCHEDULED` và `ScheduledAt` đã tới, hệ thống chuyển sang `PUBLISHED`.

## Bảo mật triển khai thực tế

- Không dùng tài khoản mẫu lâu dài.
- Giới hạn Web App cho Google Workspace/domain nếu có thể.
- Nếu công ty yêu cầu bảo mật cao, nên bổ sung Google Workspace SSO thay vì password tự quản lý.
- Drive sharing "Anyone with link" có thể bị Workspace Admin chặn; khi đó preview cần người xem có quyền Google Drive.
- Với dữ liệu hàng chục nghìn nhân viên/lượt đọc rất lớn, nên tối ưu bằng batch read/write và cache/index; bản này phù hợp để triển khai MVP/small-to-medium scale.

## GitHub

Các file trong thư mục này có thể đưa thẳng vào repository GitHub. Không commit ID spreadsheet, token, mật khẩu thật hoặc thông tin nhân viên.
