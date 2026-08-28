# tb
Kiến trúc project
company-internal-portal/
│
├── Code.gs
├── Index.html
├── app.html
├── style.html
├── javascript.html
├── README.md
└── appsscript.json

Google Apps Script sẽ dùng:
Google Sheets
    │
    ├── Employees
    ├── Users
    ├── Announcements
    ├── AnnouncementReads
    ├── ImportHistory
    └── AuditLog
             │
             ▼
       Google Apps Script
             │
       ┌─────┴─────┐
       │           │
   Công nhân       HR/Admin
       │           │
       └───── Web App ─────┘

Các loại thông báo
GENERAL       📢 Thông báo chung
WORK_TIME     🕐 Thời gian làm việc
SALARY        💰 Lương & thưởng
BENEFIT       🎁 Phúc lợi
HEALTH        🏥 Sức khỏe
HSE           🦺 An toàn / HSE
PRODUCTION    🏭 Sản xuất
HR            📋 Nhân sự
EVENT         🎉 Hoạt động công ty
URGENT        ⚠️ Khẩn cấp
