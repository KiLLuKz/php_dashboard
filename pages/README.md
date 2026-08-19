# AdminLTE 4 Starter Kit (Thai)

ชุดเทมเพลตเริ่มต้น AdminLTE v4 ที่เปลี่ยนโครงสร้างไปใช้ **Bootstrap 5** และ **Vanilla JS** (ไม่ง้อ jQuery)

## รายชื่อไฟล์
- `index.html` - หน้า Dashboard หลัก
- `forms.html` - ตัวอย่างฟอร์มรับข้อมูล (Form Components)
- `tables.html` - ตัวอย่างตารางข้อมูล (Data Table & Pagination)

## ข้อควรจำในการเปลี่ยนคลาส
1. Layout: `.wrapper` -> `.app-wrapper`, `.main-header` -> `.app-header`, `.main-sidebar` -> `.app-sidebar`
2. Data attributes: `data-widget="pushmenu"` -> `data-lte-toggle="sidebar-toggle"`, `data-toggle` -> `data-bs-toggle`
3. Utility: `ml-*`/`mr-*` -> `ms-*`/`me-*`, `float-left`/`float-right` -> `float-start`/`float-end`
