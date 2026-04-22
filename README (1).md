# MED UP Interview System
### คณะแพทยศาสตร์ มหาวิทยาลัยพะเยา

ระบบประเมินสัมภาษณ์ MED UP & วัฒนธรรมองค์กร  
รันได้เลยบน GitHub Pages — ไม่ต้องติดตั้งอะไร

---

## 🚀 วิธี Deploy บน GitHub Pages

1. สร้าง Repository ใหม่บน GitHub
2. Upload ไฟล์ `index.html` ขึ้นไปที่ root
3. ไปที่ **Settings → Pages → Source: Deploy from branch → main → / (root)**
4. กด **Save** แล้วรอ 1–2 นาที
5. เข้าใช้งานได้ที่ `https://<username>.github.io/<repo-name>`

---

## 🔗 วิธีเชื่อม Google Sheets

### ขั้นตอน

1. เปิด **Google Sheets** ใหม่
2. Copy **Sheet ID** จาก URL  
   `https://docs.google.com/spreadsheets/d/**[SHEET_ID]**/edit`
3. ไปที่เมนู **Extensions → Apps Script**
4. **ลบโค้ดเดิมทั้งหมด** แล้ววางโค้ดจากไฟล์ `google_apps_script.js`
5. แก้บรรทัด `var SHEET_ID = 'YOUR_GOOGLE_SHEET_ID_HERE';` ให้เป็น ID จริง
6. กด **Deploy → New deployment**
   - Type: **Web app**
   - Execute as: **Me**
   - Who has access: **Anyone**
7. กด **Deploy** → อนุมัติสิทธิ์ → Copy **Web App URL**
8. นำ URL ไปวางใน Dashboard แล้วกด **"เชื่อม"**

### โครงสร้าง Sheet ที่จะสร้างอัตโนมัติ

| Sheet | รายละเอียด |
|-------|-----------|
| `Summary` | คะแนนรวมทุกคน ทุกกรรมการ |
| `ผศ.นพ.สมชาย ใจดี` | คะแนนของกรรมการแต่ละคน (สร้างอัตโนมัติ) |

---

## 📁 โครงสร้างไฟล์

```
/
├── index.html              ← ตัวแอปหลัก (GitHub Pages)
├── google_apps_script.js   ← วางใน Google Apps Script
└── README.md
```

---

## ✨ ฟีเจอร์

- 📝 **ให้คะแนน** — MED UP 6 ด้าน (M E D U P C) คะแนน 1–5
- 👤 **จัดการกรรมการ** — เพิ่ม/ลบ manual หรือ Upload CSV
- 📋 **รายชื่อผู้สมัคร** — Sidebar พร้อมสถานะประเมินแล้ว/ยังไม่ได้ประเมิน
- 📊 **สรุปผล** — ค้นหา, กรอง, เรียงคะแนนทุกด้าน
- 🏆 **รับเข้าทำงาน** — ลำดับหลัก + ลำดับสำรอง
- 🖨️ **Export PDF** — ภาษาไทยอ่านออก, 4 แบบ (สรุป / รายบุคคล / คนที่เลือก / รับเข้าทำงาน)
- ☁️ **Google Sheets** — บันทึก/โหลดข้อมูลอัตโนมัติ

---

## ⚠️ ข้อผิดพลาดที่พบบ่อย

| ปัญหา | สาเหตุ | วิธีแก้ |
|-------|--------|---------|
| Apps Script error | ลืมลบโค้ดเดิม | ลบทั้งหมดก่อนวาง |
| Error: Cannot read properties | ลืมแก้ SHEET_ID | แก้ `YOUR_GOOGLE_SHEET_ID_HERE` |
| ไม่ยอม deploy | เลือก "Anyone with link" | ต้องเลือก **"Anyone"** เท่านั้น |
| CORS error | Deploy version เก่า | Deploy ใหม่ทุกครั้งที่แก้โค้ด |
