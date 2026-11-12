# Solar People — ระบบลงทะเบียนเข้างาน

- **Frontend**: GitHub Pages (โฟลเดอร์ `docs/`)
  - `index.html` แดชบอร์ด (ดูได้โดยไม่ต้องล็อกอิน)
  - `checkin.html` หน้าลงทะเบียนสำหรับเจ้าหน้าที่
  - `admin.html` หน้าจัดการผู้ใช้ (แอดมินเท่านั้น)
- **Backend**: Google Apps Script (โฟลเดอร์ `gas/`, เชื่อมด้วย clasp)
- **ฐานข้อมูล**: Google Sheet ที่ `setup()` สร้างให้ มี 3 แท็บ
  - `Companies` (PEA No., ชื่อบริษัท, เลขที่ มท., ลงวันที่, สถานะ, จำนวนคน, เวลาเข้างาน, ผู้บันทึก)
  - `Users`
  - `Log`

รายชื่อบริษัทเริ่มต้น 725 รายการมาจากไฟล์ `Copy List solar people.xlsx` ชีต `20k_N` (อยู่ใน `gas/SeedData.js`)

## ติดตั้งครั้งแรก

1. เปิด Apps Script editor: `clasp open-script`
2. เลือกฟังก์ชัน `connectSheet` แล้วกด **Run** จากนั้นกดอนุญาตสิทธิ์
   - ระบบจะใช้ Sheet ตาม `TARGET_SPREADSHEET_ID` ใน `gas/Code.js` สร้างแท็บ Companies/Users/Log และใส่รายชื่อ 725 บริษัทให้
   - บัญชีเจ้าของสคริปต์ต้องมีสิทธิ์แก้ไข Sheet นั้น
   - ถ้าจะให้สร้าง Sheet ใหม่แทน ให้รัน `setup`
3. เปิด Web App URL ใน `docs/js/config.js` (ต่อท้ายด้วย `?action=list`) ต้องเห็นข้อมูล JSON
   - ถ้าขึ้น "การเข้าถึงถูกปฏิเสธ" ให้ไปที่ Deploy → Manage deployments → แก้ไข แล้วตั้งค่า
     - Execute as: **Me**
     - Who has access: **Anyone**
4. เปิด GitHub Pages ที่ Settings → Pages → Branch `main` / โฟลเดอร์ `/docs`

## ผู้ใช้เริ่มต้น

| Username | สิทธิ์ |
|---|---|
| `admin` | ผู้ดูแลระบบ (เพิ่ม ปิด หรือลบเจ้าหน้าที่ได้ที่ `admin.html`) |
| `emd2` | เจ้าหน้าที่ |

ระบบล็อกอินด้วย Username อย่างเดียว ควรเปลี่ยนชื่อ `admin` ในแท็บ Users ให้เดายาก

## แก้ไขรายชื่อบริษัท

เพิ่มหรือแก้รายชื่อได้โดยตรงในแท็บ `Companies` ได้เลย
- ต้องมีทั้ง PEA No. และชื่อบริษัท แถวนั้นจึงจะแสดงในแอป
- PEA No. ต้องไม่ซ้ำกัน

## อัปเดตโค้ด backend

```bash
clasp push -f
clasp update-deployment AKfycbyXmFKw0Ohu-Qa0rhSKyVGj78Um_yjAdd0zzdIl1u3UUPs3gZMWHx7CGmbCJBCGpbE9Ng
```
