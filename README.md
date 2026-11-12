



## ผู้ใช้เริ่มต้น

| Username | สิทธิ์ |
|---|---|
| | ผู้ดูแลระบบ (เพิ่ม ปิด หรือลบเจ้าหน้าที่ได้ที่ `admin.html`) |
| | เจ้าหน้าที่ |

ระบบล็อกอินด้วย Username อย่างเดียว ควรเปลี่ยนชื่อ `admin` ในแท็บ Users ให้เดายาก

## แก้ไขรายชื่อบริษัท

เพิ่มหรือแก้รายชื่อได้โดยตรงในแท็บ `Companies` ได้เลย
- ต้องมีทั้ง PEA No. และชื่อบริษัท แถวนั้นจึงจะแสดงในแอป
- PEA No. ต้องไม่ซ้ำกัน

## อัปเดตโค้ด backend

```bash
clasp push -f
clasp update-deployment AKfycbxysnlfwCjY4ViL7dmN598E9C4ubrso6R2Pt8aBuUdo9pCrHz_c_oZ7g8uTdmZN0wZJXg
```
