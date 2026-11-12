// ตั้งค่าแอป — แก้ API_URL เป็น Web App URL ของ Apps Script (ลงท้ายด้วย /exec)
window.APP_CONFIG = {
  API_URL: 'https://script.google.com/macros/s/AKfycbxysnlfwCjY4ViL7dmN598E9C4ubrso6R2Pt8aBuUdo9pCrHz_c_oZ7g8uTdmZN0wZJXg/exec',
  ORG_NAME: 'การไฟฟ้าส่วนภูมิภาค (กฟภ.)',
  EVENT_NAME: 'ลงทะเบียนเข้างาน โซลาร์ภาคประชาชน', // เว้นวรรคเพื่อกำหนดจุดตัดบรรทัดบนมือถือ
  REFRESH_MS: 10000, // รีเฟรชข้อมูลอัตโนมัติทุก 10 วินาที
};
