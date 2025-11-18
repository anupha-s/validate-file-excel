# Makro Image Checker v6.0

ตรวจสอบรูปภาพจาก Google Drive ว่าเป็นจัตุรัส (1:1) หรือสี่เหลี่ยมผืนผ้า

## ฟีเจอร์

- ✅ อัปโหลดไฟล์ Excel (Makro Image.xlsx)
- ✅ ตรวจสอบรูปภาพจาก Google Drive (คอลัมน์ F)
- ✅ เติมผลลัพธ์ Y/N ลงคอลัมน์ I
- ✅ รักษาฟอร์แมต สี สูตร จากไฟล์ต้นฉบับ
- ✅ แสดง Progress bar และ Debug console
- ✅ แยก Tab ตาม Sheet

## การใช้งาน

1. เปิดเว็บไซต์
2. คลิก "Choose File" เลือกไฟล์ Excel
3. รอให้ระบบประมวลผลรูปภาพ
4. คลิก "ดาวน์โหลด Excel ที่เติม Y/N แล้ว"

## ผลลัพธ์

- **Y** = รูปจัตุรัส (ความกว้าง = ความสูง เช่น 500×500)
- **N** = รูปสี่เหลี่ยมผืนผ้า (เช่น 800×600)

## Deploy บน Vercel

### วิธีที่ 1: ผ่าน Vercel CLI

```bash
# ติดตั้ง Vercel CLI
npm install -g vercel

# Deploy
vercel
```

### วิธีที่ 2: ผ่าน Vercel Dashboard

1. ไปที่ https://vercel.com
2. คลิก "New Project"
3. Import repository นี้
4. คลิก "Deploy"

### วิธีที่ 3: Deploy ด้วย Git

1. Push โค้ดขึ้น GitHub
2. เชื่อม GitHub กับ Vercel
3. Vercel จะ auto-deploy ทุกครั้งที่ push

## โครงสร้างไฟล์

```
.
├── index.html          # หน้าหลัก
├── styles.css          # CSS styles
├── app.js              # JavaScript logic
├── vercel.json         # Vercel configuration
└── README.md           # เอกสารนี้
```

## เทคโนโลยีที่ใช้

- **ExcelJS** - อ่าน/เขียนไฟล์ Excel พร้อมรักษาฟอร์แมต
- **Google Drive API** - โหลดรูปภาพจาก Google Drive
- **Vanilla JavaScript** - ไม่ต้องใช้ framework

## License

MIT
