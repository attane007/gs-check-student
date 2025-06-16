# ระบบเช็คชื่อนักเรียน
## Google Apps Script + HTML + Tailwind CSS + Google Sheets

ระบบเช็คชื่อนักเรียนออนไลน์ที่ใช้ Google Apps Script เป็น backend และ Google Sheets เป็นฐานข้อมูล พร้อมด้วย UI สีพาสเทลที่สวยงาม

### ✨ คุณสมบัติหลัก

- 🎨 **UI สีพาสเทล** - ใช้ Tailwind CSS กับธีมสีพาสเทลที่นุ่มนวล
- 📱 **Responsive Design** - รองรับทุกขนาดหน้าจอ
- ⚡ **Real-time Updates** - อัปเดตข้อมูลแบบ Real-time
- 📊 **Dashboard** - แสดงสถิติการเข้าเรียนประจำวัน
- 🔍 **ค้นหานักเรียน** - ค้นหาข้อมูลนักเรียนได้ทันที
- 📋 **บันทึกอัตโนมัติ** - บันทึกลง Google Sheets อัตโนมัติ
- 📈 **รายงานสถิติ** - แสดงเปอร์เซ็นต์การเข้าเรียน

### 🚀 การติดตั้งและใช้งาน

#### 1. สร้าง Google Apps Script Project

1. เข้าไปที่ [Google Apps Script](https://script.google.com)
2. คลิก **"โครงการใหม่"** หรือ **"New Project"**
3. ตั้งชื่อโครงการ เช่น "ระบบเช็คชื่อนักเรียน"

#### 2. อัปโหลดไฟล์โค้ด

1. **Code.gs**: คัดลอกโค้ดจากไฟล์ `Code.gs` ใส่ใน Apps Script
2. **index.html**: สร้างไฟล์ HTML ใหม่และคัดลอกโค้ดจากไฟล์ `index.html`

#### 3. ตั้งค่า Google Sheets

1. สร้าง Google Sheets ใหม่
2. คัดลอก Sheet ID จาก URL (ส่วนระหว่าง `/d/` และ `/edit`)
   ```
   https://docs.google.com/spreadsheets/d/[SHEET_ID]/edit
   ```
3. แก้ไขค่า `SHEET_ID` ในไฟล์ `Code.gs` บรรทัดที่ 7:
   ```javascript
   const SHEET_ID = 'ใส่ SHEET_ID ของคุณที่นี่';
   ```

#### 4. รันฟังก์ชันเริ่มต้น

1. ใน Apps Script ให้เลือกฟังก์ชัน `initializeSheets`
2. คลิก **"เรียกใช้"** หรือ **"Run"**
3. อนุญาตการเข้าถึง Google Sheets เมื่อระบบขอ
4. ตรวจสอบว่า Google Sheets มีข้อมูลตัวอย่างแล้ว

#### 5. Deploy เป็น Web App

1. คลิก **"Deploy"** > **"New Deployment"**
2. เลือกประเภท: **"Web app"**
3. ตั้งค่า:
   - Execute as: **Me**
   - Who has access: **Anyone** (หรือตามต้องการ)
4. คลิก **"Deploy"**
5. คัดลอก Web App URL เพื่อใช้งาน

### 📊 โครงสร้างข้อมูลใน Google Sheets

#### Sheet: "Students" (ข้อมูลนักเรียน)
| รหัสนักเรียน | ชื่อ | นามสกุล | ห้องเรียน |
|-------------|------|---------|----------|
| 20001 | สมชาย | ใจดี | ม.1/1 |
| 20002 | สมหญิง | ใจงาม | ม.1/1 |

#### Sheet: "Attendance" (บันทึกการเข้าเรียน)
| วันที่ | เวลา | รหัสนักเรียน | ชื่อ-นามสกุล | ห้องเรียน | สถานะ |
|--------|------|-------------|-------------|----------|-------|
| 13/06/2025 | 08:30:15 | 20001 | สมชาย ใจดี | ม.1/1 | เข้าเรียน |

### 🎨 ธีมสีพาสเทล

ระบบใช้ชุดสีพาสเทลดังนี้:
- **Pink** (#FFB3BA) - สีชมพูอ่อน
- **Peach** (#FFDFBA) - สีพีชอ่อน  
- **Yellow** (#FFFFBA) - สีเหลืองอ่อน
- **Green** (#BAFFC9) - สีเขียวอ่อน
- **Blue** (#BAE1FF) - สีฟ้าอ่อน
- **Purple** (#E1BAFF) - สีม่วงอ่อน
- **Lavender** (#F0E6FF) - สีลาเวนเดอร์
- **Mint** (#E6FFE6) - สีมิ้นต์
- **Coral** (#FFE6E6) - สีคอรัล
- **Sky** (#E6F3FF) - สีฟ้าใส

### 🔧 ฟังก์ชันหลัก

#### Backend Functions (Code.gs)
- `initializeSheets()` - เตรียมฐานข้อมูล
- `getAllStudents()` - ดึงรายชื่อนักเรียนทั้งหมด
- `findStudentById(studentId)` - ค้นหานักเรียนตามรหัส
- `recordAttendance(studentId, status)` - บันทึกการเข้าเรียน
- `getTodayAttendance()` - ดึงข้อมูลการเข้าเรียนวันนี้

#### Frontend Functions (index.html)
- `checkAttendance()` - เช็คชื่อเข้าเรียน
- `searchStudent()` - ค้นหาข้อมูลนักเรียน
- `updateStats()` - อัปเดตสถิติการเข้าเรียน
- `loadInitialData()` - โหลดข้อมูลเริ่มต้น

### 📱 การใช้งาน

1. **เช็คชื่อ**: กรอกรหัสนักเรียนและกดปุ่ม "เช็คชื่อเข้าเรียน"
2. **ค้นหา**: กรอกรหัสนักเรียนเพื่อดูข้อมูล
3. **ดูสถิติ**: ดูจำนวนนักเรียนที่เข้าเรียนและเปอร์เซ็นต์
4. **รายการล่าสุด**: ดูรายการนักเรียนที่เข้าเรียนล่าสุด

### 🔒 ความปลอดภัย

- ใช้ Google OAuth สำหรับการยืนยันตัวตน
- ข้อมูลเก็บใน Google Sheets ที่ปลอดภัย
- สามารถกำหนดสิทธิ์การเข้าถึงได้

### 🐛 การแก้ไขปัญหา

#### ปัญหาที่พบบ่อย:

1. **Error: google.script.run.withSuccessHandler(...).getStudents is not a function**
   ```javascript
   // ✅ แก้ไข: ใช้ชื่อฟังก์ชันที่ถูกต้อง
   google.script.run.getStudentsList() // ถูกต้อง
   google.script.run.getStudents()     // ผิด
   ```

2. **ไม่สามารถเชื่อมต่อ Google Sheets**
   - ตรวจสอบว่า Script ผูกกับ Google Sheets แล้ว
   - รันฟังก์ชัน `initializeSheets()` ก่อน Deploy
   - ตรวจสอบสิทธิ์การเข้าถึง Sheets

3. **ข้อมูลไม่อัปเดต**
   - เปิด Console (F12) ตรวจสอบ error logs
   - รีเฟรชหน้าเว็บ
   - ตรวจสอบการเชื่อมต่ออินเทอร์เน็ต

4. **ไม่สามารถ Deploy**
   - ตรวจสอบการตั้งค่าสิทธิ์
   - ลองเปลี่ยนจาก "Anyone" เป็น "Anyone with Google account"

5. **Warning: cdn.tailwindcss.com should not be used in production**
   - สำหรับ Development ใช้ CDN ได้
   - สำหรับ Production ให้ติดตั้ง Tailwind CSS แบบ local

#### ขั้นตอนการ Debug:

1. **ตรวจสอบ Console Log**
   ```javascript
   // เปิด Developer Tools (F12)
   // ดูที่ Console tab
   console.log('Application initializing...');
   ```

2. **ตรวจสอบ Apps Script Execution**
   ```javascript
   // ใน Apps Script Editor
   // View > Execution Transcript
   ```

3. **ตรวจสอบข้อมูลใน Sheets**
   ```javascript
   // รันฟังก์ชันใน Apps Script Editor
   function testData() {
     console.log(getAllStudents());
     console.log(getTodayAttendance());
   }
   ```

### 📞 การสนับสนุน

หากมีปัญหาหรือต้องการความช่วยเหลือ:
1. ตรวจสอบ Console Log ในเบราว์เซอร์
2. ดู Execution Transcript ใน Apps Script
3. ตรวจสอบข้อมูลใน Google Sheets

### 🎯 การปรับแต่ง

#### เพิ่มข้อมูลนักเรียน:
แก้ไขใน Google Sheets หรือเพิ่มฟังก์ชันสำหรับ Admin

#### เปลี่ยนสี:
แก้ไขใน `tailwind.config` ในไฟล์ HTML

#### เพิ่มฟีเจอร์:
เพิ่มฟังก์ชันใน Code.gs และอัปเดต UI ใน HTML

---

**สร้างโดย:** GitHub Copilot  
**เวอร์ชัน:** 1.0  
**อัปเดตล่าสุด:** 13 มิถุนายน 2025
