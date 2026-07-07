# คู่มือเชื่อม Google Sheets และ Google Drive

ระบบนี้เก็บข้อมูลหลักใน Google Sheets และเก็บรูป/ไฟล์สำรองใน Google Drive ได้

## 1) สร้าง Google Sheet

สร้าง Google Sheet ใหม่ ตั้งชื่อ เช่น

```text
Donation Box Database
```

สร้างชีตย่อย 4 ชีต ชื่อตามนี้:

```text
boxes
collections
missing_reports
users
```

ระบบจะสร้างหัวตารางให้อัตโนมัติได้ ถ้ายังไม่มีข้อมูล

## 2) สร้างโฟลเดอร์ Google Drive สำหรับรูปภาพ

สร้างโฟลเดอร์ เช่น

```text
Donation Box Photos
```

คัดลอก Folder ID จาก URL เช่น

```text
https://drive.google.com/drive/folders/xxxxxxxxxxxx
```

ค่า `xxxxxxxxxxxx` คือ `GOOGLE_DRIVE_FOLDER_ID`

## 3) สร้าง Service Account

1. เข้า Google Cloud Console
2. สร้าง Project ใหม่
3. เปิด API:
   - Google Sheets API
   - Google Drive API
4. สร้าง Service Account
5. สร้าง Key แบบ JSON
6. ดาวน์โหลดไฟล์ JSON

## 4) แชร์ Google Sheet และ Drive Folder ให้ Service Account

ในไฟล์ JSON จะมีอีเมล เช่น

```text
xxxx@xxxx.iam.gserviceaccount.com
```

ให้นำอีเมลนี้ไป Share ใน Google Sheet และ Google Drive Folder โดยให้สิทธิ์ Editor

## 5) ตั้งค่า Streamlit Secrets

ใน Streamlit Cloud ไปที่ App → Settings → Secrets แล้วใส่รูปแบบนี้

```toml
APP_NAME = "ระบบเก็บกล่องบริจาค"
GOOGLE_SHEET_ID = "ใส่ Sheet ID"
GOOGLE_DRIVE_FOLDER_ID = "ใส่ Drive Folder ID"

[gcp_service_account]
type = "service_account"
project_id = "xxxxx"
private_key_id = "xxxxx"
private_key = "-----BEGIN PRIVATE KEY-----\nxxxxx\n-----END PRIVATE KEY-----\n"
client_email = "xxxxx@xxxxx.iam.gserviceaccount.com"
client_id = "xxxxx"
auth_uri = "https://accounts.google.com/o/oauth2/auth"
token_uri = "https://oauth2.googleapis.com/token"
auth_provider_x509_cert_url = "https://www.googleapis.com/oauth2/v1/certs"
client_x509_cert_url = "xxxxx"
universe_domain = "googleapis.com"
```

## 6) การเก็บข้อมูลใน Google Drive

- รายการกล่อง/ประวัติการเก็บ/แจ้งกล่องหาย/บัญชีผู้ใช้ จะบันทึกใน Google Sheets
- รูปภาพจากกล้องมือถือจะอัปโหลดเข้า Google Drive Folder
- ทุกครั้งที่บันทึกข้อมูล ระบบจะสำรองข้อมูลเป็น CSV เข้า Google Drive Folder ด้วย

## หมายเหตุความปลอดภัย

- อย่าอัปโหลดไฟล์ JSON Service Account ขึ้น GitHub
- ใส่ข้อมูล JSON เฉพาะใน Streamlit Secrets เท่านั้น
- เปลี่ยนรหัสผ่าน `admin / 1234` หลัง Deploy
