# คู่มืออัปโหลด GitHub และเปิดใช้งาน Streamlit ฟรี

## 1) แตกไฟล์ ZIP

แตกไฟล์โปรเจกต์ จะได้โฟลเดอร์

```text
donation_box_full_project_v2
```

## 2) สร้าง Repository ใน GitHub

1. เข้า https://github.com
2. กด New repository
3. ตั้งชื่อ เช่น `donation-box-app`
4. เลือก Public หรือ Private
5. กด Create repository

## 3) อัปโหลดไฟล์

เปิดโฟลเดอร์โปรเจกต์ แล้วลากไฟล์ทั้งหมดขึ้น GitHub

ต้องเห็นไฟล์เหล่านี้อยู่หน้าแรกของ Repository

```text
streamlit_app.py
requirements.txt
README.md
data/
docs/
.streamlit/
```

## 4) Deploy ผ่าน Streamlit Community Cloud

1. เข้า https://share.streamlit.io
2. Login ด้วย GitHub
3. กด New app
4. เลือก Repository ที่สร้างไว้
5. Main file path ใส่

```text
streamlit_app.py
```

6. กด Deploy

## 5) ตั้งค่า Secrets

ถ้าต้องการใช้ Google Sheets/Drive ให้เปิด Settings → Secrets แล้วใส่ค่าตาม `docs/GOOGLE_SETUP.md`

ถ้ายังไม่ใส่ Secrets ระบบจะใช้ CSV Local ก่อน เหมาะสำหรับทดสอบ

## 6) ใช้งานผ่านมือถือ

นำลิงก์ `.streamlit.app` ที่ได้ ไปเปิดบนมือถือได้ทันที

แนะนำให้เพิ่ม Shortcut บนหน้าจอมือถือ เพื่อใช้งานเหมือนแอพ
