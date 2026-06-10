# ตั้งค่า PostgreSQL

ระบบจะใช้ PostgreSQL อัตโนมัติเมื่อมีตัวแปร `DATABASE_URL` และจะใช้ SQLite เฉพาะกรณีเปิดทดสอบในเครื่องโดยยังไม่ได้ตั้งค่าตัวแปรนี้

## เปิดทดสอบในเครื่องด้วย SQLite

```bash
python3 -m venv .venv
source .venv/bin/activate
python3 -m pip install -r requirements.txt
python3 app.py
```

## เปิดด้วย PostgreSQL

สร้างฐานข้อมูล PostgreSQL ก่อน แล้วกำหนดค่า:

```bash
export DATABASE_URL='postgresql://postgres:password@localhost:5432/register_team'
export SECRET_KEY='เปลี่ยนเป็นข้อความสุ่มยาว ๆ'
python3 app.py
```

ระบบจะสร้างตารางและเพิ่มคอลัมน์ที่จำเป็นให้อัตโนมัติเมื่อเริ่มทำงาน

ตรวจสอบสถานะได้ที่:

```text
/health
```

ตัวอย่างผลลัพธ์:

```json
{"database":"postgresql","status":"ok"}
```

## ย้ายข้อมูลเดิมจาก SQLite ไป PostgreSQL

ก่อนย้ายข้อมูล ให้เปิดระบบด้วย PostgreSQL อย่างน้อยหนึ่งครั้ง เพื่อสร้างตาราง จากนั้นรัน:

```bash
export DATABASE_URL='postgresql://postgres:password@localhost:5432/register_team'
python3 migrate_sqlite_to_postgres.py
```

หากไฟล์ฐานข้อมูลเดิมไม่ได้อยู่ในโฟลเดอร์โปรเจกต์:

```bash
SQLITE_DB_PATH='/path/to/tournament_events.db' python3 migrate_sqlite_to_postgres.py
```

สคริปต์จะเพิ่มเฉพาะข้อมูลที่ยังไม่มี และไม่ลบข้อมูล PostgreSQL เดิม

## ไฟล์อัปโหลด

PostgreSQL เก็บเฉพาะข้อมูลรายการสมัคร ส่วนสลิป โลโก้ รูปลายเซ็น ตราประทับ และพื้นหลังเกียรติบัตรยังเป็นไฟล์ในโฟลเดอร์อัปโหลด

เมื่อ deploy จริง ควรกำหนด:

```bash
UPLOAD_FOLDER=/data/uploads
```

แล้ว mount persistent volume ให้ path นี้ เพื่อไม่ให้ไฟล์หายเมื่อ redeploy
