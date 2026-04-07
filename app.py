from flask import Flask, render_template, request, redirect, url_for, session, flash, send_from_directory, send_file
import sqlite3
from werkzeug.security import generate_password_hash, check_password_hash
from werkzeug.utils import secure_filename
from datetime import datetime
from io import BytesIO
import openpyxl
import os

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_NAME = os.path.join(BASE_DIR, "tournament_events.db")
UPLOAD_FOLDER = os.path.join(BASE_DIR, "static", "uploads")
ALLOWED_EXTENSIONS = {"png", "jpg", "jpeg", "pdf", "webp"}

app = Flask(__name__)
app.secret_key = "change-this-secret-key"
app.config["UPLOAD_FOLDER"] = UPLOAD_FOLDER
app.config["MAX_CONTENT_LENGTH"] = 10 * 1024 * 1024

os.makedirs(UPLOAD_FOLDER, exist_ok=True)


def get_db():
    conn = sqlite3.connect(DB_NAME)
    conn.row_factory = sqlite3.Row
    return conn


def init_db():
    conn = get_db()
    c = conn.cursor()

    c.execute("""
        CREATE TABLE IF NOT EXISTS users (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            username TEXT UNIQUE NOT NULL,
            password TEXT NOT NULL,
            role TEXT NOT NULL DEFAULT 'admin'
        )
    """)

    c.execute("""
        CREATE TABLE IF NOT EXISTS tournaments (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            title TEXT NOT NULL,
            description TEXT,
            created_by INTEGER,
            created_at TEXT,
            is_open INTEGER NOT NULL DEFAULT 1,
            FOREIGN KEY(created_by) REFERENCES users(id)
        )
    """)

    c.execute("""
        CREATE TABLE IF NOT EXISTS events (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            tournament_id INTEGER NOT NULL,
            event_name TEXT,
            category_type TEXT NOT NULL,
            gender_type TEXT NOT NULL,
            age_group TEXT NOT NULL,
            max_slots INTEGER NOT NULL,
            fee INTEGER NOT NULL DEFAULT 0,
            team_size INTEGER NOT NULL DEFAULT 1,
            is_open INTEGER NOT NULL DEFAULT 1,
            created_at TEXT,
            FOREIGN KEY(tournament_id) REFERENCES tournaments(id)
        )
    """)

    c.execute("""
        CREATE TABLE IF NOT EXISTS registrations (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            event_id INTEGER NOT NULL,
            team_name TEXT,
            contact_name TEXT NOT NULL,
            phone TEXT NOT NULL,
            slip_filename TEXT,
            notes TEXT,
            created_at TEXT,
            FOREIGN KEY(event_id) REFERENCES events(id)
        )
    """)

    c.execute("""
        CREATE TABLE IF NOT EXISTS registration_members (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            registration_id INTEGER NOT NULL,
            member_name TEXT NOT NULL,
            member_idcard TEXT,
            idcard_file TEXT,
            FOREIGN KEY(registration_id) REFERENCES registrations(id)
        )
    """)

    # migration: add missing columns for old databases
    reg_cols = [row[1] for row in c.execute("PRAGMA table_info(registrations)").fetchall()]
    if "slip_filename" not in reg_cols:
        c.execute("ALTER TABLE registrations ADD COLUMN slip_filename TEXT")
    if "notes" not in reg_cols:
        c.execute("ALTER TABLE registrations ADD COLUMN notes TEXT")
    if "created_at" not in reg_cols:
        c.execute("ALTER TABLE registrations ADD COLUMN created_at TEXT")

    member_cols = [row[1] for row in c.execute("PRAGMA table_info(registration_members)").fetchall()]
    if "idcard_file" not in member_cols:
        c.execute("ALTER TABLE registration_members ADD COLUMN idcard_file TEXT")

    existing = c.execute("SELECT id FROM users WHERE username = ?", ("admin",)).fetchone()
    if not existing:
        c.execute(
            "INSERT INTO users (username, password, role) VALUES (?, ?, ?)",
            ("admin", generate_password_hash("1234"), "admin")
        )

    conn.commit()
    conn.close()


def is_logged_in():
    return "user_id" in session


def allowed_file(filename):
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS


def category_label(value):
    return {"single": "เดี่ยว", "pair": "คู่", "team": "ทีม"}.get(value, value)


def gender_label(value):
    return {"male": "ชาย", "female": "หญิง", "mixed": "ผสม"}.get(value, value)


def age_label(value):
    return {"youth": "เยาวชน", "general": "ทั่วไป", "senior": "อาวุโส"}.get(value, value)


def event_display_name(event):
    custom = (event["event_name"] or "").strip()
    if custom:
        return custom
    return f"{category_label(event['category_type'])} {gender_label(event['gender_type'])} {age_label(event['age_group'])}"


def get_team_size(category_type, team_size):
    if category_type == "single":
        return 1
    if category_type == "pair":
        return 2
    return max(1, int(team_size or 1))


def event_reg_count(event_id):
    conn = get_db()
    total = conn.execute(
        "SELECT COUNT(*) AS total FROM registrations WHERE event_id = ?",
        (event_id,)
    ).fetchone()["total"]
    conn.close()
    return total


def save_uploaded_file(file_obj, prefix="file"):
    if not file_obj or not file_obj.filename:
        return None
    if not allowed_file(file_obj.filename):
        return None
    safe_name = secure_filename(file_obj.filename)
    ext = safe_name.rsplit(".", 1)[1].lower()
    filename = f"{prefix}_{datetime.now().strftime('%Y%m%d%H%M%S%f')}.{ext}"
    file_obj.save(os.path.join(app.config["UPLOAD_FOLDER"], filename))
    return filename


def delete_uploaded_file(filename):
    if not filename:
        return
    path = os.path.join(app.config["UPLOAD_FOLDER"], filename)
    if os.path.exists(path):
        os.remove(path)


def registration_summary_by_event(tournament_id):
    conn = get_db()
    events = conn.execute(
        "SELECT * FROM events WHERE tournament_id = ? ORDER BY id ASC",
        (tournament_id,)
    ).fetchall()
    rows = []
    total = 0
    for e in events:
        count = conn.execute(
            "SELECT COUNT(*) AS total FROM registrations WHERE event_id = ?",
            (e["id"],)
        ).fetchone()["total"]
        total += count
        rows.append({"event": e, "count": count})
    conn.close()
    return rows, total


@app.context_processor
def inject_helpers():
    return {
        "category_label": category_label,
        "gender_label": gender_label,
        "age_label": age_label,
        "event_display_name": event_display_name,
    }


@app.route("/uploads/<path:filename>")
def uploaded_file(filename):
    return send_from_directory(app.config["UPLOAD_FOLDER"], filename)


@app.route("/")
def home():
    conn = get_db()
    tournaments = conn.execute("SELECT * FROM tournaments ORDER BY id DESC").fetchall()
    event_map = {}
    count_map = {}
    for t in tournaments:
        events = conn.execute(
            "SELECT * FROM events WHERE tournament_id = ? ORDER BY id ASC",
            (t["id"],)
        ).fetchall()
        event_map[t["id"]] = events
        count_map[t["id"]] = {e["id"]: event_reg_count(e["id"]) for e in events}
    conn.close()
    return render_template("home.html", tournaments=tournaments, event_map=event_map, count_map=count_map)


@app.route("/event/<int:event_id>/register", methods=["GET", "POST"])
def register_event(event_id):
    conn = get_db()
    event = conn.execute("SELECT * FROM events WHERE id = ?", (event_id,)).fetchone()
    if not event:
        conn.close()
        flash("ไม่พบอีเวนต์")
        return redirect(url_for("home"))

    tournament = conn.execute("SELECT * FROM tournaments WHERE id = ?", (event["tournament_id"],)).fetchone()
    conn.close()

    reg_count = event_reg_count(event_id)
    member_count = get_team_size(event["category_type"], event["team_size"])

    if request.method == "POST":
        if not event["is_open"] or not tournament["is_open"]:
            flash("อีเวนต์นี้ปิดรับสมัครแล้ว")
            return redirect(url_for("register_event", event_id=event_id))

        if reg_count >= event["max_slots"]:
            flash("อีเวนต์นี้เต็มแล้ว")
            return redirect(url_for("register_event", event_id=event_id))

        contact_name = request.form.get("contact_name", "").strip()
        phone = request.form.get("phone", "").strip()
        notes = request.form.get("notes", "").strip()
        team_name = request.form.get("team_name", "").strip()

        members = []
        for i in range(1, member_count + 1):
            member_name = request.form.get(f"member_name_{i}", "").strip()
            member_idcard = request.form.get(f"member_idcard_{i}", "").strip()
            member_idcard_file = request.files.get(f"idcard_file_{i}")

            if member_name:
                idcard_file_name = None
                if member_idcard_file and member_idcard_file.filename:
                    idcard_file_name = save_uploaded_file(member_idcard_file, prefix=f"idcard_{i}")
                    if not idcard_file_name:
                        flash("อัปโหลดบัตรประชาชนได้เฉพาะไฟล์ png, jpg, jpeg, webp, pdf")
                        return redirect(url_for("register_event", event_id=event_id))
                members.append((member_name, member_idcard, idcard_file_name))

        if not contact_name or not phone:
            flash("กรุณากรอกชื่อผู้ติดต่อและเบอร์โทร")
            return redirect(url_for("register_event", event_id=event_id))

        if len(members) != member_count:
            flash(f"กรุณากรอกข้อมูลสมาชิกให้ครบ {member_count} คน")
            return redirect(url_for("register_event", event_id=event_id))

        if event["category_type"] == "team" and not team_name:
            flash("ประเภททีมต้องกรอกชื่อทีม")
            return redirect(url_for("register_event", event_id=event_id))

        slip_file = request.files.get("slip_file")
        slip_filename = None
        if slip_file and slip_file.filename:
            slip_filename = save_uploaded_file(slip_file, prefix="slip")
            if not slip_filename:
                flash("อัปโหลดสลิปได้เฉพาะไฟล์ png, jpg, jpeg, webp, pdf")
                return redirect(url_for("register_event", event_id=event_id))

        conn = get_db()
        c = conn.cursor()
        c.execute(
            """INSERT INTO registrations (event_id, team_name, contact_name, phone, slip_filename, notes, created_at)
               VALUES (?, ?, ?, ?, ?, ?, ?)""",
            (
                event_id,
                team_name or None,
                contact_name,
                phone,
                slip_filename,
                notes,
                datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            )
        )
        registration_id = c.lastrowid

        for member_name, member_idcard, idcard_file_name in members:
            c.execute(
                "INSERT INTO registration_members (registration_id, member_name, member_idcard, idcard_file) VALUES (?, ?, ?, ?)",
                (registration_id, member_name, member_idcard, idcard_file_name)
            )

        conn.commit()
        conn.close()
        flash("สมัครสำเร็จแล้ว")
        return redirect(url_for("home"))

    return render_template(
        "register_event.html",
        event=event,
        tournament=tournament,
        reg_count=reg_count,
        member_count=member_count,
    )


@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        username = request.form.get("username", "").strip()
        password = request.form.get("password", "").strip()

        conn = get_db()
        user = conn.execute("SELECT * FROM users WHERE username = ?", (username,)).fetchone()
        conn.close()

        if user and check_password_hash(user["password"], password):
            session["user_id"] = user["id"]
            session["username"] = user["username"]
            session["role"] = user["role"]
            flash("เข้าสู่ระบบสำเร็จ")
            return redirect(url_for("admin_dashboard"))
        flash("ชื่อผู้ใช้หรือรหัสผ่านไม่ถูกต้อง")

    return render_template("login.html")


@app.route("/logout")
def logout():
    session.clear()
    flash("ออกจากระบบแล้ว")
    return redirect(url_for("home"))


@app.route("/admin")
def admin_dashboard():
    if not is_logged_in():
        flash("กรุณาเข้าสู่ระบบก่อน")
        return redirect(url_for("login"))

    conn = get_db()
    tournaments = conn.execute(
        "SELECT * FROM tournaments WHERE created_by = ? ORDER BY id DESC",
        (session["user_id"],)
    ).fetchall()

    event_map = {}
    count_map = {}
    dashboard = {
        "tournaments": len(tournaments),
        "events": 0,
        "registrations": 0,
        "open_events": 0,
    }

    for t in tournaments:
        events = conn.execute(
            "SELECT * FROM events WHERE tournament_id = ? ORDER BY id ASC",
            (t["id"],)
        ).fetchall()
        event_map[t["id"]] = events
        count_map[t["id"]] = {}
        dashboard["events"] += len(events)
        for e in events:
            c = event_reg_count(e["id"])
            count_map[t["id"]][e["id"]] = c
            dashboard["registrations"] += c
            if e["is_open"]:
                dashboard["open_events"] += 1

    conn.close()
    return render_template("admin_dashboard.html", tournaments=tournaments, event_map=event_map, count_map=count_map, dashboard=dashboard)


@app.route("/admin/tournament/create", methods=["GET", "POST"])
def create_tournament():
    if not is_logged_in():
        flash("กรุณาเข้าสู่ระบบก่อน")
        return redirect(url_for("login"))

    if request.method == "POST":
        title = request.form.get("title", "").strip()
        description = request.form.get("description", "").strip()
        is_open = 1 if request.form.get("is_open") == "on" else 0

        if not title:
            flash("กรุณากรอกชื่องานแข่งขัน")
            return redirect(url_for("create_tournament"))

        conn = get_db()
        c = conn.cursor()
        c.execute(
            """INSERT INTO tournaments (title, description, created_by, created_at, is_open)
               VALUES (?, ?, ?, ?, ?)""",
            (title, description, session["user_id"], datetime.now().strftime("%Y-%m-%d %H:%M:%S"), is_open)
        )
        tournament_id = c.lastrowid
        conn.commit()
        conn.close()

        flash("สร้างงานแข่งขันเรียบร้อยแล้ว")
        return redirect(url_for("manage_events", tournament_id=tournament_id))

    return render_template("create_tournament.html")



@app.route("/admin/tournament/<int:tournament_id>/edit", methods=["GET", "POST"])
def edit_tournament(tournament_id):
    if not is_logged_in():
        flash("กรุณาเข้าสู่ระบบก่อน")
        return redirect(url_for("login"))

    conn = get_db()
    tournament = conn.execute(
        "SELECT * FROM tournaments WHERE id = ? AND created_by = ?",
        (tournament_id, session["user_id"])
    ).fetchone()

    if not tournament:
        conn.close()
        flash("ไม่พบงานแข่งขัน")
        return redirect(url_for("admin_dashboard"))

    if request.method == "POST":
        title = request.form.get("title", "").strip()
        description = request.form.get("description", "").strip()
        is_open = 1 if request.form.get("is_open") == "on" else 0

        if not title:
            conn.close()
            flash("กรุณากรอกชื่องานแข่งขัน")
            return redirect(url_for("edit_tournament", tournament_id=tournament_id))

        conn.execute(
            "UPDATE tournaments SET title = ?, description = ?, is_open = ? WHERE id = ?",
            (title, description, is_open, tournament_id)
        )
        conn.commit()
        conn.close()

        flash("แก้ไขทัวร์นาเมนต์เรียบร้อยแล้ว")
        return redirect(url_for("admin_dashboard"))

    conn.close()
    return render_template("edit_tournament.html", tournament=tournament)


@app.route("/admin/tournament/<int:tournament_id>/delete")
def delete_tournament(tournament_id):
    if not is_logged_in():
        flash("กรุณาเข้าสู่ระบบก่อน")
        return redirect(url_for("login"))

    conn = get_db()
    tournament = conn.execute(
        "SELECT * FROM tournaments WHERE id = ? AND created_by = ?",
        (tournament_id, session["user_id"])
    ).fetchone()

    if not tournament:
        conn.close()
        flash("ไม่พบงานแข่งขัน")
        return redirect(url_for("admin_dashboard"))

    events = conn.execute(
        "SELECT id FROM events WHERE tournament_id = ?",
        (tournament_id,)
    ).fetchall()

    for event in events:
        regs = conn.execute(
            "SELECT * FROM registrations WHERE event_id = ?",
            (event["id"],)
        ).fetchall()

        for reg in regs:
            members = conn.execute(
                "SELECT idcard_file FROM registration_members WHERE registration_id = ?",
                (reg["id"],)
            ).fetchall()
            for m in members:
                idcard_file = m["idcard_file"] if "idcard_file" in m.keys() else ""
                delete_uploaded_file(idcard_file)
            slip_filename = reg["slip_filename"] if "slip_filename" in reg.keys() else ""
            delete_uploaded_file(slip_filename)
            conn.execute("DELETE FROM registration_members WHERE registration_id = ?", (reg["id"],))
        conn.execute("DELETE FROM registrations WHERE event_id = ?", (event["id"],))

    conn.execute("DELETE FROM events WHERE tournament_id = ?", (tournament_id,))
    conn.execute("DELETE FROM tournaments WHERE id = ?", (tournament_id,))
    conn.commit()
    conn.close()

    flash("ลบทัวร์นาเมนต์เรียบร้อยแล้ว")
    return redirect(url_for("admin_dashboard"))


@app.route("/admin/tournament/<int:tournament_id>/events")
def manage_events(tournament_id):
    if not is_logged_in():
        flash("กรุณาเข้าสู่ระบบก่อน")
        return redirect(url_for("login"))

    conn = get_db()
    tournament = conn.execute(
        "SELECT * FROM tournaments WHERE id = ? AND created_by = ?",
        (tournament_id, session["user_id"])
    ).fetchone()

    if not tournament:
        conn.close()
        flash("ไม่พบงานแข่งขัน")
        return redirect(url_for("admin_dashboard"))

    events = conn.execute(
        "SELECT * FROM events WHERE tournament_id = ? ORDER BY id ASC",
        (tournament_id,)
    ).fetchall()
    counts = {e["id"]: event_reg_count(e["id"]) for e in events}
    summary_rows, total_regs = registration_summary_by_event(tournament_id)
    conn.close()
    return render_template("manage_events.html", tournament=tournament, events=events, counts=counts, summary_rows=summary_rows, total_regs=total_regs)


@app.route("/admin/tournament/<int:tournament_id>/event/create", methods=["GET", "POST"])
def create_event(tournament_id):
    if not is_logged_in():
        flash("กรุณาเข้าสู่ระบบก่อน")
        return redirect(url_for("login"))

    conn = get_db()
    tournament = conn.execute(
        "SELECT * FROM tournaments WHERE id = ? AND created_by = ?",
        (tournament_id, session["user_id"])
    ).fetchone()

    if not tournament:
        conn.close()
        flash("ไม่พบงานแข่งขัน")
        return redirect(url_for("admin_dashboard"))

    if request.method == "POST":
        event_name = request.form.get("event_name", "").strip()
        category_type = request.form.get("category_type", "single").strip()
        gender_type = request.form.get("gender_type", "male").strip()
        age_group = request.form.get("age_group", "general").strip()
        max_slots = request.form.get("max_slots", "0").strip()
        fee = request.form.get("fee", "0").strip()
        team_size = request.form.get("team_size", "1").strip()
        is_open = 1 if request.form.get("is_open") == "on" else 0

        if not max_slots.isdigit():
            conn.close()
            flash("กรุณากรอกจำนวนรับสมัครให้ถูกต้อง")
            return redirect(url_for("create_event", tournament_id=tournament_id))

        if not fee.isdigit():
            fee = "0"

        if category_type == "single":
            team_size = 1
        elif category_type == "pair":
            team_size = 2
        else:
            team_size = int(team_size) if str(team_size).isdigit() and int(team_size) > 0 else 3

        conn.execute(
            """INSERT INTO events
               (tournament_id, event_name, category_type, gender_type, age_group, max_slots, fee, team_size, is_open, created_at)
               VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
            (
                tournament_id,
                event_name,
                category_type,
                gender_type,
                age_group,
                int(max_slots),
                int(fee),
                int(team_size),
                is_open,
                datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            )
        )
        conn.commit()
        conn.close()

        flash("เพิ่มอีเวนต์เรียบร้อยแล้ว")
        return redirect(url_for("manage_events", tournament_id=tournament_id))

    conn.close()
    return render_template("create_event.html", tournament=tournament)


@app.route("/admin/event/<int:event_id>/edit", methods=["GET", "POST"])
def edit_event(event_id):
    if not is_logged_in():
        flash("กรุณาเข้าสู่ระบบก่อน")
        return redirect(url_for("login"))

    conn = get_db()
    event = conn.execute(
        """SELECT e.*, t.created_by
           FROM events e
           JOIN tournaments t ON e.tournament_id = t.id
           WHERE e.id = ?""",
        (event_id,)
    ).fetchone()

    if not event or event["created_by"] != session["user_id"]:
        conn.close()
        flash("ไม่พบอีเวนต์")
        return redirect(url_for("admin_dashboard"))

    if request.method == "POST":
        event_name = request.form.get("event_name", "").strip()
        category_type = request.form.get("category_type", "single").strip()
        gender_type = request.form.get("gender_type", "male").strip()
        age_group = request.form.get("age_group", "general").strip()
        max_slots = request.form.get("max_slots", "0").strip()
        fee = request.form.get("fee", "0").strip()
        team_size = request.form.get("team_size", "1").strip()
        is_open = 1 if request.form.get("is_open") == "on" else 0

        if not max_slots.isdigit():
            conn.close()
            flash("กรุณากรอกจำนวนรับสมัครให้ถูกต้อง")
            return redirect(url_for("edit_event", event_id=event_id))

        if not fee.isdigit():
            fee = "0"

        if category_type == "single":
            team_size = 1
        elif category_type == "pair":
            team_size = 2
        else:
            team_size = int(team_size) if str(team_size).isdigit() and int(team_size) > 0 else 3

        conn.execute(
            """UPDATE events
               SET event_name = ?, category_type = ?, gender_type = ?, age_group = ?, max_slots = ?, fee = ?, team_size = ?, is_open = ?
               WHERE id = ?""",
            (event_name, category_type, gender_type, age_group, int(max_slots), int(fee), int(team_size), is_open, event_id)
        )
        conn.commit()
        tournament_id = event["tournament_id"]
        conn.close()

        flash("แก้ไขอีเวนต์เรียบร้อยแล้ว")
        return redirect(url_for("manage_events", tournament_id=tournament_id))

    tournament = conn.execute("SELECT * FROM tournaments WHERE id = ?", (event["tournament_id"],)).fetchone()
    conn.close()
    return render_template("edit_event.html", tournament=tournament, event=event)


@app.route("/admin/event/<int:event_id>/delete")
def delete_event(event_id):
    if not is_logged_in():
        flash("กรุณาเข้าสู่ระบบก่อน")
        return redirect(url_for("login"))

    conn = get_db()
    event = conn.execute(
        """SELECT e.*, t.created_by
           FROM events e
           JOIN tournaments t ON e.tournament_id = t.id
           WHERE e.id = ?""",
        (event_id,)
    ).fetchone()

    if not event or event["created_by"] != session["user_id"]:
        conn.close()
        flash("ไม่พบอีเวนต์")
        return redirect(url_for("admin_dashboard"))

    regs = conn.execute("SELECT * FROM registrations WHERE event_id = ?", (event_id,)).fetchall()
    for reg in regs:
        members = conn.execute("SELECT idcard_file FROM registration_members WHERE registration_id = ?", (reg["id"],)).fetchall()
        for m in members:
            idcard_file = m["idcard_file"] if "idcard_file" in m.keys() else ""
            delete_uploaded_file(idcard_file)
        slip_filename = reg["slip_filename"] if "slip_filename" in reg.keys() else ""
        delete_uploaded_file(slip_filename)
        conn.execute("DELETE FROM registration_members WHERE registration_id = ?", (reg["id"],))
    conn.execute("DELETE FROM registrations WHERE event_id = ?", (event_id,))
    conn.execute("DELETE FROM events WHERE id = ?", (event_id,))
    conn.commit()
    tournament_id = event["tournament_id"]
    conn.close()

    flash("ลบอีเวนต์เรียบร้อยแล้ว")
    return redirect(url_for("manage_events", tournament_id=tournament_id))


@app.route("/admin/tournament/<int:tournament_id>/registrations")
def tournament_registrations(tournament_id):
    if not is_logged_in():
        flash("กรุณาเข้าสู่ระบบก่อน")
        return redirect(url_for("login"))

    conn = get_db()
    tournament = conn.execute(
        "SELECT * FROM tournaments WHERE id = ? AND created_by = ?",
        (tournament_id, session["user_id"])
    ).fetchone()

    if not tournament:
        conn.close()
        flash("ไม่พบงานแข่งขัน")
        return redirect(url_for("admin_dashboard"))

    rows = conn.execute(
        """SELECT r.*, e.event_name, e.category_type, e.gender_type, e.age_group, e.fee
           FROM registrations r
           JOIN events e ON r.event_id = e.id
           WHERE e.tournament_id = ?
           ORDER BY e.id ASC, r.id DESC""",
        (tournament_id,)
    ).fetchall()

    members_map = {}
    for row in rows:
        members_map[row["id"]] = conn.execute(
            "SELECT * FROM registration_members WHERE registration_id = ? ORDER BY id ASC",
            (row["id"],)
        ).fetchall()

    summary_rows, total_regs = registration_summary_by_event(tournament_id)
    conn.close()
    return render_template("tournament_registrations.html", tournament=tournament, rows=rows, members_map=members_map, summary_rows=summary_rows, total_regs=total_regs)


@app.route("/admin/event/<int:event_id>/export")
def export_event_excel(event_id):
    if not is_logged_in():
        flash("กรุณาเข้าสู่ระบบก่อน")
        return redirect(url_for("login"))

    conn = get_db()
    event = conn.execute(
        """SELECT e.*, t.title AS tournament_title, t.created_by
           FROM events e
           JOIN tournaments t ON e.tournament_id = t.id
           WHERE e.id = ?""",
        (event_id,)
    ).fetchone()

    if not event or event["created_by"] != session["user_id"]:
        conn.close()
        flash("ไม่พบอีเวนต์")
        return redirect(url_for("admin_dashboard"))

    regs = conn.execute(
        "SELECT * FROM registrations WHERE event_id = ? ORDER BY id ASC",
        (event_id,)
    ).fetchall()

    member_map = {}
    max_members = 0
    for reg in regs:
        members = conn.execute(
            "SELECT * FROM registration_members WHERE registration_id = ? ORDER BY id ASC",
            (reg["id"],)
        ).fetchall()
        member_map[reg["id"]] = members
        max_members = max(max_members, len(members))
    conn.close()

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "รายชื่อสมัคร"

    headers = [
        "ลำดับ", "งานแข่งขัน", "อีเวนต์", "ประเภท", "เพศ", "รุ่น",
        "ค่าสมัคร", "ชื่อทีม", "ผู้ติดต่อ", "เบอร์โทร", "สลิป", "หมายเหตุ", "สมัครเมื่อ"
    ]
    for i in range(1, max_members + 1):
        headers.extend([f"สมาชิก {i}", f"เลขบัตรสมาชิก {i}", f"ไฟล์บัตรสมาชิก {i}"])
    ws.append(headers)

    event_name = event_display_name(event)
    for idx, reg in enumerate(regs, start=1):
        team_name = reg["team_name"] if "team_name" in reg.keys() else ""
        contact_name = reg["contact_name"] if "contact_name" in reg.keys() else ""
        phone = reg["phone"] if "phone" in reg.keys() else ""
        slip_filename = reg["slip_filename"] if "slip_filename" in reg.keys() else ""
        notes = reg["notes"] if "notes" in reg.keys() else ""
        created_at = reg["created_at"] if "created_at" in reg.keys() else ""

        row = [
            idx,
            event["tournament_title"],
            event_name,
            category_label(event["category_type"]),
            gender_label(event["gender_type"]),
            age_label(event["age_group"]),
            event["fee"],
            team_name or "",
            contact_name or "",
            phone or "",
            slip_filename or "",
            notes or "",
            created_at or "",
        ]
        members = member_map[reg["id"]]
        for m in members:
            member_name = m["member_name"] if "member_name" in m.keys() else ""
            member_idcard = m["member_idcard"] if "member_idcard" in m.keys() else ""
            idcard_file_value = m["idcard_file"] if "idcard_file" in m.keys() else ""
            row.extend([member_name or "", member_idcard or "", idcard_file_value or ""])
        for _ in range(max_members - len(members)):
            row.extend(["", "", ""])
        ws.append(row)

    for col in ws.columns:
        max_len = 0
        letter = col[0].column_letter
        for cell in col:
            val = str(cell.value) if cell.value is not None else ""
            max_len = max(max_len, len(val))
        ws.column_dimensions[letter].width = min(max_len + 2, 30)

    output = BytesIO()
    wb.save(output)
    output.seek(0)

    return send_file(
        output,
        as_attachment=True,
        download_name=f"event_{event_id}_registrations.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


@app.route("/admin/registration/<int:registration_id>/delete")
def delete_registration(registration_id):
    if not is_logged_in():
        flash("กรุณาเข้าสู่ระบบก่อน")
        return redirect(url_for("login"))

    conn = get_db()
    row = conn.execute(
        """SELECT r.*, e.tournament_id, t.created_by
           FROM registrations r
           JOIN events e ON r.event_id = e.id
           JOIN tournaments t ON e.tournament_id = t.id
           WHERE r.id = ?""",
        (registration_id,)
    ).fetchone()

    if not row or row["created_by"] != session["user_id"]:
        conn.close()
        flash("ไม่พบข้อมูลผู้สมัคร")
        return redirect(url_for("admin_dashboard"))

    members = conn.execute(
        "SELECT idcard_file FROM registration_members WHERE registration_id = ?",
        (registration_id,)
    ).fetchall()
    for m in members:
        idcard_file = m["idcard_file"] if "idcard_file" in m.keys() else ""
        delete_uploaded_file(idcard_file)
    slip_filename = row["slip_filename"] if "slip_filename" in row.keys() else ""
    delete_uploaded_file(slip_filename)

    conn.execute("DELETE FROM registration_members WHERE registration_id = ?", (registration_id,))
    conn.execute("DELETE FROM registrations WHERE id = ?", (registration_id,))
    conn.commit()
    tournament_id = row["tournament_id"]
    conn.close()

    flash("ลบผู้สมัครเรียบร้อยแล้ว")
    return redirect(url_for("tournament_registrations", tournament_id=tournament_id))


if __name__ == "__main__":
    init_db()
    port = int(os.environ.get("PORT", 8000))
    app.run(host="0.0.0.0", port=port, debug=True)
