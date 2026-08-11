import os, sys, logging, asyncio
from datetime import datetime, time, timedelta
from io import BytesIO
from zoneinfo import ZoneInfo
from contextlib import contextmanager

import psycopg2, psycopg2.extras
from psycopg2.pool import ThreadedConnectionPool
from aiohttp import web
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup, ReplyKeyboardMarkup, KeyboardButton, ReplyKeyboardRemove
from telegram.ext import Application, CommandHandler, MessageHandler, CallbackQueryHandler, filters, ContextTypes

# ----------------------------------------------------------------------
# Environment variables
# ----------------------------------------------------------------------
BOT_TOKEN = os.environ.get("BOT_TOKEN")
ADMIN_GROUP_ID_STR = os.environ.get("ADMIN_GROUP_ID")
ADMIN_IDS_STR = os.environ.get("ADMIN_IDS", "")
DATABASE_URL = os.environ.get("DATABASE_URL")
WEBHOOK_URL = os.environ.get("RENDER_EXTERNAL_URL")

TOPIC_RECLAMATIONS = int(os.environ.get("TOPIC_RECLAMATIONS", "0"))
TOPIC_VALIDATION = int(os.environ.get("TOPIC_VALIDATION", "0"))
TOPIC_VIDANGE = int(os.environ.get("TOPIC_VIDANGE", "0"))
TOPIC_VEHICLE_MGMT = int(os.environ.get("TOPIC_VEHICLE_MGMT", "0"))
TOPIC_GENERAL = int(os.environ.get("TOPIC_GENERAL", "0"))
TOPIC_HISTORY = int(os.environ.get("TOPIC_HISTORY", "0"))

if not all([BOT_TOKEN, ADMIN_GROUP_ID_STR, DATABASE_URL, WEBHOOK_URL]):
    sys.exit("FATAL: BOT_TOKEN, ADMIN_GROUP_ID, DATABASE_URL and WEBHOOK_URL must be set.")
if not all([TOPIC_RECLAMATIONS, TOPIC_VALIDATION, TOPIC_VIDANGE, TOPIC_VEHICLE_MGMT, TOPIC_GENERAL, TOPIC_HISTORY]):
    sys.exit("FATAL: All TOPIC_* environment variables must be set.")

try:
    ADMIN_GROUP_ID = int(ADMIN_GROUP_ID_STR)
except ValueError:
    sys.exit("FATAL: ADMIN_GROUP_ID must be an integer.")

ADMIN_IDS = set()
if ADMIN_IDS_STR:
    for uid in ADMIN_IDS_STR.split(","):
        uid = uid.strip()
        if uid.isdigit():
            ADMIN_IDS.add(int(uid))

TZ = ZoneInfo("Africa/Algiers")

# ----------------------------------------------------------------------
# Connection pool
# ----------------------------------------------------------------------
pool = None
def get_conn():
    global pool
    if pool is None:
        pool = ThreadedConnectionPool(2, 4, DATABASE_URL, sslmode='require')
    return pool.getconn()

def put_conn(conn):
    global pool
    if pool:
        pool.putconn(conn)

@contextmanager
def db_connection():
    conn = get_conn()
    try:
        yield conn
    finally:
        put_conn(conn)

# ----------------------------------------------------------------------
# Database initialization
# ----------------------------------------------------------------------
def init_db():
    with db_connection() as conn:
        cur = conn.cursor()
        cur.execute('''
            CREATE TABLE IF NOT EXISTS drivers (
                user_id BIGINT PRIMARY KEY,
                name TEXT,
                vehicle TEXT,
                state TEXT,
                approval_status TEXT DEFAULT 'pending'
            )
        ''')
        cur.execute('''
            CREATE TABLE IF NOT EXISTS vehicles (
                code TEXT PRIMARY KEY
            )
        ''')
        cur.execute("SELECT COUNT(*) FROM vehicles")
        if cur.fetchone()[0] == 0:
            for v in ["F01","F02","H01"] + [f"M{i:02d}" for i in range(1,32)] + ["LOGAN"]:
                cur.execute("INSERT INTO vehicles (code) VALUES (%s) ON CONFLICT DO NOTHING", (v,))
        cur.execute('''
            CREATE TABLE IF NOT EXISTS problems (
                id SERIAL PRIMARY KEY,
                user_id BIGINT,
                driver_name TEXT,
                vehicle TEXT,
                problem_text TEXT,
                media_type TEXT,
                date TEXT,
                status TEXT DEFAULT 'قيد الانتظار',
                ruglee TEXT DEFAULT 'غير مُصلح',
                comments TEXT DEFAULT '',
                validation_requester BIGINT DEFAULT 0,
                group_message_id BIGINT DEFAULT 0
            )
        ''')
        cur.execute("ALTER TABLE problems ADD COLUMN IF NOT EXISTS validation_requester BIGINT DEFAULT 0")
        cur.execute("ALTER TABLE problems ADD COLUMN IF NOT EXISTS group_message_id BIGINT DEFAULT 0")
        cur.execute('''
            CREATE TABLE IF NOT EXISTS allowed_users (
                user_id BIGINT PRIMARY KEY,
                status TEXT DEFAULT 'approved'
            )
        ''')
        cur.execute('''
            CREATE TABLE IF NOT EXISTS help_videos (
                id SERIAL PRIMARY KEY,
                file_id TEXT,
                description TEXT DEFAULT ''
            )
        ''')
        cur.execute('''
            CREATE TABLE IF NOT EXISTS km_readings (
                id SERIAL PRIMARY KEY,
                vehicle TEXT,
                km INTEGER,
                date TEXT,
                driver_name TEXT DEFAULT ''
            )
        ''')
        cur.execute('''
            CREATE TABLE IF NOT EXISTS vehicle_vidange (
                vehicle TEXT PRIMARY KEY,
                last_vidange_km INTEGER DEFAULT 0
            )
        ''')
        cur.execute("SELECT code FROM vehicles")
        for v in [r[0] for r in cur.fetchall()]:
            cur.execute("INSERT INTO vehicle_vidange (vehicle, last_vidange_km) VALUES (%s, 0) ON CONFLICT DO NOTHING", (v,))
        conn.commit()

# ----------------------------------------------------------------------
# Cached vehicle status (fixed last_vid)
# ----------------------------------------------------------------------
vehicle_cache = {}
cache_dirty = set()

def invalidate_cache(vehicle: str):
    cache_dirty.add(vehicle)
    vehicle_cache.pop(vehicle, None)

def refresh_cache(vehicle: str):
    with db_connection() as conn:
        cur = conn.cursor()
        last_km = None
        last_vid = 0
        # Check open "قيد الانتظار" problems first
        cur.execute("SELECT COUNT(*) FROM problems WHERE vehicle=%s AND status='قيد الانتظار' AND ruglee != 'تم الإصلاح'", (vehicle,))
        if cur.fetchone()[0] > 0:
            status = 'bad'
        else:
            cur.execute("SELECT km FROM km_readings WHERE vehicle=%s ORDER BY date DESC LIMIT 1", (vehicle,))
            row = cur.fetchone()
            last_km = row[0] if row else None
            last_vid = get_last_vidange_km_noconn(conn, vehicle)
            if last_km and last_vid > 0 and last_km >= last_vid + 9000:
                status = 'bad'
            else:
                cur.execute("SELECT COUNT(*) FROM problems WHERE vehicle=%s AND status='قيد التصليح' AND ruglee != 'تم الإصلاح'", (vehicle,))
                if cur.fetchone()[0] > 0:
                    status = 'en_cours'
                else:
                    status = 'good'
        cur.execute("SELECT COUNT(*) FROM problems WHERE vehicle=%s AND ruglee != 'تم الإصلاح'", (vehicle,))
        open_count = cur.fetchone()[0]
        if last_km and last_vid > 0:
            remaining = (last_vid + 10000) - last_km
        else:
            remaining = None
        vehicle_cache[vehicle] = {
            "status": status,
            "open_count": open_count,
            "remaining_km": remaining
        }

def get_vehicle_status_cached(vehicle: str) -> str:
    if vehicle in cache_dirty:
        refresh_cache(vehicle)
        cache_dirty.discard(vehicle)
    if vehicle not in vehicle_cache:
        refresh_cache(vehicle)
    return vehicle_cache[vehicle]["status"]

def count_open_problems_cached(vehicle: str) -> int:
    if vehicle in cache_dirty:
        refresh_cache(vehicle)
        cache_dirty.discard(vehicle)
    if vehicle not in vehicle_cache:
        refresh_cache(vehicle)
    return vehicle_cache[vehicle]["open_count"]

def get_remaining_km_cached(vehicle: str) -> int | None:
    if vehicle in cache_dirty:
        refresh_cache(vehicle)
        cache_dirty.discard(vehicle)
    if vehicle not in vehicle_cache:
        refresh_cache(vehicle)
    return vehicle_cache[vehicle]["remaining_km"]

def get_last_vidange_km_noconn(conn, vehicle: str) -> int:
    cur = conn.cursor()
    cur.execute("SELECT last_vidange_km FROM vehicle_vidange WHERE vehicle=%s", (vehicle,))
    row = cur.fetchone()
    return row[0] if row else 0

# ----------------------------------------------------------------------
# Database functions (unchanged)
# ----------------------------------------------------------------------
def get_driver(user_id: int) -> dict | None:
    with db_connection() as conn:
        cur = conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
        cur.execute("SELECT name, vehicle, state, approval_status FROM drivers WHERE user_id = %s", (user_id,))
        row = cur.fetchone()
        return dict(row) if row else None

def set_driver(user_id: int, name=None, vehicle=None, state=None, approval_status=None):
    with db_connection() as conn:
        cur = conn.cursor()
        driver = get_driver(user_id)
        if driver:
            if name is not None:
                cur.execute("UPDATE drivers SET name=%s WHERE user_id=%s", (name, user_id))
            if vehicle is not None:
                cur.execute("UPDATE drivers SET vehicle=%s WHERE user_id=%s", (vehicle, user_id))
            if state is not None:
                cur.execute("UPDATE drivers SET state=%s WHERE user_id=%s", (state, user_id))
            if approval_status is not None:
                cur.execute("UPDATE drivers SET approval_status=%s WHERE user_id=%s", (approval_status, user_id))
        else:
            cur.execute("INSERT INTO drivers (user_id, name, vehicle, state, approval_status) VALUES (%s,%s,%s,%s,%s)",
                        (user_id, name or "", vehicle or "", state or "name_entry", approval_status or "pending"))
        conn.commit()

def get_all_vehicles():
    with db_connection() as conn:
        cur = conn.cursor()
        cur.execute("SELECT code FROM vehicles ORDER BY code")
        rows = cur.fetchall()
        return [r[0] for r in rows]

def add_vehicle(code: str):
    with db_connection() as conn:
        cur = conn.cursor()
        cur.execute("INSERT INTO vehicles (code) VALUES (%s) ON CONFLICT DO NOTHING", (code,))
        cur.execute("INSERT INTO vehicle_vidange (vehicle, last_vidange_km) VALUES (%s, 0) ON CONFLICT DO NOTHING", (code,))
        conn.commit()

def remove_vehicle(code: str):
    with db_connection() as conn:
        cur = conn.cursor()
        cur.execute("DELETE FROM vehicles WHERE code=%s", (code,))
        conn.commit()
    invalidate_cache(code)

def add_problem(user_id: int, driver_name: str, vehicle: str, problem_text: str, media_type: str, group_msg_id: int = 0) -> int:
    with db_connection() as conn:
        cur = conn.cursor()
        date = datetime.now(TZ).strftime("%Y-%m-%d %H:%M:%S")
        cur.execute(
            "INSERT INTO problems (user_id, driver_name, vehicle, problem_text, media_type, date, group_message_id) VALUES (%s,%s,%s,%s,%s,%s,%s) RETURNING id",
            (user_id, driver_name, vehicle, problem_text, media_type, date, group_msg_id)
        )
        problem_id = cur.fetchone()[0]
        conn.commit()
    invalidate_cache(vehicle)
    return problem_id

def update_problem_status(problem_id: int, status=None, ruglee=None, validation_requester=None, group_message_id=None):
    with db_connection() as conn:
        cur = conn.cursor()
        if status is not None:
            cur.execute("UPDATE problems SET status=%s WHERE id=%s", (status, problem_id))
        if ruglee is not None:
            cur.execute("UPDATE problems SET ruglee=%s WHERE id=%s", (ruglee, problem_id))
        if validation_requester is not None:
            cur.execute("UPDATE problems SET validation_requester=%s WHERE id=%s", (validation_requester, problem_id))
        if group_message_id is not None:
            cur.execute("UPDATE problems SET group_message_id=%s WHERE id=%s", (group_message_id, problem_id))
        conn.commit()
    problem = get_problem(problem_id)
    if problem:
        invalidate_cache(problem["vehicle"])

def set_problem_comment(problem_id: int, comment: str):
    with db_connection() as conn:
        cur = conn.cursor()
        cur.execute("UPDATE problems SET comments=%s WHERE id=%s", (comment, problem_id))
        conn.commit()

def delete_problem(problem_id: int):
    problem = get_problem(problem_id)
    with db_connection() as conn:
        cur = conn.cursor()
        cur.execute("DELETE FROM problems WHERE id=%s", (problem_id,))
        conn.commit()
    if problem:
        invalidate_cache(problem["vehicle"])

def get_problem(problem_id: int) -> dict | None:
    with db_connection() as conn:
        cur = conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
        cur.execute("SELECT * FROM problems WHERE id=%s", (problem_id,))
        row = cur.fetchone()
        return dict(row) if row else None

def get_all_problems():
    with db_connection() as conn:
        cur = conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
        cur.execute("SELECT * FROM problems ORDER BY date DESC")
        rows = cur.fetchall()
        return [dict(r) for r in rows]

def get_driver_problems(user_id: int, status_filter: str = None) -> list:
    with db_connection() as conn:
        cur = conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
        if status_filter:
            cur.execute("SELECT * FROM problems WHERE user_id=%s AND status=%s ORDER BY date DESC", (user_id, status_filter))
        else:
            cur.execute("SELECT * FROM problems WHERE user_id=%s ORDER BY date DESC", (user_id,))
        rows = cur.fetchall()
        return [dict(r) for r in rows]

def add_km_reading(vehicle: str, km: int, driver_name: str = ""):
    with db_connection() as conn:
        cur = conn.cursor()
        date = datetime.now(TZ).strftime("%Y-%m-%d %H:%M:%S")
        cur.execute("INSERT INTO km_readings (vehicle, km, date, driver_name) VALUES (%s,%s,%s,%s)", (vehicle, km, date, driver_name))
        conn.commit()
    invalidate_cache(vehicle)

def get_latest_km(vehicle: str) -> int | None:
    with db_connection() as conn:
        cur = conn.cursor()
        cur.execute("SELECT km FROM km_readings WHERE vehicle=%s ORDER BY date DESC LIMIT 1", (vehicle,))
        row = cur.fetchone()
        return row[0] if row else None

def get_last_vidange_km(vehicle: str) -> int:
    with db_connection() as conn:
        return get_last_vidange_km_noconn(conn, vehicle)

def set_last_vidange_km(vehicle: str, km: int):
    with db_connection() as conn:
        cur = conn.cursor()
        cur.execute("UPDATE vehicle_vidange SET last_vidange_km=%s WHERE vehicle=%s", (km, vehicle))
        conn.commit()
    invalidate_cache(vehicle)

def has_active_vidange(vehicle: str) -> bool:
    with db_connection() as conn:
        cur = conn.cursor()
        cur.execute("SELECT id FROM problems WHERE vehicle=%s AND media_type='نظام' AND ruglee != 'تم الإصلاح'", (vehicle,))
        row = cur.fetchone()
        return row is not None

def add_help_video(file_id: str, description: str = ""):
    with db_connection() as conn:
        cur = conn.cursor()
        cur.execute("INSERT INTO help_videos (file_id, description) VALUES (%s,%s)", (file_id, description))
        conn.commit()

def get_all_help_videos():
    with db_connection() as conn:
        cur = conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
        cur.execute("SELECT * FROM help_videos ORDER BY id DESC")
        rows = cur.fetchall()
        return [dict(r) for r in rows]

def delete_help_video(video_id: int):
    with db_connection() as conn:
        cur = conn.cursor()
        cur.execute("DELETE FROM help_videos WHERE id=%s", (video_id,))
        conn.commit()

def add_allowed_user(user_id: int, status: str = "approved"):
    with db_connection() as conn:
        cur = conn.cursor()
        cur.execute("INSERT INTO allowed_users (user_id, status) VALUES (%s,%s) ON CONFLICT (user_id) DO UPDATE SET status=%s", (user_id, status, status))
        conn.commit()

def is_allowed(user_id: int) -> bool:
    with db_connection() as conn:
        cur = conn.cursor()
        cur.execute("SELECT status FROM allowed_users WHERE user_id=%s", (user_id,))
        row = cur.fetchone()
        return row is not None and row[0] == 'approved'

def get_all_drivers():
    with db_connection() as conn:
        cur = conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
        cur.execute("SELECT user_id, name, vehicle FROM drivers WHERE approval_status='approved'")
        rows = cur.fetchall()
        return [dict(r) for r in rows]

def remove_driver(user_id: int):
    with db_connection() as conn:
        cur = conn.cursor()
        cur.execute("DELETE FROM drivers WHERE user_id=%s", (user_id,))
        cur.execute("DELETE FROM allowed_users WHERE user_id=%s", (user_id,))
        conn.commit()

# ----------------------------------------------------------------------
# Excel generation
# ----------------------------------------------------------------------
def generate_problems_excel() -> BytesIO:
    problems = get_all_problems()
    wb = Workbook()
    ws = wb.active
    ws.title = "المشاكل"
    headers = ["التاريخ", "السائق", "المركبة", "المشكلة", "نوع الوسائط", "الحالة", "تم الإصلاح", "تعليقات"]
    for col, h in enumerate(headers, 1):
        ws.cell(row=1, column=col, value=h).font = Font(bold=True)
    red_fill = PatternFill(start_color="FFCCCC", end_color="FFCCCC", fill_type="solid")
    orange_fill = PatternFill(start_color="FFE5CC", end_color="FFE5CC", fill_type="solid")
    green_fill = PatternFill(start_color="CCFFCC", end_color="CCFFCC", fill_type="solid")
    for row_idx, p in enumerate(problems, 2):
        if p["ruglee"] == "تم الإصلاح":
            row_fill = green_fill
        elif p["status"] == "قيد الانتظار":
            row_fill = red_fill
        elif p["status"] == "قيد التصليح":
            row_fill = orange_fill
        else:
            row_fill = None
        vals = [p["date"], p["driver_name"], p["vehicle"], p["problem_text"], p["media_type"] or "—",
                p["status"], p["ruglee"], p["comments"] or ""]
        for ci, val in enumerate(vals, 1):
            cell = ws.cell(row=row_idx, column=ci, value=val)
            if row_fill:
                cell.fill = row_fill
    for col in ws.columns:
        max_len = max((len(str(c.value)) for c in col if c.value), default=0)
        ws.column_dimensions[get_column_letter(col[0].column)].width = min(max_len+2, 50)
    out = BytesIO()
    wb.save(out)
    out.seek(0)
    return out

def generate_vidange_excel(vehicle_code: str = None) -> BytesIO:
    vehicles = [vehicle_code] if vehicle_code else get_all_vehicles()
    wb = Workbook()
    wb.remove(wb.active)
    for v in vehicles:
        ws = wb.create_sheet(title=v)
        ws.append(["التاريخ", "السائق", "العداد الحالي (KM)", "آخر فيدانج (KM)", "المسافة المتبقية للفيدانج"])
        with db_connection() as conn:
            cur = conn.cursor()
            cur.execute("SELECT date, km, driver_name FROM km_readings WHERE vehicle=%s ORDER BY date DESC", (v,))
            readings = cur.fetchall()
        last_km = get_last_vidange_km(v)
        for date_str, km, dname in readings:
            remaining = (last_km + 10000) - km if last_km > 0 else "—"
            ws.append([date_str, dname or "", km, last_km, remaining])
        for col in ws.columns:
            max_len = max((len(str(c.value)) for c in col if c.value), default=0)
            ws.column_dimensions[get_column_letter(col[0].column)].width = min(max_len+2, 50)
    out = BytesIO()
    wb.save(out)
    out.seek(0)
    return out

# ----------------------------------------------------------------------
# Keyboards & Helpers
# ----------------------------------------------------------------------
MAIN_KEYBOARD = ReplyKeyboardMarkup([
    [KeyboardButton("📝 تقديم شكوى"), KeyboardButton("✅ طلب التحقق من الإصلاح")],
    [KeyboardButton("🚗 إدخال عدد الكيلومترات"), KeyboardButton("⚙️ الإعدادات")]
], resize_keyboard=True)

ADMIN_STICKY_KEYBOARD = ReplyKeyboardMarkup([
    [KeyboardButton("/admin")]
], resize_keyboard=True, one_time_keyboard=False)

def vehicle_inline_keyboard(vehicles: list, prefix="selv_") -> InlineKeyboardMarkup:
    buttons = [InlineKeyboardButton(v, callback_data=f"{prefix}{v}") for v in vehicles]
    return InlineKeyboardMarkup([buttons[i:i+4] for i in range(0, len(buttons), 4)])

def status_emoji_cached(vehicle: str) -> str:
    s = get_vehicle_status_cached(vehicle)
    if s == 'bad': return "🔴"
    if s == 'en_cours': return "🟠"
    return "🟢"

def dashboard_button_text(vehicle: str) -> str:
    emoji = status_emoji_cached(vehicle)
    cnt = count_open_problems_cached(vehicle)
    rem = get_remaining_km_cached(vehicle)
    line1 = f"{emoji} {vehicle} ({cnt})"
    # No "كم", just the number or "--"
    line2 = f"   {rem}" if rem is not None else "   —"
    return f"{line1}\n{line2}"

def admin_main_keyboard() -> InlineKeyboardMarkup:
    return InlineKeyboardMarkup([
        [InlineKeyboardButton("🚘 تعديل المركبات", callback_data="admin_vehicles")],
        [InlineKeyboardButton("👤 إدارة السائقين", callback_data="admin_drivers")],
        [InlineKeyboardButton("🛢️ إدارة الفيدانج", callback_data="admin_vidange_menu")],
        [InlineKeyboardButton("📊 لوحة القيادة", callback_data="admin_dash")],
        [InlineKeyboardButton("📋 تصدير Excel", callback_data="admin_export_menu")],
    ])

def admin_vehicles_keyboard() -> InlineKeyboardMarkup:
    return InlineKeyboardMarkup([
        [InlineKeyboardButton("➕ إضافة مركبة", callback_data="admin_addveh"),
         InlineKeyboardButton("➖ حذف مركبة", callback_data="admin_remveh")],
        [InlineKeyboardButton("↩️ رجوع", callback_data="admin_main")]
    ])

def admin_drivers_keyboard() -> InlineKeyboardMarkup:
    return InlineKeyboardMarkup([
        [InlineKeyboardButton("✅ قبول سائقين", callback_data="admin_approve_list")],
        [InlineKeyboardButton("❌ حذف سائق", callback_data="admin_remove_driver_list")],
        [InlineKeyboardButton("📋 قائمة السائقين", callback_data="admin_drivers_list")],
        [InlineKeyboardButton("↩️ رجوع", callback_data="admin_main")]
    ])

def admin_vidange_menu_keyboard() -> InlineKeyboardMarkup:
    return InlineKeyboardMarkup([
        [InlineKeyboardButton("⚙️ تعيين فيدانج", callback_data="admin_setvid")],
        [InlineKeyboardButton("🚨 فيدانج عاجل", callback_data="admin_urgentvid")],
        [InlineKeyboardButton("↩️ رجوع", callback_data="admin_main")]
    ])

def admin_export_menu_keyboard() -> InlineKeyboardMarkup:
    return InlineKeyboardMarkup([
        [InlineKeyboardButton("📋 تصدير المشاكل", callback_data="admin_export")],
        [InlineKeyboardButton("🛢️ تصدير الفيدانج", callback_data="admin_vid")],
        [InlineKeyboardButton("↩️ رجوع", callback_data="admin_main")]
    ])

TOPIC_ACTIONS = {
    TOPIC_GENERAL: [
        [InlineKeyboardButton("📊 لوحة القيادة", callback_data="admin_dash")],
        [InlineKeyboardButton("📋 تصدير المشاكل", callback_data="admin_export")],
        [InlineKeyboardButton("🛢️ تصدير الفيدانج", callback_data="admin_vid")],
    ],
    TOPIC_RECLAMATIONS: [
        [InlineKeyboardButton("📊 لوحة القيادة", callback_data="admin_dash")],
        [InlineKeyboardButton("📋 تصدير المشاكل", callback_data="admin_export")],
    ],
    TOPIC_VALIDATION: [
        [InlineKeyboardButton("📊 لوحة القيادة", callback_data="admin_dash")],
        [InlineKeyboardButton("📋 تصدير المشاكل", callback_data="admin_export")],
    ],
    TOPIC_VIDANGE: [
        [InlineKeyboardButton("🛢️ تصدير الفيدانج", callback_data="admin_vid")],
        [InlineKeyboardButton("🚨 طلب فيدانج عاجل", callback_data="admin_urgentvid")],
        [InlineKeyboardButton("📋 تصدير المشاكل", callback_data="admin_export")],
    ],
    TOPIC_VEHICLE_MGMT: [
        [InlineKeyboardButton("➕ إضافة مركبة", callback_data="admin_addveh")],
        [InlineKeyboardButton("➖ حذف مركبة", callback_data="admin_remveh")],
        [InlineKeyboardButton("🚘 قائمة المركبات", callback_data="admin_listveh")],
    ],
    TOPIC_HISTORY: [
        [InlineKeyboardButton("📊 لوحة القيادة", callback_data="admin_dash")],
    ],
}

def get_topic_keyboard(thread_id: int) -> InlineKeyboardMarkup | None:
    actions = TOPIC_ACTIONS.get(thread_id)
    if actions:
        return InlineKeyboardMarkup(actions)
    return None

def status_icon_and_text(problem: dict) -> str:
    if problem["ruglee"] == "تم الإصلاح":
        return "🟢 مُصلح"
    if problem["status"] == "قيد الانتظار":
        return "🔴 قيد الانتظار"
    if problem["status"] == "قيد التصليح":
        return "🟠 قيد التصليح"
    return "⚪ حالة غير معروفة"

def build_problem_keyboard(problem_id: int) -> InlineKeyboardMarkup:
    if problem_id == 0:
        return InlineKeyboardMarkup([])
    problem = get_problem(problem_id)
    if not problem:
        return InlineKeyboardMarkup([])
    status = problem["status"]
    ruglee = problem["ruglee"]
    row1 = []
    if status == "قيد الانتظار":
        row1.append(InlineKeyboardButton("🔧 قيد التصليح", callback_data=f"val_{problem_id}"))
    elif status == "قيد التصليح":
        row1.append(InlineKeyboardButton("⏳ إعادة للانتظار", callback_data=f"val_{problem_id}"))
        if problem["media_type"] and not problem["comments"]:
            row1.append(InlineKeyboardButton("✅ تم الإصلاح (تعليق مطلوب)", callback_data=f"fix_comment_{problem_id}"))
        else:
            row1.append(InlineKeyboardButton("✅ تم الإصلاح", callback_data=f"rug_{problem_id}"))
    rows = []
    if row1:
        rows.append(row1)
    rows.append([InlineKeyboardButton("💬 تعليق", callback_data=f"com_{problem_id}")])
    if ruglee != "تم الإصلاح":
        rows.append([InlineKeyboardButton("🗑️ حذف", callback_data=f"del_{problem_id}")])
    return InlineKeyboardMarkup(rows)

def update_status_line(problem: dict, new_status: str = None, new_ruglee: str = None) -> str:
    dname = problem["driver_name"]
    veh = problem["vehicle"]
    prob_text = problem["problem_text"]
    status_line = f"الحالة: {status_icon_and_text({'status': new_status or problem['status'], 'ruglee': new_ruglee or problem['ruglee']})}"
    return f"السائق: {dname}\nالمركبة: {veh}\nالمشكلة: {prob_text}\n{status_line}"

async def _update_problem_message(problem: dict, new_status: str = None, new_ruglee: str = None):
    if not problem["group_message_id"]:
        return
    new_text = update_status_line(problem, new_status=new_status, new_ruglee=new_ruglee)
    try:
        if problem["media_type"]:
            await app.bot.edit_message_caption(
                chat_id=ADMIN_GROUP_ID,
                message_id=problem["group_message_id"],
                caption=new_text
            )
        else:
            await app.bot.edit_message_text(
                chat_id=ADMIN_GROUP_ID,
                message_id=problem["group_message_id"],
                text=new_text
            )
    except Exception as e:
        logging.warning(f"Could not update problem message: {e}")

# ----------------------------------------------------------------------
# Cancel handler
# ----------------------------------------------------------------------
async def cancel_input(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    context.user_data.clear()
    if update.effective_chat.type == "private":
        driver = get_driver(query.from_user.id)
        if driver and driver.get("approval_status") == "approved":
            set_driver(query.from_user.id, state="idle")
            await query.edit_message_text("تم الإلغاء.")
            await context.bot.send_message(chat_id=query.from_user.id,
                                           text="يمكنك الآن استخدام الأزرار أدناه:",
                                           reply_markup=MAIN_KEYBOARD)
            return
    await query.edit_message_text("تم الإلغاء.")

# ----------------------------------------------------------------------
# Core handlers
# ----------------------------------------------------------------------
async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    if user_id in ADMIN_IDS:
        add_allowed_user(user_id)
        set_driver(user_id, approval_status="approved")
    driver = get_driver(user_id)
    if driver and driver["name"] and driver["vehicle"] and driver.get("approval_status") == "approved":
        if driver["state"] != "idle":
            set_driver(user_id, state="idle")
        await update.message.reply_text(
            f"أهلاً بعودتك، {driver['name']}!\nمركبتك: {driver['vehicle']}",
            reply_markup=MAIN_KEYBOARD
        )
    elif driver and driver.get("approval_status") == "pending":
        await update.message.reply_text("شكراً لتسجيلك. طلب صلاحيتك قيد المراجعة. انتظر قبول المشرف.")
    else:
        set_driver(user_id, state="name_entry", approval_status="pending")
        markup = InlineKeyboardMarkup([[InlineKeyboardButton("إلغاء", callback_data="cancel_input")]])
        await update.message.reply_text("مرحباً! الرجاء إدخال اسمك الكامل:", reply_markup=markup)

async def handle_text(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.effective_chat.type != "private":
        return
    user_id = update.effective_user.id
    text = update.message.text.strip()
    driver = get_driver(user_id)
    state = driver["state"] if driver else "name_entry"

    # Block non-allowed users
    if user_id not in ADMIN_IDS and not is_allowed(user_id):
        if state == "name_entry":
            set_driver(user_id, name=text, state="awaiting_approval", approval_status="pending")
            username = update.effective_user.username
            mention = f"@{username}" if username else f"[{text}](tg://user?id={user_id})"
            msg = await context.bot.send_message(
                chat_id=ADMIN_GROUP_ID,
                message_thread_id=TOPIC_VALIDATION,
                text=f"📌 طلب صلاحية جديدة:\nالاسم: {text}\nالمستخدم: {mention}\nالحالة: ⏳ قيد الانتظار",
                reply_markup=InlineKeyboardMarkup([
                    [InlineKeyboardButton("✅ قبول", callback_data=f"approve_{user_id}"),
                     InlineKeyboardButton("❌ رفض", callback_data=f"reject_{user_id}")]
                ])
            )
            context.bot_data.setdefault("approval_msg", {})[user_id] = msg.message_id
            await update.message.reply_text("تم إرسال طلبك للمراجعة. ستصلك رسالة عند القبول.")
            return
        else:
            await update.message.reply_text("غير مصرح لك باستخدام البوت حالياً.")
            return

    # Comment session
    if context.user_data.get("awaiting_comment"):
        problem_id = context.user_data.pop("awaiting_comment")
        set_problem_comment(problem_id, text)
        await update.message.reply_text("✅ تم حفظ التعليق بنجاح.", reply_markup=MAIN_KEYBOARD)
        return

    # KM after vidange repair
    if context.user_data.get("await_km"):
        vehicle = context.user_data["await_km_vehicle"]
        if text.isdigit():
            km = int(text)
            last_km = get_latest_km(vehicle)
            if last_km is not None and km <= last_km:
                await update.message.reply_text(f"⚠️ الكيلومتر يجب أن يكون أكبر من آخر قراءة ({last_km} كم). أعد إدخال القيمة الصحيحة.",
                                                reply_markup=InlineKeyboardMarkup([[InlineKeyboardButton("إلغاء", callback_data="cancel_input")]]))
                return
            driver_name = driver["name"] if driver else "Unknown"
            msg = await context.bot.send_message(
                chat_id=ADMIN_GROUP_ID,
                message_thread_id=TOPIC_VIDANGE,
                text=f"⚙️ تأكيد تحديث الفيدانج:\nالمركبة: {vehicle}\nالقيمة المدخلة: {km} كم\nالسائق: {driver_name}\nالحالة: ⏳ في انتظار التأكيد",
                reply_markup=InlineKeyboardMarkup([
                    [InlineKeyboardButton("✅ تأكيد", callback_data=f"vidconfirm_{user_id}_{km}"),
                     InlineKeyboardButton("✏️ تعديل", callback_data=f"vidmodify_{user_id}")]
                ])
            )
            context.bot_data.setdefault("pending_vidange", {})[msg.message_id] = {
                "user_id": user_id,
                "vehicle": vehicle,
                "km": km
            }
            await update.message.reply_text("تم إرسال القيمة للمراجعة من قبل المشرفين.", reply_markup=MAIN_KEYBOARD)
            context.user_data.pop("await_km_vehicle", None)
            context.user_data.clear()
            return
        else:
            await update.message.reply_text("الرجاء إرسال رقم صحيح.",
                                            reply_markup=InlineKeyboardMarkup([[InlineKeyboardButton("إلغاء", callback_data="cancel_input")]]))
            return

    # State machine
    if state == "name_entry":
        allowed = True
        for c in text:
            if not (c.isalnum() or c.isspace() or ('\u0600' <= c <= '\u06ff') or c in '-ء'):
                allowed = False
                break
        if not text or not allowed:
            await update.message.reply_text("⚠️ الاسم يجب أن يحتوي على أحرف وأرقام ومسافات فقط. أعد الإدخال.",
                                            reply_markup=InlineKeyboardMarkup([[InlineKeyboardButton("إلغاء", callback_data="cancel_input")]]))
            return
        if driver and driver.get("vehicle"):
            set_driver(user_id, name=text, state="idle")
            await update.message.reply_text(f"تم تغيير الاسم إلى {text}.", reply_markup=MAIN_KEYBOARD)
            return
        set_driver(user_id, name=text, state="vehicle_selection")
        vehicles = get_all_vehicles()
        await update.message.reply_text("تم حفظ الاسم. اختر مركبتك:", reply_markup=vehicle_inline_keyboard(vehicles, "selv_"))
        return

    if state == "vehicle_selection":
        vehicles = get_all_vehicles()
        await update.message.reply_text("الرجاء اختيار المركبة من القائمة:", reply_markup=vehicle_inline_keyboard(vehicles, "selv_"))
        return

    # Main keyboard
    if text == "📝 تقديم شكوى":
        markup = InlineKeyboardMarkup([[InlineKeyboardButton("إلغاء", callback_data="cancel_input")]])
        await update.message.reply_text("أرسل وصف المشكلة (نص، صورة، فيديو، أو صوت).", reply_markup=markup)
        context.user_data["expecting_reclamation"] = True
        return
    if text == "✅ طلب التحقق من الإصلاح":
        problems = get_driver_problems(user_id, status_filter="قيد التصليح")
        if not problems:
            await update.message.reply_text("لا توجد مشاكل بحاجة للتحقق من إصلاحها.", reply_markup=MAIN_KEYBOARD)
            return
        pending = [p for p in problems if p["validation_requester"] == 0]
        if not pending:
            await update.message.reply_text("جميع المشاكل قيد التحقق بالفعل.", reply_markup=MAIN_KEYBOARD)
            return
        buttons = [InlineKeyboardButton(f"مشكلة #{p['id']} - {p['problem_text'][:30]}...", callback_data=f"valreq_{p['id']}") for p in pending]
        await update.message.reply_text("اختر المشكلة التي تم إصلاحها:", reply_markup=InlineKeyboardMarkup([buttons[i:i+2] for i in range(0, len(buttons), 2)]))
        return
    if text == "🚗 إدخال عدد الكيلومترات":
        if not driver or not driver["vehicle"]:
            await update.message.reply_text("يجب إكمال الملف أولاً.")
            return
        markup = InlineKeyboardMarkup([[InlineKeyboardButton("إلغاء", callback_data="cancel_input")]])
        await update.message.reply_text(f"أرسل عدد الكيلومترات الحالي للمركبة {driver['vehicle']}:", reply_markup=markup)
        context.user_data["expecting_km"] = True
        return
    if text == "⚙️ الإعدادات":
        markup = InlineKeyboardMarkup([
            [InlineKeyboardButton("🚘 تغيير المركبة", callback_data="settings_change_veh")],
            [InlineKeyboardButton("📜 سجل شكاويي", callback_data="settings_history")],
            [InlineKeyboardButton("✏️ تعديل الاسم", callback_data="settings_change_name")],
            [InlineKeyboardButton("❓ مساعدة", callback_data="settings_help")]
        ])
        await update.message.reply_text("اختر الإعداد المطلوب:", reply_markup=markup)
        return

    # Reclamation / km input
    if context.user_data.get("expecting_reclamation"):
        if not text:
            await update.message.reply_text("الرجاء كتابة وصف للمشكلة. لا يمكن إرسال شكوى فارغة.",
                                            reply_markup=InlineKeyboardMarkup([[InlineKeyboardButton("إلغاء", callback_data="cancel_input")]]))
            return
        context.user_data.pop("expecting_reclamation")
        if not driver or not driver["name"] or not driver["vehicle"]:
            await update.message.reply_text("ملفك غير مكتمل.")
            return
        status_line = "🔴 الحالة: قيد الانتظار"
        report = f"السائق: {driver['name']}\nالمركبة: {driver['vehicle']}\nالمشكلة: {text}\n{status_line}"
        msg = await context.bot.send_message(chat_id=ADMIN_GROUP_ID, message_thread_id=TOPIC_RECLAMATIONS, text=report,
                                             reply_markup=build_problem_keyboard(0))
        problem_id = add_problem(user_id, driver["name"], driver["vehicle"], text, "", group_msg_id=msg.message_id)
        await msg.edit_reply_markup(reply_markup=build_problem_keyboard(problem_id))
        await update.message.reply_text("تم إرسال الشكوى.", reply_markup=MAIN_KEYBOARD)
        return
    if context.user_data.get("expecting_km"):
        if not text.isdigit():
            await update.message.reply_text("يجب إرسال رقم.",
                                            reply_markup=InlineKeyboardMarkup([[InlineKeyboardButton("إلغاء", callback_data="cancel_input")]]))
            return
        km = int(text)
        vehicle = driver["vehicle"] if driver else None
        if not vehicle:
            await update.message.reply_text("ملف غير مكتمل.")
            return
        last_km = get_latest_km(vehicle)
        if last_km is not None and km <= last_km:
            await update.message.reply_text(f"⚠️ الكيلومتر يجب أن يكون أكبر من آخر قراءة ({last_km} كم). أعد إدخال القيمة الصحيحة.",
                                            reply_markup=InlineKeyboardMarkup([[InlineKeyboardButton("إلغاء", callback_data="cancel_input")]]))
            return
        context.user_data.pop("expecting_km")
        add_km_reading(vehicle, km, driver_name=driver["name"])
        last_vid = get_last_vidange_km(vehicle)
        # Create vidange problem at +9000
        if last_vid > 0 and km >= last_vid + 9000 and not has_active_vidange(vehicle):
            vidange_problem_id = add_problem(user_id, f"{driver['name']} (نظام)", vehicle, f"Vidange {vehicle}", "نظام")
            await context.bot.send_message(
                chat_id=ADMIN_GROUP_ID, message_thread_id=TOPIC_VIDANGE,
                text=f"⚠️ تنبيه فيدانج: المركبة {vehicle}\nالعداد الحالي: {km} كم\nآخر فيدانج: {last_vid} كم\n🔴 الحالة: قيد الانتظار",
                reply_markup=build_problem_keyboard(vidange_problem_id)
            )
        # Urgent notification at +9500
        if last_vid > 0 and km >= last_vid + 9500:
            await context.bot.send_message(
                chat_id=ADMIN_GROUP_ID, message_thread_id=TOPIC_VIDANGE,
                text=f"🚨 تنبيه متقدم: المركبة {vehicle} تجاوزت 9500 كم بعد آخر فيدانج! العداد الحالي: {km} كم. يُرجى الإسراع في الصيانة."
            )
        await update.message.reply_text(f"تم تسجيل العداد: {km} كم.", reply_markup=MAIN_KEYBOARD)
        return

    await update.message.reply_text("استخدم الأزرار أدناه.", reply_markup=MAIN_KEYBOARD)

async def handle_media(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.effective_chat.type != "private":
        return
    user_id = update.effective_user.id
    driver = get_driver(user_id)
    if not driver or not driver["name"] or not driver["vehicle"] or driver.get("approval_status") != "approved":
        await update.message.reply_text("ملفك غير مكتمل.")
        return
    caption = update.message.caption or "(مرفق وسائط)"
    header = f"السائق: {driver['name']}\nالمركبة: {driver['vehicle']}\nالمشكلة: {caption}\n🔴 الحالة: قيد الانتظار"
    if update.message.photo:
        file_id = update.message.photo[-1].file_id
        media_type = "صورة"
        msg = await context.bot.send_photo(chat_id=ADMIN_GROUP_ID, message_thread_id=TOPIC_RECLAMATIONS, photo=file_id,
                                           caption=header, reply_markup=build_problem_keyboard(0))
        problem_id = add_problem(user_id, driver["name"], driver["vehicle"], caption, media_type, group_msg_id=msg.message_id)
        await msg.edit_reply_markup(reply_markup=build_problem_keyboard(problem_id))
    elif update.message.video:
        file_id = update.message.video.file_id
        media_type = "فيديو"
        msg = await context.bot.send_video(chat_id=ADMIN_GROUP_ID, message_thread_id=TOPIC_RECLAMATIONS, video=file_id,
                                           caption=header, reply_markup=build_problem_keyboard(0))
        problem_id = add_problem(user_id, driver["name"], driver["vehicle"], caption, media_type, group_msg_id=msg.message_id)
        await msg.edit_reply_markup(reply_markup=build_problem_keyboard(problem_id))
    elif update.message.voice:
        file_id = update.message.voice.file_id
        media_type = "صوت"
        msg = await context.bot.send_voice(chat_id=ADMIN_GROUP_ID, message_thread_id=TOPIC_RECLAMATIONS, voice=file_id,
                                           caption=header, reply_markup=build_problem_keyboard(0))
        problem_id = add_problem(user_id, driver["name"], driver["vehicle"], caption, media_type, group_msg_id=msg.message_id)
        await msg.edit_reply_markup(reply_markup=build_problem_keyboard(problem_id))
    await update.message.reply_text("تم إرسال الشكوى.", reply_markup=MAIN_KEYBOARD)

# ----------------------------------------------------------------------
# Callback handlers (fixed)
# ----------------------------------------------------------------------
async def vehicle_selection_callback(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    vehicle = query.data.split("_", 1)[1]
    user_id = query.from_user.id
    context.user_data["confirm_vehicle"] = vehicle
    await query.edit_message_text(f"هل تريد تعيين المركبة {vehicle}؟", reply_markup=InlineKeyboardMarkup([
        [InlineKeyboardButton("نعم", callback_data=f"confirmveh_{vehicle}"),
         InlineKeyboardButton("إلغاء", callback_data="cancel_veh")]
    ]))

async def confirm_vehicle_callback(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    vehicle = query.data.split("_", 1)[1]
    user_id = query.from_user.id
    set_driver(user_id, vehicle=vehicle, state="idle")
    await query.edit_message_text(f"تم تعيين المركبة إلى {vehicle}.")
    await context.bot.send_message(chat_id=user_id, text="يمكنك الآن استخدام الأزرار أدناه:", reply_markup=MAIN_KEYBOARD)

async def cancel_vehicle_selection_callback(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    vehicles = get_all_vehicles()
    await query.edit_message_text("اختر مركبتك:", reply_markup=vehicle_inline_keyboard(vehicles, "selv_"))

async def valide_callback(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    problem_id = int(query.data.split("_")[1])
    problem = get_problem(problem_id)
    if not problem: return await query.answer("المشكلة غير موجودة.")
    new_status = "قيد التصليح" if problem["status"] == "قيد الانتظار" else "قيد الانتظار"
    update_problem_status(problem_id, status=new_status)
    await _update_problem_message(problem, new_status=new_status)
    await query.edit_message_reply_markup(reply_markup=build_problem_keyboard(problem_id))

async def ruglee_callback(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    problem_id = int(query.data.split("_")[1])
    problem = get_problem(problem_id)
    if not problem: return await query.answer("غير موجود.")
    if problem["media_type"] and not problem["comments"]:
        await query.answer("يجب إضافة تعليق أولاً قبل تأكيد الإصلاح.", show_alert=True)
        await context.bot.send_message(chat_id=query.from_user.id, text="يجب إضافة تعليق للمشكلة قبل وضعها كمُصلحة. أرسل التعليق هنا.")
        context.user_data["awaiting_comment"] = problem_id
        return
    new_ruglee = "تم الإصلاح" if problem["ruglee"] == "غير مُصلح" else "غير مُصلح"
    update_problem_status(problem_id, ruglee=new_ruglee)
    await _update_problem_message(problem, ruglee=new_ruglee)
    val_msg_id = context.bot_data.get("validation_msgs", {}).pop(problem_id, None)
    if val_msg_id:
        try:
            await context.bot.edit_message_text(
                chat_id=ADMIN_GROUP_ID, message_id=val_msg_id,
                text=f"📌 طلب تحقق من الإصلاح:\nالمشكلة #{problem_id} - {problem['problem_text']}\nالمركبة: {problem['vehicle']}\nالسائق: ...\nالحالة: ✅ تم الإصلاح"
            )
        except: pass
    await query.edit_message_reply_markup(reply_markup=build_problem_keyboard(problem_id))
    if problem["media_type"] == "نظام" and new_ruglee == "تم الإصلاح":
        req_id = problem.get("validation_requester") or problem.get("user_id")
        if req_id:
            try:
                await context.bot.send_message(chat_id=req_id, text=f"تم تأكيد إصلاح الفيدانج للمركبة {problem['vehicle']}. الرجاء إدخال الكيلومترات الحالية:")
                context.bot_data.setdefault("km_await", {})[req_id] = problem["vehicle"]
            except: pass

async def fix_comment_callback(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer("يجب إضافة تعليق أولاً.", show_alert=True)
    problem_id = int(query.data.split("_")[2])
    await context.bot.send_message(chat_id=query.from_user.id, text=f"يجب إضافة تعليق على المشكلة #{problem_id} أولاً.")
    context.user_data["awaiting_comment"] = problem_id

async def comment_callback(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    problem_id = int(query.data.split("_")[1])
    context.user_data["awaiting_comment"] = problem_id
    await context.bot.send_message(chat_id=query.from_user.id, text="📝 أرسل تعليقك على المشكلة:",
                                   reply_markup=InlineKeyboardMarkup([[InlineKeyboardButton("إلغاء", callback_data="cancel_input")]]))
    await query.answer("أرسل التعليق في المحادثة الخاصة.", show_alert=True)

async def delete_callback(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    if query.from_user.id not in ADMIN_IDS:
        await query.answer("⛔ غير مصرح.", show_alert=True); return
    problem_id = int(query.data.split("_")[1])
    await query.edit_message_reply_markup(reply_markup=InlineKeyboardMarkup([
        [InlineKeyboardButton("نعم", callback_data=f"confirmdel_{problem_id}"),
         InlineKeyboardButton("إلغاء", callback_data=f"cancel_{problem_id}")]
    ]))

async def confirm_delete_callback(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    if query.from_user.id not in ADMIN_IDS: return
    problem_id = int(query.data.split("_")[1])
    delete_problem(problem_id)
    await query.edit_message_text("🗑️ تم حذف المشكلة.")

async def cancel_delete_callback(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    problem_id = int(query.data.split("_")[1])
    problem = get_problem(problem_id)
    if not problem:
        await query.edit_message_reply_markup(reply_markup=None)
        return
    await query.edit_message_reply_markup(reply_markup=build_problem_keyboard(problem_id))

async def validation_request_callback(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    problem_id = int(query.data.split("_")[1])
    problem = get_problem(problem_id)
    if not problem: return await query.answer("غير موجود.")
    if problem["validation_requester"] != 0:
        await query.answer("تم إرسال طلب تحقق سابقاً.", show_alert=True)
        return
    update_problem_status(problem_id, validation_requester=query.from_user.id)
    driver = get_driver(query.from_user.id)
    driver_name = driver["name"] if driver else "Unknown"
    msg_text = f"📌 طلب تحقق من الإصلاح:\nالمشكلة #{problem_id} - {problem['problem_text']}\nالمركبة: {problem['vehicle']}\nالسائق: {driver_name}\nالحالة: 📌 في انتظار التحقق"
    sent_msg = await context.bot.send_message(chat_id=ADMIN_GROUP_ID, message_thread_id=TOPIC_VALIDATION, text=msg_text,
                                              reply_markup=InlineKeyboardMarkup([[InlineKeyboardButton("✅ تأكيد الإصلاح", callback_data=f"valrug_{problem_id}")]]))
    context.bot_data.setdefault("validation_msgs", {})[problem_id] = sent_msg.message_id
    await query.edit_message_text("تم إرسال طلب التحقق.")

async def valrug_callback(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    problem_id = int(query.data.split("_")[1])
    problem = get_problem(problem_id)
    if not problem: return await query.answer("غير موجود.")
    if problem["media_type"] and not problem["comments"]:
        await query.answer("يجب إضافة تعليق أولاً.", show_alert=True)
        await context.bot.send_message(chat_id=query.from_user.id, text="يجب إضافة تعليق للمشكلة قبل وضعها كمُصلحة. أرسل التعليق هنا.")
        context.user_data["awaiting_comment"] = problem_id
        return
    new_ruglee = "تم الإصلاح" if problem["ruglee"] == "غير مُصلح" else "غير مُصلح"
    update_problem_status(problem_id, ruglee=new_ruglee)
    await _update_problem_message(problem, ruglee=new_ruglee)
    post_markup = InlineKeyboardMarkup([
        [InlineKeyboardButton("🔒 إغلاق الطلب", callback_data=f"close_val_{problem_id}"),
         InlineKeyboardButton("🔄 إعادة فتح", callback_data=f"reopen_val_{problem_id}")]
    ])
    try:
        await query.edit_message_reply_markup(reply_markup=post_markup)
    except Exception as e:
        logging.warning(f"Failed to update val msg markup: {e}")
    context.bot_data.get("validation_msgs", {}).pop(problem_id, None)
    if problem["media_type"] == "نظام" and new_ruglee == "تم الإصلاح":
        req_id = problem.get("validation_requester") or problem.get("user_id")
        if req_id:
            try:
                await context.bot.send_message(chat_id=req_id, text=f"تم تأكيد إصلاح الفيدانج للمركبة {problem['vehicle']}. الرجاء إدخال الكيلومترات الحالية:")
                context.bot_data.setdefault("km_await", {})[req_id] = problem["vehicle"]
            except: pass

async def close_val_callback(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    problem_id = int(query.data.split("_")[2])
    problem = get_problem(problem_id)
    if not problem: return await query.edit_message_text("المشكلة غير موجودة.")
    try:
        await query.message.delete()
    except Exception as e:
        logging.warning(f"Could not delete val msg: {e}")
    driver_id = problem.get("validation_requester") or problem.get("user_id")
    if driver_id:
        try:
            await context.bot.send_message(chat_id=driver_id, text=f"تم إغلاق طلب التحقق للمشكلة #{problem_id}.")
        except: pass

async def reopen_val_callback(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    problem_id = int(query.data.split("_")[2])
    problem = get_problem(problem_id)
    if not problem: return await query.edit_message_text("المشكلة غير موجودة.")
    update_problem_status(problem_id, ruglee="غير مُصلح")
    driver_id = problem.get("validation_requester") or problem.get("user_id")
    if driver_id:
        try:
            await context.bot.send_message(chat_id=driver_id, text=f"تم إعادة فتح طلب التحقق للمشكلة #{problem_id}. الرجاء مراجعة الإصلاح.")
        except: pass
    orig_markup = InlineKeyboardMarkup([[InlineKeyboardButton("✅ تأكيد الإصلاح", callback_data=f"valrug_{problem_id}")]])
    try:
        await query.edit_message_reply_markup(reply_markup=orig_markup)
    except: pass

# Vidange confirm / modify
async def vidange_confirm_callback(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    if query.from_user.id not in ADMIN_IDS:
        await query.answer("⛔ غير مصرح.", show_alert=True); return
    data = query.data.split("_")
    user_id = int(data[1])
    km = int(data[2])
    pending = context.bot_data.get("pending_vidange", {})
    info = pending.pop(query.message.message_id, None)
    if not info:
        await query.answer("انتهت صلاحية الطلب.", show_alert=True)
        return
    vehicle = info["vehicle"]
    add_km_reading(vehicle, km, driver_name=info.get("driver_name", ""))
    set_last_vidange_km(vehicle, km)
    await query.edit_message_text(f"✅ تم تأكيد الفيدانج للمركبة {vehicle} بقيمة {km} كم.")
    try:
        await context.bot.send_message(chat_id=user_id, text=f"✅ تم اعتماد تحديث الفيدانج للمركبة {vehicle}: {km} كم.")
    except: pass

async def vidange_modify_callback(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    if query.from_user.id not in ADMIN_IDS:
        await query.answer("⛔ غير مصرح.", show_alert=True); return
    data = query.data.split("_")
    user_id = int(data[1])
    pending = context.bot_data.get("pending_vidange", {})
    info = pending.get(query.message.message_id)
    if not info:
        await query.answer("انتهت صلاحية الطلب.", show_alert=True)
        return
    context.user_data["vidange_modify"] = {
        "user_id": user_id,
        "vehicle": info["vehicle"],
        "message_id": query.message.message_id
    }
    markup = InlineKeyboardMarkup([
        [InlineKeyboardButton("إلغاء", callback_data="cancel_input"),
         InlineKeyboardButton("↩️ رجوع", callback_data="admin_main")]
    ])
    await query.edit_message_text("✏️ أرسل القيمة الصحيحة للكيلومتر بعد الفيدانج:", reply_markup=markup)

# Approval / Rejection callbacks
async def approve_callback(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    if query.from_user.id not in ADMIN_IDS: return
    user_id = int(query.data.split("_")[1])
    add_allowed_user(user_id)
    set_driver(user_id, approval_status="approved", state="vehicle_selection")
    try:
        await query.edit_message_text(f"✅ تم قبول المستخدم {user_id}")
    except: pass
    try:
        await context.bot.send_message(chat_id=user_id, text="تم قبولك. يمكنك الآن اختيار مركبتك:")
        vehicles = get_all_vehicles()
        await context.bot.send_message(chat_id=user_id, text="اختر مركبتك:", reply_markup=vehicle_inline_keyboard(vehicles, "selv_"))
    except: pass

async def reject_callback(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    if query.from_user.id not in ADMIN_IDS: return
    user_id = int(query.data.split("_")[1])
    set_driver(user_id, approval_status="rejected")
    add_allowed_user(user_id, status="rejected")
    try:
        await query.edit_message_text(f"❌ تم رفض المستخدم {user_id}")
    except: pass
    try:
        await context.bot.send_message(chat_id=user_id, text="عذراً، لم يتم قبولك. يمكنك التواصل مع الإدارة.")
    except: pass

# Admin submenu callbacks
async def admin_main_callback(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    await query.edit_message_text("🕹️ لوحة التحكم:", reply_markup=admin_main_keyboard())

async def admin_vehicles_menu(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    await query.edit_message_text("🚘 تعديل المركبات:", reply_markup=admin_vehicles_keyboard())

async def admin_drivers_menu(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    await query.edit_message_text("👤 إدارة السائقين:", reply_markup=admin_drivers_keyboard())

async def admin_vidange_menu(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    await query.edit_message_text("🛢️ إدارة الفيدانج:", reply_markup=admin_vidange_menu_keyboard())

async def admin_export_menu(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    await query.edit_message_text("📋 تصدير Excel:", reply_markup=admin_export_menu_keyboard())

async def admin_approve_list(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    with db_connection() as conn:
        cur = conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
        cur.execute("SELECT user_id, name FROM drivers WHERE approval_status='pending'")
        pending = cur.fetchall()
    if not pending:
        await query.edit_message_text("لا يوجد سائقون بانتظار القبول.")
        return
    buttons = [InlineKeyboardButton(f"{d['name']} ({d['user_id']})", callback_data=f"approve_{d['user_id']}") for d in pending]
    await query.edit_message_text("اختر سائقًا لقبوله:", reply_markup=InlineKeyboardMarkup([buttons[i:i+2] for i in range(0, len(buttons), 2)]))

async def admin_remove_driver_list(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    drivers = get_all_drivers()
    if not drivers:
        await query.edit_message_text("لا يوجد سائقون مسجلون.")
        return
    buttons = [InlineKeyboardButton(f"{d['name']} ({d['user_id']})", callback_data=f"rmdriver_{d['user_id']}") for d in drivers]
    await query.edit_message_text("اختر سائقًا لحذفه:", reply_markup=InlineKeyboardMarkup([buttons[i:i+2] for i in range(0, len(buttons), 2)]))

async def admin_drivers_list(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    drivers = get_all_drivers()
    if not drivers:
        await query.edit_message_text("لا يوجد سائقون نشطون.")
        return
    text = "👤 السائقون النشطون:\n" + "\n".join(f"• {d['name']} ({d['user_id']})" for d in drivers)
    await query.edit_message_text(text)

async def confirm_remove_driver(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    user_id = int(query.data.split("_")[1])
    await query.edit_message_text(f"هل أنت متأكد من حذف السائق {user_id}؟", reply_markup=InlineKeyboardMarkup([
        [InlineKeyboardButton("نعم", callback_data=f"confirmrm_{user_id}"),
         InlineKeyboardButton("إلغاء", callback_data="admin_drivers")]
    ]))

async def confirm_remove_driver_exec(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    if query.from_user.id not in ADMIN_IDS: return
    user_id = int(query.data.split("_")[1])
    remove_driver(user_id)
    await query.edit_message_text(f"✅ تم حذف السائق {user_id}.")
    try:
        await context.bot.send_message(chat_id=user_id, text="تم إلغاء صلاحيتك لاستخدام البوت.")
    except: pass

# Admin input handler
async def admin_input_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    if user_id not in ADMIN_IDS: return
    text = update.message.text.strip()
    if context.user_data.get("vidange_modify"):
        info = context.user_data.pop("vidange_modify")
        if not text.isdigit():
            await update.message.reply_text("يجب أن يكون الكيلومتر رقماً. حاول مرة أخرى:",
                                            reply_markup=InlineKeyboardMarkup([[InlineKeyboardButton("إلغاء", callback_data="cancel_input"),
                                                                               InlineKeyboardButton("↩️ رجوع", callback_data="admin_main")]]))
            context.user_data["vidange_modify"] = info
            return
        km = int(text)
        vehicle = info["vehicle"]
        last_km = get_latest_km(vehicle)
        if last_km is not None and km <= last_km:
            await update.message.reply_text(f"⚠️ الكيلومتر يجب أن يكون أكبر من آخر قراءة ({last_km} كم). أعد الإدخال.",
                                            reply_markup=InlineKeyboardMarkup([[InlineKeyboardButton("إلغاء", callback_data="cancel_input"),
                                                                               InlineKeyboardButton("↩️ رجوع", callback_data="admin_main")]]))
            context.user_data["vidange_modify"] = info
            return
        add_km_reading(vehicle, km, driver_name="مشرف")
        set_last_vidange_km(vehicle, km)
        try:
            await context.bot.edit_message_text(
                chat_id=ADMIN_GROUP_ID,
                message_id=info["message_id"],
                text=f"✅ تم تعديل الفيدانج للمركبة {vehicle} إلى {km} كم."
            )
        except: pass
        try:
            await context.bot.send_message(chat_id=info["user_id"], text=f"✅ تم تحديث الفيدانج للمركبة {vehicle} بقيمة {km} كم (بعد المراجعة).")
        except: pass
        await update.message.reply_text(f"✅ تم تحديث الفيدانج لـ {vehicle} = {km} كم.")
        return
    if context.user_data.get("admin_urgentvid"):
        context.user_data.pop("admin_urgentvid")
        parts = text.split()
        if len(parts) != 2 or not parts[1].isdigit():
            await update.message.reply_text("صيغة خاطئة. استخدم: CODE KM",
                                            reply_markup=InlineKeyboardMarkup([[InlineKeyboardButton("إلغاء", callback_data="cancel_input"),
                                                                               InlineKeyboardButton("↩️ رجوع", callback_data="admin_main")]]))
            return
        code = parts[0].upper()
        km = int(parts[1])
        if code not in get_all_vehicles():
            await update.message.reply_text("المركبة غير موجودة.")
            return
        set_last_vidange_km(code, km)
        latest = get_latest_km(code)
        if latest is not None and latest >= km + 9000:
            vidange_problem_id = add_problem(0, "نظام (عاجل)", code, f"Vidange عاجل {code}", "نظام")
            await context.bot.send_message(
                chat_id=ADMIN_GROUP_ID, message_thread_id=TOPIC_VIDANGE,
                text=f"🚨 فيدانج عاجل: المركبة {code}\nالعداد الحالي: {latest} كم\nآخر فيدانج (محدث): {km} كم\n🔴 الحالة: قيد الانتظار",
                reply_markup=build_problem_keyboard(vidange_problem_id)
            )
        await update.message.reply_text(f"✅ تم تعيين آخر فيدانج عاجل للمركبة {code} = {km} كم.")
        return
    if context.user_data.get("admin_add_veh"):
        context.user_data.pop("admin_add_veh")
        add_vehicle(text.upper())
        await update.message.reply_text(f"✅ تمت إضافة {text.upper()}.")
    elif context.user_data.get("admin_rem_veh"):
        context.user_data.pop("admin_rem_veh")
        remove_vehicle(text.upper())
        await update.message.reply_text(f"🗑️ تم حذف {text.upper()}.")
    elif context.user_data.get("admin_setvid"):
        context.user_data.pop("admin_setvid")
        parts = text.split()
        if len(parts) != 2 or not parts[1].isdigit():
            await update.message.reply_text("صيغة خاطئة. استخدم: CODE KM",
                                            reply_markup=InlineKeyboardMarkup([[InlineKeyboardButton("إلغاء", callback_data="cancel_input"),
                                                                               InlineKeyboardButton("↩️ رجوع", callback_data="admin_main")]]))
            return
        code = parts[0].upper()
        km = int(parts[1])
        if code not in get_all_vehicles():
            await update.message.reply_text("المركبة غير موجودة.")
            return
        set_last_vidange_km(code, km)
        await update.message.reply_text(f"✅ تم تعيين آخر فيدانج لـ {code} = {km} كم.")

# Settings callbacks
async def settings_help(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    videos = get_all_help_videos()
    if not videos:
        await query.edit_message_text("لا يوجد فيديو تعليمي حالياً.")
        return
    for v in videos:
        try:
            await context.bot.send_video(chat_id=query.from_user.id, video=v["file_id"], caption=v.get("description"))
        except: pass
    await query.edit_message_text("تم إرسال الفيديوهات التعليمية.")

async def settings_history_callback(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    user_id = query.from_user.id
    problems = get_driver_problems(user_id)
    if not problems:
        await query.edit_message_text("لا توجد شكاوي مسجلة.")
        return
    text = "📜 سجل شكاويي:\n"
    for p in problems[:10]:
        text += f"#{p['id']} | {p['date']} | {p['problem_text'][:30]} | {status_icon_and_text(p)}\n"
    await query.edit_message_text(text)

async def settings_change_name(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    set_driver(query.from_user.id, state="name_entry")
    markup = InlineKeyboardMarkup([[InlineKeyboardButton("إلغاء", callback_data="cancel_name")]])
    await query.edit_message_text("أرسل اسمك الجديد:", reply_markup=markup)

async def cancel_name_callback(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    set_driver(query.from_user.id, state="idle")
    await query.edit_message_text("تم إلغاء تغيير الاسم.")

async def settings_change_vehicle(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    vehicles = get_all_vehicles()
    await query.edit_message_text("اختر مركبتك الجديدة:", reply_markup=vehicle_inline_keyboard(vehicles, "selv_"))

# Help video commands
async def set_help_cmd(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.effective_user.id not in ADMIN_IDS: return
    if not update.message.video and not update.message.reply_to_message:
        await update.message.reply_text("أرسل الفيديو (كملف فيديو) أو رد على فيديو.")
        return
    if update.message.reply_to_message and update.message.reply_to_message.video:
        file_id = update.message.reply_to_message.video.file_id
        desc = update.message.text or ""
        add_help_video(file_id, desc)
        await update.message.reply_text("تم حفظ الفيديو بنجاح.")
    elif update.message.video:
        file_id = update.message.video.file_id
        desc = update.message.caption or ""
        add_help_video(file_id, desc)
        await update.message.reply_text("تم حفظ الفيديو بنجاح.")
    else:
        await update.message.reply_text("الرجاء إرسال فيديو أو الرد على فيديو.")

async def remove_help_cmd(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.effective_user.id not in ADMIN_IDS: return
    videos = get_all_help_videos()
    if not videos:
        await update.message.reply_text("لا توجد فيديوهات للحذف.")
        return
    buttons = []
    for v in videos:
        desc = v["description"] or f"فيديو {v['id']}"
        buttons.append([InlineKeyboardButton(f"🗑️ {desc}", callback_data=f"delhelp_{v['id']}")])
    await update.message.reply_text("اختر فيديو للحذف:", reply_markup=InlineKeyboardMarkup(buttons))

async def delete_help_video_callback(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    if query.from_user.id not in ADMIN_IDS: return
    video_id = int(query.data.split("_")[1])
    delete_help_video(video_id)
    await query.edit_message_text("✅ تم حذف الفيديو.")

# ----------------------------------------------------------------------
# Improved Broadcast command (with detailed report)
# ----------------------------------------------------------------------
async def broadcast_cmd(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.effective_user.id not in ADMIN_IDS:
        return
    if not context.args:
        await update.message.reply_text("استخدم: /broadcast <النص>")
        return

    message = " ".join(context.args)
    drivers = get_all_drivers()
    if not drivers:
        await update.message.reply_text("لا يوجد سائقون مسجلون.")
        return

    success = 0
    failed = 0
    for d in drivers:
        try:
            await context.bot.send_message(chat_id=d["user_id"], text=message)
            success += 1
        except Exception as e:
            logging.warning(f"Broadcast failed for {d['user_id']}: {e}")
            failed += 1

    await update.message.reply_text(f"✅ تم إرسال الرسالة إلى {success} سائق.\n❌ فشل الإرسال إلى {failed}.")

# ----------------------------------------------------------------------
# Dashboard functions (fixed to use thread_id from calling message)
# ----------------------------------------------------------------------
async def _send_dashboard(chat_id: int, thread_id: int = None):
    vehicles = get_all_vehicles()
    if not vehicles:
        return
    buttons = [InlineKeyboardButton(dashboard_button_text(v), callback_data=f"hist_{v}") for v in vehicles]
    markup = InlineKeyboardMarkup([buttons[i:i+2] for i in range(0, len(buttons), 2)])
    try:
        if thread_id:
            await app.bot.send_message(chat_id=chat_id, message_thread_id=thread_id,
                                       text="📊 الحالة اليومية للمركبات:", reply_markup=markup)
        else:
            await app.bot.send_message(chat_id=chat_id,
                                       text="📊 الحالة اليومية للمركبات:", reply_markup=markup)
    except Exception as e:
        logging.warning(f"Dashboard send error: {e}")

async def dashboard_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    chat_id = update.effective_chat.id
    thread_id = update.message.message_thread_id if update.message else None
    await _send_dashboard(chat_id, thread_id)
    if update.message:
        await update.message.reply_text("تم إرسال لوحة القيادة.")

async def admin_dash(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    thread_id = query.message.message_thread_id if query.message else None
    if not thread_id:
        thread_id = TOPIC_GENERAL
    await _send_dashboard(ADMIN_GROUP_ID, thread_id)
    await query.edit_message_text("تم إرسال لوحة القيادة.")

async def scheduled_dashboard(context: ContextTypes.DEFAULT_TYPE):
    await _send_dashboard(ADMIN_GROUP_ID, TOPIC_GENERAL)

async def weekly_excel(context: ContextTypes.DEFAULT_TYPE):
    file = generate_problems_excel()
    try:
        await context.bot.send_document(chat_id=ADMIN_GROUP_ID, message_thread_id=TOPIC_GENERAL, document=file, filename="المشاكل_الأسبوعي.xlsx")
    except: pass

def schedule_jobs(app: Application):
    now = datetime.now(TZ)
    target = time(7, 30, 0)
    next_daily = datetime.combine(now.date(), target, tzinfo=TZ)
    if now >= next_daily:
        next_daily += timedelta(days=1)
    app.job_queue.run_repeating(scheduled_dashboard, interval=24*60*60, first=next_daily)
    days_until_sat = (5 - now.weekday()) % 7
    next_sat = datetime.combine(now.date() + timedelta(days=days_until_sat), target, tzinfo=TZ)
    if now >= next_sat:
        next_sat += timedelta(days=7)
    app.job_queue.run_repeating(weekly_excel, interval=7*24*60*60, first=next_sat)

# ----------------------------------------------------------------------
# Export functions
# ----------------------------------------------------------------------
async def export_problems(update: Update, context: ContextTypes.DEFAULT_TYPE):
    file = generate_problems_excel()
    if update.message:
        await update.message.reply_document(document=file, filename="المشاكل.xlsx")
    else:
        await context.bot.send_document(chat_id=update.effective_chat.id, document=file, filename="المشاكل.xlsx")

async def export_vidange(update: Update, context: ContextTypes.DEFAULT_TYPE):
    file = generate_vidange_excel()
    if update.message:
        await update.message.reply_document(document=file, filename="الفيدانج.xlsx")
    else:
        await context.bot.send_document(chat_id=update.effective_chat.id, document=file, filename="الفيدانج.xlsx")

async def export_vidange_vehicle(update: Update, context: ContextTypes.DEFAULT_TYPE):
    code = context.args[0].upper() if context.args else None
    if not code or code not in get_all_vehicles():
        await update.message.reply_text("استخدم: /vidange <CODE> مع رمز مركبة صحيح.")
        return
    file = generate_vidange_excel(vehicle_code=code)
    await update.message.reply_document(document=file, filename=f"فيدانج_{code}.xlsx")

# Sticky admin panel command
async def set_admin_panel(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.effective_user.id not in ADMIN_IDS:
        return
    await update.message.reply_text(
        "تم تفعيل لوحة التحكم السريعة. استخدم الزر أدناه للوصول السريع.",
        reply_markup=ADMIN_STICKY_KEYBOARD
    )

async def remove_admin_panel(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.effective_user.id not in ADMIN_IDS:
        return
    await update.message.reply_text(
        "تم إخفاء لوحة التحكم السريعة.",
        reply_markup=ReplyKeyboardRemove()
    )

# Admin action callbacks (addveh, remveh, etc.)
async def admin_addveh(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    if query.from_user.id not in ADMIN_IDS: return
    markup = InlineKeyboardMarkup([
        [InlineKeyboardButton("إلغاء", callback_data="cancel_input"),
         InlineKeyboardButton("↩️ رجوع", callback_data="admin_main")]
    ])
    await query.edit_message_text("أرسل رمز المركبة الجديدة:", reply_markup=markup)
    context.user_data["admin_add_veh"] = True

async def admin_remveh(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    if query.from_user.id not in ADMIN_IDS: return
    markup = InlineKeyboardMarkup([
        [InlineKeyboardButton("إلغاء", callback_data="cancel_input"),
         InlineKeyboardButton("↩️ رجوع", callback_data="admin_main")]
    ])
    await query.edit_message_text("أرسل رمز المركبة المراد حذفها:", reply_markup=markup)
    context.user_data["admin_rem_veh"] = True

async def admin_setvid(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    if query.from_user.id not in ADMIN_IDS: return
    markup = InlineKeyboardMarkup([
        [InlineKeyboardButton("إلغاء", callback_data="cancel_input"),
         InlineKeyboardButton("↩️ رجوع", callback_data="admin_main")]
    ])
    await query.edit_message_text("أرسل رمز المركبة ثم الكيلومتر (مثال: M02 150000):", reply_markup=markup)
    context.user_data["admin_setvid"] = True

async def admin_urgentvid(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    if query.from_user.id not in ADMIN_IDS: return
    markup = InlineKeyboardMarkup([
        [InlineKeyboardButton("إلغاء", callback_data="cancel_input"),
         InlineKeyboardButton("↩️ رجوع", callback_data="admin_main")]
    ])
    await query.edit_message_text("أرسل رمز المركبة ثم الكيلومتر الجديد للفيدانج العاجل (مثال: M02 158000):", reply_markup=markup)
    context.user_data["admin_urgentvid"] = True

async def admin_export(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    if query.from_user.id not in ADMIN_IDS: return
    file = generate_problems_excel()
    await context.bot.send_document(chat_id=query.message.chat_id, document=file, filename="المشاكل.xlsx")
    await query.edit_message_text("تم إرسال ملف المشاكل.")

async def admin_vid(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    if query.from_user.id not in ADMIN_IDS: return
    file = generate_vidange_excel()
    await context.bot.send_document(chat_id=query.message.chat_id, document=file, filename="الفيدانج.xlsx")
    await query.edit_message_text("تم إرسال ملف الفيدانج.")

async def admin_listveh(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    if query.from_user.id not in ADMIN_IDS: return
    vehicles = get_all_vehicles()
    text = "🚘 المركبات المتاحة:\n" + "\n".join(f"• {v}" for v in vehicles) if vehicles else "لا توجد مركبات."
    await query.edit_message_text(text)

async def vehicle_history_callback(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    vehicle = query.data.split("_", 1)[1]
    with db_connection() as conn:
        cur = conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
        cur.execute("SELECT * FROM problems WHERE vehicle=%s ORDER BY date DESC", (vehicle,))
        problems = [dict(r) for r in cur.fetchall()]
        cur.execute("SELECT date, km FROM km_readings WHERE vehicle=%s ORDER BY date DESC LIMIT 5", (vehicle,))
        readings = cur.fetchall()
    text = f"🚘 تاريخ المركبة {vehicle}:\n"
    if problems:
        text += "\n📋 المشاكل:\n"
        for p in problems:
            link = f"https://t.me/c/{str(ADMIN_GROUP_ID)[4:]}/{p['group_message_id']}" if p['group_message_id'] else "#"
            text += f"  <a href='{link}'>#{p['id']}</a> | {p['date']} | {p['problem_text'][:40]} | {status_icon_and_text(p)}\n"
    else:
        text += "لا توجد مشاكل مسجلة.\n"
    if readings:
        text += "\n🛢️ آخر قراءات العداد:\n"
        for d, k in readings:
            text += f"  {d} - {k} كم\n"
    markup = InlineKeyboardMarkup([[InlineKeyboardButton("تم", callback_data="done_hist")]])
    await context.bot.send_message(chat_id=ADMIN_GROUP_ID, message_thread_id=TOPIC_HISTORY, text=text, parse_mode="HTML", reply_markup=markup)

async def done_history_callback(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    try:
        await query.message.delete()
    except Exception as e:
        logging.warning(f"Could not delete history message: {e}")

# Command handlers for /admin, /panel, etc.
async def admin_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text("🕹️ لوحة التحكم:", reply_markup=admin_main_keyboard())

async def panel_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    thread_id = update.message.message_thread_id if update.message else None
    keyboard = get_topic_keyboard(thread_id) if thread_id else None
    if keyboard:
        await update.message.reply_text("🕹️ لوحة التحكم الخاصة بهذا القسم:", reply_markup=keyboard)
    else:
        await update.message.reply_text("🕹️ لوحة التحكم الكاملة:", reply_markup=admin_main_keyboard())

# Error handler
async def error_handler(update: object, context: ContextTypes.DEFAULT_TYPE) -> None:
    logging.error(msg="Exception while handling an update:", exc_info=context.error)

# ----------------------------------------------------------------------
# Webhook and aiohttp (with app.start() for jobs)
# ----------------------------------------------------------------------
async def health(request):
    return web.Response(text="Bot is running")

async def telegram_webhook(request):
    data = await request.json()
    update = Update.de_json(data, app.bot)
    await app.process_update(update)
    return web.Response(status=200)

async def set_webhook(app: Application):
    webhook_url = f"{WEBHOOK_URL}/telegram"
    await app.bot.set_webhook(url=webhook_url)
    logging.info(f"Webhook set to {webhook_url}")

async def main():
    logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")
    init_db()
    global app
    app = Application.builder().token(BOT_TOKEN).build()

    # Command handlers
    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("admin", admin_command))
    app.add_handler(CommandHandler("panel", panel_command))
    app.add_handler(CommandHandler("dashboard", dashboard_command))
    app.add_handler(CommandHandler("sethelp", set_help_cmd))
    app.add_handler(CommandHandler("removehelp", remove_help_cmd))
    app.add_handler(CommandHandler("broadcast", broadcast_cmd))
    app.add_handler(CommandHandler("export", export_problems))
    app.add_handler(CommandHandler("export_vidange", export_vidange))
    app.add_handler(CommandHandler("vidange", export_vidange_vehicle))
    app.add_handler(CommandHandler("setadminpanel", set_admin_panel))
    app.add_handler(CommandHandler("removeadminpanel", remove_admin_panel))

    # Text / Media handlers
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND & filters.ChatType.PRIVATE, handle_text))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, admin_input_handler), group=1)
    app.add_handler(MessageHandler(filters.PHOTO | filters.VIDEO | filters.VOICE & filters.ChatType.PRIVATE, handle_media))

    # Callback handlers
    app.add_handler(CallbackQueryHandler(cancel_input, pattern="^cancel_input$"))
    app.add_handler(CallbackQueryHandler(vehicle_selection_callback, pattern="^selv_"))
    app.add_handler(CallbackQueryHandler(confirm_vehicle_callback, pattern="^confirmveh_"))
    app.add_handler(CallbackQueryHandler(cancel_vehicle_selection_callback, pattern="^cancel_veh$"))
    app.add_handler(CallbackQueryHandler(valide_callback, pattern="^val_"))
    app.add_handler(CallbackQueryHandler(ruglee_callback, pattern="^rug_"))
    app.add_handler(CallbackQueryHandler(fix_comment_callback, pattern="^fix_comment_"))
    app.add_handler(CallbackQueryHandler(comment_callback, pattern="^com_"))
    app.add_handler(CallbackQueryHandler(delete_callback, pattern="^del_"))
    app.add_handler(CallbackQueryHandler(confirm_delete_callback, pattern="^confirmdel_"))
    app.add_handler(CallbackQueryHandler(cancel_delete_callback, pattern="^cancel_"))
    app.add_handler(CallbackQueryHandler(validation_request_callback, pattern="^valreq_"))
    app.add_handler(CallbackQueryHandler(valrug_callback, pattern="^valrug_"))
    app.add_handler(CallbackQueryHandler(close_val_callback, pattern="^close_val_"))
    app.add_handler(CallbackQueryHandler(reopen_val_callback, pattern="^reopen_val_"))
    app.add_handler(CallbackQueryHandler(vidange_confirm_callback, pattern="^vidconfirm_"))
    app.add_handler(CallbackQueryHandler(vidange_modify_callback, pattern="^vidmodify_"))
    app.add_handler(CallbackQueryHandler(approve_callback, pattern="^approve_"))
    app.add_handler(CallbackQueryHandler(reject_callback, pattern="^reject_"))
    app.add_handler(CallbackQueryHandler(admin_main_callback, pattern="^admin_main$"))
    app.add_handler(CallbackQueryHandler(admin_vehicles_menu, pattern="^admin_vehicles$"))
    app.add_handler(CallbackQueryHandler(admin_drivers_menu, pattern="^admin_drivers$"))
    app.add_handler(CallbackQueryHandler(admin_vidange_menu, pattern="^admin_vidange_menu$"))
    app.add_handler(CallbackQueryHandler(admin_export_menu, pattern="^admin_export_menu$"))
    app.add_handler(CallbackQueryHandler(admin_approve_list, pattern="^admin_approve_list$"))
    app.add_handler(CallbackQueryHandler(admin_remove_driver_list, pattern="^admin_remove_driver_list$"))
    app.add_handler(CallbackQueryHandler(admin_drivers_list, pattern="^admin_drivers_list$"))
    app.add_handler(CallbackQueryHandler(confirm_remove_driver, pattern="^rmdriver_"))
    app.add_handler(CallbackQueryHandler(confirm_remove_driver_exec, pattern="^confirmrm_"))
    app.add_handler(CallbackQueryHandler(settings_help, pattern="^settings_help$"))
    app.add_handler(CallbackQueryHandler(settings_history_callback, pattern="^settings_history$"))
    app.add_handler(CallbackQueryHandler(settings_change_name, pattern="^settings_change_name$"))
    app.add_handler(CallbackQueryHandler(cancel_name_callback, pattern="^cancel_name$"))
    app.add_handler(CallbackQueryHandler(settings_change_vehicle, pattern="^settings_change_veh$"))
    app.add_handler(CallbackQueryHandler(delete_help_video_callback, pattern="^delhelp_"))
    app.add_handler(CallbackQueryHandler(admin_addveh, pattern="^admin_addveh$"))
    app.add_handler(CallbackQueryHandler(admin_remveh, pattern="^admin_remveh$"))
    app.add_handler(CallbackQueryHandler(admin_setvid, pattern="^admin_setvid$"))
    app.add_handler(CallbackQueryHandler(admin_urgentvid, pattern="^admin_urgentvid$"))
    app.add_handler(CallbackQueryHandler(admin_dash, pattern="^admin_dash$"))
    app.add_handler(CallbackQueryHandler(admin_export, pattern="^admin_export$"))
    app.add_handler(CallbackQueryHandler(admin_vid, pattern="^admin_vid$"))
    app.add_handler(CallbackQueryHandler(admin_listveh, pattern="^admin_listveh$"))
    app.add_handler(CallbackQueryHandler(vehicle_history_callback, pattern="^hist_"))
    app.add_handler(CallbackQueryHandler(done_history_callback, pattern="^done_hist$"))

    app.add_error_handler(error_handler)
    schedule_jobs(app)
    await app.initialize()
    await app.start()  # Start JobQueue
    await set_webhook(app)

    aio_app = web.Application()
    aio_app.router.add_get("/", health)
    aio_app.router.add_post("/telegram", telegram_webhook)

    port = int(os.environ.get("PORT", 5000))
    runner = web.AppRunner(aio_app)
    await runner.setup()
    site = web.TCPSite(runner, "0.0.0.0", port)
    await site.start()
    logging.info(f"Server listening on port {port}")
    await asyncio.Event().wait()

if __name__ == "__main__":
    asyncio.run(main())
