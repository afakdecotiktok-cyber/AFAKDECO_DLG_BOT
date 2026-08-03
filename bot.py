import asyncio
import json
import logging
import os
import re
from datetime import datetime
from io import BytesIO

import asyncpg
import openpyxl
import tornado.web
from PIL import Image, ImageDraw, ImageFont
from dotenv import load_dotenv
from telegram import (
    Update,
    InlineKeyboardButton,
    InlineKeyboardMarkup,
    InputFile,
    ReplyKeyboardMarkup,
    KeyboardButton,
)
from telegram.ext import (
    Application,
    CommandHandler,
    CallbackQueryHandler,
    MessageHandler,
    ConversationHandler,
    filters,
    ContextTypes,
)

load_dotenv()

# ---------- Config ----------
BOT_TOKEN = os.getenv("BOT_TOKEN")
DATABASE_URL = os.getenv("DATABASE_URL")          # Neon connection string
ADMIN_GROUP_ID = int(os.getenv("ADMIN_GROUP_ID")) # -1003913405243

# ---------- DB Pool ----------
pool = None

async def get_db_pool():
    global pool
    if pool is None:
        pool = await asyncpg.create_pool(DATABASE_URL, min_size=2, max_size=10)
    return pool

# ---------- Logging ----------
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s"
)
logger = logging.getLogger(__name__)

# ---------- Health check endpoint ----------
class HealthHandler(tornado.web.RequestHandler):
    def get(self):
        self.set_status(200)
        self.write("OK")

# ---------- Topic IDs Cache ----------
topic_cache = None

async def get_topics():
    global topic_cache
    if topic_cache is None:
        db = await get_db_pool()
        async with db.acquire() as conn:
            row = await conn.fetchrow(
                "SELECT * FROM group_settings WHERE group_id = $1", ADMIN_GROUP_ID
            )
            if row:
                topic_cache = {
                    "orders": row["orders_topic"],
                    "drivers": row["drivers_topic"],
                    "approvals": row["approvals_topic"],
                    "price": row["price_topic"],
                    "settings": row["settings_topic"],
                    "logs": row["logs_topic"],
                    "catalogue": row["catalogue_topic"],
                }
            else:
                topic_cache = {}
    return topic_cache

async def send_to_topic(context: ContextTypes.DEFAULT_TYPE, topic: str, text: str, **kwargs):
    topics = await get_topics()
    topic_id = topics.get(topic)
    if not topic_id:
        return None
    return await context.bot.send_message(
        chat_id=ADMIN_GROUP_ID, text=text, message_thread_id=topic_id, **kwargs
    )

# ---------- Driver Keyboard ----------
def get_driver_keyboard():
    return ReplyKeyboardMarkup(
        [[KeyboardButton("🔍 Search")]], resize_keyboard=True, one_time_keyboard=False
    )

# ---------- Driver List ----------
async def update_drivers_list(context: ContextTypes.DEFAULT_TYPE):
    db = await get_db_pool()
    async with db.acquire() as conn:
        drivers = await conn.fetch(
            "SELECT telegram_id, name, vehicle FROM drivers WHERE approved = true ORDER BY name"
        )
        rows_list = [f"👤 {d['name']} — 🚛 {d['vehicle']} (ID: {d['telegram_id']})" for d in drivers]
        text = "**🚚 Liste des chauffeurs**\n\n" + ("\n".join(rows_list) if rows_list else "Aucun chauffeur approuvé.")
        row = await conn.fetchrow("SELECT drivers_list_msg_id FROM group_settings WHERE group_id = $1", ADMIN_GROUP_ID)
        if not row:
            return
        drivers_topic = (await get_topics()).get("drivers")
        if not drivers_topic:
            return
        if row["drivers_list_msg_id"]:
            try:
                await context.bot.edit_message_text(
                    chat_id=ADMIN_GROUP_ID, message_id=row["drivers_list_msg_id"],
                    text=text, parse_mode='Markdown', message_thread_id=drivers_topic
                )
            except Exception:
                await send_and_pin_new(context, conn, drivers_topic, text)
        else:
            await send_and_pin_new(context, conn, drivers_topic, text)

async def send_and_pin_new(context, conn, drivers_topic, text):
    msg = await context.bot.send_message(
        chat_id=ADMIN_GROUP_ID, text=text, parse_mode='Markdown', message_thread_id=drivers_topic
    )
    await context.bot.pin_chat_message(
        chat_id=ADMIN_GROUP_ID, message_id=msg.message_id, message_thread_id=drivers_topic, disable_notification=True
    )
    await conn.execute(
        "UPDATE group_settings SET drivers_list_msg_id = $1 WHERE group_id = $2",
        msg.message_id, ADMIN_GROUP_ID
    )

# ---------- Order Number ----------
async def generate_order_number(conn):
    today_prefix = datetime.now().strftime("%Y%m%d")
    row = await conn.fetchrow(
        "SELECT order_number FROM orders WHERE order_number LIKE $1 ORDER BY order_number DESC LIMIT 1",
        f"{today_prefix}-%"
    )
    if row:
        last_num = int(row['order_number'].split('-')[-1])
        return f"{today_prefix}-{last_num + 1:03d}"
    return f"{today_prefix}-001"

# ---------- Image Generation ----------
async def generate_order_png(order):
    img = Image.new('RGB', (600, 400), color=(255, 255, 255))
    d = ImageDraw.Draw(img)
    try:
        font = ImageFont.truetype("DejaVuSans.ttf", 20)
    except:
        font = ImageFont.load_default()
    lines = [
        "**COMMANDE**",
        f"Code: {order['code']}",
        f"Collection: {order['collection']}",
        f"Forme: {order['form']}",
        f"Dimensions: {order['dimensions']}",
        f"Désign: {order['design']}",
        f"Couleur: {order['color_name']}",
        f"Quantité: {order['quantity']}",
        f"Prix unitaire: {order['unit_price']} DZD",
        f"Prix total: {order['total_price']} DZD",
        f"Date: {order['order_date'].strftime('%Y-%m-%d %H:%M')}",
    ]
    y = 20
    for line in lines:
        d.text((20, y), line, fill=(0, 0, 0), font=font)
        y += 30
    buf = BytesIO()
    img.save(buf, format='PNG')
    buf.seek(0)
    return buf

# ---------- Parse Article Code ----------
def parse_article_code(code: str):
    parts = code.strip().split('-')
    if len(parts) != 4:
        raise ValueError("Code must have 4 parts separated by '-'")
    collection = parts[0]
    form_size = parts[1]          # e.g. S57X200
    design = parts[2]
    color_code = parts[3]
    match = re.match(r'^([A-Za-z]+)(\d+X\d+)$', form_size)
    if not match:
        raise ValueError(f"Invalid form_size format: {form_size}")
    return collection, match.group(1).upper(), match.group(2), design, color_code

# ---------- Excel Processing ----------
async def process_catalogue_excel(context: ContextTypes.DEFAULT_TYPE, file_bytes: bytes):
    db = await get_db_pool()
    wb = openpyxl.load_workbook(BytesIO(file_bytes))
    ws = wb.active
    successes, errors = 0, []
    color_map = {}
    async with db.acquire() as conn:
        colors = await conn.fetch("SELECT code, name FROM color_codes")
        for c in colors:
            color_map[c['code']] = c['name']
    async with db.acquire() as conn:
        for row in ws.iter_rows(min_row=2, values_only=True):
            if not row or len(row) < 2:
                continue
            code_raw, price_raw = str(row[0]).strip(), row[1]
            if not code_raw or price_raw is None:
                continue
            try:
                price = float(price_raw)
                if price < 0:
                    errors.append(f"{code_raw}: prix négatif")
                    continue
            except:
                errors.append(f"{code_raw}: prix invalide")
                continue
            try:
                collection, form, dimensions, design, color_code = parse_article_code(code_raw)
            except Exception as e:
                errors.append(f"{code_raw}: {str(e)}")
                continue
            color_name = color_map.get(color_code)
            if not color_name:
                errors.append(f"{code_raw}: couleur inconnue ({color_code})")
                continue
            await conn.execute(
                """INSERT INTO articles (code, collection, form, dimensions, design, color_code, color_name, price)
                   VALUES ($1,$2,$3,$4,$5,$6,$7,$8)
                   ON CONFLICT (code) DO UPDATE SET
                       collection = EXCLUDED.collection, form = EXCLUDED.form,
                       dimensions = EXCLUDED.dimensions, design = EXCLUDED.design,
                       color_code = EXCLUDED.color_code, color_name = EXCLUDED.color_name,
                       price = EXCLUDED.price""",
                code_raw, collection, form, dimensions, design, color_code, color_name, price
            )
            successes += 1
    result_text = f"📤 Import catalogue terminé.\n✅ {successes} articles"
    if errors:
        result_text += f"\n❌ {len(errors)} erreurs:\n" + "\n".join(errors[:10])
    await send_to_topic(context, "logs", result_text)

# ---------- REBUILT Excel Upload Handler ----------
async def handle_excel_upload(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text("🔍 [Upload] Message received by bot.")
    if not await admin_only(update):
        await update.message.reply_text("⛔ [Upload] Only admins can upload files.")
        return
    doc = update.message.document
    if not doc:
        if update.message.photo:
            await update.message.reply_text("❌ [Upload] Photos are not accepted. Send as a Document (.xlsx).")
        elif update.message.text:
            await update.message.reply_text(f"ℹ️ [Upload] Text received: \"{update.message.text}\". Please send an Excel file.")
        else:
            await update.message.reply_text("❌ [Upload] No document found in this message.")
        return
    file_name = doc.file_name or "inconnu"
    await update.message.reply_text(f"📄 [Upload] File name: {file_name}")
    if not file_name.lower().endswith(".xlsx"):
        await update.message.reply_text("❌ [Upload] Only .xlsx files are accepted.")
        return
    thread_id = update.message.message_thread_id
    topics = await get_topics()
    catalogue_id = topics.get("catalogue")
    if thread_id != catalogue_id:
        await update.message.reply_text(
            f"ℹ️ [Upload] Wrong topic. Upload in Catalogue Upload topic (ID: {catalogue_id}). "
            f"Current topic ID: {thread_id}"
        )
        return
    await update.message.reply_text("🔄 [Upload] Processing file...")
    file = await doc.get_file()
    file_bytes = await file.download_as_bytearray()
    await process_catalogue_excel(context, file_bytes)

# ---------- Fallback handler ----------
async def fallback_log(update: Update, context: ContextTypes.DEFAULT_TYPE):
    logger.info(f"Fallback handler triggered: {update}")
    if update.message:
        content = "unknown"
        if update.message.text:
            content = f"text: {update.message.text}"
        elif update.message.document:
            content = f"document: {update.message.document.file_name}"
        elif update.message.photo:
            content = "photo"
        await update.message.reply_text(f"⚠️ [Fallback] Unhandled message: {content}")

# ---------- Admin check ----------
async def admin_only(update: Update):
    try:
        member = await update.get_bot().get_chat_member(ADMIN_GROUP_ID, update.effective_user.id)
        return member.status in ('administrator', 'creator')
    except:
        return False

# ---------- Driver Registration ----------
REG_NAME, REG_VEHICLE = range(2)

async def start_driver(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    db = await get_db_pool()
    async with db.acquire() as conn:
        driver = await conn.fetchrow("SELECT * FROM drivers WHERE telegram_id = $1 AND approved = true", user_id)
        if driver:
            await update.message.reply_text(
                "Vous êtes déjà enregistré. Utilisez le bouton 🔍 Search.",
                reply_markup=get_driver_keyboard()
            )
            return ConversationHandler.END
        pending = await conn.fetchrow("SELECT * FROM pending_drivers WHERE telegram_id = $1", user_id)
        if pending:
            await update.message.reply_text("Inscription en attente de validation.")
            return ConversationHandler.END
        context.user_data['reg'] = {}
        await update.message.reply_text("Bienvenue. Entrez votre nom complet :")
        return REG_NAME

async def reg_name(update: Update, context: ContextTypes.DEFAULT_TYPE):
    name = update.message.text.strip()
    if not name:
        await update.message.reply_text("Nom invalide. Réessayez :")
        return REG_NAME
    context.user_data['reg']['name'] = name
    await update.message.reply_text("Maintenant, le nom de votre véhicule :")
    return REG_VEHICLE

async def reg_vehicle(update: Update, context: ContextTypes.DEFAULT_TYPE):
    vehicle = update.message.text.strip()
    if not vehicle:
        await update.message.reply_text("Nom de véhicule invalide. Réessayez :")
        return REG_VEHICLE
    context.user_data['reg']['vehicle'] = vehicle
    user_id = update.effective_user.id
    db = await get_db_pool()
    async with db.acquire() as conn:
        await conn.execute(
            "INSERT INTO pending_drivers (telegram_id, name, vehicle) VALUES ($1, $2, $3)",
            user_id, context.user_data['reg']['name'], vehicle
        )
    keyboard = InlineKeyboardMarkup([
        [InlineKeyboardButton("✅ Approuver", callback_data=f"approve_{user_id}"),
         InlineKeyboardButton("❌ Rejeter", callback_data=f"reject_{user_id}")]
    ])
    await send_to_topic(context, "approvals",
        f"📝 Nouvelle inscription:\nNom: {context.user_data['reg']['name']}\nVéhicule: {vehicle}\nID: {user_id}",
        reply_markup=keyboard)
    await update.message.reply_text("Inscription envoyée. Vous recevrez une notification après validation.")
    context.user_data.pop('reg', None)
    return ConversationHandler.END

async def cancel_reg(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text("Inscription annulée.")
    context.user_data.pop('reg', None)
    return ConversationHandler.END

# ---------- Admin Approval ----------
async def handle_approval(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    if not await admin_only(update):
        await query.answer("Non autorisé.")
        return
    action, target_id_str = query.data.split('_', 1)
    target_id = int(target_id_str)
    db = await get_db_pool()
    async with db.acquire() as conn:
        pending = await conn.fetchrow("SELECT * FROM pending_drivers WHERE telegram_id = $1", target_id)
        if not pending:
            await query.message.edit_text("❌ Cette demande n'existe plus.")
            return
        if action == "approve":
            await conn.execute(
                "INSERT INTO drivers (telegram_id, name, vehicle, approved) VALUES ($1, $2, $3, true)"
                " ON CONFLICT (telegram_id) DO UPDATE SET name = $2, vehicle = $3, approved = true",
                target_id, pending['name'], pending['vehicle']
            )
            await conn.execute("DELETE FROM pending_drivers WHERE telegram_id = $1", target_id)
            await query.message.edit_text(f"✅ {pending['name']} approuvé.")
            try:
                await context.bot.send_message(chat_id=target_id,
                    text="🎉 Inscription approuvée. Utilisez le bouton 🔍 Search.",
                    reply_markup=get_driver_keyboard())
            except:
                pass
            await update_drivers_list(context)
        else:
            await conn.execute("DELETE FROM pending_drivers WHERE telegram_id = $1", target_id)
            await query.message.edit_text(f"❌ {pending['name']} rejeté.")
            try:
                await context.bot.send_message(chat_id=target_id, text="Désolé, inscription refusée.")
            except:
                pass

# ---------- Search Functions ----------
async def search_article(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    db = await get_db_pool()
    async with db.acquire() as conn:
        driver = await conn.fetchrow("SELECT * FROM drivers WHERE telegram_id = $1 AND approved = true", user_id)
        if not driver:
            await update.message.reply_text("Vous n'êtes pas un chauffeur approuvé.")
            return
        query = update.message.text.split(' ', 1)
        if len(query) < 2:
            await update.message.reply_text("Usage: /search <terme>")
            return
        term = query[1].strip()
        if len(term) < 2:
            await update.message.reply_text("Minimum 2 caractères.")
            return
        rows = await conn.fetch("""
            SELECT a.*, similarity(a.code, $1) AS sim,
                   COALESCE(h.cnt, 0) AS personal_orders, a.popularity,
                   (5 * similarity(a.code, $1) + 2 * COALESCE(h.cnt, 0) + 0.5 * a.popularity) AS combined_score
            FROM articles a
            LEFT JOIN (SELECT article_id, COUNT(*) as cnt FROM orders WHERE driver_id = $2 GROUP BY article_id) h
            ON h.article_id = a.id
            WHERE similarity(a.code, $1) > 0.1
            ORDER BY combined_score DESC LIMIT 5
        """, term, user_id)
        if not rows:
            await update.message.reply_text("Aucun article trouvé.", reply_markup=get_driver_keyboard())
            return
        keyboard = []
        for r in rows:
            avail = "" if r['available'] else " (Non dispo)"
            keyboard.append([InlineKeyboardButton(f"{r['code']} - {r['price']} DZD{avail}", callback_data=f"select_{r['id']}")])
        await update.message.reply_text(f"Résultats pour « {term} »:", reply_markup=InlineKeyboardMarkup(keyboard))

# ---------- Easy Search Conversation ----------
SEARCH_TERM = 1

async def search_button(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    db = await get_db_pool()
    async with db.acquire() as conn:
        driver = await conn.fetchrow("SELECT * FROM drivers WHERE telegram_id = $1 AND approved = true", user_id)
        if not driver:
            await update.message.reply_text("Vous n'êtes pas un chauffeur approuvé.")
            return ConversationHandler.END
    await update.message.reply_text("Que cherchez-vous ? (tapez le nom, la forme, la dimension…)")
    return SEARCH_TERM

async def receive_search_term(update: Update, context: ContextTypes.DEFAULT_TYPE):
    term = update.message.text.strip()
    if len(term) < 2:
        await update.message.reply_text("Veuillez entrer au moins 2 caractères.")
        return SEARCH_TERM
    user_id = update.effective_user.id
    db = await get_db_pool()
    async with db.acquire() as conn:
        rows = await conn.fetch("""
            SELECT a.*, similarity(a.code, $1) AS sim,
                   COALESCE(h.cnt, 0) AS personal_orders, a.popularity,
                   (5 * similarity(a.code, $1) + 2 * COALESCE(h.cnt, 0) + 0.5 * a.popularity) AS combined_score
            FROM articles a
            LEFT JOIN (SELECT article_id, COUNT(*) as cnt FROM orders WHERE driver_id = $2 GROUP BY article_id) h
            ON h.article_id = a.id
            WHERE similarity(a.code, $1) > 0.1
            ORDER BY combined_score DESC LIMIT 5
        """, term, user_id)
        if not rows:
            await update.message.reply_text("Aucun article trouvé.", reply_markup=get_driver_keyboard())
            return ConversationHandler.END
        keyboard = []
        for r in rows:
            avail = "" if r['available'] else " (Non dispo)"
            keyboard.append([InlineKeyboardButton(f"{r['code']} - {r['price']} DZD{avail}", callback_data=f"select_{r['id']}")])
        await update.message.reply_text(f"Résultats pour « {term} »:", reply_markup=InlineKeyboardMarkup(keyboard))
    return ConversationHandler.END

# ---------- Order Flow ----------
async def select_article(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    article_id = query.data.split('_', 1)[1]
    db = await get_db_pool()
    async with db.acquire() as conn:
        article = await conn.fetchrow("SELECT * FROM articles WHERE id = $1", article_id)
        if not article:
            await query.message.reply_text("Article introuvable.")
            return
        if not article['available']:
            await query.message.reply_text("❌ Article non disponible.")
            return
        context.user_data['order'] = {
            'article_id': str(article['id']),
            'code': article['code'],
            'price': float(article['price']),
            'article': article
        }
        await query.message.reply_text(
            f"Article: {article['code']}\nPrix unitaire: {article['price']} DZD\nEntrez la quantité :"
        )

async def handle_quantity(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if 'order' not in context.user_data or 'quantity' in context.user_data['order']:
        return
    try:
        qty = int(update.message.text.strip())
        if qty <= 0:
            raise ValueError
    except:
        await update.message.reply_text("Quantité invalide. Entrez un nombre positif :")
        return
    context.user_data['order']['quantity'] = qty
    article = context.user_data['order']['article']
    total = qty * context.user_data['order']['price']
    keyboard = InlineKeyboardMarkup([
        [InlineKeyboardButton("✅ Confirmer", callback_data="confirm_order"),
         InlineKeyboardButton("❌ Annuler", callback_data="cancel_order")]
    ])
    await update.message.reply_text(
        f"Récapitulatif:\nCode: {article['code']}\nCouleur: {article['color_name']}\n"
        f"Quantité: {qty}\nPrix total: {total:.2f} DZD\nConfirmer ?",
        reply_markup=keyboard
    )

async def confirm_order(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    if 'order' not in context.user_data:
        return
    order = context.user_data['order']
    driver_id = query.from_user.id
    db = await get_db_pool()
    async with db.acquire() as conn:
        driver = await conn.fetchrow("SELECT name, vehicle FROM drivers WHERE telegram_id = $1", driver_id)
        if not driver:
            await query.message.edit_text("Erreur: chauffeur non trouvé.")
            return
        article = await conn.fetchrow("SELECT * FROM articles WHERE id = $1", order['article_id'])
        if not article:
            await query.message.edit_text("Article introuvable.")
            return
        order_number = await generate_order_number(conn)
        unit_price = order['price']
        qty = order['quantity']
        total = unit_price * qty
        await conn.execute(
            "INSERT INTO orders (order_number, driver_id, article_id, quantity, unit_price, total_price) VALUES ($1, $2, $3, $4, $5, $6)",
            order_number, driver_id, order['article_id'], qty, unit_price, total
        )
        await conn.execute("UPDATE articles SET popularity = popularity + 0.01 WHERE id = $1", order['article_id'])
        order_data = {
            'code': article['code'],
            'collection': article['collection'],
            'form': article['form'],
            'dimensions': article['dimensions'],
            'design': article['design'],
            'color_name': article['color_name'],
            'quantity': qty,
            'unit_price': unit_price,
            'total_price': total,
            'order_date': datetime.now()
        }
        png = await generate_order_png(order_data)
        caption = f"🚛 Chauffeur: {driver['name']} — {driver['vehicle']}"
        topics = await get_topics()
        orders_topic = topics.get("orders")
        if orders_topic:
            await context.bot.send_photo(
                chat_id=ADMIN_GROUP_ID,
                photo=InputFile(png, filename=f"order_{order_number}.png"),
                caption=caption,
                message_thread_id=orders_topic
            )
        await query.message.edit_text(f"✅ Commande {order_number} envoyée.")
        context.user_data.pop('order', None)

async def cancel_order(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    await query.message.edit_text("Commande annulée.")
    context.user_data.pop('order', None)

# ---------- Admin Commands ----------
async def update_price(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not await admin_only(update):
        return
    args = context.args
    if len(args) != 2:
        await update.message.reply_text("Usage: /updateprice <code> <prix>")
        return
    code, price_str = args
    try:
        price = float(price_str)
    except:
        await update.message.reply_text("Prix invalide.")
        return
    db = await get_db_pool()
    async with db.acquire() as conn:
        res = await conn.execute("UPDATE articles SET price = $1 WHERE code = $2", price, code)
        if res == "UPDATE 0":
            await update.message.reply_text("Code introuvable.")
        else:
            await update.message.reply_text(f"Prix mis à jour pour {code}.")
            await send_to_topic(context, "logs", f"💰 Prix changé: {code} -> {price} DZD")

async def toggle_avail(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not await admin_only(update):
        return
    if not context.args:
        await update.message.reply_text("Usage: /toggleavail <code>")
        return
    code = context.args[0]
    db = await get_db_pool()
    async with db.acquire() as conn:
        article = await conn.fetchrow("SELECT available FROM articles WHERE code = $1", code)
        if not article:
            await update.message.reply_text("Code introuvable.")
            return
        new = not article['available']
        await conn.execute("UPDATE articles SET available = $1 WHERE code = $2", new, code)
        await update.message.reply_text(f"{code} maintenant {'disponible' if new else 'non disponible'}.")

async def remove_article(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not await admin_only(update):
        return
    if not context.args:
        await update.message.reply_text("Usage: /removearticle <code>")
        return
    code = context.args[0]
    db = await get_db_pool()
    async with db.acquire() as conn:
        res = await conn.execute("DELETE FROM articles WHERE code = $1", code)
        if res == "DELETE 0":
            await update.message.reply_text("Code introuvable.")
        else:
            await update.message.reply_text(f"✅ Article {code} supprimé.")
            await send_to_topic(context, "logs", f"🗑️ Article supprimé: {code}")

async def broadcast(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not await admin_only(update):
        return
    text = update.message.text.split(' ', 1)
    if len(text) < 2:
        await update.message.reply_text("Usage: /broadcast <message>")
        return
    message = text[1]
    db = await get_db_pool()
    async with db.acquire() as conn:
        drivers = await conn.fetch("SELECT telegram_id FROM drivers WHERE approved = true")
    sent, failed = 0, 0
    for d in drivers:
        try:
            await context.bot.send_message(chat_id=d['telegram_id'], text=message)
            sent += 1
        except:
            failed += 1
    await update.message.reply_text(f"Diffusion: {sent} envoyés, {failed} échecs.")

# ---------- /addarticle Conversation ----------
ADD_COLLECTION, ADD_FORM, ADD_DIMENSIONS, ADD_DESIGN, ADD_COLOR, ADD_PRICE = range(6)

async def add_article_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not await admin_only(update):
        return ConversationHandler.END
    await update.message.reply_text("Collection (ex: AMIRA):")
    return ADD_COLLECTION

async def add_collection(update, context):
    context.user_data['add_art'] = {'collection': update.message.text.strip()}
    await update.message.reply_text("Forme (lettres, ex: S, MQT):")
    return ADD_FORM

async def add_form(update, context):
    context.user_data['add_art']['form'] = update.message.text.strip().upper()
    await update.message.reply_text("Dimensions (ex: 57X200):")
    return ADD_DIMENSIONS

async def add_dimensions(update, context):
    dim = update.message.text.strip()
    if 'X' not in dim:
        await update.message.reply_text("Format invalide. Utilisez LargeurXHauteur.")
        return ADD_DIMENSIONS
    context.user_data['add_art']['dimensions'] = dim
    await update.message.reply_text("Désign (code):")
    return ADD_DESIGN

async def add_design(update, context):
    context.user_data['add_art']['design'] = update.message.text.strip()
    await update.message.reply_text("Code couleur (numéro):")
    return ADD_COLOR

async def add_color(update, context):
    color_code = update.message.text.strip()
    db = await get_db_pool()
    async with db.acquire() as conn:
        color = await conn.fetchrow("SELECT name FROM color_codes WHERE code = $1", color_code)
        if not color:
            await update.message.reply_text("Couleur inconnue. Ajoutez-la d'abord. Réessayez:")
            return ADD_COLOR
    context.user_data['add_art']['color_code'] = color_code
    context.user_data['add_art']['color_name'] = color['name']
    await update.message.reply_text("Prix (DZD):")
    return ADD_PRICE

async def add_price(update, context):
    try:
        price = float(update.message.text.strip())
    except:
        await update.message.reply_text("Prix invalide.")
        return ADD_PRICE
    art = context.user_data['add_art']
    code = f"{art['collection']}-{art['form']}{art['dimensions']}-{art['design']}-{art['color_code']}"
    db = await get_db_pool()
    async with db.acquire() as conn:
        try:
            await conn.execute(
                "INSERT INTO articles (code, collection, form, dimensions, design, color_code, color_name, price) VALUES ($1,$2,$3,$4,$5,$6,$7,$8)",
                code, art['collection'], art['form'], art['dimensions'], art['design'],
                art['color_code'], art['color_name'], price
            )
            await update.message.reply_text(f"✅ Article ajouté: {code}")
            await send_to_topic(context, "logs", f"➕ Article ajouté: {code}")
        except asyncpg.UniqueViolationError:
            await update.message.reply_text("Ce code existe déjà.")
    context.user_data.pop('add_art', None)
    return ConversationHandler.END

async def cancel_add(update, context):
    await update.message.reply_text("Ajout annulé.")
    context.user_data.pop('add_art', None)
    return ConversationHandler.END

# ---------- Error handler ----------
async def error_handler(update: object, context: ContextTypes.DEFAULT_TYPE) -> None:
    logger.error(msg="Exception while handling an update:", exc_info=context.error)

# ---------- Main (webhook + health endpoint) ----------
def main():
    app = Application.builder().token(BOT_TOKEN).build()

    # --- Register all handlers (keep everything exactly as before) ---
    reg_conv = ConversationHandler(
        entry_points=[CommandHandler("start", start_driver, filters.ChatType.PRIVATE)],
        states={
            REG_NAME: [MessageHandler(filters.TEXT & ~filters.COMMAND, reg_name)],
            REG_VEHICLE: [MessageHandler(filters.TEXT & ~filters.COMMAND, reg_vehicle)],
        },
        fallbacks=[CommandHandler("cancel", cancel_reg)]
    )
    app.add_handler(reg_conv)

    search_conv = ConversationHandler(
        entry_points=[MessageHandler(filters.Text(["🔍 Search"]), search_button)],
        states={SEARCH_TERM: [MessageHandler(filters.TEXT & ~filters.COMMAND, receive_search_term)]},
        fallbacks=[]
    )
    app.add_handler(search_conv)

    add_conv = ConversationHandler(
        entry_points=[CommandHandler("addarticle", add_article_start, filters.Chat(ADMIN_GROUP_ID))],
        states={
            ADD_COLLECTION: [MessageHandler(filters.TEXT & ~filters.COMMAND, add_collection)],
            ADD_FORM: [MessageHandler(filters.TEXT & ~filters.COMMAND, add_form)],
            ADD_DIMENSIONS: [MessageHandler(filters.TEXT & ~filters.COMMAND, add_dimensions)],
            ADD_DESIGN: [MessageHandler(filters.TEXT & ~filters.COMMAND, add_design)],
            ADD_COLOR: [MessageHandler(filters.TEXT & ~filters.COMMAND, add_color)],
            ADD_PRICE: [MessageHandler(filters.TEXT & ~filters.COMMAND, add_price)],
        },
        fallbacks=[CommandHandler("cancel", cancel_add)]
    )
    app.add_handler(add_conv)

    # Command handlers
    app.add_handler(CommandHandler("search", search_article, filters.ChatType.PRIVATE))
    app.add_handler(CommandHandler("s", search_article, filters.ChatType.PRIVATE))
    app.add_handler(CommandHandler("updateprice", update_price, filters.Chat(ADMIN_GROUP_ID)))
    app.add_handler(CommandHandler("toggleavail", toggle_avail, filters.Chat(ADMIN_GROUP_ID)))
    app.add_handler(CommandHandler("removearticle", remove_article, filters.Chat(ADMIN_GROUP_ID)))
    app.add_handler(CommandHandler("broadcast", broadcast, filters.Chat(ADMIN_GROUP_ID)))
    async def ping(update: Update, context: ContextTypes.DEFAULT_TYPE):
        await update.message.reply_text("pong")
    app.add_handler(CommandHandler("ping", ping, filters.Chat(ADMIN_GROUP_ID)))

    # Callback handlers
    app.add_handler(CallbackQueryHandler(handle_approval, pattern=r"^(approve_|reject_)"))
    app.add_handler(CallbackQueryHandler(select_article, pattern=r"^select_"))
    app.add_handler(CallbackQueryHandler(confirm_order, pattern="^confirm_order$"))
    app.add_handler(CallbackQueryHandler(cancel_order, pattern="^cancel_order$"))

    # Document handler (catches any document in admin group)
    app.add_handler(MessageHandler(
        filters.Document.ALL & filters.Chat(ADMIN_GROUP_ID),
        handle_excel_upload
    ))

    # Fallback for any other message in admin group
    app.add_handler(MessageHandler(
        filters.ALL & filters.Chat(ADMIN_GROUP_ID) & ~filters.COMMAND,
        fallback_log
    ))

    # Quantity handler for private chat
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND & filters.ChatType.PRIVATE, handle_quantity))

    app.add_error_handler(error_handler)

    # --- Webhook with health endpoint ---
    render_url = os.environ.get("RENDER_EXTERNAL_URL", "")
    port = int(os.environ.get("PORT", "8443"))

    if not render_url:
        logger.warning("RENDER_EXTERNAL_URL not set, falling back to polling")
        app.run_polling()
        return

    webhook_url = f"{render_url}/{BOT_TOKEN}"

    from telegram.ext import Updater
    from telegram.ext._webhookhandler import WebhookHandler

    # Custom Tornado application with a health check
    class CustomTornadoApp(tornado.web.Application):
        pass

    webhook_handler = WebhookHandler(app, webhook_url)
    tornado_app = CustomTornadoApp([
        (r"/health", HealthHandler),          # health endpoint
        (f"/{BOT_TOKEN}", webhook_handler),   # webhook path
    ])

    async def start():
        await app.initialize()
        await app.bot.set_webhook(url=webhook_url)
        updater = Updater(bot=app.bot, update_queue=app.update_queue)
        await updater.start_webhook(
            listen="0.0.0.0",
            port=port,
            url_path=BOT_TOKEN,
            webhook_url=webhook_url,
            webhook_app=tornado_app,
            drop_pending_updates=True,
        )
        await updater._started.wait()
        await updater._stop_event.wait()

    try:
        asyncio.run(start())
    except KeyboardInterrupt:
        pass
    finally:
        asyncio.run(app.shutdown())

if __name__ == "__main__":
    main()
