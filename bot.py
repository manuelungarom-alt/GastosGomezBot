import os
import re
import json
import logging
from io import BytesIO
from datetime import datetime

import openpyxl
import dropbox
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import (
    Application,
    CommandHandler,
    MessageHandler,
    CallbackQueryHandler,
    ContextTypes,
    filters,
)

logging.basicConfig(level=logging.INFO)
log = logging.getLogger("bot_familia")

# ---------------------------------------------------------------------------
# Config (todo esto viene de variables de entorno, nunca hardcodeado)
# ---------------------------------------------------------------------------
TELEGRAM_TOKEN = os.environ["TELEGRAM_TOKEN"]
DROPBOX_REFRESH_TOKEN = os.environ["DROPBOX_REFRESH_TOKEN"]
DROPBOX_APP_KEY = os.environ["DROPBOX_APP_KEY"]
DROPBOX_APP_SECRET = os.environ["DROPBOX_APP_SECRET"]
EXCEL_PATH = os.environ.get("EXCEL_PATH", "/Gastos_Familia_v3.xlsx")
USERS_PATH = os.environ.get("USERS_PATH", "/usuarios.json")

ADMIN_NAMES = ["Luis", "Vero"]
ALL_NAMES = ["Luis", "Vero", "Malena", "Luka", "Mateo"]

FONDOS = {
    "vacaciones": "Vacaciones",
    "jubilacion": "Jubilación",
    "jubilación": "Jubilación",
    "emergencia": "Fondo de Emergencia",
}

MEDIOS = {
    "efectivo": "Efectivo",
    "debito": "Débito",
    "débito": "Débito",
    "transferencia": "Transferencia",
}

RETIRO_WORDS = ["retiro", "retiré", "retire", "saque", "saqué", "saco"]
INGRESO_WORDS = ["ingreso", "ingresé", "ingrese", "cobre", "cobré", "cobro"]

dbx = dropbox.Dropbox(
    oauth2_refresh_token=DROPBOX_REFRESH_TOKEN,
    app_key=DROPBOX_APP_KEY,
    app_secret=DROPBOX_APP_SECRET,
)

# ---------------------------------------------------------------------------
# Dropbox helpers
# ---------------------------------------------------------------------------
def load_users() -> dict:
    try:
        _, res = dbx.files_download(USERS_PATH)
        return json.loads(res.content.decode("utf-8"))
    except Exception as e:
        log.warning(f"No se pudo leer {USERS_PATH} (puede ser normal si aun no existe): {e}")
        return {}


def save_users(users: dict):
    dbx.files_upload(
        json.dumps(users, ensure_ascii=False, indent=2).encode("utf-8"),
        USERS_PATH,
        mode=dropbox.files.WriteMode.overwrite,
    )


def download_excel() -> openpyxl.Workbook:
    _, res = dbx.files_download(EXCEL_PATH)
    return openpyxl.load_workbook(BytesIO(res.content))


def upload_excel(wb: openpyxl.Workbook):
    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)
    dbx.files_upload(buf.read(), EXCEL_PATH, mode=dropbox.files.WriteMode.overwrite)


def first_empty_row(ws, start=3, col=1) -> int:
    r = start
    while ws.cell(row=r, column=col).value not in (None, ""):
        r += 1
    return r


# ---------------------------------------------------------------------------
# Parseo del mensaje
# ---------------------------------------------------------------------------
def parse_message(text: str) -> dict | None:
    text_low = text.lower()

    m = re.search(r"\$?\s*(\d{1,3}(?:[.,]\d{3})+|\d+)", text)
    if not m:
        return None
    monto_str = m.group(1).replace(".", "").replace(",", "")
    try:
        monto = float(monto_str)
    except ValueError:
        return None

    fondo = None
    for k, v in FONDOS.items():
        if k in text_low:
            fondo = v
            break

    is_retiro = any(w in text_low for w in RETIRO_WORDS)
    is_ingreso = any(w in text_low for w in INGRESO_WORDS)

    medio = "Transferencia"  # default razonable si no lo especifican
    medio_encontrado = None
    for k, v in MEDIOS.items():
        if k in text_low:
            medio = v
            medio_encontrado = k
            break

    # Motivo: el mensaje original, sacando el monto y las palabras clave usadas
    motivo = text.replace(m.group(0), "")
    palabras_a_sacar = list(FONDOS.keys()) + RETIRO_WORDS + INGRESO_WORDS
    if medio_encontrado:
        palabras_a_sacar.append(medio_encontrado)
    for palabra in palabras_a_sacar:
        motivo = re.sub(rf"\b{re.escape(palabra)}\b", "", motivo, flags=re.IGNORECASE)
    motivo = re.sub(r"\s+", " ", motivo).strip(" ,.-")
    if not motivo:
        motivo = "Sin descripción"

    return {
        "monto": monto,
        "fondo": fondo,
        "is_retiro": is_retiro,
        "is_ingreso": is_ingreso,
        "medio": medio,
        "motivo": motivo,
    }


# ---------------------------------------------------------------------------
# Escritura en Excel
# ---------------------------------------------------------------------------
def registrar_movimiento(nombre: str, tipo: str, monto: float, motivo: str, medio: str):
    wb = download_excel()
    ws = wb["Movimientos"]
    r = first_empty_row(ws, start=3, col=1)
    ws.cell(row=r, column=1, value=datetime.now())
    ws.cell(row=r, column=1).number_format = "DD/MM/YYYY"
    ws.cell(row=r, column=2, value=tipo)
    ws.cell(row=r, column=3, value=monto)
    ws.cell(row=r, column=3).number_format = "$#,##0"
    ws.cell(row=r, column=4, value=motivo)
    ws.cell(row=r, column=5, value=medio)
    ws.cell(row=r, column=6, value=nombre)
    ws.cell(row=r, column=7, value=f'=IF(A{r}="","",TEXT(A{r},"YYYY-MM"))')
    upload_excel(wb)


def registrar_ahorro(nombre: str, fondo: str, tipo: str, monto: float):
    wb = download_excel()
    ws = wb["Ahorro_Detalle"]
    r = first_empty_row(ws, start=3, col=1)
    ws.cell(row=r, column=1, value=datetime.now())
    ws.cell(row=r, column=1).number_format = "DD/MM/YYYY"
    ws.cell(row=r, column=2, value=fondo)
    ws.cell(row=r, column=3, value=tipo)
    ws.cell(row=r, column=4, value=monto)
    ws.cell(row=r, column=4).number_format = "$#,##0"
    ws.cell(row=r, column=5, value=nombre)
    ws.cell(row=r, column=6, value=f'=IF(A{r}="","",TEXT(A{r},"YYYY-MM"))')
    upload_excel(wb)


def calcular_totales() -> dict:
    """Recalcula a mano (no depende de la cache de formulas de Excel)."""
    wb = download_excel()

    ws = wb["Movimientos"]
    ingresos = egresos = 0.0
    r = 3
    while ws.cell(row=r, column=1).value not in (None, ""):
        tipo = ws.cell(row=r, column=2).value
        monto = ws.cell(row=r, column=3).value or 0
        if tipo == "Ingreso":
            ingresos += monto
        elif tipo == "Egreso":
            egresos += monto
        r += 1

    wsd = wb["Ahorro_Detalle"]
    fondos_total = {"Vacaciones": 0.0, "Jubilación": 0.0, "Fondo de Emergencia": 0.0}
    r = 3
    while wsd.cell(row=r, column=1).value not in (None, ""):
        fondo = wsd.cell(row=r, column=2).value
        tipo = wsd.cell(row=r, column=3).value
        monto = wsd.cell(row=r, column=4).value or 0
        if fondo in fondos_total:
            fondos_total[fondo] += monto if tipo == "Aporte" else -monto
        r += 1

    return {
        "ingresos": ingresos,
        "egresos": egresos,
        "balance": ingresos - egresos,
        "fondos": fondos_total,
    }


def fmt(n: float) -> str:
    return f"${n:,.0f}".replace(",", ".")


# ---------------------------------------------------------------------------
# Handlers de Telegram
# ---------------------------------------------------------------------------
async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    chat_id = str(update.effective_chat.id)
    users = load_users()
    if chat_id in users:
        await update.message.reply_text(f"Hola {users[chat_id]}, ya te tengo registrado ✅")
        return
    buttons = [[InlineKeyboardButton(n, callback_data=f"reg:{n}")] for n in ALL_NAMES]
    await update.message.reply_text(
        "¡Hola! ¿Quién sos?", reply_markup=InlineKeyboardMarkup(buttons)
    )


async def register_callback(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    name = query.data.split(":")[1]
    chat_id = str(query.message.chat_id)
    users = load_users()
    users[chat_id] = name
    save_users(users)
    await query.answer()
    await query.edit_message_text(f"Listo, quedaste registrado como {name} ✅")


async def total_cmd(update: Update, context: ContextTypes.DEFAULT_TYPE):
    chat_id = str(update.effective_chat.id)
    users = load_users()
    nombre = users.get(chat_id)
    if nombre not in ADMIN_NAMES:
        await update.message.reply_text("Este comando es solo para Luis y Vero.")
        return
    t = calcular_totales()
    msg = (
        f"📊 *Resumen*\n"
        f"Ingresos: {fmt(t['ingresos'])}\n"
        f"Egresos: {fmt(t['egresos'])}\n"
        f"Balance: {fmt(t['balance'])}\n\n"
        f"💰 *Ahorros*\n"
        f"Fondo de Emergencia: {fmt(t['fondos']['Fondo de Emergencia'])}\n"
        f"Jubilación: {fmt(t['fondos']['Jubilación'])}\n"
        f"Vacaciones: {fmt(t['fondos']['Vacaciones'])}\n"
    )
    await update.message.reply_text(msg, parse_mode="Markdown")


async def handle_message(update: Update, context: ContextTypes.DEFAULT_TYPE):
    chat_id = str(update.effective_chat.id)
    users = load_users()
    nombre = users.get(chat_id)

    if not nombre:
        await update.message.reply_text("Antes que nada, decime quién sos con /start")
        return

    text = update.message.text or ""
    parsed = parse_message(text)
    if not parsed:
        await update.message.reply_text(
            "No pude entender el monto. Probá algo como: '50.000 comida transferencia'"
        )
        return

    es_admin = nombre in ADMIN_NAMES

    # --- Ahorro (solo admins) ---
    if parsed["fondo"] and es_admin:
        tipo = "Retiro" if parsed["is_retiro"] else "Aporte"
        registrar_ahorro(nombre, parsed["fondo"], tipo, parsed["monto"])
        await update.message.reply_text(
            f"Listo ✅ {tipo} de {fmt(parsed['monto'])} en {parsed['fondo']}"
        )
        return

    # --- Ingreso (solo admins) ---
    if parsed["is_ingreso"] and es_admin:
        registrar_movimiento(nombre, "Ingreso", parsed["monto"], parsed["motivo"], parsed["medio"])
        await update.message.reply_text(f"Listo ✅ Ingreso de {fmt(parsed['monto'])} anotado")
        return

    # --- Egreso normal (todos) ---
    registrar_movimiento(nombre, "Egreso", parsed["monto"], parsed["motivo"], parsed["medio"])
    if es_admin:
        await update.message.reply_text(
            f"Listo ✅ Gasto de {fmt(parsed['monto'])} en {parsed['motivo']} ({parsed['medio']})"
        )
    else:
        await update.message.reply_text("Listo, anotado ✅")


# ---------------------------------------------------------------------------
# Arranque en modo webhook (Telegram nos avisa a nosotros, en vez de que
# nosotros le preguntemos todo el tiempo). Esto es clave para el plan free
# de Render: un pedido HTTP entrante es lo unico que puede "despertar" al
# servicio si estaba dormido por inactividad.
# ---------------------------------------------------------------------------
def main():
    application = Application.builder().token(TELEGRAM_TOKEN).build()
    application.add_handler(CommandHandler("start", start))
    application.add_handler(CommandHandler("total", total_cmd))
    application.add_handler(CallbackQueryHandler(register_callback, pattern=r"^reg:"))
    application.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_message))

    port = int(os.environ.get("PORT", 10000))
    webhook_base = os.environ["WEBHOOK_URL"].rstrip("/")

    log.info("Bot arrancado en modo webhook...")
    application.run_webhook(
        listen="0.0.0.0",
        port=port,
        url_path=TELEGRAM_TOKEN,
        webhook_url=f"{webhook_base}/{TELEGRAM_TOKEN}",
    )


if __name__ == "__main__":
    main()
