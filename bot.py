import os
import re
import json
import logging
import unicodedata
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
    ConversationHandler,
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
PUEDE_CREDITO = ["Luis", "Vero", "Malena"]
ALL_NAMES = ["Luis", "Vero", "Malena", "Luka", "Mateo"]

FONDOS = {
    "vacaciones": "Vacaciones",
    "jubilacion": "Jubilación",
    "jubilación": "Jubilación",
    "emergencia": "Fondo de Emergencia",
}

MEDIOS = {
    "efectivo": "Efectivo",
    "mercado pago": "Mercado Pago",
    "mercadopago": "Mercado Pago",
    "mp": "Mercado Pago",
    "galicia": "Galicia",
    "santander": "Santander",
}
LUGARES = ["Efectivo", "Mercado Pago", "Galicia", "Santander"]

RETIRO_WORDS = ["retiro", "retiré", "retire", "saque", "saqué", "saco"]
INGRESO_WORDS = ["ingreso", "ingresé", "ingrese", "cobre", "cobré", "cobro"]
CREDITO_WORDS = ["credito", "crédito"]

TARJETAS = ["Visa", "American"]
BANCOS = ["Galicia", "Santander"]

# Estados de la conversacion de credito
TARJETA, BANCO, CUOTAS, MES = range(4)

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
ACCENT_MAP = {"a": "[aá]", "e": "[eé]", "i": "[ií]", "o": "[oó]", "u": "[uúü]"}


def patron_sin_acentos(palabra: str) -> str:
    return "".join(ACCENT_MAP.get(c, re.escape(c)) for c in palabra)


def quitar_acentos(s: str) -> str:
    return "".join(
        c for c in unicodedata.normalize("NFD", s) if unicodedata.category(c) != "Mn"
    )


def parse_monto(s: str) -> float:
    """Interpreta '50.000' (miles) o '2479054,80' (decimales) correctamente."""
    partes = re.split(r"[.,]", s)
    if len(partes) == 1:
        return float(partes[0])
    ultima = partes[-1]
    if len(ultima) in (1, 2):
        entero = "".join(partes[:-1])
        return float(f"{entero}.{ultima}")
    return float("".join(partes))


def parse_message(text: str) -> dict | None:
    text_low = text.lower()
    text_low_sa = quitar_acentos(text_low)

    m = re.search(r"\$?\s*(\d+(?:[.,]\d{3})*(?:[.,]\d{1,2})?)", text)
    if not m:
        return None
    try:
        monto = parse_monto(m.group(1))
    except ValueError:
        return None

    fondo = None
    for k, v in FONDOS.items():
        if k in text_low_sa:
            fondo = v
            break

    is_retiro = any(w in text_low_sa for w in RETIRO_WORDS)
    is_ingreso = any(w in text_low_sa for w in INGRESO_WORDS)
    is_credito = any(w in text_low_sa for w in CREDITO_WORDS)

    medio = "Sin especificar"  # si no dicen el lugar, no se cuenta en la diversificación
    medio_encontrado = None
    for k, v in MEDIOS.items():
        if k in text_low:
            medio = v
            medio_encontrado = k
            break

    # Motivo: el mensaje original, sacando el monto y las palabras clave usadas
    motivo = text.replace(m.group(0), "")
    palabras_a_sacar = list(FONDOS.keys()) + RETIRO_WORDS + INGRESO_WORDS + CREDITO_WORDS
    if medio_encontrado:
        palabras_a_sacar.append(medio_encontrado)
    for palabra in palabras_a_sacar:
        motivo = re.sub(rf"\b{patron_sin_acentos(palabra)}\b", "", motivo, flags=re.IGNORECASE)
    motivo = re.sub(r"\s+", " ", motivo).strip(" ,.-")
    if not motivo:
        motivo = "Sin descripción"

    return {
        "monto": monto,
        "fondo": fondo,
        "is_retiro": is_retiro,
        "is_ingreso": is_ingreso,
        "is_credito": is_credito,
        "medio": medio,
        "motivo": motivo,
    }


# ---------------------------------------------------------------------------
# Escritura en Excel
# ---------------------------------------------------------------------------
def registrar_movimiento(nombre: str, tipo: str, monto: float, motivo: str, medio: str, fecha_override: datetime = None):
    wb = download_excel()
    ws = wb["Movimientos"]
    r = first_empty_row(ws, start=3, col=1)
    fecha = fecha_override or datetime.now()
    ws.cell(row=r, column=1, value=fecha)
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


MESES_ES = {
    "enero": 1, "febrero": 2, "marzo": 3, "abril": 4, "mayo": 5, "junio": 6,
    "julio": 7, "agosto": 8, "septiembre": 9, "setiembre": 9, "octubre": 10,
    "noviembre": 11, "diciembre": 12,
}


def parse_mes_input(texto: str) -> str | None:
    """Convierte 'agosto' o 'agosto2026' en 'YYYY-MM'."""
    texto_sa = quitar_acentos(texto.strip().lower())
    m = re.match(r"([a-z]+)\s*(\d{4})?", texto_sa)
    if not m:
        return None
    nombre_mes, anio_str = m.group(1), m.group(2)
    mes_num = MESES_ES.get(nombre_mes)
    if not mes_num:
        return None
    if anio_str:
        anio = int(anio_str)
    else:
        hoy = datetime.now()
        anio = hoy.year if mes_num >= hoy.month else hoy.year + 1
    return f"{anio:04d}-{mes_num:02d}"


def calcular_totales_mes(mes_sheet: str) -> dict:
    wb = download_excel()
    ws = wb["Movimientos"]
    ingresos = egresos = 0.0
    r = 3
    while ws.cell(row=r, column=1).value not in (None, ""):
        fecha = ws.cell(row=r, column=1).value
        if fecha and fecha.strftime("%Y-%m") == mes_sheet:
            tipo = ws.cell(row=r, column=2).value
            monto = ws.cell(row=r, column=3).value or 0
            if tipo == "Ingreso":
                ingresos += monto
            elif tipo == "Egreso":
                egresos += monto
        r += 1
    return {"ingresos": ingresos, "egresos": egresos, "balance": ingresos - egresos}


def mes_sheet_name(dt: datetime) -> str:
    return dt.strftime("%Y-%m")


def calcular_gasto_persona_mes(nombre_persona: str, mes_sheet: str) -> float:
    wb = download_excel()
    ws = wb["Movimientos"]
    total = 0.0
    r = 3
    while ws.cell(row=r, column=1).value not in (None, ""):
        fecha = ws.cell(row=r, column=1).value
        tipo = ws.cell(row=r, column=2).value
        persona = ws.cell(row=r, column=6).value
        if (
            fecha and fecha.strftime("%Y-%m") == mes_sheet
            and tipo == "Egreso"
            and persona == nombre_persona
        ):
            total += ws.cell(row=r, column=3).value or 0
        r += 1
    return total


def calcular_totales() -> dict:
    """Recalcula a mano (no depende de la cache de formulas de Excel)."""
    wb = download_excel()

    ws = wb["Movimientos"]
    ingresos = egresos = 0.0
    lugares_total = {l: 0.0 for l in LUGARES}
    r = 3
    while ws.cell(row=r, column=1).value not in (None, ""):
        tipo = ws.cell(row=r, column=2).value
        monto = ws.cell(row=r, column=3).value or 0
        medio = ws.cell(row=r, column=5).value
        if tipo == "Ingreso":
            ingresos += monto
            if medio in lugares_total:
                lugares_total[medio] += monto
        elif tipo == "Egreso":
            egresos += monto
            if medio in lugares_total:
                lugares_total[medio] -= monto
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
        "lugares": lugares_total,
    }


def fmt(n: float) -> str:
    return f"${n:,.0f}".replace(",", ".")


# ---------------------------------------------------------------------------
# Handlers de Telegram
# ---------------------------------------------------------------------------
async def reset_cmd(update: Update, context: ContextTypes.DEFAULT_TYPE):
    chat_id = str(update.effective_chat.id)
    users = load_users()
    if chat_id in users:
        del users[chat_id]
        save_users(users)
    await start(update, context)


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

    texto_comando = update.message.text.split()[0][1:]  # sin la barra "/"
    sufijo = texto_comando[len("total"):].strip()

    if sufijo:
        mes_sheet = parse_mes_input(sufijo)
        if mes_sheet:
            tm = calcular_totales_mes(mes_sheet)
            msg = (
                f"📊 *{mes_sheet}*\n"
                f"Ingresos: {fmt(tm['ingresos'])}\n"
                f"Egresos: {fmt(tm['egresos'])}\n"
                f"Balance del mes: {fmt(tm['balance'])}\n"
            )
            await update.message.reply_text(msg, parse_mode="Markdown")
            return

        sufijo_norm = quitar_acentos(sufijo.lower())
        persona_encontrada = None
        for persona in ALL_NAMES:
            if quitar_acentos(persona.lower()) == sufijo_norm:
                persona_encontrada = persona
                break
        if persona_encontrada:
            mes_actual = mes_sheet_name(datetime.now())
            gasto = calcular_gasto_persona_mes(persona_encontrada, mes_actual)
            await update.message.reply_text(
                f"{persona_encontrada} gastó {fmt(gasto)} en {mes_actual}"
            )
            return

        await update.message.reply_text(
            f"No entendí '{sufijo}'. Probá '/totalagosto' (un mes) o '/total{sufijo.lower()}' con el nombre de alguien de la familia."
        )
        return

    t = calcular_totales()
    ahorrado_total = sum(t["fondos"].values())
    balance_disponible = t["balance"] - ahorrado_total
    msg = (
        f"📊 *Resumen*\n"
        f"Ingresos: {fmt(t['ingresos'])}\n"
        f"Egresos: {fmt(t['egresos'])}\n"
        f"Ahorrado (separado): {fmt(ahorrado_total)}\n"
        f"Balance disponible: {fmt(balance_disponible)}\n\n"
        f"📍 *Dónde está la plata*\n"
        f"Efectivo: {fmt(t['lugares']['Efectivo'])}\n"
        f"Mercado Pago: {fmt(t['lugares']['Mercado Pago'])}\n"
        f"Galicia: {fmt(t['lugares']['Galicia'])}\n"
        f"Santander: {fmt(t['lugares']['Santander'])}\n\n"
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
        return ConversationHandler.END

    text = update.message.text or ""
    parsed = parse_message(text)
    if not parsed:
        await update.message.reply_text(
            "No reconocí ese mensaje como un gasto/ingreso. Probá algo como: '50.000 comida transferencia'"
        )
        return ConversationHandler.END

    es_admin = nombre in ADMIN_NAMES

    # --- Credito (Luis, Vero, Malena; solo gastos) ---
    if parsed["is_credito"] and not parsed["is_ingreso"] and not parsed["fondo"] and nombre in PUEDE_CREDITO:
        context.user_data["pendiente"] = {
            "nombre": nombre,
            "monto": parsed["monto"],
            "motivo": parsed["motivo"],
        }
        buttons = [[InlineKeyboardButton(t, callback_data=f"tarjeta:{t}")] for t in TARJETAS]
        await update.message.reply_text(
            "¿Qué tarjeta usaste?", reply_markup=InlineKeyboardMarkup(buttons)
        )
        return TARJETA

    # --- Ahorro (solo admins) ---
    if parsed["fondo"] and es_admin:
        tipo = "Retiro" if parsed["is_retiro"] else "Aporte"
        registrar_ahorro(nombre, parsed["fondo"], tipo, parsed["monto"])
        await update.message.reply_text(
            f"Listo ✅ {tipo} de {fmt(parsed['monto'])} en {parsed['fondo']}"
        )
        return ConversationHandler.END

    # --- Ingreso (solo admins) ---
    if parsed["is_ingreso"] and es_admin:
        registrar_movimiento(nombre, "Ingreso", parsed["monto"], parsed["motivo"], parsed["medio"])
        await update.message.reply_text(f"Listo ✅ Ingreso de {fmt(parsed['monto'])} anotado")
        return ConversationHandler.END

    # --- Egreso normal (todos) ---
    registrar_movimiento(nombre, "Egreso", parsed["monto"], parsed["motivo"], parsed["medio"])
    if es_admin:
        await update.message.reply_text(
            f"Listo ✅ Gasto de {fmt(parsed['monto'])} en {parsed['motivo']} ({parsed['medio']})"
        )
    else:
        await update.message.reply_text("Listo, anotado ✅")
    return ConversationHandler.END


# ---------------------------------------------------------------------------
# Conversacion de credito: tarjeta -> banco -> cuotas -> mes
# ---------------------------------------------------------------------------
def sumar_meses(mes_sheet: str, n: int) -> str:
    anio, mes = map(int, mes_sheet.split("-"))
    total = (anio * 12 + (mes - 1)) + n
    return f"{total // 12:04d}-{total % 12 + 1:02d}"


async def tarjeta_callback(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    tarjeta = query.data.split(":")[1]
    context.user_data["pendiente"]["tarjeta"] = tarjeta
    await query.answer()
    buttons = [[InlineKeyboardButton(b, callback_data=f"banco:{b}")] for b in BANCOS]
    await query.edit_message_text(f"Tarjeta: {tarjeta}\n¿De qué banco?")
    await query.message.reply_text("Elegí el banco:", reply_markup=InlineKeyboardMarkup(buttons))
    return BANCO


async def banco_callback(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    banco = query.data.split(":")[1]
    context.user_data["pendiente"]["banco"] = banco
    await query.answer()
    await query.edit_message_text(f"Banco: {banco}\n¿Cuántas cuotas? (escribí el número)")
    return CUOTAS


async def cuotas_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    texto = update.message.text.strip()
    if not texto.isdigit() or int(texto) < 1:
        await update.message.reply_text("Escribí un número válido de cuotas (ej. 1, 3, 12).")
        return CUOTAS
    context.user_data["pendiente"]["cuotas"] = int(texto)
    await update.message.reply_text(
        "¿En qué mes empieza a cobrarse la primera cuota? (ej. 'agosto' o 'agosto 2026')"
    )
    return MES


async def mes_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    texto = update.message.text.strip()
    mes_sheet = parse_mes_input(texto)
    if not mes_sheet:
        await update.message.reply_text("No entendí el mes. Escribilo así: 'agosto' o 'agosto 2026'.")
        return MES

    pend = context.user_data.get("pendiente")
    if not pend:
        await update.message.reply_text("Se perdió el gasto en curso, probá de nuevo desde cero.")
        return ConversationHandler.END

    cuotas = pend["cuotas"]
    monto_total = pend["monto"]
    monto_cuota = round(monto_total / cuotas)
    diferencia = monto_total - monto_cuota * cuotas

    for i in range(cuotas):
        monto_i = monto_cuota + (diferencia if i == cuotas - 1 else 0)
        mes_i = sumar_meses(mes_sheet, i)
        anio_i, mes_num_i = map(int, mes_i.split("-"))
        fecha_i = datetime(anio_i, mes_num_i, 1)
        motivo_i = f"{pend['motivo']} ({pend['tarjeta']}, cuota {i + 1}/{cuotas})"
        registrar_movimiento(pend["nombre"], "Egreso", monto_i, motivo_i, pend["banco"], fecha_override=fecha_i)

    await update.message.reply_text(
        f"Listo ✅ {fmt(monto_total)} en {cuotas} cuota(s) de {fmt(monto_cuota)}, "
        f"empezando en {mes_sheet} ({pend['tarjeta']} {pend['banco']})"
    )
    context.user_data.pop("pendiente", None)
    return ConversationHandler.END


async def cancelar(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data.pop("pendiente", None)
    await update.message.reply_text("Cancelado.")
    return ConversationHandler.END


# ---------------------------------------------------------------------------
# Arranque en modo webhook (Telegram nos avisa a nosotros, en vez de que
# nosotros le preguntemos todo el tiempo). Esto es clave para el plan free
# de Render: un pedido HTTP entrante es lo unico que puede "despertar" al
# servicio si estaba dormido por inactividad.
# ---------------------------------------------------------------------------
def main():
    application = Application.builder().token(TELEGRAM_TOKEN).build()

    conv_handler = ConversationHandler(
        entry_points=[MessageHandler(filters.TEXT & ~filters.COMMAND, handle_message)],
        states={
            TARJETA: [CallbackQueryHandler(tarjeta_callback, pattern=r"^tarjeta:")],
            BANCO: [CallbackQueryHandler(banco_callback, pattern=r"^banco:")],
            CUOTAS: [MessageHandler(filters.TEXT & ~filters.COMMAND, cuotas_handler)],
            MES: [MessageHandler(filters.TEXT & ~filters.COMMAND, mes_handler)],
        },
        fallbacks=[CommandHandler("cancelar", cancelar)],
    )

    application.add_handler(CommandHandler("start", start))
    application.add_handler(CommandHandler("reset", reset_cmd))
    application.add_handler(MessageHandler(filters.Regex(r"^/total") & filters.COMMAND, total_cmd))
    application.add_handler(CallbackQueryHandler(register_callback, pattern=r"^reg:"))
    application.add_handler(conv_handler)

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
