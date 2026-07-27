# Cómo poner el bot en funcionamiento

## Paso 0: subir el Excel a Dropbox
Subí el archivo `Gastos_Familia_v3.xlsx` a la carpeta que Dropbox creó para esta app
(algo como `/Apps/GastosGomezBot/`). Tiene que llamarse exactamente igual y estar
en esa carpeta — es lo único a lo que el token de Dropbox tiene acceso.

## Paso 1: subir el código a GitHub (gratis)
Render necesita conectarse a un repositorio para poder desplegar.

1. Andá a github.com y creá una cuenta (si no tenés)
2. Click en "New repository" → nombre, por ejemplo `bot-familia` → marcalo como **Private**
3. Subí los 3 archivos de esta carpeta (`bot.py`, `requirements.txt`, este `DEPLOY.md`)
   usando el botón "Add file" → "Upload files" en la web de GitHub (no hace falta usar
   la terminal)

## Paso 2: crear el servicio en Render
1. Andá a render.com y creá una cuenta gratis (podés entrar con GitHub directamente)
2. Click en "New +" → "Web Service"
3. Conectá el repositorio `bot-familia` que acabás de crear
4. Configuración:
   - **Runtime**: Python 3
   - **Build Command**: `pip install -r requirements.txt`
   - **Start Command**: `python bot.py`
   - **Plan**: Free
5. Antes de crear el servicio, bajá hasta "Environment Variables" y agregá estas 4:

   | Variable | Valor |
   |---|---|
   | `TELEGRAM_TOKEN` | el token que te dio BotFather |
   | `DROPBOX_TOKEN` | el token que generaste en Dropbox |
   | `EXCEL_PATH` | `/Gastos_Familia_v3.xlsx` |
   | `USERS_PATH` | `/usuarios.json` |

6. Click en "Create Web Service"

Render va a instalar todo y arrancar el bot solo. La primera vez puede tardar
unos minutos. Cuando el log diga "Bot arrancado, escuchando mensajes...", ya
está funcionando.

## Paso 3: probarlo
1. Cada uno de los 5 (Luis, Vero, Malena, Luka, Mateo) le escribe `/start` al bot
   en Telegram y toca su nombre en los botones que aparecen
2. Después ya pueden mandar mensajes como "50.000 comida transferencia"
3. Luis o Vero pueden escribir `/total` en cualquier momento para ver el resumen

## Notas
- El plan Free de Render "duerme" el servicio si pasa mucho tiempo sin uso;
  el primer mensaje después de la inactividad puede tardar 20-30 segundos en
  responder mientras se despierta. Es normal.
- Si algún día cambiás el nombre del archivo Excel o su ubicación en Dropbox,
  solo hay que actualizar la variable `EXCEL_PATH` en Render, no el código.
