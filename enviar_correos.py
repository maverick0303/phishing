import os
import smtplib
import pandas as pd
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage

# CONFIGURA TU CUENTA GMAIL AQUI 👇
GMAIL_USER = "pruebasunisimple@gmail.com"
GMAIL_PASSWORD = "ukpu xgyo wpuc grvo"

# ================================
# 🔴 ENVÍO A JEFATURAS
# ================================

print("📤 Enviando correos a jefaturas...")

archivo_jefes = "jefaturas_template.xlsx"
carpeta_jefes = "correos_jefatura"
registro_jefes = "correos_enviados_jefes.csv"

df_jefes = pd.read_excel(archivo_jefes)
df_jefes.columns = df_jefes.columns.str.strip().str.lower().str.replace(" ", "_")

if os.path.exists(registro_jefes):
    enviados_jefes = pd.read_csv(registro_jefes)["correo"].tolist()
else:
    enviados_jefes = []

server = smtplib.SMTP_SSL("smtp.gmail.com", 465)
server.login(GMAIL_USER, GMAIL_PASSWORD)

jefes_nuevos_enviados = []

for _, fila in df_jefes.iterrows():
    jefe = fila["jefe_nombre"]
    correo_destino = fila["jefe_email"]

    if correo_destino in enviados_jefes:
        print(f"⏭ Ya enviado a: {correo_destino} (jefatura)")
        continue

    nombre_archivo = f"{jefe.replace(' ', '_')}.html"
    ruta_html = os.path.join(carpeta_jefes, nombre_archivo)

    if not os.path.exists(ruta_html):
        print(f"❌ No se encontró HTML para: {jefe}")
        continue

    with open(ruta_html, "r", encoding="utf-8") as f:
        contenido = f.read()

    # 🧩 Inyectamos imagen al final del contenido
    contenido += '<br><img src="cid:claroimg" style="width:100%; max-width:600px; margin-top:20px;">'

    msg = MIMEMultipart("related")
    msg["Subject"] = "Medida Preventiva Jefaturas – Curso de Phishing"
    msg["From"] = GMAIL_USER
    msg["To"] = correo_destino

    alt = MIMEMultipart("alternative")
    alt.attach(MIMEText(contenido, "html"))
    msg.attach(alt)

    if os.path.exists("claro.png"):
        with open("claro.png", "rb") as f:
            img = MIMEImage(f.read())
            img.add_header("Content-ID", "<claroimg>")
            img.add_header("Content-Disposition", "inline", filename="claro.png")
            msg.attach(img)
    else:
        print("⚠️ No se encontró claro.png")

    try:
        server.sendmail(GMAIL_USER, correo_destino, msg.as_string())
        print(f"✅ Enviado a jefatura: {correo_destino}")
        jefes_nuevos_enviados.append(correo_destino)
    except Exception as e:
        print(f"❌ Error enviando a jefatura {correo_destino}: {e}")

# Guardar jefes enviados
if jefes_nuevos_enviados:
    df_nuevos = pd.DataFrame({"correo": jefes_nuevos_enviados})
    if os.path.exists(registro_jefes):
        df_nuevos.to_csv(registro_jefes, mode="a", header=False, index=False)
    else:
        df_nuevos.to_csv(registro_jefes, index=False)

# ================================
# 🟢 ENVÍO A USUARIOS
# ================================

print("\n📤 Enviando correos a usuarios...")

archivo_usuarios = "usuarios.xlsx"
carpeta_usuarios = "correos_usuarios"
registro_usuarios = "correos_enviados_usuarios.csv"

df_usuarios = pd.read_excel(archivo_usuarios)
df_usuarios.columns = df_usuarios.columns.str.strip().str.lower().str.replace(" ", "_")

if os.path.exists(registro_usuarios):
    enviados_usuarios = pd.read_csv(registro_usuarios)["correo"].tolist()
else:
    enviados_usuarios = []

usuarios_nuevos_enviados = []

for _, fila in df_usuarios.iterrows():
    nombre = fila["nombres"]
    correo_destino = fila["correo"]

    if correo_destino in enviados_usuarios:
        print(f"⏭ Ya enviado a: {correo_destino} (usuario)")
        continue

    nombre_archivo = f"{nombre.replace(' ', '_')}.html"
    ruta_html = os.path.join(carpeta_usuarios, nombre_archivo)

    if not os.path.exists(ruta_html):
        print(f"❌ No se encontró HTML para: {nombre}")
        continue

    with open(ruta_html, "r", encoding="utf-8") as f:
        contenido = f.read()

    # Inyectamos imagen al final
    contenido += '<br><img src="cid:claroimg" style="width:100%; max-width:600px; margin-top:20px;">'

    msg = MIMEMultipart("related")
    msg["Subject"] = "Resultado Simulación de Phishing – ClaroVTR"
    msg["From"] = GMAIL_USER
    msg["To"] = correo_destino

    alt = MIMEMultipart("alternative")
    alt.attach(MIMEText(contenido, "html"))
    msg.attach(alt)

    if os.path.exists("claro.png"):
        with open("claro.png", "rb") as f:
            img = MIMEImage(f.read())
            img.add_header("Content-ID", "<claroimg>")
            img.add_header("Content-Disposition", "inline", filename="claro.png")
            msg.attach(img)

    try:
        server.sendmail(GMAIL_USER, correo_destino, msg.as_string())
        print(f"✅ Enviado a usuario: {correo_destino}")
        usuarios_nuevos_enviados.append(correo_destino)
    except Exception as e:
        print(f"❌ Error enviando a usuario {correo_destino}: {e}")

server.quit()

# Guardar usuarios enviados
if usuarios_nuevos_enviados:
    df_nuevos = pd.DataFrame({"correo": usuarios_nuevos_enviados})
    if os.path.exists(registro_usuarios):
        df_nuevos.to_csv(registro_usuarios, mode="a", header=False, index=False)
    else:
        df_nuevos.to_csv(registro_usuarios, index=False)

print("\n✅ Correos enviados correctamente.")
