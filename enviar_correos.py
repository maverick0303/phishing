import os
import smtplib
import pandas as pd
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

# CONFIGURA TU CUENTA GMAIL AQUI 👇
GMAIL_USER = "pruebasunisimple@gmail.com"
GMAIL_PASSWORD = "ukpu xgyo wpuc grvo"

# Ruta al archivo de registro
registro_path = "correos_enviados.csv"

# Cargar jefes desde Excel
df = pd.read_excel("jefaturas_template.xlsx")
df.columns = df.columns.str.strip().str.lower().str.replace(" ", "_")

# Cargar lista de correos ya enviados (si existe)
if os.path.exists(registro_path):
    enviados = pd.read_csv(registro_path)["correo"].tolist()
else:
    enviados = []

# Conexión segura
server = smtplib.SMTP_SSL("smtp.gmail.com", 465)
server.login(GMAIL_USER, GMAIL_PASSWORD)

# Lista para actualizar después
correos_nuevos_enviados = []

# Enviar uno por uno
for _, fila in df.iterrows():
    jefe = fila["jefe_nombre"]
    correo_destino = fila["jefe_email"]

    if correo_destino in enviados:
        print(f"⏭ Ya fue enviado a: {correo_destino}, se omite.")
        continue

    nombre_archivo = f"{jefe.replace(' ', '_')}.html"
    ruta_html = os.path.join("correos_jefatura", nombre_archivo)

    if not os.path.exists(ruta_html):
        print(f"❌ No se encontró HTML para: {jefe}")
        continue

    with open(ruta_html, "r", encoding="utf-8") as f:
        contenido = f.read()

    msg = MIMEMultipart("alternative")
    msg["Subject"] = "Medida Preventiva Jefaturas – Curso de Phishing"
    msg["From"] = GMAIL_USER
    msg["To"] = correo_destino
    msg.attach(MIMEText(contenido, "html"))

    try:
        server.sendmail(GMAIL_USER, correo_destino, msg.as_string())
        print(f"✅ Enviado a: {correo_destino}")
        correos_nuevos_enviados.append(correo_destino)
    except Exception as e:
        print(f"❌ Error con {correo_destino}: {e}")

server.quit()

# Guardar los nuevos correos enviados al archivo CSV
if correos_nuevos_enviados:
    df_nuevos = pd.DataFrame({"correo": correos_nuevos_enviados})
    if os.path.exists(registro_path):
        df_nuevos.to_csv(registro_path, mode="a", header=False, index=False)
    else:
        df_nuevos.to_csv(registro_path, index=False)
