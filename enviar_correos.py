import os
import smtplib
import pandas as pd
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

# CONFIGURA TU CUENTA GMAIL AQUI 👇
GMAIL_USER = "pruebasunisimple@gmail.com"  # Tu cuenta Gmail
GMAIL_PASSWORD = "ukpu xgyo wpuc grvo"  # Contraseña de aplicación generada

# Cargar la lista de jefes desde el Excel
df = pd.read_excel("jefaturas_template.xlsx")
df.columns = df.columns.str.strip().str.lower().str.replace(" ", "_")

# Conexión segura con Gmail usando SSL
server = smtplib.SMTP_SSL("smtp.gmail.com", 465)
server.login(GMAIL_USER, GMAIL_PASSWORD)

# Enviar un correo por cada HTML generado
for _, fila in df.iterrows():
    jefe = fila["jefe_nombre"]
    
    # SOLO PARA PRUEBAS: envía todos a tu correo personal
    correo_destino = "mariavyeguezp@gmail.com"  # 👈 Aquí va tu correo de prueba

    # Buscar el archivo HTML del jefe correspondiente
    nombre_archivo = f"{jefe.replace(' ', '_')}.html"
    ruta_html = os.path.join("correos_jefatura", nombre_archivo)

    if not os.path.exists(ruta_html):
        print(f"❌ No se encontró HTML para: {jefe}")
        continue

    with open(ruta_html, "r", encoding="utf-8") as f:
        contenido = f.read()

    # Crear el correo
    msg = MIMEMultipart("alternative")
    msg["Subject"] = "Medida Preventiva Jefaturas – Curso de Phishing"
    msg["From"] = GMAIL_USER
    msg["To"] = correo_destino
    msg.attach(MIMEText(contenido, "html"))

    try:
        server.sendmail(GMAIL_USER, correo_destino, msg.as_string())
        print(f"✅ Enviado a: {correo_destino} (contenido de: {jefe})")
    except Exception as e:
        print(f"❌ Error con {correo_destino}: {e}")

# Cerrar servidor SMTP
server.quit()
