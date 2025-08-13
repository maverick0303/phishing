import pandas as pd
from jinja2 import Template
import os

# 1️⃣ Cargar Excel de jefaturas
df = pd.read_excel("jefaturas_template.xlsx")

# 2️⃣ Normalizar columnas
df.columns = df.columns.str.strip().str.lower().str.replace(' ', '_')

# 3️⃣ Cargar plantilla HTML
with open("jefes.html", "r", encoding="utf-8") as f:
    template = Template(f.read())

# 4️⃣ Crear carpeta de destino si no existe
carpeta_destino = "correos_jefatura"
os.makedirs(carpeta_destino, exist_ok=True)

# 5️⃣ Generar un HTML por jefe
for _, row in df.iterrows():
    html_content = template.render(
        jefe_nombre=row['jefe_nombre'],
        jefe_email=row['jefe_email'],
        colaboradores=row['colaboradores'],
        cantidad_colaboradores=row['cantidad_colaboradores'],
        periodo_meses=row['periodo_meses'],
        curso_titulo=row['curso_titulo'],
        curso_plataforma=row['curso_plataforma'],
        curso_link=row['curso_link'],
        correo_soporte=row['correo_soporte']
    )

    # Guardar archivo HTML en la carpeta
    nombre_archivo = f"{row['jefe_nombre'].replace(' ', '_')}.html"
    ruta_completa = os.path.join(carpeta_destino, nombre_archivo)
    with open(ruta_completa, "w", encoding="utf-8") as salida:
        salida.write(html_content)

    print(f"✅ Generado {ruta_completa}")
