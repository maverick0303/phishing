import pandas as pd
from jinja2 import Template
import os

# 1. Cargar Excel
df = pd.read_excel("usuarios.xlsx")

# 2. Normalizar columnas
df.columns = df.columns.str.strip().str.lower().str.replace(' ', '_')

# 3. Cargar plantilla HTML
with open("usuarios.html", "r", encoding="utf-8") as f:
    template = Template(f.read())

# 4. Crear carpeta si no existe
carpeta_destino = "correos_usuarios"
os.makedirs(carpeta_destino, exist_ok=True)

# 5. Generar un HTML por persona
for _, row in df.iterrows():
    html_content = template.render(
        nombre=row['nombres'],
        cargo=row['cargo'],
        resultado_abril=row['primer_ejercicio_concurso-abril'],
        resultado_mayo=row['segundo_ejercicio_bad_bunny_-_mayo'],
        resultado_junio=row['tercer_ejercicio_remuneraciones_-_junio']
    )

    # Guardar archivo HTML personalizado en la carpeta
    nombre_archivo = f"{row['nombres'].replace(' ', '_')}.html"
    ruta_completa = os.path.join(carpeta_destino, nombre_archivo)
    with open(ruta_completa, "w", encoding="utf-8") as salida:
        salida.write(html_content)

    print(f"✅ Generado {ruta_completa}")
