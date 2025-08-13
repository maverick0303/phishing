import pandas as pd
from jinja2 import Template
import os

# 1. Cargar archivos Excel
df_jefes = pd.read_excel("jefaturas_template.xlsx")
df_usuarios = pd.read_excel("usuarios.xlsx")

# 2. Normalizar nombres de columnas
df_jefes.columns = df_jefes.columns.str.strip().str.lower().str.replace(' ', '_')
df_usuarios.columns = df_usuarios.columns.str.strip().str.lower().str.replace(' ', '_')

# 3. Cargar plantilla HTML
with open("jefes.html", "r", encoding="utf-8") as f:
    template = Template(f.read())

# 4. Crear carpeta de salida
carpeta_destino = "correos_jefatura"
os.makedirs(carpeta_destino, exist_ok=True)

# 5. Hacer merge entre jefaturas y usuarios por ID
df_merged = pd.merge(df_jefes, df_usuarios, how="inner", left_on="id-j", right_on="id")

# 6. Agrupar por jefatura
from collections import defaultdict

grupos = defaultdict(list)
for _, row in df_merged.iterrows():
    grupos[row["id-j"]].append(row)

# 7. Generar HTML por jefe
for id_jefe, filas in grupos.items():
    jefe_info = filas[0]
    jefe_nombre = jefe_info["jefe_nombre"]

    colaboradores = []
    for r in filas:
        colaboradores.append({
            "nombre": r["nombres"],
            "cargo": r["cargo"],
            "ej1": r["primer_ejercicio_concurso-abril"],
            "ej2": r["segundo_ejercicio_bad_bunny_-_mayo"],
            "ej3": r["tercer_ejercicio_remuneraciones_-_junio"],
        })

    html_renderizado = template.render(
        jefe_nombre=jefe_nombre,
        colaboradores=colaboradores
    )

    nombre_archivo = f"{jefe_nombre.replace(' ', '_')}.html"
    ruta_completa = os.path.join(carpeta_destino, nombre_archivo)

    with open(ruta_completa, "w", encoding="utf-8") as salida:
        salida.write(html_renderizado)

    print(f"✅ Generado: {ruta_completa}")
