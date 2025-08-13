import pandas as pd
from jinja2 import Template

# 1. Cargar Excel
df = pd.read_excel("usuarios.xlsx")

# 2. Normalizar columnas
df.columns = df.columns.str.strip().str.lower()

# 3. Cargar plantilla HTML
with open("usuarios.html", "r", encoding="utf-8") as f:
    template = Template(f.read())

# 4. Generar un HTML por persona
for _, row in df.iterrows():
    html_content = template.render(
        nombre=row['nombres'],
        cargo=row['cargo'],
        resultado_abril=row['primer ejercicio concurso-abril'],
        resultado_mayo=row['segundo ejercicio bad bunny - mayo'],
        resultado_junio=row['tercer ejercicio remuneraciones - junio']
    )

    # Guardar archivo HTML personalizado
    nombre_archivo = f"{row['nombres'].replace(' ', '_')}.html"
    with open(nombre_archivo, "w", encoding="utf-8") as salida:
        salida.write(html_content)

    print(f"✅ Generado {nombre_archivo}")
