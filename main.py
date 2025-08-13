import pandas as pd
from jinja2 import Environment, FileSystemLoader
import os

# Archivos
EXCEL_FILE = "usuarios.xlsx"        # Tu archivo Excel
TEMPLATE_FILE = "usuarios.html"  # Plantilla HTML
OUTPUT_DIR = "correos_html"

# Crear carpeta de salida si no existe
os.makedirs(OUTPUT_DIR, exist_ok=True)

# Cargar datos
df = pd.read_excel(EXCEL_FILE)

# Configurar Jinja2
env = Environment(loader=FileSystemLoader('.'))
template = env.get_template(TEMPLATE_FILE)

# Generar HTML para cada fila
for _, row in df.iterrows():
    html_content = template.render(
        nombre=row['nombres'],
        cargo=row['cargo'],
        resultado_abril=row['primer ejercicio concurso-abril'],
        resultado_mayo=row['segundo ejercicio bad bunny - mayo'],
        resultado_junio=row['tercer ejercicio remuneraciones - junio']
    )
    
    # Guardar archivo HTML
    output_path = os.path.join(OUTPUT_DIR, f"{row['nombres']}.html")
    with open(output_path, "w", encoding="utf-8") as f:
        f.write(html_content)

print(f"Archivos generados en '{OUTPUT_DIR}'")
