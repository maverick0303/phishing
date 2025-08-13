import pandas as pd
from pathlib import Path
from jinja2 import Template
from html import escape

# === Config ===
RUTA_EXCEL = Path("jefaturas_template.xlsx")   # pon aquí tu archivo real
RUTA_TEMPLATE = Path("jefes.html")
OUTDIR = Path("correos_html_jefes")

# === Cargar Excel ===
df = pd.read_excel(RUTA_EXCEL)

# Normalizar nombres de columnas a minúsculas y sin espacios
df.columns = df.columns.str.strip().str.lower()

# Columnas esperadas (se intentan crear si faltan para no romper)
esperadas = [
    "jefe_nombre","jefe_email",
    "colaborador_nombre","colaborador_cargo",
    "ej1_tema","ej1_mes","ej1_resultado",
    "ej2_tema","ej2_mes","ej2_resultado",
    "ej3_tema","ej3_mes","ej3_resultado",
]
for col in esperadas:
    if col not in df.columns:
        df[col] = ""

# === Cargar plantilla Jinja ===
with open(RUTA_TEMPLATE, "r", encoding="utf-8") as f:
    template = Template(f.read())

# === Helpers ===
def safe(x):
    return escape(str(x)) if pd.notna(x) and str(x) != "nan" else ""

def encabezado(grupo, col_tema, col_mes, fallback):
    tema = next((safe(v) for v in grupo[col_tema].tolist() if pd.notna(v) and str(v).strip()), "")
    mes  = next((safe(v) for v in grupo[col_mes].tolist()  if pd.notna(v) and str(v).strip()), "")
    if tema or mes:
        return f"{tema} - {mes}".strip(" -")
    return fallback

# === Generación por jefe ===
OUTDIR.mkdir(parents=True, exist_ok=True)

for (jefe_nombre, jefe_email), g in df.groupby(["jefe_nombre","jefe_email"], dropna=False):
    g = g.copy()

    ej1_header = encabezado(g, "ej1_tema", "ej1_mes", "Primer ejercicio")
    ej2_header = encabezado(g, "ej2_tema", "ej2_mes", "Segundo ejercicio")
    ej3_header = encabezado(g, "ej3_tema", "ej3_mes", "Tercer ejercicio")

    colaboradores = []
    for _, r in g.iterrows():
        colaboradores.append({
            "nombre": safe(r.get("colaborador_nombre","")),
            "cargo":  safe(r.get("colaborador_cargo","")),
            "ej1":    safe(r.get("ej1_resultado","")),
            "ej2":    safe(r.get("ej2_resultado","")),
            "ej3":    safe(r.get("ej3_resultado","")),
        })

    html = template.render(
        jefe_nombre=safe(jefe_nombre),
        ej1_header=ej1_header,
        ej2_header=ej2_header,
        ej3_header=ej3_header,
        colaboradores=colaboradores
    )

    # Nombre de archivo por jefe
    safe_name = "".join(ch for ch in str(jefe_nombre) if ch.isalnum() or ch in (" ","-","_")).strip().replace(" ","_")
    out_file = OUTDIR / f"correo_{safe_name or 'sin_nombre'}.html"
    out_file.write_text(html, encoding="utf-8")
    print(f"✅ Generado {out_file}")
