import os
import sqlite3
import tempfile
import subprocess
import shutil
from datetime import date

import pandas as pd
import streamlit as st
from docxtpl import DocxTemplate, RichText

st.set_page_config(page_title="APP Área Comercial Buses y Vans", layout="wide")

TEMPLATE_FILE = "plantilla_cotizacion_foton_u9.docx"
DB_FILE = "cotizaciones.db"

# =========================================================
# COTIZANTES
# =========================================================
COTIZANTES = {
    "Diego Vejar": {
        "prefijo": "DV",
        "firma_nombre": "Diego Vejar",
        "firma_cargo": "Subgerente Comercial Buses y Vans",
        "firma_correo": "dvejar@andesmotor.cl",
        "firma_telefono": "981774604",
    },
    "Sergio Silva": {
        "prefijo": "SS",
        "firma_nombre": "Sergio Silva",
        "firma_cargo": "Ejecutivo Zona Sur",
        "firma_correo": "sergio.silva@andesmotor.cl",
        "firma_telefono": "",
    },
    "Alvaro Correa": {
        "prefijo": "AC",
        "firma_nombre": "Alvaro Correa",
        "firma_cargo": "Ejecutivo Zona Norte",
        "firma_correo": "alvaro.correa@andesmotor.cl",
        "firma_telefono": "",
    },
}

# =========================================================
# UTILIDADES
# =========================================================
def fecha_larga_es(fecha):
    meses = ["enero","febrero","marzo","abril","mayo","junio",
             "julio","agosto","septiembre","octubre","noviembre","diciembre"]
    return f"Santiago, {fecha.day} de {meses[fecha.month-1]} de {fecha.year}"

def usd_fmt(valor):
    return f"USD$ {valor:,.0f}".replace(",", ".")

def limpiar(texto):
    return texto.replace(" ", "_")

def formatear_texto(texto):
    rt = RichText()
    for linea in texto.split("\n"):
        rt.add(linea)
        rt.add("\n")
    return rt

# =========================================================
# BASE DE DATOS
# =========================================================
def init_db():
    conn = sqlite3.connect(DB_FILE)
    cur = conn.cursor()

    cur.execute("""
    CREATE TABLE IF NOT EXISTS cotizaciones (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        cliente TEXT,
        cotizante TEXT,
        numero TEXT
    )
    """)
    conn.commit()
    conn.close()

def siguiente_correlativo(cotizante):
    conn = sqlite3.connect(DB_FILE)
    cur = conn.cursor()

    prefijo = COTIZANTES[cotizante]["prefijo"]

    cur.execute("SELECT numero FROM cotizaciones WHERE numero LIKE ?", (f"{prefijo}-%",))
    numeros = cur.fetchall()

    if not numeros:
        return 1

    ult = max([int(n[0].split("-")[1]) for n in numeros])
    return ult + 1

def guardar(cliente, cotizante, numero):
    conn = sqlite3.connect(DB_FILE)
    cur = conn.cursor()
    cur.execute("INSERT INTO cotizaciones (cliente, cotizante, numero) VALUES (?,?,?)",
                (cliente, cotizante, numero))
    conn.commit()
    conn.close()

# =========================================================
# PDF
# =========================================================
def convertir_pdf(docx):
    if shutil.which("soffice") is None:
        return None

    subprocess.run([
        "soffice",
        "--headless",
        "--convert-to", "pdf",
        docx,
        "--outdir", os.path.dirname(docx)
    ])

    return docx.replace(".docx", ".pdf")

# =========================================================
# DOCX
# =========================================================
def generar_doc(contexto, nombre):
    doc = DocxTemplate(TEMPLATE_FILE)
    doc.render(contexto)

    ruta = f"{nombre}.docx"
    doc.save(ruta)
    return ruta

# =========================================================
# APP
# =========================================================
init_db()

st.title("APP Área Comercial Buses y Vans")

tab1, tab2, tab3 = st.tabs([
    "Nueva cotización",
    "Historial",
    "Eficiencia energética"
])

# =========================================================
# TAB 1
# =========================================================
with tab1:
    st.subheader("Cotización FOTON U9")

    cliente = st.text_input("Cliente")
    cotizante = st.selectbox("Cotizante", list(COTIZANTES.keys()))
    precio = st.number_input("Precio USD", value=130491.0)

    texto = st.text_area("Texto mantenimiento", value="""I. Mantención incluida
II. Bono repuestos
III. Telemetría incluida""")

    if st.button("Generar cotización"):

        correlativo = siguiente_correlativo(cotizante)
        prefijo = COTIZANTES[cotizante]["prefijo"]
        numero = f"{prefijo}-{correlativo:02d}"

        contexto = {
            "cliente": cliente,
            "numero_cotizacion": numero,
            "fecha_larga": fecha_larga_es(date.today()),
            "precio_unitario": usd_fmt(precio),
            "texto_mantto": formatear_texto(texto),
            "firma_nombre": COTIZANTES[cotizante]["firma_nombre"],
            "firma_cargo": COTIZANTES[cotizante]["firma_cargo"],
        }

        nombre = f"Propuesta_{limpiar(cliente)}"

        archivo = generar_doc(contexto, nombre)
        guardar(cliente, cotizante, numero)

        st.success(f"Generada {numero}")

        with open(archivo, "rb") as f:
            st.download_button("Descargar Word", f, file_name=archivo)

        pdf = convertir_pdf(archivo)

        if pdf:
            with open(pdf, "rb") as f:
                st.download_button("Descargar PDF", f, file_name=pdf)

# =========================================================
# TAB 2
# =========================================================
with tab2:
    conn = sqlite3.connect(DB_FILE)
    df = pd.read_sql("SELECT * FROM cotizaciones", conn)
    st.dataframe(df)

# =========================================================
# TAB 3
# =========================================================
with tab3:
    st.subheader("Eficiencia energética")

    archivo = st.file_uploader("Subir Excel")

    if archivo:
        df = pd.read_excel(archivo, engine="openpyxl")

        st.dataframe(df.head())

        col_kwh = st.selectbox("Col energía", df.columns)
        col_km = st.selectbox("Col distancia", df.columns)

        if st.button("Calcular"):

            df["kwh_km"] = df[col_kwh] / df[col_km]

            st.metric("Consumo promedio", f"{df['kwh_km'].mean():.2f} kWh/km")

            st.line_chart(df["kwh_km"])pip install streamlit pandas docxtpl python-docx openpyxl
sudo apt-get install -y libreoffice