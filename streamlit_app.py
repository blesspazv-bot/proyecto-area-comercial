import os
import sqlite3
import tempfile
from datetime import date

import pandas as pd
import streamlit as st
from docxtpl import DocxTemplate

st.set_page_config(page_title="APP Área Comercial Buses y Vans", layout="wide")

TEMPLATE_FILE = "plantilla_cotizacion_foton_u9.docx"
DB_FILE = "cotizaciones.db"

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
        "firma_cargo": "Ejecutivo de ventas Zona Sur Buses y Vans",
        "firma_correo": "sergio.silva@andesmotor.cl",
        "firma_telefono": "",
    },
    "Alvaro Correa": {
        "prefijo": "AC",
        "firma_nombre": "Alvaro Correa",
        "firma_cargo": "Ejecutivo de ventas Zona Norte Buses y Vans",
        "firma_correo": "alvaro.correa@andesmotor.cl",
        "firma_telefono": "",
    },
    "Fabian Orellana": {
        "prefijo": "FO",
        "firma_nombre": "Fabian Orellana",
        "firma_cargo": "Product Manager Senior Zona Buses y Vans",
        "firma_correo": "forellana@andesmotor.cl",
        "firma_telefono": "989356449",
    },
    "Rodrigo Sepulveda": {
        "prefijo": "RS",
        "firma_nombre": "Rodrigo Sepulveda",
        "firma_cargo": "Gerente de Buses y Vans",
        "firma_correo": "rsepulveda@andesmotor.cl",
        "firma_telefono": "979783254",
    },
}


# =========================================================
# BASE DE DATOS
# =========================================================
def get_conn():
    return sqlite3.connect(DB_FILE, check_same_thread=False)


def init_db():
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("""
        CREATE TABLE IF NOT EXISTS cotizaciones (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            fecha TEXT NOT NULL,
            cliente TEXT NOT NULL,
            cotizante TEXT NOT NULL,
            prefijo TEXT NOT NULL,
            correlativo INTEGER NOT NULL,
            numero_cotizacion TEXT NOT NULL UNIQUE,
            cantidad_unidades INTEGER NOT NULL,
            precio_unitario REAL NOT NULL,
            total_negocio REAL NOT NULL,
            contrato_mantto TEXT,
            texto_mantto TEXT,
            creado_en TEXT NOT NULL DEFAULT (datetime('now'))
        )
    """)
    conn.commit()
    conn.close()


def siguiente_correlativo(cotizante):
    prefijo = COTIZANTES[cotizante]["prefijo"]
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("""
        SELECT COALESCE(MAX(correlativo), 0) + 1
        FROM cotizaciones
        WHERE prefijo = ?
    """, (prefijo,))
    valor = cur.fetchone()[0]
    conn.close()
    return valor


def guardar_cotizacion(data):
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("""
        INSERT INTO cotizaciones (
            fecha, cliente, cotizante, prefijo, correlativo, numero_cotizacion,
            cantidad_unidades, precio_unitario, total_negocio,
            contrato_mantto, texto_mantto, creado_en
        )
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, datetime('now'))
    """, (
        data["fecha_iso"],
        data["cliente"],
        data["cotizante"],
        data["prefijo"],
        data["correlativo"],
        data["numero_cotizacion"],
        data["cantidad_unidades"],
        data["precio_unitario_raw"],
        data["total_negocio_raw"],
        data["contrato_mantto"],
        data["texto_mantto"],
    ))
    conn.commit()
    conn.close()


def cargar_historial():
    conn = get_conn()
    df = pd.read_sql_query("""
        SELECT
            fecha,
            cliente,
            cotizante,
            numero_cotizacion,
            cantidad_unidades,
            precio_unitario,
            total_negocio,
            creado_en
        FROM cotizaciones
        ORDER BY id DESC
    """, conn)
    conn.close()
    return df


# =========================================================
# UTILIDADES
# =========================================================
def fecha_larga_es(fecha):
    dias = ["lunes", "martes", "miércoles", "jueves", "viernes", "sábado", "domingo"]
    meses = [
        "enero", "febrero", "marzo", "abril", "mayo", "junio",
        "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre"
    ]
    return f"Santiago, {dias[fecha.weekday()]} {fecha.day} de {meses[fecha.month - 1]} de {fecha.year}"


def usd_fmt(valor):
    return f"USD {valor:,.0f}".replace(",", ".")


def generar_docx(contexto):
    if not os.path.exists(TEMPLATE_FILE):
        raise FileNotFoundError(
            f"No se encontró la plantilla '{TEMPLATE_FILE}'. "
            f"Debes dejarla en la misma carpeta del proyecto."
        )

    doc = DocxTemplate(TEMPLATE_FILE)
    doc.render(contexto)

    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".docx")
    doc.save(tmp.name)
    tmp.close()
    return tmp.name


# =========================================================
# APP
# =========================================================
init_db()

st.title("APP Área Comercial Buses y Vans")
st.subheader("Generador de Cotización FOTON U9")

tab1, tab2 = st.tabs(["Nueva cotización", "Historial"])

with tab1:
    col1, col2 = st.columns(2)

    with col1:
        fecha = st.date_input("Fecha", value=date.today())
        cliente = st.text_input("Cliente", value="")
        cotizante = st.selectbox("Cotizante", list(COTIZANTES.keys()))

    with col2:
        cantidad_unidades = st.number_input("Cantidad de unidades", min_value=1, value=1, step=1)
        precio_unitario = st.number_input("Precio unitario USD", min_value=0.0, value=130491.0, step=1000.0)
        contrato_mantto = st.text_input("Contrato mantto", value="48 meses")
        capacidad_bateria = st.selectbox("Capacidad nominal batería",["231,8 kWh", "255 kWh"]) 

    texto_mantto = st.text_area(
        "Bloque beneficios / mantenimiento",
        value=(
            "I. La oferta incluye 48 meses de mantenimiento preventivo y correctivo "
            "(Correctivo de desgaste Disco y Pastilla de frenos) sin costo para el cliente, "
            "con el fin de entregar conocimientos técnicos y aprendizaje continuo del mantenimiento "
            "para este tipo de vehículos, durante este periodo.\n\n"
            "II. Adicionalmente se entregarán USD 1.500.- por bus, para la compra de repuestos de desgaste "
            "a elección del cliente (este ítem deberá ser utilizado dentro de los primeros 12 meses realizada la entrega de flota).\n\n"
            "III. La oferta incluye 8 años de telemetría sin costo para el cliente."
        ),
        height=220
    )

    if cotizante:
        prox = siguiente_correlativo(cotizante)
        pref = COTIZANTES[cotizante]["prefijo"]
        st.info(f"Próximo correlativo automático: {pref}-{prox:02d}")

    if st.button("Generar cotización", use_container_width=True):
        if not cliente.strip():
            st.error("Debes ingresar el nombre del cliente.")
        else:
            datos_firma = COTIZANTES[cotizante]
            correlativo = siguiente_correlativo(cotizante)
            numero_cotizacion = f'{datos_firma["prefijo"]}-{correlativo:02d}'
            total_negocio = precio_unitario * cantidad_unidades

            contexto = {
                "fecha_larga": fecha_larga_es(fecha),
                "numero_cotizacion": numero_cotizacion,
                "cliente": cliente.strip(),
                "cantidad_unidades": cantidad_unidades,
                "precio_unitario": usd_fmt(precio_unitario),
                "texto_mantto": texto_mantto,
                "capacidad_bateria": capacidad_bateria,
                "firma_nombre": datos_firma["firma_nombre"],
                "firma_cargo": datos_firma["firma_cargo"],
                "firma_correo": datos_firma["firma_correo"],
                "firma_telefono": datos_firma["firma_telefono"],
            }

            registro = {
                "fecha_iso": fecha.isoformat(),
                "cliente": cliente.strip(),
                "cotizante": cotizante,
                "prefijo": datos_firma["prefijo"],
                "correlativo": correlativo,
                "numero_cotizacion": numero_cotizacion,
                "cantidad_unidades": int(cantidad_unidades),
                "precio_unitario_raw": float(precio_unitario),
                "total_negocio_raw": float(total_negocio),
                "contrato_mantto": contrato_mantto,
                "texto_mantto": texto_mantto,
            }

            try:
                archivo_path = generar_docx(contexto)
                guardar_cotizacion(registro)

                with open(archivo_path, "rb") as f:
                    contenido = f.read()

                st.success(f"Cotización {numero_cotizacion} generada correctamente.")
                st.write(f"**Cliente:** {cliente.strip()}")
                st.write(f"**Cotizante:** {cotizante}")
                st.write(f"**Precio unitario:** {usd_fmt(precio_unitario)}")
                st.write(f"**Total negocio:** {usd_fmt(total_negocio)}")

                st.download_button(
                    "Descargar cotización Word",
                    data=contenido,
                    file_name=f"Cotizacion_{numero_cotizacion}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )

            except Exception as e:
                st.error(f"Error al generar la cotización: {e}")

with tab2:
    st.subheader("Historial de cotizaciones")

    try:
        df = cargar_historial()
        if df.empty:
            st.info("Aún no hay cotizaciones registradas.")
        else:
            df["precio_unitario"] = df["precio_unitario"].apply(usd_fmt)
            df["total_negocio"] = df["total_negocio"].apply(usd_fmt)
            st.dataframe(df, use_container_width=True)
    except Exception as e:
        st.error(f"No fue posible cargar el historial: {e}")