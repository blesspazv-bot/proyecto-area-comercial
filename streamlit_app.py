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
        "firma_cargo": "Product Manager Senior Buses y Vans",
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
# UTILIDADES
# =========================================================
def fecha_larga_es(fecha):
    dias = ["lunes", "martes", "miércoles", "jueves", "viernes", "sábado", "domingo"]
    meses = [
        "enero", "febrero", "marzo", "abril", "mayo", "junio",
        "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre"
    ]
    return f"Santiago, {dias[fecha.weekday()]} {fecha.day} de {meses[fecha.month - 1]} de {fecha.year}"


def fecha_corta(fecha):
    return fecha.strftime("%Y%m%d")


def usd_fmt(valor):
    return f"USD$ {valor:,.0f}".replace(",", ".")


def clp_fmt(valor):
    return f"$ {valor:,.0f}".replace(",", ".")


def limpiar_nombre_archivo(texto):
    texto = str(texto).strip().replace(" ", "_")
    reemplazos = {
        "á": "a", "é": "e", "í": "i", "ó": "o", "ú": "u",
        "Á": "A", "É": "E", "Í": "I", "Ó": "O", "Ú": "U",
        "ñ": "n", "Ñ": "N",
        ",": "", ".": "", "/": "-", "\\": "-",
        ":": "-", ";": "", "(": "", ")": "", "[": "", "]": "",
    }
    for a, b in reemplazos.items():
        texto = texto.replace(a, b)
    return texto


def formatear_texto_mantto(texto):
    rt = RichText()
    bloques = [b.strip() for b in texto.split("\n") if b.strip()]
    for i, bloque in enumerate(bloques):
        rt.add(bloque)
        if i < len(bloques) - 1:
            rt.add("\n\n")
    return rt


def detectar_motor_excel(uploaded_file):
    nombre = uploaded_file.name.lower()
    if nombre.endswith(".xlsx"):
        return "openpyxl"
    if nombre.endswith(".xls"):
        return None
    return "openpyxl"


def leer_excel_seguro(uploaded_file):
    engine = detectar_motor_excel(uploaded_file)
    if engine:
        return pd.read_excel(uploaded_file, engine=engine)
    return pd.read_excel(uploaded_file)


def semaforo_kwh(valor, objetivo):
    if pd.isna(valor):
        return "Sin dato"
    if valor <= objetivo:
        return "🟢 Dentro objetivo"
    if valor <= objetivo * 1.10:
        return "🟡 Leve desvío"
    return "🔴 Sobre objetivo"

# =========================================================
# BASE DE DATOS
# =========================================================
def get_conn():
    return sqlite3.connect(DB_FILE, check_same_thread=False)


def table_exists(conn, table_name):
    cur = conn.cursor()
    cur.execute("""
        SELECT name
        FROM sqlite_master
        WHERE type='table' AND name=?
    """, (table_name,))
    return cur.fetchone() is not None


def get_columns(conn, table_name):
    cur = conn.cursor()
    cur.execute(f"PRAGMA table_info({table_name})")
    return [row[1] for row in cur.fetchall()]


def crear_tabla_cotizaciones(conn):
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
            capacidad_bateria TEXT,
            creado_en TEXT NOT NULL DEFAULT (datetime('now'))
        )
    """)
    conn.commit()


def migrar_base_si_corresponde():
    conn = get_conn()

    if not table_exists(conn, "cotizaciones"):
        crear_tabla_cotizaciones(conn)
        conn.close()
        return

    columnas = get_columns(conn, "cotizaciones")

    # Si ya está la estructura buena, solo agrega faltantes menores
    if "numero_cotizacion" in columnas:
        cur = conn.cursor()
        faltantes = {
            "contrato_mantto": "TEXT",
            "texto_mantto": "TEXT",
            "capacidad_bateria": "TEXT",
            "creado_en": "TEXT DEFAULT (datetime('now'))",
            "fecha": "TEXT",
            "prefijo": "TEXT",
            "correlativo": "INTEGER",
            "cantidad_unidades": "INTEGER DEFAULT 1",
            "precio_unitario": "REAL DEFAULT 0",
            "total_negocio": "REAL DEFAULT 0"
        }
        for col, tipo in faltantes.items():
            if col not in columnas:
                cur.execute(f"ALTER TABLE cotizaciones ADD COLUMN {col} {tipo}")
        conn.commit()
        conn.close()
        return

    # Si existe una tabla vieja con columna "numero", la migramos
    if "numero" in columnas:
        cur = conn.cursor()

        cur.execute("""
            ALTER TABLE cotizaciones RENAME TO cotizaciones_old
        """)
        conn.commit()

        crear_tabla_cotizaciones(conn)

        columnas_old = get_columns(conn, "cotizaciones_old")

        fecha_expr = "fecha" if "fecha" in columnas_old else "''"
        cliente_expr = "cliente" if "cliente" in columnas_old else "''"
        cotizante_expr = "cotizante" if "cotizante" in columnas_old else "''"
        numero_expr = "numero" if "numero" in columnas_old else "''"

        cur.execute(f"""
            INSERT INTO cotizaciones (
                fecha, cliente, cotizante, prefijo, correlativo, numero_cotizacion,
                cantidad_unidades, precio_unitario, total_negocio,
                contrato_mantto, texto_mantto, capacidad_bateria, creado_en
            )
            SELECT
                {fecha_expr},
                {cliente_expr},
                {cotizante_expr},
                '',
                0,
                {numero_expr},
                1,
                0,
                0,
                '',
                '',
                '',
                datetime('now')
            FROM cotizaciones_old
        """)
        conn.commit()
        conn.close()
        return

    # Si la tabla es irreconocible, se recrea limpia
    cur = conn.cursor()
    cur.execute("DROP TABLE IF EXISTS cotizaciones")
    conn.commit()
    crear_tabla_cotizaciones(conn)
    conn.close()


def init_db():
    migrar_base_si_corresponde()


def siguiente_correlativo(cotizante):
    prefijo = COTIZANTES[cotizante]["prefijo"]
    conn = get_conn()
    cur = conn.cursor()

    cur.execute("""
        SELECT MAX(correlativo)
        FROM cotizaciones
        WHERE prefijo = ?
    """, (prefijo,))
    row = cur.fetchone()
    conn.close()

    ultimo = row[0] if row and row[0] is not None else 0
    return ultimo + 1


def guardar_cotizacion(data):
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("""
        INSERT INTO cotizaciones (
            fecha, cliente, cotizante, prefijo, correlativo, numero_cotizacion,
            cantidad_unidades, precio_unitario, total_negocio,
            contrato_mantto, texto_mantto, capacidad_bateria, creado_en
        )
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, datetime('now'))
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
        data["capacidad_bateria"],
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
            COALESCE(capacidad_bateria, '') AS capacidad_bateria,
            COALESCE(creado_en, '') AS creado_en
        FROM cotizaciones
        ORDER BY id DESC
    """, conn)
    conn.close()
    return df

# =========================================================
# DOCX / PDF
# =========================================================
def generar_docx(contexto, nombre_salida):
    if not os.path.exists(TEMPLATE_FILE):
        raise FileNotFoundError(
            f"No se encontró la plantilla '{TEMPLATE_FILE}'. "
            f"Debes dejarla en la misma carpeta del proyecto."
        )

    doc = DocxTemplate(TEMPLATE_FILE)
    doc.render(contexto)

    tmp_dir = tempfile.mkdtemp()
    salida = os.path.join(tmp_dir, f"{nombre_salida}.docx")
    doc.save(salida)
    return salida


def convertir_docx_a_pdf(docx_path):
    posibles = [
        shutil.which("soffice"),
        "/usr/bin/soffice",
        "/usr/local/bin/soffice",
    ]
    soffice_path = next((p for p in posibles if p and os.path.exists(p)), None)

    if soffice_path is None:
        raise RuntimeError("LibreOffice no está instalado en el entorno.")

    output_dir = os.path.dirname(docx_path)

    comando = [
        soffice_path,
        "--headless",
        "--convert-to", "pdf",
        "--outdir", output_dir,
        docx_path
    ]

    resultado = subprocess.run(comando, capture_output=True, text=True)

    if resultado.returncode != 0:
        raise RuntimeError(
            f"Error al convertir a PDF.\nSTDOUT: {resultado.stdout}\nSTDERR: {resultado.stderr}"
        )

    pdf_path = os.path.splitext(docx_path)[0] + ".pdf"
    if not os.path.exists(pdf_path):
        raise RuntimeError("No se generó el archivo PDF.")

    return pdf_path

# =========================================================
# EXCEL RESULTADOS
# =========================================================
def exportar_resultados_excel(df_detalle, df_resumen, df_ruta=None):
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
    with pd.ExcelWriter(tmp.name, engine="openpyxl") as writer:
        df_detalle.to_excel(writer, sheet_name="Detalle", index=False)
        df_resumen.to_excel(writer, sheet_name="Resumen", index=False)
        if df_ruta is not None and not df_ruta.empty:
            df_ruta.to_excel(writer, sheet_name="Mapa_Coordenadas", index=False)

    with open(tmp.name, "rb") as f:
        data = f.read()
    return data

# =========================================================
# APP
# =========================================================
init_db()

st.title("APP Área Comercial Buses y Vans")

tab_cot, tab_hist, tab_efi = st.tabs([
    "Nueva cotización",
    "Historial",
    "Eficiencia energética"
])

# =========================================================
# TAB 1 - COTIZACION
# =========================================================
with tab_cot:
    st.subheader("Cotización FOTON U9")

    c1, c2 = st.columns(2)

    with c1:
        fecha = st.date_input("Fecha", value=date.today())
        cliente = st.text_input("Cliente", value="")
        cotizante = st.selectbox("Cotizante", list(COTIZANTES.keys()))

    with c2:
        cantidad_unidades = st.number_input("Cantidad de unidades", min_value=1, value=1, step=1)
        precio_unitario = st.number_input("Precio USD", min_value=0.0, value=130491.0, step=1000.0)
        contrato_mantto = st.text_input("Contrato mantto", value="48 meses")

    capacidad_bateria = st.selectbox(
        "Capacidad nominal batería",
        ["231,8 kWh", "255 kWh"]
    )

    texto = st.text_area(
        "Texto mantenimiento",
        value="""I. Mantención incluida

II. Bono repuestos

III. Telemetría incluida""",
        height=180
    )

    prox = siguiente_correlativo(cotizante)
    prefijo = COTIZANTES[cotizante]["prefijo"]
    st.info(f"Próximo correlativo automático: {prefijo}-{prox:02d}")

    if st.button("Generar cotización", use_container_width=True):
        if not cliente.strip():
            st.error("Debes ingresar el nombre del cliente.")
        else:
            correlativo = siguiente_correlativo(cotizante)
            numero = f"{prefijo}-{correlativo:02d}"
            total_negocio = precio_unitario * cantidad_unidades

            nombre_archivo = (
                f"Propuesta_foton_U9_"
                f"{limpiar_nombre_archivo(cliente)}_"
                f"{limpiar_nombre_archivo(capacidad_bateria)}_"
                f"{fecha_corta(fecha)}"
            )

            contexto = {
                "cliente": cliente.strip(),
                "numero_cotizacion": numero,
                "fecha_larga": fecha_larga_es(fecha),
                "fecha_corta": fecha_corta(fecha),
                "cantidad_unidades": cantidad_unidades,
                "precio_unitario": usd_fmt(precio_unitario),
                "capacidad_bateria": capacidad_bateria,
                "texto_mantto": formatear_texto_mantto(texto) if texto.strip() else "",
                "firma_nombre": COTIZANTES[cotizante]["firma_nombre"],
                "firma_cargo": COTIZANTES[cotizante]["firma_cargo"],
                "firma_correo": COTIZANTES[cotizante]["firma_correo"],
                "firma_telefono": COTIZANTES[cotizante]["firma_telefono"],
            }

            registro = {
                "fecha_iso": fecha.isoformat(),
                "cliente": cliente.strip(),
                "cotizante": cotizante,
                "prefijo": prefijo,
                "correlativo": correlativo,
                "numero_cotizacion": numero,
                "cantidad_unidades": int(cantidad_unidades),
                "precio_unitario_raw": float(precio_unitario),
                "total_negocio_raw": float(total_negocio),
                "contrato_mantto": contrato_mantto,
                "texto_mantto": texto,
                "capacidad_bateria": capacidad_bateria,
            }

            try:
                archivo_docx = generar_docx(contexto, nombre_archivo)
                guardar_cotizacion(registro)

                with open(archivo_docx, "rb") as f:
                    contenido_docx = f.read()

                st.success(f"Cotización {numero} generada correctamente.")
                st.write(f"**Fecha:** {fecha_larga_es(fecha)}")
                st.write(f"**Cliente:** {cliente}")
                st.write(f"**Cotizante:** {cotizante}")
                st.write(f"**Precio unitario:** {usd_fmt(precio_unitario)}")
                st.write(f"**Total negocio:** {usd_fmt(total_negocio)}")

                d1, d2 = st.columns(2)

                with d1:
                    st.download_button(
                        "Descargar Word",
                        data=contenido_docx,
                        file_name=f"{nombre_archivo}.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    )

                try:
                    archivo_pdf = convertir_docx_a_pdf(archivo_docx)
                    with open(archivo_pdf, "rb") as f:
                        contenido_pdf = f.read()

                    with d2:
                        st.download_button(
                            "Descargar PDF",
                            data=contenido_pdf,
                            file_name=f"{nombre_archivo}.pdf",
                            mime="application/pdf"
                        )

                except Exception as e_pdf:
                    st.warning(f"Se generó el Word, pero no fue posible convertir a PDF: {e_pdf}")

            except Exception as e:
                st.error(f"Error al generar la cotización: {e}")

# =========================================================
# TAB 2 - HISTORIAL
# =========================================================
with tab_hist:
    st.subheader("Historial")
    try:
        df_hist = cargar_historial()
        if df_hist.empty:
            st.info("Aún no hay cotizaciones registradas.")
        else:
            df_hist["precio_unitario"] = df_hist["precio_unitario"].apply(usd_fmt)
            df_hist["total_negocio"] = df_hist["total_negocio"].apply(usd_fmt)
            st.dataframe(df_hist, use_container_width=True)
    except Exception as e:
        st.error(f"No fue posible cargar el historial: {e}")

# =========================================================
# TAB 3 - EFICIENCIA
# =========================================================
with tab_efi:
    st.subheader("Eficiencia energética")

    archivo = st.file_uploader("Subir Excel", type=["xlsx", "xls"], key="excel_efi")

    if archivo:
        try:
            df = leer_excel_seguro(archivo)
            st.dataframe(df.head(), use_container_width=True)

            col_kwh = st.selectbox("Columna energía", df.columns)
            col_km = st.selectbox("Columna distancia", df.columns)

            if st.button("Calcular", use_container_width=True):
                df["kwh_km"] = pd.to_numeric(df[col_kwh], errors="coerce") / pd.to_numeric(df[col_km], errors="coerce")
                df = df.dropna(subset=["kwh_km"])

                st.metric("Consumo promedio", f"{df['kwh_km'].mean():.2f} kWh/km")
                st.line_chart(df["kwh_km"])

        except Exception as e:
            st.error(f"Error al procesar el archivo: {e}")