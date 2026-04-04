import os
import sqlite3
import tempfile
import subprocess
import shutil
import base64
import time
from datetime import datetime
from zoneinfo import ZoneInfo

import unicodedata
import plotly.express as px
import pandas as pd
import streamlit as st
import altair as alt
from docxtpl import DocxTemplate, RichText

try:
    import plotly.graph_objects as go
except ImportError:
    st.error("Falta instalar plotly. Ejecuta: pip install plotly")
    st.stop()

try:
    import pydeck as pdk
except ImportError:
    st.error("Falta instalar pydeck. Ejecuta: pip install pydeck")
    st.stop()

# =========================================================
# CONFIGURACION GENERAL
# =========================================================
st.set_page_config(
    page_title="APP Área Comercial Buses y Vans",
    page_icon="🚌",
    layout="wide"
)

DB_FILE = "cotizaciones.db"
LOGO_FILE = "logo_andes_motor.png"

USUARIOS = {
    "dvejar": {"nombre": "Diego Vejar"},
    "ssilva": {"nombre": "Sergio Silva"},
    "acorrea": {"nombre": "Alvaro Correa"},
    "forellana": {"nombre": "Fabian Orellana"},
    "rsepulveda": {"nombre": "Rodrigo Sepulveda"},
}

COTIZANTES = {
    "Diego Vejar": {
        "prefijo": "DV",
        "firma_nombre": "Diego Vejar",
        "firma_cargo": "Subgerente Comercial Buses y Vans (IVECO)",
        "firma_correo": "dvejar@andesmotor.cl",
        "firma_telefono": "981774604",
    },
    "Sergio Silva": {
        "prefijo": "SS",
        "firma_nombre": "Sergio Silva",
        "firma_cargo": "Ejecutivo de ventas Zona Sur Buses y Vans (IVECO)",
        "firma_correo": "sergio.silva@andesmotor.cl",
        "firma_telefono": "",
    },
    "Alvaro Correa": {
        "prefijo": "AC",
        "firma_nombre": "Alvaro Correa",
        "firma_cargo": "Ejecutivo de ventas Zona Norte Buses y Vans (IVECO)",
        "firma_correo": "alvaro.correa@andesmotor.cl",
        "firma_telefono": "",
    },
    "Fabian Orellana": {
        "prefijo": "FO",
        "firma_nombre": "Fabian Orellana",
        "firma_cargo": "Product Manager Senior Buses y Vans (IVECO)",
        "firma_correo": "forellana@andesmotor.cl",
        "firma_telefono": "989356449",
    },
    "Rodrigo Sepulveda": {
        "prefijo": "RS",
        "firma_nombre": "Rodrigo Sepulveda",
        "firma_cargo": "Gerente de Buses y Vans (IVECO)",
        "firma_correo": "rsepulveda@andesmotor.cl",
        "firma_telefono": "979783254",
    },
}

MODELOS = {
    "Foton U9": {
        "tipo": "electrico",
        "template": "plantilla_cotizacion_foton_u9.docx",
        "capacidades": ["231,8 kWh", "255 kWh", "266 kWh"],
    },
    "Foton U10": {
        "tipo": "electrico",
        "template": "plantilla_cotizacion_foton_u10.docx",
        "capacidades": ["266 kWh", "310 kWh"],
    },
    "Foton U12": {
        "tipo": "electrico",
        "template": "plantilla_cotizacion_foton_u12.docx",
        "capacidades": ["247 kWh", "382 kWh"],
    },
    "Foton DU9": {
        "tipo": "diesel",
        "template": "plantilla_cotizacion_foton_du9.docx",
        "capacidades": [],
    },
    "Foton DU10": {
        "tipo": "diesel",
        "template": "plantilla_cotizacion_foton_du10.docx",
        "capacidades": [],
    },
}

# =========================================================
# ESTILO / LOGO CENTRAL TENUE
# =========================================================
st.markdown("""
<style>
.stApp {
    background-image: url("logo_andes_motor.png");
    background-repeat: no-repeat;
    background-position: center;
    background-size: 500px;
}

.stApp::before {
    content: "";
    position: fixed;
    top: 0;
    left: 0;
    width: 100%;
    height: 100%;
    background-image: url("logo_andes_motor.png");
    background-repeat: no-repeat;
    background-position: center;
    background-size: 500px;
    opacity: 0.05; /* 🔥 aquí controlas lo tenue */
    z-index: -1;
}
</style>
""", unsafe_allow_html=True)

agregar_logo_central_tenue(LOGO_FILE)

# =========================================================
# UTILIDADES
# =========================================================
def ahora_santiago():
    return datetime.now(ZoneInfo("America/Santiago"))


def hoy_santiago():
    return ahora_santiago().date()


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
    try:
        return f"USD {valor:,.0f}".replace(",", ".")
    except Exception:
        return "USD 0"


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


def leer_excel_hoja(uploaded_file, sheet_name):
    engine = detectar_motor_excel(uploaded_file)
    if engine:
        return pd.read_excel(uploaded_file, sheet_name=sheet_name, engine=engine)
    return pd.read_excel(uploaded_file, sheet_name=sheet_name)


def sugerir_columna(columnas, candidatos):
    columnas_lower = {str(c).lower(): c for c in columnas}
    for cand in candidatos:
        for col_lower, col_real in columnas_lower.items():
            if cand in col_lower:
                return col_real
    return None


def obtener_template_por_modelo(modelo):
    return MODELOS[modelo]["template"]

def normalizar_texto(texto):
    if texto is None:
        return ""
    texto = str(texto).strip().lower()
    texto = unicodedata.normalize("NFKD", texto).encode("ascii", "ignore").decode("utf-8")
    return texto


def normalizar_columnas(df):
    nuevas = {}
    for c in df.columns:
        c_norm = normalizar_texto(c)
        nuevas[c] = c_norm
    return df.rename(columns=nuevas)
# =========================================================
# BASE DE DATOS
# =========================================================
def get_conn():
    return sqlite3.connect(DB_FILE, check_same_thread=False)


def table_exists(conn, table_name):
    cur = conn.cursor()
    cur.execute(
        "SELECT name FROM sqlite_master WHERE type='table' AND name=?",
        (table_name,)
    )
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
            modelo TEXT,
            capacidad_bateria TEXT,
            cantidad_unidades INTEGER NOT NULL,
            precio_unitario REAL NOT NULL,
            total_negocio REAL NOT NULL,
            lugar_entrega TEXT,
            contrato_mantto TEXT,
            texto_mantto TEXT,
            creado_en TEXT NOT NULL
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
    cur = conn.cursor()

    faltantes = {
        "modelo": "TEXT",
        "capacidad_bateria": "TEXT",
        "lugar_entrega": "TEXT",
        "contrato_mantto": "TEXT",
        "texto_mantto": "TEXT",
        "creado_en": "TEXT",
        "fecha": "TEXT",
        "prefijo": "TEXT",
        "correlativo": "INTEGER",
        "cantidad_unidades": "INTEGER DEFAULT 1",
        "precio_unitario": "REAL DEFAULT 0",
        "total_negocio": "REAL DEFAULT 0",
    }

    if "numero_cotizacion" in columnas:
        for col, tipo in faltantes.items():
            if col not in columnas:
                cur.execute(f"ALTER TABLE cotizaciones ADD COLUMN {col} {tipo}")
        conn.commit()
        conn.close()
        return

    if "numero" in columnas:
        cur.execute("ALTER TABLE cotizaciones RENAME TO cotizaciones_old")
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
                modelo, capacidad_bateria, cantidad_unidades, precio_unitario, total_negocio,
                lugar_entrega, contrato_mantto, texto_mantto, creado_en
            )
            SELECT
                {fecha_expr},
                {cliente_expr},
                {cotizante_expr},
                '',
                0,
                {numero_expr},
                '',
                '',
                1,
                0,
                0,
                '',
                '',
                '',
                ''
            FROM cotizaciones_old
        """)
        conn.commit()
        conn.close()
        return

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
        SELECT correlativo
        FROM cotizaciones
        WHERE prefijo = ?
        ORDER BY correlativo ASC
    """, (prefijo,))
    usados = [row[0] for row in cur.fetchall() if row[0] is not None]
    conn.close()

    correlativo = 1
    while correlativo in usados:
        correlativo += 1
    return correlativo


def guardar_cotizacion(data):
    conn = get_conn()
    cur = conn.cursor()
    creado_en = ahora_santiago().strftime("%Y-%m-%d %H:%M:%S")

    cur.execute("""
        INSERT INTO cotizaciones (
            fecha, cliente, cotizante, prefijo, correlativo, numero_cotizacion,
            modelo, capacidad_bateria, cantidad_unidades, precio_unitario, total_negocio,
            lugar_entrega, contrato_mantto, texto_mantto, creado_en
        )
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    """, (
        data["fecha_iso"],
        data["cliente"],
        data["cotizante"],
        data["prefijo"],
        data["correlativo"],
        data["numero_cotizacion"],
        data["modelo"],
        data["capacidad_bateria"],
        data["cantidad_unidades"],
        data["precio_unitario_raw"],
        data["total_negocio_raw"],
        data["lugar_entrega"],
        data["contrato_mantto"],
        data["texto_mantto"],
        creado_en,
    ))
    conn.commit()
    conn.close()


def cargar_historial():
    conn = get_conn()
    df = pd.read_sql_query("""
        SELECT
            id,
            fecha,
            cliente,
            cotizante,
            numero_cotizacion,
            modelo,
            cantidad_unidades,
            precio_unitario,
            total_negocio,
            COALESCE(capacidad_bateria, '') AS capacidad_bateria,
            COALESCE(lugar_entrega, '') AS lugar_entrega,
            COALESCE(creado_en, '') AS creado_en
        FROM cotizaciones
        ORDER BY id DESC
    """, conn)
    conn.close()
    return df


def eliminar_cotizacion_por_id(cotizacion_id):
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("DELETE FROM cotizaciones WHERE id = ?", (cotizacion_id,))
    conn.commit()
    conn.close()

# =========================================================
# DOCX / PDF
# =========================================================
def generar_docx(contexto, template_file, nombre_salida):
    if not os.path.exists(template_file):
        raise FileNotFoundError(f"No se encontró la plantilla '{template_file}'.")

    doc = DocxTemplate(template_file)
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
        raise RuntimeError(
            "No se encontró LibreOffice/soffice en el servidor. Instala LibreOffice para habilitar PDF."
        )

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
            f"Error al convertir PDF. STDOUT: {resultado.stdout} | STDERR: {resultado.stderr}"
        )

    pdf_path = os.path.splitext(docx_path)[0] + ".pdf"
    if not os.path.exists(pdf_path):
        raise RuntimeError("No se generó el archivo PDF.")

    return pdf_path

# =========================================================
# LOGIN
# =========================================================
if "usuario" not in st.session_state:
    st.session_state.usuario = None

if st.session_state.usuario is None:
    c1, c2, c3 = st.columns([1, 2, 1])
    with c2:
        if os.path.exists(LOGO_FILE):
            st.image(LOGO_FILE, width=180)
        st.title("Ingreso Área Comercial")
        user = st.text_input("Usuario")
        login_btn = st.button("Ingresar", use_container_width=True)

        if login_btn:
            if user in USUARIOS:
                st.session_state.usuario = user
                st.rerun()
            else:
                st.error("Usuario no válido")
    st.stop()

usuario_actual = USUARIOS[st.session_state.usuario]["nombre"]

# =========================================================
# APP
# =========================================================
init_db()

header1, header2 = st.columns([1.2, 6])

with header1:
    if os.path.exists(LOGO_FILE):
        st.image(LOGO_FILE, width=180)

with header2:
    st.markdown(
        """
        <div style="padding-top: 8px;">
            <h1 style='margin-bottom:0;'>APP Área Comercial Buses y Vans</h1>
            <p style='margin-top:0;color:gray;'>Andes Motor - Plataforma Comercial</p>
        </div>
        """,
        unsafe_allow_html=True
    )

st.markdown("---")

st.sidebar.success(f"Usuario: {usuario_actual}")
st.sidebar.caption(f"Hora Santiago: {ahora_santiago().strftime('%d-%m-%Y %H:%M:%S')}")
if st.sidebar.button("Cerrar sesión"):
    st.session_state.usuario = None
    st.rerun()

tab_cot, tab_hist, tab_efi, tab_dash = st.tabs([
    "🧾 Nueva cotización",
    "📚 Historial",
    "⚡ Eficiencia energética",
    "📊 Dashboard Comercial"
])

# =========================================================
# TAB 1 - COTIZACION
# =========================================================
with tab_cot:
    st.subheader("Cotizaciones Foton")

    cotizante = usuario_actual
    modelo = st.selectbox("Modelo", list(MODELOS.keys()))
    modelo_info = MODELOS[modelo]

    c1, c2 = st.columns(2)

    with c1:
        fecha = st.date_input("Fecha", value=hoy_santiago())
        cliente = st.text_input("Cliente", value="")
        st.text_input("Cotizante", value=cotizante, disabled=True)
        lugar_entrega = st.text_input("Lugar de entrega", value="")

    with c2:
        cantidad_unidades = st.number_input("Cantidad de unidades", min_value=1, value=1, step=1)
        precio_unitario = st.number_input("Precio unitario USD", min_value=0.0, value=130491.0, step=1000.0)
        contrato_mantto = st.text_input("Contrato mantto", value="48 meses")

        if modelo_info["tipo"] == "electrico":
            capacidad_bateria = st.selectbox("Capacidad nominal batería", modelo_info["capacidades"])
        else:
            capacidad_bateria = ""

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
        elif not lugar_entrega.strip():
            st.error("Debes ingresar el lugar de entrega.")
        else:
            correlativo = siguiente_correlativo(cotizante)
            numero = f"{prefijo}-{correlativo:02d}"
            total_negocio = precio_unitario * cantidad_unidades

            nombre_archivo = (
                f"Propuesta_{limpiar_nombre_archivo(modelo)}_"
                f"{limpiar_nombre_archivo(cliente)}_"
                f"{limpiar_nombre_archivo(capacidad_bateria) if capacidad_bateria else 'diesel'}_"
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
                "lugar_entrega": lugar_entrega.strip(),
            }

            registro = {
                "fecha_iso": fecha.isoformat(),
                "cliente": cliente.strip(),
                "cotizante": cotizante,
                "prefijo": prefijo,
                "correlativo": correlativo,
                "numero_cotizacion": numero,
                "modelo": modelo,
                "capacidad_bateria": capacidad_bateria,
                "cantidad_unidades": int(cantidad_unidades),
                "precio_unitario_raw": float(precio_unitario),
                "total_negocio_raw": float(total_negocio),
                "lugar_entrega": lugar_entrega.strip(),
                "contrato_mantto": contrato_mantto,
                "texto_mantto": texto,
            }

            try:
                template_file = obtener_template_por_modelo(modelo)
                archivo_docx = generar_docx(contexto, template_file, nombre_archivo)
                guardar_cotizacion(registro)

                with open(archivo_docx, "rb") as f:
                    contenido_docx = f.read()

                st.success(f"Cotización {numero} generada correctamente.")
                st.write(f"**Fecha:** {fecha_larga_es(fecha)}")
                st.write(f"**Cliente:** {cliente}")
                st.write(f"**Cotizante:** {cotizante}")
                st.write(f"**Modelo:** {modelo}")
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
                    inicio_pdf = time.time()
                    with st.spinner("Generando PDF..."):
                        archivo_pdf = convertir_docx_a_pdf(archivo_docx)

                    fin_pdf = time.time()
                    segundos_pdf = round(fin_pdf - inicio_pdf, 2)

                    with open(archivo_pdf, "rb") as f:
                        contenido_pdf = f.read()

                    with d2:
                        st.download_button(
                            "Descargar PDF",
                            data=contenido_pdf,
                            file_name=f"{nombre_archivo}.pdf",
                            mime="application/pdf"
                        )

                    st.caption(f"PDF generado en {segundos_pdf} segundos.")

                except Exception as e_pdf:
                    st.warning(f"No fue posible habilitar el PDF: {e_pdf}")

            except Exception as e:
                st.error(f"Error al generar la cotización: {e}")

# =========================================================
# TAB 2 - HISTORIAL / ELIMINAR
# =========================================================
with tab_hist:
    st.subheader("Historial")

    try:
        df_hist = cargar_historial()
        if df_hist.empty:
            st.info("Aún no hay cotizaciones registradas.")
        else:
            vista = df_hist.copy()
            vista["precio_unitario"] = vista["precio_unitario"].apply(usd_fmt)
            vista["total_negocio"] = vista["total_negocio"].apply(usd_fmt)
            st.dataframe(vista, use_container_width=True)

            st.markdown("### Eliminar cotización")
            opciones = [
                f"{row['id']} | {row['numero_cotizacion']} | {row['cliente']} | {row['modelo']}"
                for _, row in df_hist.iterrows()
            ]
            seleccion = st.selectbox("Selecciona una cotización para eliminar", opciones)

            if st.button("Eliminar cotización seleccionada", type="secondary"):
                cotizacion_id = int(seleccion.split("|")[0].strip())
                eliminar_cotizacion_por_id(cotizacion_id)
                st.success("Cotización eliminada. El correlativo queda disponible nuevamente.")
                st.rerun()

    except Exception as e:
        st.error(f"No fue posible cargar el historial: {e}")


# =========================================================
# TAB 3 - EFICIENCIA ENERGÉTICA
# =========================================================
with tab_efi:
    import pandas as pd
    import plotly.graph_objects as go
    import plotly.express as px
    import unicodedata
    from io import BytesIO
    from reportlab.lib.pagesizes import A4
    from reportlab.lib import colors
    from reportlab.lib.units import cm
    from reportlab.platypus import (
        SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image
    )
    from reportlab.lib.styles import getSampleStyleSheet

    st.subheader("⚡ Eficiencia energética")

    archivo = st.file_uploader("Subir Excel (.xlsx)", type=["xlsx"], key="efi_tab3")

    # -----------------------------------------------------
    # FUNCIONES AUXILIARES
    # -----------------------------------------------------
    def norm(txt):
        if pd.isna(txt):
            return ""
        txt = str(txt).lower().strip()
        return unicodedata.normalize("NFKD", txt).encode("ascii", "ignore").decode("utf-8")

    def buscar_columna(cols, candidatos):
        cols_norm = {norm(c): c for c in cols}
        for cand in candidatos:
            cand_n = norm(cand)
            for c_n, c_real in cols_norm.items():
                if cand_n in c_n:
                    return c_real
        return None

    def fig_to_png_bytes(fig, width=1000, height=500):
        return fig.to_image(format="png", width=width, height=height, scale=2)

    def generar_pdf_ejecutivo(
        trazado,
        bateria,
        distancia,
        consumo,
        rendimiento,
        autonomia_total,
        autonomia_15,
        vel_prom,
        fig_vs,
        fig_map,
        fig_alt
    ):
        buffer = BytesIO()
        doc = SimpleDocTemplate(
            buffer,
            pagesize=A4,
            rightMargin=1.2 * cm,
            leftMargin=1.2 * cm,
            topMargin=1.0 * cm,
            bottomMargin=1.0 * cm
        )

        styles = getSampleStyleSheet()
        story = []

        titulo = Paragraph(f"<b>Informe Eficiencia Energética - {trazado}</b>", styles["Title"])
        story.append(titulo)
        story.append(Spacer(1, 0.3 * cm))

        subt = Paragraph("Resumen ejecutivo del recorrido seleccionado", styles["Heading2"])
        story.append(subt)
        story.append(Spacer(1, 0.2 * cm))

        data = [
            ["Indicador", "Valor"],
            ["Trazado", str(trazado)],
            ["Batería HV", f"{bateria:.1f} kWh"],
            ["Distancia", f"{distancia:.1f} km"],
            ["Consumo energético", f"{consumo:.2f} kWh"],
            ["Rendimiento", f"{rendimiento:.3f} kWh/km"],
            ["Autonomía proyectada", f"{autonomia_total:.0f} km"],
            ["Autonomía útil al 15% SoC", f"{autonomia_15:.0f} km"],
            ["Velocidad promedio", f"{vel_prom:.1f} km/h"],
        ]

        tabla = Table(data, colWidths=[7 * cm, 7 * cm])
        tabla.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1f4e78")),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
            ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
            ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
            ("BACKGROUND", (0, 1), (-1, -1), colors.whitesmoke),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ]))
        story.append(tabla)
        story.append(Spacer(1, 0.5 * cm))

        story.append(Paragraph("<b>Velocidad y Estado de Carga</b>", styles["Heading3"]))
        img1 = Image(BytesIO(fig_to_png_bytes(fig_vs, 1200, 450)))
        img1.drawWidth = 17 * cm
        img1.drawHeight = 6.5 * cm
        story.append(img1)
        story.append(Spacer(1, 0.3 * cm))

        story.append(Paragraph("<b>Mapa del recorrido</b>", styles["Heading3"]))
        img2 = Image(BytesIO(fig_to_png_bytes(fig_map, 1200, 600)))
        img2.drawWidth = 17 * cm
        img2.drawHeight = 8.2 * cm
        story.append(img2)
        story.append(Spacer(1, 0.3 * cm))

        story.append(Paragraph("<b>Perfil de altura</b>", styles["Heading3"]))
        img3 = Image(BytesIO(fig_to_png_bytes(fig_alt, 1200, 400)))
        img3.drawWidth = 17 * cm
        img3.drawHeight = 5.8 * cm
        story.append(img3)

        doc.build(story)
        buffer.seek(0)
        return buffer.getvalue()

    if archivo:
        try:
            # =================================================
            # 1) CARGA DE HOJAS
            # =================================================
            xls = pd.ExcelFile(archivo)
            hojas = xls.sheet_names

            if len(hojas) < 2:
                st.error("El archivo debe traer al menos 2 hojas: base y resumen.")
                st.stop()

            df_base = pd.read_excel(archivo, sheet_name=0)
            df_resumen = pd.read_excel(archivo, sheet_name=1)

            df_base_original = df_base.copy()
            df_resumen_original = df_resumen.copy()

            df_base.columns = [norm(c) for c in df_base.columns]
            df_resumen.columns = [norm(c) for c in df_resumen.columns]

            # =================================================
            # 2) DETECCIÓN FLEXIBLE DE COLUMNAS
            # =================================================
            col_trazado_base = buscar_columna(df_base.columns, ["trazado", "ruta"])
            col_odo = buscar_columna(df_base.columns, ["odometro", "odómetro"])
            col_vel = buscar_columna(df_base.columns, ["velocidad"])
            col_soc = buscar_columna(df_base.columns, ["soc", "estado de carga"])
            col_alt = buscar_columna(df_base.columns, ["altitud", "altura"])
            col_lat = buscar_columna(df_base.columns, ["latitud", "latitude", "lat"])
            col_lon = buscar_columna(df_base.columns, ["longitud", "longitude", "long", "lon"])

            col_trazado_res = buscar_columna(df_resumen.columns, ["trazado", "ruta"])
            col_distancia = buscar_columna(df_resumen.columns, ["distancia"])
            col_consumo = buscar_columna(df_resumen.columns, ["consumo energetico", "consumo energético", "consumo"])

            faltantes_base = []
            if col_trazado_base is None: faltantes_base.append("trazado")
            if col_odo is None: faltantes_base.append("odometro")
            if col_vel is None: faltantes_base.append("velocidad")
            if col_soc is None: faltantes_base.append("soc")
            if col_alt is None: faltantes_base.append("altitud")
            if col_lat is None: faltantes_base.append("latitud")
            if col_lon is None: faltantes_base.append("longitud")

            faltantes_res = []
            if col_trazado_res is None: faltantes_res.append("trazado")
            if col_distancia is None: faltantes_res.append("distancia")
            if col_consumo is None: faltantes_res.append("consumo energetico")

            if faltantes_base:
                st.error(f"Faltan columnas en hoja base: {', '.join(faltantes_base)}")
                st.stop()

            if faltantes_res:
                st.error(f"Faltan columnas en hoja resumen: {', '.join(faltantes_res)}")
                st.stop()

            # =================================================
            # 3) NORMALIZACIÓN DE DATOS
            # =================================================
            df_base["trazado_key"] = df_base[col_trazado_base].astype(str).apply(norm)
            df_resumen["trazado_key"] = df_resumen[col_trazado_res].astype(str).apply(norm)

            df_base["odometro"] = pd.to_numeric(df_base[col_odo], errors="coerce")
            df_base["velocidad"] = pd.to_numeric(df_base[col_vel], errors="coerce")
            df_base["soc"] = pd.to_numeric(df_base[col_soc], errors="coerce")
            df_base["altitud"] = pd.to_numeric(df_base[col_alt], errors="coerce")
            df_base["lat"] = pd.to_numeric(df_base[col_lat], errors="coerce")
            df_base["lon"] = pd.to_numeric(df_base[col_lon], errors="coerce")

            df_resumen["distancia_calc"] = pd.to_numeric(df_resumen[col_distancia], errors="coerce")
            df_resumen["consumo_calc"] = pd.to_numeric(df_resumen[col_consumo], errors="coerce")

            trazados = (
                df_resumen[[col_trazado_res, "trazado_key"]]
                .dropna()
                .drop_duplicates()
                .rename(columns={col_trazado_res: "trazado"})
                .sort_values("trazado")
            )

            if trazados.empty:
                st.error("No se encontraron trazados válidos.")
                st.stop()

            # =================================================
            # 4) SELECTOR Y BATERÍA
            # =================================================
            csel1, csel2 = st.columns([2, 1])

            with csel1:
                trazado_sel = st.selectbox("Seleccionar trazado", trazados["trazado"].tolist())

            with csel2:
                bateria = st.selectbox("Batería HV (kWh)", [231.8, 255.0, 266.0, 310.0, 382.0], index=1)

            key = norm(trazado_sel)

            base = df_base[df_base["trazado_key"] == key].copy()
            base = base.dropna(subset=["odometro"]).sort_values("odometro")

            resumen_sel = df_resumen[df_resumen["trazado_key"] == key]

            if resumen_sel.empty:
                st.error("No se encontró el trazado en la hoja resumen.")
                st.stop()

            res = resumen_sel.iloc[0]

            # =================================================
            # 5) CÁLCULOS SOLO DESDE HOJA 2
            # =================================================
            distancia = float(res["distancia_calc"]) if pd.notna(res["distancia_calc"]) else None
            consumo = float(res["consumo_calc"]) if pd.notna(res["consumo_calc"]) else None

            if distancia is None or distancia <= 0:
                st.error("La distancia de la hoja resumen no es válida.")
                st.stop()

            if consumo is None or consumo <= 0:
                st.error("El consumo energético de la hoja resumen no es válido.")
                st.stop()

            rendimiento = consumo / distancia
            autonomia = bateria / rendimiento if rendimiento > 0 else None

            # nueva autonomía considerando 15% de reserva SoC
            bateria_util_15 = bateria * 0.85
            autonomia_15 = bateria_util_15 / rendimiento if rendimiento > 0 else None

            vel_prom = base["velocidad"].mean() if base["velocidad"].notna().any() else None

            # =================================================
            # 6) KPIs
            # =================================================
            c1, c2, c3, c4 = st.columns(4)
            c1.metric("Rendimiento", f"{rendimiento:.3f} kWh/km")
            c2.metric("Autonomía proyectada", f"{autonomia:.0f} km" if autonomia is not None else "Sin dato")
            c3.metric("Velocidad promedio", f"{vel_prom:.1f} km/h" if vel_prom is not None else "Sin dato")
            c4.metric("Distancia", f"{distancia:.1f} km")

            # =================================================
            # 7) RELOJES / GAUGES COMPACTOS
            # =================================================
            st.markdown("<div style='height:25px;'></div>", unsafe_allow_html=True)

            g1, g2, g3 = st.columns(3)

            with g1:
                fig_g1 = go.Figure(go.Indicator(
                    mode="gauge+number",
                    value=rendimiento,
                    number={"suffix": " kWh/km", "font": {"size": 26}},
                    title={"text": "Rendimiento", "font": {"size": 18}},
                    gauge={
                        "axis": {"range": [0, 1.2], "tickfont": {"size": 11}},
                        "steps": [
                            {"range": [0, 0.90], "color": "#22c55e"},
                            {"range": [0.90, 1.00], "color": "#facc15"},
                            {"range": [1.00, 1.20], "color": "#ef4444"},
                        ],
                        "bar": {"color": "#15803d"},
                    }
                ))
                fig_g1.update_layout(height=240, margin=dict(l=5, r=5, t=45, b=10))
                st.plotly_chart(fig_g1, use_container_width=True)

            with g2:
                fig_g2 = go.Figure(go.Indicator(
                    mode="gauge+number",
                    value=autonomia if autonomia is not None else 0,
                    number={"suffix": " km", "font": {"size": 26}},
                    title={"text": f"Autonomía total ({bateria:.0f} kWh)", "font": {"size": 16}},
                    gauge={
                        "axis": {"range": [0, 500], "tickfont": {"size": 11}},
                        "steps": [
                            {"range": [0, 280], "color": "#ef4444"},
                            {"range": [280, 350], "color": "#facc15"},
                            {"range": [350, 500], "color": "#22c55e"},
                        ],
                        "bar": {"color": "#15803d"},
                    }
                ))
                fig_g2.update_layout(height=240, margin=dict(l=5, r=5, t=45, b=10))
                st.plotly_chart(fig_g2, use_container_width=True)

            with g3:
                fig_g3 = go.Figure(go.Indicator(
                    mode="gauge+number",
                    value=autonomia_15 if autonomia_15 is not None else 0,
                    number={"suffix": " km", "font": {"size": 26}},
                    title={"text": f"Autonomía útil al 15% SoC", "font": {"size": 16}},
                    gauge={
                        "axis": {"range": [0, 500], "tickfont": {"size": 11}},
                        "steps": [
                            {"range": [0, 250], "color": "#ef4444"},
                            {"range": [250, 320], "color": "#facc15"},
                            {"range": [320, 500], "color": "#22c55e"},
                        ],
                        "bar": {"color": "#166534"},
                    }
                ))
                fig_g3.update_layout(height=240, margin=dict(l=5, r=5, t=45, b=10))
                st.plotly_chart(fig_g3, use_container_width=True)

            # =================================================
            # 8) VELOCIDAD VS SOC
            # =================================================
            st.markdown("### Velocidad y Estado de Carga")

            fig_vs = go.Figure()

            fig_vs.add_trace(go.Scatter(
                x=base["odometro"],
                y=base["velocidad"],
                mode="lines+markers",
                name="Velocidad",
                line=dict(color="#2563eb", width=3, shape="spline"),
                marker=dict(size=5)
            ))

            if base["soc"].notna().any():
                fig_vs.add_trace(go.Scatter(
                    x=base["odometro"],
                    y=base["soc"],
                    mode="lines+markers",
                    name="SoC",
                    yaxis="y2",
                    line=dict(color="#60a5fa", width=3, shape="spline"),
                    marker=dict(size=4)
                ))

            fig_vs.update_layout(
                height=380,
                margin=dict(l=10, r=10, t=20, b=10),
                xaxis=dict(title="Odómetro km"),
                yaxis=dict(title="Velocidad km/h"),
                yaxis2=dict(
                    title="Estado de carga SoC",
                    overlaying="y",
                    side="right"
                ),
                legend=dict(orientation="h")
            )

            st.plotly_chart(fig_vs, use_container_width=True)

            # =================================================
            # 9) MAPA DEL RECORRIDO
            # =================================================
            st.markdown("### 🛰️ Mapa del recorrido")

            mapa = base.dropna(subset=["lat", "lon"]).copy()

            if mapa.empty:
                st.warning("No hay coordenadas válidas para mostrar el mapa.")
                fig_map = go.Figure()
            else:
                mapa["tipo_punto"] = "Punto recorrido"
                mapa_inicio = mapa.iloc[[0]].copy()
                mapa_inicio["tipo_punto"] = "Inicio"
                mapa_fin = mapa.iloc[[-1]].copy()
                mapa_fin["tipo_punto"] = "Fin"

                mapa_plot = pd.concat([mapa, mapa_inicio, mapa_fin], ignore_index=True)

                fig_map = px.scatter_mapbox(
                    mapa_plot,
                    lat="lat",
                    lon="lon",
                    color="tipo_punto",
                    color_discrete_map={
                        "Punto recorrido": "#1d4ed8",
                        "Inicio": "#16a34a",
                        "Fin": "#dc2626"
                    },
                    hover_data={
                        "odometro": True,
                        "velocidad": True,
                        "soc": True,
                        "lat": False,
                        "lon": False
                    },
                    zoom=12,
                    height=520
                )

                fig_map.add_trace(go.Scattermapbox(
                    lat=mapa["lat"],
                    lon=mapa["lon"],
                    mode="lines",
                    line=dict(width=4, color="#2563eb"),
                    name="Ruta"
                ))

                fig_map.update_layout(
                    mapbox_style="open-street-map",
                    margin=dict(r=0, t=0, l=0, b=0),
                    legend=dict(orientation="h", yanchor="bottom", y=1.01, xanchor="right", x=1)
                )

                st.plotly_chart(fig_map, use_container_width=True)

            # =================================================
            # 10) PERFIL DE ALTURA
            # =================================================
            if base["altitud"].notna().any():
                st.markdown("### Perfil de altura")

                fig_alt = go.Figure()
                fig_alt.add_trace(go.Scatter(
                    x=base["odometro"],
                    y=base["altitud"],
                    fill="tozeroy",
                    mode="lines",
                    line=dict(color="#3b82f6", width=3, shape="spline")
                ))

                fig_alt.update_layout(
                    height=320,
                    margin=dict(l=10, r=10, t=20, b=10),
                    xaxis=dict(title="Odómetro"),
                    yaxis=dict(title="Altura (m)")
                )

                st.plotly_chart(fig_alt, use_container_width=True)
            else:
                fig_alt = go.Figure()

            # =================================================
            # 11) BOTÓN PDF EJECUTIVO
            # =================================================
            try:
                pdf_bytes = generar_pdf_ejecutivo(
                    trazado=trazado_sel,
                    bateria=bateria,
                    distancia=distancia,
                    consumo=consumo,
                    rendimiento=rendimiento,
                    autonomia_total=autonomia,
                    autonomia_15=autonomia_15,
                    vel_prom=vel_prom,
                    fig_vs=fig_vs,
                    fig_map=fig_map,
                    fig_alt=fig_alt
                )

                st.download_button(
                    "📄 Descargar informe PDF",
                    data=pdf_bytes,
                    file_name=f"Informe_Eficiencia_{trazado_sel.replace(' ', '_')}.pdf",
                    mime="application/pdf"
                )
            except Exception as e_pdf:
                st.warning(f"No fue posible generar el PDF: {e_pdf}")

            # =================================================
            # 12) TABLA RESUMEN VISIBLE
            # =================================================
            st.markdown("### Tabla resumen")
            st.dataframe(df_resumen_original, use_container_width=True)

            # =================================================
            # 13) DETALLE OCULTO
            # =================================================
            with st.expander("Ver detalle de la base"):
                st.dataframe(df_base_original.head(200), use_container_width=True)

        except Exception as e:
            st.error(f"Error: {e}")
# =========================================================
# TAB 4 - DASHBOARD
# =========================================================
with tab_dash:
    st.subheader("Dashboard Comercial")

    try:
        df_dash = cargar_historial()

        if df_dash.empty:
            st.info("No hay datos aún.")
        else:
            df_dash["total_negocio"] = pd.to_numeric(df_dash["total_negocio"], errors="coerce").fillna(0)
            df_dash["precio_unitario"] = pd.to_numeric(df_dash["precio_unitario"], errors="coerce").fillna(0)
            df_dash["cantidad_unidades"] = pd.to_numeric(df_dash["cantidad_unidades"], errors="coerce").fillna(0)
            df_dash["fecha_dt"] = pd.to_datetime(df_dash["fecha"], errors="coerce")

            c1, c2, c3 = st.columns(3)
            c1.metric("Total cotizaciones", len(df_dash))
            c2.metric("Monto total negocio", usd_fmt(df_dash["total_negocio"].sum()))
            c3.metric("Promedio por cotización", usd_fmt(df_dash["total_negocio"].mean()))

            df_precio = df_dash.dropna(subset=["fecha_dt"]).sort_values("fecha_dt")

            st.markdown("### Evolución del precio unitario")
            graf_precio = alt.Chart(df_precio).mark_line(point=True).encode(
                x=alt.X("fecha_dt:T", title="Fecha"),
                y=alt.Y("precio_unitario:Q", title="Precio unitario (USD)"),
                tooltip=["fecha", "cliente", "cotizante", "precio_unitario"]
            ).properties(height=320)
            st.altair_chart(graf_precio, use_container_width=True)

            st.markdown("### Evolución del total negocio")
            graf_total = alt.Chart(df_precio).mark_bar().encode(
                x=alt.X("fecha_dt:T", title="Fecha"),
                y=alt.Y("total_negocio:Q", title="Total negocio (USD)"),
                color=alt.Color("cotizante:N", title="Cotizante"),
                tooltip=["fecha", "cliente", "cotizante", "total_negocio"]
            ).properties(height=320)
            st.altair_chart(graf_total, use_container_width=True)

            st.markdown("### Relación entre precio unitario y total negocio")
            graf_scatter = alt.Chart(df_dash).mark_circle(size=120).encode(
                x=alt.X("precio_unitario:Q", title="Precio unitario (USD)"),
                y=alt.Y("total_negocio:Q", title="Total negocio (USD)"),
                color=alt.Color("cotizante:N", title="Cotizante"),
                size=alt.Size("cantidad_unidades:Q", title="Cantidad unidades"),
                tooltip=["cliente", "cotizante", "cantidad_unidades", "precio_unitario", "total_negocio"]
            ).properties(height=350)
            st.altair_chart(graf_scatter, use_container_width=True)

            st.markdown("### Detalle")
            df_v = df_dash.copy()
            df_v["precio_unitario"] = df_v["precio_unitario"].apply(usd_fmt)
            df_v["total_negocio"] = df_v["total_negocio"].apply(usd_fmt)
            st.dataframe(df_v, use_container_width=True)

    except Exception as e:
        st.error(f"Error dashboard: {e}")