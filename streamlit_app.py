import os
import sqlite3
import tempfile
import subprocess
import shutil
import base64
import time
from datetime import datetime
from zoneinfo import ZoneInfo

import pandas as pd
import streamlit as st
import altair as alt
from docxtpl import DocxTemplate, RichText

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
def agregar_logo_central_tenue(ruta_logo: str):
    if not os.path.exists(ruta_logo):
        return

    with open(ruta_logo, "rb") as f:
        logo_base64 = base64.b64encode(f.read()).decode()

    st.markdown(
        f"""
        <style>
        .stApp {{
            background-image: url("data:image/png;base64,{logo_base64}");
            background-repeat: no-repeat;
            background-position: center 66%;
            background-size: 260px;
            background-attachment: fixed;
        }}

        .stApp::before {{
            content: "";
            position: fixed;
            inset: 0;
            background: rgba(255,255,255,0.975);
            z-index: -1;
        }}
        </style>
        """,
        unsafe_allow_html=True
    )

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


def clp_fmt(valor):
    try:
        return f"$ {valor:,.0f}".replace(",", ".")
    except Exception:
        return "$ 0"


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


def sugerir_columna(columnas, candidatos):
    columnas_lower = {c.lower(): c for c in columnas}
    for cand in candidatos:
        for col_lower, col_real in columnas_lower.items():
            if cand in col_lower:
                return col_real
    return None


def obtener_template_por_modelo(modelo):
    return MODELOS[modelo]["template"]

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
        return f.read()

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
        value="""I. La oferta incluye 48 meses de mantenimiento preventivo y correctivo (Correctivo de desgaste Disco y Pastilla de frenos) sin costo para el cliente, con el fin de entregar conocimientos técnicos y aprendizaje continuo del mantenimiento para este tipo de vehículos, durante este periodo.

II. Adicionalmente se entregarán USD 1.500.- por bus, para la compra de repuestos de desgaste a elección del cliente (este ítem deberá ser utilizado dentro de los primeros 12 meses realizada la entrega de flota).

III. La oferta incluye 8 años de telemetría sin costo para el cliente.""",
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
# TAB 3 - EFICIENCIA
# =========================================================
with tab_efi:
    st.subheader("Eficiencia energética")
    st.write("Sube una planilla Excel con la hoja de eventos y la hoja resumen por trazado.")

    archivo = st.file_uploader("Subir Excel", type=["xlsx", "xls"], key="excel_efi")

    if archivo:
        try:
            # -----------------------------
            # CARGA DE HOJAS
            # -----------------------------
            xls = pd.ExcelFile(archivo)
            hojas = xls.sheet_names

            if len(hojas) < 1:
                st.error("El archivo no contiene hojas válidas.")
            else:
                hoja_datos = hojas[0]
                hoja_resumen = hojas[1] if len(hojas) > 1 else None

                df = pd.read_excel(archivo, sheet_name=hoja_datos, engine="openpyxl")
                df_resumen_excel = None
                if hoja_resumen:
                    df_resumen_excel = pd.read_excel(archivo, sheet_name=hoja_resumen, engine="openpyxl")

                st.markdown("### Vista previa hoja principal")
                st.dataframe(df.head(10), use_container_width=True)

                if df_resumen_excel is not None:
                    st.markdown("### Vista previa hoja resumen")
                    st.dataframe(df_resumen_excel, use_container_width=True)

                # -----------------------------
                # NORMALIZACION DE COLUMNAS
                # -----------------------------
                columnas = list(df.columns)

                def buscar_columna(candidatos):
                    columnas_lower = {c.lower(): c for c in columnas}
                    for cand in candidatos:
                        for col_lower, col_real in columnas_lower.items():
                            if cand in col_lower:
                                return col_real
                    return None

                col_fecha = buscar_columna(["fecha evento", "fecha"])
                col_lon = buscar_columna(["longitud", "lon"])
                col_lat = buscar_columna(["latitud", "lat"])
                col_alt = buscar_columna(["altitud", "altura"])
                col_odo = buscar_columna(["odometro", "odómetro"])
                col_soc = buscar_columna(["soc"])
                col_vel = buscar_columna(["velocidad"])
                col_energia = buscar_columna(["energia consumida por viaje", "energía consumida por viaje"])
                col_consumo_kw = buscar_columna(["consumo kw"])
                col_trazado = buscar_columna(["trazado", "ruta"])

                obligatorias = {
                    "Trazado": col_trazado,
                    "Odometro": col_odo,
                    "Velocidad": col_vel,
                    "SoC": col_soc,
                    "Latitud": col_lat,
                    "Longitud": col_lon,
                    "Altitud": col_alt,
                }

                faltantes = [k for k, v in obligatorias.items() if v is None]
                if faltantes:
                    st.error(f"Faltan columnas clave en la hoja principal: {', '.join(faltantes)}")
                else:
                    trabajo = df.copy()

                    if col_fecha:
                        trabajo["fecha_evento"] = pd.to_datetime(trabajo[col_fecha], errors="coerce")
                    else:
                        trabajo["fecha_evento"] = pd.NaT

                    trabajo["trazado"] = trabajo[col_trazado].astype(str)
                    trabajo["odometro"] = pd.to_numeric(trabajo[col_odo], errors="coerce")
                    trabajo["velocidad"] = pd.to_numeric(trabajo[col_vel], errors="coerce")
                    trabajo["soc"] = pd.to_numeric(trabajo[col_soc], errors="coerce")
                    trabajo["lat"] = pd.to_numeric(trabajo[col_lat], errors="coerce")
                    trabajo["lon"] = pd.to_numeric(trabajo[col_lon], errors="coerce")
                    trabajo["altitud"] = pd.to_numeric(trabajo[col_alt], errors="coerce")

                    if col_energia:
                        trabajo["energia_viaje"] = pd.to_numeric(trabajo[col_energia], errors="coerce")
                    else:
                        trabajo["energia_viaje"] = None

                    if col_consumo_kw:
                        trabajo["consumo_kw"] = pd.to_numeric(trabajo[col_consumo_kw], errors="coerce")
                    else:
                        trabajo["consumo_kw"] = None

                    trabajo = trabajo.dropna(subset=["trazado", "odometro"]).copy()

                    trazados = sorted(trabajo["trazado"].dropna().unique().tolist())

                    st.markdown("### Selección de trazado")
                    trazado_sel = st.selectbox("Trazado", trazados)

                    vista = trabajo[trabajo["trazado"] == trazado_sel].copy().sort_values("odometro")

                    if vista.empty:
                        st.warning("No hay datos para el trazado seleccionado.")
                    else:
                        # -----------------------------
                        # DISTANCIA RECORRIDA
                        # -----------------------------
                        kms_recorridos = float(vista["odometro"].max() - vista["odometro"].min())

                        # -----------------------------
                        # RESUMEN POR TRAZADO DESDE HOJA2
                        # -----------------------------
                        rendimiento = None
                        autonomia = None
                        distancia_resumen = None
                        pasajeros = None

                        if df_resumen_excel is not None and "Trazado" in df_resumen_excel.columns:
                            resumen_sel = df_resumen_excel[df_resumen_excel["Trazado"].astype(str) == trazado_sel]

                            if not resumen_sel.empty:
                                fila = resumen_sel.iloc[0]

                                if "Distancia" in resumen_sel.columns:
                                    distancia_resumen = pd.to_numeric(fila["Distancia"], errors="coerce")

                                if "Autonomia" in resumen_sel.columns:
                                    autonomia = pd.to_numeric(fila["Autonomia"], errors="coerce")

                                if "Rendimiento Prueba" in resumen_sel.columns:
                                    rendimiento = pd.to_numeric(fila["Rendimiento Prueba"], errors="coerce")

                                if "Cantidad de  pasajeros" in resumen_sel.columns:
                                    pasajeros = str(fila["Cantidad de  pasajeros"])

                        # fallback si faltan datos resumen
                        if rendimiento is None or pd.isna(rendimiento):
                            if vista["consumo_kw"].notna().any() and kms_recorridos > 0:
                                energia_total = vista["consumo_kw"].fillna(0).sum() / 60.0
                                rendimiento = energia_total / kms_recorridos if kms_recorridos > 0 else None

                        if distancia_resumen is None or pd.isna(distancia_resumen):
                            distancia_resumen = kms_recorridos

                        velocidad_prom = vista["velocidad"].mean() if vista["velocidad"].notna().any() else None
                        soc_inicio = vista["soc"].iloc[0] if vista["soc"].notna().any() else None
                        soc_fin = vista["soc"].iloc[-1] if vista["soc"].notna().any() else None

                        # -----------------------------
                        # KPIS
                        # -----------------------------
                        k1, k2, k3, k4 = st.columns(4)

                        k1.metric(
                            "Rendimiento kWh/km",
                            f"{rendimiento:.3f}" if rendimiento is not None and not pd.isna(rendimiento) else "Sin dato"
                        )
                        k2.metric(
                            "Autonomía proyectada",
                            f"{autonomia:.0f} km" if autonomia is not None and not pd.isna(autonomia) else "Sin dato"
                        )
                        k3.metric(
                            "Velocidad promedio",
                            f"{velocidad_prom:.1f} km/h" if velocidad_prom is not None and not pd.isna(velocidad_prom) else "Sin dato"
                        )
                        k4.metric(
                            "Kms recorridos",
                            f"{distancia_resumen:.1f} km" if distancia_resumen is not None and not pd.isna(distancia_resumen) else f"{kms_recorridos:.1f} km"
                        )

                        if pasajeros:
                            st.caption(f"Pasajeros: {pasajeros}")

                        # -----------------------------
                        # GRAFICO 1: VELOCIDAD + SOC
                        # -----------------------------
                        g1, g2 = st.columns([1.1, 1])

                        with g1:
                            st.markdown("### Velocidad y Estado de Carga")

                            base = alt.Chart(vista).encode(
                                x=alt.X("odometro:Q", title="Odómetro km")
                            )

                            line_vel = base.mark_line(point=True).encode(
                                y=alt.Y("velocidad:Q", title="Velocidad km/h"),
                                tooltip=["odometro", "velocidad", "soc"]
                            )

                            line_soc = base.mark_line(point=True).encode(
                                y=alt.Y("soc:Q", title="Estado de Carga SoC")
                            )

                            chart = alt.layer(line_vel, line_soc).resolve_scale(
                                y="independent"
                            ).properties(height=340)

                            st.altair_chart(chart, use_container_width=True)

                        # -----------------------------
                        # GRAFICO 2: MAPA
                        # -----------------------------
                        with g2:
                            st.markdown("### Trazado")
                            mapa = vista.dropna(subset=["lat", "lon"]).copy()
                            if not mapa.empty:
                                st.map(
                                    mapa.rename(columns={"lat": "latitude", "lon": "longitude"})[
                                        ["latitude", "longitude"]
                                    ]
                                )
                            else:
                                st.info("No hay coordenadas válidas para el trazado seleccionado.")

                        # -----------------------------
                        # GRAFICO 3: ALTURA VS ODOMETRO
                        # -----------------------------
                        st.markdown("### Perfil de altura")
                        if vista["altitud"].notna().any():
                            chart_alt = alt.Chart(vista).mark_area(opacity=0.6).encode(
                                x=alt.X("odometro:Q", title="Odómetro"),
                                y=alt.Y("altitud:Q", title="Altura (m)"),
                                tooltip=["odometro", "altitud"]
                            ).properties(height=320)
                            st.altair_chart(chart_alt, use_container_width=True)
                        else:
                            st.info("No hay datos de altitud.")

                        # -----------------------------
                        # TABLA DETALLE
                        # -----------------------------
                        st.markdown("### Detalle del trazado")
                        columnas_detalle = [
                            c for c in [
                                "fecha_evento", "trazado", "odometro", "velocidad",
                                "soc", "altitud", "lat", "lon", "energia_viaje", "consumo_kw"
                            ] if c in vista.columns
                        ]
                        st.dataframe(vista[columnas_detalle], use_container_width=True)

        except Exception as e:
            st.error(f"Error al procesar el archivo: {e}")

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