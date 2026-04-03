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


def sugerir_columna(columnas, candidatos):
    columnas_lower = {c.lower(): c for c in columnas}
    for cand in candidatos:
        for col_lower, col_real in columnas_lower.items():
            if cand in col_lower:
                return col_real
    return None


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

    if "numero" in columnas:
        cur = conn.cursor()
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
            f"No se encontró la plantilla '{TEMPLATE_FILE}'. Debes dejarla en la misma carpeta del proyecto."
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
        return f.read()


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
# TAB 3 - EFICIENCIA ADAPTADA OSORNO
# =========================================================
with tab_efi:
    st.subheader("Eficiencia energética")
    st.write("Sube una planilla Excel y selecciona las columnas. La app intentará sugerir automáticamente columnas parecidas al caso Osorno.")

    archivo = st.file_uploader("Subir Excel", type=["xlsx", "xls"], key="excel_efi")

    if archivo:
        try:
            df = leer_excel_seguro(archivo)
            st.dataframe(df.head(10), use_container_width=True)

            columnas = list(df.columns)

            sug_trazado = sugerir_columna(columnas, ["trazado", "ruta", "servicio"])
            sug_odo = sugerir_columna(columnas, ["odometro", "odómetro", "odom", "km"])
            sug_vel = sugerir_columna(columnas, ["velocidad", "speed"])
            sug_soc = sugerir_columna(columnas, ["soc", "estado de carga", "carga"])
            sug_alt = sugerir_columna(columnas, ["altura", "altitud", "elev"])
            sug_lat = sugerir_columna(columnas, ["lat"])
            sug_lon = sugerir_columna(columnas, ["lon", "lng", "long"])

            st.markdown("### Mapeo de columnas")

            m1, m2, m3 = st.columns(3)

            with m1:
                col_trazado = st.selectbox(
                    "Columna trazado / ruta",
                    columnas,
                    index=columnas.index(sug_trazado) if sug_trazado in columnas else 0
                )
                col_odo = st.selectbox(
                    "Columna odómetro / km acumulado",
                    columnas,
                    index=columnas.index(sug_odo) if sug_odo in columnas else 0
                )

            with m2:
                col_vel = st.selectbox(
                    "Columna velocidad",
                    ["(No usar)"] + columnas,
                    index=(["(No usar)"] + columnas).index(sug_vel) if sug_vel in columnas else 0
                )
                col_soc = st.selectbox(
                    "Columna SoC / estado de carga",
                    ["(No usar)"] + columnas,
                    index=(["(No usar)"] + columnas).index(sug_soc) if sug_soc in columnas else 0
                )

            with m3:
                col_alt = st.selectbox(
                    "Columna altura / altitud",
                    ["(No usar)"] + columnas,
                    index=(["(No usar)"] + columnas).index(sug_alt) if sug_alt in columnas else 0
                )
                col_lat = st.selectbox(
                    "Columna latitud",
                    ["(No usar)"] + columnas,
                    index=(["(No usar)"] + columnas).index(sug_lat) if sug_lat in columnas else 0
                )
                col_lon = st.selectbox(
                    "Columna longitud",
                    ["(No usar)"] + columnas,
                    index=(["(No usar)"] + columnas).index(sug_lon) if sug_lon in columnas else 0
                )

            st.markdown("### Parámetros")

            p1, p2, p3, p4 = st.columns(4)

            with p1:
                bateria_kwh = st.selectbox("Batería", [231.8, 255.0], index=1)
            with p2:
                reserva_pct = st.slider("Reserva batería (%)", 0, 30, 10)
            with p3:
                rendimiento_objetivo = st.number_input("Objetivo kWh/km", min_value=0.0, value=0.95, step=0.01, format="%.2f")
            with p4:
                energia_total_trazado = st.number_input("Energía total del trazado (kWh)", min_value=0.0, value=10.45, step=0.01)

            if st.button("Calcular rendimiento", use_container_width=True):
                trabajo = df.copy()
                trabajo["trazado"] = trabajo[col_trazado].astype(str)
                trabajo["odometro"] = pd.to_numeric(trabajo[col_odo], errors="coerce")

                if col_vel != "(No usar)":
                    trabajo["velocidad"] = pd.to_numeric(trabajo[col_vel], errors="coerce")
                else:
                    trabajo["velocidad"] = None

                if col_soc != "(No usar)":
                    trabajo["soc"] = pd.to_numeric(trabajo[col_soc], errors="coerce")
                else:
                    trabajo["soc"] = None

                if col_alt != "(No usar)":
                    trabajo["altitud"] = pd.to_numeric(trabajo[col_alt], errors="coerce")
                else:
                    trabajo["altitud"] = None

                if col_lat != "(No usar)" and col_lon != "(No usar)":
                    trabajo["lat"] = pd.to_numeric(trabajo[col_lat], errors="coerce")
                    trabajo["lon"] = pd.to_numeric(trabajo[col_lon], errors="coerce")
                else:
                    trabajo["lat"] = None
                    trabajo["lon"] = None

                trabajo = trabajo.dropna(subset=["odometro"]).copy()

                grupos = []
                for _, g in trabajo.groupby("trazado", dropna=False):
                    g = g.copy()
                    g["distancia_parcial"] = g["odometro"].diff().abs().fillna(0)
                    grupos.append(g)

                trabajo = pd.concat(grupos, ignore_index=True)
                distancia_total = trabajo["distancia_parcial"].sum()

                if distancia_total <= 0:
                    st.error("No fue posible calcular distancia. Revisa la columna de odómetro.")
                else:
                    capacidad_disponible = bateria_kwh * (1 - reserva_pct / 100)
                    kwh_km = energia_total_trazado / distancia_total
                    km_kwh = distancia_total / energia_total_trazado if energia_total_trazado > 0 else None
                    autonomia = capacidad_disponible / kwh_km if kwh_km > 0 else None

                    trabajo["kwh_km_ref"] = kwh_km

                    st.markdown("### Indicadores principales")
                    k1, k2, k3, k4 = st.columns(4)

                    k1.metric("Rendimiento kWh/km", f"{kwh_km:.3f}")
                    k2.metric("Autonomía proyectada", f"{autonomia:.0f} km" if autonomia else "Sin dato")
                    k3.metric("Kms recorridos", f"{distancia_total:.1f} km")
                    k4.metric(
                        "Velocidad promedio",
                        f"{trabajo['velocidad'].mean():.1f} km/h" if col_vel != "(No usar)" and trabajo["velocidad"].notna().any() else "Sin dato"
                    )

                    trazados = sorted(trabajo["trazado"].dropna().unique().tolist())
                    trazado_sel = st.selectbox("Trazado a visualizar", trazados)
                    vista = trabajo[trabajo["trazado"] == trazado_sel].copy()

                    g1, g2 = st.columns([1.1, 1])

                    with g1:
                        st.markdown("### Velocidad / SoC sobre odómetro")
                        df_chart = pd.DataFrame(index=vista["odometro"])
                        if col_vel != "(No usar)" and vista["velocidad"].notna().any():
                            df_chart["Velocidad"] = vista["velocidad"]
                        if col_soc != "(No usar)" and vista["soc"].notna().any():
                            df_chart["SoC"] = vista["soc"]
                        if not df_chart.empty:
                            st.line_chart(df_chart)
                        else:
                            st.info("No hay columnas de velocidad o SoC válidas para graficar.")

                    with g2:
                        st.markdown("### Trazado")
                        if col_lat != "(No usar)" and col_lon != "(No usar)":
                            mapa = vista.dropna(subset=["lat", "lon"]).copy()
                            if not mapa.empty:
                                st.map(mapa.rename(columns={"lat": "latitude", "lon": "longitude"})[["latitude", "longitude"]])
                            else:
                                st.info("No hay coordenadas válidas para el trazado seleccionado.")
                        else:
                            st.info("Selecciona latitud y longitud para mostrar el mapa.")

                    if col_alt != "(No usar)" and vista["altitud"].notna().any():
                        st.markdown("### Perfil de altura")
                        st.area_chart(pd.DataFrame({"Altura (m)": vista["altitud"].values}, index=vista["odometro"].values))

                    st.markdown("### Resumen por trazado")
                    resumen = trabajo.groupby("trazado", dropna=False).agg(
                        odometro_inicial=("odometro", "min"),
                        odometro_final=("odometro", "max"),
                        kms_recorridos=("distancia_parcial", "sum"),
                        velocidad_promedio=("velocidad", "mean") if col_vel != "(No usar)" else ("odometro", "count")
                    ).reset_index()

                    if col_vel == "(No usar)":
                        resumen = resumen.drop(columns=["velocidad_promedio"])

                    resumen["kwh_km"] = kwh_km
                    resumen["km_kwh"] = km_kwh if km_kwh is not None else 0
                    resumen["autonomia_proyectada_km"] = autonomia if autonomia is not None else 0
                    resumen["semaforo"] = resumen["kwh_km"].apply(lambda x: semaforo_kwh(x, rendimiento_objetivo))

                    st.dataframe(resumen, use_container_width=True)

                    detalle = trabajo.copy()

                    export_excel = exportar_resultados_excel(
                        detalle,
                        resumen,
                        trabajo[["trazado", "lat", "lon"]].dropna() if "lat" in trabajo.columns and "lon" in trabajo.columns else None
                    )

                    st.download_button(
                        "Descargar resultados Excel",
                        data=export_excel,
                        file_name=f"Rendimiento_energetico_{fecha_corta(date.today())}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

        except Exception as e:
            st.error(f"Error al procesar el archivo: {e}")