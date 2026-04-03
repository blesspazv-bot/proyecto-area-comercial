import os
import sqlite3
import tempfile
from datetime import date

import pandas as pd
import streamlit as st
from docxtpl import DocxTemplate, RichText

st.set_page_config(page_title="APP Área Comercial Buses y Vans", layout="wide")

# =========================================================
# CONFIGURACION GENERAL
# =========================================================
TEMPLATE_FILE = "plantilla_cotizacion_foton_u9.docx"
DB_FILE = "cotizaciones.db"

COTIZANTES = {
    "Diego Vejar": {
        "prefijo": "DV",
        "firma_nombre": "Diego Vejar",
        "firma_cargo": "Subgerente Comercial Buses y Vans(IVECO)",
        "firma_correo": "dvejar@andesmotor.cl",
        "firma_telefono": "981774604",
    },
    "Sergio Silva": {
        "prefijo": "SS",
        "firma_nombre": "Sergio Silva",
        "firma_cargo": "Ejecutivo de ventas Zona Sur Buses y Vans(IVECO)",
        "firma_correo": "sergio.silva@andesmotor.cl",
        "firma_telefono": "",
    },
    "Alvaro Correa": {
        "prefijo": "AC",
        "firma_nombre": "Alvaro Correa",
        "firma_cargo": "Ejecutivo de ventas Zona Norte Buses y Vans(IVECO)",
        "firma_correo": "alvaro.correa@andesmotor.cl",
        "firma_telefono": "",
    },
    "Fabian Orellana": {
        "prefijo": "FO",
        "firma_nombre": "Fabian Orellana",
        "firma_cargo": "Product Manager Senior Buses y Vans(IVECO)",
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
# UTILIDADES GENERALES
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
    lineas = texto.split("\n")
    for i, linea in enumerate(lineas):
        rt.add(linea, size=11)
        if i < len(lineas) - 1:
            rt.add("\n")
    return rt


# =========================================================
# BASE DE DATOS COTIZACIONES
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
            capacidad_bateria TEXT,
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
            capacidad_bateria,
            creado_en
        FROM cotizaciones
        ORDER BY id DESC
    """, conn)
    conn.close()
    return df


# =========================================================
# GENERACION DOCX
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


# =========================================================
# EXPORTAR RESULTADOS EXCEL
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
# INICIO APP
# =========================================================
init_db()

st.title("APP Área Comercial Buses y Vans")

tab_cot, tab_hist, tab_efi = st.tabs([
    "Nueva cotización",
    "Historial cotizaciones",
    "Eficiencia energética"
])

# =========================================================
# TAB 1 - COTIZACIONES
# =========================================================
with tab_cot:
    st.subheader("Generador de Cotización FOTON U9")

    col1, col2 = st.columns(2)

    with col1:
        fecha = st.date_input("Fecha", value=date.today())
        cliente = st.text_input("Cliente", value="")
        cotizante = st.selectbox("Cotizante", list(COTIZANTES.keys()))

    with col2:
        cantidad_unidades = st.number_input("Cantidad de unidades", min_value=1, value=1, step=1)
        precio_unitario = st.number_input("Precio unitario USD", min_value=0.0, value=130491.0, step=1000.0)
        

    capacidad_bateria = st.selectbox(
        "Capacidad nominal batería",
        ["231,8 kWh", "255 kWh"]
    )

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
        height=180
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

            nombre_archivo = (
                f"Propuesta_foton_U9_"
                f"{limpiar_nombre_archivo(cliente.strip())}_"
                f"{limpiar_nombre_archivo(capacidad_bateria)}_"
                f"{fecha_corta(fecha)}"
            )

            contexto = {
                "fecha_larga": fecha_larga_es(fecha),
                "fecha_corta": fecha_corta(fecha),
                "numero_cotizacion": numero_cotizacion,
                "cliente": cliente.strip(),
                "cantidad_unidades": cantidad_unidades,
                "precio_unitario": usd_fmt(precio_unitario),
                "capacidad_bateria": capacidad_bateria,
                "texto_mantto": formatear_texto_mantto(texto_mantto),
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
                "texto_mantto": texto_mantto,
                "capacidad_bateria": capacidad_bateria,
            }

            try:
                archivo_path = generar_docx(contexto, nombre_archivo)
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
                    file_name=f"{nombre_archivo}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )

            except Exception as e:
                st.error(f"Error al generar la cotización: {e}")


# =========================================================
# TAB 2 - HISTORIAL
# =========================================================
with tab_hist:
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


# =========================================================
# TAB 3 - EFICIENCIA ENERGETICA
# =========================================================
with tab_efi:
    st.subheader("Rendimiento energético")
    st.write("Sube una planilla Excel `.xlsx` y selecciona las columnas reales del archivo.")

    archivo = st.file_uploader(
        "Subir archivo Excel",
        type=["xlsx"],
        key="archivo_eficiencia"
    )

    if archivo is not None:
        try:
            df_raw = pd.read_excel(archivo)

            st.markdown("### Vista previa")
            st.dataframe(df_raw.head(20), use_container_width=True)

            columnas = list(df_raw.columns)

            st.markdown("### Mapeo de columnas")
            c1, c2, c3 = st.columns(3)

            with c1:
                col_fecha = st.selectbox("Columna fecha", ["(No usar)"] + columnas, index=0)
                col_ruta = st.selectbox("Columna ruta / servicio", columnas)
                col_bus = st.selectbox("Columna bus / patente / VIN", ["(No usar)"] + columnas, index=0)

            with c2:
                col_energia = st.selectbox("Columna energía consumida (kWh)", columnas)
                col_distancia = st.selectbox("Columna distancia recorrida (km)", columnas)
                col_velocidad = st.selectbox("Columna velocidad promedio", ["(No usar)"] + columnas, index=0)

            with c3:
                col_lat = st.selectbox("Columna latitud", ["(No usar)"] + columnas, index=0)
                col_lon = st.selectbox("Columna longitud", ["(No usar)"] + columnas, index=0)
                col_altitud = st.selectbox("Columna altitud", ["(No usar)"] + columnas, index=0)

            st.markdown("### Parámetros de cálculo")
            p1, p2, p3, p4 = st.columns(4)

            with p1:
                capacidad_util = st.number_input(
                    "Capacidad útil batería (kWh)",
                    min_value=1.0,
                    value=255.0,
                    step=1.0
                )
            with p2:
                reserva_pct = st.slider(
                    "Reserva batería (%)",
                    min_value=0,
                    max_value=30,
                    value=10,
                    step=1
                )
            with p3:
                tarifa_energia = st.number_input(
                    "Tarifa energía CLP/kWh",
                    min_value=0.0,
                    value=180.0,
                    step=1.0
                )
            with p4:
                rendimiento_objetivo = st.number_input(
                    "Objetivo consumo (kWh/km)",
                    min_value=0.0,
                    value=1.20,
                    step=0.01,
                    format="%.2f"
                )

            if st.button("Calcular rendimiento energético", use_container_width=True):
                trabajo = df_raw.copy()

                trabajo["ruta_servicio"] = trabajo[col_ruta].astype(str)
                trabajo["energia_kwh"] = pd.to_numeric(trabajo[col_energia], errors="coerce")
                trabajo["distancia_km"] = pd.to_numeric(trabajo[col_distancia], errors="coerce")

                if col_fecha != "(No usar)":
                    trabajo["fecha"] = trabajo[col_fecha]
                else:
                    trabajo["fecha"] = ""

                if col_bus != "(No usar)":
                    trabajo["bus"] = trabajo[col_bus].astype(str)
                else:
                    trabajo["bus"] = ""

                if col_velocidad != "(No usar)":
                    trabajo["velocidad_prom"] = pd.to_numeric(trabajo[col_velocidad], errors="coerce")
                else:
                    trabajo["velocidad_prom"] = None

                if col_altitud != "(No usar)":
                    trabajo["altitud"] = pd.to_numeric(trabajo[col_altitud], errors="coerce")
                else:
                    trabajo["altitud"] = None

                if col_lat != "(No usar)" and col_lon != "(No usar)":
                    trabajo["lat"] = pd.to_numeric(trabajo[col_lat], errors="coerce")
                    trabajo["lon"] = pd.to_numeric(trabajo[col_lon], errors="coerce")
                else:
                    trabajo["lat"] = None
                    trabajo["lon"] = None

                trabajo = trabajo.dropna(subset=["energia_kwh", "distancia_km"])
                trabajo = trabajo[trabajo["distancia_km"] > 0].copy()

                if trabajo.empty:
                    st.error("No hay registros válidos para calcular.")
                else:
                    capacidad_disponible = capacidad_util * (1 - reserva_pct / 100)

                    trabajo["kwh_km"] = trabajo["energia_kwh"] / trabajo["distancia_km"]
                    trabajo["km_kwh"] = trabajo["distancia_km"] / trabajo["energia_kwh"]
                    trabajo["autonomia_estimada_km"] = capacidad_disponible / trabajo["kwh_km"]
                    trabajo["costo_energia_clp"] = trabajo["energia_kwh"] * tarifa_energia
                    trabajo["costo_km_clp"] = trabajo["costo_energia_clp"] / trabajo["distancia_km"]
                    trabajo["desviacion_vs_objetivo_pct"] = (
                        (trabajo["kwh_km"] - rendimiento_objetivo) / rendimiento_objetivo
                    ) * 100

                    trabajo["estado_rendimiento"] = trabajo["kwh_km"].apply(
                        lambda x: "Sobre objetivo" if x > rendimiento_objetivo else "Dentro objetivo"
                    )

                    columnas_detalle = [
                        "fecha",
                        "ruta_servicio",
                        "bus",
                        "energia_kwh",
                        "distancia_km",
                        "kwh_km",
                        "km_kwh",
                        "autonomia_estimada_km",
                        "costo_energia_clp",
                        "costo_km_clp",
                        "desviacion_vs_objetivo_pct",
                        "estado_rendimiento",
                    ]

                    if col_velocidad != "(No usar)":
                        columnas_detalle.append("velocidad_prom")
                    if col_altitud != "(No usar)":
                        columnas_detalle.append("altitud")
                    if col_lat != "(No usar)" and col_lon != "(No usar)":
                        columnas_detalle.extend(["lat", "lon"])

                    df_detalle = trabajo[columnas_detalle].copy()

                    agg_dict = {
                        "energia_kwh": "sum",
                        "distancia_km": "sum",
                        "kwh_km": "mean",
                        "km_kwh": "mean",
                        "autonomia_estimada_km": "mean",
                        "costo_energia_clp": "sum",
                        "costo_km_clp": "mean",
                        "desviacion_vs_objetivo_pct": "mean",
                    }

                    if col_velocidad != "(No usar)":
                        agg_dict["velocidad_prom"] = "mean"

                    df_resumen = (
                        trabajo.groupby("ruta_servicio", dropna=False)
                        .agg(agg_dict)
                        .reset_index()
                    )

                    st.markdown("### Indicadores generales")
                    m1, m2, m3, m4 = st.columns(4)
                    m1.metric("Consumo promedio", f'{trabajo["kwh_km"].mean():.2f} kWh/km')
                    m2.metric("Rendimiento promedio", f'{trabajo["km_kwh"].mean():.2f} km/kWh')
                    m3.metric("Autonomía promedio", f'{trabajo["autonomia_estimada_km"].mean():.0f} km')
                    m4.metric("Costo promedio por km", clp_fmt(trabajo["costo_km_clp"].mean()))

                    s1, s2, s3 = st.columns(3)
                    s1.metric("Energía total consumida", f'{trabajo["energia_kwh"].sum():,.1f} kWh')
                    s2.metric("Distancia total", f'{trabajo["distancia_km"].sum():,.1f} km')
                    s3.metric("Costo total energía", clp_fmt(trabajo["costo_energia_clp"].sum()))

                    st.markdown("### Detalle calculado")
                    st.dataframe(df_detalle, use_container_width=True)

                    st.markdown("### Resumen por ruta / servicio")
                    st.dataframe(df_resumen, use_container_width=True)

                    df_mapa = pd.DataFrame()
                    if col_lat != "(No usar)" and col_lon != "(No usar)":
                        df_mapa = trabajo.dropna(subset=["lat", "lon"]).copy()
                        if not df_mapa.empty:
                            st.markdown("### Mapa de coordenadas")
                            st.map(
                                df_mapa.rename(columns={"lat": "latitude", "lon": "longitude"})[
                                    ["latitude", "longitude"]
                                ]
                            )

                    excel_data = exportar_resultados_excel(df_detalle, df_resumen, df_mapa)

                    st.download_button(
                        "Descargar resultados Excel",
                        data=excel_data,
                        file_name=f"Rendimiento_energetico_{date.today().strftime('%Y%m%d')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

        except Exception as e:
            st.error(f"Error al procesar el archivo: {e}")