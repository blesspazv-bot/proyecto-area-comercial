import io
from datetime import date

import streamlit as st
from docx import Document

st.set_page_config(page_title="APP Área Comercial Buses y Vans", layout="wide")

COTIZANTES = {
    "Diego Vejar": "DV",
    "Sergio Silva": "SS",
    "Alvaro Correa": "AC",
    "Fabian Orellana": "FO",
    "Rodrigo Sepulveda": "RS",
}

FIRMAS = {
    "Diego Vejar": {
        "nombre": "Diego Vejar",
        "cargo": "Ejecutivo Comercial",
        "correo": "",
        "telefono": ""
    },
    "Sergio Silva": {
        "nombre": "Sergio Silva",
        "cargo": "Ejecutivo Comercial",
        "correo": "",
        "telefono": ""
    },
    "Alvaro Correa": {
        "nombre": "Alvaro Correa",
        "cargo": "Ejecutivo Comercial",
        "correo": "",
        "telefono": ""
    },
    "Fabian Orellana": {
        "nombre": "Fabian Orellana",
        "cargo": "Ejecutivo Comercial",
        "correo": "",
        "telefono": ""
    },
    "Rodrigo Sepulveda": {
        "nombre": "Rodrigo Sepulveda Toepfer",
        "cargo": "Gerente de Buses y Vans (Iveco)",
        "correo": "rsepulveda@andesmotor.cl",
        "telefono": "227202221 / 979783254"
    },
}

TEMPLATE_PATH = "Propuesta_FOTON_Electrico_U9.docx"


def format_usd(valor: float) -> str:
    return f"USD$ {valor:,.0f}".replace(",", ".")


def numero_cotizacion(cotizante: str, correlativo: int) -> str:
    prefijo = COTIZANTES[cotizante]
    return f"{prefijo}-{correlativo:02d}"


def reemplazar_en_parrafo(paragraph, reemplazos):
    texto = "".join(run.text for run in paragraph.runs)
    nuevo = texto
    for k, v in reemplazos.items():
        nuevo = nuevo.replace(k, v)
    if nuevo != texto:
        if paragraph.runs:
            paragraph.runs[0].text = nuevo
            for run in paragraph.runs[1:]:
                run.text = ""
        else:
            paragraph.add_run(nuevo)


def reemplazar_en_doc(doc, reemplazos):
    for p in doc.paragraphs:
        reemplazar_en_parrafo(p, reemplazos)

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    reemplazar_en_parrafo(p, reemplazos)


def generar_docx(datos):
    doc = Document(TEMPLATE_PATH)

    firma = FIRMAS[datos["cotizante"]]

    reemplazos = {
        "Santiago, miércoles 17 de diciembre de 2025": datos["fecha"],
        "Cotización n° RS-03": f'Cotización n° {datos["numero_cotizacion"]}',
        "Conecta": datos["cliente"],
        "1 Buses Eléctricos, Marca FOTON, Modelo U9 de carrocería monocasco Urbano estándar RED, año 2026.": (
            f'{datos["cantidad_unidades"]} Buses Eléctricos, Marca FOTON, Modelo U9 de carrocería monocasco Urbano estándar RED, año 2026.'
        ),
        "· Valor unitario por bus: USD$  130.491.-  + IVA.": (
            f'· Valor unitario por bus: {datos["monto"]} + IVA.'
        ),
        "· Valor cuota flota por unidad: . + IVA por 96 meses, con una tasa leasing de %.": (
            f'· Valor total negocio: {datos["total_negocio"]} + IVA.'
        ),
        "48 meses": datos["contrato_mantto"],
        "Rodrigo Sepúlveda Toepfer": firma["nombre"],
        "Gerente de Buses y Vans (Iveco)": firma["cargo"],
        "rsepulveda@andesmotor.cl": firma["correo"],
        "T: 227202221": f'T: {firma["telefono"]}' if firma["telefono"] else "T:",
    }

    reemplazar_en_doc(doc, reemplazos)

    salida = io.BytesIO()
    doc.save(salida)
    salida.seek(0)
    return salida


st.title("APP Área Comercial Buses y Vans")
st.subheader("Generador de Cotización FOTON U9")

with st.form("form_cotizacion"):
    col1, col2 = st.columns(2)

    with col1:
        fecha = st.date_input("Fecha", value=date.today())
        cliente = st.text_input("Cliente")
        cotizante = st.selectbox("Cotizante", list(COTIZANTES.keys()))
        correlativo = st.number_input("Correlativo", min_value=1, value=1, step=1)

    with col2:
        monto = st.number_input("Monto unitario USD", min_value=0.0, value=130491.0, step=1000.0)
        contrato_mantto = st.text_input("Contrato mantto", value="48 meses")
        cantidad_unidades = st.number_input("Cantidad de unidades", min_value=1, value=1, step=1)

    total_negocio = monto * cantidad_unidades

    submitted = st.form_submit_button("Generar cotización")

if submitted:
    try:
        num_cot = numero_cotizacion(cotizante, correlativo)

        datos = {
            "fecha": fecha.strftime("%d-%m-%Y"),
            "cliente": cliente.strip() if cliente.strip() else "Cliente",
            "cotizante": cotizante,
            "numero_cotizacion": num_cot,
            "monto": format_usd(monto),
            "contrato_mantto": contrato_mantto,
            "cantidad_unidades": str(cantidad_unidades),
            "total_negocio": format_usd(total_negocio),
        }

        archivo = generar_docx(datos)

        st.success(f"Cotización {num_cot} generada correctamente.")
        st.write(f"Total negocio: {format_usd(total_negocio)}")

        st.download_button(
            label="Descargar cotización Word",
            data=archivo,
            file_name=f"Cotizacion_{num_cot}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )

    except Exception as e:
        st.error(f"Error al generar la cotización: {e}")