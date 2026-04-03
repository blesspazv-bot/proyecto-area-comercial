import io
from datetime import date
import streamlit as st
from docx import Document

st.set_page_config(page_title="APP Área Comercial Buses y Vans", layout="wide")

# ================================
# CONFIG COTIZANTES
# ================================
COTIZANTES = {
    "Diego Vejar": {
        "prefijo": "DV",
        "cargo": "Subgerente Comercial Buses y Vans",
        "correo": "dvejar@andesmotor.cl",
        "telefono": "981774604"
    },
    "Sergio Silva": {
        "prefijo": "SS",
        "cargo": "Ejecutivo de ventas Zona Sur Buses y Vans",
        "correo": "sergio.silva@andesmotor.cl",
        "telefono": ""
    },
    "Alvaro Correa": {
        "prefijo": "AC",
        "cargo": "Ejecutivo de ventas Zona Norte Buses y Vans",
        "correo": "alvaro.correa@andesmotor.cl",
        "telefono": ""
    },
    "Fabian Orellana": {
        "prefijo": "FO",
        "cargo": "Product Manager Senior Zona Buses y Vans",
        "correo": "forellana@andesmotor.cl",
        "telefono": "989356449"
    },
    "Rodrigo Sepulveda": {
        "prefijo": "RS",
        "cargo": "Gerente de Buses y Vans",
        "correo": "rsepulveda@andesmotor.cl",
        "telefono": "979783254"
    },
}

TEMPLATE = "Propuesta_FOTON_Electrico_U9.docx"

# ================================
# FUNCIONES
# ================================
def fecha_larga(fecha):
    meses = ["enero","febrero","marzo","abril","mayo","junio","julio","agosto","septiembre","octubre","noviembre","diciembre"]
    dias = ["lunes","martes","miércoles","jueves","viernes","sábado","domingo"]
    return f"Santiago, {dias[fecha.weekday()]} {fecha.day} de {meses[fecha.month-1]} de {fecha.year}"

def formato_usd(valor):
    return f"USD$ {valor:,.0f}".replace(",", ".")

def reemplazar(doc, dic):
    for p in doc.paragraphs:
        for k, v in dic.items():
            if k in p.text:
                p.text = p.text.replace(k, v)

    for t in doc.tables:
        for r in t.rows:
            for c in r.cells:
                for p in c.paragraphs:
                    for k, v in dic.items():
                        if k in p.text:
                            p.text = p.text.replace(k, v)

def generar_docx(data):
    doc = Document(TEMPLATE)

    reemplazos = {
        "Santiago, miércoles 17 de diciembre de 2025": data["fecha"],
        "Cotización n° RS-03": f'Cotización n° {data["numero"]}',
        "Conecta": data["cliente"],
        "USD$  130.491.-": data["precio"],
        "1 Buses Eléctricos": f'{data["cantidad"]} Buses Eléctricos',
        "48 meses": data["mantto"],
        "Rodrigo Sepúlveda Toepfer": data["nombre"],
        "Gerente de Buses y Vans (Iveco)": data["cargo"],
        "rsepulveda@andesmotor.cl": data["correo"],
        "M: 979783254": f'M: {data["telefono"]}'
    }

    reemplazar(doc, reemplazos)

    # Reemplazo bloque mantenimiento completo
    for p in doc.paragraphs:
        if "La oferta incluye 48 meses" in p.text:
            p.text = data["texto_mantto"]

    output = io.BytesIO()
    doc.save(output)
    output.seek(0)
    return output

# ================================
# UI
# ================================
st.title("Generador Cotización FOTON U9")

col1, col2 = st.columns(2)

with col1:
    cliente = st.text_input("Cliente", value="Juan Perez")
    fecha = st.date_input("Fecha", value=date.today())
    cotizante = st.selectbox("Cotizante", list(COTIZANTES.keys()))
    correlativo = st.number_input("Correlativo", value=1)

with col2:
    precio = st.number_input("Precio unitario USD", value=130491)
    cantidad = st.number_input("Cantidad unidades", value=1)
    mantto = st.text_input("Contrato mantto", value="48 meses")

texto_mantto = st.text_area(
    "Texto mantenimiento (editable)",
    """I. La oferta incluye 48 meses de mantenimiento preventivo y correctivo...
II. Adicionalmente se entregarán USD 1.500...
III. La oferta incluye 8 años de telemetría sin costo para el cliente."""
)

# ================================
# LOGICA
# ================================
data_cot = COTIZANTES[cotizante]
numero = f'{data_cot["prefijo"]}-{int(correlativo):02d}'

st.write("### Vista previa")
st.write("Número:", numero)
st.write("Total:", formato_usd(precio * cantidad))

if st.button("Generar cotización"):

    data = {
        "fecha": fecha_larga(fecha),
        "cliente": cliente,
        "numero": numero,
        "precio": formato_usd(precio),
        "cantidad": cantidad,
        "mantto": mantto,
        "texto_mantto": texto_mantto,
        "nombre": cotizante,
        "cargo": data_cot["cargo"],
        "correo": data_cot["correo"],
        "telefono": data_cot["telefono"]
    }

    archivo = generar_docx(data)

    st.download_button(
        "Descargar Cotización",
        archivo,
        file_name=f"Cotizacion_{numero}.docx"
    )