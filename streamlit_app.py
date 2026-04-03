import io
from datetime import date
import streamlit as st
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH

st.set_page_config(page_title="APP Área Comercial Buses y Vans", layout="wide")

TEMPLATE_FILE = "Propuesta_FOTON_Electrico_U9.docx"

COTIZANTES = {
    "Diego Vejar": {
        "prefijo": "DV",
        "nombre_firma": "Diego Vejar",
        "cargo": "Subgerente Comercial Buses y Vans",
        "correo": "dvejar@andesmotor.cl",
        "telefono_fijo": "",
        "telefono_movil": "981774604",
    },
    "Sergio Silva": {
        "prefijo": "SS",
        "nombre_firma": "Sergio Silva",
        "cargo": "Ejecutivo de ventas Zona Sur Buses y Vans",
        "correo": "sergio.silva@andesmotor.cl",
        "telefono_fijo": "",
        "telefono_movil": "",
    },
    "Alvaro Correa": {
        "prefijo": "AC",
        "nombre_firma": "Alvaro Correa",
        "cargo": "Ejecutivo de ventas Zona Norte Buses y Vans",
        "correo": "alvaro.correa@andesmotor.cl",
        "telefono_fijo": "",
        "telefono_movil": "",
    },
    "Fabian Orellana": {
        "prefijo": "FO",
        "nombre_firma": "Fabian Orellana",
        "cargo": "Product Manager Senior Zona Buses y Vans",
        "correo": "forellana@andesmotor.cl",
        "telefono_fijo": "",
        "telefono_movil": "989356449",
    },
    "Rodrigo Sepulveda": {
        "prefijo": "RS",
        "nombre_firma": "Rodrigo Sepulveda",
        "cargo": "Gerente de Buses y Vans",
        "correo": "rsepulveda@andesmotor.cl",
        "telefono_fijo": "",
        "telefono_movil": "979783254",
    },
}


def fecha_larga_es(fecha: date) -> str:
    dias = ["lunes", "martes", "miércoles", "jueves", "viernes", "sábado", "domingo"]
    meses = [
        "enero", "febrero", "marzo", "abril", "mayo", "junio",
        "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre"
    ]
    return f"Santiago, {dias[fecha.weekday()]} {fecha.day} de {meses[fecha.month-1]} de {fecha.year}"


def usd_fmt(valor: float) -> str:
    return f"USD$ {valor:,.0f}".replace(",", ".")


def limpiar_parrafo(paragraph):
    p = paragraph._element
    for child in list(p):
        p.remove(child)


def escribir_parrafo_simple(paragraph, texto, bold=False, align=None, font_size=Pt(11)):
    limpiar_parrafo(paragraph)
    run = paragraph.add_run(texto)
    run.bold = bold
    run.font.size = font_size
    if align is not None:
        paragraph.alignment = align


def escribir_parrafo_precio(paragraph, precio_texto):
    limpiar_parrafo(paragraph)
    r1 = paragraph.add_run("•    Valor unitario por bus: ")
    r1.font.size = Pt(11)

    r2 = paragraph.add_run(precio_texto)
    r2.bold = True
    r2.font.size = Pt(11)

    r3 = paragraph.add_run(" + IVA.")
    r3.font.size = Pt(11)


def reemplazar_en_parrafos(doc, datos):
    cotizante_info = COTIZANTES[datos["cotizante"]]

    for p in doc.paragraphs:
        txt = p.text.strip()

        # Fecha
        if txt.startswith("Santiago,"):
            escribir_parrafo_simple(p, datos["fecha_larga"])

        # Número de cotización
        elif "Cotización n°" in txt or "Cotización n°" in txt.replace("º", "°"):
            escribir_parrafo_simple(p, f'Cotización n° {datos["numero_cotizacion"]}')

        # Cliente
        elif txt == "Conecta" or txt == datos.get("cliente_original", ""):
            escribir_parrafo_simple(p, datos["cliente"])

        # Cantidad de buses / descripción
        elif txt.startswith("1) DESCRIPCIÓN:"):
            nuevo = (
                f'1) DESCRIPCIÓN: {datos["cantidad_unidades"]} Buses Eléctricos, '
                f'Marca FOTON, Modelo U9 de carrocería monocasco Urbano estándar RED, año 2026.'
            )
            escribir_parrafo_simple(p, nuevo)

        # Precio
        elif "Valor unitario por bus:" in txt:
            escribir_parrafo_precio(p, datos["precio_unitario"])

        # Bloque mantenimiento original: se borra
        elif txt.startswith("I.") or txt.startswith("II.") or txt.startswith("III."):
            escribir_parrafo_simple(p, "")

        # Firma nombre
        elif "Rodrigo Sepúlveda Toepfer" in txt or "Rodrigo Sepulveda Toepfer" in txt:
            escribir_parrafo_simple(p, cotizante_info["nombre_firma"])

        # Firma cargo
        elif "Gerente de Buses y Vans" in txt:
            escribir_parrafo_simple(p, cotizante_info["cargo"])

        # Firma correo
        elif "@andesmotor.cl" in txt:
            escribir_parrafo_simple(p, cotizante_info["correo"])

        # Firma teléfono fijo
        elif txt.startswith("T:"):
            valor = f'T: {cotizante_info["telefono_fijo"]}' if cotizante_info["telefono_fijo"] else ""
            escribir_parrafo_simple(p, valor)

        # Firma teléfono móvil
        elif txt.startswith("M:"):
            valor = f'M: {cotizante_info["telefono_movil"]}' if cotizante_info["telefono_movil"] else ""
            escribir_parrafo_simple(p, valor)


def insertar_bloque_mantenimiento(doc, texto_mantto):
    for i, p in enumerate(doc.paragraphs):
        if p.text.strip().startswith("5.") and "Precio" in p.text:
            insert_after_index = i + 2
            break
    else:
        return

    lineas = [x.strip() for x in texto_mantto.split("\n") if x.strip()]

    # insertar después del precio
    for offset, linea in enumerate(lineas):
        nuevo_p = doc.paragraphs[insert_after_index + offset]._element
        p_new = doc.paragraphs[insert_after_index + offset]

        if p_new.text.strip() == "":
            escribir_parrafo_simple(p_new, linea)
        else:
            # insertar nuevo párrafo antes del existente
            paragraph = doc.add_paragraph()
            escribir_parrafo_simple(paragraph, linea)

            ref = doc.paragraphs[insert_after_index + offset]._element
            ref.addprevious(paragraph._element)


def reemplazar_en_tablas(doc, datos):
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    txt = p.text.strip()
                    if txt == "Conecta":
                        escribir_parrafo_simple(p, datos["cliente"])


def generar_docx(datos):
    doc = Document(TEMPLATE_FILE)

    reemplazar_en_parrafos(doc, datos)
    reemplazar_en_tablas(doc, datos)
    insertar_bloque_mantenimiento(doc, datos["texto_mantto"])

    output = io.BytesIO()
    doc.save(output)
    output.seek(0)
    return output


st.title("APP Área Comercial Buses y Vans")
st.subheader("Generador de Cotización FOTON U9")

with st.form("form_cotizacion"):
    c1, c2 = st.columns(2)

    with c1:
        fecha = st.date_input("Fecha", value=date.today())
        cliente = st.text_input("Cliente", value="Juan Perez")
        cotizante = st.selectbox("Cotizante", list(COTIZANTES.keys()))
        correlativo = st.number_input("Correlativo cotización", min_value=1, value=1, step=1)

    with c2:
        cantidad_unidades = st.number_input("Cantidad de unidades", min_value=1, value=1, step=1)
        precio_unitario = st.number_input("Precio unitario USD", min_value=0.0, value=130491.0, step=1000.0)
        contrato_mantto = st.text_input("Contrato mantto", value="48 meses")

    texto_mantto = st.text_area(
        "Bloque mantenimiento / beneficios",
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

    enviar = st.form_submit_button("Generar cotización")

if enviar:
    info = COTIZANTES[cotizante]
    numero_cotizacion = f'{info["prefijo"]}-{int(correlativo):02d}'

    datos = {
        "fecha_larga": fecha_larga_es(fecha),
        "cliente": cliente,
        "cotizante": cotizante,
        "numero_cotizacion": numero_cotizacion,
        "cantidad_unidades": int(cantidad_unidades),
        "precio_unitario": usd_fmt(precio_unitario),
        "contrato_mantto": contrato_mantto,
        "texto_mantto": texto_mantto,
    }

    try:
        archivo = generar_docx(datos)

        st.success(f"Cotización {numero_cotizacion} generada correctamente.")
        st.write(f"**Cliente:** {cliente}")
        st.write(f"**Cotizante:** {cotizante}")
        st.write(f"**Precio unitario:** {usd_fmt(precio_unitario)}")
        st.write(f"**Total negocio:** {usd_fmt(precio_unitario * cantidad_unidades)}")

        st.download_button(
            "Descargar cotización Word",
            data=archivo,
            file_name=f"Cotizacion_{numero_cotizacion}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )

    except Exception as e:
        st.error(f"Ocurrió un error al generar la cotización: {e}")