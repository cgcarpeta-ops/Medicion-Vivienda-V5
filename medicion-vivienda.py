from io import BytesIO
import json
import unicodedata

from fpdf import FPDF
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
import streamlit as st

st.set_page_config(page_title="Medición de Vivienda", page_icon="📐", layout="centered")
st.title("📐 Medición de Vivienda")

TIPOS_SUELO = ["Porcelánico / Cerámico", "Tarima"]

if "estancias" not in st.session_state:
    st.session_state.estancias = []


def _sanitizar_pdf(texto: str) -> str:
    """Adapta cadenas a fuentes estándar latin-1 de FPDF evitando '?' innecesarios."""
    reemplazos = {
        "–": "-",
        "—": "-",
        "−": "-",
        "“": '"',
        "”": '"',
        "’": "'",
        "‘": "'",
        "€": "EUR",
        "²": "2",
        "³": "3",
        "°": " deg",
    }
    for orig, dest in reemplazos.items():
        texto = texto.replace(orig, dest)

    texto_norm = unicodedata.normalize("NFKD", texto)
    return texto_norm.encode("latin-1", errors="replace").decode("latin-1")


def _truncar_texto(texto: str, max_chars: int = 24) -> str:
    texto = texto.strip()
    if len(texto) <= max_chars:
        return texto
    return texto[: max_chars - 1].rstrip() + "…"


def _calcular_totales(estancias: list[dict]) -> dict:
    """Calcula totales generales y desglose de suelo por material sin redondeos intermedios."""
    tot_paredes = sum(float(e["Paredes (m²)"]) for e in estancias)
    tot_suelos = sum(float(e["Suelo (m²)"]) for e in estancias)
    tot_techos = sum(float(e["Techo (m²)"]) for e in estancias)

    por_tipo = {tipo: 0.0 for tipo in TIPOS_SUELO}
    for e in estancias:
        tipo_suelo = str(e.get("Tipo de suelo", TIPOS_SUELO[0]))
        por_tipo[tipo_suelo] = por_tipo.get(tipo_suelo, 0.0) + float(e["Suelo (m²)"])

    return {
        "paredes": round(tot_paredes, 2),
        "suelos": round(tot_suelos, 2),
        "techos": round(tot_techos, 2),
        "suelos_por_tipo": {k: round(v, 2) for k, v in por_tipo.items()},
    }


@st.cache_data(show_spinner=False)
def _generar_pdf_desde_json(estancias_json: str) -> bytes:
    estancias: list[dict] = json.loads(estancias_json)
    totales = _calcular_totales(estancias)

    pdf = FPDF(orientation="P", unit="mm", format="A4")
    pdf.set_auto_page_break(auto=True, margin=15)
    pdf.set_margins(left=10, top=12, right=10)
    pdf.add_page()

    # Suma de anchos = 190 mm (A4 210 mm - 20 mm márgenes)
    anchos = [38, 28, 17, 17, 17, 23, 22, 20]
    headers = ["Estancia", "Tipo Suelo", "Largo", "Ancho", "Alto", "Paredes", "Suelo", "Techo"]

    def _dibujar_cabecera() -> None:
        pdf.set_font("Helvetica", "B", 8)
        pdf.set_fill_color(230, 236, 245)
        for h, w in zip(headers, anchos):
            pdf.cell(w, 8, _sanitizar_pdf(h), border=1, align="C", fill=True)
        pdf.ln()

    pdf.set_font("Helvetica", "B", 14)
    pdf.cell(0, 10, _sanitizar_pdf("Informe de Mediciones y Superficies de Vivienda"), ln=True)
    pdf.set_font("Helvetica", "I", 8)
    pdf.cell(
        0,
        6,
        _sanitizar_pdf("Nota: Cálculo en planta rectangular sin descontar huecos (puertas/ventanas)."),
        ln=True,
    )
    pdf.ln(3)

    _dibujar_cabecera()

    pdf.set_font("Helvetica", "", 8)
    for e in estancias:
        if pdf.get_y() + 16 > pdf.page_break_trigger:
            pdf.add_page()
            _dibujar_cabecera()
            pdf.set_font("Helvetica", "", 8)

        fila = [
            _sanitizar_pdf(_truncar_texto(str(e["Estancia"]), 20)),
            _sanitizar_pdf(_truncar_texto(str(e.get("Tipo de suelo", "")), 15)),
            f"{float(e['Largo (m)']):.2f}m",
            f"{float(e['Ancho (m)']):.2f}m",
            f"{float(e['Altura (m)']):.2f}m",
            f"{float(e['Paredes (m²)']):.2f}m2",
            f"{float(e['Suelo (m²)']):.2f}m2",
            f"{float(e['Techo (m²)']):.2f}m2",
        ]
        for idx, (val, w) in enumerate(zip(fila, anchos)):
            alineacion = "L" if idx in (0, 1) else "R"
            pdf.cell(w, 7, val, border=1, align=alineacion)
        pdf.ln()

    if pdf.get_y() + 28 > pdf.page_break_trigger:
        pdf.add_page()

    pdf.ln(2)
    pdf.set_font("Helvetica", "B", 8.5)
    pdf.set_fill_color(242, 244, 248)
    pdf.cell(sum(anchos[:5]), 8, _sanitizar_pdf("TOTALES GENERALES"), border=1, align="R", fill=True)
    pdf.cell(anchos[5], 8, f"{totales['paredes']:.2f} m2", border=1, align="R", fill=True)
    pdf.cell(anchos[6], 8, f"{totales['suelos']:.2f} m2", border=1, align="R", fill=True)
    pdf.cell(anchos[7], 8, f"{totales['techos']:.2f} m2", border=1, align="R", fill=True)
    pdf.ln()

    pdf.ln(4)
    pdf.set_font("Helvetica", "B", 10)
    pdf.cell(0, 7, _sanitizar_pdf("Desglose de Suelos por Material"), ln=True)
    pdf.set_font("Helvetica", "", 9)
    for tipo, metros in totales["suelos_por_tipo"].items():
        pdf.cell(
            0,
            6.5,
            _sanitizar_pdf(f"- {tipo}: {metros:.2f} m2"),
            ln=True,
        )

    salida = pdf.output()
    if isinstance(salida, (bytes, bytearray)):
        return bytes(salida)
    return salida.encode("latin-1")


@st.cache_data(show_spinner=False)
def _generar_xlsx_desde_json(estancias_json: str) -> bytes:
    estancias: list[dict] = json.loads(estancias_json)
    totales = _calcular_totales(estancias)

    wb = Workbook()
    ws = wb.active
    ws.title = "Mediciones"

    columnas = [
        "Estancia",
        "Tipo de suelo",
        "Largo (m)",
        "Ancho (m)",
        "Altura (m)",
        "Paredes (m²)",
        "Suelo (m²)",
        "Techo (m²)",
    ]

    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    total_font = Font(bold=True)
    total_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    sub_fill = PatternFill(start_color="EDF2F8", end_color="EDF2F8", fill_type="solid")

    ws.append(columnas)
    for col_idx in range(1, len(columnas) + 1):
        celda = ws.cell(row=1, column=col_idx)
        celda.font = header_font
        celda.fill = header_fill
        celda.alignment = Alignment(horizontal="center", vertical="center")

    for fila_idx, e in enumerate(estancias, start=2):
        valores = [
            str(e["Estancia"]),
            str(e.get("Tipo de suelo", TIPOS_SUELO[0])),
            round(float(e["Largo (m)"]), 2),
            round(float(e["Ancho (m)"]), 2),
            round(float(e["Altura (m)"]), 2),
            round(float(e["Paredes (m²)"]), 2),
            round(float(e["Suelo (m²)"]), 2),
            round(float(e["Techo (m²)"]), 2),
        ]
        ws.append(valores)
        for col_idx, val in enumerate(valores, start=1):
            celda = ws.cell(row=fila_idx, column=col_idx)
            if col_idx in (1, 2):
                celda.alignment = Alignment(horizontal="left", vertical="center")
            else:
                celda.number_format = "0.00"
                celda.alignment = Alignment(horizontal="right", vertical="center")

    fila_total = len(estancias) + 2
    ws.cell(row=fila_total, column=1, value="TOTALES GENERALES").alignment = Alignment(
        horizontal="right", vertical="center"
    )
    for col_idx in range(2, 6):
        ws.cell(row=fila_total, column=col_idx, value="")

    totales_col = {
        6: totales["paredes"],
        7: totales["suelos"],
        8: totales["techos"],
    }
    for col_idx, valor_tot in totales_col.items():
        celda = ws.cell(row=fila_total, column=col_idx, value=valor_tot)
        celda.number_format = "0.00"
        celda.alignment = Alignment(horizontal="right", vertical="center")

    for col_idx in range(1, len(columnas) + 1):
        celda = ws.cell(row=fila_total, column=col_idx)
        celda.font = total_font
        celda.fill = total_fill

    # Bloque de desglose de suelos por material
    fila_desglose_titulo = fila_total + 2
    ws.cell(row=fila_desglose_titulo, column=1, value="Desglose de Suelos por Material").font = total_font
    ws.cell(row=fila_desglose_titulo, column=2, value="Superficie (m²)").font = total_font
    for col_idx in (1, 2):
        ws.cell(row=fila_desglose_titulo, column=col_idx).fill = sub_fill

    fila_actual = fila_desglose_titulo + 1
    for tipo, metros in totales["suelos_por_tipo"].items():
        ws.cell(row=fila_actual, column=1, value=tipo).alignment = Alignment(horizontal="left", vertical="center")
        celda_m2 = ws.cell(row=fila_actual, column=2, value=metros)
        celda_m2.number_format = "0.00"
        celda_m2.alignment = Alignment(horizontal="right", vertical="center")
        fila_actual += 1

    # Ajuste automático de ancho de columnas con margen
    for col_idx in range(1, len(columnas) + 1):
        col_letter = get_column_letter(col_idx)
        max_len = max(
            len(str(ws.cell(row=r, column=col_idx).value or ""))
            for r in range(1, fila_actual)
        )
        ws.column_dimensions[col_letter].width = max(max_len + 4, 13)

    buffer = BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer.getvalue()


with st.form("form_estancia", clear_on_submit=True):
    st.subheader("Añadir nueva estancia")

    opciones = [
        "Hall",
        "Pasillo 1",
        "Pasillo 2",
        "Cocina",
        "Baño",
        "Aseo",
        "Despensa",
        "Salón",
        "Dormitorio 1",
        "Dormitorio 2",
        "Dormitorio 3",
        "Terraza",
        "Galería",
        "Otra (escribir abajo)",
    ]
    seleccion = st.selectbox("Estancia", opciones)
    nombre_custom = st.text_input("Nombre personalizado (si elegiste 'Otra')")

    tipo_suelo = st.radio(
        "Tipo de suelo de la estancia",
        TIPOS_SUELO,
        horizontal=True,
    )

    col1, col2, col3 = st.columns(3)
    with col1:
        largo = st.number_input("Largo (m)", min_value=0.0, step=0.05, format="%.2f")
    with col2:
        ancho = st.number_input("Ancho (m)", min_value=0.0, step=0.05, format="%.2f")
    with col3:
        altura = st.number_input("Altura (m)", min_value=0.0, value=2.60, step=0.05, format="%.2f")

    enviado = st.form_submit_button("➕ Añadir estancia", use_container_width=True)

    if enviado:
        nombre_limpio = nombre_custom.strip()
        if seleccion == "Otra (escribir abajo)":
            nombre_final = nombre_limpio if nombre_limpio else "Estancia personalizada"
        else:
            nombre_final = seleccion

        if largo > 0 and ancho > 0 and altura > 0:
            largo_f = float(largo)
            ancho_f = float(ancho)
            altura_f = float(altura)

            paredes = 2.0 * (largo_f + ancho_f) * altura_f
            suelo = largo_f * ancho_f
            techo = suelo

            st.session_state.estancias.append(
                {
                    "Estancia": nombre_final,
                    "Tipo de suelo": tipo_suelo,
                    "Largo (m)": round(largo_f, 2),
                    "Ancho (m)": round(ancho_f, 2),
                    "Altura (m)": round(altura_f, 2),
                    "Paredes (m²)": round(paredes, 2),
                    "Suelo (m²)": round(suelo, 2),
                    "Techo (m²)": round(techo, 2),
                }
            )
            st.success(f"Añadida: {nombre_final} ({tipo_suelo})")
        else:
            st.error("Introduce medidas mayores que 0 en largo, ancho y altura.")


if st.session_state.estancias:
    st.subheader("Resumen de estancias medidas")
    st.dataframe(st.session_state.estancias, use_container_width=True)

    totales = _calcular_totales(st.session_state.estancias)

    c1, c2, c3 = st.columns(3)
    c1.metric("Total Paredes", f"{totales['paredes']:.2f} m²")
    c2.metric("Total Suelos", f"{totales['suelos']:.2f} m²")
    c3.metric("Total Techos", f"{totales['techos']:.2f} m²")

    st.markdown("**Desglose de suelos por material**")
    cols_suelo = st.columns(len(TIPOS_SUELO))
    for idx, tipo in enumerate(TIPOS_SUELO):
        cols_suelo[idx].metric(f"Suelo {tipo}", f"{totales['suelos_por_tipo'].get(tipo, 0.0):.2f} m²")

    estancias_json = json.dumps(st.session_state.estancias, ensure_ascii=False)
    pdf_bytes = _generar_pdf_desde_json(estancias_json)
    xlsx_bytes = _generar_xlsx_desde_json(estancias_json)

    col_dl1, col_dl2 = st.columns(2)
    with col_dl1:
        st.download_button(
            label="📄 Descargar Informe PDF (DIN A4)",
            data=pdf_bytes,
            file_name="informe_calculo_vivienda.pdf",
            mime="application/pdf",
            use_container_width=True,
        )
    with col_dl2:
        st.download_button(
            label="📊 Descargar Hoja Excel (.xlsx)",
            data=xlsx_bytes,
            file_name="informe_calculo_vivienda.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

    col_act1, col_act2 = st.columns(2)
    with col_act1:
        if st.button("🗑️ Borrar última estancia", use_container_width=True):
            st.session_state.estancias.pop()
            st.rerun()
    with col_act2:
        if st.button("🧹 Vaciar todas las estancias", use_container_width=True):
            st.session_state.estancias.clear()
            st.rerun()
