import streamlit as st
import openpyxl
from io import BytesIO
from datetime import datetime

st.set_page_config(page_title="Distribución de Repuestos", layout="wide")
st.title("📦 Distribución de Repuestos")


# ---------- BOTÓN REINICIAR ----------

if st.button("🔄 Reiniciar"):
    st.session_state.clear()
    st.rerun()


archivo = st.file_uploader("Cargar archivo Excel", type=["xlsx"])


def buscar_columna(encabezados, posibles):
    for p in posibles:
        for e in encabezados:
            if e and p.lower() in str(e).lower():
                return encabezados.index(e)
    return None


if archivo:

    wb = openpyxl.load_workbook(archivo)
    ws = wb.active

    encabezados = [c.value for c in next(ws.iter_rows(min_row=1, max_row=1))]

    col_caso = buscar_columna(encabezados, ["caso"])
    col_centro = buscar_columna(encabezados, ["centro"])
    col_fecha = 10  # Columna K

    if None in (col_caso, col_centro):
        st.error("❌ Archivo no válido")
        st.stop()

    datos_originales = []
    casos_vistos = set()
    datos = []

    casos_woden = set()
    casos_logy = set()

    # -------- ELIMINAR DUPLICADOS --------

    for fila in ws.iter_rows(min_row=2, values_only=True):

        caso = fila[col_caso]
        centro = str(fila[col_centro]).upper() if fila[col_centro] else ""

        datos_originales.append(list(fila))

        if not caso or caso in casos_vistos:
            continue

        casos_vistos.add(caso)

        if "WODEN" in centro:
            casos_woden.add(caso)

        if "LOGYTECH" in centro:
            casos_logy.add(caso)

        datos.append(list(fila))

    total_rep_original = len(datos_originales)
    total_casos = len(datos)

    # -------- MÉTRICAS --------

    c1, c2, c3, c4 = st.columns(4)

    c1.metric("Casos únicos", total_casos)
    c2.metric("Casos WODEN", len(casos_woden))
    c3.metric("Casos LOGYTECH", len(casos_logy))
    c4.metric("Filas originales", total_rep_original)

    # -------- INPUTS --------

    colA, colB = st.columns(2)

    personas = colA.number_input("Número de líderes técnicos", min_value=1, step=1)
    por_persona = colB.number_input("Casos por líder técnico", min_value=1, step=1)

    total_asignar = personas * por_persona

    st.write(f"📦 Total de casos a asignar: {total_asignar}")

    if total_casos >= total_asignar:
        st.success(f"Datos suficientes. Sobrantes: {total_casos - total_asignar}")
    else:
        st.error(f"No hay suficientes casos. Faltan {total_asignar - total_casos}")

    # -------- GENERAR SOLO UNA VEZ --------

    if st.button("Generar distribución"):

        datos.sort(key=lambda x: int(x[col_caso]))

        asignados = datos[:total_asignar]
        sobrantes = datos[total_asignar:]

        # ROUND ROBIN
        grupos = [[] for _ in range(personas)]

        for i, fila in enumerate(asignados):
            grupos[i % personas].append(fila)

        # ---------- ARCHIVO PRINCIPAL ----------

        wb_out = openpyxl.Workbook()
        wb_out.remove(wb_out.active)

        for i, grupo in enumerate(grupos):

            prioridad = []
            resto = []

            for fila in grupo:
                centro = str(fila[col_centro]).upper() if fila[col_centro] else ""

                if "WODEN" in centro or "LOGYTECH" in centro:
                    prioridad.append(fila)
                else:
                    resto.append(fila)

            resto.sort(key=lambda x: int(x[col_caso]))

            resto_fecha = sorted(
                resto,
                key=lambda x: x[col_fecha]
                if isinstance(x[col_fecha], datetime)
                else datetime.max
            )

            mitad = len(resto) // 2
            organizados = prioridad + resto[:mitad] + resto_fecha[:mitad]

            ws_out = wb_out.create_sheet(f"Tec_lid{i+1}")
            ws_out.append(encabezados)

            for fila in organizados:
                ws_out.append(fila)

        buffer1 = BytesIO()
        wb_out.save(buffer1)

        # ---------- SOBRANTES ----------

        wb_rest = openpyxl.Workbook()
        ws_rest = wb_rest.active
        ws_rest.title = "Repuestos no asignados"
        ws_rest.append(encabezados)

        sobrantes.sort(key=lambda x: int(x[col_caso]))

        for fila in sobrantes:
            ws_rest.append(fila)

        buffer2 = BytesIO()
        wb_rest.save(buffer2)

        # 🔒 GUARDAR EN MEMORIA
        st.session_state["dist"] = buffer1.getvalue()
        st.session_state["sobrantes"] = buffer2.getvalue()

    # -------- MOSTRAR SI YA EXISTEN --------

    if "dist" in st.session_state:

        st.success("Distribución generada ✅")

        st.download_button(
            "⬇️ Descargar distribución",
            data=st.session_state["dist"],
            file_name="Distribucion_Casos.xlsx"
        )

        st.download_button(
            "⬇️ Descargar repuestos no asignados",
            data=st.session_state["sobrantes"],
            file_name="Repuestos_No_Asignados.xlsx"
        )