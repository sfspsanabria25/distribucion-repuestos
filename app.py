import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="Distribución de Repuestos", layout="wide")

st.title("🔧 Distribución de Repuestos")

archivo = st.file_uploader("Cargar archivo Excel", type=["xlsx"])

if archivo:

    df = pd.read_excel(archivo)

    # ===== NORMALIZAR NOMBRES DE COLUMNAS =====
    df.columns = (
        df.columns
        .astype(str)
        .str.strip()
        .str.upper()
    )

    # ===== DETECCIÓN AUTOMÁTICA DE COLUMNAS =====

    # CASO
    posibles_caso = [c for c in df.columns if "CASO" in c]
    if not posibles_caso:
        st.error("No se encontró la columna de CASO")
        st.stop()

    col_caso = posibles_caso[0]

    # CENTRO DE SERVICIO
    posibles_centro = [
        c for c in df.columns
        if "CENTRO" in c and "SERVICIO" in c
    ]

    if not posibles_centro:
        st.error("No se encontró la columna de Centro de servicio")
        st.stop()

    col_centro = posibles_centro[0]

    # FECHA SOLICITUD (Columna K usualmente)
    if len(df.columns) >= 11:
        col_fecha = df.columns[10]
    else:
        col_fecha = df.columns[-1]

    # ===== ELIMINAR DUPLICADOS POR CASO =====
    df = df.drop_duplicates(subset=[col_caso])

    total_casos = len(df)

    # ===== PRIORIDAD =====
    centros_prioridad = ["WODEN S.A.S.", "LOGYTECH MOBILE"]

    df_prioridad = df[df[col_centro].isin(centros_prioridad)].copy()
    df_otros = df[~df[col_centro].isin(centros_prioridad)].copy()

    casos_woden = len(df[df[col_centro] == "WODEN S.A.S."])
    casos_logy = len(df[df[col_centro] == "LOGYTECH MOBILE"])

    st.info(f"Total de casos: {total_casos}")
    st.info(f"WODEN S.A.S.: {casos_woden}")
    st.info(f"LOGYTECH MOBILE: {casos_logy}")

    # ===== PARÁMETROS =====
    col1, col2 = st.columns(2)

    with col1:
        num_lideres = st.number_input(
            "Número de líderes técnicos",
            min_value=1,
            value=2
        )

    with col2:
        repuestos_por_lider = st.number_input(
            "Cantidad de repuestos por líder",
            min_value=1,
            value=50
        )

    total_requerido = num_lideres * repuestos_por_lider
    st.write(f"Total requerido: {total_requerido}")

    # ===== GENERAR DISTRIBUCIÓN =====
    if st.button("Generar distribución"):

        # PRIORIDAD COMPLETA SIEMPRE
        asignados = df_prioridad.copy()

        restantes_cupo = total_requerido - len(asignados)

        df_otros = df_otros.sort_values(by=col_caso)

        if restantes_cupo > 0:
            asignados = pd.concat(
                [asignados, df_otros.head(restantes_cupo)],
                ignore_index=True
            )

        # ===== SOBRANTES =====
        sobrantes = df[~df[col_caso].isin(asignados[col_caso])].copy()
        sobrantes = sobrantes.sort_values(by=col_caso)

        # ===== ALEATORIZAR =====
        asignados = asignados.sample(frac=1).reset_index(drop=True)

        # ===== DISTRIBUCIÓN ENTRE LÍDERES =====
        distribucion = []
        inicio = 0

        for i in range(num_lideres):
            fin = inicio + repuestos_por_lider
            distribucion.append(asignados.iloc[inicio:fin].copy())
            inicio = fin

        # ===== ORGANIZACIÓN INTERNA =====
        hojas_finales = []

        for df_lider in distribucion:

            prioridad = df_lider[df_lider[col_centro].isin(centros_prioridad)]
            otros = df_lider[~df_lider[col_centro].isin(centros_prioridad)]

            mitad = len(otros) // 2

            por_caso = otros.sort_values(by=col_caso).head(mitad)
            por_fecha = otros.sort_values(by=col_fecha).iloc[mitad:]

            hoja = pd.concat([prioridad, por_caso, por_fecha])
            hojas_finales.append(hoja)

        # ===== EXCEL DISTRIBUCIÓN =====
        buffer = BytesIO()

        with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:

            for i, hoja in enumerate(hojas_finales, start=1):
                hoja.to_excel(writer, sheet_name=f"Tec_lid{i}", index=False)

        st.download_button(
            label="Descargar distribución",
            data=buffer.getvalue(),
            file_name="distribucion_repuestos.xlsx"
        )

        # ===== EXCEL SOBRANTES =====
        buffer2 = BytesIO()

        with pd.ExcelWriter(buffer2, engine="xlsxwriter") as writer:
            sobrantes.to_excel(writer, index=False)

        st.download_button(
            label="Descargar repuestos no asignados",
            data=buffer2.getvalue(),
            file_name="repuestos_no_asignados.xlsx"
        )

        st.success("Distribución generada correctamente")

    if st.button("Reiniciar"):
        st.rerun()
