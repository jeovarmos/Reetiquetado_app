import streamlit as st
import pandas as pd
import math
from datetime import datetime
import os
import warnings
from io import BytesIO

warnings.filterwarnings("ignore", category=UserWarning, module='openpyxl')

def main():
    st.set_page_config(page_title="Programación de Reetiquetado", layout="wide")

    st.title("Programación de Reetiquetado")

    # Initialize session state
    if 'lineas_disponibles' not in st.session_state:
        st.session_state.lineas_disponibles = 12
    if 'horas_lineas' not in st.session_state:
        st.session_state.horas_lineas = {f"L{i:02d}": 37.5 for i in range(1, 13)}
    if 'df' not in st.session_state:
        st.session_state.df = None
    if 'priorizacion_df' not in st.session_state:
        st.session_state.priorizacion_df = None
    if 'file_name' not in st.session_state:
        st.session_state.file_name = None
    if 'priorizacion_file_name' not in st.session_state:
        st.session_state.priorizacion_file_name = None


    with st.sidebar:
        st.header("Configuración")

        # --- Sección de archivo principal ---
        st.subheader("Cargar Archivo Excel")
        uploaded_file = st.file_uploader("Seleccionar Archivo Principal", type=["xlsx", "xls"])
        if uploaded_file:
            try:
                st.session_state.df = pd.read_excel(
                    uploaded_file,
                    sheet_name='Consolidado',
                    dtype={'PRTNUM': str}
                )
                if st.session_state.df.empty:
                    st.error("La hoja 'Consolidado' está vacía")
                    st.session_state.df = None
                else:
                    st.session_state.file_name = uploaded_file.name
                    st.success(f"Archivo '{uploaded_file.name}' cargado.")
            except Exception as e:
                st.error(f"Error al cargar: {str(e)}")
                st.session_state.df = None

        if st.session_state.df is not None:
            if st.button("Eliminar Archivo Principal"):
                st.session_state.df = None
                st.session_state.file_name = None
                st.rerun()

        # --- Sección para archivo de priorización ---
        st.subheader("Priorización Externa")
        uploaded_priorizacion_file = st.file_uploader("Seleccionar Archivo de Priorización", type=["xlsx", "xls", "csv"])
        if uploaded_priorizacion_file:
            try:
                if uploaded_priorizacion_file.name.endswith('.csv'):
                    st.session_state.priorizacion_df = pd.read_csv(uploaded_priorizacion_file, dtype={'PRTNUM': str})
                else:
                    st.session_state.priorizacion_df = pd.read_excel(uploaded_priorizacion_file, dtype={'PRTNUM': str})

                if 'PRTNUM' not in st.session_state.priorizacion_df.columns or 'PRIORIDAD' not in st.session_state.priorizacion_df.columns:
                    st.error("El archivo debe contener columnas 'PRTNUM' y 'PRIORIDAD'")
                    st.session_state.priorizacion_df = None
                elif not pd.api.types.is_numeric_dtype(st.session_state.priorizacion_df['PRIORIDAD']):
                    st.error("La columna 'PRIORIDAD' debe contener valores numéricos.")
                    st.session_state.priorizacion_df = None
                else:
                    st.session_state.priorizacion_file_name = uploaded_priorizacion_file.name
                    st.success(f"Archivo de priorización '{uploaded_priorizacion_file.name}' cargado.")
            except Exception as e:
                st.error(f"Error al cargar el archivo de priorización: {str(e)}")
                st.session_state.priorizacion_df = None

        if st.session_state.priorizacion_df is not None:
            if st.button("Eliminar Priorización"):
                st.session_state.priorizacion_df = None
                st.session_state.priorizacion_file_name = None
                st.rerun()

    # --- Sección de horas por línea ---
    with st.expander("Horas por Línea (semana)", expanded=True):
        col1, col2, col3 = st.columns(3)
        with col1:
            if st.button("➕ Añadir Línea"):
                if st.session_state.lineas_disponibles < 24:
                    st.session_state.lineas_disponibles += 1
                    line_name = f"L{st.session_state.lineas_disponibles:02d}"
                    st.session_state.horas_lineas[line_name] = 37.5
                else:
                    st.warning("Máximo 24 líneas permitidas")
        with col2:
            if st.button("➖ Eliminar Línea"):
                if st.session_state.lineas_disponibles > 1:
                    line_name = f"L{st.session_state.lineas_disponibles:02d}"
                    del st.session_state.horas_lineas[line_name]
                    st.session_state.lineas_disponibles -= 1
                else:
                    st.warning("Debe haber al menos 1 línea")
        with col3:
            if st.button("Reset a 12 líneas"):
                st.session_state.lineas_disponibles = 12
                st.session_state.horas_lineas = {f"L{i:02d}": 37.5 for i in range(1, 13)}


        line_cols = st.columns(6)
        lineas_activas = [f"L{i:02d}" for i in range(1, st.session_state.lineas_disponibles + 1)]
        for i, line_name in enumerate(lineas_activas):
            with line_cols[i % 6]:
                st.session_state.horas_lineas[line_name] = st.number_input(
                    f"{line_name} (horas)",
                    min_value=0.0,
                    value=st.session_state.horas_lineas.get(line_name, 37.5),
                    step=0.5,
                    key=f"horas_{line_name}"
                )

    # --- Priorización estándar ---
    st.subheader("Priorización de Inventario")
    prioridad_inventario = st.radio(
        "Seleccione la prioridad del inventario:",
        ("Mayor inventario primero", "Menor inventario primero"),
        horizontal=True,
        label_visibility="collapsed"
    )

    # --- Semana de inicio ---
    st.subheader("Semana de Inicio")
    current_week = datetime.now().isocalendar()[1]
    semana_inicio = st.number_input("Semana:", min_value=1, max_value=52, value=current_week)
    st.info(f"Semana actual: {current_week}")


    if st.button("Generar Programación", type="primary"):
        if st.session_state.df is None:
            st.error("Debe cargar un archivo Excel primero")
        else:
            with st.spinner("Generando programación..."):
                try:
                    MIN_UNIDADES_POR_ASIGNACION = 10
                    df = st.session_state.df.copy()
                    df.columns = [col.strip().upper() for col in df.columns]
                    required_columns = ['PRTNUM', 'INVENTARIO', 'PROD. HORA', 'CLASIFICACION ABC']
                    missing_columns = [col for col in required_columns if col not in df.columns]

                    if missing_columns:
                        st.error(f"Faltan columnas requeridas en el archivo principal: {', '.join(missing_columns)}")
                        return

                    if st.session_state.priorizacion_df is not None:
                        df = pd.merge(df, st.session_state.priorizacion_df[['PRTNUM', 'PRIORIDAD']], on='PRTNUM', how='left')
                        df['PRIORIDAD'] = df['PRIORIDAD'].fillna(99)
                    else:
                        df['PRIORIDAD'] = 99

                    df['HORAS_REQUERIDAS'] = df['INVENTARIO'] / df['PROD. HORA']
                    abc_priority = {'A': 1, 'B': 2, 'C': 3}
                    df['ABC_PRIORITY'] = df['CLASIFICACION ABC'].map(abc_priority).fillna(99)

                    sort_ascending = (prioridad_inventario == "Menor inventario primero")
                    df_sorted = df.sort_values(
                        by=['PRIORIDAD', 'ABC_PRIORITY', 'INVENTARIO', 'PROD. HORA'],
                        ascending=[True, True, sort_ascending, False]
                    )

                    TOTAL_SEMANAS = 52
                    asignaciones = []
                    inventario_restante = df_sorted.set_index('PRTNUM')['INVENTARIO'].copy()
                    df_sorted.set_index('PRTNUM', inplace=True)

                    for semana in range(semana_inicio, semana_inicio + TOTAL_SEMANAS):
                        lineas_activas = [f"L{i:02d}" for i in range(1, st.session_state.lineas_disponibles + 1)]
                        for linea in lineas_activas:
                            horas_disponibles = st.session_state.horas_lineas[linea]
                            for prtnum, producto in df_sorted.iterrows():
                                if horas_disponibles <= 0.01:
                                    break
                                if inventario_restante.loc[prtnum] <= 0:
                                    continue

                                unidades_max_linea = math.floor(horas_disponibles * producto['PROD. HORA'])
                                unidades_disponibles_stock = int(inventario_restante.loc[prtnum])
                                unidades_a_asignar = min(unidades_disponibles_stock, unidades_max_linea)
                                es_el_remate_final = (unidades_a_asignar == unidades_disponibles_stock)
                                es_un_lote_significativo = (unidades_a_asignar >= MIN_UNIDADES_POR_ASIGNACION)

                                if unidades_a_asignar > 0 and (es_un_lote_significativo or es_el_remate_final):
                                    descripcion = producto.get('DESCRIPCION', 'N/A')
                                    prioridad_externa = producto.get('PRIORIDAD', 'Pred.')
                                    if prioridad_externa == 99:
                                        prioridad_externa = 'Pred.'
                                    horas_necesarias = unidades_a_asignar / producto['PROD. HORA']
                                    nueva_asignacion = {
                                        'Semana': semana,
                                        'Linea': linea,
                                        'PRTNUM': prtnum,
                                        'Descripcion': descripcion,
                                        'Clasificacion_ABC': producto['CLASIFICACION ABC'],
                                        'Prioridad_Externa': prioridad_externa,
                                        'Unidades_Asignadas': int(unidades_a_asignar),
                                        'Horas_Utilizadas': round(horas_necesarias, 2),
                                        'Productividad': producto['PROD. HORA'],
                                        'Unidades Reales': '',
                                        'Horas reales': ''
                                    }
                                    asignaciones.append(nueva_asignacion)
                                    inventario_restante.loc[prtnum] -= unidades_a_asignar
                                    horas_disponibles -= horas_necesarias

                    if asignaciones:
                        df_asignaciones = pd.DataFrame(asignaciones)
                        columnas_finales = [
                            'Semana', 'Linea', 'PRTNUM', 'Descripcion', 'Clasificacion_ABC',
                            'Prioridad_Externa', 'Unidades_Asignadas', 'Horas_Utilizadas',
                            'Productividad', 'Unidades Reales', 'Horas reales'
                        ]
                        df_asignaciones = df_asignaciones[columnas_finales]
                        fecha_creacion = datetime.now().strftime('%Y-%m-%d')
                        df_asignaciones.insert(0, 'Fecha_Creacion', fecha_creacion)
                        
                        st.session_state.df_asignaciones = df_asignaciones
                        st.success("Programación generada con éxito.")

                except Exception as e:
                    st.error(f"Ocurrió un error: {str(e)}")

    if 'df_asignaciones' in st.session_state and st.session_state.df_asignaciones is not None:
        st.dataframe(st.session_state.df_asignaciones)
        
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            st.session_state.df_asignaciones.to_excel(writer, index=False, sheet_name='Programacion')
        
        base_name = f"Programacion_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"

        st.download_button(
            label="Descargar Programación",
            data=output.getvalue(),
            file_name=base_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )


if __name__ == "__main__":
    main()
