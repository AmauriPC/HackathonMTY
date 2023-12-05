import pandas as pd
from collections import Counter
import re
import streamlit as st
import altair as alt
from io import BytesIO
from fpdf import FPDF
import toml

# Configurar el diseño de la página a "wide"
st.set_page_config(layout="wide")

# Leer el archivo TOML
config_data = toml.load("config.toml")

# Acceder a los valores en el archivo TOML
app_title = config_data["app"]["title"]
app_description = config_data["app"]["description"]

# Configurar la aplicación Streamlit
st.title(app_title)
st.write(app_description)

# Inicializar DataFrame vacío
df = pd.DataFrame()

# ---- READ EXCEL ----
uploaded_file = st.file_uploader("Subir archivo tipo Excel", type="xlsx")
if uploaded_file is not None:
    df = pd.read_excel(uploaded_file, engine="openpyxl", sheet_name="Input_llamadas", usecols="A:I")
else:
    st.warning("Agrega el registro de llamadas en un archivo tipo Excel")

# Check if the DataFrame is not empty
if not df.empty:
    # Convert DataFrame to HTML
    html_table = df.to_html(index=False)

    # Use JavaScript to adjust the height of the div container and prevent scrolling
    js_code = f"""
    <script>
        document.addEventListener('DOMContentLoaded', function() {{
            var container = document.getElementById('table-container');
            container.style.height = '100vh';
            container.style.overflow = 'hidden';
        }});
    </script>
    """
    st.markdown(js_code, unsafe_allow_html=True)

    #############  OPERACIONES sobre un nuevo dataframe de auxiliar para el promedio de duracion
    # Crear un nuevo DataFrame para realizar operaciones
    df_operaciones = df.copy()

    # Convert 'Duración' column to datetime
    df_operaciones['Duración_Prueba'] = pd.to_datetime(df_operaciones['Duración'], format='%H:%M')

    # Extraer los minutos de la columna 'Duración'
    df_operaciones['Duración_minutes'] = df_operaciones['Duración_Prueba'].dt.hour

    # Calcular el promedio de duración en minutos
    average_duration = df_operaciones['Duración_minutes'].mean()

    #############  FORMATEOS
    # Convert 'Hora' column to datetime and format
    df['Hora'] = pd.to_datetime(df['Hora']).dt.strftime('%H:%M')
    # Convert 'Duración' column to datetime and format
    df['Duración'] = pd.to_datetime(df['Duración']).dt.strftime('%H:%M')
    # Extraer solo la fecha (sin la hora)
    df['Fecha'] = df['Fecha'].dt.date

    # Dividir la transcripción en palabras y contar la cantidad de palabras
    df['Palabras_Transcripcion'] = df['Transcripción'].apply(lambda x: len(str(x).split()))

    # Calcular el promedio de palabras en las transcripciones
    promedio_palabras = round(df['Palabras_Transcripcion'].mean())


    # Configurar 'ID_llamada' como índice
    df.set_index('ID_llamada', inplace=True)

    # Eliminar la columna predeterminada de índices (si existe)
    if 'index' in df.columns:
        df.drop(columns='index', inplace=True)

        # SIDEBAR
    st.sidebar.header("Please Filter Here:")

    sentimiento = st.sidebar.multiselect(
        "Select Sentimiento:",
        options=df["Sentimiento"].unique(),
        default=df["Sentimiento"].unique()
    )

    score = st.sidebar.multiselect(
        "Select Score:",
        options=df["Score"].unique(),
        default=df["Score"].unique()
    )

    # Filtrar el DataFrame según los filtros seleccionados en el sidebar
    df_selection = df.query("Sentimiento == @sentimiento & Score == @score")

    # Calcular el promedio de palabras en las transcripciones
    promedio_palabras = round(df_selection['Palabras_Transcripcion'].mean())

    # Calcular la hora del día con más llamadas
    df_selection['Hora'] = pd.to_datetime(df_selection['Hora'])
    df_selection['Hora_del_dia'] = df_selection['Hora'].dt.hour
    frecuencia_horas = df_selection['Hora_del_dia'].value_counts()
    hora_mas_frecuente = frecuencia_horas.idxmax()

    # Actualizar el número total de registros en base a la selección actual
    total_registros = len(df_selection)

    # ---- MAINPAGE ----
    st.title(":telephone_receiver: NPS Inteligente - Equipo MTY")
    st.markdown("##")

    # TOP KPI's
    average_rating = int(round(df_selection["Score"].mean(), 0))
    star_rating = ":star:" * int(round(average_rating, 0))
    average_sale_by_transaction = df_selection["Duración"]

    # Lista de stop words (artículos y preposiciones comunes en español)
    stop_words = ["cómo", "se", "el", "la", "los", "las", "un", "una", "unos",
                  "unas", "de", "en", "con", "por", "para", "a", "su", "sus",
                  "al", "del", "sobre"]

    # Seleccionar la columna para análisis
    columna_seleccionada = "Transcripción"

    # Obtener todas las celdas de la columna seleccionada
    celdas = df_selection[columna_seleccionada].astype(str).tolist()

    # Combinar todas las celdas en un solo texto
    texto_completo = " ".join(celdas)

    # Limpiar el texto (eliminar signos de puntuación y convertir a minúsculas)
    texto_limpiado = re.sub(r'[^\w\s]', '', texto_completo).lower()

    # Dividir el texto en palabras
    palabras = texto_limpiado.split()

    # Filtrar stop words
    palabras_filtradas = [palabra for palabra in palabras if palabra not in stop_words]

    # Calcular las 10 palabras más frecuentes
    contador_palabras = Counter(palabras_filtradas)
    palabras_mas_frecuentes = contador_palabras.most_common(10)

    left_column, middle_column, right_column = st.columns(3)
    with left_column:
        st.subheader("Total de Registros Ingresados: " f"{total_registros:,} Registros")
    with right_column:
        st.subheader("Calificación promedio:")
        st.subheader(f"{average_rating} {star_rating}")

    # left_column, middle_column, right_column = st.columns(3)
    with left_column:
        st.subheader("Duración promedio por llamada: " f"{average_duration:.2f} minutos")
    with right_column:
        st.subheader("Promedio de palabras por transcripción: " f"{promedio_palabras} palabras")
    with left_column:
        st.subheader(f"Hora del día con más llamadas: {hora_mas_frecuente}:00")
        
        
    # Dividir la lista de palabras en dos partes
    mitad = len(palabras_mas_frecuentes) // 2
    parte1 = palabras_mas_frecuentes[:mitad]
    parte2 = palabras_mas_frecuentes[mitad:]

    # Crear listas separadas para palabras y frecuencias de ambas partes
    palabras_parte1, frecuencias_parte1 = zip(*parte1)
    palabras_parte2, frecuencias_parte2 = zip(*parte2)

    # Crear un DataFrame combinado para ambas partes
    df_combined = pd.DataFrame({
        "Palabras": palabras_parte1,
        "Frecuencia": frecuencias_parte1,
        "Palabras ": palabras_parte2,
        "Frecuencia ": frecuencias_parte2
    })

    # Convertir DataFrame a formato Markdown
    table_markdown = df_combined.to_markdown(index=False)


    # Mostrar resultados en Streamlit como texto en formato Markdown
    st.subheader("Las 10 palabras más frecuentes (sin artículos y preposiciones):")
    left_column, middle_column, right_column = st.columns(3)
    with middle_column:
        st.markdown(table_markdown)

    # Contar la cantidad de cada categoría
    count_df = df_selection['Sentimiento'].value_counts().reset_index()
    count_df.columns = ['Sentimiento', 'Cantidad']

    # Mapear colores
    color_scale = alt.Scale(domain=['Negativo', 'Neutral', 'Positivo'],
                            range=['red', 'yellow', 'green'])

    # Crear la gráfica de barras con altair
    chart = alt.Chart(count_df).mark_bar().encode(
        x=alt.X('Sentimiento', title='Sentimiento', axis=alt.Axis(labelAngle=0)),
        y='Cantidad',
        color=alt.Color('Sentimiento:N', scale=color_scale)
    ).properties(
        title='Cantidad de Llamadas por Sentimiento'
    )

    # Mostrar gráfica en Streamlit
    st.altair_chart(chart, use_container_width=True)

    # # Imprime archivo ingresado en formato de tabla
    st.title("Archivo Ingresado")
    st.dataframe(df)

    # Agregar botones de descarga para el DataFrame en diferentes formatos
    st.download_button(
        label="Descargar como CSV",
        data=df.to_csv(index=False).encode('utf-8'),
        file_name="dataframe.csv",
        mime="text/csv"
    )

    # Descargar como Excel
    excel_buffer = BytesIO()
    df.to_excel(excel_buffer, index=False, engine="openpyxl", encoding='utf-8')
    excel_buffer.seek(0)
    st.download_button(
        label="Descargar como Excel",
        data=excel_buffer,
        file_name="dataframe.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

else:
    st.warning("El archivo está vacío. Favor de subir uno válido.")


# ---- HIDE STREAMLIT STYLE ----
hide_st_style = """
            <style>
            #MainMenu {visibility: hidden;}
            footer {visibility: hidden;}
            header {visibility: hidden;}
            </style>
            """
st.markdown(hide_st_style, unsafe_allow_html=True)