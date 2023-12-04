import pandas as pd
from collections import Counter
import re
import tabulate
import plotly.express as px
import altair as alt  # Specify Altair import
from st_aggrid import AgGrid, GridOptionsBuilder
import streamlit_pandas as sp
import streamlit as st

st.set_page_config(page_title="Sales Dashboaradsfd", page_icon=":telephone_receiver:", layout="wide")

# ---- READ EXCEL ----
@st.cache_data
def get_data_from_excel():
    df = pd.read_excel(
        io="input1.xlsx",
        engine="openpyxl",
        sheet_name="Input_llamadas",
        usecols="A:I",
        nrows=151,
    )

    return df

df = get_data_from_excel()



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

# Configurar 'ID_llamada' como índice
df.set_index('ID_llamada', inplace=True)

# Eliminar la columna predeterminada de índices (si existe)
if 'index' in df.columns:
    df.drop(columns='index', inplace=True)



# ---- SIDEBAR ----
st.sidebar.header("Please Filter Here:")

sentimiento = st.sidebar.multiselect(
    "Select the Sentimiento:",
    options=df["Sentimiento"].unique(),
    default=df["Sentimiento"].unique()
)

score = st.sidebar.multiselect(
    "Select the Score:",
    options=df["Score"].unique(),
    default=df["Score"].unique()
)

df_selection = df.query(
    "Sentimiento == @sentimiento & Score == @score"
)

# Filtrar el DataFrame según los filtros seleccionados en el sidebar
df_selection = df.query("Sentimiento == @sentimiento & Score == @score")

# Actualizar el número total de registros en base a la selección actual
total_sales = len(df_selection)

# ---- MAINPAGE ----

st.title(":telephone_receiver: NPS Inteligente - Equipo MTY")
st.markdown("##")


# TOP KPI's
average_rating = int(round(df_selection["Score"].mean(), 0))
star_rating = ":star:" * int(round(average_rating, 0))
average_sale_by_transaction = df_selection["Duración"]


# Lista de stop words (artículos y preposiciones comunes en español)
stop_words = ["cómo","se","el", "la", "los", "las", "un", "una", "unos", "unas", "de", "en", "con", "por", "para", "a", "su", "sus", "al", "del", "sobre"]

# Seleccionar la columna para análisis
columna_seleccionada = "Transcripción"

# Obtener todas las celdas de la columna seleccionada
celdas = df[columna_seleccionada].astype(str).tolist()

# Combinar todas las celdas en un solo texto
texto_completo = " ".join(celdas)

# Limpiar el texto (eliminar signos de puntuación y convertir a minúsculas)
texto_limpiado = re.sub(r'[^\w\s]', '', texto_completo).lower()

# Dividir el texto en palabras
palabras = texto_limpiado.split()

# Filtrar stop words
palabras_filtradas = [palabra for palabra in palabras if palabra not in stop_words]

# Calcular las 5 palabras más frecuentes
contador_palabras = Counter(palabras_filtradas)
palabras_mas_frecuentes = contador_palabras.most_common(10)



left_column, middle_column, right_column = st.columns(3)
with left_column:
    st.subheader("Total de Registros Ingresados:")
    st.subheader(f"{total_sales:,} Registros")
with right_column:
    st.subheader("Calificación promedio:")
    st.subheader(f"{average_rating} {star_rating}")

left_column, middle_column, right_column = st.columns(3)
with left_column:
    st.subheader("Duración promedio por llamada:")
    st.subheader(f"{average_duration:.2f} minutos")


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




st.title("Archivo Ingresado")
st.write(df)





# st.markdown("""---""")

# # SALES BY PRODUCT LINE [BAR CHART]
# sales_by_product_line = df_selection.groupby(by=["Sentimiento"])[["Score"]].sum().sort_values(by="Score")
# fig_product_sales = px.bar(
#     sales_by_product_line,
#     x="Total",
#     y=sales_by_product_line.index,
#     orientation="h",
#     title="<b>Sales by Product Line</b>",
#     color_discrete_sequence=["#0083B8"] * len(sales_by_product_line),
#     template="plotly_white",
# )
# fig_product_sales.update_layout(
#     plot_bgcolor="rgba(0,0,0,0)",
#     xaxis=(dict(showgrid=False))
# )

# # SALES BY HOUR [BAR CHART]
# sales_by_hour = df_selection.groupby(by=["hour"])[["Total"]].sum()
# fig_hourly_sales = px.bar(
#     sales_by_hour,
#     x=sales_by_hour.index,
#     y="Total",
#     title="<b>Sales by hour</b>",
#     color_discrete_sequence=["#0083B8"] * len(sales_by_hour),
#     template="plotly_white",
# )
# fig_hourly_sales.update_layout(
#     xaxis=dict(tickmode="linear"),
#     plot_bgcolor="rgba(0,0,0,0)",
#     yaxis=(dict(showgrid=False)),
# )


# left_column, right_column = st.columns(2)
# left_column.plotly_chart(fig_hourly_sales, use_container_width=True)
# right_column.plotly_chart(fig_product_sales, use_container_width=True)


# # ---- HIDE STREAMLIT STYLE ----
# hide_st_style = """
#             <style>
#             #MainMenu {visibility: hidden;}
#             footer {visibility: hidden;}
#             header {visibility: hidden;}
#             </style>
#             """
# st.markdown(hide_st_style, unsafe_allow_html=True)
