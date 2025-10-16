import pandas as pd
import streamlit as st
import altair as alt
from datetime import date
import locale

# === CONFIGURAÇÕES REGIONAIS ===
locale.setlocale(locale.LC_TIME, "pt_BR.utf8")

# === CONFIGURAÇÃO DE PÁGINA ===
st.set_page_config(
    page_title="Dashboard Cirúrgico Evolutivo",
    page_icon="💉",
    layout="wide",
    initial_sidebar_state="expanded"
)

# === ESTILO CLARO E PREMIUM ===
st.markdown("""
<style>
    body { background-color: #F5F5F5; color: #000000; font-family: 'Segoe UI', sans-serif; }
    .titulo { font-size:2.8em; font-weight:700; color:#007ACC; margin-bottom:-5px; }
    .subtitulo { font-size:1.1em; color: #555555; margin-bottom:30px; }
    .metric-card { background: #FFFFFF; padding:25px; border-radius:15px; text-align:center; box-shadow:0 8px 30px rgba(0,0,0,0.2); transition: transform 0.2s; }
    .metric-card:hover { transform: translateY(-5px); }
    .metric-value { font-size:2.2em; font-weight:bold; color:#007ACC; }
    .metric-label { font-size:0.9em; color:#333333; text-transform:uppercase; letter-spacing:0.05em; }
    hr { border:none; height:1px; background-color:#007ACC; margin:35px 0; }
    .stMultiSelect { font-size:0.8em !important; }
</style>
""", unsafe_allow_html=True)

# === CABEÇALHO ===
st.markdown("<div class='titulo'>💉 Dashboard Cirúrgico Evolutivo</div>", unsafe_allow_html=True)
st.markdown("<div class='subtitulo'>Evolução mensal do volume de cirurgias por Sala</div>", unsafe_allow_html=True)

# === UPLOAD DE ARQUIVO ===
uploaded_file = st.file_uploader("📁 Envie seu arquivo Excel (.xlsx ou .xls)", type=["xlsx","xls"])
if uploaded_file:
    arquivo = pd.read_excel(uploaded_file)
    arquivo["DATA Inicial"] = pd.to_datetime(arquivo["DATA Inicial"], errors="coerce")
    arquivo = arquivo.dropna(subset=["DATA Inicial"])
    
    # === FILTRO DE PERÍODO ===
    min_data = arquivo["DATA Inicial"].min().date()
    max_data = arquivo["DATA Inicial"].max().date()
    st.sidebar.header("📅 Filtro de Período")
    data_inicial, data_final = st.sidebar.date_input(
        "Selecione o período",
        [min_data, max_data],
        min_value=min_data,
        max_value=max_data,
        format="DD/MM/YYYY"
    )
    
    # === FILTRO MULTI-SELEÇÃO DE SALA ===
    st.sidebar.header("🏥 Filtrar por Sala")
    salas_disponiveis = arquivo["SALA"].dropna().unique().tolist()
    salas_selecionadas = st.sidebar.multiselect(
        "Selecione uma ou mais Salas",
        options=salas_disponiveis,
        default=salas_disponiveis
    )
    
    # === APLICAR FILTROS ===
    arquivo_filtrado = arquivo[
        (arquivo["DATA Inicial"].dt.date >= data_inicial) &
        (arquivo["DATA Inicial"].dt.date <= data_final)
    ]
    if salas_selecionadas:
        arquivo_filtrado = arquivo_filtrado[arquivo_filtrado["SALA"].isin(salas_selecionadas)]
    
    # === AGRUPAMENTO MENSAL ===
    arquivo_filtrado["Mes_Ano"] = arquivo_filtrado["DATA Inicial"].dt.to_period("M")
    cirurgias_mensais = (
        arquivo_filtrado.groupby("Mes_Ano")["RESERVA"]
        .count()
        .reset_index()
        .rename(columns={"RESERVA":"Quantidade"})
    )
    cirurgias_mensais["Mes_Ano"] = cirurgias_mensais["Mes_Ano"].dt.to_timestamp()
    cirurgias_mensais = cirurgias_mensais.sort_values("Mes_Ano")  # ORDEM CRONOLÓGICA
    cirurgias_mensais["Mes_Formatado"] = cirurgias_mensais["Mes_Ano"].dt.strftime("%b/%Y").str.capitalize()
    cirurgias_mensais["Quantidade"] = cirurgias_mensais["Quantidade"].astype(int)

    # === KPIs ---
    col1, col2, col3 = st.columns(3)
    total = f"{int(cirurgias_mensais['Quantidade'].sum()):,}".replace(",", ".")
    media = f"{round(cirurgias_mensais['Quantidade'].mean()):,}".replace(",", ".")
    maior_mes = cirurgias_mensais.loc[cirurgias_mensais["Quantidade"].idxmax(),"Mes_Formatado"]
    
    col1.markdown(f"<div class='metric-card'><div class='metric-value'>{total}</div><div class='metric-label'>Total de Cirurgias</div></div>", unsafe_allow_html=True)
    col2.markdown(f"<div class='metric-card'><div class='metric-value'>{media}</div><div class='metric-label'>Média Mensal</div></div>", unsafe_allow_html=True)
    col3.markdown(f"<div class='metric-card'><div class='metric-value'>{maior_mes}</div><div class='metric-label'>Mês de Maior Volume</div></div>", unsafe_allow_html=True)

    st.markdown("<hr>", unsafe_allow_html=True)
    
    # === GRÁFICO PREMIUM COM LINHAS VERTICAIS E RÓTULOS FORMATADOS ---
    linha = alt.Chart(cirurgias_mensais).mark_line(
        point=alt.OverlayMarkDef(filled=True, color="#007ACC", size=200),
        color=alt.Gradient(
            gradient='linear',
            stops=[alt.GradientStop(color='#00BFFF', offset=0),
                   alt.GradientStop(color='#007ACC', offset=1)]
        ),
        strokeWidth=4
    ).encode(
        x=alt.X("Mes_Ano:T",
                axis=alt.Axis(format="%b/%Y", labelAngle=-35, labelFontSize=12, title=None, labelFontWeight='bold')),
        y=alt.Y("Quantidade:Q", axis=alt.Axis(title=None, labels=False)),
        tooltip=[alt.Tooltip("Mes_Ano:T", title="Mês/Ano", format="%b/%Y"),
                 alt.Tooltip("Quantidade:Q", title="Quantidade", format=",")]
    ).interactive()

    # --- linhas verticais sutis ---
    linhas_verticais = alt.Chart(cirurgias_mensais).mark_rule(color='lightgray', strokeWidth=1, opacity=0.3).encode(
        x="Mes_Ano:T",
        y=alt.Y("Quantidade:Q"),
        y2=alt.value(0)
    )

    # --- rótulos dos volumes mensais ---
    rotulos = alt.Chart(cirurgias_mensais).mark_text(
        align="center", dy=-20, color="#007ACC",
        fontWeight="bold", fontSize=14
    ).encode(
        x="Mes_Ano:T",
        y="Quantidade:Q",
        text=alt.Text("Quantidade:Q", format=",")
    )

    grafico = (linhas_verticais + linha + rotulos).properties(
        width="container",
        height=550,
        title="📈 Evolução Mensal de Cirurgias por Sala"
    ).configure_title(
        fontSize=24,
        fontWeight="bold",
        color="#007ACC",
        anchor="start"
    ).configure_axis(
        labelFontSize=12
    )

    st.altair_chart(grafico, use_container_width=True)

    # === DOWNLOAD DE DADOS ---
    st.markdown("### ⬇️ Exportar Dados Filtrados")
    csv_bytes = cirurgias_mensais.to_csv(index=False).encode()
    st.download_button("Download CSV", data=csv_bytes, file_name="cirurgias_agrupadas.csv", mime="text/csv")
