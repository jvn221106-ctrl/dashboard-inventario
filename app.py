import streamlit as st
import pandas as pd
import plotly.express as px
import re

# ==============================================================================
# CONFIGURAÇÃO DA PÁGINA E CSS
# ==============================================================================
st.set_page_config(
    page_title="Dashboard Executivo de Inventário",
    page_icon="📊",
    layout="wide"
)

st.markdown("""
<style>
    .stApp {
        background-color: #0e1117;
    }
    div[data-baseweb="select"] > div {
        background-color: #1a1c23;
    }
    span[data-baseweb="tag"] {
        background-color: #e53935 !important;
        color: white !important;
    }
</style>
""", unsafe_allow_html=True)

# ==============================================================================
# 1. PERMISSÕES DE ACESSO
# ==============================================================================
PERMISSOES_EMAIL = {
    "sara.leite@vonnycosmeticos.com.br": "B001",
    "julio.fonseca@vonnycosmeticos.com.br": "B002",
    "fabiana.bertassi@vonnycosmeticos.com.br": "B006",
    "vanessa.tais@vonnycosmeticos.com.br": "B007",
    "yara.silva@vonnycosmeticos.com.br": "B008",
    "josemary.bezerra@vonnycosmeticos.com.br": "B009",
    "maria.beserra@vonnycosmeticos.com.br": "B010",
    "gislaine.barra@vonnycosmeticos.com.br": "B011",
    "thamires.conceicao@vonnycosmeticos.com.br": "B012",
    "vera.silva@vonnycosmeticos.com.br": "B013",
    "vanessa.amaral@vonnycosmeticos.com.br": "B015",
    "claudineia.mendes@vonnycosmeticos.com.br": "B016",
    "thatiane.ferreira@vonnycosmeticos.com.br": "B017",
    "katiane.silva@vonnycosmeticos.com.br": "B018",
    "lanny.andryelly@vonnycosmeticos.com.br": "B019",
    "suzana.silveira@vonnycosmeticos.com.br": "B020",
    "luciana.vasconcelos@vonnycosmeticos.com.br": "B021",
    "daiane.martins@vonnycosmeticos.com.br": "B022",
    "gisele.trampusch@vonnycosmeticos.com.br": "B023",
    "raquel.lopes@vonnycosmeticos.com.br": "B024",
    "claudinea.santos@vonnycosmeticos.com.br": "B025",
    "rosania.chagas@vonnycosmeticos.com.br": "B026",
    "luana.costa@vonnycosmeticos.com.br": "B027",
    "rosangela.botelho@vonnycosmeticos.com.br": "B028",
    "elza.silva@vonnycosmeticos.com.br": "B029",
    "joao.pereira@vonnycosmeticos.com.br": "B030",
    "controladoriaprevencao@gmail.com": "B031",
    "sergio.oliveira@vonnycosmeticos.com.br": "TODAS",
    "jvn221106@gmail.com": "TODAS"
}

if "usuario_logado" not in st.session_state:
    st.session_state["usuario_logado"] = None

if not st.session_state["usuario_logado"]:
    st.title("🔒 Acesso Restrito")
    with st.form("form_acesso"):
        email_input = st.text_input("E-mail corporativo:").strip().lower()
        if st.form_submit_button("Acessar Painel"):
            if email_input in PERMISSOES_EMAIL:
                st.session_state["usuario_logado"] = email_input
                st.rerun()
            else:
                st.error("E-mail não cadastrado.")
    st.stop()

email_usuario = st.session_state["usuario_logado"]
loja_permitida = PERMISSOES_EMAIL[email_usuario]

# ==============================================================================
# 2. CARREGAMENTO E LEITURA DOS DADOS
# ==============================================================================
NOME_ARQUIVO_EXCEL = "STATUS DOS INVENTÁRIOS.xlsm"

def e_codigo_centro(val):
    return bool(re.match(r"^B\d{3}$", str(val).strip().upper()))

@st.cache_data(ttl=300)
def carregar_dados():
    try:
        excel_file = pd.ExcelFile(NOME_ARQUIVO_EXCEL, engine="openpyxl")
        df_perdas = None
        
        for sheet in excel_file.sheet_names:
            df_temp = pd.read_excel(excel_file, sheet_name=sheet)
            cols = [str(c).strip() for c in df_temp.columns]
            if "Fornecedor2" in cols or "Montante em MI" in cols:
                df_perdas = df_temp
                break
        
        if df_perdas is None:
            df_perdas = pd.read_excel(excel_file, sheet_name=0)

        df = df_perdas.copy()
        df = df.loc[:, ~df.columns.astype(str).str.contains('^Unnamed')]
        df.columns = [str(c).strip() for c in df.columns]

        # Centro
        if "Centro" in df.columns:
            df["_CENTRO_COD_"] = df["Centro"].astype(str).str.strip().str.upper()
        else:
            df["_CENTRO_COD_"] = "B000"

        # Marca
        if "Fornecedor2" in df.columns:
            marca_col = "Fornecedor2"
        elif "Rótulos de Linha" in df.columns:
            marca_col = "Rótulos de Linha"
        else:
            marca_col = None

        if marca_col:
            df["_MARCA_"] = df[marca_col].astype(str).str.strip()
            df["_MARCA_"] = df["_MARCA_"].apply(lambda x: "" if e_codigo_centro(x) else x)
        else:
            df["_MARCA_"] = ""

        # Qtd
        col_qtd = next((c for c in ["Qtd. UM registro", "Soma de Qtd. UM registro", "Qtd"] if c in df.columns), None)
        if col_qtd:
            df["_QTD_PERDA_"] = pd.to_numeric(df[col_qtd], errors="coerce").fillna(0)
        else:
            df["_QTD_PERDA_"] = 0.0

        # Valor
        col_val = next((c for c in ["Montante em MI", "Soma de Montante em MI", "Valor"] if c in df.columns), None)
        if col_val:
            df["_VALOR_PERDA_"] = pd.to_numeric(df[col_val], errors="coerce").fillna(0)
        else:
            df["_VALOR_PERDA_"] = 0.0

        return df
    except Exception as e:
        st.error(f"Erro ao carregar a planilha: {e}")
        st.stop()

df_bruto = carregar_dados()

# ==============================================================================
# 3. FILTROS NA SIDEBAR
# ==============================================================================
st.sidebar.title("Filtros")

centros_disponiveis = sorted([str(x) for x in df_bruto["_CENTRO_COD_"].dropna().unique() if str(x).strip() not in ["", "NAN", "NONE"]])

if loja_permitida != "TODAS":
    centros_permitidos = [loja_permitida] if loja_permitida in centros_disponiveis else []
    centros_selecionados = st.sidebar.multiselect("Selecione os Centros:", options=centros_permitidos, default=centros_permitidos)
else:
    centros_selecionados = st.sidebar.multiselect("Selecione os Centros:", options=centros_disponiveis, default=centros_disponiveis)

df_filtrado = df_bruto[df_bruto["_CENTRO_COD_"].isin(centros_selecionados)]

marcas_validas = df_filtrado[
    ~df_filtrado["_MARCA_"].str.upper().isin(["", "NAN", "NONE", "NÃO INFORMADO", "NAO INFORMADO"])
]
marcas_disponiveis = sorted([str(x) for x in marcas_validas["_MARCA_"].dropna().unique()])
marcas_selecionadas = st.sidebar.multiselect("Selecione as Marcas:", options=marcas_disponiveis, default=marcas_disponiveis)

if marcas_selecionadas:
    df_filtrado = df_filtrado[df_filtrado["_MARCA_"].isin(marcas_selecionadas)]

# ==============================================================================
# 4. TÍTULO PRINCIPAL E KPIS (TELA 1)
# ==============================================================================
st.title("📊 Dashboard Executivo de Inventário")
st.markdown("---")

qtd_perda = df_filtrado[df_filtrado["_QTD_PERDA_"] < 0]["_QTD_PERDA_"].sum()
valor_perda = df_filtrado[df_filtrado["_VALOR_PERDA_"] < 0]["_VALOR_PERDA_"].sum()
valor_sobras = df_filtrado[df_filtrado["_VALOR_PERDA_"] > 0]["_VALOR_PERDA_"].sum()
resultado_net = df_filtrado["_VALOR_PERDA_"].sum()

c1, c2, c3, c4 = st.columns(4)

with c1:
    st.caption("Total de Perdas (Qtd)")
    st.title(f"{qtd_perda:,.0f} UN".replace(",", "."))

with c2:
    st.caption("Perda Total (R$)")
    st.title(f"-R$ {abs(valor_perda):,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))

with c3:
    st.caption("Sobras / Ajustes (+)")
    st.title(f"R$ {valor_sobras:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))

with c4:
    st.caption("Resultado Net (Caixa)")
    sinal = "-R$" if resultado_net < 0 else "R$"
    st.title(f"{sinal} {abs(resultado_net):,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))

st.markdown("<br>", unsafe_allow_html=True)

# ==============================================================================
# 5. GRÁFICOS DE PERDA POR CENTRO (BARRA - TELA 1)
# ==============================================================================
col_bar1, col_bar2 = st.columns(2)

df_centro_agrupado = df_filtrado.groupby("_CENTRO_COD_").agg({
    "_QTD_PERDA_": "sum",
    "_VALOR_PERDA_": "sum"
}).reset_index()

with col_bar1:
    st.subheader("📦 Perda por Centro (Qtd)")
    df_c_qtd = df_centro_agrupado.sort_values(by="_QTD_PERDA_", ascending=True)
    fig_c_qtd = px.bar(
        df_c_qtd,
        x="_CENTRO_COD_",
        y="_QTD_PERDA_",
        text_auto=".0f",
        color_discrete_sequence=["#3b82f6"]
    )
    fig_c_qtd.update_layout(xaxis_title=None, yaxis_title=None, plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)")
    st.plotly_chart(fig_c_qtd, use_container_width=True)

with col_bar2:
    st.subheader("🎯 Perda por Centro (R$)")
    df_c_val = df_centro_agrupado.sort_values(by="_VALOR_PERDA_", ascending=True)
    fig_c_val = px.bar(
        df_c_val,
        x="_CENTRO_COD_",
        y="_VALOR_PERDA_",
        text_auto=",.2f",
        color_discrete_sequence=["#3b82f6"]
    )
    fig_c_val.update_layout(xaxis_title=None, yaxis_title=None, plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)")
    st.plotly_chart(fig_c_val, use_container_width=True)

st.markdown("<br><hr><br>", unsafe_allow_html=True)

# ==============================================================================
# 6. RANKING TOP 10 CENTROS (TABELA - TELA 2)
# ==============================================================================
st.subheader("🏢 Ranking: Top 10 Centros com Maior Perda")

df_top_c = df_centro_agrupado.sort_values(by="_VALOR_PERDA_", ascending=True).head(10).reset_index(drop=True)
df_top_c["Posição"] = [f"{i+1}º" for i in range(len(df_top_c))]
df_top_c["Perda (Qtd)"] = df_top_c["_QTD_PERDA_"].apply(lambda x: f"{x:,.0f} un".replace(",", "."))
df_top_c["Perda (R$)"] = df_top_c["_VALOR_PERDA_"].apply(lambda x: f"R$ {x:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))

st.dataframe(
    df_top_c[["Posição", "_CENTRO_COD_", "Perda (Qtd)", "Perda (R$)"]].rename(columns={"_CENTRO_COD_": "Centro"}),
    use_container_width=True,
    hide_index=True
)

st.markdown("<br><hr><br>", unsafe_allow_html=True)

# ==============================================================================
# 7. RANKING TOP 10 MARCAS (TABELA - TELA 3)
# ==============================================================================
st.subheader("⚠️ Ranking: Top 10 Marcas com Maior Perda")

df_marcas_validas = df_filtrado[
    ~df_filtrado["_MARCA_"].str.upper().isin(["", "NAN", "NONE", "NÃO INFORMADO", "NAO INFORMADO"])
]

df_marca_agrupado = df_marcas_validas.groupby("_MARCA_").agg({
    "_QTD_PERDA_": "sum",
    "_VALOR_PERDA_": "sum"
}).reset_index()

df_top_m = df_marca_agrupado.sort_values(by="_VALOR_PERDA_", ascending=True).head(10).reset_index(drop=True)
df_top_m["Posição"] = [f"{i+1}º" for i in range(len(df_top_m))]
df_top_m["Perda (Qtd)"] = df_top_m["_QTD_PERDA_"].apply(lambda x: f"{x:,.0f} un".replace(",", "."))
df_top_m["Perda (R$)"] = df_top_m["_VALOR_PERDA_"].apply(lambda x: f"R$ {x:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))

st.dataframe(
    df_top_m[["Posição", "_MARCA_", "Perda (Qtd)", "Perda (R$)"]].rename(columns={"_MARCA_": "Marca"}),
    use_container_width=True,
    hide_index=True
)

st.markdown("<br><hr><br>", unsafe_allow_html=True)

# ==============================================================================
# 8. GRÁFICOS DE LINHA - PERDAS POR MARCA (TELA 4)
# ==============================================================================
col_line1, col_line2 = st.columns(2)

with col_line1:
    st.subheader("📦 Perdas por Marca - Todas (Qtd)")
    df_m_qtd = df_marca_agrupado.sort_values(by="_QTD_PERDA_", ascending=True)
    fig_m_qtd = px.line(
        df_m_qtd,
        x="_MARCA_",
        y="_QTD_PERDA_",
        markers=True,
        text="_QTD_PERDA_",
        color_discrete_sequence=["#f97316"]
    )
    fig_m_qtd.update_traces(texttemplate='%{text:,.0f} un', textposition='top center')
    fig_m_qtd.update_layout(xaxis_title=None, yaxis_title=None, xaxis_tickangle=-45, plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)")
    st.plotly_chart(fig_m_qtd, use_container_width=True)

with col_line2:
    st.subheader("🏷️ Perdas por Marca - Todas (R$)")
    df_m_val = df_marca_agrupado.sort_values(by="_VALOR_PERDA_", ascending=True)
    fig_m_val = px.line(
        df_m_val,
        x="_MARCA_",
        y="_VALOR_PERDA_",
        markers=True,
        text="_VALOR_PERDA_",
        color_discrete_sequence=["#38bdf8"]
    )
    fig_m_val.update_traces(texttemplate='%{text:,.0f}', textposition='top center')
    fig_m_val.update_layout(xaxis_title=None, yaxis_title=None, xaxis_tickangle=-45, plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)")
    st.plotly_chart(fig_m_val, use_container_width=True)