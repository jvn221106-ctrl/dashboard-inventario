import streamlit as st
import pandas as pd
import plotly.express as px
import re

# ==============================================================================
# CONFIGURAÇÃO INICIAL DA PÁGINA
# ==============================================================================
st.set_page_config(
    page_title="Painel de Inventários - Análise de Perdas",
    page_icon="📊",
    layout="wide"
)

# ==============================================================================
# 1. MAPEAMENTO DE E-MAILS PARA CENTROS (LOJAS)
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

# ==============================================================================
# 2. CONTROLE DE SESSÃO E LOGIN
# ==============================================================================
if "usuario_logado" not in st.session_state:
    st.session_state["usuario_logado"] = None

if not st.session_state["usuario_logado"]:
    st.title("🔒 Sistema de Inventários - Acesso Restrito")
    st.write("Digite o seu e-mail corporativo cadastrado para acessar o painel.")
    
    with st.form("form_acesso"):
        email_input = st.text_input("E-mail de Acesso:").strip().lower()
        btn_entrar = st.form_submit_button("Acessar Painel")
        
        if btn_entrar:
            if not email_input:
                st.warning("Por favor, digite o seu e-mail.")
            elif email_input in PERMISSOES_EMAIL:
                st.session_state["usuario_logado"] = email_input
                st.rerun()
            else:
                st.error(f"⛔ O e-mail '{email_input}' não está cadastrado ou não possui permissão de acesso.")
    st.stop()

email_usuario = st.session_state["usuario_logado"]
loja_permitida = PERMISSOES_EMAIL[email_usuario]

# ==============================================================================
# 3. LEITURA E TRATAMENTO DA PLANILHA
# ==============================================================================
NOME_ARQUIVO_EXCEL = "STATUS DOS INVENTÁRIOS.xlsm"

def e_codigo_centro(val):
    val_str = str(val).strip().upper()
    return bool(re.match(r"^B\d{3}$", val_str))

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

        # Remove colunas Unnamed
        df = df.loc[:, ~df.columns.astype(str).str.contains('^Unnamed')]
        df.columns = [str(c).strip() for c in df.columns]

        # Centro / Loja
        if "Centro" in df.columns:
            df["_CENTRO_COD_"] = df["Centro"].astype(str).str.strip().str.upper()
        else:
            df["_CENTRO_COD_"] = "B000"

        if "Nome 1" in df.columns:
            df["_LOJA_NOME_"] = df["Nome 1"].astype(str).str.strip()
        else:
            df["_LOJA_NOME_"] = df["_CENTRO_COD_"]

        # Trata Marca
        if "Fornecedor2" in df.columns:
            df["_MARCA_raw"] = df["Fornecedor2"].astype(str).str.strip()
        elif "Rótulos de Linha" in df.columns:
            df["_MARCA_raw"] = df["Rótulos de Linha"].astype(str).str.strip()
        else:
            df["_MARCA_raw"] = ""

        df["_MARCA_"] = df["_MARCA_raw"].apply(lambda x: "" if e_codigo_centro(x) else x)

        # Trata valores
        if "Qtd. UM registro" in df.columns:
            df["_QTD_PERDA_"] = pd.to_numeric(df["Qtd. UM registro"], errors="coerce").fillna(0).round(2)
        elif "Soma de Qtd. UM registro" in df.columns:
            df["_QTD_PERDA_"] = pd.to_numeric(df["Soma de Qtd. UM registro"], errors="coerce").fillna(0).round(2)
        else:
            df["_QTD_PERDA_"] = 0.0

        if "Montante em MI" in df.columns:
            df["_VALOR_PERDA_"] = pd.to_numeric(df["Montante em MI"], errors="coerce").fillna(0).round(2)
        elif "Soma de Montante em MI" in df.columns:
            df["_VALOR_PERDA_"] = pd.to_numeric(df["Soma de Montante em MI"], errors="coerce").fillna(0).round(2)
        else:
            df["_VALOR_PERDA_"] = 0.0

        return df

    except Exception as e:
        st.error(f"❌ Erro ao ler a planilha: {e}")
        st.stop()

df_bruto = carregar_dados()

# ==============================================================================
# 4. BARRA LATERAL (FILTROS)
# ==============================================================================
st.sidebar.title("🔐 Painel de Controle")
st.sidebar.write(f"**Usuário:** `{email_usuario}`")
st.sidebar.write(f"**Escopo:** `{loja_permitida}`")

if st.sidebar.button("🚪 Sair"):
    st.session_state["usuario_logado"] = None
    st.rerun()

st.sidebar.divider()
st.sidebar.header("🔍 Filtros de Visualização")

# 1. Filtro Centro / Loja
if loja_permitida == "TODAS":
    centros_unicos = [str(x) for x in df_bruto["_CENTRO_COD_"].dropna().unique() if str(x).strip() not in ["", "NAN", "NONE"]]
    lista_centros = ["Todos os Centros"] + sorted(centros_unicos)
    centro_sel = st.sidebar.selectbox("Filtrar por Centro/Loja:", lista_centros)
    
    if centro_sel != "Todos os Centros":
        df_filtrado = df_bruto[df_bruto["_CENTRO_COD_"].astype(str) == centro_sel]
    else:
        df_filtrado = df_bruto.copy()
else:
    df_filtrado = df_bruto[df_bruto["_CENTRO_COD_"].astype(str) == str(loja_permitida)]

# 2. Filtro de Marca
marcas_validas = df_filtrado[
    ~df_filtrado["_MARCA_"].str.upper().isin(["", "NAN", "NONE", "NÃO INFORMADO", "NAO INFORMADO"])
]
marcas_unicas = [str(x) for x in marcas_validas["_MARCA_"].dropna().unique()]
lista_marcas = ["Todas as Marcas"] + sorted(marcas_unicas)
marca_sel = st.sidebar.selectbox("Filtrar por Marca/Fornecedor:", lista_marcas)

if marca_sel != "Todas as Marcas":
    df_filtrado = df_filtrado[df_filtrado["_MARCA_"].astype(str) == marca_sel]

# ==============================================================================
# 5. DASHBOARD - MÉTRICAS E TOP 10 PERDAS
# ==============================================================================
st.title("📊 Dashboard de Análise de Perdas de Inventário")

total_registros = len(df_filtrado)
qtd_total = df_filtrado["_QTD_PERDA_"].sum()
valor_total = df_filtrado["_VALOR_PERDA_"].sum()
marcas_contagem = df_filtrado[~df_filtrado["_MARCA_"].str.upper().isin(["", "NAN", "NONE", "NÃO INFORMADO", "NAO INFORMADO"])]["_MARCA_"].nunique()

m1, m2, m3, m4 = st.columns(4)
with m1:
    st.metric("Total de Registros", total_registros)
with m2:
    st.metric("Qtd. Total Perdida", f"{qtd_total:,.0f}")
with m3:
    st.metric("Valor Total Perda (R$)", f"R$ {valor_total:,.2f}")
with m4:
    st.metric("Marcas Envolvidas", marcas_contagem)

st.divider()

col_g1, col_g2 = st.columns(2)

with col_g1:
    st.subheader("🏆 Top 10 Lojas com Maiores Perdas (R$)")
    top_lojas = df_filtrado.groupby("_LOJA_NOME_")["_VALOR_PERDA_"].sum().reset_index()
    top_lojas["_VALOR_PERDA_"] = top_lojas["_VALOR_PERDA_"].round(2)
    top_lojas = top_lojas.sort_values(by="_VALOR_PERDA_", ascending=False).head(10)
    
    fig_lojas = px.bar(
        top_lojas, 
        x="_VALOR_PERDA_", 
        y="_LOJA_NOME_", 
        orientation="h",
        labels={"_VALOR_PERDA_": "Perda em R$", "_LOJA_NOME_": "Loja"},
        text_auto=",.2f",
        color_discrete_sequence=["#d62728"]
    )
    fig_lojas.update_layout(yaxis={"categoryorder": "total ascending"})
    st.plotly_chart(fig_lojas, use_container_width=True)

with col_g2:
    st.subheader("📉 Top 10 Marcas com Maiores Perdas (R$)")
    df_marcas_grafico = df_filtrado[
        ~df_filtrado["_MARCA_"].str.upper().isin(["", "NAN", "NONE", "NÃO INFORMADO", "NAO INFORMADO"])
    ]
    top_marcas = df_marcas_grafico.groupby("_MARCA_")["_VALOR_PERDA_"].sum().reset_index()
    top_marcas["_VALOR_PERDA_"] = top_marcas["_VALOR_PERDA_"].round(2)
    top_marcas = top_marcas.sort_values(by="_VALOR_PERDA_", ascending=False).head(10)
    
    fig_marcas = px.bar(
        top_marcas, 
        x="_VALOR_PERDA_", 
        y="_MARCA_", 
        orientation="h",
        labels={"_VALOR_PERDA_": "Perda em R$", "_MARCA_": "Marca"},
        text_auto=",.2f",
        color_discrete_sequence=["#ff7f0e"]
    )
    fig_marcas.update_layout(yaxis={"categoryorder": "total ascending"})
    st.plotly_chart(fig_marcas, use_container_width=True)

st.divider()

# ==============================================================================
# 6. TABELA DETALHADA DE DADOS (OCULTANDO AS TRÊS COLUNAS SOLICITADAS)
# ==============================================================================
st.subheader("📋 Detalhamento dos Registros")

colunas_bloqueadas = [
    "_CENTRO_COD_", "_LOJA_NOME_", "_MARCA_raw", "_MARCA_", "_QTD_PERDA_", "_VALOR_PERDA_",
    "Rótulos de Linha", "Soma de Qtd. UM registro", "Soma de Montante em MI"
]

colunas_exibir = [
    c for c in df_filtrado.columns 
    if c not in colunas_bloqueadas 
    and not c.startswith("Unnamed") 
    and not c.startswith("Soma de")
]

st.dataframe(df_filtrado[colunas_exibir], use_container_width=True, hide_index=True)
