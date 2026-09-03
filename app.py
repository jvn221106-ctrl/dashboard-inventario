import streamlit as st
import pandas as pd
import plotly.express as px

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
    # --- GERENTES DE LOJA ---
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

    # --- ADMINISTRADORES / DIRETORIA (Acesso Total) ---
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
# 3. CARREGAMENTO E TRATAMENTO DA PLANILHA EXCEL
# ==============================================================================
NOME_ARQUIVO_EXCEL = "STATUS DOS INVENTÁRIOS.xlsx"

@st.cache_data(ttl=300)
def carregar_dados():
    try:
        df = pd.read_excel(NOME_ARQUIVO_EXCEL)
        df.columns = [str(c).strip() for c in df.columns]

        # Normaliza a coluna Centro (ex: B008) para aplicação de permissão
        if "Centro" in df.columns:
            df["_CODIGO_CENTRO_"] = df["Centro"].astype(str).str.strip().str.upper()
        else:
            df["_CODIGO_CENTRO_"] = "B000"

        # Converte valores numéricos e garante sinal absoluto de perdas
        if "Qtd. UM registro" in df.columns:
            df["Qtd_Absoluta"] = pd.to_numeric(df["Qtd. UM registro"], errors="coerce").fillna(0).abs()
        else:
            df["Qtd_Absoluta"] = 0

        if "Montante em MI" in df.columns:
            df["Valor_Absoluto"] = pd.to_numeric(df["Montante em MI"], errors="coerce").fillna(0).abs()
        else:
            df["Valor_Absoluto"] = 0

        # Preenche fornecedores em branco
        if "Fornecedor2" in df.columns:
            df["Fornecedor2"] = df["Fornecedor2"].fillna("NÃO INFORMADO")

        return df

    except FileNotFoundError:
        st.error(f"❌ O arquivo `{NOME_ARQUIVO_EXCEL}` não foi localizado no repositório.")
        st.stop()
    except Exception as e:
        st.error(f"❌ Erro ao ler a planilha: {e}")
        st.stop()

df_bruto = carregar_dados()

# ==============================================================================
# 4. BARRA LATERAL (FILTROS DE CENTRO E MARCA)
# ==============================================================================
st.sidebar.title("🔐 Painel de Controle")
st.sidebar.write(f"**Usuário:** `{email_usuario}`")
st.sidebar.write(f"**Escopo:** `{loja_permitida}`")

if st.sidebar.button("🚪 Sair"):
    st.session_state["usuario_logado"] = None
    st.rerun()

st.sidebar.divider()
st.sidebar.header("🔍 Filtros de Visualização")

# Aplicar trava por e-mail primeiro
if loja_permitida == "TODAS":
    lista_centros = ["Todos os Centros"] + sorted(df_bruto["_CODIGO_CENTRO_"].unique().tolist())
    centro_sel = st.sidebar.selectbox("Filtrar por Centro/Loja:", lista_centros)
    
    if centro_sel != "Todos os Centros":
        df_filtrado = df_bruto[df_bruto["_CODIGO_CENTRO_"] == centro_sel]
    else:
        df_filtrado = df_bruto.copy()
else:
    df_filtrado = df_bruto[df_bruto["_CODIGO_CENTRO_"] == loja_permitida]

# Filtro dinâmico por Marcas/Fornecedores
if "Fornecedor2" in df_filtrado.columns:
    lista_marcas = ["Todas as Marcas"] + sorted(df_filtrado["Fornecedor2"].astype(str).unique().tolist())
    marca_sel = st.sidebar.selectbox("Filtrar por Marca/Fornecedor:", lista_marcas)
    
    if marca_sel != "Todas as Marcas":
        df_filtrado = df_filtrado[df_filtrado["Fornecedor2"] == marca_sel]

# ==============================================================================
# 5. DASHBOARD - MÉTRICAS E TOP 10 PERDAS
# ==============================================================================
st.title("📊 Dashboard de Análise de Perdas de Inventário")

# CARDS DE MÉTRICAS
m1, m2, m3, m4 = st.columns(4)
with m1:
    st.metric("Total de Ocorrências", len(df_filtrado))
with m2:
    st.metric("Qtd. Total Perdida", f"{df_filtrado['Qtd_Absoluta'].sum():,.0f}")
with m3:
    st.metric("Valor Total Perda (R$)", f"R$ {df_filtrado['Valor_Absoluto'].sum():,.2f}")
with m4:
    total_m = df_filtrado['Fornecedor2'].nunique() if "Fornecedor2" in df_filtrado.columns else 0
    st.metric("Marcas Envolvidas", total_m)

st.divider()

# GRÁFICOS TOP 10
col_g1, col_g2 = st.columns(2)

with col_g1:
    st.subheader("🏆 Top 10 Lojas com Maiores Perdas (R$)")
    if "Nome 1" in df_filtrado.columns:
        top_lojas = df_filtrado.groupby("Nome 1")["Valor_Absoluto"].sum().reset_index()
        top_lojas = top_lojas.sort_values(by="Valor_Absoluto", ascending=False).head(10)
        
        fig_lojas = px.bar(
            top_lojas, 
            x="Valor_Absoluto", 
            y="Nome 1", 
            orientation="h",
            labels={"Valor_Absoluto": "Perda em R$", "Nome 1": "Loja"},
            text_auto=",.2f",
            color_discrete_sequence=["#d62728"]
        )
        fig_lojas.update_layout(yaxis={"categoryorder": "total ascending"})
        st.plotly_chart(fig_lojas, use_container_width=True)

with col_g2:
    st.subheader("📉 Top 10 Marcas com Maiores Perdas (R$)")
    if "Fornecedor2" in df_filtrado.columns:
        top_marcas = df_filtrado.groupby("Fornecedor2")["Valor_Absoluto"].sum().reset_index()
        top_marcas = top_marcas.sort_values(by="Valor_Absoluto", ascending=False).head(10)
        
        fig_marcas = px.bar(
            top_marcas, 
            x="Valor_Absoluto", 
            y="Fornecedor2", 
            orientation="h",
            labels={"Valor_Absoluto": "Perda em R$", "Fornecedor2": "Marca"},
            text_auto=",.2f",
            color_discrete_sequence=["#ff7f0e"]
        )
        fig_marcas.update_layout(yaxis={"categoryorder": "total ascending"})
        st.plotly_chart(fig_marcas, use_container_width=True)

st.divider()

# ==============================================================================
# 6. TABELA DETALHADA DE DADOS
# ==============================================================================
st.subheader("📋 Detalhamento dos Registros")
colunas_para_exibir = [
    c for c in df_filtrado.columns 
    if c not in ["_CODIGO_CENTRO_", "Qtd_Absoluta", "Valor_Absoluto"]
]
st.dataframe(df_filtrado[colunas_para_exibir], use_container_width=True, hide_index=True)