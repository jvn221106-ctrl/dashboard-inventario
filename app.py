import streamlit as st
import pandas as pd

# ==============================================================================
# CONFIGURAÇÃO INICIAL DA PÁGINA
# ==============================================================================
st.set_page_config(
    page_title="Painel de Inventários",
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

# Tela de Formuário de Login
if not st.session_state["usuario_logado"]:
    st.title("🔒 Sistema de Inventários - Acesso Restrito")
    st.write("Digite o seu e-mail corporativo cadastrado para acessar o painel da sua unidade.")
    
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

# Recupera credenciais
email_usuario = st.session_state["usuario_logado"]
loja_permitida = PERMISSOES_EMAIL[email_usuario]

# ==============================================================================
# 3. CARREGAMENTO E TRATAMENTO DA BASE DE DADOS
# ==============================================================================
NOME_ARQUIVO_EXCEL = "STATUS DOS INVENTÁRIOS.xlsm"  # Altere para o nome do seu arquivo Excel se for diferente

@st.cache_data(ttl=300)
def carregar_dados():
    try:
        df = pd.read_excel(NOME_ARQUIVO_EXCEL)
        if "Centro" in df.columns:
            df["Centro"] = df["Centro"].astype(str).str.strip().str.upper()
        return df
    except FileNotFoundError:
        st.error(f"❌ O arquivo `{NOME_ARQUIVO_EXCEL}` não foi localizado no repositório.")
        st.stop()
    except Exception as e:
        st.error(f"❌ Erro ao processar a base de dados: {e}")
        st.stop()

df_bruto = carregar_dados()

# ==============================================================================
# 4. BARRA LATERAL E REGRAS DE ISOLAMENTO
# ==============================================================================
st.sidebar.title("🔐 Painel de Controle")
st.sidebar.write(f"**Usuário:** `{email_usuario}`")
st.sidebar.write(f"**Escopo:** `{loja_permitida}`")

if st.sidebar.button("🚪 Sair"):
    st.session_state["usuario_logado"] = None
    st.rerun()

st.sidebar.divider()

if loja_permitida == "TODAS":
    st.sidebar.subheader("Visão de Administrador")
    lista_centros = ["Todos os Centros"] + sorted(df_bruto["Centro"].unique().tolist())
    centro_selecionado = st.sidebar.selectbox("Filtrar por unidade:", lista_centros)
    
    if centro_selecionado != "Todos os Centros":
        df_exibicao = df_bruto[df_bruto["Centro"] == centro_selecionado]
    else:
        df_exibicao = df_bruto.copy()
else:
    # Trava rigorosa: filtra apenas a loja vinculada ao gerente logado
    df_exibicao = df_bruto[df_bruto["Centro"] == loja_permitida]

# ==============================================================================
# 5. DASHBOARD E APRESENTAÇÃO
# ==============================================================================
st.title("📊 Painel Geral de Inventário")

if loja_permitida == "TODAS":
    st.info("Acesso global ativado (Perfil Administrador).")
else:
    st.success(f"Acesso liberado exclusivamente para a unidade **{loja_permitida}**.")

col1, col2, col3 = st.columns(3)
with col1:
    st.metric("Total de Registros", len(df_exibicao))
with col2:
    if "Quantidade" in df_exibicao.columns:
        st.metric("Total de Itens", f"{df_exibicao['Quantidade'].sum():,.0f}")
    else:
        st.metric("Total de Centros Exibidos", df_exibicao["Centro"].nunique())
with col3:
    if "Valor Total" in df_exibicao.columns:
        st.metric("Valor Total", f"R$ {df_exibicao['Valor Total'].sum():,.2f}")
    else:
        st.metric("Status da Base", "Carregada")

st.divider()

st.subheader("📋 Tabela de Dados")
if df_exibicao.empty:
    st.warning("Nenhum dado localizado para esta unidade.")
else:
    st.dataframe(
        df_exibicao,
        use_container_width=True,
        hide_index=True
    )
