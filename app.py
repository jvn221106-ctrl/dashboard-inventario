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
# Relacione aqui o e-mail de cada gerente ao seu respectivo código de loja.
# ==============================================================================
PERMISSOES_EMAIL = {
    # --- GERENTES DE LOJA ---
    "sara.leite@vonnycosmeticos.com.br": "B001",
    "julio.fonseca@vonnycosmeticos.com.br": "B002",
    "fabiana.bertassi@vonnycosmeticos.com.br": "B006",
    "Vanessa.tais@vonnycosmeticos.com.br": "B007",
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
    "Suzana.silveira@vonnycosmeticos.com.br": "B020",
    "luciana.vasconcelos@vonnycosmeticos.com.br": "B021",
    "daiane.martins@vonnycosmeticos.com.br": "B022",
    "gisele.trampusch@vonnycosmeticos.com.br": "B023",
    "raquel.lopes@vonnycosmeticos.com.br": "B024",
    "claudinea.santos@vonnycosmeticos.com.br": "B025",
    "Rosania.chagas@vonnycosmeticos.com.br": "B026",
    "luana.costa@vonnycosmeticos.com.br": "B027",
    "rosangela.botelho@vonnycosmeticos.com.br": "B028",
    "elza.silva@vonnycosmeticos.com.br ": "B029",
    "joao.pereira@vonnycosmeticos.com.br": "B030",
    "controladoriaprevencao@gmail.com": "B031",
    
    # --- ADMINISTRADORES / DIRETORIA (Acesso Total) ---
    "sergio.oliveira@vonnycosmeticos.com.br": "TODAS",
    "jvn221106@gmail.com": "TODAS"
}

# ==============================================================================
# 2. CONTROLE DE AUTENTICAÇÃO E SEGURANÇA (Compatível)
# ==============================================================================
# Obtém o e-mail do usuário autenticado no Streamlit Cloud
email_usuario = getattr(st.user, "email", None)

# Se o usuário não estiver logado
if not email_usuario:
    st.title("🔒 Sistema de Inventários - Acesso Restrito")
    st.write("Por favor, faça login com a sua conta Google cadastrada para acessar o painel.")
    st.login()
    st.stop()

# Limpa e padroniza o e-mail
email_usuario = email_usuario.lower().strip()

# Valida permissão do e-mail
if email_usuario not in PERMISSOES_EMAIL:
    st.error(f"⛔ O e-mail ({email_usuario}) não possui autorização de acesso.")
    st.info("Entre em contato com o administrador para solicitar liberação.")
    if st.button("Trocar de Conta"):
        st.logout()
    st.stop()

# Define a loja permitida
loja_permitida = PERMISSOES_EMAIL[email_usuario]

# ==============================================================================
# 3. CARREGAMENTO E TRATAMENTO DA BASE DE DADOS
# ==============================================================================
NOME_ARQUIVO_EXCEL = "STATUS DOS INVENTÁRIOS.xlsm" # Nome do seu arquivo na pasta do projeto

@st.cache_data(ttl=600)  # Recarrega a base a cada 10 minutos
def carregar_dados():
    try:
        df = pd.read_excel(NOME_ARQUIVO_EXCEL)
        # Garante que a coluna Centro seja tratada como texto
        if "Centro" in df.columns:
            df["Centro"] = df["Centro"].astype(str).str.strip().str.upper()
        return df
    except FileNotFoundError:
        st.error(f"❌ O arquivo `{NOME_ARQUIVO_EXCEL}` não foi encontrado no servidor.")
        st.stop()
    except Exception as e:
        st.error(f"❌ Erro ao ler o arquivo de dados: {e}")
        st.stop()

df_bruto = carregar_dados()

# ==============================================================================
# 4. APLICAÇÃO DOS FILTROS DE SEGURANÇA E BARRA LATERAL
# ==============================================================================
# Informações do usuário na barra lateral
st.sidebar.title("🔐 Painel de Controle")
st.sidebar.write(f"**Usuário:** `{email_usuario}`")
st.sidebar.write(f"**Escopo de Acesso:** `{loja_permitida}`")

if st.sidebar.button("🚪 Sair da Conta"):
    st.logout()

st.sidebar.divider()

# Aplicando regras de isolamento dos dados
if loja_permitida == "TODAS":
    st.sidebar.subheader("Modo Administrador")
    lista_centros = ["Todos os Centros"] + sorted(df_bruto["Centro"].unique().tolist())
    centro_selecionado = st.sidebar.selectbox("Filtrar por centro específico:", lista_centros)
    
    if centro_selecionado != "Todos os Centros":
        df_exibicao = df_bruto[df_bruto["Centro"] == centro_selecionado]
    else:
        df_exibicao = df_bruto.copy()
else:
    # Trava rigorosa: recupera apenas as linhas pertencentes à loja do gerente
    df_exibicao = df_bruto[df_bruto["Centro"] == loja_permitida]

# ==============================================================================
# 5. DASHBOARD E APRESENTAÇÃO DOS DADOS
# ==============================================================================
st.title("📊 Painel Geral de Inventário")

if loja_permitida == "TODAS":
    st.info("Visão Administrativa: Acesso liberado a todas as unidades.")
else:
    st.success(f"Visão de Gerente: Exibindo dados exclusivos da unidade **{loja_permitida}**.")

# Resumo em Métricas principais
col1, col2, col3 = st.columns(3)
with col1:
    st.metric("Total de Registros", len(df_exibicao))
with col2:
    if "Quantidade" in df_exibicao.columns:
        st.metric("Total de Itens", f"{df_exibicao['Quantidade'].sum():,.0f}")
    else:
        st.metric("Total de Centros", df_exibicao["Centro"].nunique())
with col3:
    if "Valor Total" in df_exibicao.columns:
        st.metric("Valor Total (R$)", f"R$ {df_exibicao['Valor Total'].sum():,.2f}")
    else:
        st.metric("Status da Base", "Atualizada")

st.divider()

# Exibição do relatório em tabela interativa
st.subheader("📋 Detalhamento do Estoque")
if df_exibicao.empty:
    st.warning("Nenhum dado localizado para esta unidade na base atual.")
else:
    st.dataframe(
        df_exibicao,
        use_container_width=True,
        hide_index=True
    )