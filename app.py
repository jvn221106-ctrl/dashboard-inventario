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
# 3. CARREGAMENTO E DETECÇÃO AUTOMÁTICA DA COLUNA
# ==============================================================================
NOME_ARQUIVO_EXCEL = "STATUS DOS INVENTÁRIOS.xlsm"  # Confirme o nome correto da planilha aqui

@st.cache_data(ttl=300)
def carregar_dados():
    try:
        df = pd.read_excel(NOME_ARQUIVO_EXCEL)
        
        # Identifica automaticamente a coluna que guarda o código do centro/loja
        coluna_centro = None
        opcoes_nomes = ["centro", "loja", "cd_centro", "unidade", "lojas", "centros"]
        
        for col in df.columns:
            if str(col).strip().lower() in opcoes_nomes:
                coluna_centro = col
                break
        
        # Se não achar por nome exato, pega a primeira coluna que contém 'centro' ou 'loja'
        if not coluna_centro:
            for col in df.columns:
                if "centro" in str(col).lower() or "loja" in str(col).lower():
                    coluna_centro = col
                    break
        
        # Se ainda assim não achar, define a primeira coluna como padrão
        if not coluna_centro:
            coluna_centro = df.columns[0]
            
        # Padroniza a coluna encontrada para texto sem espaços
        df["_CENTRO_LOGICA_"] = df[coluna_centro].astype(str).str.strip().str.upper()
        return df, coluna_centro
        
    except FileNotFoundError:
        st.error(f"❌ O arquivo `{NOME_ARQUIVO_EXCEL}` não foi localizado no repositório.")
        st.stop()
    except Exception as e:
        st.error(f"❌ Erro ao ler a base de dados: {e}")
        st.stop()

df_bruto, nome_coluna_original = carregar_dados()

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

# Filtragem segura dos dados
if loja_permitida == "TODAS":
    st.sidebar.subheader("Visão de Administrador")
    lista_centros = ["Todos os Centros"] + sorted(df_bruto["_CENTRO_LOGICA_"].unique().tolist())
    centro_selecionado = st.sidebar.selectbox("Filtrar por unidade:", lista_centros)
    
    if centro_selecionado != "Todos os Centros":
        df_exibicao = df_bruto[df_bruto["_CENTRO_LOGICA_"] == centro_selecionado]
    else:
        df_exibicao = df_bruto.copy()
else:
    # Trava rigorosa para o gerente de loja
    df_exibicao = df_bruto[df_bruto["_CENTRO_LOGICA_"] == loja_permitida]

# Remove a coluna temporária antes de exibir na tela
df_para_mostrar = df_exibicao.drop(columns=["_CENTRO_LOGICA_"])

# ==============================================================================
# 5. DASHBOARD E APRESENTAÇÃO
# ==============================================================================
st.title("📊 Painel Geral de Inventário")

if loja_permitida == "TODAS":
    st.info("Acesso global ativado (Perfil Administrador).")
else:
    st.success(f"Acesso liberado exclusivamente para a unidade **{loja_permitida}**.")

# Indicadores simples
col1, col2 = st.columns(2)
with col1:
    st.metric("Total de Registros Exibidos", len(df_para_mostrar))
with col2:
    st.metric("Coluna de Unidade Detectada", f"'{nome_coluna_original}'")

st.divider()

st.subheader("📋 Tabela de Dados")
if df_para_mostrar.empty:
    st.warning("Nenhum dado localizado para esta unidade na base atual.")
else:
    st.dataframe(
        df_para_mostrar,
        use_container_width=True,
        hide_index=True
    )
