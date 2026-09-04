import streamlit as st
import json
import os
import hashlib

# Configuração da página
st.set_page_config(
    page_title="Sistema Vonny Cosméticos",
    page_icon="🔒",
    layout="wide"
)

DB_FILE = "usuarios_db.json"

# --- MAPEAMENTO DE E-MAILS PERMITIDOS E SUAS RESPECTIVAS LOJAS ---
EMAILS_PERMITIDOS = {
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
    "jvn221106@gmail.com": "TODAS"  # Administrador
}

# E-mails com privilégio de Administrador
ADMINS = ["jvn221106@gmail.com"]


# --- FUNÇÕES DE BANCO DE DADOS (JSON) ---
def carregar_usuarios():
    """Carrega o banco de dados do JSON ou sincroniza com a lista oficial."""
    if os.path.exists(DB_FILE):
        with open(DB_FILE, "r") as f:
            usuarios = json.load(f)
    else:
        usuarios = {}

    # Sincroniza novos e-mails e atualiza o código da loja se mudou
    atualizou = False
    for email, loja in EMAILS_PERMITIDOS.items():
        if email not in usuarios:
            usuarios[email] = {
                "loja": loja,
                "senha": None  # None = primeiro acesso pendente
            }
            atualizou = True
        else:
            # Garante que a informação da loja esteja sempre sincronizada
            if usuarios[email].get("loja") != loja:
                usuarios[email]["loja"] = loja
                atualizou = True

    if atualizou or not os.path.exists(DB_FILE):
        salvar_usuarios(usuarios)

    return usuarios

def salvar_usuarios(usuarios):
    """Salva a estrutura atualizada de usuários no JSON."""
    with open(DB_FILE, "w") as f:
        json.dump(usuarios, f, indent=4)

def gerar_hash(senha):
    """Gera hash SHA-256 da senha."""
    return hashlib.sha256(senha.encode()).hexdigest()


# --- ESTADO DE SESSÃO ---
if "logado" not in st.session_state:
    st.session_state["logado"] = False
if "usuario_atual" not in st.session_state:
    st.session_state["usuario_atual"] = None


# --- TELA DE LOGIN / PRIMEIRO ACESSO ---
def renderizar_tela_login():
    st.title("🔒 Vonny Cosméticos - Acesso ao Sistema")
    st.write("Digite seu e-mail para acessar.")

    usuarios = carregar_usuarios()

    with st.form("form_login"):
        email_input = st.text_input("E-mail corporativo:").strip().lower()
        senha_input = st.text_input("Senha:", type="password")
        btn_entrar = st.form_submit_button("Entrar / Cadastrar", type="primary")

    if btn_entrar:
        if not email_input:
            st.error("Por favor, digite seu e-mail.")
            return

        if email_input not in usuarios:
            st.error("E-mail não autorizado. Entre em contato com a controladoria/administrador.")
            return

        dados_usuario = usuarios[email_input]

        # PRIMEIRO ACESSO
        if dados_usuario["senha"] is None:
            if not senha_input or len(senha_input) < 6:
                st.warning("⚠️ **Primeiro Acesso:** Defina uma senha de no mínimo 6 caracteres e clique em entrar novamente.")
            else:
                usuarios[email_input]["senha"] = gerar_hash(senha_input)
                salvar_usuarios(usuarios)
                
                st.session_state["logado"] = True
                st.session_state["usuario_atual"] = email_input
                st.success("🎉 Primeiro acesso registrado com sucesso! Entrando...")
                st.rerun()

        # LOGIN NORMAL
        else:
            if gerar_hash(senha_input) == dados_usuario["senha"]:
                st.session_state["logado"] = True
                st.session_state["usuario_atual"] = email_input
                st.success("Login efetuado com sucesso!")
                st.rerun()
            else:
                st.error("Senha incorreta. Tente novamente.")


# --- ABA ADMINISTRAÇÃO ---
def renderizar_aba_admin():
    st.header("⚙️ Painel de Administração")
    
    usuarios = carregar_usuarios()

    # 1. Tabela de Usuários Cadastrados
    st.subheader("👥 Lista de Usuários e Status de Acesso")
    
    dados_tabela = []
    for email, dados in usuarios.items():
        dados_tabela.append({
            "E-mail": email,
            "Loja / Unidade": dados.get("loja", "N/A"),
            "Perfil": "Administrador" if email in ADMINS else "Gerente",
            "Primeiro Acesso": "✅ Concluído" if dados.get("senha") else "⏳ Pendente",
            "Hash da Senha": (dados.get("senha")[:12] + "...") if dados.get("senha") else "Sem senha"
        })
    
    st.dataframe(dados_tabela, use_container_width=True)

    st.markdown("---")

    # 2. Reset de Senha pelo Admin
    st.subheader("🔑 Redefinir Senha de Usuário")
    st.caption("Escolha um usuário para redefinir a senha caso ele esqueça.")

    col1, col2 = st.columns([2, 1])

    with col1:
        usuario_selecionado = st.selectbox(
            "Selecione o e-mail:",
            options=list(usuarios.keys())
        )
        
        nova_senha = st.text_input(
            "Nova Senha Temporária:", 
            type="password", 
            key="input_nova_senha"
        )

    with col2:
        st.write("##") # Alinhamento visual
        if st.button("Redefinir Senha", type="primary"):
            if not nova_senha or len(nova_senha) < 6:
                st.error("A nova senha deve ter pelo menos 6 caracteres.")
            else:
                usuarios[usuario_selecionado]["senha"] = gerar_hash(nova_senha)
                salvar_usuarios(usuarios)
                st.success(f"✅ Senha do e-mail **{usuario_selecionado}** redefinida com sucesso!")


# --- DASHBOARD PRINCIPAL ---
def renderizar_dashboard():
    usuarios = carregar_usuarios()
    email_logado = st.session_state["usuario_atual"]
    loja_usuario = usuarios[email_logado].get("loja", "N/A")

    st.title("📊 Dashboard Vonny Cosméticos")
    st.write(f"Bem-vindo(a), **{email_logado}** | Loja Vinculada: **{loja_usuario}**")

    st.markdown("---")
    
    # Exemplo de filtragem por loja
    if loja_usuario == "TODAS":
        st.info("ℹ️ **Visão Geral:** Como Administrador, você possui visão consolidada de todas as lojas.")
    else:
        st.info(f"ℹ️ Exibindo métricas exclusivas da loja **{loja_usuario}**.")

    col1, col2, col3 = st.columns(3)
    col1.metric("Unidade", loja_usuario)
    col2.metric("Vendas do Mês", "R$ 45.200", "+8%")
    col3.metric("Meta Alcançada", "92%", "+4%")


# --- FLUXO PRINCIPAL ---
if not st.session_state["logado"]:
    renderizar_tela_login()
else:
    # Sidebar
    st.sidebar.write(f"👤 **Usuário:** {st.session_state['usuario_atual']}")
    
    if st.sidebar.button("Sair / Logout"):
        st.session_state["logado"] = False
        st.session_state["usuario_atual"] = None
        st.rerun()

    usuario_logado = st.session_state["usuario_atual"]
    
    # Exibe a aba do Admin apenas para administradores
    if usuario_logado in ADMINS:
        aba_dash, aba_admin = st.tabs(["📊 Dashboard", "⚙️ Painel Admin"])
        
        with aba_dash:
            renderizar_dashboard()
            
        with aba_admin:
            renderizar_aba_admin()
    else:
        # Gerentes normais visualizam apenas o Dashboard
        renderizar_dashboard()