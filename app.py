import streamlit as st
import pandas as pd
import plotly.express as px
import json
import os
import hashlib
import requests
import io
import datetime

# --- CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(
    page_title="Dashboard Executivo de Inventário - Vonny Cosméticos",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# --- ESTILIZAÇÃO VISUAL (TEMA ESCURO) ---
st.markdown("""
    <style>
    .stApp {
        background-color: #0e1117;
        color: #ffffff;
    }
    div[data-testid="stMetricValue"] {
        font-size: 1.8rem !important;
        font-weight: bold;
    }
    </style>
""", unsafe_allow_html=True)

URL_EXCEL_NUVEM = "https://vonnycosmeticos-my.sharepoint.com/:x:/g/personal/josue_pereira_vonnycosmeticos_onmicrosoft_com/IQAVAJHO0KlcS73eMCZZkJMEAdrs0fKrEhefibx1ieyMW_Y?e=B8QkcG&download=1"

DB_FILE = "usuarios_db.json"

# ==============================================================================
# MAPEAMENTO EXATO DAS REGIONAIS
# ==============================================================================
CENTROS_REGIONAL_1 = ["B013", "B015", "B016", "B017", "B019", "B020", "B021", "B022", "B023", "B024", "B025", "B026", "B027", "B028", "B029", "B031", "B032"]
CENTROS_REGIONAL_2 = ["B001", "B002", "B006", "B007", "B008", "B009", "B010", "B011", "B012", "B018", "B030"]

STR_REGIONAL_1 = "B013,B015,B016,B017,B019,B020,B021,B022,B023,B024,B025,B026,B027,B028,B029,B031,B032"
STR_REGIONAL_2 = "B001,B002,B006,B007,B008,B009,B010,B011,B012,B018,B030"

EMAILS_PERMITIDOS_PADRAO = {
    "sara.leite@vonnycosmeticos.com.br": ("B001", "Gerente"),
    "julio.fonseca@vonnycosmeticos.com.br": ("B002", "Gerente"),
    "fabiana.bertassi@vonnycosmeticos.com.br": ("B006", "Gerente"),
    "vanessa.tais@vonnycosmeticos.com.br": ("B007", "Gerente"),
    "yara.silva@vonnycosmeticos.com.br": ("B008", "Gerente"),
    "josemary.bezerra@vonnycosmeticos.com.br": ("B009", "Gerente"),
    "maria.beserra@vonnycosmeticos.com.br": ("B010", "Gerente"),
    "gislaine.barra@vonnycosmeticos.com.br": ("B011", "Gerente"),
    "thamires.conceicao@vonnycosmeticos.com.br": ("B012", "Gerente"),
    "vera.silva@vonnycosmeticos.com.br": ("B013", "Gerente"),
    "vanessa.amaral@vonnycosmeticos.com.br": ("B015", "Gerente"),
    "claudineia.mendes@vonnycosmeticos.com.br": ("B016", "Gerente"),
    "thatiane.ferreira@vonnycosmeticos.com.br": ("B017", "Gerente"),
    "katiane.silva@vonnycosmeticos.com.br": ("B018", "Gerente"),
    "lanny.andryelly@vonnycosmeticos.com.br": ("B019", "Gerente"),
    "suzana.silveira@vonnycosmeticos.com.br": ("B020", "Gerente"),
    "luciana.vasconcelos@vonnycosmeticos.com.br": ("B021", "Gerente"),
    "wagner.valle@casadolojista.com.br": (STR_REGIONAL_2, "Regional 2"),
    "daiane.martins@vonnycosmeticos.com.br": ("B022", "Gerente"),
    "gisele.trampusch@vonnycosmeticos.com.br": ("B023", "Gerente"),
    "raquel.lopes@vonnycosmeticos.com.br": ("B024", "Gerente"),
    "claudinea.santos@vonnycosmeticos.com.br": ("B025", "Gerente"),
    "rosania.chagas@vonnycosmeticos.com.br": ("B026", "Gerente"),
    "luana.costa@vonnycosmeticos.com.br": ("B027", "Gerente"),
    "rosangela.botelho@vonnycosmeticos.com.br": ("B028", "Gerente"),
    "elza.silva@vonnycosmeticos.com.br": ("B029", "Gerente"),
    "joao.pereira@vonnycosmeticos.com.br": ("B030", "Gerente"),
    "jorgiane.aragao@vonnycosmeticos.com.br": ("B031", "Gerente"),
    "jvn221106@gmail.com": ("TODAS", "Administrador"),
    "sergio.oliveira@vonnycosmeticos.com.br": ("TODAS", "Administrador"),
    "controladoriaprevencao@gmail.com": ("TODAS", "Administrador"),
    "josue.victor@vonnycosmeticos.com.br": ("TODAS", "Administrador"),
    "vanusia.garcia@casadolojista.com.br": ("TODAS", "Administrador"),
    "luciana.valle@vonnycosmeticos.com.br": (STR_REGIONAL_1, "Regional 1")
}

OPCOES_PERFIL = ["Gerente", "Líder de Loja", "Regional 1", "Regional 2", "Administrador"]

# --- FUNÇÕES DE BANCO DE DADOS E AUTENTICAÇÃO ---
def carregar_dados_db():
    if os.path.exists(DB_FILE):
        with open(DB_FILE, "r") as f:
            data = json.load(f)
            if "usuarios" in data:
                usuarios = data["usuarios"]
                removidos = set(data.get("removidos", []))
                historico_remocoes = data.get("historico_remocoes", [])
            else:
                usuarios = data
                removidos = set()
                historico_remocoes = []
    else:
        usuarios = {}
        removidos = set()
        historico_remocoes = []

    atualizou = False
    for email, (loja, perfil_padrao) in EMAILS_PERMITIDOS_PADRAO.items():
        email_limpo = email.strip().lower()
        if email_limpo not in usuarios and email_limpo not in removidos:
            usuarios[email_limpo] = {
                "loja": loja,
                "perfil": perfil_padrao,
                "senha": None,
                "forcar_redefinicao": False,
                "historico_senhas": []
            }
            atualizou = True
        elif email_limpo in usuarios:
            if "perfil" not in usuarios[email_limpo]:
                usuarios[email_limpo]["perfil"] = perfil_padrao
                atualizou = True
            if "historico_senhas" not in usuarios[email_limpo]:
                usuarios[email_limpo]["historico_senhas"] = []
                atualizou = True

    if atualizou or not os.path.exists(DB_FILE):
        salvar_dados_db(usuarios, removidos, historico_remocoes)

    return usuarios, removidos, historico_remocoes

def salvar_dados_db(usuarios, removidos, historico_remocoes=None):
    if historico_remocoes is None:
        historico_remocoes = []
    with open(DB_FILE, "w") as f:
        json.dump({
            "usuarios": usuarios,
            "removidos": list(removidos),
            "historico_remocoes": historico_remocoes
        }, f, indent=4)

def gerar_hash(senha):
    return hashlib.sha256(senha.encode()).hexdigest()

def atualizar_senha_com_historico(email, nova_senha_texto, usuarios_dict, removidos_set, historico_remocoes):
    novo_hash = gerar_hash(nova_senha_texto)
    dados_usr = usuarios_dict[email]
    
    historico = dados_usr.get("historico_senhas", [])
    
    if novo_hash == dados_usr.get("senha") or novo_hash in historico:
        return False, "⚠️ Por motivos de segurança, você não pode reutilizar nenhuma das suas últimas 3 senhas."
    
    if dados_usr.get("senha"):
        historico.insert(0, dados_usr["senha"])
    
    dados_usr["historico_senhas"] = historico[:3]
    dados_usr["senha"] = novo_hash
    dados_usr["forcar_redefinicao"] = False
    
    salvar_dados_db(usuarios_dict, removidos_set, historico_remocoes)
    return True, "✅ Senha alterada com sucesso!"

# --- LEITURA E TRATAMENTO DA PLANILHA NUVEM ---
@st.cache_data(ttl=60)
def load_data():
    response = requests.get(URL_EXCEL_NUVEM)
    response.raise_for_status()
    
    excel_file = io.BytesIO(response.content)
    xls = pd.ExcelFile(excel_file)
    df = pd.read_excel(xls, sheet_name="VALORES INVENTÁRIOS")
    df.columns = [str(col).strip() for col in df.columns]
    
    def achar_coluna(df_target, termos_prioritarios):
        for termo in termos_prioritarios:
            for col in df_target.columns:
                if termo.lower() in str(col).lower():
                    return col
        return None

    col_qtd = achar_coluna(df, ['qtd. um registro', 'qtd', 'registro'])
    col_valor = achar_coluna(df, ['montante em mi', 'montante', 'mi'])
    col_loja = achar_coluna(df, ['centro'])
    col_marca = achar_coluna(df, ['fornecedor2', 'fornecedor', 'marca'])

    if not col_qtd: col_qtd = df.columns[0]
    if not col_valor: col_valor = df.columns[1]
    if not col_loja: col_loja = df.columns[2]
    if not col_marca: col_marca = df.columns[3]

    df['Qtd_Limpa'] = pd.to_numeric(df[col_qtd], errors='coerce').fillna(0.0)
    df['Valor_Limpo'] = pd.to_numeric(df[col_valor], errors='coerce').fillna(0.0)
    
    df['Loja_Nome'] = df[col_loja].astype(str).fillna('').str.strip()
    df['Loja_Nome'] = df['Loja_Nome'].replace(['nan', 'None', 'NaN', 'none', ''], 'S/ Centro')
    
    df['Marca_Nome'] = df[col_marca].astype(str).fillna('').str.strip()
    df['Marca_Nome'] = df['Marca_Nome'].replace(['nan', 'None', 'NaN', 'none', ''], 'Sem Marca')

    df = df[(df['Loja_Nome'] != 'S/ Centro') & (df['Marca_Nome'] != 'Sem Marca')].copy()

    def classificar_centro(centro):
        c = str(centro).strip().upper()
        if c in CENTROS_REGIONAL_1:
            return 'Regional 1'
        elif c in CENTROS_REGIONAL_2:
            return 'Regional 2'
        else:
            return 'Sem Regional'

    df['Regional_Nome'] = df['Loja_Nome'].apply(classificar_centro)

    return df

def formatar_moeda(val):
    if val < 0:
        return f"-R$ {abs(val):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    else:
        return f"R$ {val:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

def formatar_qtd(val):
    if val < 0:
        return f"-{abs(val):,.0f} UN".replace(",", ".")
    else:
        return f"{val:,.0f} UN".replace(",", ".")

# --- GERENCIAMENTO DE SESSÃO ---
if "logado" not in st.session_state:
    st.session_state["logado"] = False
if "usuario_atual" not in st.session_state:
    st.session_state["usuario_atual"] = None
if "troca_obrigatoria" not in st.session_state:
    st.session_state["troca_obrigatoria"] = False

# --- TELA DE LOGIN ---
def renderizar_tela_login():
    st.title("🔒 Vonny Cosméticos - Acesso ao Sistema")
    st.write("Digite seu e-mail corporativo para acessar os indicadores.")

    usuarios, removidos, historico_remocoes = carregar_dados_db()

    with st.form("form_login"):
        email_input = st.text_input("E-mail corporativo:").strip().lower()
        senha_input = st.text_input("Senha:", type="password")
        btn_entrar = st.form_submit_button("Entrar", type="primary")

    if btn_entrar:
        if not email_input:
            st.error("Por favor, digite seu e-mail.")
            return

        if email_input in removidos or email_input not in usuarios:
            st.error("E-mail não autorizado ou acesso revogado. Entre em contato com o administrador.")
            return

        dados_usuario = usuarios[email_input]

        if dados_usuario["senha"] is None:
            if not senha_input or len(senha_input) < 6:
                st.warning("⚠️ **Primeiro Acesso:** Defina uma senha de no mínimo 6 caracteres e clique em entrar novamente.")
            else:
                sucesso, msg = atualizar_senha_com_historico(
                    email_input, senha_input, usuarios, removidos, historico_remocoes
                )
                if sucesso:
                    st.session_state["logado"] = True
                    st.session_state["usuario_atual"] = email_input
                    st.success("🎉 Primeiro acesso realizado! Entrando...")
                    st.rerun()
                else:
                    st.error(msg)
        else:
            if gerar_hash(senha_input) == dados_usuario["senha"]:
                st.session_state["logado"] = True
                st.session_state["usuario_atual"] = email_input
                
                if dados_usuario.get("forcar_redefinicao", False):
                    st.session_state["troca_obrigatoria"] = True
                
                st.success("Login efetuado!")
                st.rerun()
            else:
                st.error("Senha incorreta.")

    st.markdown("---")
    with st.expander("❓ Esqueceu a senha?"):
        st.info("📩 Por favor, abra um chamado para o setor de **Controladoria / Prevenção de Perdas** solicitando a redefinição de senha.")

# --- TELA OBRIGATÓRIA DE REDEFINIÇÃO DE SENHA ---
def renderizar_tela_troca_obrigatoria():
    st.title("🔑 Redefinição de Senha Obrigatória")
    st.warning("Você acessou com uma **senha temporária**. Escolha uma nova senha definitiva para continuar.")

    usuarios, removidos, historico_remocoes = carregar_dados_db()
    email_logado = st.session_state["usuario_atual"]

    with st.form("form_troca_obrigatoria"):
        nova_senha = st.text_input("Nova Senha (mínimo 6 caracteres):", type="password")
        confirma_nova = st.text_input("Confirme a Nova Senha:", type="password")
        btn_salvar = st.form_submit_button("Salvar Nova Senha", type="primary")

    if btn_salvar:
        if len(nova_senha) < 6:
            st.error("A nova senha deve ter no mínimo 6 caracteres.")
            return

        if nova_senha != confirma_nova:
            st.error("As senhas não coincidem.")
            return

        sucesso, msg = atualizar_senha_com_historico(
            email_logado, nova_senha, usuarios, removidos, historico_remocoes
        )
        if sucesso:
            st.session_state["troca_obrigatoria"] = False
            st.success("✅ Senha atualizada com sucesso!")
            st.rerun()
        else:
            st.error(msg)

# --- ABA PAINEL ADMIN ---
def renderizar_aba_admin():
    st.header("⚙️ Painel do Administrador")
    usuarios, removidos, historico_remocoes = carregar_dados_db()

    st.subheader("👥 Lista de Usuários e Status")
    dados_tabela = []
    for email, dados in usuarios.items():
        dados_tabela.append({
            "E-mail": email,
            "Loja / Centro": dados.get("loja", "N/A"),
            "Perfil / Cargo": dados.get("perfil", "Gerente"),
            "Primeiro Acesso": "✅ Concluído" if dados.get("senha") else "⏳ Pendente",
            "Senha Temporária Ativa": "⚠️ Sim" if dados.get("forcar_redefinicao") else "Não"
        })
    st.dataframe(dados_tabela, use_container_width=True)

    st.markdown("---")

    st.subheader("➕ Adicionar ou Editar Usuário")
    col_add1, col_add2, col_add3, col_add4 = st.columns([2, 2, 1, 1])
    with col_add1:
        novo_email = st.text_input("E-mail corporativo:", key="input_novo_email").strip().lower()
    with col_add2:
        nova_loja = st.text_input("Centro (Ex: B001,B002 ou TODAS):", key="input_nova_loja").strip().upper()
    with col_add3:
        novo_perfil = st.selectbox("Perfil / Cargo:", options=OPCOES_PERFIL, key="select_novo_perfil")
    with col_add4:
        st.write("##")
        if st.button("Salvar Usuário", type="primary"):
            if not novo_email or "@" not in novo_email:
                st.error("Por favor, digite um e-mail válido.")
            elif not nova_loja:
                st.error("Por favor, informe a loja / centro.")
            else:
                if novo_email in removidos:
                    removidos.remove(novo_email)

                if novo_email in usuarios:
                    usuarios[novo_email]["loja"] = nova_loja
                    usuarios[novo_email]["perfil"] = novo_perfil
                    st.success(f"✅ Usuário **{novo_email}** atualizado!")
                else:
                    usuarios[novo_email] = {
                        "loja": nova_loja,
                        "perfil": novo_perfil,
                        "senha": None,
                        "forcar_redefinicao": False,
                        "historico_senhas": []
                    }
                    st.success(f"🎉 Usuário **{novo_email}** cadastrado!")
                
                salvar_dados_db(usuarios, removidos, historico_remocoes)
                st.rerun()

    st.markdown("---")

    st.subheader("🔑 Resetar Senha / Gerar Senha Temporária")
    col1, col2 = st.columns([2, 1])
    with col1:
        usuario_selecionado = st.selectbox("Selecione o e-mail:", options=list(usuarios.keys()), key="select_reset_senha")
        senha_temp = st.text_input("Senha Temporária:", type="password", key="input_senha_temp")

    with col2:
        st.write("##")
        if st.button("Definir Senha Temporária"):
            if not senha_temp or len(senha_temp) < 6:
                st.error("A senha deve ter pelo menos 6 caracteres.")
            else:
                usuarios[usuario_selecionado]["senha"] = gerar_hash(senha_temp)
                usuarios[usuario_selecionado]["forcar_redefinicao"] = True
                salvar_dados_db(usuarios, removidos, historico_remocoes)
                st.success(f"✅ Senha temporária definida para **{usuario_selecionado}**!")

    st.markdown("---")

    st.subheader("🗑️ Remover Usuário Permanentemente")
    col_del1, col_del2 = st.columns([2, 1])
    with col_del1:
        user_para_deletar = st.selectbox("Selecione para remover:", options=list(usuarios.keys()), key="select_del_user")
    with col_del2:
        st.write("##")
        if st.button("Remover Usuário", type="secondary"):
            admin_atual = st.session_state["usuario_atual"]
            
            if user_para_deletar == admin_atual:
                st.error("Você não pode remover seu próprio usuário logado.")
            else:
                registro_log = {
                    "usuario_removido": user_para_deletar,
                    "removido_por": admin_atual,
                    "data_hora": datetime.datetime.now().strftime("%d/%m/%Y %H:%M:%S")
                }
                historico_remocoes.append(registro_log)

                del usuarios[user_para_deletar]
                removidos.add(user_para_deletar)

                salvar_dados_db(usuarios, removidos, historico_remocoes)
                st.success(f"🗑️ Usuário **{user_para_deletar}** removido por **{admin_atual}**!")
                st.rerun()

    st.markdown("---")

    st.subheader("📋 Histórico de Remoções (Auditoria)")
    if historico_remocoes:
        df_historico = pd.DataFrame(historico_remocoes)
        df_historico.rename(columns={
            "usuario_removido": "Usuário Removido",
            "removido_por": "Removido por (Admin)",
            "data_hora": "Data e Horário"
        }, inplace=True)
        st.dataframe(df_historico, use_container_width=True, hide_index=True)
    else:
        st.info("Nenhuma remoção registrada até o momento.")

# --- DASHBOARD VISUAL DE INVENTÁRIO ---
def renderizar_dashboard():
    try:
        usuarios, _, _ = carregar_dados_db()
        email_logado = st.session_state["usuario_atual"]
        dados_usr = usuarios.get(email_logado, {})
        
        loja_usuario = dados_usr.get("loja", "N/A")
        perfil_usuario = dados_usr.get("perfil", "Gerente")

        df = load_data()

        st.sidebar.title("Filtros")

        if st.sidebar.button("🔄 Atualizar Dados"):
            st.cache_data.clear()
            st.rerun()

        regionais_disponiveis = [str(r) for r in ["Regional 1", "Regional 2"] if r in df['Regional_Nome'].unique()]

        if perfil_usuario == "Administrador":
            regionais_sel = st.sidebar.multiselect(
                "Selecione a Divisão Regional:", 
                options=regionais_disponiveis, 
                default=regionais_disponiveis
            )
        elif perfil_usuario in ["Regional 1", "Regional 2"]:
            regionais_sel = [perfil_usuario]
        else:
            regionais_sel = regionais_disponiveis

        lojas_unicas = [str(x) for x in df[df['Regional_Nome'].isin(regionais_sel)]['Loja_Nome'].unique() if str(x).lower() not in ['nan', 'none', '', 'sem centro', 's/ centro']]
        lojas_disponiveis = sorted(lojas_unicas)

        if perfil_usuario == "Administrador":
            lojas_sel = st.sidebar.multiselect(
                "Selecione os Centros:", 
                options=lojas_disponiveis, 
                default=lojas_disponiveis
            )
        elif perfil_usuario in ["Regional 1", "Regional 2"]:
            lojas_permitidas_usr = [x.strip() for x in str(loja_usuario).replace(" ", ",").split(",") if x.strip()]
            lojas_filtradas_usr = [x for x in lojas_disponiveis if x in lojas_permitidas_usr] if lojas_permitidas_usr else lojas_disponiveis
            lojas_sel = st.sidebar.multiselect(
                "Selecione os Centros:", 
                options=lojas_filtradas_usr, 
                default=lojas_filtradas_usr
            )
        else:
            lojas_permitidas_usr = [x.strip() for x in str(loja_usuario).replace(" ", ",").split(",") if x.strip()]
            lojas_sel = [x for x in lojas_disponiveis if x in lojas_permitidas_usr]
            if not lojas_sel:
                lojas_sel = lojas_disponiveis
            st.sidebar.info(f"📍 **Centro Vinculado:** {', '.join(lojas_sel)}")

        marcas_unicas = [str(x) for x in df['Marca_Nome'].unique() if str(x).lower() not in ['nan', 'none', '', 'sem marca']]
        marcas = sorted(marcas_unicas)
        marcas_sel = st.sidebar.multiselect("Selecione as Marcas:", options=marcas, default=marcas)

        df_filtered = df[
            (df['Regional_Nome'].isin(regionais_sel)) &
            (df['Loja_Nome'].isin(lojas_sel)) &
            (df['Marca_Nome'].isin(marcas_sel))
        ]

        st.title("📊 Dashboard Executivo de Inventário")
        st.markdown(f"**Usuário:** `{email_logado}` | **Perfil:** `{perfil_usuario}`")
        st.markdown("---")

        perda_total_rs = float(df_filtered[df_filtered['Valor_Limpo'] < 0]['Valor_Limpo'].sum())
        perda_total_un = float(df_filtered[df_filtered['Qtd_Limpa'] < 0]['Qtd_Limpa'].sum())
        sobra_total_rs = float(df_filtered[df_filtered['Valor_Limpo'] > 0]['Valor_Limpo'].sum())
        resultado_net = sobra_total_rs + perda_total_rs

        kpi1, kpi2, kpi3, kpi4 = st.columns(4)
        kpi1.metric("Total de Perdas (Qtd)", formatar_qtd(perda_total_un))
        kpi2.metric("Perda Total (R$)", formatar_moeda(perda_total_rs))
        kpi3.metric("Sobras / Ajustes (+)", formatar_moeda(sobra_total_rs))
        kpi4.metric("Resultado Net (Caixa)", formatar_moeda(resultado_net))

        st.markdown("<br>", unsafe_allow_html=True)

        if perfil_usuario == "Administrador":
            st.subheader("🗺️ Comparativo por Divisão Regional (Regional 1 vs Regional 2)")

            df_reg_comp = (
                df_filtered[df_filtered['Valor_Limpo'] < 0]
                .groupby('Regional_Nome')
                .agg({
                    'Qtd_Limpa': lambda x: abs(x.sum()),
                    'Valor_Limpo': lambda x: abs(x.sum())
                })
                .reset_index()
                .sort_values(by='Valor_Limpo', ascending=False)
            )
            df_reg_comp['Texto_Valor'] = df_reg_comp['Valor_Limpo'].apply(lambda x: f"-R$ {x:,.2f}")

            fig_reg_comp = px.bar(
                df_reg_comp,
                x='Regional_Nome',
                y='Valor_Limpo',
                text='Texto_Valor',
                color='Regional_Nome',
                title="Comparativo por Divisão Regional",
                color_discrete_map={
                    'Regional 1': '#4ba3e3',
                    'Regional 2': '#ff7f0e',
                },
                labels={'Valor_Limpo': 'Perda (R$)', 'Regional_Nome': 'Divisão Regional'}
            )
            fig_reg_comp.update_traces(textposition='inside')
            fig_reg_comp.update_layout(
                template="plotly_dark",
                paper_bgcolor="rgba(0,0,0,0)",
                plot_bgcolor="rgba(0,0,0,0)",
                xaxis_title="",
                yaxis_title="",
                showlegend=False
            )
            st.plotly_chart(fig_reg_comp, use_container_width=True)
            st.markdown("<br>", unsafe_allow_html=True)

        if perfil_usuario in ["Administrador", "Regional 1", "Regional 2"]: 
            graf_col1, graf_col2 = st.columns(2)

            with graf_col1:
                st.subheader("📦 Perda por Centro (Qtd)")
                df_qtd_lojas = (
                    df_filtered[df_filtered['Qtd_Limpa'] < 0]
                    .groupby('Loja_Nome')['Qtd_Limpa']
                    .sum()
                    .abs()
                    .reset_index()
                    .sort_values(by='Qtd_Limpa', ascending=False)
                )
                df_qtd_lojas['Texto_Qtd'] = df_qtd_lojas['Qtd_Limpa'].apply(lambda x: f"-{x:,.0f} un")

                fig_qtd_lojas = px.bar(
                    df_qtd_lojas,
                    x='Loja_Nome',
                    y='Qtd_Limpa',
                    text='Texto_Qtd',
                    title="Perda por Centro (Qtd)",
                    labels={'Qtd_Limpa': 'Perda (Qtd)', 'Loja_Nome': 'Centro'}
                )
                fig_qtd_lojas.update_traces(marker_color='#4ba3e3', textposition='inside')
                fig_qtd_lojas.update_layout(
                    template="plotly_dark",
                    paper_bgcolor="rgba(0,0,0,0)",
                    plot_bgcolor="rgba(0,0,0,0)",
                    xaxis_title="",
                    yaxis_title=""
                )
                st.plotly_chart(fig_qtd_lojas, use_container_width=True)

            with graf_col2:
                st.subheader("🎯 Perda por Centro (R$)")
                df_lojas = (
                    df_filtered[df_filtered['Valor_Limpo'] < 0]
                    .groupby('Loja_Nome')['Valor_Limpo']
                    .sum()
                    .abs()
                    .reset_index()
                    .sort_values(by='Valor_Limpo', ascending=False)
                )
                df_lojas['Texto_Valor'] = df_lojas['Valor_Limpo'].apply(lambda x: f"-{x:,.2f}")

                fig_lojas = px.bar(
                    df_lojas,
                    x='Loja_Nome',
                    y='Valor_Limpo',
                    text='Texto_Valor',
                    title="Perda por Centro (R$)",
                    labels={'Valor_Limpo': 'Perda (R$)', 'Loja_Nome': 'Centro'}
                )
                fig_lojas.update_traces(marker_color='#70bbfd', textposition='inside')
                fig_lojas.update_layout(
                    template="plotly_dark",
                    paper_bgcolor="rgba(0,0,0,0)",
                    plot_bgcolor="rgba(0,0,0,0)",
                    xaxis_title="",
                    yaxis_title=""
                )
                st.plotly_chart(fig_lojas, use_container_width=True)

            st.markdown("<br>", unsafe_allow_html=True)
            st.subheader("🏢 Ranking: Top 10 Centros com Maior Perda")

            df_centros_completo = (
                df_filtered[(df_filtered['Valor_Limpo'] < 0) | (df_filtered['Qtd_Limpa'] < 0)]
                .groupby(['Loja_Nome', 'Regional_Nome'])
                .agg({
                    'Qtd_Limpa': lambda x: abs(x[x < 0].sum()),
                    'Valor_Limpo': lambda x: abs(x[x < 0].sum())
                })
                .reset_index()
                .sort_values(by='Valor_Limpo', ascending=False)
            )

            df_centros_completo.insert(0, 'Posição', [f"{i+1}º" for i in range(len(df_centros_completo))])
            df_centros_completo = df_centros_completo[['Posição', 'Loja_Nome', 'Regional_Nome', 'Qtd_Limpa', 'Valor_Limpo']]
            df_centros_completo.rename(columns={'Loja_Nome': 'Centro', 'Regional_Nome': 'Divisão Regional', 'Qtd_Limpa': 'Perda (Qtd)', 'Valor_Limpo': 'Perda (R$)'}, inplace=True)

            df_top10_centros = df_centros_completo.head(10).copy()
            df_top10_centros['Perda (Qtd)'] = df_top10_centros['Perda (Qtd)'].apply(lambda x: f"-{x:,.0f} un")
            df_top10_centros['Perda (R$)'] = df_top10_centros['Perda (R$)'].apply(lambda x: f"R$ -{x:,.2f}")

            st.dataframe(df_top10_centros, use_container_width=True, hide_index=True)

        st.markdown("<br>", unsafe_allow_html=True)
        st.subheader("⚠️ Ranking: Top 10 Marcas com Maior Perda")

        df_marcas_completo = (
            df_filtered[(df_filtered['Valor_Limpo'] < 0) | (df_filtered['Qtd_Limpa'] < 0)]
            .groupby('Marca_Nome')
            .agg({
                'Qtd_Limpa': lambda x: abs(x[x < 0].sum()),
                'Valor_Limpo': lambda x: abs(x[x < 0].sum())
            })
            .reset_index()
            .sort_values(by='Valor_Limpo', ascending=False)
        )

        df_marcas_completo.insert(0, 'Posição', [f"{i+1}º" for i in range(len(df_marcas_completo))])
        df_marcas_completo = df_marcas_completo[['Posição', 'Marca_Nome', 'Qtd_Limpa', 'Valor_Limpo']]
        df_marcas_completo.rename(columns={'Marca_Nome': 'Marca', 'Qtd_Limpa': 'Perda (Qtd)', 'Valor_Limpo': 'Perda (R$)'}, inplace=True)

        df_top10_marcas = df_marcas_completo.head(10).copy()
        df_top10_marcas['Perda (Qtd)'] = df_top10_marcas['Perda (Qtd)'].apply(lambda x: f"-{x:,.0f} un")
        df_top10_marcas['Perda (R$)'] = df_top10_marcas['Perda (R$)'].apply(lambda x: f"R$ -{x:,.2f}")

        st.dataframe(df_top10_marcas, use_container_width=True, hide_index=True)

        st.markdown("<br>", unsafe_allow_html=True)
        marca_col1, marca_col2 = st.columns(2)

        df_perdas_marcas = df_filtered[(df_filtered['Valor_Limpo'] < 0) | (df_filtered['Qtd_Limpa'] < 0)]

        with marca_col1:
            st.subheader("📦 Perdas por Marca - Todas (Qtd)")
            df_marca_qtd = (
                df_perdas_marcas[df_perdas_marcas['Qtd_Limpa'] < 0]
                .groupby('Marca_Nome')['Qtd_Limpa']
                .sum()
                .abs()
                .reset_index()
                .sort_values(by='Qtd_Limpa', ascending=False)
            )
            df_marca_qtd['Texto_Qtd'] = df_marca_qtd['Qtd_Limpa'].apply(lambda x: f"-{x:,.0f} un")

            fig_marca_qtd = px.line(
                df_marca_qtd,
                x='Marca_Nome',
                y='Qtd_Limpa',
                text='Texto_Qtd',
                markers=True,
                title="Perdas por Marca - Todas (Qtd)",
                labels={'Qtd_Limpa': 'Perda (Qtd)', 'Marca_Nome': 'Marca'}
            )
            fig_marca_qtd.update_traces(line_color='#ff7f0e', line_width=3, marker_size=7, textposition='top center')
            fig_marca_qtd.update_layout(
                template="plotly_dark",
                paper_bgcolor="rgba(0,0,0,0)",
                plot_bgcolor="rgba(0,0,0,0)",
                xaxis_title="",
                yaxis_title="",
                xaxis_tickangle=-45
            )
            st.plotly_chart(fig_marca_qtd, use_container_width=True)

        with marca_col2:
            st.subheader("🏷️ Perdas por Marca - Todas (R$)")
            df_marca_rs = (
                df_perdas_marcas[df_perdas_marcas['Valor_Limpo'] < 0]
                .groupby('Marca_Nome')['Valor_Limpo']
                .sum()
                .abs()
                .reset_index()
                .sort_values(by='Valor_Limpo', ascending=False)
            )
            df_marca_rs['Texto_RS'] = df_marca_rs['Valor_Limpo'].apply(lambda x: f"-{x:,.0f}")

            fig_marca_rs = px.line(
                df_marca_rs,
                x='Marca_Nome',
                y='Valor_Limpo',
                text='Texto_RS',
                markers=True,
                title="Perdas por Marca - Todas (R$)",
                labels={'Valor_Limpo': 'Perda (R$)', 'Marca_Nome': 'Marca'}
            )
            fig_marca_rs.update_traces(line_color='#4ba3e3', line_width=3, marker_size=7, textposition='top center')
            fig_marca_rs.update_layout(
                template="plotly_dark",
                paper_bgcolor="rgba(0,0,0,0)",
                plot_bgcolor="rgba(0,0,0,0)",
                xaxis_title="",
                yaxis_title="",
                xaxis_tickangle=-45
            )
            st.plotly_chart(fig_marca_rs, use_container_width=True)

    except Exception as e:
        st.error(f"Erro ao carregar os dados do arquivo Excel na nuvem: {e}")

# --- FLUXO PRINCIPAL DA APLICAÇÃO ---
if not st.session_state["logado"]:
    renderizar_tela_login()

elif st.session_state["troca_obrigatoria"]:
    renderizar_tela_troca_obrigatoria()

else:
    usuarios_db, removidos_set, historico_remocoes = carregar_dados_db()
    usr_atual = st.session_state["usuario_atual"]

    # --- REVALIDAÇÃO DE SESSÃO EM TEMPO REAL ---
    # Se o usuário foi removido do JSON ou está na lista de removidos, cancela a sessão imediatamente
    if usr_atual in removidos_set or usr_atual not in usuarios_db:
        st.session_state["logado"] = False
        st.session_state["usuario_atual"] = None
        st.session_state["troca_obrigatoria"] = False
        st.error("🔒 Sua conta foi desativada ou removida. Você foi desconectado.")
        st.rerun()

    dados_logado = usuarios_db.get(usr_atual, {})

    st.sidebar.markdown(f"👤 **Usuário:** `{usr_atual}`\n\n💼 **Cargo:** `{dados_logado.get('perfil', 'Gerente')}`")
    
    with st.sidebar.expander("🔑 Alterar minha senha"):
        with st.form("form_mudar_senha_sidebar"):
            senha_antiga_sb = st.text_input("Senha Atual:", type="password")
            nova_senha_sb = st.text_input("Nova Senha:", type="password")
            confirma_sb = st.text_input("Confirme a Nova Senha:", type="password")
            btn_mudar_sb = st.form_submit_button("Atualizar Senha")

            if btn_mudar_sb:
                usuarios_dict, removidos_set, historico_remocoes = carregar_dados_db()
                
                if gerar_hash(senha_antiga_sb) != usuarios_dict[usr_atual]["senha"]:
                    st.error("Senha atual incorreta.")
                elif len(nova_senha_sb) < 6:
                    st.error("Mínimo de 6 caracteres.")
                elif nova_senha_sb != confirma_sb:
                    st.error("Senhas não conferem.")
                else:
                    sucesso, msg = atualizar_senha_com_historico(
                        usr_atual, nova_senha_sb, usuarios_dict, removidos_set, historico_remocoes
                    )
                    if sucesso:
                        st.success(msg)
                    else:
                        st.error(msg)

    if st.sidebar.button("🚪 Sair / Logout"):
        st.session_state["logado"] = False
        st.session_state["usuario_atual"] = None
        st.session_state["troca_obrigatoria"] = False
        st.rerun()

    perfil_logado = dados_logado.get("perfil", "Gerente")

    if perfil_logado == "Administrador":
        aba_dash, aba_admin = st.tabs(["📊 Dashboard Geral", "⚙️ Painel Admin"])
        with aba_dash:
            renderizar_dashboard()
        with aba_admin:
            renderizar_aba_admin()
    else:
        renderizar_dashboard()