import streamlit as st
import pandas as pd
import plotly.express as px
import json
import os
import hashlib
import requests
import io

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

# Link do arquivo principal de Inventário
URL_EXCEL_NUVEM = "https://vonnycosmeticos-my.sharepoint.com/:x:/g/personal/josue_pereira_vonnycosmeticos_onmicrosoft_com/IQAVAJHO0KlcS73eMCZZkJMEAdrs0fKrEhefibx1ieyMW_Y?e=B8QkcG&download=1"

# LINK PARA O ARQUIVO DE DIVISÃO REGIONAL (se for arquivo separado)
URL_EXCEL_REGIONAL = "https://vonnycosmeticos-my.sharepoint.com/:x:/g/personal/josue_pereira_vonnycosmeticos_onmicrosoft_com/IQD4e3tUtaZxQJBi3c-RKo4GAcZgXBpTvtj8g25Q6jiJQYU?e=6zHbkR?download=1"

DB_FILE = "usuarios_db.json"

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
    "luciana.vasconcelos@vonnycosmeticos.com.br":("B013,B015,B016,B017,B019,B020,B021,B022,B023,B024,B025,B026,B027,B028,B029,B031,B032", "Regional 1"),
    "wagner.valle@casadolojista.com.br":("B001,B002,B006,B007,B008,B009,B010,B011,B012,B018,B030", "Regional 2"),
}

OPCOES_PERFIL = ["Gerente", "Líder de Loja", "Administrador", "Regional 1", "Regional 2"]


# --- PERSISTÊNCIA E USUÁRIOS (JSON) ---
def carregar_dados_db():
    if os.path.exists(DB_FILE):
        with open(DB_FILE, "r") as f:
            data = json.load(f)
            if "usuarios" in data:
                usuarios = data["usuarios"]
                removidos = set(data.get("removidos", []))
            else:
                usuarios = data
                removidos = set()
    else:
        usuarios = {}
        removidos = set()

    atualizou = False
    for email, (loja, perfil_padrao) in EMAILS_PERMITIDOS_PADRAO.items():
        email_limpo = email.strip().lower()
        if email_limpo not in usuarios and email_limpo not in removidos:
            usuarios[email_limpo] = {
                "loja": loja,
                "perfil": perfil_padrao,
                "senha": None,
                "forcar_redefinicao": False
            }
            atualizou = True
        elif email_limpo in usuarios and "perfil" not in usuarios[email_limpo]:
            usuarios[email_limpo]["perfil"] = perfil_padrao
            atualizou = True

    if atualizou or not os.path.exists(DB_FILE):
        salvar_dados_db(usuarios, removidos)

    return usuarios, removidos

def salvar_dados_db(usuarios, removidos):
    with open(DB_FILE, "w") as f:
        json.dump({
            "usuarios": usuarios,
            "removidos": list(removidos)
        }, f, indent=4)

def gerar_hash(senha):
    return hashlib.sha256(senha.encode()).hexdigest()


# --- LEITURA E TRATAMENTO DA PLANILHA NUVEM ---
@st.cache_data(ttl=60)
def load_data():
    # 1. Carregar Planilha de Inventário
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

    df['Qtd_Limpa'] = pd.to_numeric(df[col_qtd], errors='coerce').fillna(0)
    df['Valor_Limpo'] = pd.to_numeric(df[col_valor], errors='coerce').fillna(0)
    df['Loja_Nome'] = df[col_loja].fillna('S/ Centro').astype(str).str.strip()
    df['Marca_Nome'] = df[col_marca].fillna('Sem Marca').astype(str).str.strip()

    # 2. Carregar e Mapear Divisão Regional (Regional 1 - Luciana / Regional 2 - Wagner)
    try:
        if "PADRÃO_DIVISÃO REGIONAL" in xls.sheet_names:
            df_regional = pd.read_excel(xls, sheet_name="PADRÃO_DIVISÃO REGIONAL")
        elif "DIVISÃO REGIONAL" in xls.sheet_names:
            df_regional = pd.read_excel(xls, sheet_name="DIVISÃO REGIONAL")
        else:
            resp_reg = requests.get(URL_EXCEL_REGIONAL)
            df_regional = pd.read_excel(io.BytesIO(resp_reg.content))

        df_regional.columns = [str(col).strip() for col in df_regional.columns]
        col_reg_loja = achar_coluna(df_regional, ['centro', 'loja', 'filial'])
        col_reg_divisao = achar_coluna(df_regional, ['regional', 'divisao', 'divisão', 'gerente', 'responsavel'])

        if col_reg_loja and col_reg_divisao:
            df_regional['Loja_Key'] = df_regional[col_reg_loja].astype(str).str.strip()
            
            # Função para classificar em Regional 1 ou Regional 2
            def classificar_regional(val):
                val_str = str(val).lower()
                if 'luciana' in val_str or 'regional 1' in val_str or 'reg 1' in val_str or val_str == '1':
                    return 'Regional 1 - Luciana'
                elif 'wagner' in val_str or 'regional 2' in val_str or 'reg 2' in val_str or val_str == '2':
                    return 'Regional 2 - Wagner'
                else:
                    return f"Outros ({val})"

            df_regional['Regional_Nome'] = df_regional[col_reg_divisao].apply(classificar_regional)
            
            df = df.merge(df_regional[['Loja_Key', 'Regional_Nome']], left_on='Loja_Nome', right_on='Loja_Key', how='left')
            df['Regional_Nome'] = df['Regional_Nome'].fillna('Sem Regional')
            if 'Loja_Key' in df.columns:
                df.drop(columns=['Loja_Key'], inplace=True)
        else:
            df['Regional_Nome'] = 'Sem Regional'
    except Exception:
        df['Regional_Nome'] = 'Sem Regional'

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

    usuarios, removidos = carregar_dados_db()

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
                usuarios[email_input]["senha"] = gerar_hash(senha_input)
                usuarios[email_input]["forcar_redefinicao"] = False
                salvar_dados_db(usuarios, removidos)
                st.session_state["logado"] = True
                st.session_state["usuario_atual"] = email_input
                st.success("🎉 Primeiro acesso realizado! Entrando...")
                st.rerun()
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

    usuarios, removidos = carregar_dados_db()
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

        usuarios[email_logado]["senha"] = gerar_hash(nova_senha)
        usuarios[email_logado]["forcar_redefinicao"] = False
        salvar_dados_db(usuarios, removidos)

        st.session_state["troca_obrigatoria"] = False
        st.success("✅ Senha atualizada com sucesso!")
        st.rerun()


# --- ABA PAINEL ADMIN ---
def renderizar_aba_admin():
    st.header("⚙️ Painel do Administrador")
    usuarios, removidos = carregar_dados_db()

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
    col_add1, col_add2, col_add3, col_add4 = st.columns([2, 1, 1, 1])
    with col_add1:
        novo_email = st.text_input("E-mail corporativo:", key="input_novo_email").strip().lower()
    with col_add2:
        nova_loja = st.text_input("Centro (Ex: B001):", key="input_nova_loja").strip().upper()
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
                        "forcar_redefinicao": False
                    }
                    st.success(f"🎉 Usuário **{novo_email}** cadastrado!")
                
                salvar_dados_db(usuarios, removidos)
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
                salvar_dados_db(usuarios, removidos)
                st.success(f"✅ Senha temporária definida para **{usuario_selecionado}**!")

    st.markdown("---")

    st.subheader("🗑️ Remover Usuário Permanentemente")
    col_del1, col_del2 = st.columns([2, 1])
    with col_del1:
        user_para_deletar = st.selectbox("Selecione para remover:", options=list(usuarios.keys()), key="select_del_user")
    with col_del2:
        st.write("##")
        if st.button("Remover Usuário", type="secondary"):
            if user_para_deletar == st.session_state["usuario_atual"]:
                st.error("Você não pode remover seu próprio usuário logado.")
            else:
                del usuarios[user_para_deletar]
                removidos.add(user_para_deletar)
                salvar_dados_db(usuarios, removidos)
                st.success(f"🗑️ Usuário **{user_para_deletar}** removido permanentemente!")
                st.rerun()


# --- DASHBOARD VISUAL DE INVENTÁRIO ---
def renderizar_dashboard():
    usuarios, _ = carregar_dados_db()
    email_logado = st.session_state["usuario_atual"]
    dados_usr = usuarios.get(email_logado, {})
    
    loja_usuario = dados_usr.get("loja", "N/A")
    perfil_usuario = dados_usr.get("perfil", "Gerente")

    try:
        df = load_data()

        st.sidebar.title("Filtros")

        if st.sidebar.button("🔄 Atualizar Dados"):
            st.cache_data.clear()
            st.rerun()

        # Filtro de Divisões Regionais (Regional 1 - Luciana / Regional 2 - Wagner)
        regionais_disponiveis = sorted(df['Regional_Nome'].unique())
        regionais_sel = st.sidebar.multiselect(
            "Selecione a Divisão Regional:",
            options=regionais_disponiveis,
            default=regionais_disponiveis
        )

        lojas_disponiveis = [x for x in sorted(df[df['Regional_Nome'].isin(regionais_sel)]['Loja_Nome'].unique()) if x.lower() not in ['nan', 'none', '', 's/ centro']]
        
        if perfil_usuario == "Administrador":
            lojas_sel = st.sidebar.multiselect("Selecione os Centros:", options=lojas_disponiveis, default=lojas_disponiveis)
        else:
            if loja_usuario in lojas_disponiveis:
                lojas_sel = [loja_usuario]
                st.sidebar.info(f"📍 **Centro Vinculado:** {loja_usuario}")
            else:
                lojas_sel = lojas_disponiveis
                st.sidebar.warning(f"Centro {loja_usuario} não encontrado nos registros atuais.")

        marcas = [x for x in sorted(df['Marca_Nome'].unique()) if x.lower() not in ['nan', 'none', '', 'sem marca']]
        marcas_sel = st.sidebar.multiselect("Selecione as Marcas:", options=marcas, default=marcas)

        df_filtered = df[
            (df['Regional_Nome'].isin(regionais_sel)) &
            (df['Loja_Nome'].isin(lojas_sel)) &
            (df['Marca_Nome'].isin(marcas_sel))
        ]

        st.title("📊 Dashboard Executivo de Inventário")
        st.markdown(f"**Usuário:** `{email_logado}` | **Perfil:** `{perfil_usuario}` | **Loja:** `{loja_usuario}`")
        st.markdown("---")

        perda_total_rs = df_filtered[df_filtered['Valor_Limpo'] < 0]['Valor_Limpo'].sum()
        perda_total_un = df_filtered[df_filtered['Qtd_Limpa'] < 0]['Qtd_Limpa'].sum()
        sobra_total_rs = df_filtered[df_filtered['Valor_Limpo'] > 0]['Valor_Limpo'].sum()
        resultado_net = sobra_total_rs + perda_total_rs

        kpi1, kpi2, kpi3, kpi4 = st.columns(4)
        kpi1.metric("Total de Perdas (Qtd)", formatar_qtd(perda_total_un))
        kpi2.metric("Perda Total (R$)", formatar_moeda(perda_total_rs))
        kpi3.metric("Sobras / Ajustes (+)", formatar_moeda(sobra_total_rs))
        kpi4.metric("Resultado Net (Caixa)", formatar_moeda(resultado_net))

        st.markdown("<br>", unsafe_allow_html=True)

        # --- VISÃO EXCLUSIVA PARA ADMINISTRADORES ---
        if perfil_usuario == "Administrador":
            # Comparativo entre Regional 1 e Regional 2
            st.subheader("🗺️ Comparativo por Divisão Regional (Regional 1 - Luciana vs Regional 2 - Wagner)")
            
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
                color_discrete_map={
                    'Regional 1 - Luciana': '#4ba3e3',
                    'Regional 2 - Wagner': '#ff7f0e',
                    'Sem Regional': '#888888'
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

            df_top10_centros = (
                df_filtered[(df_filtered['Valor_Limpo'] < 0) | (df_filtered['Qtd_Limpa'] < 0)]
                .groupby(['Loja_Nome', 'Regional_Nome'])
                .agg({
                    'Qtd_Limpa': lambda x: abs(x[x < 0].sum()),
                    'Valor_Limpo': lambda x: abs(x[x < 0].sum())
                })
                .reset_index()
                .sort_values(by='Valor_Limpo', ascending=False)
                .head(10)
            )

            df_top10_centros.insert(0, 'Posição', [f"{i+1}º" for i in range(len(df_top10_centros))])
            df_top10_centros = df_top10_centros[['Posição', 'Loja_Nome', 'Regional_Nome', 'Qtd_Limpa', 'Valor_Limpo']]
            df_top10_centros.rename(columns={'Loja_Nome': 'Centro', 'Regional_Nome': 'Divisão Regional', 'Qtd_Limpa': 'Perda (Qtd)', 'Valor_Limpo': 'Perda (R$)'}, inplace=True)

            df_top10_centros['Perda (Qtd)'] = df_top10_centros['Perda (Qtd)'].apply(lambda x: f"-{x:,.0f} un")
            df_top10_centros['Perda (R$)'] = df_top10_centros['Perda (R$)'].apply(lambda x: f"R$ -{x:,.2f}")

            st.dataframe(df_top10_centros, use_container_width=True, hide_index=True)

        # --- MARCAS (VISÍVEL PARA TODOS OS PERFIS) ---
        st.markdown("<br>", unsafe_allow_html=True)
        st.subheader("⚠️ Ranking: Top 10 Marcas com Maior Perda")

        df_top10_marcas = (
            df_filtered[(df_filtered['Valor_Limpo'] < 0) | (df_filtered['Qtd_Limpa'] < 0)]
            .groupby('Marca_Nome')
            .agg({
                'Qtd_Limpa': lambda x: abs(x[x < 0].sum()),
                'Valor_Limpo': lambda x: abs(x[x < 0].sum())
            })
            .reset_index()
            .sort_values(by='Valor_Limpo', ascending=False)
            .head(10)
        )

        df_top10_marcas.insert(0, 'Posição', [f"{i+1}º" for i in range(len(df_top10_marcas))])
        df_top10_marcas = df_top10_marcas[['Posição', 'Marca_Nome', 'Qtd_Limpa', 'Valor_Limpo']]
        df_top10_marcas.rename(columns={'Marca_Nome': 'Marca', 'Qtd_Limpa': 'Perda (Qtd)', 'Valor_Limpo': 'Perda (R$)'}, inplace=True)

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

        st.markdown("---")
        st.subheader("📥 Exportação de Dados Filtrados")
        
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_filtered.to_excel(writer, index=False, sheet_name="Inventario_Filtrado")
        processed_data = output.getvalue()

        st.download_button(
            label="📄 Baixar Relatório em Excel (.xlsx)",
            data=processed_data,
            file_name="relatorio_inventario_filtrado.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"Erro ao carregar os dados do arquivo Excel na nuvem: {e}")


# --- FLUXO PRINCIPAL DA APLICAÇÃO ---
if not st.session_state["logado"]:
    renderizar_tela_login()

elif st.session_state["troca_obrigatoria"]:
    renderizar_tela_troca_obrigatoria()

else:
    usuarios_db, _ = carregar_dados_db()
    usr_atual = st.session_state["usuario_atual"]
    dados_logado = usuarios_db.get(usr_atual, {})

    st.sidebar.markdown(f"👤 **Usuário:** `{usr_atual}`\n\n💼 **Cargo:** `{dados_logado.get('perfil', 'Gerente')}`")
    
    with st.sidebar.expander("🔑 Alterar minha senha"):
        with st.form("form_mudar_senha_sidebar"):
            senha_antiga_sb = st.text_input("Senha Atual:", type="password")
            nova_senha_sb = st.text_input("Nova Senha:", type="password")
            confirma_sb = st.text_input("Confirme a Nova Senha:", type="password")
            btn_mudar_sb = st.form_submit_button("Atualizar Senha")

            if btn_mudar_sb:
                usuarios_dict, removidos_set = carregar_dados_db()
                
                if gerar_hash(senha_antiga_sb) != usuarios_dict[usr_atual]["senha"]:
                    st.error("Senha atual incorreta.")
                elif len(nova_senha_sb) < 6:
                    st.error("Mínimo de 6 caracteres.")
                elif nova_senha_sb != confirma_sb:
                    st.error("Senhas não conferem.")
                else:
                    usuarios_dict[usr_atual]["senha"] = gerar_hash(nova_senha_sb)
                    usuarios_dict[usr_atual]["forcar_redefinicao"] = False
                    salvar_dados_db(usuarios_dict, removidos_set)
                    st.success("✅ Senha alterada com sucesso!")

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