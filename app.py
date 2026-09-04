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

# URL do arquivo no OneDrive / SharePoint / Nuvem
URL_EXCEL_NUVEM = "https://vonnycosmeticos-my.sharepoint.com/:x:/g/personal/josue_pereira_vonnycosmeticos_onmicrosoft_com/IQAVAJHO0KlcS73eMCZZkJMEAdrs0fKrEhefibx1ieyMW_Y?e=B8QkcG&download=1"

DB_FILE = "usuarios_db.json"

# Usuários Administradores Iniciais / Padrão (Criados automaticamente se o JSON não existir)
ADMINS_INICIAIS = {
    "jvn221106@gmail.com": {"loja": "TODAS", "is_admin": True},
    "sergio.oliveira@vonnycosmeticos.com.br": {"loja": "TODAS", "is_admin": True},
    "josue.victor@vonnycosmeticos.com.br": {"loja": "TODAS", "is_admin": True},
    "controladoriaprevencao@gmail.com": {"loja": "TODAS", "is_admin": True}
}

# --- PERSISTÊNCIA DE USUÁRIOS E PERMISSÕES NO SITE ---
def carregar_usuarios():
    if os.path.exists(DB_FILE):
        with open(DB_FILE, "r") as f:
            usuarios = json.load(f)
    else:
        usuarios = {}

    # Garante que os administradores iniciais estejam presentes
    atualizou = False
    for email, dados in ADMINS_INICIAIS.items():
        if email not in usuarios:
            usuarios[email] = {
                "loja": dados["loja"],
                "is_admin": dados["is_admin"],
                "senha": None,
                "forcar_redefinicao": False
            }
            atualizou = True

    if atualizou or not os.path.exists(DB_FILE):
        salvar_usuarios(usuarios)

    return usuarios

def salvar_usuarios(usuarios):
    with open(DB_FILE, "w") as f:
        json.dump(usuarios, f, indent=4)

def gerar_hash(senha):
    return hashlib.sha256(senha.encode()).hexdigest()

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


# --- LEITURA E TRATAMENTO DA PLANILHA DE INVENTÁRIO (NUVEM) ---
@st.cache_data(ttl=300)  # Recarrega a planilha de dados da nuvem a cada 5 minutos
def load_data():
    response = requests.get(URL_EXCEL_NUVEM)
    response.raise_for_status()
    excel_file = io.BytesIO(response.content)
    df = pd.read_excel(excel_file, sheet_name="VALORES INVENTÁRIOS")
    df.columns = [str(col).strip() for col in df.columns]
    
    def achar_coluna(termos_prioritarios):
        for termo in termos_prioritarios:
            for col in df.columns:
                if termo.lower() in str(col).lower():
                    return col
        return None

    col_qtd = achar_coluna(['qtd. um registro', 'qtd', 'registro'])
    col_valor = achar_coluna(['montante em mi', 'montante', 'mi'])
    col_loja = achar_coluna(['centro'])
    col_marca = achar_coluna(['fornecedor2', 'fornecedor', 'marca'])

    if not col_qtd: col_qtd = df.columns[0]
    if not col_valor: col_valor = df.columns[1]
    if not col_loja: col_loja = df.columns[2]
    if not col_marca: col_marca = df.columns[3]

    df['Qtd_Limpa'] = pd.to_numeric(df[col_qtd], errors='coerce').fillna(0)
    df['Valor_Limpo'] = pd.to_numeric(df[col_valor], errors='coerce').fillna(0)
    
    df['Loja_Nome'] = df[col_loja].fillna('S/ Centro').astype(str).str.strip()
    df['Marca_Nome'] = df[col_marca].fillna('Sem Marca').astype(str).str.strip()
    
    return df


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

    usuarios = carregar_usuarios()

    with st.form("form_login"):
        email_input = st.text_input("E-mail corporativo:").strip().lower()
        senha_input = st.text_input("Senha:", type="password")
        btn_entrar = st.form_submit_button("Entrar", type="primary")

    if btn_entrar:
        if not email_input:
            st.error("Por favor, digite seu e-mail.")
            return

        if email_input not in usuarios:
            st.error("E-mail não autorizado. Solicite o cadastro ao Administrador.")
            return

        dados_usuario = usuarios[email_input]

        if dados_usuario["senha"] is None:
            if not senha_input or len(senha_input) < 6:
                st.warning("⚠️ **Primeiro Acesso:** Defina uma senha de no mínimo 6 caracteres e clique em entrar novamente.")
            else:
                usuarios[email_input]["senha"] = gerar_hash(senha_input)
                usuarios[email_input]["forcar_redefinicao"] = False
                salvar_usuarios(usuarios)
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
        st.info("📩 Por favor, solicite a redefinição de senha ao Administrador do sistema.")


# --- TELA OBRIGATÓRIA DE REDEFINIÇÃO DE SENHA ---
def renderizar_tela_troca_obrigatoria():
    st.title("🔑 Redefinição de Senha Obrigatória")
    st.warning("Você acessou com uma **senha temporária**. Por segurança, escolha uma nova senha definitiva para continuar.")

    usuarios = carregar_usuarios()
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
        salvar_usuarios(usuarios)

        st.session_state["troca_obrigatoria"] = False
        st.success("✅ Senha atualizada com sucesso!")
        st.rerun()


# --- ABA PAINEL ADMIN ---
def renderizar_aba_admin():
    st.header("⚙️ Painel de Gestão de Usuários e Permissões")
    usuarios = carregar_usuarios()

    # Busca lista de lojas disponíveis na planilha para preencher o select box
    try:
        df = load_data()
        lojas_lista = ["TODAS"] + [x for x in sorted(df['Loja_Nome'].unique()) if x.lower() not in ['nan', 'none', '', 's/ centro']]
    except Exception:
        lojas_lista = ["TODAS"]

    st.subheader("➕ Cadastrar ou Editar Usuário")
    with st.form("form_gerenciar_usuario"):
        col_usr, col_loja, col_perfil = st.columns([2, 1, 1])
        with col_usr:
            email_novo = st.text_input("E-mail do Usuário:").strip().lower()
        with col_loja:
            loja_selecionada = st.selectbox("Centro / Loja Autorizada:", options=lojas_lista)
        with col_perfil:
            e_admin = st.checkbox("Perfil Administrador?", value=False)
        
        btn_salvar_usr = st.form_submit_button("Salvar Usuário / Permissão", type="primary")

    if btn_salvar_usr:
        if not email_novo or "@" not in email_novo:
            st.error("Digite um e-mail válido.")
        else:
            if email_novo in usuarios:
                usuarios[email_novo]["loja"] = loja_selecionada
                usuarios[email_novo]["is_admin"] = e_admin
                st.success(f"✅ Permissões do usuário **{email_novo}** atualizadas!")
            else:
                usuarios[email_novo] = {
                    "loja": loja_selecionada,
                    "is_admin": e_admin,
                    "senha": None,
                    "forcar_redefinicao": False
                }
                st.success(f"✅ Usuário **{email_novo}** cadastrado com sucesso!")
            salvar_usuarios(usuarios)
            st.rerun()

    st.markdown("---")
    st.subheader("👥 Usuários Cadastrados")

    dados_tabela = []
    for email, dados in usuarios.items():
        dados_tabela.append({
            "E-mail": email,
            "Centro / Loja": dados.get("loja", "N/A"),
            "Perfil": "Administrador" if dados.get("is_admin", False) else "Gerente",
            "Primeiro Acesso": "✅ Concluído" if dados.get("senha") else "⏳ Pendente",
            "Senha Temporária": "⚠️ Pendente Troca" if dados.get("forcar_redefinicao") else "OK"
        })
    st.dataframe(dados_tabela, use_container_width=True)

    st.markdown("---")
    col_reset, col_del = st.columns(2)

    with col_reset:
        st.subheader("🔑 Resetar Senha / Definir Senha Temporária")
        usr_reset = st.selectbox("Selecione o e-mail:", options=list(usuarios.keys()), key="select_reset")
        senha_temp = st.text_input("Nova Senha Temporária:", type="password")
        if st.button("Aplicar Senha Temporária", type="primary"):
            if len(senha_temp) < 6:
                st.error("A senha deve ter pelo menos 6 caracteres.")
            else:
                usuarios[usr_reset]["senha"] = gerar_hash(senha_temp)
                usuarios[usr_reset]["forcar_redefinicao"] = True
                salvar_usuarios(usuarios)
                st.success(f"✅ Senha temporária definida para **{usr_reset}**!")

    with col_del:
        st.subheader("🗑️ Remover Acesso de Usuário")
        usr_del = st.selectbox("Selecione o e-mail para remover:", options=list(usuarios.keys()), key="select_del")
        if st.button("Remover Usuário", type="secondary"):
            if usr_del == st.session_state["usuario_atual"]:
                st.error("Você não pode remover seu próprio acesso enquanto está logado.")
            else:
                del usuarios[usr_del]
                salvar_usuarios(usuarios)
                st.success(f"🗑️ Usuário **{usr_del}** removido do sistema!")
                st.rerun()


# --- DASHBOARD VISUAL DE INVENTÁRIO ---
def renderizar_dashboard():
    usuarios = carregar_usuarios()
    email_logado = st.session_state["usuario_atual"]
    
    dados_usr = usuarios.get(email_logado, {})
    loja_usuario = dados_usr.get("loja", "N/A")
    e_admin = dados_usr.get("is_admin", False)

    try:
        df = load_data()

        st.sidebar.title("Filtros")

        lojas_disponiveis = [x for x in sorted(df['Loja_Nome'].unique()) if x.lower() not in ['nan', 'none', '', 's/ centro']]
        
        if e_admin or loja_usuario == "TODAS":
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
            (df['Loja_Nome'].isin(lojas_sel)) & 
            (df['Marca_Nome'].isin(marcas_sel))
        ]

        st.title("📊 Dashboard Executivo de Inventário")
        st.markdown(f"**Usuário:** `{email_logado}` | **Centro Autorizado:** `{loja_usuario}`")
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
        if e_admin:
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
                .groupby('Loja_Nome')
                .agg({
                    'Qtd_Limpa': lambda x: abs(x[x < 0].sum()),
                    'Valor_Limpo': lambda x: abs(x[x < 0].sum())
                })
                .reset_index()
                .sort_values(by='Valor_Limpo', ascending=False)
                .head(10)
            )

            df_top10_centros.insert(0, 'Posição', [f"{i+1}º" for i in range(len(df_top10_centros))])
            df_top10_centros = df_top10_centros[['Posição', 'Loja_Nome', 'Qtd_Limpa', 'Valor_Limpo']]
            df_top10_centros.rename(columns={'Loja_Nome': 'Centro', 'Qtd_Limpa': 'Perda (Qtd)', 'Valor_Limpo': 'Perda (R$)'}, inplace=True)

            df_top10_centros['Perda (Qtd)'] = df_top10_centros['Perda (Qtd)'].apply(lambda x: f"-{x:,.0f} un")
            df_top10_centros['Perda (R$)'] = df_top10_centros['Perda (R$)'].apply(lambda x: f"R$ -{x:,.2f}")

            st.dataframe(df_top10_centros, use_container_width=True, hide_index=True)

        # --- MARCAS (VISÍVEL PARA TODOS) ---
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
                .sort_values(by='Marca_Nome', ascending=False)
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

    except Exception as e:
        st.error(f"Erro ao carregar os dados do arquivo Excel na nuvem: {e}")


# --- FLUXO PRINCIPAL DA APLICAÇÃO ---
if not st.session_state["logado"]:
    renderizar_tela_login()

elif st.session_state["troca_obrigatoria"]:
    renderizar_tela_troca_obrigatoria()

else:
    # Barra Lateral
    st.sidebar.markdown(f"👤 **Usuário Conectado:**\n`{st.session_state['usuario_atual']}`")
    
    with st.sidebar.expander("🔑 Alterar minha senha"):
        with st.form("form_mudar_senha_sidebar"):
            senha_antiga_sb = st.text_input("Senha Atual:", type="password")
            nova_senha_sb = st.text_input("Nova Senha:", type="password")
            confirma_sb = st.text_input("Confirme a Nova Senha:", type="password")
            btn_mudar_sb = st.form_submit_button("Atualizar Senha")

            if btn_mudar_sb:
                usuarios_dict = carregar_usuarios()
                usr_atual = st.session_state["usuario_atual"]
                
                if gerar_hash(senha_antiga_sb) != usuarios_dict[usr_atual]["senha"]:
                    st.error("Senha atual incorreta.")
                elif len(nova_senha_sb) < 6:
                    st.error("Mínimo de 6 caracteres.")
                elif nova_senha_sb != confirma_sb:
                    st.error("Senhas não conferem.")
                else:
                    usuarios_dict[usr_atual]["senha"] = gerar_hash(nova_senha_sb)
                    usuarios_dict[usr_atual]["forcar_redefinicao"] = False
                    salvar_usuarios(usuarios_dict)
                    st.success("✅ Senha alterada com sucesso!")

    if st.sidebar.button("🚪 Sair / Logout"):
        st.session_state["logado"] = False
        st.session_state["usuario_atual"] = None
        st.session_state["troca_obrigatoria"] = False
        st.rerun()

    usuario_logado = st.session_state["usuario_atual"]
    todos_usuarios = carregar_usuarios()
    e_admin_logado = todos_usuarios.get(usuario_logado, {}).get("is_admin", False)

    # Controle de visualização de abas por perfil
    if e_admin_logado:
        aba_dash, aba_admin = st.tabs(["📊 Dashboard Geral", "⚙️ Painel Admin"])
        with aba_dash:
            renderizar_dashboard()
        with aba_admin:
            renderizar_aba_admin()
    else:
        renderizar_dashboard()