import streamlit as st
import pandas as pd
import plotly.express as px
import json
import os
import hashlib
import requests
import io
from PIL import Image, ImageDraw, ImageFont

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
    "controladoriaprevencao@gmail.com": (STR_REGIONAL_1, "Regional 1"),
    "josue.victor@vonnycosmeticos.com.br": ("TODAS", "Administrador"),
    "vanusia.garcia@casadolojista.com.br": ("TODAS", "Administrador")
}

OPCOES_PERFIL = ["Gerente", "Líder de Loja", "Regional 1", "Regional 2", "Administrador"]


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
        elif email_limpo in usuarios:
            if "perfil" not in usuarios[email_limpo]:
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


# --- DESENHAR TABELAS NA IMAGEM (PIL) ---
def desenhar_tabela_pil(draw, df_tabela, titulo, start_x, start_y, largura_max, font_titulo, font_corpo):
    draw.text((start_x, start_y), titulo, fill="#ffffff", font=font_titulo)
    y = start_y + 35

    colunas = list(df_tabela.columns)
    num_cols = len(colunas)
    largura_col = largura_max // num_cols

    # Cabeçalho
    draw.rectangle([start_x, y, start_x + largura_max, y + 30], fill="#1e232a", outline="#30363d", width=1)
    for i, col in enumerate(colunas):
        cx = start_x + (i * largura_col) + 10
        draw.text((cx, y + 6), str(col), fill="#4ba3e3", font=font_corpo)
    y += 30

    # Linhas de dados
    for idx, row in df_tabela.iterrows():
        cor_fundo = "#161b22" if idx % 2 == 0 else "#0e1117"
        draw.rectangle([start_x, y, start_x + largura_max, y + 26], fill=cor_fundo, outline="#21262d", width=1)
        for i, col in enumerate(colunas):
            cx = start_x + (i * largura_col) + 10
            draw.text((cx, y + 5), str(row[col]), fill="#d0d7de", font=font_corpo)
        y += 26

    return y + 30


# --- FUNÇÃO COMPLETA PARA COMPILAR VISÃO GERAL EM IMAGEM ---
def gerar_imagem_dashboard(figuras_lista, kpis_dict, df_top_centros=None, df_top_marcas=None):
    imagens_bytes = []
    for fig in figuras_lista:
        fig_temp = fig.full_figure_for_development(warn=False)
        fig_temp.update_layout(width=1190, height=450, margin=dict(l=40, r=40, t=50, b=50))
        img_bytes = fig_temp.to_image(format="png", width=1190, height=450, scale=2)
        imagens_bytes.append(Image.open(io.BytesIO(img_bytes)))

    largura = 1250
    altura_cabecalho = 180
    altura_por_grafico = 460
    
    num_linhas_centros = len(df_top_centros) if df_top_centros is not None else 0
    altura_tabela_centros = (60 + (num_linhas_centros * 28) + 40) if num_linhas_centros > 0 else 0

    num_linhas_marcas = len(df_top_marcas) if df_top_marcas is not None else 0
    altura_tabela_marcas = (60 + (num_linhas_marcas * 28) + 40) if num_linhas_marcas > 0 else 0

    altura_total = altura_cabecalho + (len(imagens_bytes) * altura_por_grafico) + altura_tabela_centros + altura_tabela_marcas + 80

    imagem_final = Image.new("RGB", (largura, altura_total), color="#0e1117")
    draw = ImageDraw.Draw(imagem_final)

    try:
        font_titulo_gen = ImageFont.truetype("arial.ttf", 26)
        font_secao = ImageFont.truetype("arial.ttf", 20)
        font_kpi_rotulo = ImageFont.truetype("arial.ttf", 13)
        font_kpi_valor = ImageFont.truetype("arial.ttf", 18)
        font_tabela = ImageFont.truetype("arial.ttf", 12)
    except IOError:
        font_titulo_gen = ImageFont.load_default()
        font_secao = ImageFont.load_default()
        font_kpi_rotulo = ImageFont.load_default()
        font_kpi_valor = ImageFont.load_default()
        font_tabela = ImageFont.load_default()

    # Cabeçalho
    draw.text((30, 25), "📊 Dashboard Executivo de Inventário - Visão Geral", fill="#ffffff", font=font_titulo_gen)

    # KPIs
    col_x = 30
    largura_card = 270
    for rotulo, valor in kpis_dict.items():
        draw.rectangle([col_x, 80, col_x + largura_card, 150], fill="#1e232a", outline="#30363d", width=1)
        draw.text((col_x + 15, 90), rotulo, fill="#a3a8b2", font=font_kpi_rotulo)
        draw.text((col_x + 15, 115), str(valor), fill="#4ba3e3", font=font_kpi_valor)
        col_x += largura_card + 20

    y_offset = altura_cabecalho

    # Renderizar Gráficos
    for img in imagens_bytes:
        imagem_final.paste(img, (30, y_offset))
        y_offset += altura_por_grafico

    # Renderizar Tabela Top 10 Centros
    if df_top_centros is not None and not df_top_centros.empty:
        y_offset = desenhar_tabela_pil(
            draw, df_top_centros, "🏢 Ranking: Top 10 Centros com Maior Perda", 
            start_x=30, start_y=y_offset, largura_max=1190, 
            font_titulo=font_secao, font_corpo=font_tabela
        )

    # Renderizar Tabela Top 10 Marcas
    if df_top_marcas is not None and not df_top_marcas.empty:
        y_offset = desenhar_tabela_pil(
            draw, df_top_marcas, "⚠️ Ranking: Top 10 Marcas com Maior Perda", 
            start_x=30, start_y=y_offset, largura_max=1190, 
            font_titulo=font_secao, font_corpo=font_tabela
        )

    buf = io.BytesIO()
    imagem_final.save(buf, format="PNG")
    return buf.getvalue()

    # ==============================================================================
# INTERFACE DE USUÁRIO - TELAS
# ==============================================================================

def renderizar_troca_senha_obrigatoria(usuarios_db, email_usuario, removidos_set):
    st.title("🔒 Alteração de Senha Obrigatória")
    st.warning("É necessário redefinir sua senha no primeiro acesso para garantir a segurança da conta.")

    with st.form("form_troca_obrigatoria"):
        nova_senha = st.text_input("Nova Senha", type="password")
        confirma_senha = st.text_input("Confirme a Nova Senha", type="password")
        btn_salvar = st.form_submit_button("Salvar e Acessar Dashboard")

        if btn_salvar:
            if not nova_senha:
                st.error("A senha não pode estar em branco.")
            elif nova_senha != confirma_senha:
                st.error("As senhas digitadas não coincidem!")
            else:
                usuarios_db[email_usuario]["senha"] = gerar_hash(nova_senha)
                usuarios_db[email_usuario]["forcar_redefinicao"] = False
                salvar_dados_db(usuarios_db, removidos_set)
                st.session_state["forcar_redefinicao"] = False
                st.success("Senha cadastrada com sucesso!")
                st.rerun()


def renderizar_gestao_usuarios(usuarios_db, email_admin, removidos_set):
    st.header("👥 Painel de Gestão de Usuários")
    st.info("Cadastre novos usuários, redefina senhas ou remova permissões de acesso.")

    # 1. Formulário para Adicionar / Editar Usuário
    st.subheader("➕ Adicionar / Modificar Usuário")
    with st.form("form_novo_usuario"):
        col1, col2, col3 = st.columns([3, 2, 2])
        with col1:
            novo_email = st.text_input("E-mail corporativo").strip().lower()
        with col2:
            novo_perfil = st.selectbox("Perfil de Acesso", OPCOES_PERFIL)
        with col3:
            nova_loja = st.text_input("Lojas/Centros (ex: B001 ou B001,B002)").strip()

        btn_cadastrar = st.form_submit_button("Salvar Usuário")

        if btn_cadastrar:
            if not novo_email:
                st.error("Por favor, digite o e-mail.")
            else:
                loja_final = "TODAS" if novo_perfil == "Administrador" else (nova_loja if nova_loja else "TODAS")
                
                if novo_email in removidos_set:
                    removidos_set.remove(novo_email)

                usuarios_db[novo_email] = {
                    "loja": loja_final,
                    "perfil": novo_perfil,
                    "senha": None,
                    "forcar_redefinicao": True
                }
                salvar_dados_db(usuarios_db, removidos_set)
                st.success(f"Usuário {novo_email} salvo com sucesso!")
                st.rerun()

    st.markdown("---")

    # 2. Tabela de Usuários Existentes
    st.subheader("📋 Usuários Cadastrados")
    
    for email, dados in list(usuarios_db.items()):
        with st.expander(f"✉️ {email} - [{dados.get('perfil', 'Gerente')}] - Lojas: {dados.get('loja', 'N/A')}"):
            col_a, col_b = st.columns(2)
            with col_a:
                st.write(f"**Perfil:** {dados.get('perfil', 'Gerente')}")
                st.write(f"**Lojas Vinculadas:** {dados.get('loja', 'N/A')}")
                st.write(f"**Status da Senha:** {'Cadastrada' if dados.get('senha') else 'Pendente (Primeiro Acesso)'}")
            with col_b:
                if st.button(f"🔄 Redefinir Senha", key=f"reset_{email}"):
                    dados["senha"] = None
                    dados["forcar_redefinicao"] = True
                    salvar_dados_db(usuarios_db, removidos_set)
                    st.success(f"Senha de {email} redefinida. O usuário deverá criar uma nova senha no próximo acesso.")
                    st.rerun()

                if email != email_admin:
                    if st.button(f"🗑️ Remover Usuário", key=f"del_{email}"):
                        del usuarios_db[email]
                        removidos_set.add(email)
                        salvar_dados_db(usuarios_db, removidos_set)
                        st.warning(f"Usuário {email} removido.")
                        st.rerun()


def renderizar_dashboard(df_raw, perfil_usuario, lojas_permitidas_str):
    st.sidebar.markdown("---")
    st.sidebar.title("🔍 Filtros Dinâmicos")

    # Tratamento das lojas permitidas pelo perfil
    lojas_permitidas = [l.strip().upper() for l in lojas_permitidas_str.split(",")] if lojas_permitidas_str != "TODAS" else []

    # Filtrar dataframe por acesso
    if perfil_usuario in ["Administrador", "Controle"]:
        df_filtrado = df_raw.copy()
    else:
        df_filtrado = df_raw[df_raw['Loja_Nome'].str.upper().isin(lojas_permitidas)].copy()

    # Filtro de Regional na Sidebar
    regionais_disponiveis = sorted(df_filtrado['Regional_Nome'].dropna().unique().tolist())
    sel_regional = st.sidebar.multiselect("Regional", options=regionais_disponiveis, default=regionais_disponiveis)
    if sel_regional:
        df_filtrado = df_filtrado[df_filtrado['Regional_Nome'].isin(sel_regional)]

    # Filtro de Centro/Loja na Sidebar
    centros_disponiveis = sorted(df_filtrado['Loja_Nome'].dropna().unique().tolist())
    sel_centro = st.sidebar.multiselect("Centro / Loja", options=centros_disponiveis, default=centros_disponiveis)
    if sel_centro:
        df_filtrado = df_filtrado[df_filtrado['Loja_Nome'].isin(sel_centro)]

    # Filtro de Marca na Sidebar
    marcas_disponiveis = sorted(df_filtrado['Marca_Nome'].dropna().unique().tolist())
    sel_marca = st.sidebar.multiselect("Marca / Fornecedor", options=marcas_disponiveis)
    if sel_marca:
        df_filtrado = df_filtrado[df_filtrado['Marca_Nome'].isin(sel_marca)]

    st.title("📊 Dashboard Executivo de Inventário")
    st.markdown("---")

    # Cálculo dos KPIs Globais
    total_perdas_qtd = df_filtrado[df_filtrado['Qtd_Limpa'] < 0]['Qtd_Limpa'].sum()
    total_perdas_rs = df_filtrado[df_filtrado['Valor_Limpo'] < 0]['Valor_Limpo'].sum()
    total_sobras_rs = df_filtrado[df_filtrado['Valor_Limpo'] > 0]['Valor_Limpo'].sum()
    resultado_net_rs = df_filtrado['Valor_Limpo'].sum()

    kpis_dict = {
        "Total de Perdas (Qtd)": formatar_qtd(total_perdas_qtd),
        "Perda Total (R$)": formatar_moeda(total_perdas_rs),
        "Sobras / Ajustes (R$)": formatar_moeda(total_sobras_rs),
        "Resultado Net (R$)": formatar_moeda(resultado_net_rs)
    }

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Total de Perdas", formatar_qtd(total_perdas_qtd))
    c2.metric("Perda Total (R$)", formatar_moeda(total_perdas_rs))
    c3.metric("Sobras / Ajustes", formatar_moeda(total_sobras_rs))
    c4.metric("Resultado Net", formatar_moeda(resultado_net_rs))

    st.markdown("---")

    # --------------------------------------------------------------------------
    # CONSTRUÇÃO DOS GRÁFICOS
    # --------------------------------------------------------------------------
    figuras_exportacao = []

    # 1. Comparativo por Divisão Regional
    df_reg = df_filtrado.groupby('Regional_Nome')['Valor_Limpo'].sum().reset_index()
    fig_reg = px.bar(
        df_reg, x='Regional_Nome', y='Valor_Limpo',
        title="Comparativo por Divisão Regional",
        labels={'Regional_Nome': 'Regional', 'Valor_Limpo': 'Valor Net (R$)'},
        color_discrete_sequence=['#ff7f0e'], text_auto=True
    )
    fig_reg.update_layout(template="plotly_dark", height=400)
    st.plotly_chart(fig_reg, use_container_width=True)
    figuras_exportacao.append(fig_reg)

    # 2. Perda por Centro (Qtd)
    df_centro_qtd = df_filtrado[df_filtrado['Qtd_Limpa'] < 0].groupby('Loja_Nome')['Qtd_Limpa'].sum().reset_index().sort_values('Qtd_Limpa')
    fig_centro_qtd = px.bar(
        df_centro_qtd, x='Loja_Nome', y='Qtd_Limpa',
        title="Perda por Centro (Qtd)",
        labels={'Loja_Nome': 'Centro', 'Qtd_Limpa': 'Perda (Qtd)'},
        color_discrete_sequence=['#1f77b4'], text_auto=True
    )
    fig_centro_qtd.update_layout(template="plotly_dark", height=400)
    st.plotly_chart(fig_centro_qtd, use_container_width=True)
    figuras_exportacao.append(fig_centro_qtd)

    # 3. Perda por Centro (R$)
    df_centro_rs = df_filtrado[df_filtrado['Valor_Limpo'] < 0].groupby('Loja_Nome')['Valor_Limpo'].sum().reset_index().sort_values('Valor_Limpo')
    fig_centro_rs = px.bar(
        df_centro_rs, x='Loja_Nome', y='Valor_Limpo',
        title="Perda por Centro (R$)",
        labels={'Loja_Nome': 'Centro', 'Valor_Limpo': 'Perda (R$)'},
        color_discrete_sequence=['#1f77b4'], text_auto=True
    )
    fig_centro_rs.update_layout(template="plotly_dark", height=400)
    st.plotly_chart(fig_centro_rs, use_container_width=True)
    figuras_exportacao.append(fig_centro_rs)

    # 4. Perdas por Marca - Qtd
    df_marca_qtd = df_filtrado[df_filtrado['Qtd_Limpa'] < 0].groupby('Marca_Nome')['Qtd_Limpa'].sum().reset_index().sort_values('Qtd_Limpa').head(15)
    fig_marca_qtd = px.line(
        df_marca_qtd, x='Marca_Nome', y='Qtd_Limpa', markers=True,
        title="Perdas por Marca - Top 15 (Qtd)",
        labels={'Marca_Nome': 'Marca', 'Qtd_Limpa': 'Perda (Qtd)'},
        color_discrete_sequence=['#ff7f0e'], text_auto=True
    )
    fig_marca_qtd.update_layout(template="plotly_dark", height=400)
    st.plotly_chart(fig_marca_qtd, use_container_width=True)
    figuras_exportacao.append(fig_marca_qtd)

    # 5. Perdas por Marca - R$
    df_marca_rs = df_filtrado[df_filtrado['Valor_Limpo'] < 0].groupby('Marca_Nome')['Valor_Limpo'].sum().reset_index().sort_values('Valor_Limpo').head(15)
    fig_marca_rs = px.line(
        df_marca_rs, x='Marca_Nome', y='Valor_Limpo', markers=True,
        title="Perdas por Marca - Top 15 (R$)",
        labels={'Marca_Nome': 'Marca', 'Valor_Limpo': 'Perda (R$)'},
        color_discrete_sequence=['#1f77b4'], text_auto=True
    )
    fig_marca_rs.update_layout(template="plotly_dark", height=400)
    st.plotly_chart(fig_marca_rs, use_container_width=True)
    figuras_exportacao.append(fig_marca_rs)

    st.markdown("---")

    # --------------------------------------------------------------------------
    # TABELAS RANKING TOP 10
    # --------------------------------------------------------------------------
    col_t1, col_t2 = st.columns(2)

    with col_t1:
        st.subheader("🏢 Top 10 Centros com Maior Perda")
        df_top_c = df_filtrado[df_filtrado['Valor_Limpo'] < 0].groupby(['Loja_Nome', 'Regional_Nome']).agg(
            Perda_Qtd=('Qtd_Limpa', 'sum'),
            Perda_RS=('Valor_Limpo', 'sum')
        ).reset_index().sort_values('Perda_RS').head(10)

        df_top_c_exibicao = df_top_c.copy()
        df_top_c_exibicao['Perda_Qtd'] = df_top_c_exibicao['Perda_Qtd'].apply(formatar_qtd)
        df_top_c_exibicao['Perda_RS'] = df_top_c_exibicao['Perda_RS'].apply(formatar_moeda)
        df_top_c_exibicao.columns = ['Centro', 'Divisão Regional', 'Perda (Qtd)', 'Perda (R$)']
        st.dataframe(df_top_c_exibicao, use_container_width=True, hide_index=True)

    with col_t2:
        st.subheader("⚠️ Top 10 Marcas com Maior Perda")
        df_top_m = df_filtrado[df_filtrado['Valor_Limpo'] < 0].groupby('Marca_Nome').agg(
            Perda_Qtd=('Qtd_Limpa', 'sum'),
            Perda_RS=('Valor_Limpo', 'sum')
        ).reset_index().sort_values('Perda_RS').head(10)

        df_top_m_exibicao = df_top_m.copy()
        df_top_m_exibicao['Perda_Qtd'] = df_top_m_exibicao['Perda_Qtd'].apply(formatar_qtd)
        df_top_m_exibicao['Perda_RS'] = df_top_m_exibicao['Perda_RS'].apply(formatar_moeda)
        df_top_m_exibicao.columns = ['Marca', 'Perda (Qtd)', 'Perda (R$)']
        st.dataframe(df_top_m_exibicao, use_container_width=True, hide_index=True)

    st.markdown("---")

    # Botão de Exportação de Imagem do Dashboard Completo
    st.subheader("📸 Exportar Visão Geral")
    if st.button("🖼️ Gerar Imagem da Visão Geral (.png)"):
        with st.spinner("Gerando imagem em alta resolução sem cortes..."):
            img_bytes = gerar_imagem_dashboard(
                figuras_lista=figuras_exportacao,
                kpis_dict=kpis_dict,
                df_top_centros=df_top_c_exibicao,
                df_top_marcas=df_top_m_exibicao
            )
            st.download_button(
                label="⬇️ Baixar Imagem do Dashboard (.png)",
                data=img_bytes,
                file_name="visao_geral_dashboard_completa.png",
                mime="image/png"
            )


# ==============================================================================
# FLUXO PRINCIPAL DA APLICAÇÃO (MAIN)
# ==============================================================================
def main():
    usuarios_db, removidos_set = carregar_dados_db()

    if "usuario_logado" not in st.session_state:
        st.session_state["usuario_logado"] = None

    # TELA DE LOGIN
    if st.session_state["usuario_logado"] is None:
        st.title("🔐 Acesso ao Dashboard de Inventário")
        st.markdown("---")
        
        with st.form("form_login"):
            email_input = st.text_input("E-mail corporativo").strip().lower()
            senha_input = st.text_input("Senha", type="password")
            btn_entrar = st.form_submit_button("Entrar")

            if btn_entrar:
                if email_input in usuarios_db:
                    user_data = usuarios_db[email_input]
                    senha_salva = user_data.get("senha")

                    # Caso 1: Primeiro acesso (sem senha definida)
                    if senha_salva is None:
                        st.session_state["usuario_logado"] = email_input
                        st.session_state["forcar_redefinicao"] = True
                        st.rerun()

                    # Caso 2: Validação da senha digitada
                    elif senha_salva == gerar_hash(senha_input):
                        st.session_state["usuario_logado"] = email_input
                        st.session_state["forcar_redefinicao"] = user_data.get("forcar_redefinicao", False)
                        st.rerun()
                    else:
                        st.error("Senha incorreta.")
                else:
                    st.error("E-mail não cadastrado ou sem permissão de acesso.")
        return

    # SESSÃO ATIVA
    usuario_atual = st.session_state["usuario_logado"]
    dados_user = usuarios_db.get(usuario_atual, {})
    perfil_act = dados_user.get("perfil", "Gerente")
    lojas_act = dados_user.get("loja", "")

    # Barra lateral de navegação do usuário
    st.sidebar.write(f"👤 **Usuário:** {usuario_atual}")
    st.sidebar.write(f"🔰 **Perfil:** {perfil_act}")
    
    if st.sidebar.button("🚪 Sair (Logout)"):
        st.session_state["usuario_logado"] = None
        st.session_state["forcar_redefinicao"] = False
        st.rerun()

    # Redirecionamento obrigatório caso exija redefinição de senha
    if st.session_state.get("forcar_redefinicao", False):
        renderizar_troca_senha_obrigatoria(usuarios_db, usuario_atual, removidos_set)
        return

    # Navegação por abas
    if perfil_act == "Administrador":
        aba1, aba2 = st.tabs(["📊 Dashboard Executivo", "👥 Gestão de Usuários"])
        with aba1:
            try:
                df = load_data()
                renderizar_dashboard(df, perfil_act, lojas_act)
            except Exception as e:
                st.error(f"Erro ao carregar dados da nuvem: {e}")
        with aba2:
            renderizar_gestao_usuarios(usuarios_db, usuario_atual, removidos_set)
    else:
        try:
            df = load_data()
            renderizar_dashboard(df, perfil_act, lojas_act)
        except Exception as e:
            st.error(f"Erro ao carregar dados da nuvem: {e}")

if __name__ == "__main__":
    main()