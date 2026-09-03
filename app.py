import streamlit as st
import pandas as pd
import plotly.express as px

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

email_usuario = st.session_state["usuario_logado"]
loja_permitida = PERMISSOES_EMAIL[email_usuario]

# ==============================================================================
# 3. TRATAMENTO INTELIGENTE DA PLANILHA EXCEL
# ==============================================================================
NOME_ARQUIVO_EXCEL = "STATUS DOS INVENTÁRIOS.xlsm"

@st.cache_data(ttl=300)
def carregar_dados():
    try:
        # Lê a planilha sem cabeçalho fixo para tratar títulos mesclados
        df_raw = pd.read_excel(NOME_ARQUIVO_EXCEL, header=None)
        
        # Encontra a linha real onde estão os nomes das colunas (LOJAS, STATUS, etc.)
        linha_cabecalho = 0
        for i, row in df_raw.iterrows():
            linha_texto = " ".join(row.dropna().astype(str)).upper()
            if "LOJAS" in linha_texto or "LOJA" in linha_texto or "STATUS" in linha_texto:
                linha_cabecalho = i
                break
                
        # Recarrega a planilha usando a linha correta como cabeçalho
        df = pd.read_excel(NOME_ARQUIVO_EXCEL, header=linha_cabecalho)
        df.columns = [str(col).strip() for col in df.columns]
        
        # Localiza a coluna de Loja/Centro
        col_loja = None
        for col in df.columns:
            if col.upper() in ["LOJAS", "LOJA", "CENTRO", "UNIDADE"]:
                col_loja = col
                break
        if not col_loja:
            col_loja = df.columns[0]
            
        # Extrai o código do centro (ex: "B001" de "B001 - ARICANDUVA")
        df["_CODIGO_CENTRO_"] = df[col_loja].astype(str).str.extract(r'(B\d{3})', expand=False).fillna(df[col_loja].astype(str))
        
        # Tratamento numérico para Quantidade e Valor se existirem
        for c in df.columns:
            if "QUANTIDADE" in c.upper() or "QTD" in c.upper():
                df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0)
            elif "VALOR" in c.upper() or "PREÇO" in c.upper():
                df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0)

        return df, col_loja
        
    except FileNotFoundError:
        st.error(f"❌ O arquivo `{NOME_ARQUIVO_EXCEL}` não foi localizado no repositório.")
        st.stop()
    except Exception as e:
        st.error(f"❌ Erro ao processar a planilha: {e}")
        st.stop()

df_bruto, col_loja_nome = carregar_dados()

# ==============================================================================
# 4. BARRA LATERAL E REGRAS DE ACESSO
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
    lista_centros = ["Todos os Centros"] + sorted(df_bruto["_CODIGO_CENTRO_"].unique().tolist())
    centro_selecionado = st.sidebar.selectbox("Filtrar por unidade:", lista_centros)
    
    if centro_selecionado != "Todos os Centros":
        df_exibicao = df_bruto[df_bruto["_CODIGO_CENTRO_"] == centro_selecionado]
    else:
        df_exibicao = df_bruto.copy()
else:
    df_exibicao = df_bruto[df_bruto["_CODIGO_CENTRO_"] == loja_permitida]

# ==============================================================================
# 5. DASHBOARD COMPLETO (MÉTRICAS, GRÁFICOS E TOP 10)
# ==============================================================================
st.title("📊 Painel Geral de Inventário")

if loja_permitida == "TODAS":
    st.info("Acesso global ativado (Perfil Administrador).")
else:
    st.success(f"Acesso liberado exclusivamente para a unidade **{loja_permitida}**.")

# Identifica colunas de interesse
col_marcas = next((c for c in df_exibicao.columns if "MARCA" in c.upper() or "FORNECEDOR" in c.upper()), None)
col_qtd = next((c for c in df_exibicao.columns if "QUANTIDADE" in c.upper() or "QTD" in c.upper()), None)
col_valor = next((c for c in df_exibicao.columns if "VALOR" in c.upper() or "PREÇO" in c.upper()), None)
col_status = next((c for c in df_exibicao.columns if "STATUS" in c.upper()), None)

# --- CARDS DE MÉTRICAS ---
m1, m2, m3, m4 = st.columns(4)
with m1:
    st.metric("Total de Lojas/Registros", len(df_exibicao))
with m2:
    val_qtd = df_exibicao[col_qtd].sum() if col_qtd else 0
    st.metric("Total de Itens", f"{val_qtd:,.0f}")
with m3:
    val_tot = df_exibicao[col_valor].sum() if col_valor else 0
    st.metric("Valor Total (R$)", f"R$ {val_tot:,.2f}")
with m4:
    total_marcas = df_exibicao[col_marcas].nunique() if col_marcas else "N/A"
    st.metric("Total de Marcas/Fornecedores", total_marcas)

st.divider()

# --- SEÇÃO DE GRÁFICOS E TOP 10 ---
g1, g2 = st.columns(2)

with g1:
    st.subheader("🏆 Top 10 Lojas")
    metrica_grafico = col_valor if col_valor else (col_qtd if col_qtd else None)
    
    if metrica_grafico:
        top_lojas = df_exibicao.groupby(col_loja_nome)[metrica_grafico].sum().reset_index()
        top_lojas = top_lojas.sort_values(by=metrica_grafico, ascending=False).head(10)
        fig_lojas = px.bar(top_lojas, x=metrica_grafico, y=col_loja_nome, orientation='h', 
                           title=f"Top 10 Lojas por {metrica_grafico}", text_auto=True)
        fig_lojas.update_layout(yaxis={'categoryorder':'total ascending'})
        st.plotly_chart(fig_lojas, use_container_width=True)
    else:
        st.write("Exibindo contagem por Status das Lojas:")
        if col_status:
            st.bar_chart(df_exibicao[col_status].value_counts().head(10))

with g2:
    st.subheader("🏷️ Top 10 Marcas / Fornecedores")
    if col_marcas and metrica_grafico:
        top_marcas = df_exibicao.groupby(col_marcas)[metrica_grafico].sum().reset_index()
        top_marcas = top_marcas.sort_values(by=metrica_grafico, ascending=False).head(10)
        fig_marcas = px.bar(top_marcas, x=metrica_grafico, y=col_marcas, orientation='h', 
                            title=f"Top 10 Marcas por {metrica_grafico}", text_auto=True, color_discrete_sequence=['#ff7f0e'])
        fig_marcas.update_layout(yaxis={'categoryorder':'total ascending'})
        st.plotly_chart(fig_marcas, use_container_width=True)
    elif col_marcas:
        top_m = df_exibicao[col_marcas].value_counts().head(10)
        st.bar_chart(top_m)
    else:
        st.info("Coluna de Marcas/Fornecedores não encontrada na planilha.")

st.divider()

# --- TABELA FINAL DE DADOS TRATADA ---
st.subheader("📋 Tabela Completa de Dados")
df_tabela = df_exibicao.drop(columns=["_CODIGO_CENTRO_"], errors="ignore")
st.dataframe(df_tabela, use_container_width=True, hide_index=True)
