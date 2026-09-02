import streamlit as st
import pandas as pd
import plotly.express as px

st.set_page_config(
    page_title="Dashboard Executivo de Inventário",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown("""
    <style>
    .stApp { background-color: #0e1117; color: #ffffff; }
    div[data-testid="stMetricValue"] { font-size: 1.8rem !important; font-weight: bold; }
    </style>
""", unsafe_allow_html=True)

@st.cache_data(ttl=300)
def load_data(uploaded_file):
    # Procura a linha real do cabeçalho nas primeiras 30 linhas
    df_preview = pd.read_excel(uploaded_file, engine='openpyxl', header=None, nrows=30)
    
    header_idx = 0
    for idx, row in df_preview.iterrows():
        row_str = [str(val).lower() for val in row.dropna().tolist()]
        texto_linha = " ".join(row_str)
        # Procura termos chave do cabeçalho real
        if any(term in texto_linha for term in ['qtd', 'montante', 'fornecedor', 'centro', 'nome 1', 'material']):
            header_idx = idx
            break

    df = pd.read_excel(uploaded_file, engine='openpyxl', header=header_idx)
    df = df.dropna(how='all')
    
    # Limpa os nomes das colunas
    df.columns = [str(c).replace('\n', ' ').strip() for c in df.columns]
    return df

def formatar_moeda(val):
    return f"R$ {val:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

def formatar_qtd(val):
    return f"{val:,.0f} UN".replace(",", ".")

# --- INTERFACE ---
st.sidebar.title("📁 Carregar Planilha")
uploaded_file = st.sidebar.file_uploader("Envie seu arquivo Excel (.xlsx ou .xlsm):", type=['xlsx', 'xlsm'])

if uploaded_file is not None:
    try:
        df_raw = load_data(uploaded_file)
        
        st.sidebar.title("⚙️ Mapeamento de Colunas")
        colunas_disponiveis = list(df_raw.columns)

        # Seleção manual das colunas com sugestão automática
        col_qtd = st.sidebar.selectbox("Coluna de Quantidade:", colunas_disponiveis, index=0)
        col_valor = st.sidebar.selectbox("Coluna de Valor (R$):", colunas_disponiveis, index=min(1, len(colunas_disponiveis)-1))
        col_loja = st.sidebar.selectbox("Coluna de Centro / Loja:", colunas_disponiveis, index=min(2, len(colunas_disponiveis)-1))
        col_marca = st.sidebar.selectbox("Coluna de Marca / Fornecedor:", colunas_disponiveis, index=min(3, len(colunas_disponiveis)-1))

        df = df_raw.copy()

        # Tratamento Numérico
        df['Qtd_Limpa'] = pd.to_numeric(df[col_qtd], errors='coerce').fillna(0.0)
        df['Valor_Limpo'] = pd.to_numeric(df[col_valor], errors='coerce').fillna(0.0)

        # Tratamento de Texto
        df['Loja_Nome'] = df[col_loja].astype(str).str.strip()
        df['Marca_Nome'] = df[col_marca].astype(str).str.strip()

        # Remove linhas de totalizadores ou cabeçalhos duplicados
        descartar = ['NAN', 'NONE', 'TOTAL', 'RESULTADO', 'SUM', 'SOMA DE QUANTIDADE', 'SOMA DE VALOR', 'UNNAMED']
        df = df[~df['Loja_Nome'].str.upper().isin(descartar)]
        df = df[~df['Marca_Nome'].str.upper().isin(descartar)]
        df = df[(df['Qtd_Limpa'] != 0) | (df['Valor_Limpo'] != 0)]

        # Filtros no Menu Lateral
        st.sidebar.title("🎯 Filtros")
        lojas = sorted([x for x in df['Loja_Nome'].unique() if x])
        marcas = sorted([x for x in df['Marca_Nome'].unique() if x])

        lojas_sel = st.sidebar.multiselect("Selecione os Centros:", options=lojas, default=lojas)
        marcas_sel = st.sidebar.multiselect("Selecione as Marcas/Fornecedores:", options=marcas, default=marcas)

        df_filtered = df[
            (df['Loja_Nome'].isin(lojas_sel if lojas_sel else lojas)) & 
            (df['Marca_Nome'].isin(marcas_sel if marcas_sel else marcas))
        ]

        st.title("📊 Dashboard Executivo de Inventário")
        st.markdown("---")

        # Trabalha com os valores absolutos para garantir exibição visual correta independente de sinal (+/-)
        df_filtered['Qtd_Abs'] = df_filtered['Qtd_Limpa'].abs()
        df_filtered['Valor_Abs'] = df_filtered['Valor_Limpo'].abs()

        # KPIs
        perda_qtd = df_filtered['Qtd_Abs'].sum()
        perda_rs = df_filtered['Valor_Abs'].sum()

        kpi1, kpi2 = st.columns(2)
        kpi1.metric("Total Quantidade", formatar_qtd(perda_qtd))
        kpi2.metric("Total Valor (R$)", formatar_moeda(perda_rs))

        st.markdown("<br>", unsafe_allow_html=True)

        # RANKING POR MARCA / FORNECEDOR
        st.subheader("⚠️ Ranking: Marcas / Fornecedores")

        df_marcas_rank = (
            df_filtered.groupby('Marca_Nome')
            .agg({'Qtd_Abs': 'sum', 'Valor_Abs': 'sum'})
            .reset_index()
            .sort_values(by='Valor_Abs', ascending=False)
        )

        df_marcas_rank.insert(0, 'Posição', [f"{i+1}º" for i in range(len(df_marcas_rank))])
        df_marcas_rank['Perda (Qtd)'] = df_marcas_rank['Qtd_Abs'].apply(formatar_qtd)
        df_marcas_rank['Perda (R$)'] = df_marcas_rank['Valor_Abs'].apply(formatar_moeda)

        st.dataframe(
            df_marcas_rank[['Posição', 'Marca_Nome', 'Perda (Qtd)', 'Perda (R$)']].rename(columns={'Marca_Nome': 'Marca / Fornecedor'}),
            use_container_width=True,
            hide_index=True
        )

        # GRÁFICOS DE MARCAS
        st.markdown("<br>", unsafe_allow_html=True)
        m_col1, m_col2 = st.columns(2)

        with m_col1:
            st.subheader("📦 Perdas por Marca (Qtd)")
            fig_m_qtd = px.bar(
                df_marcas_rank.head(15), x='Marca_Nome', y='Qtd_Abs', text='Perda (Qtd)',
                labels={'Qtd_Abs': 'Quantidade', 'Marca_Nome': 'Marca'}
            )
            fig_m_qtd.update_traces(marker_color='#ff7f0e', textposition='outside')
            fig_m_qtd.update_layout(template="plotly_dark", paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)", xaxis_tickangle=-45)
            st.plotly_chart(fig_m_qtd, use_container_width=True)

        with m_col2:
            st.subheader("🏷️ Perdas por Marca (R$)")
            fig_m_rs = px.bar(
                df_marcas_rank.head(15), x='Marca_Nome', y='Valor_Abs', text='Perda (R$)',
                labels={'Valor_Abs': 'Valor (R$)', 'Marca_Nome': 'Marca'}
            )
            fig_m_rs.update_traces(marker_color='#4ba3e3', textposition='outside')
            fig_m_rs.update_layout(template="plotly_dark", paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)", xaxis_tickangle=-45)
            st.plotly_chart(fig_m_rs, use_container_width=True)

    except Exception as e:
        st.error(f"Erro ao processar: {e}")
else:
    st.info("👈 Faça o upload da sua planilha Excel no menu à esquerda.")