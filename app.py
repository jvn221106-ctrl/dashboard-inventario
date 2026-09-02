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

@st.cache_data(ttl=300)
def load_data(uploaded_file):
    # Lê os dados brutos sem assumir linha de cabeçalho
    df_raw = pd.read_excel(uploaded_file, engine='openpyxl', header=None)
    
    # Encontra a primeira linha que não seja totalmente vazia para usar como cabeçalho
    header_idx = 0
    for idx, row in df_raw.iterrows():
        if row.dropna().count() >= 3:
            header_idx = idx
            break

    # Recarrega ajustando o cabeçalho correto
    df = pd.read_excel(uploaded_file, engine='openpyxl', header=header_idx)
    df = df.dropna(how='all')
    
    # Limpa o nome das colunas
    df.columns = [str(col).strip() for col in df.columns]
    
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

# --- INTERFACE ---
st.sidebar.title("📁 Carregar Planilha")
uploaded_file = st.sidebar.file_uploader("Envie seu arquivo Excel (.xlsx ou .xlsm):", type=['xlsx', 'xlsm'])

if uploaded_file is not None:
    try:
        df_raw = load_data(uploaded_file)
        
        st.sidebar.title("⚙️ Mapeamento de Colunas")
        colunas_disponiveis = list(df_raw.columns)
        
        def pegar_indice_padrao(padroes, default_idx=0, indices_ocupados=None):
            if indices_ocupados is None:
                indices_ocupados = []
            for i, col in enumerate(colunas_disponiveis):
                if i in indices_ocupados:
                    continue
                for p in padroes:
                    if p.lower() in str(col).lower():
                        return i
            for i in range(len(colunas_disponiveis)):
                if i not in indices_ocupados:
                    return i
            return min(default_idx, len(colunas_disponiveis) - 1)

        idx_qtd = pegar_indice_padrao(['qtd', 'quantidade', 'registro'], 0)
        idx_val = pegar_indice_padrao(['montante', 'valor', 'soma de valor', 'r$'], 1, [idx_qtd])
        idx_centro = pegar_indice_padrao(['centro', 'loja', 'unidade', 'filial'], 2, [idx_qtd, idx_val])
        idx_marca = pegar_indice_padrao(['fornecedor', 'marca', 'brand', 'inventários'], 3, [idx_qtd, idx_val, idx_centro])

        col_qtd = st.sidebar.selectbox("Coluna de Quantidade:", colunas_disponiveis, index=idx_qtd)
        col_valor = st.sidebar.selectbox("Coluna de Valor (R$):", colunas_disponiveis, index=idx_val)
        col_loja = st.sidebar.selectbox("Coluna de Centro / Loja:", colunas_disponiveis, index=idx_centro)
        col_marca = st.sidebar.selectbox("Coluna de Marca / Fornecedor:", colunas_disponiveis, index=idx_marca)

        # Processamento e sanitização rigorosa de dados
        df = df_raw.copy()

        # Limpeza Numérica (Força conversão para float e substitui textos por 0)
        df['Qtd_Limpa'] = pd.to_numeric(df[col_qtd], errors='coerce').fillna(0.0).astype(float)
        df['Valor_Limpo'] = pd.to_numeric(df[col_valor], errors='coerce').fillna(0.0).astype(float)

        # Limpeza de Textos
        df['Loja_Nome'] = df[col_loja].astype(str).str.strip().replace({'nan': 'Sem Centro', 'None': 'Sem Centro', '': 'Sem Centro', 'NaN': 'Sem Centro'})
        df['Marca_Nome'] = df[col_marca].astype(str).str.strip().replace({'nan': 'Sem Marca', 'None': 'Sem Marca', '': 'Sem Marca', 'NaN': 'Sem Marca'})

        # Filtra linhas de erros do Excel
        erros = ['#ERROR!', '#ERRO!', '#N/A', '#REF!', '#VALUE!']
        df = df[~df['Loja_Nome'].str.upper().isin(erros)]
        df = df[~df['Marca_Nome'].str.upper().isin(erros)]

        # Filtros na Barra Lateral
        st.sidebar.title("🎯 Filtros")

        lojas = sorted([str(x) for x in df['Loja_Nome'].unique()])
        marcas = sorted([str(x) for x in df['Marca_Nome'].unique()])

        lojas_sel = st.sidebar.multiselect("Selecione os Centros:", options=lojas, default=lojas)
        marcas_sel = st.sidebar.multiselect("Selecione as Marcas:", options=marcas, default=marcas)

        if not lojas_sel:
            lojas_sel = lojas
        if not marcas_sel:
            marcas_sel = marcas

        df_filtered = df[
            (df['Loja_Nome'].isin(lojas_sel)) & 
            (df['Marca_Nome'].isin(marcas_sel))
        ]

        st.title("📊 Dashboard Executivo de Inventário")
        st.markdown("---")

        # KPIs
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

        # GRÁFICOS (CENTROS)
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
                df_qtd_lojas, x='Loja_Nome', y='Qtd_Limpa', text='Texto_Qtd',
                labels={'Qtd_Limpa': 'Perda (Qtd)', 'Loja_Nome': 'Centro'}
            )
            fig_qtd_lojas.update_traces(marker_color='#4ba3e3', textposition='inside')
            fig_qtd_lojas.update_layout(template="plotly_dark", paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)", xaxis_title="", yaxis_title="")
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
                df_lojas, x='Loja_Nome', y='Valor_Limpo', text='Texto_Valor',
                labels={'Valor_Limpo': 'Perda (R$)', 'Loja_Nome': 'Centro'}
            )
            fig_lojas.update_traces(marker_color='#70bbfd', textposition='inside')
            fig_lojas.update_layout(template="plotly_dark", paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)", xaxis_title="", yaxis_title="")
            st.plotly_chart(fig_lojas, use_container_width=True)

        # TABELAS DE RANKING
        st.markdown("<br>", unsafe_allow_html=True)
        st.subheader("🏢 Ranking: Top 10 Centros com Maior Perda")

        df_top10_centros = (
            df_filtered[(df_filtered['Valor_Limpo'] < 0) | (df_filtered['Qtd_Limpa'] < 0)]
            .groupby('Loja_Nome')
            .agg({'Qtd_Limpa': lambda x: abs(x[x < 0].sum()), 'Valor_Limpo': lambda x: abs(x[x < 0].sum())})
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

        st.markdown("<br>", unsafe_allow_html=True)
        st.subheader("⚠️ Ranking: Top 10 Marcas com Maior Perda")

        df_top10_marcas = (
            df_filtered[(df_filtered['Valor_Limpo'] < 0) | (df_filtered['Qtd_Limpa'] < 0)]
            .groupby('Marca_Nome')
            .agg({'Qtd_Limpa': lambda x: abs(x[x < 0].sum()), 'Valor_Limpo': lambda x: abs(x[x < 0].sum())})
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

        # GRÁFICOS DE LINHAS (MARCAS)
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
                df_marca_qtd, x='Marca_Nome', y='Qtd_Limpa', text='Texto_Qtd', markers=True,
                labels={'Qtd_Limpa': 'Perda (Qtd)', 'Marca_Nome': 'Marca'}
            )
            fig_marca_qtd.update_traces(line_color='#ff7f0e', line_width=3, marker_size=7, textposition='top center')
            fig_marca_qtd.update_layout(template="plotly_dark", paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)", xaxis_title="", yaxis_title="", xaxis_tickangle=-45)
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
                df_marca_rs, x='Marca_Nome', y='Valor_Limpo', text='Texto_RS', markers=True,
                labels={'Valor_Limpo': 'Perda (R$)', 'Marca_Nome': 'Marca'}
            )
            fig_marca_rs.update_traces(line_color='#4ba3e3', line_width=3, marker_size=7, textposition='top center')
            fig_marca_rs.update_layout(template="plotly_dark", paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)", xaxis_title="", yaxis_title="", xaxis_tickangle=-45)
            st.plotly_chart(fig_marca_rs, use_container_width=True)

    except Exception as e:
        st.error(f"Erro ao processar a planilha: {e}")
else:
    st.info("👈 Faça o upload da sua planilha Excel no menu à esquerda para carregar o dashboard.")