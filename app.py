import streamlit as st
import pandas as pd
import plotly.express as px

st.set_page_config(
    page_title="Dashboard Executivo de Inventário",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Estilização visual para manter a tela escura igual ao print
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
    # Encontra a linha exata do cabeçalho procurando pelos nomes das colunas
    df_preview = pd.read_excel(uploaded_file, engine='openpyxl', header=None, nrows=30)
    
    header_idx = 0
    for idx, row in df_preview.iterrows():
        row_str = [str(val).lower() for val in row.dropna().tolist()]
        texto_linha = " ".join(row_str)
        if any(term in texto_linha for term in ['qtd. um registro', 'montante em mi', 'fornecedor2', 'centro']):
            header_idx = idx
            break

    df = pd.read_excel(uploaded_file, engine='openpyxl', header=header_idx)
    df = df.dropna(how='all')
    
    # Normaliza nome das colunas
    df.columns = [str(c).replace('\n', ' ').strip() for c in df.columns]
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

def limpa_texto_seguro(val, padrao="Sem informação"):
    if pd.isna(val) or val is None:
        return padrao
    s = str(val).strip()
    if s.lower() in ['nan', 'none', '', 'null', '#n/a', '#valor!', '#ref!']:
        return padrao
    return s

# --- INTERFACE ---
st.sidebar.title("📁 Carregar Planilha")
uploaded_file = st.sidebar.file_uploader("Arraste e solte seu arquivo aqui (.xlsx ou .xlsm):", type=['xlsx', 'xlsm'])

if uploaded_file is not None:
    try:
        df_raw = load_data(uploaded_file)
        colunas = list(df_raw.columns)

        # Mapeamento 100% Automático das colunas exatas do seu relatório
        def buscar_coluna(nomes_prioritarios):
            for nome in nomes_prioritarios:
                for c in colunas:
                    if nome.lower() in c.lower():
                        return c
            return colunas[0]

        col_qtd = buscar_coluna(['Qtd. UM registro', 'qtd'])
        col_valor = buscar_coluna(['Montante em MI', 'montante', 'valor'])
        col_loja = buscar_coluna(['Centro', 'loja'])
        col_marca = buscar_coluna(['Fornecedor2', 'fornecedor 2', 'marca'])

        df = df_raw.copy()

        # Tratamento Numérico
        df['Qtd_Limpa'] = pd.to_numeric(df[col_qtd], errors='coerce').fillna(0.0).astype(float)
        df['Valor_Limpo'] = pd.to_numeric(df[col_valor], errors='coerce').fillna(0.0).astype(float)

        # Tratamento de Texto
        df['Loja_Nome'] = df[col_loja].apply(lambda x: limpa_texto_seguro(x, "Sem Centro"))
        df['Marca_Nome'] = df[col_marca].apply(lambda x: limpa_texto_seguro(x, "Sem Marca"))

        # Descarta linhas de totais ou sujeiras de cabeçalhos repetidos
        invalidos = ['TOTAL', 'RESULTADO', 'FORNECEDOR2', 'CENTRO', 'RÓTULOS DE LINHA', 'NAN', 'NONE']
        df = df[~df['Loja_Nome'].str.upper().isin(invalidos)]
        df = df[~df['Marca_Nome'].str.upper().isin(invalidos)]

        # --- FILTROS NO MENU LATERAL ---
        st.sidebar.title("Filtros")

        lojas = sorted([str(x) for x in df['Loja_Nome'].unique() if x])
        marcas = sorted([str(x) for x in df['Marca_Nome'].unique() if x])

        lojas_sel = st.sidebar.multiselect("Selecione os Centros:", options=lojas, default=lojas)
        marcas_sel = st.sidebar.multiselect("Selecione as Marcas:", options=marcas, default=marcas)

        df_filtered = df[
            (df['Loja_Nome'].isin(lojas_sel if lojas_sel else lojas)) & 
            (df['Marca_Nome'].isin(marcas_sel if marcas_sel else marcas))
        ]

        # --- KPIS EXECUTIVOS ---
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

        # --- GRÁFICOS LADO A LADO ---
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
                labels={'Qtd_Limpa': '', 'Loja_Nome': ''}
            )
            fig_qtd_lojas.update_traces(marker_color='#4ba3e3', textposition='inside')
            fig_qtd_lojas.update_layout(
                template="plotly_dark",
                paper_bgcolor="rgba(0,0,0,0)",
                plot_bgcolor="rgba(0,0,0,0)",
                xaxis_title="",
                yaxis_title="",
                xaxis_tickangle=-90
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
                df_lojas, x='Loja_Nome', y='Valor_Limpo', text='Texto_Valor',
                labels={'Valor_Limpo': '', 'Loja_Nome': ''}
            )
            fig_lojas.update_traces(marker_color='#70bbfd', textposition='inside')
            fig_lojas.update_layout(
                template="plotly_dark",
                paper_bgcolor="rgba(0,0,0,0)",
                plot_bgcolor="rgba(0,0,0,0)",
                xaxis_title="",
                yaxis_title="",
                xaxis_tickangle=-90
            )
            st.plotly_chart(fig_lojas, use_container_width=True)

    except Exception as e:
        st.error(f"Erro ao processar o arquivo: {e}")
else:
    st.info("👈 Arraste e solte a sua planilha no menu à esquerda para carregar o dashboard instantaneamente.")