import streamlit as st
import pandas as pd
import plotly.express as px
import os

# --- CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(
    page_title="Dashboard Executivo de Inventário",
    layout="wide",
    initial_sidebar_state="expanded"
)

# --- ESTILIZAÇÃO VISUAL (CSS) ---
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

# --- MÓDULO DE LOGIN E PERMISSÕES ---
def realizar_login():
    if "autenticado" not in st.session_state:
        st.session_state["autenticado"] = False

    if not st.session_state["autenticado"]:
        col1, col2, col3 = st.columns([1, 2, 1])
        with col2:
            st.title("🔐 Acesso ao Sistema")
            st.subheader("Faça login para continuar")
            
            with st.form("form_login"):
                usuario = st.text_input("Usuário")
                senha = st.text_input("Senha", type="password")
                btn_entrar = st.form_submit_button("Entrar")

                if btn_entrar:
                    # ALTERE AQUI PARA A SUA LÓGICA DE LOGIN OU CONSULTA NO BANCO
                    if usuario == "admin" and senha == "123456":
                        st.session_state["autenticado"] = True
                        st.session_state["usuario"] = usuario
                        st.success("Login realizado com sucesso!")
                        st.rerun()
                    else:
                        st.error("Usuário ou senha incorretos.")
        return False
    return True

# --- CARREGAMENTO E TRATAMENTO DE DADOS ---
def load_data():
    pasta_atual = os.path.dirname(os.path.abspath(__file__))
    file_path = os.path.join(pasta_atual, "STATUS DOS INVENTÁRIOS.xlsm")
    
    df = pd.read_excel(file_path, sheet_name="VALORES INVENTÁRIOS")
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

# --- EXECUÇÃO PRINCIPAL ---
if realizar_login():
    # Botão para Sair no Menu Lateral
    st.sidebar.write(f"👤 Usuário: **{st.session_state.get('usuario', 'Admin')}**")
    if st.sidebar.button("Sair / Logout"):
        st.session_state["autenticado"] = False
        st.rerun()

    try:
        df = load_data()

        # --- FILTROS SIDEBAR ---
        st.sidebar.title("Filtros")

        lojas = [x for x in sorted(df['Loja_Nome'].unique()) if x.lower() not in ['nan', 'none', '', 's/ centro']]
        lojas_sel = st.sidebar.multiselect("Selecione os Centros:", options=lojas, default=lojas)

        marcas = [x for x in sorted(df['Marca_Nome'].unique()) if x.lower() not in ['nan', 'none', '', 'sem marca']]
        marcas_sel = st.sidebar.multiselect("Selecione as Marcas:", options=marcas, default=marcas)

        df_filtered = df[
            (df['Loja_Nome'].isin(lojas_sel)) & 
            (df['Marca_Nome'].isin(marcas_sel))
        ]

        # --- PAINEL PRINCIPAL ---
        st.title("📊 Dashboard Executivo de Inventário")
        st.markdown("---")

        # KPIS
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

        # GRÁFICOS POR CENTRO
        graf_col1, graf_col2 = st.columns(2)

        with graf_col1:
            st.subheader("📦 Perda por Centro (Qtd)")
            df_qtd_lojas = (
                df_filtered[df_filtered['Qtd_Limpa'] < 0]
                .groupby('Loja_Nome')['Qtd_Limpa']
                .sum().abs().reset_index()
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
                .sum().abs().reset_index()
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

        # RANKINGS
        st.markdown("<br>", unsafe_allow_html=True)
        st.subheader("🏢 Ranking: Top 10 Centros com Maior Perda")

        df_top10_centros = (
            df_filtered[(df_filtered['Valor_Limpo'] < 0) | (df_filtered['Qtd_Limpa'] < 0)]
            .groupby('Loja_Nome')
            .agg({'Qtd_Limpa': lambda x: abs(x[x < 0].sum()), 'Valor_Limpo': lambda x: abs(x[x < 0].sum())})
            .reset_index().sort_values(by='Valor_Limpo', ascending=False).head(10)
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
            .reset_index().sort_values(by='Valor_Limpo', ascending=False).head(10)
        )
        df_top10_marcas.insert(0, 'Posição', [f"{i+1}º" for i in range(len(df_top10_marcas))])
        df_top10_marcas = df_top10_marcas[['Posição', 'Marca_Nome', 'Qtd_Limpa', 'Valor_Limpo']]
        df_top10_marcas.rename(columns={'Marca_Nome': 'Marca', 'Qtd_Limpa': 'Perda (Qtd)', 'Valor_Limpo': 'Perda (R$)'}, inplace=True)
        df_top10_marcas['Perda (Qtd)'] = df_top10_marcas['Perda (Qtd)'].apply(lambda x: f"-{x:,.0f} un")
        df_top10_marcas['Perda (R$)'] = df_top10_marcas['Perda (R$)'].apply(lambda x: f"R$ -{x:,.2f}")

        st.dataframe(df_top10_marcas, use_container_width=True, hide_index=True)

        # GRÁFICOS MARCAS
        st.markdown("<br>", unsafe_allow_html=True)
        marca_col1, marca_col2 = st.columns(2)
        df_perdas_marcas = df_filtered[(df_filtered['Valor_Limpo'] < 0) | (df_filtered['Qtd_Limpa'] < 0)]

        with marca_col1:
            st.subheader("📦 Perdas por Marca - Todas (Qtd)")
            df_marca_qtd = (
                df_perdas_marcas[df_perdas_marcas['Qtd_Limpa'] < 0]
                .groupby('Marca_Nome')['Qtd_Limpa'].sum().abs().reset_index().sort_values(by='Qtd_Limpa', ascending=False)
            )
            df_marca_qtd['Texto_Qtd'] = df_marca_qtd['Qtd_Limpa'].apply(lambda x: f"-{x:,.0f} un")

            fig_marca_qtd = px.line(df_marca_qtd, x='Marca_Nome', y='Qtd_Limpa', text='Texto_Qtd', markers=True, labels={'Qtd_Limpa': 'Perda (Qtd)', 'Marca_Nome': 'Marca'})
            fig_marca_qtd.update_traces(line_color='#ff7f0e', line_width=3, marker_size=7, textposition='top center')
            fig_marca_qtd.update_layout(template="plotly_dark", paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)", xaxis_title="", yaxis_title="", xaxis_tickangle=-45)
            st.plotly_chart(fig_marca_qtd, use_container_width=True)

        with marca_col2:
            st.subheader("🏷️ Perdas por Marca - Todas (R$)")
            df_marca_rs = (
                df_perdas_marcas[df_perdas_marcas['Valor_Limpo'] < 0]
                .groupby('Marca_Nome')['Valor_Limpo'].sum().abs().reset_index().sort_values(by='Valor_Limpo', ascending=False)
            )
            df_marca_rs['Texto_RS'] = df_marca_rs['Valor_Limpo'].apply(lambda x: f"-{x:,.0f}")

            fig_marca_rs = px.line(df_marca_rs, x='Marca_Nome', y='Valor_Limpo', text='Texto_RS', markers=True, labels={'Valor_Limpo': 'Perda (R$)', 'Marca_Nome': 'Marca'})
            fig_marca_rs.update_traces(line_color='#4ba3e3', line_width=3, marker_size=7, textposition='top center')
            fig_marca_rs.update_layout(template="plotly_dark", paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)", xaxis_title="", yaxis_title="", xaxis_tickangle=-45)
            st.plotly_chart(fig_marca_rs, use_container_width=True)

    except Exception as e:
        st.error(f"Erro ao carregar dados: {e}")