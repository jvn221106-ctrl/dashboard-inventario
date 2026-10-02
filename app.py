import streamlit as st
import pandas as pd
import plotly.express as px
import json
import os
import hashlib
import requests
import io
import datetime
import re

# --- CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(
    page_title="Dashboard Executivo de Inventário - Vonny Cosméticos",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# --- ESTILIZAÇÃO VISUAL (TEMA ESCURO + BARRA DE ROLAGEM NOS MULTISELECTS) ---
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
    
    /* Permite rolagem horizontal e limita altura nos multiselects */
    div[data-baseweb="select"] > div:first-child {
        max-height: 80px !important;
        overflow-y: auto !important;
        overflow-x: auto !important;
        white-space: nowrap !important;
        flex-wrap: nowrap !important;
    }
    
    /* Estilização da barra de rolagem */
    div[data-baseweb="select"] > div:first-child::-webkit-scrollbar {
        height: 6px;
        width: 6px;
    }
    div[data-baseweb="select"] > div:first-child::-webkit-scrollbar-thumb {
        background: #4ba3e3;
        border-radius: 4px;
    }
    div[data-baseweb="select"] > div:first-child::-webkit-scrollbar-track {
        background: #1e222d;
    }
    </style>
""", unsafe_allow_html=True)

URL_EXCEL_NUVEM = "https://vonnycosmeticos-my.sharepoint.com/:x:/g/personal/josue_pereira_vonnycosmeticos_onmicrosoft_com/IQBPk7RDywHuR53pgILohLKbARZ0TkXuAjeJEfuUpfNehRM?download=1"

DB_FILE = "usuarios_db.json"
LOJAS_DB_FILE = "lojas_contatos_db.json"

# ==============================================================================
# DADOS INICIAIS DA TABELA DE LOJAS E CONTATOS
# ==============================================================================
DADOS_LOJAS_INICIAIS = [
    {
        "Nº LOJA": "B001",
        "REFERENCIA LOJA": "ARICANDUVA",
        "UF": "SP",
        "ESTADO": "SÃO PAULO",
        "NOME": "Jaqueline Aparecida Santos Laurentino",
        "EMAIL": "-",
        "TELEFONE": "11 95476-8160",
        "SETOR": "LOJA",
        "CARGO": "-",
        "N DO GRUPO": "2",
        "GRUPO": "LÍDERES LOJA",
        "SAP ID": "-",
        "Gmail padrão Recebimento": "-",
        "Coluna 1": "Atualizado em 22/09/2026 Whats Gerente Sara",
    },
    {
        "Nº LOJA": "B001",
        "REFERENCIA LOJA": "ARICANDUVA",
        "UF": "SP",
        "ESTADO": "SÃO PAULO",
        "NOME": "Jose Marcello Lobo Junior",
        "EMAIL": "jose.marcello@vonnycosmeticos.com.br",
        "TELEFONE": "11 97353-0039",
        "SETOR": "COORDENADOR",
        "CARGO": "1045 - Líder de Setor",
        "N DO GRUPO": "2",
        "GRUPO": "LÍDERES LOJA",
        "SAP ID": "JMARCELLO4863",
        "Gmail padrão Recebimento": "-",
        "Coluna 1": "Atualizado em 22/09/2026 Whats Gerente Sara",
    },
    {
        "Nº LOJA": "B001",
        "REFERENCIA LOJA": "ARICANDUVA",
        "UF": "SP",
        "ESTADO": "SÃO PAULO",
        "NOME": "Jaqueline Aparecida Santos Laurentino",
        "EMAIL": "-",
        "TELEFONE": "11 95476-8160",
        "SETOR": "LOJA",
        "CARGO": "-",
        "N DO GRUPO": "2",
        "GRUPO": "LÍDERES LOJA",
        "SAP ID": "-",
        "Gmail padrão Recebimento": "-",
        "Coluna 1": "Atualizado em 22/09/2026 Whats Gerente Sara",
    },
    {
        "Nº LOJA": "B001",
        "REFERENCIA LOJA": "ARICANDUVA",
        "UF": "SP",
        "ESTADO": "SÃO PAULO",
        "NOME": "Valéria Francisco Avelino",
        "EMAIL": "adm.aricanduva@vonnycosmeticos.com.br",
        "TELEFONE": "11 99634-8636",
        "SETOR": "ADMINISTRATIVO",
        "CARGO": "-",
        "N DO GRUPO": "3",
        "GRUPO": "ADMINISTRATIVO",
        "SAP ID": "TROCAS14 e TROCAS15",
        "Gmail padrão Recebimento": "-",
        "Coluna 1": "Atualizado em 22/09/2026 Whats Gerente Sara",
    },
    {
        "Nº LOJA": "B001",
        "REFERENCIA LOJA": "ARICANDUVA",
        "UF": "SP",
        "ESTADO": "SÃO PAULO",
        "NOME": "Esther Silva de Oliveira Pelizzari",
        "EMAIL": "trocas@vonnycosmeticos.com.br",
        "TELEFONE": "11 97234-3202",
        "SETOR": "RECEBIMENTO / TROCAS",
        "CARGO": "-",
        "N DO GRUPO": "2",
        "GRUPO": "LÍDERES LOJA",
        "SAP ID": "-",
        "Gmail padrão Recebimento": "-",
        "Coluna 1": "Atualizado em 22/09/2026 Whats Gerente Sara",
    },
    {
        "Nº LOJA": "B001",
        "REFERENCIA LOJA": "ARICANDUVA",
        "UF": "SP",
        "ESTADO": "SÃO PAULO",
        "NOME": "Fabio Joao de Andrade",
        "EMAIL": "recebimento@vonnycosmeticos.com.br",
        "TELEFONE": "11 97234-3202",
        "SETOR": "RECEBIMENTO",
        "CARGO": "1051 - Atendente de Loja (exp.) inicio em 01/09/2024",
        "N DO GRUPO": "5",
        "GRUPO": "RECEBIMENTO",
        "SAP ID": "-",
        "Gmail padrão Recebimento": "recebmpb001@gmail.com",
        "Coluna 1": "Atualizado em 22/09/2026 Whats Gerente Sara",
    },
    {
        "Nº LOJA": "B001",
        "REFERENCIA LOJA": "ARICANDUVA",
        "UF": "SP",
        "ESTADO": "SÃO PAULO",
        "NOME": "Veronica Bernardo de O Costa",
        "EMAIL": "veronica.bernardo@vonnycosmeticos.com.br",
        "TELEFONE": "11 97234-3202",
        "SETOR": "LOJA",
        "CARGO": "1051 - Atendente de Loja",
        "N DO GRUPO": "2",
        "GRUPO": "LÍDERES LOJA",
        "SAP ID": "VBERNARDO4956",
        "Gmail padrão Recebimento": "-",
        "Coluna 1": "Atualizado em 22/09/2026 Whats Gerente Sara",
    },
    {
        "Nº LOJA": "B002",
        "REFERENCIA LOJA": "SÃO JOSÉ DOS CAMPOS",
        "UF": "SP",
        "ESTADO": "SÃO JOSÉ DOS CAMPOS",
        "NOME": "Julio Fonseca",
        "EMAIL": "julio.fonseca@vonnycosmeticos.com.br",
        "TELEFONE": "11 94487-4329",
        "SETOR": "LOJA",
        "CARGO": "105 - Gerente",
        "N DO GRUPO": "1",
        "GRUPO": "GERENTES LOJA",
        "SAP ID": "-",
        "Gmail padrão Recebimento": "-",
        "Coluna 1": "Atualizado em 22/09/2026 Whats Gerente Julio Fonseca",
    },
    {
        "Nº LOJA": "B002",
        "REFERENCIA LOJA": "SÃO JOSÉ DOS CAMPOS",
        "UF": "SP",
        "ESTADO": "SÃO JOSÉ DOS CAMPOS",
        "NOME": "Adriana Alves de A Carneira",
        "EMAIL": "-",
        "TELEFONE": "11 97238-8029",
        "SETOR": "LOJA",
        "CARGO": "1045 - Lider de setor",
        "N DO GRUPO": "2",
        "GRUPO": "LÍDERES LOJA",
        "SAP ID": "8324",
        "Gmail padrão Recebimento": "-",
        "Coluna 1": "Atualizado em 22/09/2026 Whats Gerente Julio Fonseca",
    },
    {
        "Nº LOJA": "B002",
        "REFERENCIA LOJA": "SÃO JOSÉ DOS CAMPOS",
        "UF": "SP",
        "ESTADO": "SÃO JOSÉ DOS CAMPOS",
        "NOME": "Tatiane Tamizara Ferreira",
        "EMAIL": "tatiane.tamizara@vonnycosmeticos.com.br",
        "TELEFONE": "11 97238-8029",
        "SETOR": "COORDENADOR",
        "CARGO": "1045 - Lider de setor",
        "N DO GRUPO": "2",
        "GRUPO": "LÍDERES LOJA",
        "SAP ID": "-",
        "Gmail padrão Recebimento": "-",
        "Coluna 1": "Atualizado em 22/09/2026 Whats Gerente Julio Fonseca",
    },
    {
        "Nº LOJA": "B002",
        "REFERENCIA LOJA": "SÃO JOSÉ DOS CAMPOS",
        "UF": "SP",
        "ESTADO": "SÃO JOSÉ DOS CAMPOS",
        "NOME": "Adriana Leite",
        "EMAIL": "-",
        "TELEFONE": "11 97238-8029",
        "SETOR": "LOJA",
        "CARGO": "1045 - Líder de Setor",
        "N DO GRUPO": "2",
        "GRUPO": "LÍDERES LOJA",
        "SAP ID": "-",
        "Gmail padrão Recebimento": "-",
        "Coluna 1": "Atualizado em 22/09/2026 Whats Gerente Julio Fonseca",
    },
    {
        "Nº LOJA": "B002",
        "REFERENCIA LOJA": "SÃO JOSÉ DOS CAMPOS",
        "UF": "SP",
        "ESTADO": "SÃO JOSÉ DOS CAMPOS",
        "NOME": "Fernanda Oliveira de Castro",
        "EMAIL": "-",
        "TELEFONE": "11 97238-8029",
        "SETOR": "LOJA",
        "CARGO": "1045 - Lider de setor",
        "N DO GRUPO": "2",
        "GRUPO": "LÍDERES LOJA",
        "SAP ID": "-",
        "Gmail padrão Recebimento": "-",
        "Coluna 1": "Atualizado em 22/09/2026 Whats Gerente Julio Fonseca",
    },
    {
        "Nº LOJA": "B002",
        "REFERENCIA LOJA": "SÃO JOSÉ DOS CAMPOS",
        "UF": "SP",
        "ESTADO": "SÃO JOSÉ DOS CAMPOS",
        "NOME": "Stephany Mariana dos Reis Alves",
        "EMAIL": "adm.sjc@vonnycosmeticos.com.br",
        "TELEFONE": "11 99710-6402",
        "SETOR": "ADMINISTRATIVO",
        "CARGO": "-",
        "N DO GRUPO": "3",
        "GRUPO": "ADMINISTRATIVO",
        "SAP ID": "-",
        "Gmail padrão Recebimento": "-",
        "Coluna 1": "Atualizado em 22/09/2026 Whats Gerente Julio Fonseca",
    },
    {
        "Nº LOJA": "B002",
        "REFERENCIA LOJA": "SÃO JOSÉ DOS CAMPOS",
        "UF": "SP",
        "ESTADO": "SÃO JOSÉ DOS CAMPOS",
        "NOME": "Lassira Nascimento Campelo",
        "EMAIL": "trocas.sjc@vonnycosmeticos.com.br",
        "TELEFONE": "11 97381-3840",
        "SETOR": "TROCAS",
        "CARGO": "em experiencia",
        "N DO GRUPO": "4",
        "GRUPO": "TROCAS",
        "SAP ID": "8121",
        "Gmail padrão Recebimento": "-",
        "Coluna 1": "Atualizado em 22/09/2026 Whats Gerente Julio Fonseca",
    },
    {
        "Nº LOJA": "B002",
        "REFERENCIA LOJA": "SÃO JOSÉ DOS CAMPOS",
        "UF": "SP",
        "ESTADO": "SÃO JOSÉ DOS CAMPOS",
        "NOME": "Ashley Raiza Santas Alves",
        "EMAIL": "recebmpsj@vonnycosmeticos.com.br",
        "TELEFONE": "11 97381-3840",
        "SETOR": "RECEBIMENTO",
        "CARGO": "em experiencia",
        "N DO GRUPO": "5",
        "GRUPO": "RECEBIMENTO",
        "SAP ID": "7744",
        "Gmail padrão Recebimento": "recebmpb002@gmail.com",
        "Coluna 1": "Atualizado em 22/09/2026 Whats Gerente Julio Fonseca",
    },
    {
        "Nº LOJA": "B006",
        "REFERENCIA LOJA": "SÃO BERNARDO DO CAMPO",
        "UF": "SP",
        "ESTADO": "SÃO PAULO",
        "NOME": "Fabiana Aparecida Bertassi",
        "EMAIL": "fabiana.bertassi@vonnycosmeticos.com.br",
        "TELEFONE": "11 97350-7627",
        "SETOR": "LOJA",
        "CARGO": "105 - Gerente",
        "N DO GRUPO": "1",
        "GRUPO": "GERENTES LOJA",
        "SAP ID": "-",
        "Gmail padrão Recebimento": "-",
        "Coluna 1": "Atualizado em 22/09/2026 Whats Gerente Fabiana Bertassi",
    },
    {
        "Nº LOJA": "B006",
        "REFERENCIA LOJA": "SÃO BERNARDO DO CAMPO",
        "UF": "SP",
        "ESTADO": "SÃO PAULO",
        "NOME": "Angelica do Carmo Santos Romao",
        "EMAIL": "angelica.santos@vonnycosmeticos.com.br",
        "TELEFONE": "12 94456-9174",
        "SETOR": "LOJA",
        "CARGO": "-",
        "N DO GRUPO": "2",
        "GRUPO": "LÍDERES LOJA",
        "SAP ID": "6575",
        "Gmail padrão Recebimento": "-",
        "Coluna 1": "Atualizado em 22/09/2026 Whats Gerente Fabiana Bertassi",
    },
    {
        "Nº LOJA": "B006",
        "REFERENCIA LOJA": "SÃO BERNARDO DO CAMPO",
        "UF": "SP",
        "ESTADO": "SÃO PAULO",
        "NOME": "Daniela Rodrigues Martins",
        "EMAIL": "adm.sbc@vonnycosmeticos.com.br",
        "TELEFONE": "11 99737-6499",
        "SETOR": "ADMINISTRATIVO",
        "CARGO": "1045 - Lider de Setor",
        "N DO GRUPO": "3",
        "GRUPO": "ADMINISTRATIVO",
        "SAP ID": "CGOMES / TROCASSBC6988",
        "Gmail padrão Recebimento": "-",
        "Coluna 1": "Atualizado em 22/09/2026 Whats Gerente Fabiana Bertassi",
    },
    {
        "Nº LOJA": "B006",
        "REFERENCIA LOJA": "SÃO BERNARDO DO CAMPO",
        "UF": "SP",
        "ESTADO": "SÃO PAULO",
        "NOME": "Melissa Santos Nunes",
        "EMAIL": "trocas.sbc@vonnycosmeticos.com.br",
        "TELEFONE": "11 99799-2986",
        "SETOR": "TROCAS",
        "CARGO": "52 - Auxiliar de Recebimento",
        "N DO GRUPO": "4",
        "GRUPO": "TROCAS",
        "SAP ID": "-",
        "Gmail padrão Recebimento": "-",
        "Coluna 1": "Atualizado em 22/09/2026 Whats Gerente Fabiana Bertassi",
    },
    {
        "Nº LOJA": "B006",
        "REFERENCIA LOJA": "SÃO BERNARDO DO CAMPO",
        "UF": "SP",
        "ESTADO": "SÃO PAULO",
        "NOME": "Giovanna Luppi da Silva",
        "EMAIL": "recebmp06@vonnycosmeticos.com.br",
        "TELEFONE": "12 99799-2986",
        "SETOR": "RECEBIMENTO",
        "CARGO": "-",
        "N DO GRUPO": "5",
        "GRUPO": "RECEBIMENTO",
        "SAP ID": "8920",
        "Gmail padrão Recebimento": "-",
        "Coluna 1": "Atualizado em 22/09/2026 Whats Gerente Fabiana Bertassi",
    },
    {
        "Nº LOJA": "B007",
        "REFERENCIA LOJA": "PINHEIROS",
        "UF": "SP",
        "ESTADO": "SÃO PAULO",
        "NOME": "Vanessa Tais da Silva",
        "EMAIL": "Vanessa.tais@vonnycosmeticos.com.br",
        "TELEFONE": "11 96865-4097",
        "SETOR": "LOJA",
        "CARGO": "105 - Gerente",
        "N DO GRUPO": "1",
        "GRUPO": "GERENTES LOJA",
        "SAP ID": "-",
        "Gmail padrão Recebimento": "-",
        "Coluna 1": "Atualizado em 22/09/2026 Whats Gerente VANESSA TAIS",
    },
    {
        "Nº LOJA": "B007",
        "REFERENCIA LOJA": "PINHEIROS",
        "UF": "SP",
        "ESTADO": "SÃO PAULO",
        "NOME": "Elenilza tavares (Eli)",
        "EMAIL": "elenilza.tavares@vonnycosmeticos.com.br",
        "TELEFONE": "11 94448-5852",
        "SETOR": "LOJA",
        "CARGO": "-",
        "N DO GRUPO": "2",
        "GRUPO": "LÍDERES LOJA",
        "SAP ID": "-",
        "Gmail padrão Recebimento": "-",
        "Coluna 1": "Atualizado em 22/09/2026 Whats Gerente VANESSA TAIS",
    },
    {
        "Nº LOJA": "B007",
        "REFERENCIA LOJA": "PINHEIROS",
        "UF": "SP",
        "ESTADO": "SÃO PAULO",
        "NOME": "Zirlene Rosa da Silva",
        "EMAIL": "adm.pinheiros@vonnycosmeticos.com.br",
        "TELEFONE": "11 99713-4472",
        "SETOR": "ADMINISTRATIVO",
        "CARGO": "-",
        "N DO GRUPO": "3",
        "GRUPO": "ADMINISTRATIVO",
        "SAP ID": "5589",
        "Gmail padrão Recebimento": "-",
        "Coluna 1": "Atualizado em 22/09/2026 Whats Gerente VANESSA TAIS",
    },
    {
        "Nº LOJA": "B007",
        "REFERENCIA LOJA": "PINHEIROS",
        "UF": "SP",
        "ESTADO": "SÃO PAULO",
        "NOME": "Juan Lopes Ferreira",
        "EMAIL": "trocas.pinheiros@vonnycosmeticos.com.br",
        "TELEFONE": "11 97472-3774",
        "SETOR": "RECEBIMENTO / TROCAS",
        "CARGO": "-",
        "N DO GRUPO": "7",
        "GRUPO": "RECEBIMENTO / TROCAS",
        "SAP ID": "-",
        "Gmail padrão Recebimento": "-",
        "Coluna 1": "Atualizado em 22/09/2026 Whats Gerente VANESSA TAIS",
    },
    {
        "Nº LOJA": "B008",
        "REFERENCIA LOJA": "CARAPICUIBA",
        "UF": "SP",
        "ESTADO": "SÃO PAULO",
        "NOME": "Yara Valéria da Silva Barros",
        "EMAIL": "yara.silva@vonnycosmeticos.com.br",
        "TELEFONE": "11 97144-7354",
        "SETOR": "LOJA",
        "CARGO": "105 - Gerente",
        "N DO GRUPO": "1",
        "GRUPO": "GERENTES LOJA",
        "SAP ID": "-",
        "Gmail padrão Recebimento": "-",
        "Coluna 1": "Atualizado em 22/09/2026 Whats Gerente Yara Valéria",
    },
    {
        "Nº LOJA": "B008",
        "REFERENCIA LOJA": "CARAPICUIBA",
        "UF": "SP",
        "ESTADO": "SÃO PAULO",
        "NOME": "Cibele Cristina Santos Gonçalves",
        "EMAIL": "cibele.santos@vonnycosmeticos.com.br",
        "TELEFONE": "11 97295-0870",
        "SETOR": "LOJA",
        "CARGO": "-",
        "N DO GRUPO": "2",
        "GRUPO": "LÍDERES LOJA",
        "SAP ID": "-",
        "Gmail padrão Recebimento": "-",
        "Coluna 1": "Atualizado em 22/09/2026 Whats Gerente Yara Valéria",
    },
    {
        "Nº LOJA": "B008",
        "REFERENCIA LOJA": "CARAPICUIBA",
        "UF": "SP",
        "ESTADO": "SÃO PAULO",
        "NOME": "Sirlei Nair da Silva",
        "EMAIL": "sirlei.nair@vonnycosmeticos.com.br",
        "TELEFONE": "11 95772-2374",
        "SETOR": "ADMINISTRATIVO",
        "CARGO": "1045 - Lider de Setor",
        "N DO GRUPO": "3",
        "GRUPO": "ADMINISTRATIVO",
        "SAP ID": "-",
        "Gmail padrão Recebimento": "-",
        "Coluna 1": "Atualizado em 22/09/2026 Whats Gerente Yara Valéria",
    },
    {
        "Nº LOJA": "B008",
        "REFERENCIA LOJA": "CARAPICUIBA",
        "UF": "SP",
        "ESTADO": "SÃO PAULO",
        "NOME": "Bruno Matheus Trindade Rocha",
        "EMAIL": "trocas.carapicuiba@vonnycosmeticos.com.br",
        "TELEFONE": "11 95051-2839",
        "SETOR": "RECEBIMENTO / TROCAS",
        "CARGO": "-",
        "N DO GRUPO": "7",
        "GRUPO": "RECEBIMENTO / TROCAS",
        "SAP ID": "-",
        "Gmail padrão Recebimento": "recebmpb008@gmail.com",
        "Coluna 1": "Atualizado em 22/09/2026 Whats Gerente Yara Valéria",
    },
    {
        "Nº LOJA": "B009",
        "REFERENCIA LOJA": "TATUAPÉ",
        "UF": "SP",
        "ESTADO": "SÃO PAULO",
        "NOME": "Josemary Bezerra Barbosa",
        "EMAIL": "josemary.bezerra@vonnycosmeticos.com.br",
        "TELEFONE": "11 99903-9657",
        "SETOR": "LOJA",
        "CARGO": "105 - Gerente",
        "N DO GRUPO": "1",
        "GRUPO": "GERENTES LOJA",
        "SAP ID": "-",
        "Gmail padrão Recebimento": "-",
        "Coluna 1": "Atualizado em 24/09/2026 Whats Gerente Josemary Bezerra",
    },
    {
        "Nº LOJA": "B009",
        "REFERENCIA LOJA": "TATUAPÉ",
        "UF": "SP",
        "ESTADO": "SÃO PAULO",
        "NOME": "Kelly Cristina dos Santos",
        "EMAIL": "Kelly.santos@vonnycosmeticos.com.br",
        "TELEFONE": "11 97454-8417",
        "SETOR": "LOJA",
        "CARGO": "Transferida de Paulista",
        "N DO GRUPO": "2",
        "GRUPO": "LÍDERES LOJA",
        "SAP ID": "-",
        "Gmail padrão Recebimento": "-",
        "Coluna 1": "Atualizado em 24/09/2026 Whats Gerente Josemary Bezerra",
    },
    {
        "Nº LOJA": "B009",
        "REFERENCIA LOJA": "TATUAPÉ",
        "UF": "SP",
        "ESTADO": "SÃO PAULO",
        "NOME": "Fabiana Silva Oliveira Paes Landim",
        "EMAIL": "adm.tatuape@vonnycosmeticos.com.br",
        "TELEFONE": "11 96398-2127",
        "SETOR": "ADMINISTRATIVO",
        "CARGO": "-",
        "N DO GRUPO": "3",
        "GRUPO": "ADMINISTRATIVO",
        "SAP ID": "-",
        "Gmail padrão Recebimento": "-",
        "Coluna 1": "Atualizado em 24/09/2026 Whats Gerente Josemary Bezerra",
    },
    {
        "Nº LOJA": "B009",
        "REFERENCIA LOJA": "TATUAPÉ",
        "UF": "SP",
        "ESTADO": "SÃO PAULO",
        "NOME": "Gabriel Silva Cardoso",
        "EMAIL": "trocas.tatuape@vonnycosmeticos.com.br",
        "TELEFONE": "11 94308-1547",
        "SETOR": "TROCAS",
        "CARGO": "-",
        "N DO GRUPO": "4",
        "GRUPO": "TROCAS",
        "SAP ID": "8587",
        "Gmail padrão Recebimento": "recebmpb009@gmail.com",
        "Coluna 1": "Atualizado em 24/09/2026 Whats Gerente Josemary Bezerra",
    },
    {
        "Nº LOJA": "B009",
        "REFERENCIA LOJA": "TATUAPÉ",
        "UF": "SP",
        "ESTADO": "SÃO PAULO",
        "NOME": "Luan Santiago da Silva",
        "EMAIL": "trocas.tatuape@vonnycosmeticos.com.br",
        "TELEFONE": "11 94308-1547",
        "SETOR": "RECEBIMENTO",
        "CARGO": "-",
        "N DO GRUPO": "5",
        "GRUPO": "RECEBIMENTO",
        "SAP ID": "7385",
        "Gmail padrão Recebimento": "recebmpb009@gmail.com",
        "Coluna 1": "Atualizado em 24/09/2026 Whats Gerente Josemary Bezerra",
    },
    {
        "Nº LOJA": "B009",
        "REFERENCIA LOJA": "TATUAPÉ",
        "UF": "SP",
        "ESTADO": "SÃO PAULO",
        "NOME": "Katiusia B. Borges Vieira",
        "EMAIL": "trocas.tatuape@vonnycosmeticos.com.br",
        "TELEFONE": "11 94308-1547",
        "SETOR": "RECEBIMENTO",
        "CARGO": "Lider do Recebimento",
        "N DO GRUPO": "5",
        "GRUPO": "RECEBIMENTO",
        "SAP ID": "5685",
        "Gmail padrão Recebimento": "recebmpb009@gmail.com",
        "Coluna 1": "Atualizado em 24/09/2026 Whats Gerente Josemary Bezerra",
    },
    {
        "Nº LOJA": "B010",
        "REFERENCIA LOJA": "BRÁS I",
        "UF": "SP",
        "ESTADO": "SÃO PAULO",
        "NOME": "Maria José Beserra A. Oliveira (Mazé)",
        "EMAIL": "maria.beserra@vonnycosmeticos.com.br",
        "TELEFONE": "11 91123-6833",
        "SETOR": "LOJA",
        "CARGO": "105 - Gerente",
        "N DO GRUPO": "1",
        "GRUPO": "GERENTES LOJA",
        "SAP ID": "MBESERRA5541",
        "Gmail padrão Recebimento": "-",
        "Coluna 1": "Atualizado em 22/09/2026 Whats Gerente Maria José (Maze)",
    },
    {
        "Nº LOJA": "B010",
        "REFERENCIA LOJA": "BRÁS I",
        "UF": "SP",
        "ESTADO": "SÃO PAULO",
        "NOME": "Jeane Lopes da Rocha",
        "EMAIL": "jeane.lopes@vonnycosmeticos.com.br",
        "TELEFONE": "11 99877-6081",
        "SETOR": "LOJA",
        "CARGO": "1031 - Coordenador de Setor",
        "N DO GRUPO": "2",
        "GRUPO": "LÍDERES LOJA",
        "SAP ID": "CNUNES3723",
        "Gmail padrão Recebimento": "-",
        "Coluna 1": "Atualizado em 22/09/2026 Whats Gerente Maria José (Maze)",
    },
    {
        "Nº LOJA": "B010",
        "REFERENCIA LOJA": "BRÁS I",
        "UF": "SP",
        "ESTADO": "SÃO PAULO",
        "NOME": "Liliane Tamires Souza Ribeiro",
        "EMAIL": "adm.bras@vonnycosmeticos.com.br",
        "TELEFONE": "11 94154-5229",
        "SETOR": "ADMINISTRATIVO",
        "CARGO": "-",
        "N DO GRUPO": "3",
        "GRUPO": "ADMINISTRATIVO",
        "SAP ID": "5970",
        "Gmail padrão Recebimento": "-",
        "Coluna 1": "Atualizado em 22/09/2026 Whats Gerente Maria José (Maze)",
    },
    {
        "Nº LOJA": "B010",
        "REFERENCIA LOJA": "BRÁS I",
        "UF": "SP",
        "ESTADO": "SÃO PAULO",
        "NOME": "Elaine Cruz de Oliveira",
        "EMAIL": "trocas.bras@vonnycosmeticos.com.br",
        "TELEFONE": "11 94136-7959",
        "SETOR": "TROCAS",
        "CARGO": "52 - Auxiliar de Recebimento",
        "N DO GRUPO": "4",
        "GRUPO": "TROCAS",
        "SAP ID": "MPBR2 / JMOREIRA6081",
        "Gmail padrão Recebimento": "-",
        "Coluna 1": "Atualizado em 22/09/2026 Whats Gerente Maria José (Maze)",
    },
    {
        "Nº LOJA": "B010",
        "REFERENCIA LOJA": "BRÁS I",
        "UF": "SP",
        "ESTADO": "SÃO PAULO",
        "NOME": "Daniel da Silva Melo",
        "EMAIL": "recebmp10@vonnycosmeticos.com.br",
        "TELEFONE": "11 94136-7959",
        "SETOR": "RECEBIMENTO",
        "CARGO": "-",
        "N DO GRUPO": "5",
        "GRUPO": "RECEBIMENTO",
        "SAP ID": "MPBR16503",
        "Gmail padrão Recebimento": "recebmpb010@gmail.com",
        "Coluna 1": "Atualizado em 22/09/2026 Whats Gerente Maria José (Maze)",
    },
    {
        "Nº LOJA": "B011",
        "REFERENCIA LOJA": "SÃO MIGUEL",
        "UF": "SP",
        "ESTADO": "SÃO PAULO",
        "NOME": "Gislaine Barra",
        "EMAIL": "gislaine.barra@vonnycosmeticos.com.br",
        "TELEFONE": "11 99585-0678",
        "SETOR": "LOJA",
        "CARGO": "105 - Gerente",
        "N DO GRUPO": "1",
        "GRUPO": "GERENTES LOJA",
        "SAP ID": "JNASCIMENTO",
        "Gmail padrão Recebimento": "-",
        "Coluna 1": "Atualizado em 22/09/2026 Whats Gerente Gislaine Barra",
    },
    {
        "Nº LOJA": "B011",
        "REFERENCIA LOJA": "SÃO MIGUEL",
        "UF": "SP",
        "ESTADO": "SÃO PAULO",
        "NOME": "Thais Santana Vilela Amorim",
        "EMAIL": "thais.amorim@vonnycosmeticos.com.br",
        "TELEFONE": "11 94133-2226",
        "SETOR": "LOJA",
        "CARGO": "-",
        "N DO GRUPO": "2",
        "GRUPO": "LÍDERES LOJA",
        "SAP ID": "-",
        "Gmail padrão Recebimento": "-",
        "Coluna 1": "Atualizado em 22/09/2026 Whats Gerente Gislaine Barra",
    },
    {
        "Nº LOJA": "B011",
        "REFERENCIA LOJA": "SÃO MIGUEL",
        "UF": "SP",
        "ESTADO": "SÃO PAULO",
        "NOME": "Maria José de Lima Andrade",
        "EMAIL": "adm.saomiguel@vonnycosmeticos.com.br",
        "TELEFONE": "11 97554-4567",
        "SETOR": "ADMINISTRATIVO",
        "CARGO": "-",
        "N DO GRUPO": "3",
        "GRUPO": "ADMINISTRATIVO",
        "SAP ID": "TROCAS5078",
        "Gmail padrão Recebimento": "-",
        "Coluna 1": "Atualizado em 22/09/2026 Whats Gerente Gislaine Barra",
    },
    {
        "Nº LOJA": "B011",
        "REFERENCIA LOJA": "SÃO MIGUEL",
        "UF": "SP",
        "ESTADO": "SÃO PAULO",
        "NOME": "Bianca Alves dos Santos",
        "EMAIL": "recebmp11@vonnycosmeticos.com.br",
        "TELEFONE": "11 97318-5784",
        "SETOR": "RECEBIMENTO / TROCAS",
        "CARGO": "-",
        "N DO GRUPO": "5",
        "GRUPO": "RECEBIMENTO / TROCAS",
        "SAP ID": "-",
        "Gmail padrão Recebimento": "-",
        "Coluna 1": "Atualizado em 22/09/2026 Whats Gerente Gislaine Barra",
    },
    {
        "Nº LOJA": "B011",
        "REFERENCIA LOJA": "SÃO MIGUEL",
        "UF": "SP",
        "ESTADO": "SÃO PAULO",
        "NOME": "Italo David Souza Marcelino",
        "EMAIL": "recebmp11@vonnycosmeticos.com.br",
        "TELEFONE": "11 97318-5784",
        "SETOR": "RECEBIMENTO / TROCAS",
        "CARGO": "LIDER RECEBIMENTO",
        "N DO GRUPO": "5",
        "GRUPO": "RECEBIMENTO / TROCAS",
        "SAP ID": "-",
        "Gmail padrão Recebimento": "recebmpb011@gmail.com",
        "Coluna 1": "Atualizado em 22/09/2026 Whats Gerente Gislaine Barra",
    },
    {
        "Nº LOJA": "B012",
        "REFERENCIA LOJA": "BRÁS II",
        "UF": "SP",
        "ESTADO": "SÃO PAULO",
        "NOME": "Thamires Conceição dos Santos",
        "EMAIL": "thamires.conceicao@vonnycosmeticos.com.br",
        "TELEFONE": "11 99762-5542",
        "SETOR": "LOJA",
        "CARGO": "105 - Gerente",
        "N DO GRUPO": "1",
        "GRUPO": "GERENTES LOJA",
        "SAP ID": "TCONCEIÇÃO5509",
        "Gmail padrão Recebimento": "-",
        "Coluna 1": "Atualizado em 29/09/2026 Whats Gerente Thamires Conceição",
    },
    {
        "Nº LOJA": "B012",
        "REFERENCIA LOJA": "BRÁS II",
        "UF": "SP",
        "ESTADO": "SÃO PAULO",
        "NOME": "Caio vinicios rabelo primo",
        "EMAIL": "-",
        "TELEFONE": "11 97182-8306",
        "SETOR": "LOJA",
        "CARGO": "em Experiência",
        "N DO GRUPO": "2",
        "GRUPO": "LÍDERES LOJA",
        "SAP ID": "-",
        "Gmail padrão Recebimento": "-",
        "Coluna 1": "Atualizado em 29/09/2026 Whats Gerente Thamires Conceição",
    },
    {
        "Nº LOJA": "B012",
        "REFERENCIA LOJA": "BRÁS II",
        "UF": "SP",
        "ESTADO": "SÃO PAULO",
        "NOME": "Ingrid dos Santos Queiroz de Almeida",
        "EMAIL": "adm.bras02@vonnycosmeticos.com.br",
        "TELEFONE": "11 94750-8810",
        "SETOR": "ADMINISTRATIVO",
        "CARGO": "-",
        "N DO GRUPO": "3",
        "GRUPO": "ADMINISTRATIVO",
        "SAP ID": "TROCASBRAS2",
        "Gmail padrão Recebimento": "-",
        "Coluna 1": "Atualizado em 29/09/2026 Whats Gerente Thamires Conceição",
    },
    {
        "Nº LOJA": "B012",
        "REFERENCIA LOJA": "BRÁS II",
        "UF": "SP",
        "ESTADO": "SÃO PAULO",
        "NOME": "Priscila Rosa de s.c legere",
        "EMAIL": "trocas.bras02@vonnycosmeticos.com.br",
        "TELEFONE": "11 96193-5405",
        "SETOR": "RECEBIMENTO",
        "CARGO": "-",
        "N DO GRUPO": "7",
        "GRUPO": "RECEBIMENTO / TROCAS",
        "SAP ID": "-",
        "Gmail padrão Recebimento": "-",
        "Coluna 1": "Atualizado em 29/09/2026 Whats Gerente Thamires Conceição",
    },
    {
        "Nº LOJA": "B012",
        "REFERENCIA LOJA": "BRÁS II",
        "UF": "SP",
        "ESTADO": "SÃO PAULO",
        "NOME": "Marli Pereira da Rocha",
        "EMAIL": "trocas.bras02@vonnycosmeticos.com.br",
        "TELEFONE": "11 96193-5405",
        "SETOR": "TROCAS",
        "CARGO": "-",
        "N DO GRUPO": "7",
        "GRUPO": "RECEBIMENTO / TROCAS",
        "SAP ID": "-",
        "Gmail padrão Recebimento": "recebmpb012@gmail.com",
        "Coluna 1": "Atualizado em 29/09/2026 Whats Gerente Thamires Conceição",
    },
    {
        "Nº LOJA": "B013",
        "REFERENCIA LOJA": "FLORIANOPOLIS",
        "UF": "SC",
        "ESTADO": "SANTA CATARINA",
        "NOME": "Vera Lúcia dos Santos Silva",
        "EMAIL": "vera.silva@vonnycosmeticos.com.br",
        "TELEFONE": "11 97590-1340",
        "SETOR": "LOJA",
        "CARGO": "105 - Gerente",
        "N DO GRUPO": "1",
        "GRUPO": "GERENTES LOJA",
        "SAP ID": "-",
        "Gmail padrão Recebimento": "-",
        "Coluna 1": "Atualizado em 23/09/2026 Whats Gerente Vera Lucia",
    },
    {
        "Nº LOJA": "B013",
        "REFERENCIA LOJA": "FLORIANOPOLIS",
        "UF": "SC",
        "ESTADO": "SANTA CATARINA",
        "NOME": "Juliana Macário de Lima",
        "EMAIL": "juliana.lima@vonnycosmeticos.com.br",
        "TELEFONE": "11 94377-4346",
        "SETOR": "LOJA",
        "CARGO": "gerente operacional",
        "N DO GRUPO": "2",
        "GRUPO": "LÍDERES LOJA",
        "SAP ID": "-",
        "Gmail padrão Recebimento": "-",
        "Coluna 1": "Atualizado em 23/09/2026 Whats Gerente Vera Lucia",
    },
    {
        "Nº LOJA": "B013",
        "REFERENCIA LOJA": "FLORIANOPOLIS",
        "UF": "SC",
        "ESTADO": "SANTA CATARINA",
        "NOME": "Micheli das graças vieira",
        "EMAIL": "adm.florianopolis@vonnycosmeticos.com.br",
        "TELEFONE": "11 97104-3293",
        "SETOR": "ADMINISTRATIVO",
        "CARGO": "-",
        "N DO GRUPO": "3",
        "GRUPO": "ADMINISTRATIVO",
        "SAP ID": "TROCAFLO",
        "Gmail padrão Recebimento": "-",
        "Coluna 1": "Atualizado em 23/09/2026 Whats Gerente Vera Lucia",
    },
    {
        "Nº LOJA": "B013",
        "REFERENCIA LOJA": "FLORIANOPOLIS",
        "UF": "SC",
        "ESTADO": "SANTA CATARINA",
        "NOME": "Paloma dos Santos Costa",
        "EMAIL": "recebmp13@vonnycosmeticos.com.br",
        "TELEFONE": "11 94212-8324",
        "SETOR": "RECEBIMENTO / TROCAS",
        "CARGO": "-",
        "N DO GRUPO": "7",
        "GRUPO": "RECEBIMENTO / TROCAS",
        "SAP ID": "-",
        "Gmail padrão Recebimento": "recebmpb013@gmail.com",
        "Coluna 1": "Atualizado em 23/09/2026 Whats Gerente Vera Lucia",
    },
    {
        "Nº LOJA": "B015",
        "REFERENCIA LOJA": "PORTO ALEGRE",
        "UF": "RS",
        "ESTADO": "RIO GRANDE DO SUL",
        "NOME": "Vanessa de Mello Amaral",
        "EMAIL": "vanessa.amaral@vonnycosmeticos.com.br",
        "TELEFONE": "11 95577-0860",
        "SETOR": "LOJA",
        "CARGO": "105 - Gerente",
        "N DO GRUPO": "1",
        "GRUPO": "GERENTES LOJA",
        "SAP ID": "-",
        "Gmail padrão Recebimento": "-",
        "Coluna 1": "Atualizado em 22/09/2026 Whats Gerente Vanessa de Mello",
    },
    {
        "Nº LOJA": "B015",
        "REFERENCIA LOJA": "PORTO ALEGRE",
        "UF": "RS",
        "ESTADO": "RIO GRANDE DO SUL",
        "NOME": "Luciana Inácio",
        "EMAIL": "luciana.inacio@vonnycosmeticos.com.br",
        "TELEFONE": "11 99878-2591",
        "SETOR": "LOJA",
        "CARGO": "-",
        "N DO GRUPO": "2",
        "GRUPO": "LÍDERES LOJA",
        "SAP ID": "-",
        "Gmail padrão Recebimento": "-",
        "Coluna 1": "Atualizado em 22/09/2026 Whats Gerente Vanessa de Mello",
    },
    {
        "Nº LOJA": "B015",
        "REFERENCIA LOJA": "PORTO ALEGRE",
        "UF": "RS",
        "ESTADO": "RIO GRANDE DO SUL",
        "NOME": "Mariana sasha Bastos",
        "EMAIL": "adm.portoalegre@vonnycosmeticos.com.br",
        "TELEFONE": "11 95669-0471",
        "SETOR": "ADMINISTRATIVO",
        "CARGO": "-",
        "N DO GRUPO": "3",
        "GRUPO": "ADMINISTRATIVO",
        "SAP ID": "-",
        "Gmail padrão Recebimento": "-",
        "Coluna 1": "Atualizado em 22/09/2026 Whats Gerente Vanessa de Mello",
    },
    {
        "Nº LOJA": "B015",
        "REFERENCIA LOJA": "PORTO ALEGRE",
        "UF": "RS",
        "ESTADO": "RIO GRANDE DO SUL",
        "NOME": "Tamires de Oliveira",
        "EMAIL": "recebmp15@vonnycosmeticos.com.br",
        "TELEFONE": "11 96417-3698",
        "SETOR": "RECEBIMENTO / TROCAS",
        "CARGO": "-",
        "N DO GRUPO": "7",
        "GRUPO": "RECEBIMENTO / TROCAS",
        "SAP ID": "-",
        "Gmail padrão Recebimento": "recebmpb015@gmail.com",
        "Coluna 1": "Atualizado em 22/09/2026 Whats Gerente Vanessa de Mello",
    },
    {
        "Nº LOJA": "B016",
        "REFERENCIA LOJA": "BALNEÁRIO CAMBORIÚ",
        "UF": "SC",
        "ESTADO": "SANTA CATARINA",
        "NOME": "Claudineia Mendes de L Araujo",
        "EMAIL": "claudineia.mendes@vonnycosmeticos.com.br",
        "TELEFONE": "11 97398-3216",
        "SETOR": "LOJA",
        "CARGO": "105 - Gerente",
        "N DO GRUPO": "1",
        "GRUPO": "GERENTES LOJA",
        "SAP ID": "CMENDES6042",
        "Gmail padrão Recebimento": "-",
        "Coluna 1": "Atualizado em 22/09/2026 Whats Gerente Claudineia Mendes",
    },
    {
        "Nº LOJA": "B016",
        "REFERENCIA LOJA": "BALNEÁRIO CAMBORIÚ",
        "UF": "SC",
        "ESTADO": "SANTA CATARINA",
        "NOME": "Tábatta Sarah da Silva",
        "EMAIL": "Tabatta.silva@vonnycosmeticos.com.br",
        "TELEFONE": "11 97361-3084",
        "SETOR": "LOJA",
        "CARGO": "gerente operacional",
        "N DO GRUPO": "2",
        "GRUPO": "LÍDERES LOJA",
        "SAP ID": "TSARAH6449",
        "Gmail padrão Recebimento": "-",
        "Coluna 1": "Atualizado em 22/09/2026 Whats Gerente Claudineia Mendes",
    },
    {
        "Nº LOJA": "B016",
        "REFERENCIA LOJA": "BALNEÁRIO CAMBORIÚ",
        "UF": "SC",
        "ESTADO": "SANTA CATARINA",
        "NOME": "Ianara Inácio Bossi",
        "EMAIL": "Ianara.bossi@vonnycosmeticos.com.br",
        "TELEFONE": "11 97361-3084",
        "SETOR": "LOJA",
        "CARGO": "gerente operacional",
        "N DO GRUPO": "2",
        "GRUPO": "LÍDERES LOJA",
        "SAP ID": "-",
        "Gmail padrão Recebimento": "-",
        "Coluna 1": "Atualizado em 22/09/2026 Whats Gerente Claudineia Mendes",
    },
    {
        "Nº LOJA": "B016",
        "REFERENCIA LOJA": "BALNEÁRIO CAMBORIÚ",
        "UF": "SC",
        "ESTADO": "SANTA CATARINA",
        "NOME": "Maria Jaqueline Felix de Lemos",
        "EMAIL": "adm.camboriu@vonnycosmeticos.com.br",
        "TELEFONE": "11 97496-6333",
        "SETOR": "ADMINISTRATIVO",
        "CARGO": "1027 - Operador de Loja (exp.) inicio em 19/08/2024",
        "N DO GRUPO": "3",
        "GRUPO": "ADMINISTRATIVO",
        "SAP ID": "TROCABCA6422",
        "Gmail padrão Recebimento": "-",
        "Coluna 1": "Atualizado em 22/09/2026 Whats Gerente Claudineia Mendes",
    },
    {
        "Nº LOJA": "B016",
        "REFERENCIA LOJA": "BALNEÁRIO CAMBORIÚ",
        "UF": "SC",
        "ESTADO": "SANTA CATARINA",
        "NOME": "Jaqueline de Jesus",
        "EMAIL": "recebmp16@vonnycosmeticos.com.br",
        "TELEFONE": "11 97540-5147",
        "SETOR": "RECEBIMENTO / TROCAS",
        "CARGO": "1027 - Operador de Loja (exp.) inicio em 06/06/2024",
        "N DO GRUPO": "7",
        "GRUPO": "RECEBIMENTO / TROCAS",
        "SAP ID": "-",
        "Gmail padrão Recebimento": "recebmpb016@gmail.com",
        "Coluna 1": "Atualizado em 22/09/2026 Whats Gerente Claudineia Mendes",
    },
    {
        "Nº LOJA": "B017",
        "REFERENCIA LOJA": "SÃO LEOPOLDO",
        "UF": "RS",
        "ESTADO": "RIO GRANDE DO SUL",
        "NOME": "Thatiane Ferreira",
        "EMAIL": "thatiane.ferreira@vonnycosmeticos.com.br",
        "TELEFONE": "11 93364-8424",
        "SETOR": "LOJA",
        "CARGO": "-",
        "N DO GRUPO": "1",
        "GRUPO": "GERENTES LOJA",
        "SAP ID": "-",
        "Gmail padrão Recebimento": "-",
        "Coluna 1": "Atualizado em 22/09/2026 Whats Gerente Thatiane Ferreira",
    },
    {
        "Nº LOJA": "B017",
        "REFERENCIA LOJA": "SÃO LEOPOLDO",
        "UF": "RS",
        "ESTADO": "RIO GRANDE DO SUL",
        "NOME": "Mariana Ribeiro de Oliveira",
        "EMAIL": "mariana.ribeiro@vonnycosmeticos.com.br",
        "TELEFONE": "11 99739-2155",
        "SETOR": "LOJA",
        "CARGO": "-",
        "N DO GRUPO": "2",
        "GRUPO": "LÍDERES LOJA",
        "SAP ID": "-",
        "Gmail padrão Recebimento": "-",
        "Coluna 1": "Atualizado em 22/09/2026 Whats Gerente Thatiane Ferreira",
    },
    {
        "Nº LOJA": "B017",
        "REFERENCIA LOJA": "SÃO LEOPOLDO",
        "UF": "RS",
        "ESTADO": "RIO GRANDE DO SUL",
        "NOME": "Camila Peixoto",
        "EMAIL": "adm.saoleopoldo@vonnycosmeticos.com.br",
        "TELEFONE": "11 93498-1451",
        "SETOR": "ADMINISTRATIVO",
        "CARGO": "-",
        "N DO GRUPO": "3",
        "GRUPO": "ADMINISTRATIVO",
        "SAP ID": "-",
        "Gmail padrão Recebimento": "-",
        "Coluna 1": "Atualizado em 22/09/2026 Whats Gerente Thatiane Ferreira",
    },
    {
        "Nº LOJA": "B017",
        "REFERENCIA LOJA": "SÃO LEOPOLDO",
        "UF": "RS",
        "ESTADO": "RIO GRANDE DO SUL",
        "NOME": "Giovana Teixeira dos Santos",
        "EMAIL": "recebmp17@vonnycosmeticos.com.br",
        "TELEFONE": "11 99894-7834",
        "SETOR": "TROCAS",
        "CARGO": "-",
        "N DO GRUPO": "4",
        "GRUPO": "TROCAS",
        "SAP ID": "-",
        "Gmail padrão Recebimento": "-",
        "Coluna 1": "Atualizado em 22/09/2026 Whats Gerente Thatiane Ferreira",
    },
    {
        "Nº LOJA": "B017",
        "REFERENCIA LOJA": "SÃO LEOPOLDO",
        "UF": "RS",
        "ESTADO": "RIO GRANDE DO SUL",
        "NOME": "Lucas Adriano Savallisch Rodrigues",
        "EMAIL": "recebmp17@vonnycosmeticos.com.br",
        "TELEFONE": "11 99894-7834",
        "SETOR": "RECEBIMENTO",
        "CARGO": "-",
        "N DO GRUPO": "5",
        "GRUPO": "RECEBIMENTO",
        "SAP ID": "-",
        "Gmail padrão Recebimento": "recebmpb017@gmail.com",
        "Coluna 1": "Atualizado em 22/09/2026 Whats Gerente Thatiane Ferreira",
    },
    {
        "Nº LOJA": "B018",
        "REFERENCIA LOJA": "SÃO BERNARDO DO CAMPO 2",
        "UF": "SP",
        "ESTADO": "SÃO PAULO",
        "NOME": "Katiane da Silva S Maximiano",
        "EMAIL": "katiane.silva@vonnycosmeticos.com.br",
        "TELEFONE": "11 91829-6671",
        "SETOR": "LOJA",
        "CARGO": "105 - Gerente",
        "N DO GRUPO": "1",
        "GRUPO": "GERENTES LOJA",
        "SAP ID": "KSOUZA5793",
        "Gmail padrão Recebimento": "-",
        "Coluna 1": "Atualizado em 22/09/2026 Whats Gerente Fabiana Bertassi",
    },
    {
        "Nº LOJA": "B018",
        "REFERENCIA LOJA": "SÃO BERNARDO DO CAMPO 2",
        "UF": "SP",
        "ESTADO": "SÃO PAULO",
        "NOME": "Hector Ramon Oliveros Vicent",
        "EMAIL": "Hector.ramon@vonnycosmeticos.com.br",
        "TELEFONE": "-",
        "SETOR": "LOJA",
        "CARGO": "-",
        "N DO GRUPO": "2",
        "GRUPO": "LÍDERES LOJA",
        "SAP ID": "6912",
        "Gmail padrão Recebimento": "-",
        "Coluna 1": "Atualizado em 22/09/2026 Whats Gerente Fabiana Bertassi",
    },
    {
        "Nº LOJA": "B018",
        "REFERENCIA LOJA": "SÃO BERNARDO DO CAMPO 2",
        "UF": "SP",
        "ESTADO": "SÃO PAULO",
        "NOME": "Jéssica alexandrina Barros",
        "EMAIL": "Jessica.barros@vonnycosmeticos.com.br",
        "TELEFONE": "11 91829-9055",
        "SETOR": "LOJA",
        "CARGO": "-",
        "N DO GRUPO": "2",
        "GRUPO": "LÍDERES LOJA",
        "SAP ID": "-",
        "Gmail padrão Recebimento": "-",
        "Coluna 1": "Atualizado em 22/09/2026 Whats Gerente Fabiana Bertassi",
    },
    {
        "Nº LOJA": "B018",
        "REFERENCIA LOJA": "SÃO BERNARDO DO CAMPO 2",
        "UF": "SP",
        "ESTADO": "SÃO PAULO",
        "NOME": "Gisele Aparecida Correa",
        "EMAIL": "gisele.correa@vonnycosmeticos.com.br",
        "TELEFONE": "11 91829-6614",
        "SETOR": "LOJA",
        "CARGO": "1045 - Lider de Setor",
        "N DO GRUPO": "2",
        "GRUPO": "LÍDERES LOJA",
        "SAP ID": "GAPARECIDA5517",
        "Gmail padrão Recebimento": "-",
        "Coluna 1": "Atualizado em 22/09/2026 Whats Gerente Fabiana Bertassi",
    },
    {
        "Nº LOJA": "B018",
        "REFERENCIA LOJA": "SÃO BERNARDO DO CAMPO 2",
        "UF": "SP",
        "ESTADO": "SÃO PAULO",
        "NOME": "Givanilda Maria da Silva",
        "EMAIL": "adm.sbc02@vonnycosmeticos.com.br",
        "TELEFONE": "11 91829-7435",
        "SETOR": "ADMINISTRATIVO",
        "CARGO": "1045 - Lider de Setor",
        "N DO GRUPO": "3",
        "GRUPO": "ADMINISTRATIVO",
        "SAP ID": "4872",
        "Gmail padrão Recebimento": "-",
        "Coluna 1": "Atualizado em 22/09/2026 Whats Gerente Fabiana Bertassi",
    },
    {
        "Nº LOJA": "B018",
        "REFERENCIA LOJA": "SÃO BERNARDO DO CAMPO 2",
        "UF": "SP",
        "ESTADO": "SÃO PAULO",
        "NOME": "Lucas Henrique M da S Felix",
        "EMAIL": "troca.sbc2@vonnycosmeticos.com.br",
        "TELEFONE": "11 97595-9102",
        "SETOR": "RECEBIMENTO / TROCAS",
        "CARGO": "52 - Auxiliar de Recebimento",
        "N DO GRUPO": "7",
        "GRUPO": "RECEBIMENTO / TROCAS",
        "SAP ID": "LFELIX6189",
        "Gmail padrão Recebimento": "recebmpb018@gmail.com",
        "Coluna 1": "Atualizado em 22/09/2026 Whats Gerente Fabiana Bertassi",
    },
    {
        "Nº LOJA": "B019",
        "REFERENCIA LOJA": "GOIANIA",
        "UF": "GO",
        "ESTADO": "GOIAS",
        "NOME": "Lanny Andryelly de Jesus Costa",
        "EMAIL": "lanny.andryelly@vonnycosmeticos.com.br",
        "TELEFONE": "11 91829-7233",
        "SETOR": "LOJA",
        "CARGO": "-",
        "N DO GRUPO": "1",
        "GRUPO": "GERENTES LOJA",
        "SAP ID": "-",
        "Gmail padrão Recebimento": "-",
        "Coluna 1": "Atualizado em 22/09/2026 Whats Gerente Lanny Andryelly",
    },
    {
        "Nº LOJA": "B019",
        "REFERENCIA LOJA": "GOIANIA",
        "UF": "GO",
        "ESTADO": "GOIAS",
        "NOME": "Bárbara Cristina Oliveira Gomes",
        "EMAIL": "barbara.oliveira@vonnycosmeticos.com.br",
        "TELEFONE": "11 91829-8934",
        "SETOR": "LOJA",
        "CARGO": "-",
        "N DO GRUPO": "2",
        "GRUPO": "LÍDERES LOJA",
        "SAP ID": "-",
        "Gmail padrão Recebimento": "-",
        "Coluna 1": "Atualizado em 22/09/2026 Whats Gerente Lanny Andryelly",
    },
    {
        "Nº LOJA": "B019",
        "REFERENCIA LOJA": "GOIANIA",
        "UF": "GO",
        "ESTADO": "GOIAS",
        "NOME": "-",
        "EMAIL": "adm.goiania@vonnycosmeticos.com.br",
        "TELEFONE": "11 91829-8653",
        "SETOR": "ADMINISTRATIVO",
        "CARGO": "-",
        "N DO GRUPO": "3",
        "GRUPO": "ADMINISTRATIVO",
        "SAP ID": "-",
        "Gmail padrão Recebimento": "-",
        "Coluna 1": "Atualizado em 22/09/2026 Whats Gerente Lanny Andryelly",
    },
    {
        "Nº LOJA": "B019",
        "REFERENCIA LOJA": "GOIANIA",
        "UF": "GO",
        "ESTADO": "GOIAS",
        "NOME": "Liliane Miranda da Silva",
        "EMAIL": "recebmp19@vonnycosmeticos.com.br",
        "TELEFONE": "11 91829-7290",
        "SETOR": "RECEBIMENTO / TROCAS",
        "CARGO": "-",
        "N DO GRUPO": "7",
        "GRUPO": "RECEBIMENTO / TROCAS",
        "SAP ID": "-",
        "Gmail padrão Recebimento": "recebmpb019@gmail.com",
        "Coluna 1": "Atualizado em 22/09/2026 Whats Gerente Lanny Andryelly",
    },
    {
        "Nº LOJA": "B020",
        "REFERENCIA LOJA": "PORTO ALEGRE 2",
        "UF": "RS",
        "ESTADO": "RIO GRANDE DO SUL",
        "NOME": "Suzana Vasconcellos da Silveira",
        "EMAIL": "Suzana.silveira@vonnycosmeticos.com.br",
        "TELEFONE": "11 91829-8542",
        "SETOR": "LOJA",
        "CARGO": "-",
        "N DO GRUPO": "1",
        "GRUPO": "GERENTES LOJA",
        "SAP ID": "-",
        "Gmail padrão Recebimento": "-",
        "Coluna 1": "Atualizado em 22/09/2026 Whats Gerente Suzana Vasconcellos",
    },
    {
        "Nº LOJA": "B020",
        "REFERENCIA LOJA": "PORTO ALEGRE 2",
        "UF": "RS",
        "ESTADO": "RIO GRANDE DO SUL",
        "NOME": "Celena Monteiro",
        "EMAIL": "celena.monteiro@vonnycosmeticos.com.br",
        "TELEFONE": "11 91829-8348",
        "SETOR": "LOJA",
        "CARGO": "-",
        "N DO GRUPO": "2",
        "GRUPO": "LÍDERES LOJA",
        "SAP ID": "-",
        "Gmail padrão Recebimento": "-",
        "Coluna 1": "Atualizado em 22/09/2026 Whats Gerente Suzana Vasconcellos",
    },
    {
        "Nº LOJA": "B020",
        "REFERENCIA LOJA": "PORTO ALEGRE 2",
        "UF": "RS",
        "ESTADO": "RIO GRANDE DO SUL",
        "NOME": "Alessandra Fernandes",
        "EMAIL": "adm.portoalegre02@vonnycosmeticos.com.br",
        "TELEFONE": "11 91829-8473",
        "SETOR": "ADMINISTRATIVO",
        "CARGO": "-",
        "N DO GRUPO": "3",
        "GRUPO": "ADMINISTRATIVO",
        "SAP ID": "-",
        "Gmail padrão Recebimento": "-",
        "Coluna 1": "Atualizado em 22/09/2026 Whats Gerente Suzana Vasconcellos",
    },
    {
        "Nº LOJA": "B020",
        "REFERENCIA LOJA": "PORTO ALEGRE 2",
        "UF": "RS",
        "ESTADO": "RIO GRANDE DO SUL",
        "NOME": "Jéssica Borges Vaz",
        "EMAIL": "recebmp20@vonnycosmeticos.com.br",
        "TELEFONE": "11 91829-8580",
        "SETOR": "RECEBIMENTO / TROCAS",
        "CARGO": "-",
        "N DO GRUPO": "7",
        "GRUPO": "RECEBIMENTO / TROCAS",
        "SAP ID": "-",
        "Gmail padrão Recebimento": "recebmpb020@gmail.com",
        "Coluna 1": "Atualizado em 22/09/2026 Whats Gerente Suzana Vasconcellos",
    },
    {
        "Nº LOJA": "B021",
        "REFERENCIA LOJA": "PORTO ALEGRE 3",
        "UF": "RS",
        "ESTADO": "RIO GRANDE DO SUL",
        "NOME": "Luciana Vasconcellos da Silveira",
        "EMAIL": "luciana.vasconcelos@vonnycosmeticos.com.br",
        "TELEFONE": "11 91829-7390",
        "SETOR": "LOJA",
        "CARGO": "-",
        "N DO GRUPO": "1",
        "GRUPO": "GERENTES LOJA",
        "SAP ID": "-",
        "Gmail padrão Recebimento": "-",
        "Coluna 1": "Atualizado em 22/09/2026 Whats Gerente Luciana Vasconcellos",
    },
    {
        "Nº LOJA": "B021",
        "REFERENCIA LOJA": "PORTO ALEGRE 3",
        "UF": "RS",
        "ESTADO": "RIO GRANDE DO SUL",
        "NOME": "Tamires Ferreira de Oliveira",
        "EMAIL": "tamires.oliveira@vonnycosmeticos.com.br",
        "TELEFONE": "1191829-6455",
        "SETOR": "LOJA",
        "CARGO": "1045 - Lider de Setor",
        "N DO GRUPO": "2",
        "GRUPO": "LÍDERES LOJA",
        "SAP ID": "-",
        "Gmail padrão Recebimento": "-",
        "Coluna 1": "Atualizado em 22/09/2026 Whats Gerente Luciana Vasconcellos",
    },
    {
        "Nº LOJA": "B021",
        "REFERENCIA LOJA": "PORTO ALEGRE 3",
        "UF": "RS",
        "ESTADO": "RIO GRANDE DO SUL",
        "NOME": "Nataly Fabiane da S Rodrígues",
        "EMAIL": "adm.portoalegre03@vonnycosmeticos.com.br",
        "TELEFONE": "11 91829-7088",
        "SETOR": "ADMINISTRATIVO",
        "CARGO": "-",
        "N DO GRUPO": "3",
        "GRUPO": "ADMINISTRATIVO",
        "SAP ID": "-",
        "Gmail padrão Recebimento": "-",
        "Coluna 1": "Atualizado em 22/09/2026 Whats Gerente Luciana Vasconcellos",
    },
    {
        "Nº LOJA": "B021",
        "REFERENCIA LOJA": "PORTO ALEGRE 3",
        "UF": "RS",
        "ESTADO": "RIO GRANDE DO SUL",
        "NOME": "Gustavo Pinto do Nascimento",
        "EMAIL": "recebmp21@vonnycosmeticos.com.br",
        "TELEFONE": "11 91829-7020",
        "SETOR": "RECEBIMENTO / TROCAS",
        "CARGO": "-",
        "N DO GRUPO": "7",
        "GRUPO": "RECEBIMENTO / TROCAS",
        "SAP ID": "-",
        "Gmail padrão Recebimento": "recebmpb021@gmail.com",
        "Coluna 1": "Atualizado em 22/09/2026 Whats Gerente Luciana Vasconcellos",
    },
    {
        "Nº LOJA": "B022",
        "REFERENCIA LOJA": "CURITIBA",
        "UF": "PR",
        "ESTADO": "PARANÁ",
        "NOME": "Suzana Pires",
        "EMAIL": "suzana.pires@vonnycosmeticos.com.br",
        "TELEFONE": "11 91829-6571",
        "SETOR": "LOJA",
        "CARGO": "-",
        "N DO GRUPO": "1",
        "GRUPO": "GERENTES LOJA",
        "SAP ID": "-",
        "Gmail padrão Recebimento": "-",
        "Coluna 1": "Atualizado em 21/09/2026 Whats Regional Luciana",
    },
    {
        "Nº LOJA": "B022",
        "REFERENCIA LOJA": "CURITIBA",
        "UF": "PR",
        "ESTADO": "PARANÁ",
        "NOME": "Thacilla Portela da Paixão",
        "EMAIL": "thacilla.paixao@vonnycosmeticos.com.br",
        "TELEFONE": "11 91829-6502",
        "SETOR": "LOJA",
        "CARGO": "-",
        "N DO GRUPO": "2",
        "GRUPO": "LÍDERES LOJA",
        "SAP ID": "-",
        "Gmail padrão Recebimento": "-",
        "Coluna 1": "Atualizado em 29/09/2026 Whats Gerente Suzana Pires",
    },
    {
        "Nº LOJA": "B022",
        "REFERENCIA LOJA": "CURITIBA",
        "UF": "PR",
        "ESTADO": "PARANÁ",
        "NOME": "Ketlyn Caroline de Melo",
        "EMAIL": "adm.curitiba@vonnycosmeticos.com.br",
        "TELEFONE": "11 91829-7891",
        "SETOR": "ADMINISTRATIVO",
        "CARGO": "-",
        "N DO GRUPO": "3",
        "GRUPO": "ADMINISTRATIVO",
        "SAP ID": "-",
        "Gmail padrão Recebimento": "-",
        "Coluna 1": "Atualizado em 29/09/2026 Whats Gerente Suzana Pires",
    },
    {
        "Nº LOJA": "B022",
        "REFERENCIA LOJA": "CURITIBA",
        "UF": "PR",
        "ESTADO": "PARANÁ",
        "NOME": "Erik Soares Rodrigues",
        "EMAIL": "recebmp22@vonnycosmeticos.com.br",
        "TELEFONE": "11 91829-9368",
        "SETOR": "RECEBIMENTO / TROCAS",
        "CARGO": "-",
        "N DO GRUPO": "7",
        "GRUPO": "RECEBIMENTO / TROCAS",
        "SAP ID": "-",
        "Gmail padrão Recebimento": "-",
        "Coluna 1": "Atualizado em 29/09/2026 Whats Gerente Suzana Pires",
    },
    {
        "Nº LOJA": "B023",
        "REFERENCIA LOJA": "PORTO ALEGRE 4",
        "UF": "RS",
        "ESTADO": "RIO GRANDE DO SUL",
        "NOME": "Gisele Trampusch",
        "EMAIL": "gisele.trampusch@vonnycosmeticos.com.br",
        "TELEFONE": "11 99930-4924",
        "SETOR": "LOJA",
        "CARGO": "-",
        "N DO GRUPO": "1",
        "GRUPO": "GERENTES LOJA",
        "SAP ID": "-",
        "Gmail padrão Recebimento": "-",
        "Coluna 1": "Atualizado em 22/09/2026 Whats Gerente Gisele Trampusch",
    },
    {
        "Nº LOJA": "B023",
        "REFERENCIA LOJA": "PORTO ALEGRE 4",
        "UF": "RS",
        "ESTADO": "RIO GRANDE DO SUL",
        "NOME": "Debora Reis mazzili",
        "EMAIL": "debora.mazzilli@vonnycosmeticos.com.br",
        "TELEFONE": "11 99734-5609",
        "SETOR": "LOJA",
        "CARGO": "-",
        "N DO GRUPO": "2",
        "GRUPO": "LÍDERES LOJA",
        "SAP ID": "-",
        "Gmail padrão Recebimento": "-",
        "Coluna 1": "Atualizado em 22/09/2026 Whats Gerente Gisele Trampusch",
    },
    {
        "Nº LOJA": "B023",
        "REFERENCIA LOJA": "PORTO ALEGRE 4",
        "UF": "RS",
        "ESTADO": "RIO GRANDE DO SUL",
        "NOME": "Sue Ellen Mendes",
        "EMAIL": "adm.portoalegre04@vonnycosmeticos.com.br",
        "TELEFONE": "11 97259-4268",
        "SETOR": "ADMINISTRATIVO",
        "CARGO": "-",
        "N DO GRUPO": "3",
        "GRUPO": "ADMINISTRATIVO",
        "SAP ID": "-",
        "Gmail padrão Recebimento": "-",
        "Coluna 1": "Atualizado em 22/09/2026 Whats Gerente Gisele Trampusch",
    },
    {
        "Nº LOJA": "B023",
        "REFERENCIA LOJA": "PORTO ALEGRE 4",
        "UF": "RS",
        "ESTADO": "RIO GRANDE DO SUL",
        "NOME": "Juliana Correa Silveira",
        "EMAIL": "recebmp23@vonnycosmeticos.com.br",
        "TELEFONE": "11 99720-3539",
        "SETOR": "RECEBIMENTO / TROCAS",
        "CARGO": "-",
        "N DO GRUPO": "7",
        "GRUPO": "RECEBIMENTO / TROCAS",
        "SAP ID": "-",
        "Gmail padrão Recebimento": "recebmpb023@gmail.com",
        "Coluna 1": "Atualizado em 22/09/2026 Whats Gerente Gisele Trampusch",
    },
    {
        "Nº LOJA": "B024",
        "REFERENCIA LOJA": "FORTALEZA",
        "UF": "CE",
        "ESTADO": "CEARÁ",
        "NOME": "Raquel Ketlen Lopes A. de Oliveira",
        "EMAIL": "raquel.lopes@vonnycosmeticos.com.br",
        "TELEFONE": "11 99771-2651",
        "SETOR": "LOJA",
        "CARGO": "105 - Gerente",
        "N DO GRUPO": "1",
        "GRUPO": "GERENTES LOJA",
        "SAP ID": "RLOPES",
        "Gmail padrão Recebimento": "-",
        "Coluna 1": "Atualizado em 22/09/2026 Whats Gerente Raquel Ketlen",
    },
    {
        "Nº LOJA": "B024",
        "REFERENCIA LOJA": "FORTALEZA",
        "UF": "CE",
        "ESTADO": "CEARÁ",
        "NOME": "Tays Gomes da Silva",
        "EMAIL": "Tays.silva@vonnycosmeticos.com.br",
        "TELEFONE": "11 99592-8761",
        "SETOR": "LOJA",
        "CARGO": "-",
        "N DO GRUPO": "2",
        "GRUPO": "LÍDERES LOJA",
        "SAP ID": "-",
        "Gmail padrão Recebimento": "-",
        "Coluna 1": "Atualizado em 22/09/2026 Whats Gerente Raquel Ketlen",
    },
    {
        "Nº LOJA": "B024",
        "REFERENCIA LOJA": "FORTALEZA",
        "UF": "CE",
        "ESTADO": "CEARÁ",
        "NOME": "Ana Paula monte",
        "EMAIL": "Ana.monte@vonnycosmeticos.com.br",
        "TELEFONE": "11 99592-8761",
        "SETOR": "LOJA",
        "CARGO": "-",
        "N DO GRUPO": "2",
        "GRUPO": "LÍDERES LOJA",
        "SAP ID": "-",
        "Gmail padrão Recebimento": "-",
        "Coluna 1": "Atualizado em 22/09/2026 Whats Gerente Raquel Ketlen",
    },
    {
        "Nº LOJA": "B024",
        "REFERENCIA LOJA": "FORTALEZA",
        "UF": "CE",
        "ESTADO": "CEARÁ",
        "NOME": "Maria Ingrid Pereira da Silva",
        "EMAIL": "adm.fortaleza@vonnycosmeticos.com.br",
        "TELEFONE": "11 99669-8980",
        "SETOR": "ADMINISTRATIVO",
        "CARGO": "-",
        "N DO GRUPO": "3",
        "GRUPO": "ADMINISTRATIVO",
        "SAP ID": "-",
        "Gmail padrão Recebimento": "-",
        "Coluna 1": "Atualizado em 22/09/2026 Whats Gerente Raquel Ketlen",
    },
    {
        "Nº LOJA": "B024",
        "REFERENCIA LOJA": "FORTALEZA",
        "UF": "CE",
        "ESTADO": "CEARÁ",
        "NOME": "João Victor",
        "EMAIL": "recebmp24@vonnycosmeticos.com.br",
        "TELEFONE": "11 99612-7212",
        "SETOR": "RECEBIMENTO / TROCAS",
        "CARGO": "Em treinamento",
        "N DO GRUPO": "7",
        "GRUPO": "RECEBIMENTO / TROCAS",
        "SAP ID": "-",
        "Gmail padrão Recebimento": "-",
        "Coluna 1": "Atualizado em 22/09/2026 Whats Gerente Raquel Ketlen",
    },
    {
        "Nº LOJA": "B024",
        "REFERENCIA LOJA": "FORTALEZA",
        "UF": "CE",
        "ESTADO": "CEARÁ",
        "NOME": "Nathaniel Silva Quinto",
        "EMAIL": "recebmp24@vonnycosmeticos.com.br",
        "TELEFONE": "11 99612-7212",
        "SETOR": "RECEBIMENTO / TROCAS",
        "CARGO": "-",
        "N DO GRUPO": "7",
        "GRUPO": "RECEBIMENTO / TROCAS",
        "SAP ID": "-",
        "Gmail padrão Recebimento": "recebmpb024@gmail.com",
        "Coluna 1": "Atualizado em 22/09/2026 Whats Gerente Raquel Ketlen",
    },
    {
        "Nº LOJA": "B025",
        "REFERENCIA LOJA": "SALVADOR",
        "UF": "BA",
        "ESTADO": "BAHIA",
        "NOME": "Rosania Lacerda chagas",
        "EMAIL": "Rosania.chagas@vonnycosmeticos.com.br",
        "TELEFONE": "11 9 9542-2142",
        "SETOR": "LOJA",
        "CARGO": "-",
        "N DO GRUPO": "1",
        "GRUPO": "GERENTES LOJA",
        "SAP ID": "-",
        "Gmail padrão Recebimento": "-",
        "Coluna 1": "Atualizado em 22/09/2026 Whats Gerente RosaniaSALVADOR SHOPPING",
    },
    {
        "Nº LOJA": "B025",
        "REFERENCIA LOJA": "SALVADOR",
        "UF": "BA",
        "ESTADO": "BAHIA",
        "NOME": "Pâmela Julia Barreto dos Santos",
        "EMAIL": "pamela.santos@vonnycosmeticos.com.br",
        "TELEFONE": "11 97402-2727",
        "SETOR": "LOJA",
        "CARGO": "-",
        "N DO GRUPO": "2",
        "GRUPO": "LÍDERES LOJA",
        "SAP ID": "-",
        "Gmail padrão Recebimento": "-",
        "Coluna 1": "Atualizado em 22/09/2026 Whats Gerente RosaniaSALVADOR SHOPPING",
    },
    {
        "Nº LOJA": "B025",
        "REFERENCIA LOJA": "SALVADOR",
        "UF": "BA",
        "ESTADO": "BAHIA",
        "NOME": "Priscila Portela",
        "EMAIL": "Priscila.portela@vonnycosmeticos.com.br",
        "TELEFONE": "11 97402-2727",
        "SETOR": "LOJA",
        "CARGO": "Em Treinamento",
        "N DO GRUPO": "2",
        "GRUPO": "LÍDERES LOJA",
        "SAP ID": "-",
        "Gmail padrão Recebimento": "-",
        "Coluna 1": "Atualizado em 22/09/2026 Whats Gerente RosaniaSALVADOR SHOPPING",
    },
    {
        "Nº LOJA": "B025",
        "REFERENCIA LOJA": "SALVADOR",
        "UF": "BA",
        "ESTADO": "BAHIA",
        "NOME": "Debora Santos Sá Barreto",
        "EMAIL": "adm.salvador@vonnycosmeticos.com.br",
        "TELEFONE": "11 99518-8071",
        "SETOR": "ADMINISTRATIVO",
        "CARGO": "-",
        "N DO GRUPO": "3",
        "GRUPO": "ADMINISTRATIVO",
        "SAP ID": "-",
        "Gmail padrão Recebimento": "-",
        "Coluna 1": "Atualizado em 22/09/2026 Whats Gerente RosaniaSALVADOR SHOPPING",
    },
    {
        "Nº LOJA": "B025",
        "REFERENCIA LOJA": "SALVADOR",
        "UF": "BA",
        "ESTADO": "BAHIA",
        "NOME": "Joyce dos Santos Pereira",
        "EMAIL": "recebmp25@vonnycosmetico.com.br",
        "TELEFONE": "11 97534-5299",
        "SETOR": "RECEBIMENTO / TROCAS",
        "CARGO": "-",
        "N DO GRUPO": "7",
        "GRUPO": "RECEBIMENTO / TROCAS",
        "SAP ID": "-",
        "Gmail padrão Recebimento": "recebmpb025@gmail.com",
        "Coluna 1": "Atualizado em 22/09/2026 Whats Gerente RosaniaSALVADOR SHOPPING",
    },
    {
        "Nº LOJA": "B026",
        "REFERENCIA LOJA": "SALVADOR 2",
        "UF": "BA",
        "ESTADO": "BAHIA",
        "NOME": "Claudinea Lopes dos Santos (Nea)",
        "EMAIL": "claudinea.santos@vonnycosmeticos.com.br",
        "TELEFONE": "11 99542-2142",
        "SETOR": "LOJA",
        "CARGO": "-",
        "N DO GRUPO": "1",
        "GRUPO": "GERENTES LOJA",
        "SAP ID": "-",
        "Gmail padrão Recebimento": "-",
        "Coluna 1": "Atualizado em 22/09/2026 Whats Gerente NeaSALVADOR NORTE",
    },
    {
        "Nº LOJA": "B026",
        "REFERENCIA LOJA": "SALVADOR 2",
        "UF": "BA",
        "ESTADO": "BAHIA",
        "NOME": "Mirela Santos",
        "EMAIL": "Mirela.santos@vonnycosmeticos.com.br",
        "TELEFONE": "11 94171-5565",
        "SETOR": "LOJA",
        "CARGO": "-",
        "N DO GRUPO": "2",
        "GRUPO": "LÍDERES LOJA",
        "SAP ID": "-",
        "Gmail padrão Recebimento": "-",
        "Coluna 1": "Atualizado em 22/09/2026 Whats Gerente NeaSALVADOR NORTE",
    },
    {
        "Nº LOJA": "B026",
        "REFERENCIA LOJA": "SALVADOR 2",
        "UF": "BA",
        "ESTADO": "BAHIA",
        "NOME": "Natália do Rosário Bispo dos Santos",
        "EMAIL": "adm.salvador2@vonnycosmeticos.com.br",
        "TELEFONE": "11 97134-3181",
        "SETOR": "ADMINISTRATIVO",
        "CARGO": "-",
        "N DO GRUPO": "3",
        "GRUPO": "ADMINISTRATIVO",
        "SAP ID": "-",
        "Gmail padrão Recebimento": "-",
        "Coluna 1": "Atualizado em 22/09/2026 Whats Gerente NeaSALVADOR NORTE",
    },
    {
        "Nº LOJA": "B026",
        "REFERENCIA LOJA": "SALVADOR 2",
        "UF": "BA",
        "ESTADO": "BAHIA",
        "NOME": "Daniela Souza Barbosa Araujo",
        "EMAIL": "recebmp26@vonnycosmeticos.com.br",
        "TELEFONE": "11 96492-5212",
        "SETOR": "RECEBIMENTO / TROCAS",
        "CARGO": "-",
        "N DO GRUPO": "7",
        "GRUPO": "RECEBIMENTO / TROCAS",
        "SAP ID": "8436",
        "Gmail padrão Recebimento": "recebmpb026@gmail.com",
        "Coluna 1": "Atualizado em 22/09/2026 Whats Gerente NeaSALVADOR NORTE",
    },
    {
        "Nº LOJA": "B027",
        "REFERENCIA LOJA": "OLINDA",
        "UF": "PE",
        "ESTADO": "PERNAMBUCO",
        "NOME": "Luana Moreira da costa",
        "EMAIL": "luana.costa@vonnycosmeticos.com.br",
        "TELEFONE": "11 96197-6327",
        "SETOR": "LOJA",
        "CARGO": "-",
        "N DO GRUPO": "1",
        "GRUPO": "GERENTES LOJA",
        "SAP ID": "-",
        "Gmail padrão Recebimento": "-",
        "Coluna 1": "Atualizado em 22/09/2026 Whats Gerente Luana Moreiram",
    },
    {
        "Nº LOJA": "B027",
        "REFERENCIA LOJA": "OLINDA",
        "UF": "PE",
        "ESTADO": "PERNAMBUCO",
        "NOME": "Acácia Lima",
        "EMAIL": "acacia.lima@vonnycosmeticos.com.br",
        "TELEFONE": "11 97457-0177",
        "SETOR": "LOJA",
        "CARGO": "-",
        "N DO GRUPO": "2",
        "GRUPO": "LÍDERES LOJA",
        "SAP ID": "-",
        "Gmail padrão Recebimento": "-",
        "Coluna 1": "Atualizado em 22/09/2026 Whats Gerente Luana Moreiram",
    },
    {
        "Nº LOJA": "B027",
        "REFERENCIA LOJA": "OLINDA",
        "UF": "PE",
        "ESTADO": "PERNAMBUCO",
        "NOME": "Marilane Oliveira",
        "EMAIL": "marilane.oliveira@vonnycosmeticos.com.br",
        "TELEFONE": "11 97457-0177",
        "SETOR": "LOJA",
        "CARGO": "-",
        "N DO GRUPO": "2",
        "GRUPO": "LÍDERES LOJA",
        "SAP ID": "-",
        "Gmail padrão Recebimento": "-",
        "Coluna 1": "Atualizado em 22/09/2026 Whats Gerente Luana Moreiram",
    },
    {
        "Nº LOJA": "B027",
        "REFERENCIA LOJA": "OLINDA",
        "UF": "PE",
        "ESTADO": "PERNAMBUCO",
        "NOME": "Fabiana Ferreira Da Silva",
        "EMAIL": "adm.olinda@vonnycosmeticos.com.br",
        "TELEFONE": "11 94393-0888",
        "SETOR": "ADMINISTRATIVO",
        "CARGO": "-",
        "N DO GRUPO": "3",
        "GRUPO": "ADMINISTRATIVO",
        "SAP ID": "-",
        "Gmail padrão Recebimento": "-",
        "Coluna 1": "Atualizado em 22/09/2026 Whats Gerente Luana Moreiram",
    },
    {
        "Nº LOJA": "B027",
        "REFERENCIA LOJA": "OLINDA",
        "UF": "PE",
        "ESTADO": "PERNAMBUCO",
        "NOME": "Andresa Maria do Nascimento Sena",
        "EMAIL": "recebmp27@vonnycosmeticos.com.br",
        "TELEFONE": "11 96487-7242",
        "SETOR": "RECEBIMENTO / TROCAS",
        "CARGO": "-",
        "N DO GRUPO": "7",
        "GRUPO": "RECEBIMENTO / TROCAS",
        "SAP ID": "-",
        "Gmail padrão Recebimento": "recebmpb027@gmail.com",
        "Coluna 1": "Atualizado em 22/09/2026 Whats Gerente Luana Moreiram",
    },
    {
        "Nº LOJA": "B028",
        "REFERENCIA LOJA": "MARINGÁ",
        "UF": "PR",
        "ESTADO": "PARANÁ",
        "NOME": "Rosangela Cristiane Botelho Santos",
        "EMAIL": "rosangela.botelho@vonnycosmeticos.com.br",
        "TELEFONE": "11 94320-0405",
        "SETOR": "LOJA",
        "CARGO": "-",
        "N DO GRUPO": "1",
        "GRUPO": "GERENTES LOJA",
        "SAP ID": "-",
        "Gmail padrão Recebimento": "-",
        "Coluna 1": "Atualizado em 22/09/2026 Gerente Rosangela Cristiane",
    },
    {
        "Nº LOJA": "B028",
        "REFERENCIA LOJA": "MARINGÁ",
        "UF": "PR",
        "ESTADO": "PARANÁ",
        "NOME": "Pablo Henrique campelo dos Santos",
        "EMAIL": "Pablo.campelo@vonnycosmeticos.com.br",
        "TELEFONE": "11 97493-1293",
        "SETOR": "LOJA",
        "CARGO": "-",
        "N DO GRUPO": "2",
        "GRUPO": "LÍDERES LOJA",
        "SAP ID": "-",
        "Gmail padrão Recebimento": "-",
        "Coluna 1": "Atualizado em 22/09/2026 Gerente Rosangela Cristiane",
    },
    {
        "Nº LOJA": "B028",
        "REFERENCIA LOJA": "MARINGÁ",
        "UF": "PR",
        "ESTADO": "PARANÁ",
        "NOME": "Tais Baltazar Viana dos Santos",
        "EMAIL": "adm.maringa@vonnycosmeticos.com.br",
        "TELEFONE": "11 94271-3082",
        "SETOR": "ADMINISTRATIVO",
        "CARGO": "-",
        "N DO GRUPO": "3",
        "GRUPO": "ADMINISTRATIVO",
        "SAP ID": "-",
        "Gmail padrão Recebimento": "-",
        "Coluna 1": "Atualizado em 22/09/2026 Gerente Rosangela Cristiane",
    },
    {
        "Nº LOJA": "B028",
        "REFERENCIA LOJA": "MARINGÁ",
        "UF": "PR",
        "ESTADO": "PARANÁ",
        "NOME": "Alessandro Guilherme Oliveira e Silva",
        "EMAIL": "recebmp28@vonnycosmeticos.com.br",
        "TELEFONE": "11 97529-8528",
        "SETOR": "RECEBIMENTO / TROCAS",
        "CARGO": "-",
        "N DO GRUPO": "7",
        "GRUPO": "RECEBIMENTO / TROCAS",
        "SAP ID": "-",
        "Gmail padrão Recebimento": "recebmpb028@gmail.com",
        "Coluna 1": "Atualizado em 22/09/2026 Gerente Rosangela Cristiane",
    },
    {
        "Nº LOJA": "B029",
        "REFERENCIA LOJA": "NATAL",
        "UF": "RN",
        "ESTADO": "RIO GRANDE DO NORTE",
        "NOME": "Elza dos Santos Silva",
        "EMAIL": "elza.silva@vonnycosmeticos.com.br",
        "TELEFONE": "11 94349-7994",
        "SETOR": "LOJA",
        "CARGO": "-",
        "N DO GRUPO": "1",
        "GRUPO": "GERENTES LOJA",
        "SAP ID": "-",
        "Gmail padrão Recebimento": "-",
        "Coluna 1": "Atualizado em 29/09/2026 Whats Gerente Elza",
    },
    {
        "Nº LOJA": "B029",
        "REFERENCIA LOJA": "NATAL",
        "UF": "RN",
        "ESTADO": "RIO GRANDE DO NORTE",
        "NOME": "Elayne Costa Coutinho",
        "EMAIL": "elayne.coutinho@vonnycosmeticos.com.br",
        "TELEFONE": "1194217-8468",
        "SETOR": "LOJA",
        "CARGO": "-",
        "N DO GRUPO": "2",
        "GRUPO": "LÍDERES LOJA",
        "SAP ID": "-",
        "Gmail padrão Recebimento": "-",
        "Coluna 1": "Atualizado em 29/09/2026 Whats Gerente Elza",
    },
    {
        "Nº LOJA": "B029",
        "REFERENCIA LOJA": "NATAL",
        "UF": "RN",
        "ESTADO": "RIO GRANDE DO NORTE",
        "NOME": "Tatiana Medeiros de Lima melo",
        "EMAIL": "adm.natal@vonnycosmeticos.com br",
        "TELEFONE": "11 97292-5678",
        "SETOR": "ADMINISTRATIVO",
        "CARGO": "-",
        "N DO GRUPO": "3",
        "GRUPO": "ADMINISTRATIVO",
        "SAP ID": "-",
        "Gmail padrão Recebimento": "-",
        "Coluna 1": "Atualizado em 29/09/2026 Whats Gerente Elza",
    },
    {
        "Nº LOJA": "B029",
        "REFERENCIA LOJA": "NATAL",
        "UF": "RN",
        "ESTADO": "RIO GRANDE DO NORTE",
        "NOME": "Bruno do Nacimento Silva",
        "EMAIL": "recebmp29@vonnycosmeticos.com.br",
        "TELEFONE": "11 95652-9862",
        "SETOR": "RECEBIMENTO / TROCAS",
        "CARGO": "-",
        "N DO GRUPO": "7",
        "GRUPO": "RECEBIMENTO / TROCAS",
        "SAP ID": "-",
        "Gmail padrão Recebimento": "recebmpb029@gmail.com",
        "Coluna 1": "Atualizado em 29/09/2026 Whats Gerente Elza",
    },
    {
        "Nº LOJA": "B029",
        "REFERENCIA LOJA": "NATAL",
        "UF": "RN",
        "ESTADO": "RIO GRANDE DO NORTE",
        "NOME": "Andrew Raphael Barros",
        "EMAIL": "recebmp29@vonnycosmeticos.com.br",
        "TELEFONE": "11 95652-9862",
        "SETOR": "RECEBIMENTO / TROCAS",
        "CARGO": "-",
        "N DO GRUPO": "7",
        "GRUPO": "RECEBIMENTO / TROCAS",
        "SAP ID": "-",
        "Gmail padrão Recebimento": "-",
        "Coluna 1": "Atualizado em 29/09/2026 Whats Gerente Elza",
    },
    {
        "Nº LOJA": "B030",
        "REFERENCIA LOJA": "PAULISTA",
        "UF": "SP",
        "ESTADO": "SÃO PAULO",
        "NOME": "Joao Pereira Nascimento",
        "EMAIL": "joao.pereira@vonnycosmeticos.com.br",
        "TELEFONE": "11 95786-8606",
        "SETOR": "LOJA",
        "CARGO": "-",
        "N DO GRUPO": "1",
        "GRUPO": "GERENTES LOJA",
        "SAP ID": "-",
        "Gmail padrão Recebimento": "-",
        "Coluna 1": "Atualizado em 24/09/2026 Whats Gerente João Pereira",
    },
    {
        "Nº LOJA": "B030",
        "REFERENCIA LOJA": "PAULISTA",
        "UF": "SP",
        "ESTADO": "SÃO PAULO",
        "NOME": "Tassiana Gomes de Souza",
        "EMAIL": "tassiana.gomes@vonnycosmeticos.com.br",
        "TELEFONE": "11 92576-5436",
        "SETOR": "LOJA",
        "CARGO": "Subgerente",
        "N DO GRUPO": "2",
        "GRUPO": "LÍDERES LOJA",
        "SAP ID": "-",
        "Gmail padrão Recebimento": "-",
        "Coluna 1": "Atualizado em 24/09/2026 Whats Gerente João Pereira",
    },
    {
        "Nº LOJA": "B030",
        "REFERENCIA LOJA": "PAULISTA",
        "UF": "SP",
        "ESTADO": "SÃO PAULO",
        "NOME": "Giovanna Flor Gonçalves",
        "EMAIL": "giovanna.flor@vonnycosmeticos.com.br",
        "TELEFONE": "11 94484-9625",
        "SETOR": "LOJA",
        "CARGO": "-",
        "N DO GRUPO": "2",
        "GRUPO": "LÍDERES LOJA",
        "SAP ID": "-",
        "Gmail padrão Recebimento": "-",
        "Coluna 1": "Atualizado em 24/09/2026 Whats Gerente João Pereira",
    },
    {
        "Nº LOJA": "B030",
        "REFERENCIA LOJA": "PAULISTA",
        "UF": "SP",
        "ESTADO": "SÃO PAULO",
        "NOME": "-",
        "EMAIL": "-",
        "TELEFONE": "11 92533-7862",
        "SETOR": "LOJA",
        "CARGO": "-",
        "N DO GRUPO": "2",
        "GRUPO": "LÍDERES LOJA",
        "SAP ID": "-",
        "Gmail padrão Recebimento": "-",
        "Coluna 1": "Atualizado em 24/09/2026 Whats Gerente João Pereira",
    },
    {
        "Nº LOJA": "B030",
        "REFERENCIA LOJA": "PAULISTA",
        "UF": "SP",
        "ESTADO": "SÃO PAULO",
        "NOME": "Juliana Bezerra de Lima",
        "EMAIL": "adm.paulista@vonnycoamwticos.com.br",
        "TELEFONE": "11 97228-5149",
        "SETOR": "ADMINISTRATIVO",
        "CARGO": "-",
        "N DO GRUPO": "3",
        "GRUPO": "ADMINISTRATIVO",
        "SAP ID": "-",
        "Gmail padrão Recebimento": "-",
        "Coluna 1": "Atualizado em 24/09/2026 Whats Gerente João Pereira",
    },
    {
        "Nº LOJA": "B030",
        "REFERENCIA LOJA": "PAULISTA",
        "UF": "SP",
        "ESTADO": "SÃO PAULO",
        "NOME": "Mateus Nunes Crispim",
        "EMAIL": "recebmp@vonnycosmeticos.com.br",
        "TELEFONE": "11 97580-6383",
        "SETOR": "RECEBIMENTO / TROCAS",
        "CARGO": "-",
        "N DO GRUPO": "7",
        "GRUPO": "RECEBIMENTO / TROCAS",
        "SAP ID": "-",
        "Gmail padrão Recebimento": "recebmpb030@gmail.com",
        "Coluna 1": "Atualizado em 24/09/2026 Whats Gerente João Pereira",
    },
    {
        "Nº LOJA": "B030",
        "REFERENCIA LOJA": "PAULISTA",
        "UF": "SP",
        "ESTADO": "SÃO PAULO",
        "NOME": "Higor Andrade Nogueira",
        "EMAIL": "trocas.tatuape@vonnycosmeticos.com.br",
        "TELEFONE": "11 97580-6383",
        "SETOR": "RECEBIMENTO",
        "CARGO": "-",
        "N DO GRUPO": "5",
        "GRUPO": "RECEBIMENTO",
        "SAP ID": "7644",
        "Gmail padrão Recebimento": "-",
        "Coluna 1": "Atualizado em 24/09/2026 Whats Gerente João Pereira",
    },
    {
        "Nº LOJA": "B031",
        "REFERENCIA LOJA": "FORTALEZA 2",
        "UF": "CE",
        "ESTADO": "CEARÁ",
        "NOME": "Jorgiane Aragão alves",
        "EMAIL": "jorgiane.aragao@vonnycosmeticos.com.br",
        "TELEFONE": "11 973081917",
        "SETOR": "LOJA",
        "CARGO": "-",
        "N DO GRUPO": "1",
        "GRUPO": "GERENTES LOJA",
        "SAP ID": "-",
        "Gmail padrão Recebimento": "-",
        "Coluna 1": "Atualizado em 22/09/2026 Whats Gerente Jorgiane Aragão",
    },
    {
        "Nº LOJA": "B031",
        "REFERENCIA LOJA": "FORTALEZA 2",
        "UF": "CE",
        "ESTADO": "CEARÁ",
        "NOME": "Magnum baçal Torres de Sousa",
        "EMAIL": "magnum.torres@vonnycosmeticos.com.br",
        "TELEFONE": "11 925404382",
        "SETOR": "LOJA",
        "CARGO": "-",
        "N DO GRUPO": "2",
        "GRUPO": "LÍDERES LOJA",
        "SAP ID": "-",
        "Gmail padrão Recebimento": "-",
        "Coluna 1": "Atualizado em 22/09/2026 Whats Gerente Jorgiane Aragão",
    },
    {
        "Nº LOJA": "B031",
        "REFERENCIA LOJA": "FORTALEZA 2",
        "UF": "CE",
        "ESTADO": "CEARÁ",
        "NOME": "Ana Flávia dos Santos Cunha Oliveira",
        "EMAIL": "adm.fortaleza2@vonnycosmeticos.com.br",
        "TELEFONE": "11 926839870",
        "SETOR": "ADMINISTRATIVO",
        "CARGO": "-",
        "N DO GRUPO": "3",
        "GRUPO": "ADMINISTRATIVO",
        "SAP ID": "-",
        "Gmail padrão Recebimento": "-",
        "Coluna 1": "Atualizado em 22/09/2026 Whats Gerente Jorgiane Aragão",
    },
    {
        "Nº LOJA": "B031",
        "REFERENCIA LOJA": "FORTALEZA 2",
        "UF": "CE",
        "ESTADO": "CEARÁ",
        "NOME": "Francisco Romeu Ferreira Davi",
        "EMAIL": "recebmp31@vonnycosmeticos.com.br",
        "TELEFONE": "11 93727-2587",
        "SETOR": "RECEBIMENTO / TROCAS",
        "CARGO": "-",
        "N DO GRUPO": "7",
        "GRUPO": "RECEBIMENTO / TROCAS",
        "SAP ID": "-",
        "Gmail padrão Recebimento": "-",
        "Coluna 1": "Atualizado em 22/09/2026 Whats Gerente Jorgiane Aragão",
    },
    {
        "Nº LOJA": "B032",
        "REFERENCIA LOJA": "SALVADOR 2",
        "UF": "BA",
        "ESTADO": "BAHIA",
        "NOME": "-",
        "EMAIL": "-",
        "TELEFONE": "-",
        "SETOR": "LOJA",
        "CARGO": "-",
        "N DO GRUPO": "1",
        "GRUPO": "GERENTES LOJA",
        "SAP ID": "-",
        "Gmail padrão Recebimento": "-",
        "Coluna 1": "-",
    },
    {
        "Nº LOJA": "B032",
        "REFERENCIA LOJA": "SALVADOR 2",
        "UF": "BA",
        "ESTADO": "BAHIA",
        "NOME": "-",
        "EMAIL": "-",
        "TELEFONE": "-",
        "SETOR": "LOJA",
        "CARGO": "-",
        "N DO GRUPO": "2",
        "GRUPO": "LÍDERES LOJA",
        "SAP ID": "-",
        "Gmail padrão Recebimento": "-",
        "Coluna 1": "-",
    },
    {
        "Nº LOJA": "B032",
        "REFERENCIA LOJA": "SALVADOR 2",
        "UF": "BA",
        "ESTADO": "BAHIA",
        "NOME": "-",
        "EMAIL": "-",
        "TELEFONE": "-",
        "SETOR": "ADMINISTRATIVO",
        "CARGO": "-",
        "N DO GRUPO": "3",
        "GRUPO": "ADMINISTRATIVO",
        "SAP ID": "-",
        "Gmail padrão Recebimento": "-",
        "Coluna 1": "-",
    },
    {
        "Nº LOJA": "B032",
        "REFERENCIA LOJA": "SALVADOR 2",
        "UF": "BA",
        "ESTADO": "BAHIA",
        "NOME": "-",
        "EMAIL": "-",
        "TELEFONE": "-",
        "SETOR": "RECEBIMENTO / TROCAS",
        "CARGO": "-",
        "N DO GRUPO": "7",
        "GRUPO": "RECEBIMENTO / TROCAS",
        "SAP ID": "-",
        "Gmail padrão Recebimento": "-",
        "Coluna 1": "-",
    },
]
# ==============================================================================
# MAPEAMENTO EXATO DAS REGIONAIS
# ==============================================================================
CENTROS_REGIONAL_1 = ["B013", "B015", "B016", "B017", "B019", "B020", "B021", "B022", "B023", "B024", "B025", "B026", "B027", "B028", "B029", "B031", "B032"]
CENTROS_REGIONAL_2 = ["B001", "B002", "B006", "B007", "B008", "B009", "B010", "B011", "B012", "B018", "B030"]

STR_REGIONAL_1 = ",".join(CENTROS_REGIONAL_1)
STR_REGIONAL_2 = ",".join(CENTROS_REGIONAL_2)

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
    "suzana.pires@vonnycosmeticos.com.br": ("B022", "Gerente"),
    "gisele.trampusch@vonnycosmeticos.com.br": ("B023", "Gerente"),
    "raquel.lopes@vonnycosmeticos.com.br": ("B024", "Gerente"),
    "claudinea.santos@vonnycosmeticos.com.br": ("B025", "Gerente"),
    "rosania.chagas@vonnycosmeticos.com.br": ("B026", "Gerente"),
    "luana.costa@vonnycosmeticos.com.br": ("B027", "Gerente"),
    "rosangela.botelho@vonnycosmeticos.com.br": ("B028", "Gerente"),
    "elza.silva@vonnycosmeticos.com.br": ("B029", "Gerente"),
    "joao.pereira@vonnycosmeticos.com.br": ("B030", "Gerente"),
    "jorgiane.aragao@vonnycosmeticos.com.br": ("B031", "Gerente"),
    "manuelle.ramos@vonnycosmeticos.com.br": ("B032", "Gerente"),
    "sergio.oliveira@vonnycosmeticos.com.br": ("TODAS", "Administrador"),
    "controladoriaprevencao@gmail.com": ("TODAS", "Administrador"),
    "josue.victor@vonnycosmeticos.com.br": ("TODAS", "Administrador"),
    "Jvn221106@gmail.com": ("TODAS", "Administrador"),
    "vanusia.garcia@casadolojista.com.br": ("TODAS", "Controladoria"),
    "luciana.valle@vonnycosmeticos.com.br": (STR_REGIONAL_1, "Regional 1"),
    "diego.clodes@vonnycosmeticos.com.br": (STR_REGIONAL_2, "Gerente de produtos 2"),
    "anderson.rodrigues@vonnycosmeticos.com.br": (STR_REGIONAL_1, "Gerente de produtos 1"),
    "jose.marcello@vonnycosmeticos.com.br": ("B001", "Líder de Loja"),
    "trocas@vonnycosmeticos.com.br": ("B001", "Líder de Loja"),
    "veronica.bernardo@vonnycosmeticos.com.br": ("B001", "Líder de Loja"),
    "tatiane.tamizara@vonnycosmeticos.com.br": ("B002", "Líder de Loja"),
    "jessica.barros@vonnycosmeticos.com.br": ("B006", "Líder de Loja"),
    "elenilza.tavares@vonnycosmeticos.com.br": ("B007", "Líder de Loja"),
    "cibele.santos@vonnycosmeticos.com.br": ("B008", "Líder de Loja"),
    "jeane.lopes@vonnycosmeticos.com.br": ("B010", "Líder de Loja"),
    "thais.amorim@vonnycosmeticos.com.br": ("B011", "Líder de Loja"),
    "juliana.lima@vonnycosmeticos.com.br": ("B013", "Líder de Loja"),
    "luciana.inacio@vonnycosmeticos.com.br": ("B015", "Líder de Loja"),
    "tabatta.silva@vonnycosmeticos.com.br": ("B016", "Líder de Loja"),
    "ianara.bossi@vonnycosmeticos.com.br": ("B016", "Líder de Loja"),
    "mariana.ribeiro@vonnycosmeticos.com.br": ("B017", "Líder de Loja"),
    "hector.ramon@vonnycosmeticos.com.br": ("B018", "Líder de Loja"),
    "angelica.santos@vonnycosmeticos.com.br": ("B018", "Líder de Loja"),
    "gisele.correa@vonnycosmeticos.com.br": ("B018", "Líder de Loja"),
    "barbara.oliveira@vonnycosmeticos.com.br": ("B019", "Líder de Loja"),
    "celena.monteiro@vonnycosmeticos.com.br": ("B020", "Líder de Loja"),
    "tamires.oliveira@vonnycosmeticos.com.br": ("B021", "Líder de Loja"),
    "thacilla.paixao@vonnycosmeticos.com.br": ("B022", "Líder de Loja"),
    "debora.mazzilli@vonnycosmeticos.com.br": ("B023", "Líder de Loja"),
    "tays.silva@vonnycosmeticos.com.br": ("B024", "Líder de Loja"),
    "priscila.portela@vonnycosmeticos.com.br": ("B025", "Líder de Loja"),
    "mirela.santos@vonnycosmeticos.com.br": ("B026", "Líder de Loja"),
    "acacia.lima@vonnycosmeticos.com.br": ("B027", "Líder de Loja"),
    "marilane.oliveira@vonnycosmeticos.com.br": ("B027", "Líder de Loja"),
    "pablo.campelo@vonnycosmeticos.com.br": ("B028", "Líder de Loja"),
    "elayne.coutinho@vonnycosmeticos.com.br": ("B029", "Líder de Loja"),
    "tassiana.gomes@vonnycosmeticos.com.br": ("B030", "Líder de Loja"),
    "giovanna.flor@vonnycosmeticos.com.br": ("B030", "Líder de Loja"),
    "kelly.santos@vonnycosmeticos.com.br": ("B009", "Líder de Loja"),
    "magnum.torres@vonnycosmeticos.com.br": ("B031", "Líder de Loja")
}

OPCOES_PERFIL = ["Gerente", "Líder de Loja", "Regional 1", "Regional 2", "Administrador", "Gerente de produtos 1", "Gerente de produtos 2", "Controladoria"]

# --- VALIDAÇÃO DE COMPLEXIDADE DE SENHA ---
def validar_complexidade_senha(senha):
    padrao = r"^(?=.*[a-z])(?=.*[A-Z])(?=.*\d).{6,}$"
    if not re.match(padrao, senha):
        return False, "⚠️ A senha deve conter pelo menos 6 caracteres, incluindo uma letra maiúscula, uma letra minúscula e um número."
    return True, ""

# --- FUNÇÕES DE BANCO DE DADOS E AUTENTICAÇÃO ---
def carregar_dados_db():
    if os.path.exists(DB_FILE):
        try:
            with open(DB_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
                if "usuarios" in data:
                    usuarios = data["usuarios"]
                    removidos = set(data.get("removidos", []))
                    historico_remocoes = data.get("historico_remocoes", [])
                    historico_resets = data.get("historico_resets", [])
                else:
                    usuarios = data
                    removidos = set()
                    historico_remocoes = []
                    historico_resets = []
        except Exception:
            usuarios, removidos, historico_remocoes, historico_resets = {}, set(), [], []
    else:
        usuarios, removidos, historico_remocoes, historico_resets = {}, set(), [], []

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
        salvar_dados_db(usuarios, removidos, historico_remocoes, historico_resets)

    return usuarios, removidos, historico_remocoes, historico_resets

def salvar_dados_db(usuarios, removidos, historico_remocoes=None, historico_resets=None):
    if historico_remocoes is None:
        historico_remocoes = []
    if historico_resets is None:
        historico_resets = []
    with open(DB_FILE, "w", encoding="utf-8") as f:
        json.dump({
            "usuarios": usuarios,
            "removidos": list(removidos),
            "historico_remocoes": historico_remocoes,
            "historico_resets": historico_resets
        }, f, indent=4, ensure_ascii=False)

def gerar_hash(senha):
    return hashlib.sha256(senha.encode('utf-8')).hexdigest()

def atualizar_senha_com_historico(email, nova_senha_texto, usuarios_dict, removidos_set, historico_remocoes, historico_resets):
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
    
    salvar_dados_db(usuarios_dict, removidos_set, historico_remocoes, historico_resets)
    return True, "✅ Senha alterada com sucesso!"

# --- FUNÇÕES DE CARREGAMENTO E SALVAMENTO DE LOJAS/CONTATOS ---
def carregar_dados_lojas():
    if os.path.exists(LOJAS_DB_FILE):
        try:
            with open(LOJAS_DB_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
                return data.get("dados_lojas", DADOS_LOJAS_INICIAIS), data.get("historico_mudancas", [])
        except Exception:
            return DADOS_LOJAS_INICIAIS, []
    else:
        return DADOS_LOJAS_INICIAIS, []

def salvar_dados_lojas(dados_lojas, historico_mudancas):
    with open(LOJAS_DB_FILE, "w", encoding="utf-8") as f:
        json.dump({
            "dados_lojas": dados_lojas,
            "historico_mudancas": historico_mudancas
        }, f, indent=4, ensure_ascii=False)

# --- LEITURA E TRATAMENTO DA PLANILHA NUVEM ---
@st.cache_data(ttl=60)
def load_data():
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"
    }
    response = requests.get(URL_EXCEL_NUVEM, headers=headers, timeout=15)
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

    col_qtd = achar_coluna(df, ['qtd. um registro', 'qtd', 'registro']) or df.columns[0]
    col_valor = achar_coluna(df, ['montante em mi', 'montante', 'mi']) or df.columns[1]
    col_loja = achar_coluna(df, ['centro']) or df.columns[2]
    col_marca = achar_coluna(df, ['fornecedor2', 'fornecedor', 'marca']) or df.columns[3]
    col_material = achar_coluna(df, ['material', 'código', 'codigo'])
    col_texto_mat = achar_coluna(df, ['texto breve material', 'descrição', 'descricao', 'texto breve'])
    col_data = achar_coluna(df, ['data de lançamento', 'lançamento', 'data'])

    df['Qtd_Limpa'] = pd.to_numeric(df[col_qtd], errors='coerce').fillna(0.0)
    df['Valor_Limpo'] = pd.to_numeric(df[col_valor], errors='coerce').fillna(0.0)
    
    df['Loja_Nome'] = df[col_loja].astype(str).fillna('').str.strip()
    df['Loja_Nome'] = df['Loja_Nome'].replace(['nan', 'None', 'NaN', 'none', ''], 'S/ Centro')
    
    df['Marca_Nome'] = df[col_marca].astype(str).fillna('').str.strip()
    df['Marca_Nome'] = df['Marca_Nome'].replace(['nan', 'None', 'NaN', 'none', ''], 'Sem Marca')

    if col_material:
        df['Material_Codigo'] = (
            pd.to_numeric(df[col_material], errors='coerce')
            .fillna(0)
            .astype(int)
            .astype(str)
            .replace('0', 'S/ Codigo')
        )
    else:
        df['Material_Codigo'] = 'S/ Codigo'

    df['Material_Nome'] = df[col_texto_mat].astype(str).fillna('').str.strip() if col_texto_mat else 'S/ Descrição'

    if col_data:
        df['Data_dt'] = pd.to_datetime(df[col_data], dayfirst=True, errors='coerce')
        df['Mes_Ano'] = df['Data_dt'].dt.strftime('%m/%Y').fillna('Sem Data')
    else:
        df['Mes_Ano'] = 'Sem Data'

    df = df[(df['Loja_Nome'] != 'S/ Centro') & (df['Marca_Nome'] != 'Sem Marca')].copy()

    def classificar_centro(centro):
        c = str(centro).strip().upper()
        if c in CENTROS_REGIONAL_1:
            return 'Regional 1'
        elif c in CENTROS_REGIONAL_2:
            return 'Regional 2'
        return 'Sem Regional'

    df['Regional_Nome'] = df['Loja_Nome'].apply(classificar_centro)
    return df

def formatar_moeda(val):
    prefix = "-R$" if val < 0 else "R$"
    return f"{prefix} {abs(val):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

def formatar_qtd(val):
    prefix = "-" if val < 0 else ""
    return f"{prefix}{abs(val):,.0f} UN".replace(",", ".")

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

    usuarios, removidos, historico_remocoes, historico_resets = carregar_dados_db()

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
            valida, msg_validacao = validar_complexidade_senha(senha_input)
            if not valida:
                st.warning(f"⚠️ **Primeiro Acesso:** {msg_validacao}")
            else:
                sucesso, msg = atualizar_senha_com_historico(
                    email_input, senha_input, usuarios, removidos, historico_remocoes, historico_resets
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
        st.info("📩 Por favor, abra um chamado para o setor de **Controladoria / Prevenção de Perdas** solicitando a redefinição de sua senha.\n\n⚠️ *Lembre-se: Por motivos de segurança, você não poderá reutilizar nenhuma das suas últimas 3 senhas.*")

# --- TELA OBRIGATÓRIA DE REDEFINIÇÃO DE SENHA ---
def renderizar_tela_troca_obrigatoria():
    st.title("🔑 Redefinição de Senha Obrigatória")
    st.warning("Você acessou com uma **senha temporária**. Escolha uma nova senha definitiva para continuar.")

    usuarios, removidos, historico_remocoes, historico_resets = carregar_dados_db()
    email_logado = st.session_state["usuario_atual"]

    with st.form("form_troca_obrigatoria"):
        nova_senha = st.text_input("Nova Senha:", type="password")
        confirma_nova = st.text_input("Confirme a Nova Senha:", type="password")
        btn_salvar = st.form_submit_button("Salvar Nova Senha", type="primary")

    if btn_salvar:
        valida, msg_validacao = validar_complexidade_senha(nova_senha)
        if not valida:
            st.error(msg_validacao)
            return

        if nova_senha != confirma_nova:
            st.error("As senhas não coincidem.")
            return

        sucesso, msg = atualizar_senha_com_historico(
            email_logado, nova_senha, usuarios, removidos, historico_remocoes, historico_resets
        )
        if sucesso:
            st.session_state["troca_obrigatoria"] = False
            st.success("✅ Senha atualizada com sucesso!")
            st.rerun()
        else:
            st.error(msg)

# --- ABA DE GESTÃO DE LOJAS E CONTATOS (NOVO MÓDULO) ---
def renderizar_aba_gestao_lojas():
    st.header("🏬 Gestão de Contatos das Lojas e Histórico de Mudanças")
    dados_lojas, historico_mudancas = carregar_dados_lojas()
    df_lojas = pd.DataFrame(dados_lojas)

    tab_vis, tab_edit, tab_logs = st.tabs(["📊 Visão Geral (Tabela)", "✏️ Atualizar / Editar Dados", "📜 Histórico de Mudanças"])

    with tab_vis:
        st.subheader("📋 Tabela Consolidada de Lojas e Contatos")

        # Filtra a tabela conforme o perfil e as lojas vinculadas ao usuário.
        email_logado = st.session_state.get("usuario_atual")
        usuarios, _, _, _ = carregar_dados_db()
        dados_usr = usuarios.get(email_logado, {})
        perfil_usuario = dados_usr.get("perfil", "Gerente")
        loja_usuario = dados_usr.get("loja", "")

        if perfil_usuario in ["Administrador", "Controladoria"]:
            df_lojas_vis = df_lojas.copy()
        else:
            lojas_permitidas = [
                x.strip().upper()
                for x in str(loja_usuario).replace(" ", ",").split(",")
                if x.strip()
            ]
            df_lojas_vis = df_lojas[
                df_lojas["Nº LOJA"].astype(str).str.strip().str.upper().isin(lojas_permitidas)
            ].copy()

        # Oculta o SAP ID somente na tabela consolidada.
        df_lojas_vis = df_lojas_vis.drop(columns=["SAP ID"], errors="ignore")
        st.dataframe(df_lojas_vis, use_container_width=True)

        # Permite preencher diretamente os registros que ainda estão vazios.
        # Registros já preenchidos continuam sendo alterados pela aba de edição.
        def _campo_vazio(valor):
            if pd.isna(valor):
                return True
            texto = str(valor).strip()
            return texto == "" or texto in {"-", "nan", "None"}

        if not df_lojas_vis.empty:
            mascara_vazios = df_lojas_vis["NOME"].apply(_campo_vazio) if "NOME" in df_lojas_vis.columns else pd.Series(False, index=df_lojas_vis.index)
            df_vazios = df_lojas_vis[mascara_vazios].copy()
        else:
            df_vazios = pd.DataFrame()

        if not df_vazios.empty:
            st.markdown("### ✏️ Preencher registros vazios diretamente na tabela")
            st.caption("Use esta área somente para os espaços ainda vazios. Registros que já possuem colaborador continuam sendo atualizados pela aba 'Atualizar / Editar Dados'.")

            # Grupos e respectivos números já cadastrados no banco.
            mapa_grupo_numero = {}
            if "GRUPO" in df_lojas.columns and "N DO GRUPO" in df_lojas.columns:
                for _, rg in df_lojas[["GRUPO", "N DO GRUPO"]].dropna(how="all").iterrows():
                    g = str(rg.get("GRUPO", "")).strip()
                    n = str(rg.get("N DO GRUPO", "")).strip()
                    if g and g not in {"-", "nan", "None"} and n and n not in {"-", "nan", "None"}:
                        mapa_grupo_numero.setdefault(g, n)
            grupos_disponiveis = sorted(mapa_grupo_numero.keys())

            colunas_editaveis = [
                "Nº LOJA", "REFERENCIA LOJA", "UF", "ESTADO", "NOME", "EMAIL",
                "TELEFONE", "SETOR", "CARGO", "GRUPO", "N DO GRUPO",
                "Gmail padrão Recebimento", "Coluna 1"
            ]
            colunas_editaveis = [c for c in colunas_editaveis if c in df_vazios.columns]
            editor_df = df_vazios[colunas_editaveis].copy()

            # Campos técnicos/preenchidos pela loja ficam bloqueados.
            disabled_cols = [c for c in ["Nº LOJA", "REFERENCIA LOJA", "UF", "ESTADO", "SETOR", "N DO GRUPO", "Coluna 1"] if c in editor_df.columns]
            if perfil_usuario != "Administrador" and "Gmail padrão Recebimento" in editor_df.columns:
                disabled_cols.append("Gmail padrão Recebimento")

            column_config = {}
            if "GRUPO" in editor_df.columns and grupos_disponiveis:
                column_config["GRUPO"] = st.column_config.SelectboxColumn(
                    "GRUPO", options=grupos_disponiveis, required=False
                )
            if "N DO GRUPO" in editor_df.columns:
                column_config["N DO GRUPO"] = st.column_config.TextColumn("Nº do Grupo")
            if "Coluna 1" in editor_df.columns:
                column_config["Coluna 1"] = st.column_config.TextColumn("Observações", disabled=True)

            edited_vazios = st.data_editor(
                editor_df,
                key="editor_registros_vazios",
                use_container_width=True,
                hide_index=False,
                disabled=disabled_cols,
                column_config=column_config,
                num_rows="fixed"
            )

            if st.button("💾 Salvar preenchimentos da tabela", type="primary", key="salvar_vazios_tabela"):
                dados_lojas_atual = list(dados_lojas)
                historico_atual = list(historico_mudancas)
                usr_atual = st.session_state.get("usuario_atual", "Administrador")
                alterou = False
                erros = []

                # Processa de baixo para cima para que remoções por transferência não
                # alterem os índices das vagas que ainda serão preenchidas.
                linhas_editadas = list(edited_vazios.iterrows())
                linhas_editadas.sort(key=lambda item: item[0], reverse=True)
                for idx, linha_editada in linhas_editadas:
                    nome_novo = str(linha_editada.get("NOME", "")).strip()
                    if not nome_novo or nome_novo in {"-", "nan", "None"}:
                        continue

                    # Localiza o registro original correspondente ao índice do DataFrame.
                    idx_original = idx
                    if idx_original not in df_lojas.index:
                        continue
                    registro_original = dados_lojas[idx_original]
                    loja_destino_num = str(registro_original.get("Nº LOJA", linha_editada.get("Nº LOJA", ""))).strip()
                    loja_destino_ref = str(registro_original.get("REFERENCIA LOJA", linha_editada.get("REFERENCIA LOJA", ""))).strip()

                    # Segurança: gerente/líder só podem preencher a própria loja.
                    if perfil_usuario in ["Gerente", "Líder de Loja"]:
                        permitidas = [x.strip().upper() for x in str(loja_usuario).replace(" ", ",").split(",") if x.strip()]
                        if loja_destino_num.upper() not in permitidas:
                            erros.append(f"{nome_novo}: loja não autorizada.")
                            continue

                    grupo_novo = str(linha_editada.get("GRUPO", "")).strip()
                    num_grupo_novo = mapa_grupo_numero.get(grupo_novo, str(linha_editada.get("N DO GRUPO", "")).strip())
                    email_rec_novo = str(linha_editada.get("Gmail padrão Recebimento", registro_original.get("Gmail padrão Recebimento", ""))).strip()
                    observacao_auto = f"Atualizado em {datetime.datetime.now().strftime('%d/%m/%Y')} por {usr_atual}"

                    novo_registro = dict(registro_original)
                    novo_registro.update({
                        "Nº LOJA": loja_destino_num,
                        "REFERENCIA LOJA": loja_destino_ref,
                        "NOME": nome_novo,
                        "EMAIL": str(linha_editada.get("EMAIL", "")).strip(),
                        "TELEFONE": str(linha_editada.get("TELEFONE", "")).strip(),
                        "CARGO": str(linha_editada.get("CARGO", "")).strip(),
                        "GRUPO": grupo_novo,
                        "N DO GRUPO": num_grupo_novo,
                        "Gmail padrão Recebimento": email_rec_novo if perfil_usuario == "Administrador" else registro_original.get("Gmail padrão Recebimento", ""),
                        "Coluna 1": observacao_auto
                    })

                    nome_norm = nome_novo.casefold()
                    indices_nome = [
                        i for i, r in enumerate(dados_lojas_atual)
                        if i != idx_original and nome_norm and str(r.get("NOME", "")).strip().casefold() == nome_norm
                    ]

                    if indices_nome:
                        origem_idx = indices_nome[0]
                        origem = dados_lojas_atual[origem_idx]
                        origem_num = str(origem.get("Nº LOJA", "")).strip()
                        origem_ref = str(origem.get("REFERENCIA LOJA", "")).strip()

                        # Para preencher uma vaga, a transferência de alguém de outra loja
                        # também respeita as permissões de gerente/líder.
                        if perfil_usuario in ["Gerente", "Líder de Loja"]:
                            permitidas = [x.strip().upper() for x in str(loja_usuario).replace(" ", ",").split(",") if x.strip()]
                            if origem_num.upper() not in permitidas:
                                erros.append(f"{nome_novo}: colaborador pertence a outra loja.")
                                continue

                        if not email_rec_novo:
                            novo_registro["Gmail padrão Recebimento"] = origem.get("Gmail padrão Recebimento", "")
                        # Remove a pessoa da loja de origem e preenche a vaga.
                        dados_lojas_atual.pop(origem_idx)
                        # Ajusta o índice da vaga caso ela esteja depois da origem.
                        idx_insercao = idx_original - 1 if origem_idx < idx_original else idx_original
                        if 0 <= idx_insercao < len(dados_lojas_atual):
                            dados_lojas_atual[idx_insercao] = novo_registro
                        else:
                            dados_lojas_atual.append(novo_registro)

                        historico_atual.append({
                            "Data/Hora": datetime.datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
                            "Usuário Responsável": usr_atual,
                            "Loja Afetada": f"{loja_destino_num} - {loja_destino_ref}",
                            "Colaborador": nome_novo,
                            "Descrição da Mudança": f"Colaborador transferido de {origem_num} - {origem_ref} para {loja_destino_num} - {loja_destino_ref}.",
                            "Observações": observacao_auto
                        })
                    else:
                        dados_lojas_atual[idx_original] = novo_registro
                        historico_atual.append({
                            "Data/Hora": datetime.datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
                            "Usuário Responsável": usr_atual,
                            "Loja Afetada": f"{loja_destino_num} - {loja_destino_ref}",
                            "Colaborador": nome_novo,
                            "Descrição da Mudança": "Cadastro de colaborador preenchido diretamente pela tabela.",
                            "Observações": observacao_auto
                        })

                    alterou = True

                if erros:
                    for erro in erros:
                        st.error(f"❌ {erro}")

                if alterou and not erros:
                    salvar_dados_lojas(dados_lojas_atual, historico_atual)
                    st.success("✅ Registros vazios preenchidos e salvos com sucesso!")
                    st.rerun()
                elif alterou:
                    salvar_dados_lojas(dados_lojas_atual, historico_atual)
                    st.warning("⚠️ Alguns registros foram salvos e outros não puderam ser alterados.")
                    st.rerun()
        else:
            st.info("Não há registros vazios disponíveis para preenchimento direto nesta loja/visão.")

    with tab_edit:
        st.subheader("📝 Adicionar ou Modificar Registro de Loja")

        # Cadastro de uma nova loja: disponível para perfis administrativos/regionais.
        # Gerentes e Líderes de Loja continuam limitados à própria loja e não podem criar lojas.
        if perfil_usuario in ["Administrador", "Controladoria", "Regional 1", "Regional 2"]:
            with st.expander("➕ Cadastrar nova loja", expanded=False):
                st.caption("Cadastre primeiro a loja. Depois, os colaboradores podem ser preenchidos pela tabela de registros vazios.")
                nova_col1, nova_col2, nova_col3 = st.columns(3)
                with nova_col1:
                    nova_num_loja = st.text_input("Nº da Loja *", key="nova_loja_num")
                    nova_ref_loja = st.text_input("Referência da Loja *", key="nova_loja_ref")
                with nova_col2:
                    nova_uf = st.text_input("UF", value="SP", key="nova_loja_uf")
                    nova_estado = st.text_input("Estado / Cidade", value="SÃO PAULO", key="nova_loja_estado")
                with nova_col3:
                    nova_setor = st.selectbox("Setor inicial", ["LOJA", "ADMINISTRATIVO", "RECEBIMENTO / TROCAS", "RECEBIMENTO"], key="nova_loja_setor")

                if st.button("➕ Criar Nova Loja", type="primary", key="criar_nova_loja"):
                    num_novo = str(nova_num_loja).strip().upper()
                    ref_nova = str(nova_ref_loja).strip().upper()

                    if not num_novo or not ref_nova:
                        st.error("❌ Informe o Nº da Loja e a Referência da Loja.")
                    else:
                        loja_ja_existe = any(
                            str(r.get("Nº LOJA", "")).strip().upper() == num_novo or
                            str(r.get("REFERENCIA LOJA", "")).strip().upper() == ref_nova
                            for r in dados_lojas
                        )

                        if loja_ja_existe:
                            st.error("❌ Essa loja já está cadastrada pelo Nº ou pela Referência.")
                        else:
                            usr_novo = st.session_state.get("usuario_atual", "Administrador")
                            obs_nova = f"Loja criada em {datetime.datetime.now().strftime('%d/%m/%Y')} por {usr_novo}"
                            novo_registro_loja = {
                                "Nº LOJA": num_novo,
                                "REFERENCIA LOJA": ref_nova,
                                "UF": str(nova_uf).strip().upper(),
                                "ESTADO": str(nova_estado).strip().upper(),
                                "NOME": "",
                                "EMAIL": "",
                                "TELEFONE": "",
                                "SETOR": nova_setor,
                                "CARGO": "",
                                "N DO GRUPO": "",
                                "GRUPO": "",
                                "Gmail padrão Recebimento": "",
                                "Coluna 1": obs_nova
                            }
                            dados_lojas.append(novo_registro_loja)
                            historico_mudancas.append({
                                "Data/Hora": datetime.datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
                                "Usuário Responsável": usr_novo,
                                "Loja Afetada": f"{num_novo} - {ref_nova}",
                                "Colaborador": "",
                                "Descrição da Mudança": "Nova loja cadastrada.",
                                "Observações": obs_nova
                            })
                            salvar_dados_lojas(dados_lojas, historico_mudancas)
                            st.success(f"✅ Loja {num_novo} - {ref_nova} criada com sucesso!")
                            st.rerun()

        # Gerentes e Líderes de Loja só podem trabalhar com a própria loja.
        if perfil_usuario in ["Gerente", "Líder de Loja"]:
            lojas_permitidas_edit = [
                x.strip().upper()
                for x in str(loja_usuario).replace(" ", ",").split(",")
                if x.strip()
            ]
            lojas_edit_df = df_lojas[
                df_lojas["Nº LOJA"].astype(str).str.strip().str.upper().isin(lojas_permitidas_edit)
            ].copy()
            lojas_edit_df["_OPCAO"] = (
                lojas_edit_df["Nº LOJA"].astype(str).str.strip() + " - " +
                lojas_edit_df["REFERENCIA LOJA"].astype(str).str.strip()
            )
            loja_opcoes = list(dict.fromkeys(lojas_edit_df["_OPCAO"].tolist()))
            if not loja_opcoes:
                st.error("❌ Nenhuma loja está vinculada ao seu usuário.")
                return
            loja_sel = st.selectbox("Loja autorizada para edição:", loja_opcoes, disabled=True)
            loja_num_autorizada, loja_ref_autorizada = [x.strip() for x in loja_sel.split(" - ", 1)]
        else:
            loja_opcoes = ["-- Nova Entrada --"] + list(df_lojas["REFERENCIA LOJA"].unique()) if not df_lojas.empty else ["-- Nova Entrada --"]
            loja_sel = st.selectbox("Selecione uma Loja Existente para Editar ou Crie Uma Nova:", loja_opcoes)
            loja_num_autorizada = "" if loja_sel == "-- Nova Entrada --" else loja_sel
            loja_ref_autorizada = "" if loja_sel == "-- Nova Entrada --" else loja_sel

        # Ação: permite retirar um colaborador sem apagar a loja inteira.
        acao = st.radio(
            "Ação:",
            ["Adicionar / Atualizar colaborador", "🚪 Retirar colaborador da loja"],
            horizontal=True
        )

        if acao == "🚪 Retirar colaborador da loja":
            if loja_sel == "-- Nova Entrada --":
                st.info("Selecione uma loja existente para retirar um colaborador.")
            else:
                registros_loja = df_lojas[
                    df_lojas["REFERENCIA LOJA"].astype(str).str.strip() == str(loja_ref_autorizada).strip()
                ].copy()
                registros_loja = registros_loja[registros_loja["NOME"].fillna("").astype(str).str.strip() != ""]

                if registros_loja.empty:
                    st.info("Nenhum colaborador cadastrado nesta loja.")
                else:
                    registros_loja["_COLAB"] = (
                        registros_loja["NOME"].astype(str).str.strip() +
                        " — " + registros_loja["GRUPO"].fillna("").astype(str).str.strip()
                    )
                    colaborador_sel = st.selectbox(
                        "Selecione o colaborador que saiu da loja:",
                        registros_loja["_COLAB"].tolist()
                    )
                    idx_colaborador = registros_loja.index[registros_loja["_COLAB"] == colaborador_sel][0]
                    observacao_saida = st.text_area(
                        "Observações da saída *",
                        placeholder="Ex.: Desligamento, transferência para outra loja, mudança de função..."
                    )

                    if st.button("🚪 Retirar Colaborador", type="primary"):
                        if not observacao_saida.strip():
                            st.error("❌ Informe uma observação para registrar a saída do colaborador.")
                        else:
                            colaborador = dados_lojas[idx_colaborador]
                            nome_removido = str(colaborador.get("NOME", "")).strip()
                            dados_lojas.pop(idx_colaborador)

                            usr_atual = st.session_state.get("usuario_atual", "Administrador")
                            historico_mudancas.append({
                                "Data/Hora": datetime.datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
                                "Usuário Responsável": usr_atual,
                                "Loja Afetada": f"{colaborador.get('Nº LOJA', '')} - {colaborador.get('REFERENCIA LOJA', '')}",
                                "Colaborador": nome_removido,
                                "Descrição da Mudança": "Colaborador retirado da loja.",
                                "Observações": observacao_saida.strip()
                            })
                            salvar_dados_lojas(dados_lojas, historico_mudancas)
                            st.success(f"✅ {nome_removido} foi retirado da loja e a saída foi registrada.")
                            st.rerun()
        else:
            col1, col2, col3 = st.columns(3)
            with col1:
                if perfil_usuario in ["Gerente", "Líder de Loja"]:
                    num_loja = st.text_input("Nº Loja:", value=loja_num_autorizada, disabled=True)
                    ref_loja = st.text_input("Referência Loja:", value=loja_ref_autorizada, disabled=True)
                else:
                    num_loja = st.text_input("Nº Loja:", value="" if loja_sel == "-- Nova Entrada --" else loja_sel)
                    ref_loja = st.text_input("Referência Loja:", value="" if loja_sel == "-- Nova Entrada --" else loja_sel)
                uf = st.text_input("UF:", value="SP")
                estado = st.text_input("Estado / Cidade:", value="SÃO PAULO")
                nome = st.text_input("Nome do Colaborador:")

            with col2:
                email = st.text_input("E-mail:")
                telefone = st.text_input("Telefone / Celular:")
                setor = st.selectbox("Setor:", ["LOJA", "COORDENADOR", "ADMINISTRATIVO", "RECEBIMENTO / TROCAS", "RECEBIMENTO", "OUTRO"])
                cargo = st.text_input("Cargo:")

                # O Nº do Grupo é vinculado automaticamente ao Grupo.
                mapa_grupo_numero = {}
                if not df_lojas.empty and "GRUPO" in df_lojas.columns and "N DO GRUPO" in df_lojas.columns:
                    for _, registro_grupo in df_lojas[["GRUPO", "N DO GRUPO"]].dropna().iterrows():
                        nome_grupo = str(registro_grupo["GRUPO"]).strip()
                        numero_grupo = str(registro_grupo["N DO GRUPO"]).strip()
                        if nome_grupo and numero_grupo and nome_grupo.lower() not in ["nan", "none"]:
                            mapa_grupo_numero.setdefault(nome_grupo, numero_grupo)

                grupos_disponiveis = sorted(mapa_grupo_numero.keys())
                grupo_atual_registro = ""
                if loja_sel != "-- Nova Entrada --" and not df_lojas.empty:
                    registros_loja = df_lojas[df_lojas["REFERENCIA LOJA"].astype(str) == str(loja_ref_autorizada)]
                    if not registros_loja.empty and "GRUPO" in registros_loja.columns:
                        valores_grupo = registros_loja["GRUPO"].dropna().astype(str).str.strip()
                        if not valores_grupo.empty:
                            grupo_atual_registro = valores_grupo.iloc[0]

                grupo_inicial = grupo_atual_registro if grupo_atual_registro in grupos_disponiveis else (grupos_disponiveis[0] if grupos_disponiveis else "")
                grupo = st.selectbox("Grupo:", grupos_disponiveis, index=grupos_disponiveis.index(grupo_inicial) if grupo_inicial in grupos_disponiveis else 0, disabled=False) if grupos_disponiveis else st.text_input("Grupo:")
                num_grupo = mapa_grupo_numero.get(grupo, "")
                st.text_input("Nº do Grupo:", value=num_grupo, disabled=True)
                gmail_rec = st.text_input(
                    "Gmail Padrão Recebimento:",
                    disabled=(perfil_usuario != "Administrador")
                )
                email_atualizacao = str(st.session_state.get("usuario_atual", "")).strip()
                observacoes = st.text_area(
                    "Observações:",
                    value=f"Atualizado em {datetime.datetime.now().strftime('%d/%m/%Y')} por {email_atualizacao}",
                    disabled=True
                )

            if st.button("💾 Salvar Registro e Registrar Mudança", type="primary"):
                if perfil_usuario in ["Gerente", "Líder de Loja"]:
                    lojas_permitidas_save = [
                        x.strip().upper()
                        for x in str(loja_usuario).replace(" ", ",").split(",")
                        if x.strip()
                    ]
                    if str(num_loja).strip().upper() not in lojas_permitidas_save:
                        st.error("❌ Você só pode alterar registros da sua própria loja.")
                        return

                novo_registro = {
                    "Nº LOJA": num_loja,
                    "REFERENCIA LOJA": ref_loja,
                    "UF": uf,
                    "ESTADO": estado,
                    "NOME": nome,
                    "EMAIL": email,
                    "TELEFONE": telefone,
                    "SETOR": setor,
                    "CARGO": cargo,
                    "N DO GRUPO": num_grupo,
                    "GRUPO": grupo,
                    "Gmail padrão Recebimento": gmail_rec,
                    "Coluna 1": observacoes
                }

                # Se o colaborador já existir pelo nome, não cria duplicidade.
                # Se estiver em outra loja, realiza a transferência automaticamente.
                nome_normalizado = str(nome).strip().casefold()
                indices_mesmo_nome = [
                    i for i, registro in enumerate(dados_lojas)
                    if nome_normalizado and str(registro.get("NOME", "")).strip().casefold() == nome_normalizado
                ]

                usr_atual = st.session_state.get("usuario_atual", "Administrador")
                registro_existente = None
                indice_existente = None

                if indices_mesmo_nome:
                    # Preferimos um registro já existente na loja de destino.
                    for i in indices_mesmo_nome:
                        if (str(dados_lojas[i].get("Nº LOJA", "")).strip().casefold() == str(num_loja).strip().casefold()
                                or str(dados_lojas[i].get("REFERENCIA LOJA", "")).strip().casefold() == str(ref_loja).strip().casefold()):
                            indice_existente = i
                            registro_existente = dados_lojas[i]
                            break

                    # Caso não esteja na loja de destino, é uma transferência.
                    if indice_existente is None:
                        indice_existente = indices_mesmo_nome[0]
                        registro_existente = dados_lojas[indice_existente]
                        loja_origem_num = str(registro_existente.get("Nº LOJA", "")).strip()
                        loja_origem_ref = str(registro_existente.get("REFERENCIA LOJA", "")).strip()
                        loja_destino_num = str(num_loja).strip()
                        loja_destino_ref = str(ref_loja).strip()

                        # Gerente e Líder de Loja não podem retirar/transferir alguém de outra loja.
                        if perfil_usuario in ["Gerente", "Líder de Loja"]:
                            lojas_permitidas_transfer = [
                                x.strip().upper()
                                for x in str(loja_usuario).replace(" ", ",").split(",")
                                if x.strip()
                            ]
                            if loja_origem_num.upper() not in lojas_permitidas_transfer:
                                st.error("❌ Este colaborador está vinculado a outra loja. Apenas Administrador, Controladoria e Regionais podem realizar essa transferência.")
                                return

                        # Remove o registro da loja antiga antes de inserir na nova.
                        dados_lojas.pop(indice_existente)
                        novo_registro["Gmail padrão Recebimento"] = registro_existente.get("Gmail padrão Recebimento", "")
                        dados_lojas.append(novo_registro)

                        historico_mudancas.append({
                            "Data/Hora": datetime.datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
                            "Usuário Responsável": usr_atual,
                            "Loja Afetada": f"{loja_destino_num} - {loja_destino_ref}",
                            "Colaborador": nome,
                            "Descrição da Mudança": f"Colaborador transferido de {loja_origem_num} - {loja_origem_ref} para {loja_destino_num} - {loja_destino_ref}.",
                            "Observações": observacoes
                        })
                        salvar_dados_lojas(dados_lojas, historico_mudancas)
                        st.success(f"✅ {nome} foi transferido automaticamente de {loja_origem_num} - {loja_origem_ref} para {loja_destino_num} - {loja_destino_ref}.")
                        st.rerun()

                    # Se já estiver na mesma loja, atualiza o cadastro existente.
                    else:
                        dados_lojas[indice_existente] = novo_registro
                        historico_mudancas.append({
                            "Data/Hora": datetime.datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
                            "Usuário Responsável": usr_atual,
                            "Loja Afetada": f"{num_loja} - {ref_loja}",
                            "Colaborador": nome,
                            "Descrição da Mudança": f"Cadastro do colaborador atualizado no setor {setor} ({cargo}).",
                            "Observações": observacoes
                        })
                else:
                    dados_lojas.append(novo_registro)
                    historico_mudancas.append({
                        "Data/Hora": datetime.datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
                        "Usuário Responsável": usr_atual,
                        "Loja Afetada": f"{num_loja} - {ref_loja}",
                        "Colaborador": nome,
                        "Descrição da Mudança": f"Cadastrado contato do setor {setor} ({cargo}).",
                        "Observações": observacoes
                    })

                salvar_dados_lojas(dados_lojas, historico_mudancas)
                st.success("✅ Informações salvas com sucesso no banco de dados!")
                st.rerun()

    with tab_logs:
        st.subheader("📜 Histórico e Registro de Mudanças")
        if historico_mudancas:
            df_logs = pd.DataFrame(historico_mudancas)
            if "Observações" not in df_logs.columns:
                df_logs["Observações"] = ""
            st.dataframe(df_logs, use_container_width=True)
        else:
            st.info("Nenhuma mudança registrada até o momento.")

# --- ABA PAINEL ADMIN ---
def renderizar_aba_admin():
    st.header("⚙️ Painel do Administrador")
    usuarios, removidos, historico_remocoes, historico_resets = carregar_dados_db()

    st.subheader("👥 Lista de Usuários e Status")
    dados_tabela = [
        {
            "E-mail": email,
            "Loja / Centro": dados.get("loja", "N/A"),
            "Perfil / Cargo": dados.get("perfil", "Gerente"),
            "Primeiro Acesso": "✅ Concluído" if dados.get("senha") else "⏳ Pendente",
            "Senha Temporária Ativa": "⚠️ Sim" if dados.get("forcar_redefinicao") else "Não"
        }
        for email, dados in usuarios.items()
    ]
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
                
                salvar_dados_db(usuarios, removidos, historico_remocoes, historico_resets)
                st.rerun()

    st.markdown("---")

    st.subheader("🔑 Resetar Senha / Gerar Senha Temporária")
    col1, col2 = st.columns([2, 1])
    with col1:
        usuario_selecionado = st.selectbox("Selecione o e-mail:", options=list(usuarios.keys()), key="select_reset_senha")
        senha_temp = st.text_input("Senha Temporária:", value="Controladoria", type="password", key="input_senha_temp")
        st.caption("🔑 Senha padrão inicial: **Controladoria**")

    with col2:
        st.write("##")
        if st.button("Definir Senha Temporária"):
            valida, msg_validacao = validar_complexidade_senha(senha_temp)
            if not valida:
                st.error(msg_validacao)
            else:
                admin_atual = st.session_state["usuario_atual"]
                novo_hash_temp = gerar_hash(senha_temp)
                dados_usr = usuarios[usuario_selecionado]
                historico = dados_usr.get("historico_senhas", [])
                senha_atual_esquecida = dados_usr.get("senha")

                if senha_atual_esquecida and senha_atual_esquecida not in historico:
                    historico.insert(0, senha_atual_esquecida)
                    dados_usr["historico_senhas"] = historico[:3]

                dados_usr["senha"] = novo_hash_temp
                dados_usr["forcar_redefinicao"] = True

                registro_reset = {
                    "usuario_afetado": usuario_selecionado,
                    "resetado_por": admin_atual,
                    "data_hora": datetime.datetime.now().strftime("%d/%m/%Y %H:%M:%S")
                }
                historico_resets.append(registro_reset)

                salvar_dados_db(usuarios, removidos, historico_remocoes, historico_resets)
                st.success(f"✅ Senha temporária definida para **{usuario_selecionado}** por **{admin_atual}**!")

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

                salvar_dados_db(usuarios, removidos, historico_remocoes, historico_resets)
                st.success(f"🗑️ Usuário **{user_para_deletar}** removido por **{admin_atual}**!")
                st.rerun()

    st.markdown("---")

    st.subheader("📦 Backup e Restauração dos Históricos (Audit Logs)")
    
    col_exp, col_imp = st.columns(2)

    with col_exp:
        df_exp_resets = pd.DataFrame(historico_resets) if historico_resets else pd.DataFrame(columns=["usuario_afetado", "resetado_por", "data_hora"])
        df_exp_remocoes = pd.DataFrame(historico_remocoes) if historico_remocoes else pd.DataFrame(columns=["usuario_removido", "removido_por", "data_hora"])

        output_audit = io.BytesIO()
        with pd.ExcelWriter(output_audit, engine='openpyxl') as writer:
            df_exp_resets.to_excel(writer, index=False, sheet_name='Historico_Resets')
            df_exp_remocoes.to_excel(writer, index=False, sheet_name='Historico_Remocoes')
        
        excel_audit_bytes = output_audit.getvalue()

        st.download_button(
            label="📥 Exportar Históricos (Excel)",
            data=excel_audit_bytes,
            file_name=f"historico_auditoria_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="secondary"
        )

    with col_imp:
        uploaded_file = st.file_uploader("Importar Arquivo de Históricos (Excel)", type=["xlsx", "xls"], key="upload_audit_logs")
        if uploaded_file is not None:
            try:
                xls_import = pd.ExcelFile(uploaded_file)
                novos_resets, novas_remocoes = 0, 0
                
                if "Historico_Resets" in xls_import.sheet_names:
                    df_imp_resets = pd.read_excel(xls_import, sheet_name="Historico_Resets").fillna("")
                    for item in df_imp_resets.to_dict(orient="records"):
                        if item.get("usuario_afetado") and item not in historico_resets:
                            historico_resets.append(item)
                            novos_resets += 1

                if "Historico_Remocoes" in xls_import.sheet_names:
                    df_imp_remocoes = pd.read_excel(xls_import, sheet_name="Historico_Remocoes").fillna("")
                    for item in df_imp_remocoes.to_dict(orient="records"):
                        if item.get("usuario_removido") and item not in historico_remocoes:
                            historico_remocoes.append(item)
                            novas_remocoes += 1

                if novos_resets > 0 or novas_remocoes > 0:
                    salvar_dados_db(usuarios, removidos, historico_remocoes, historico_resets)
                    st.success(f"✅ Históricos importados com sucesso! ({novos_resets} resets e {novas_remocoes} remoções adicionados)")
                    st.rerun()
                else:
                    st.info("ℹ️ Os históricos importados já existem na base atual ou a planilha está vazia.")
            except Exception as e:
                st.error(f"Erro ao processar o arquivo de importação: {e}")

    st.markdown("---")

    col_audit1, col_audit2 = st.columns(2)

    with col_audit1:
        st.subheader("🔑 Histórico de Resets de Senha")
        if historico_resets:
            df_resets = pd.DataFrame(historico_resets).rename(columns={
                "usuario_afetado": "Usuário Afetado",
                "resetado_por": "Resetado por (Admin)",
                "data_hora": "Data e Horário"
            })
            st.dataframe(df_resets, use_container_width=True, hide_index=True)
        else:
            st.info("Nenhum reset de senha registrado até o momento.")

    with col_audit2:
        st.subheader("📋 Histórico de Remoções")
        if historico_remocoes:
            df_historico = pd.DataFrame(historico_remocoes).rename(columns={
                "usuario_removido": "Usuário Removido",
                "removido_por": "Removido por (Admin)",
                "data_hora": "Data e Horário"
            })
            st.dataframe(df_historico, use_container_width=True, hide_index=True)
        else:
            st.info("Nenhuma remoção registrada até o momento.")

# --- DASHBOARD VISUAL DE INVENTÁRIO ---
def renderizar_dashboard():
    try:
        usuarios, _, _, _ = carregar_dados_db()
        email_logado = st.session_state["usuario_atual"]
        dados_usr = usuarios.get(email_logado, {})
        
        loja_usuario = dados_usr.get("loja", "N/A")
        perfil_usuario = dados_usr.get("perfil", "Gerente")

        df = load_data()

        st.sidebar.title("📌 Filtro Geral")

        if st.sidebar.button("🔄 Atualizar Dados"):
            st.cache_data.clear()
            st.rerun()

        meses_unicos = sorted([x for x in df['Mes_Ano'].unique() if x != 'Sem Data'])
        meses_sel = st.sidebar.multiselect("Mês/Ano (Geral):", options=meses_unicos, default=meses_unicos)

        regionais_disponiveis = [str(r) for r in ["Regional 1", "Regional 2"] if r in df['Regional_Nome'].unique()]

        if perfil_usuario in ["Administrador", "Controladoria"]:
            regionais_sel = st.sidebar.multiselect(
                "Divisão Regional (Geral):", 
                options=regionais_disponiveis, 
                default=regionais_disponiveis
            )
        elif perfil_usuario in ["Regional 1", "Regional 2"]:
            regionais_sel = [perfil_usuario]
        else:
            regionais_sel = regionais_disponiveis

        lojas_unicas = [str(x) for x in df[df['Regional_Nome'].isin(regionais_sel)]['Loja_Nome'].unique() if str(x).lower() not in ['nan', 'none', '', 'sem centro', 's/ centro']]
        lojas_disponiveis = sorted(lojas_unicas)

        if perfil_usuario in ["Administrador", "Controladoria"]:
            lojas_sel = st.sidebar.multiselect("Centros (Geral):", options=lojas_disponiveis, default=lojas_disponiveis)
        elif perfil_usuario in ["Regional 1", "Regional 2"]:
            lojas_permitidas_usr = [x.strip() for x in str(loja_usuario).replace(" ", ",").split(",") if x.strip()]
            lojas_filtradas_usr = [x for x in lojas_disponiveis if x in lojas_permitidas_usr] if lojas_permitidas_usr else lojas_disponiveis
            lojas_sel = st.sidebar.multiselect("Centros (Geral):", options=lojas_filtradas_usr, default=lojas_filtradas_usr)
        else:
            lojas_permitidas_usr = [x.strip() for x in str(loja_usuario).replace(" ", ",").split(",") if x.strip()]
            lojas_sel = [x for x in lojas_disponiveis if x in lojas_permitidas_usr] or lojas_disponiveis

        marcas_unicas = [str(x) for x in df['Marca_Nome'].unique() if str(x).lower() not in ['nan', 'none', '', 'sem marca']]
        marcas_sel = st.sidebar.multiselect("Marcas (Geral):", options=sorted(marcas_unicas), default=sorted(marcas_unicas))

        # Base Global Filtrada
        df_filtered = df[
            (df['Mes_Ano'].isin(meses_sel)) &
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

        if perfil_usuario in ["Administrador", "Regional 1", "Regional 2"]:
            # ANÁLISE DE PRODUTOS
            st.subheader("🛍️ Análise de Produtos por Nível de Perda")
            tab_unificado, tab_por_loja = st.tabs(["🌐 Unificado (Geral)", "🏬 Por Loja (Centro)"])
            
            df_perdas_prod = df_filtered[(df_filtered['Qtd_Limpa'] < 0) | (df_filtered['Valor_Limpo'] < 0)].copy()
            
            with tab_unificado:
                with st.expander("🔍 Filtro Local: Produtos (Unificado)"):
                    col_fu1, col_fu2, col_fu3 = st.columns(3)
                    with col_fu1:
                        m_unif_sel = st.multiselect("Filtrar Marcas:", options=sorted(df_perdas_prod['Marca_Nome'].unique()), default=sorted(df_perdas_prod['Marca_Nome'].unique()), key="f_prod_unif_marca")
                    with col_fu2:
                        mes_unif_sel = st.multiselect("Filtrar Mês/Ano:", options=sorted(df_perdas_prod['Mes_Ano'].unique()), default=sorted(df_perdas_prod['Mes_Ano'].unique()), key="f_prod_unif_mes")
                    with col_fu3:
                        top_n_unif = st.slider("Selecione o Top (Produtos Unificado):", min_value=10, max_value=100, value=15, step=5, key="top_prod_unif")
    
                df_perdas_unif_f = df_perdas_prod[
                    (df_perdas_prod['Marca_Nome'].isin(m_unif_sel)) &
                    (df_perdas_prod['Mes_Ano'].isin(mes_unif_sel))
                ]
    
                st.markdown("#### 🚨 Maiores Perdas de Produtos (Unificado)")
                col_unif_qtd, col_unif_val = st.columns(2)
                
                with col_unif_qtd:
                    st.markdown(f"##### 📦 Top {top_n_unif} MAIORES por Quantidade (UN)")
                    df_prod_qtd_unif = (
                        df_perdas_unif_f[df_perdas_unif_f['Qtd_Limpa'] < 0]
                        .groupby(['Material_Codigo', 'Material_Nome'])['Qtd_Limpa']
                        .sum().abs().reset_index()
                        .sort_values(by='Qtd_Limpa', ascending=False)
                    )
                    df_prod_qtd_unif.insert(0, 'Posição', [f"{i+1}º" for i in range(len(df_prod_qtd_unif))])
                    df_top_qtd_unif = df_prod_qtd_unif.head(top_n_unif).copy()
                    df_top_qtd_unif['Qtd_Limpa'] = df_top_qtd_unif['Qtd_Limpa'].apply(lambda x: f"-{x:,.0f} un")
                    df_top_qtd_unif.rename(columns={'Material_Codigo': 'Material', 'Material_Nome': 'Descrição do Produto', 'Qtd_Limpa': 'Quantidade Perdida'}, inplace=True)
                    st.dataframe(df_top_qtd_unif, use_container_width=True, hide_index=True)
    
                with col_unif_val:
                    st.markdown(f"##### 💰 Top {top_n_unif} MAIORES por Valor (R$)")
                    df_prod_val_unif = (
                        df_perdas_unif_f[df_perdas_unif_f['Valor_Limpo'] < 0]
                        .groupby(['Material_Codigo', 'Material_Nome'])['Valor_Limpo']
                        .sum().abs().reset_index()
                        .sort_values(by='Valor_Limpo', ascending=False)
                    )
                    df_prod_val_unif.insert(0, 'Posição', [f"{i+1}º" for i in range(len(df_prod_val_unif))])
                    df_top_val_unif = df_prod_val_unif.head(top_n_unif).copy()
                    df_top_val_unif['Valor_Limpo'] = df_top_val_unif['Valor_Limpo'].apply(lambda x: f"R$ -{x:,.2f}")
                    df_top_val_unif.rename(columns={'Material_Codigo': 'Material', 'Material_Nome': 'Descrição do Produto', 'Valor_Limpo': 'Valor Perdido'}, inplace=True)
                    st.dataframe(df_top_val_unif, use_container_width=True, hide_index=True)
    
                st.markdown("---")
                st.markdown("#### ✅ Menores Perdas de Produtos (Unificado)")
                col_unif_qtd_min, col_unif_val_min = st.columns(2)
    
                with col_unif_qtd_min:
                    st.markdown(f"##### 📦 Top {top_n_unif} MENORES por Quantidade (UN)")
                    df_prod_qtd_unif_min = (
                        df_perdas_unif_f[df_perdas_unif_f['Qtd_Limpa'] < 0]
                        .groupby(['Material_Codigo', 'Material_Nome'])['Qtd_Limpa']
                        .sum().abs().reset_index()
                        .sort_values(by='Qtd_Limpa', ascending=True)
                    )
                    df_prod_qtd_unif_min.insert(0, 'Posição', [f"{i+1}º" for i in range(len(df_prod_qtd_unif_min))])
                    df_min_qtd_unif = df_prod_qtd_unif_min.head(top_n_unif).copy()
                    df_min_qtd_unif['Qtd_Limpa'] = df_min_qtd_unif['Qtd_Limpa'].apply(lambda x: f"-{x:,.0f} un")
                    df_min_qtd_unif.rename(columns={'Material_Codigo': 'Material', 'Material_Nome': 'Descrição do Produto', 'Qtd_Limpa': 'Quantidade Perdida'}, inplace=True)
                    st.dataframe(df_min_qtd_unif, use_container_width=True, hide_index=True)
    
                with col_unif_val_min:
                    st.markdown(f"##### 💰 Top {top_n_unif} MENORES por Valor (R$)")
                    df_prod_val_unif_min = (
                        df_perdas_unif_f[df_perdas_unif_f['Valor_Limpo'] < 0]
                        .groupby(['Material_Codigo', 'Material_Nome'])['Valor_Limpo']
                        .sum().abs().reset_index()
                        .sort_values(by='Valor_Limpo', ascending=True)
                    )
                    df_prod_val_unif_min.insert(0, 'Posição', [f"{i+1}º" for i in range(len(df_prod_val_unif_min))])
                    df_min_val_unif = df_prod_val_unif_min.head(top_n_unif).copy()
                    df_min_val_unif['Valor_Limpo'] = df_min_val_unif['Valor_Limpo'].apply(lambda x: f"R$ -{x:,.2f}")
                    df_min_val_unif.rename(columns={'Material_Codigo': 'Material', 'Material_Nome': 'Descrição do Produto', 'Valor_Limpo': 'Valor Perdido'}, inplace=True)
                    st.dataframe(df_min_val_unif, use_container_width=True, hide_index=True)
    
            with tab_por_loja:
                lojas_existentes = sorted([x for x in df_perdas_prod['Loja_Nome'].unique() if x])
                if lojas_existentes:
                    centro_selecionado = st.selectbox("Selecione o Centro (Loja):", options=lojas_existentes, key="f_prod_loja_centro")
                    
                    with st.expander("🔍 Filtro Local: Marcas, Meses e Top N da Loja Selecionada"):
                        col_fl1, col_fl2, col_fl3 = st.columns(3)
                        with col_fl1:
                            marcas_loja_opts = sorted(df_perdas_prod[df_perdas_prod['Loja_Nome'] == centro_selecionado]['Marca_Nome'].unique())
                            m_loja_sel = st.multiselect("Filtrar Marcas:", options=marcas_loja_opts, default=marcas_loja_opts, key="f_prod_loja_marca")
                        with col_fl2:
                            meses_loja_opts = sorted(df_perdas_prod[df_perdas_prod['Loja_Nome'] == centro_selecionado]['Mes_Ano'].unique())
                            mes_loja_sel = st.multiselect("Filtrar Mês/Ano:", options=meses_loja_opts, default=meses_loja_opts, key="f_prod_loja_mes")
                        with col_fl3:
                            top_n_loja = st.slider("Selecione o Top (Produtos por Loja):", min_value=10, max_value=100, value=10, step=5, key="top_prod_loja")
    
                    df_loja_prod = df_perdas_prod[
                        (df_perdas_prod['Loja_Nome'] == centro_selecionado) &
                        (df_perdas_prod['Marca_Nome'].isin(m_loja_sel)) &
                        (df_perdas_prod['Mes_Ano'].isin(mes_loja_sel))
                    ]
                    
                    st.markdown(f"#### 🚨 Maiores Perdas em `{centro_selecionado}`")
                    col_loja_qtd, col_loja_val = st.columns(2)
                    
                    with col_loja_qtd:
                        st.markdown(f"##### 📦 Top {top_n_loja} MAIORES por Qtd em `{centro_selecionado}`")
                        df_prod_qtd_loja = (
                            df_loja_prod[df_loja_prod['Qtd_Limpa'] < 0]
                            .groupby(['Material_Codigo', 'Material_Nome'])['Qtd_Limpa']
                            .sum().abs().reset_index()
                            .sort_values(by='Qtd_Limpa', ascending=False)
                        )
                        df_prod_qtd_loja.insert(0, 'Posição', [f"{i+1}º" for i in range(len(df_prod_qtd_loja))])
                        df_top_qtd_loja = df_prod_qtd_loja.head(top_n_loja).copy()
                        df_top_qtd_loja['Qtd_Limpa'] = df_top_qtd_loja['Qtd_Limpa'].apply(lambda x: f"-{x:,.0f} un")
                        df_top_qtd_loja.rename(columns={'Material_Codigo': 'Material', 'Material_Nome': 'Descrição do Produto', 'Qtd_Limpa': 'Quantidade'}, inplace=True)
                        st.dataframe(df_top_qtd_loja, use_container_width=True, hide_index=True)
    
                    with col_loja_val:
                        st.markdown(f"##### 💰 Top {top_n_loja} MAIORES por Valor em `{centro_selecionado}`")
                        df_prod_val_loja = (
                            df_loja_prod[df_loja_prod['Valor_Limpo'] < 0]
                            .groupby(['Material_Codigo', 'Material_Nome'])['Valor_Limpo']
                            .sum().abs().reset_index()
                            .sort_values(by='Valor_Limpo', ascending=False)
                        )
                        df_prod_val_loja.insert(0, 'Posição', [f"{i+1}º" for i in range(len(df_prod_val_loja))])
                        df_top_val_loja = df_prod_val_loja.head(top_n_loja).copy()
                        df_top_val_loja['Valor_Limpo'] = df_top_val_loja['Valor_Limpo'].apply(lambda x: f"R$ -{x:,.2f}")
                        df_top_val_loja.rename(columns={'Material_Codigo': 'Material', 'Material_Nome': 'Descrição do Produto', 'Valor_Limpo': 'Valor'}, inplace=True)
                        st.dataframe(df_top_val_loja, use_container_width=True, hide_index=True)
    
                    st.markdown("---")
                    st.markdown(f"#### ✅ Menores Perdas em `{centro_selecionado}`")
                    col_loja_qtd_min, col_loja_val_min = st.columns(2)
    
                    with col_loja_qtd_min:
                        st.markdown(f"##### 📦 Top {top_n_loja} MENORES por Qtd em `{centro_selecionado}`")
                        df_prod_qtd_loja_min = (
                            df_loja_prod[df_loja_prod['Qtd_Limpa'] < 0]
                            .groupby(['Material_Codigo', 'Material_Nome'])['Qtd_Limpa']
                            .sum().abs().reset_index()
                            .sort_values(by='Qtd_Limpa', ascending=True)
                        )
                        df_prod_qtd_loja_min.insert(0, 'Posição', [f"{i+1}º" for i in range(len(df_prod_qtd_loja_min))])
                        df_min_qtd_loja = df_prod_qtd_loja_min.head(top_n_loja).copy()
                        df_min_qtd_loja['Qtd_Limpa'] = df_min_qtd_loja['Qtd_Limpa'].apply(lambda x: f"-{x:,.0f} un")
                        df_min_qtd_loja.rename(columns={'Material_Codigo': 'Material', 'Material_Nome': 'Descrição do Produto', 'Qtd_Limpa': 'Quantidade'}, inplace=True)
                        st.dataframe(df_min_qtd_loja, use_container_width=True, hide_index=True)
    
                    with col_loja_val_min:
                        st.markdown(f"##### 💰 Top {top_n_loja} MENORES por Valor em `{centro_selecionado}`")
                        df_prod_val_loja_min = (
                            df_loja_prod[df_loja_prod['Valor_Limpo'] < 0]
                            .groupby(['Material_Codigo', 'Material_Nome'])['Valor_Limpo']
                            .sum().abs().reset_index()
                            .sort_values(by='Valor_Limpo', ascending=True)
                        )
                        df_prod_val_loja_min.insert(0, 'Posição', [f"{i+1}º" for i in range(len(df_prod_val_loja_min))])
                        df_min_val_loja = df_prod_val_loja_min.head(top_n_loja).copy()
                        df_min_val_loja['Valor_Limpo'] = df_min_val_loja['Valor_Limpo'].apply(lambda x: f"R$ -{x:,.2f}")
                        df_min_val_loja.rename(columns={'Material_Codigo': 'Material', 'Material_Nome': 'Descrição do Produto', 'Valor_Limpo': 'Valor'}, inplace=True)
                        st.dataframe(df_min_val_loja, use_container_width=True, hide_index=True)
                else:
                    st.info("Nenhum registro de perda encontrado para os filtros selecionados.")
    
            st.markdown("<br>", unsafe_allow_html=True)

        # COMPARATIVO REGIONAL
        if perfil_usuario in ["Administrador", "Controladoria"]:
            st.subheader("🗺️ Comparativo por Divisão Regional (Regional 1 vs Regional 2)")
            
            with st.expander("🔍 Filtro Local: Comparativo Regional"):
                col_fr1, col_fr2 = st.columns(2)
                with col_fr1:
                    m_reg_sel = st.multiselect("Filtrar Marcas:", options=sorted(df_filtered['Marca_Nome'].unique()), default=sorted(df_filtered['Marca_Nome'].unique()), key="f_reg_marcas")
                with col_fr2:
                    mes_reg_sel = st.multiselect("Filtrar Mês/Ano:", options=sorted(df_filtered['Mes_Ano'].unique()), default=sorted(df_filtered['Mes_Ano'].unique()), key="f_reg_mes")
            
            df_reg_filtered = df_filtered[
                (df_filtered['Marca_Nome'].isin(m_reg_sel)) &
                (df_filtered['Mes_Ano'].isin(mes_reg_sel))
            ]

            df_reg_comp = (
                df_reg_filtered[df_reg_filtered['Valor_Limpo'] < 0]
                .groupby('Regional_Nome')
                .agg({'Qtd_Limpa': lambda x: abs(x.sum()), 'Valor_Limpo': lambda x: abs(x.sum())})
                .reset_index()
                .sort_values(by='Valor_Limpo', ascending=False)
            )
            if not df_reg_comp.empty:
                df_reg_comp['Texto_Valor'] = df_reg_comp['Valor_Limpo'].apply(lambda x: f"-R$ {x:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
                total_reg_val = df_reg_comp['Valor_Limpo'].sum()

                fig_reg_comp = px.bar(
                    df_reg_comp,
                    x='Regional_Nome',
                    y='Valor_Limpo',
                    text='Texto_Valor',
                    color='Regional_Nome',
                    title=f"Comparativo por Divisão Regional — Total: -R$ {total_reg_val:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."),
                    color_discrete_map={'Regional 1': '#4ba3e3', 'Regional 2': '#ff7f0e'},
                    labels={'Valor_Limpo': 'Perda (R$)', 'Regional_Nome': 'Divisão Regional'}
                )
                fig_reg_comp.update_traces(textposition='inside')
                fig_reg_comp.update_layout(template="plotly_dark", paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)", showlegend=False)
                st.plotly_chart(fig_reg_comp, use_container_width=True)
                st.markdown("<br>", unsafe_allow_html=True)

        # RANKING POR CENTRO
        if perfil_usuario in ["Administrador", "Regional 1", "Regional 2", "Controladoria", "Gerente de produtos 1", "Gerente de produtos 2"]: 
            st.subheader("🏬 Análise e Ranking por Centro")

            with st.expander("🔍 Filtro Local: Perdas por Centro"):
                if perfil_usuario == "Administrador":
                    col_fc1, col_fc2, col_fc3, col_fc4 = st.columns(4)
                    with col_fc4:
                        top_n_centros = st.slider("Selecione o Top (Centros):", min_value=10, max_value=100, value=10, step=5, key="top_centros_adm")
                else:
                    col_fc1, col_fc2, col_fc3 = st.columns(3)
                    top_n_centros = 10

                with col_fc1:
                    c_lojas_sel = st.multiselect("Filtrar Centros Específicos:", options=sorted(df_filtered['Loja_Nome'].unique()), default=sorted(df_filtered['Loja_Nome'].unique()), key="f_centro_lojas")
                with col_fc2:
                    c_marcas_sel = st.multiselect("Filtrar Marcas:", options=sorted(df_filtered['Marca_Nome'].unique()), default=sorted(df_filtered['Marca_Nome'].unique()), key="f_centro_marcas")
                with col_fc3:
                    c_meses_sel = st.multiselect("Filtrar Mês/Ano:", options=sorted(df_filtered['Mes_Ano'].unique()), default=sorted(df_filtered['Mes_Ano'].unique()), key="f_centro_meses")

            df_centros_local = df_filtered[
                (df_filtered['Loja_Nome'].isin(c_lojas_sel)) &
                (df_filtered['Marca_Nome'].isin(c_marcas_sel)) &
                (df_filtered['Mes_Ano'].isin(c_meses_sel))
            ]

            graf_col1, graf_col2 = st.columns(2)

            with graf_col1:
                df_qtd_lojas = (
                    df_centros_local[df_centros_local['Qtd_Limpa'] < 0]
                    .groupby('Loja_Nome')['Qtd_Limpa']
                    .sum().abs().reset_index()
                    .sort_values(by='Qtd_Limpa', ascending=False)
                )
                if not df_qtd_lojas.empty:
                    df_qtd_lojas['Texto_Qtd'] = df_qtd_lojas['Qtd_Limpa'].apply(lambda x: f"-{x:,.0f} un")
                    total_centros_qtd = df_qtd_lojas['Qtd_Limpa'].sum()

                    fig_qtd_lojas = px.bar(
                        df_qtd_lojas,
                        x='Loja_Nome',
                        y='Qtd_Limpa',
                        text='Texto_Qtd',
                        title=f"Perda por Centro (Qtd) — Total: -{total_centros_qtd:,.0f} un".replace(",", "."),
                        labels={'Qtd_Limpa': 'Perda (Qtd)', 'Loja_Nome': 'Centro'}
                    )
                    fig_qtd_lojas.update_traces(marker_color='#4ba3e3', textposition='inside')
                    fig_qtd_lojas.update_layout(template="plotly_dark", paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)")
                    st.plotly_chart(fig_qtd_lojas, use_container_width=True)

            with graf_col2:
                df_lojas = (
                    df_centros_local[df_centros_local['Valor_Limpo'] < 0]
                    .groupby('Loja_Nome')['Valor_Limpo']
                    .sum().abs().reset_index()
                    .sort_values(by='Valor_Limpo', ascending=False)
                )
                if not df_lojas.empty:
                    df_lojas['Texto_Valor'] = df_lojas['Valor_Limpo'].apply(lambda x: f"-R$ {x:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
                    total_centros_val = df_lojas['Valor_Limpo'].sum()

                    fig_lojas = px.bar(
                        df_lojas,
                        x='Loja_Nome',
                        y='Valor_Limpo',
                        text='Texto_Valor',
                        title=f"Perda por Centro (R$) — Total: -R$ {total_centros_val:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."),
                        labels={'Valor_Limpo': 'Perda (R$)', 'Loja_Nome': 'Centro'}
                    )
                    fig_lojas.update_traces(marker_color='#70bbfd', textposition='inside')
                    fig_lojas.update_layout(template="plotly_dark", paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)")
                    st.plotly_chart(fig_lojas, use_container_width=True)

            col_tb_c1, col_tb_c2 = st.columns(2)

            with col_tb_c1:
                st.markdown(f"##### 🚨 Ranking: Top {top_n_centros} Centros com MAIOR Perda")
                df_centros_completo = (
                    df_centros_local[(df_centros_local['Valor_Limpo'] < 0) | (df_centros_local['Qtd_Limpa'] < 0)]
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

                df_top_centros = df_centros_completo.head(top_n_centros).copy()
                df_top_centros['Perda (Qtd)'] = df_top_centros['Perda (Qtd)'].apply(lambda x: f"-{x:,.0f} un")
                df_top_centros['Perda (R$)'] = df_top_centros['Perda (R$)'].apply(lambda x: f"R$ -{x:,.2f}")
                st.dataframe(df_top_centros, use_container_width=True, hide_index=True)

            with col_tb_c2:
                st.markdown(f"##### ✅ Ranking: Top {top_n_centros} Centros com MENOR Perda")
                df_centros_completo_min = (
                    df_centros_local[(df_centros_local['Valor_Limpo'] < 0) | (df_centros_local['Qtd_Limpa'] < 0)]
                    .groupby(['Loja_Nome', 'Regional_Nome'])
                    .agg({
                        'Qtd_Limpa': lambda x: abs(x[x < 0].sum()),
                        'Valor_Limpo': lambda x: abs(x[x < 0].sum())
                    })
                    .reset_index()
                    .sort_values(by='Valor_Limpo', ascending=True)
                )
                df_centros_completo_min.insert(0, 'Posição', [f"{i+1}º" for i in range(len(df_centros_completo_min))])
                df_centros_completo_min = df_centros_completo_min[['Posição', 'Loja_Nome', 'Regional_Nome', 'Qtd_Limpa', 'Valor_Limpo']]
                df_centros_completo_min.rename(columns={'Loja_Nome': 'Centro', 'Regional_Nome': 'Divisão Regional', 'Qtd_Limpa': 'Perda (Qtd)', 'Valor_Limpo': 'Perda (R$)'}, inplace=True)

                df_min_centros = df_centros_completo_min.head(top_n_centros).copy()
                df_min_centros['Perda (Qtd)'] = df_min_centros['Perda (Qtd)'].apply(lambda x: f"-{x:,.0f} un")
                df_min_centros['Perda (R$)'] = df_min_centros['Perda (R$)'].apply(lambda x: f"R$ -{x:,.2f}")
                st.dataframe(df_min_centros, use_container_width=True, hide_index=True)

        st.markdown("<br>", unsafe_allow_html=True)

        # PERDAS POR MARCA
        st.subheader("🏷️ Análise e Ranking por Marca")

        with st.expander("🔍 Filtro Local: Perdas por Marca"):
            if perfil_usuario in ["Gerente", "Líder de Loja"]:
                col_fm1, col_fm2, col_fm3 = st.columns(3)
                with col_fm1:
                    m_marcas_sel = st.multiselect("Filtrar Marcas Específicas:", options=sorted(df_filtered['Marca_Nome'].unique()), default=sorted(df_filtered['Marca_Nome'].unique()), key="f_marca_marcas")
                with col_fm2:
                    m_meses_sel = st.multiselect("Filtrar Mês/Ano:", options=sorted(df_filtered['Mes_Ano'].unique()), default=sorted(df_filtered['Mes_Ano'].unique()), key="f_marca_meses")
                with col_fm3:
                    top_n_marcas = st.slider("Selecione o Top (Marcas):", min_value=10, max_value=100, value=10, step=5, key="top_marcas")
                m_centros_sel = sorted(df_filtered['Loja_Nome'].unique())
            else:
                col_fm1, col_fm2, col_fm3, col_fm4 = st.columns(4)
                with col_fm1:
                    m_marcas_sel = st.multiselect("Filtrar Marcas Específicas:", options=sorted(df_filtered['Marca_Nome'].unique()), default=sorted(df_filtered['Marca_Nome'].unique()), key="f_marca_marcas")
                with col_fm2:
                    m_centros_sel = st.multiselect("Filtrar Centros/Lojas:", options=sorted(df_filtered['Loja_Nome'].unique()), default=sorted(df_filtered['Loja_Nome'].unique()), key="f_marca_centros")
                with col_fm3:
                    m_meses_sel = st.multiselect("Filtrar Mês/Ano:", options=sorted(df_filtered['Mes_Ano'].unique()), default=sorted(df_filtered['Mes_Ano'].unique()), key="f_marca_meses")
                with col_fm4:
                    top_n_marcas = st.slider("Selecione o Top (Marcas):", min_value=10, max_value=100, value=10, step=5, key="top_marcas")

        df_marcas_local = df_filtered[
            (df_filtered['Marca_Nome'].isin(m_marcas_sel)) &
            (df_filtered['Loja_Nome'].isin(m_centros_sel)) &
            (df_filtered['Mes_Ano'].isin(m_meses_sel))
        ]

        col_tb_m1, col_tb_m2 = st.columns(2)

        with col_tb_m1:
            st.markdown(f"##### 🚨 Ranking: Top {top_n_marcas} Marcas com MAIOR Perda")
            df_marcas_completo = (
                df_marcas_local[(df_marcas_local['Valor_Limpo'] < 0) | (df_marcas_local['Qtd_Limpa'] < 0)]
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

            df_top_marcas = df_marcas_completo.head(top_n_marcas).copy()
            df_top_marcas['Perda (Qtd)'] = df_top_marcas['Perda (Qtd)'].apply(lambda x: f"-{x:,.0f} un")
            df_top_marcas['Perda (R$)'] = df_top_marcas['Perda (R$)'].apply(lambda x: f"R$ -{x:,.2f}")
            st.dataframe(df_top_marcas, use_container_width=True, hide_index=True)

        with col_tb_m2:
            st.markdown(f"##### ✅ Ranking: Top {top_n_marcas} Marcas com MENOR Perda")
            df_marcas_completo_min = (
                df_marcas_local[(df_marcas_local['Valor_Limpo'] < 0) | (df_marcas_local['Qtd_Limpa'] < 0)]
                .groupby('Marca_Nome')
                .agg({
                    'Qtd_Limpa': lambda x: abs(x[x < 0].sum()),
                    'Valor_Limpo': lambda x: abs(x[x < 0].sum())
                })
                .reset_index()
                .sort_values(by='Valor_Limpo', ascending=True)
            )
            df_marcas_completo_min.insert(0, 'Posição', [f"{i+1}º" for i in range(len(df_marcas_completo_min))])
            df_marcas_completo_min = df_marcas_completo_min[['Posição', 'Marca_Nome', 'Qtd_Limpa', 'Valor_Limpo']]
            df_marcas_completo_min.rename(columns={'Marca_Nome': 'Marca', 'Qtd_Limpa': 'Perda (Qtd)', 'Valor_Limpo': 'Perda (R$)'}, inplace=True)

            df_min_marcas = df_marcas_completo_min.head(top_n_marcas).copy()
            df_min_marcas['Perda (Qtd)'] = df_min_marcas['Perda (Qtd)'].apply(lambda x: f"-{x:,.0f} un")
            df_min_marcas['Perda (R$)'] = df_min_marcas['Perda (R$)'].apply(lambda x: f"R$ -{x:,.2f}")
            st.dataframe(df_min_marcas, use_container_width=True, hide_index=True)

        st.markdown("<br>", unsafe_allow_html=True)
        marca_col1, marca_col2 = st.columns(2)

        df_perdas_marcas = df_marcas_local[(df_marcas_local['Valor_Limpo'] < 0) | (df_marcas_local['Qtd_Limpa'] < 0)]

        with marca_col1:
            st.markdown(f"##### 📦 Top {top_n_marcas} Perdas por Marca - Tendência (Qtd)")
            df_marca_qtd = (
                df_perdas_marcas[df_perdas_marcas['Qtd_Limpa'] < 0]
                .groupby('Marca_Nome')['Qtd_Limpa']
                .sum().abs().reset_index()
                .sort_values(by='Qtd_Limpa', ascending=False)
                .head(top_n_marcas)
            )
            if not df_marca_qtd.empty:
                df_marca_qtd['Texto_Qtd'] = df_marca_qtd['Qtd_Limpa'].apply(lambda x: f"-{x:,.0f} un")
                total_marcas_qtd = df_marca_qtd['Qtd_Limpa'].sum()

                fig_marca_qtd = px.line(
                    df_marca_qtd,
                    x='Marca_Nome',
                    y='Qtd_Limpa',
                    text='Texto_Qtd',
                    markers=True,
                    title=f"Perdas por Marca (Top {top_n_marcas} - Qtd) — Total Top {top_n_marcas}: -{total_marcas_qtd:,.0f} un".replace(",", "."),
                    labels={'Qtd_Limpa': 'Perda (Qtd)', 'Marca_Nome': 'Marca'}
                )
                fig_marca_qtd.update_traces(line_color='#ff7f0e', line_width=3, marker_size=7, textposition='top center')
                fig_marca_qtd.update_layout(template="plotly_dark", paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)", xaxis_tickangle=-45)
                st.plotly_chart(fig_marca_qtd, use_container_width=True)

        with marca_col2:
            st.markdown(f"##### 🏷️ Top {top_n_marcas} Perdas por Marca - Tendência (R$)")
            df_marca_rs = (
                df_perdas_marcas[df_perdas_marcas['Valor_Limpo'] < 0]
                .groupby('Marca_Nome')['Valor_Limpo']
                .sum().abs().reset_index()
                .sort_values(by='Valor_Limpo', ascending=False)
                .head(top_n_marcas)
            )
            if not df_marca_rs.empty:
                df_marca_rs['Texto_RS'] = df_marca_rs['Valor_Limpo'].apply(lambda x: f"-R$ {x:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
                total_marcas_val = df_marca_rs['Valor_Limpo'].sum()

                fig_marca_rs = px.line(
                    df_marca_rs,
                    x='Marca_Nome',
                    y='Valor_Limpo',
                    text='Texto_RS',
                    markers=True,
                    title=f"Perdas por Marca (Top {top_n_marcas} - R$) — Total Top {top_n_marcas}: -R$ {total_marcas_val:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."),
                    labels={'Valor_Limpo': 'Perda (R$)', 'Marca_Nome': 'Marca'}
                )
                fig_marca_rs.update_traces(line_color='#4ba3e3', line_width=3, marker_size=7, textposition='top center')
                fig_marca_rs.update_layout(template="plotly_dark", paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)", xaxis_tickangle=-45)
                st.plotly_chart(fig_marca_rs, use_container_width=True)

    except requests.exceptions.HTTPError as http_err:
        st.error(f"⚠️ Erro HTTP ao baixar do SharePoint ({http_err.response.status_code}). Verifique as permissões do link.")
    except Exception as e:
        st.error(f"Erro ao carregar os dados do arquivo Excel na nuvem: {e}")

# --- FLUXO PRINCIPAL DA APLICAÇÃO ---
if not st.session_state["logado"]:
    renderizar_tela_login()

elif st.session_state["troca_obrigatoria"]:
    renderizar_tela_troca_obrigatoria()

else:
    usuarios_db, removidos_set, historico_remocoes, historico_resets = carregar_dados_db()
    usr_atual = st.session_state["usuario_atual"]

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
                usuarios_dict, removidos_set, historico_remocoes, historico_resets = carregar_dados_db()
                valida, msg_validacao = validar_complexidade_senha(nova_senha_sb)
                
                if gerar_hash(senha_antiga_sb) != usuarios_dict[usr_atual]["senha"]:
                    st.error("Senha atual incorreta.")
                elif not valida:
                    st.error(msg_validacao)
                elif nova_senha_sb != confirma_sb:
                    st.error("Senhas não conferem.")
                else:
                    sucesso, msg = atualizar_senha_com_historico(
                        usr_atual, nova_senha_sb, usuarios_dict, removidos_set, historico_remocoes, historico_resets
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
        aba_dash, aba_lojas, aba_admin = st.tabs(["📊 Dashboard Geral", "🏬 Gestão de Lojas e Mudanças", "⚙️ Painel Admin"])
        with aba_dash:
            renderizar_dashboard()
        with aba_lojas:
            renderizar_aba_gestao_lojas()
        with aba_admin:
            renderizar_aba_admin()
    else:
        aba_dash, aba_lojas = st.tabs(["📊 Dashboard Geral", "🏬 Gestão de Lojas e Mudanças"])
        with aba_dash:
            renderizar_dashboard()
        with aba_lojas:
            renderizar_aba_gestao_lojas()
