
import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime

# Funções de identificação e classificação
def identificar_empresa(nome):
    nome = str(nome).upper()
    if "OCEAN" in nome:
        return "OCEAN"
    elif "AV09" in nome:
        return "AV09"
    elif "SULFOODS" in nome or "LEGOUR" in nome:
        return "SULFOODS + LEGOUR"
    return "OUTROS"

def classificar_combustivel(tipo):
    if pd.isna(tipo):
        return 'OUTROS'
    tipo = tipo.upper()
    if 'DIESEL' in tipo:
        return 'DIESEL'
    elif 'GASOLINA' in tipo or 'ETANOL' in tipo:
        return 'GASOLINA'
    elif 'ARLA' in tipo:
        return 'ARLA'
    return 'OUTROS'

# Configuração do Streamlit
st.set_page_config(page_title="Painel de Combustível", layout="wide")
st.title("📊 Painel de Consumo de Combustível (Nuvem)")

# Upload múltiplo
uploaded_files = st.file_uploader("📂 Envie os arquivos .xlsx das empresas", type="xlsx", accept_multiple_files=True)

if uploaded_files:
    dfs = []
    for uploaded_file in uploaded_files:
        df = pd.read_excel(uploaded_file)
        df['EMPRESA_ARQUIVO'] = uploaded_file.name.replace('.xlsx', '')
        dfs.append(df)

    df = pd.concat(dfs, ignore_index=True)

    # Pré-processamento
    df['EMPRESA'] = df['NOME REDUZIDO'].apply(identificar_empresa)
    df['DATA TRANSACAO'] = pd.to_datetime(df['DATA TRANSACAO'], errors='coerce')
    df['MES_ANO'] = df['DATA TRANSACAO'].dt.to_period("M").astype(str)
    df['COMBUSTIVEL_TIPO'] = df['TIPO COMBUSTIVEL'].apply(classificar_combustivel)

    # Filtros
    empresas = sorted(df['EMPRESA'].unique())
    meses = sorted(df['MES_ANO'].unique())
    tipos_comb = ['TODOS'] + ['DIESEL', 'GASOLINA', 'ARLA']

    col1, col2, col3 = st.columns(3)
    empresa_sel = col1.multiselect("Empresa", empresas, default=empresas)
    mes_sel = col2.multiselect("Mês/Ano", meses, default=meses)
    combustivel_sel = col3.selectbox("Tipo de Combustível", tipos_comb)

    df_filtrado = df[df['EMPRESA'].isin(empresa_sel) & df['MES_ANO'].isin(mes_sel)]
    if combustivel_sel != 'TODOS':
        df_filtrado = df_filtrado[df_filtrado['COMBUSTIVEL_TIPO'] == combustivel_sel]

    # Métricas principais
    st.markdown("### 🔍 Visão Geral Filtrada")
    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Total Transações", len(df_filtrado))
    col2.metric("Total R$", f"{df_filtrado['VALOR EMISSAO'].sum():,.2f}")
    col3.metric("Litros Diesel", round(df_filtrado[df_filtrado['COMBUSTIVEL_TIPO'] == 'DIESEL']['LITROS'].sum(), 2))
    col4.metric("Litros Gasolina", round(df_filtrado[df_filtrado['COMBUSTIVEL_TIPO'] == 'GASOLINA']['LITROS'].sum(), 2))

    # Gráfico de consumo mensal
    st.markdown("### 📈 Consumo Mensal (Litros)")
    graf = df_filtrado.groupby(['MES_ANO', 'COMBUSTIVEL_TIPO'])['LITROS'].sum().reset_index()
    graf_pivot = graf.pivot(index='MES_ANO', columns='COMBUSTIVEL_TIPO', values='LITROS').fillna(0)
    st.bar_chart(graf_pivot)

    # Dados detalhados
    with st.expander("🔎 Ver dados detalhados"):
        st.dataframe(df_filtrado)

    # Exportação CSV
    st.download_button("📥 Baixar dados filtrados", df_filtrado.to_csv(index=False).encode(), "dados_filtrados.csv", "text/csv")
else:
    st.info("Envie os arquivos para iniciar a análise.")
