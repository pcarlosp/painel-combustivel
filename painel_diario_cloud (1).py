
import streamlit as st
import pandas as pd
from io import BytesIO

# Configuração da página
st.set_page_config(page_title="Painel Diário de Consumo", layout="wide")
st.title("📆 Painel Diário de Consumo de Combustível")

# Upload do arquivo
uploaded_file = st.file_uploader("📂 Envie o arquivo de relatório diário (.xlsx)", type="xlsx")

if uploaded_file:
    df = pd.read_excel(uploaded_file)

    # Tratamento inicial
    df['DATA TRANSACAO'] = pd.to_datetime(df['DATA TRANSACAO'], errors='coerce')
    df['DIA'] = df['DATA TRANSACAO'].dt.date
    df['HODOMETRO'] = pd.to_numeric(df['HODOMETRO OU HORIMETRO'], errors='coerce')
    df['KM_RODADOS'] = pd.to_numeric(df['KM RODADOS OU HORAS TRABALHADAS'], errors='coerce')
    df['LITROS_ESTIMADO'] = df['KM_RODADOS'] / pd.to_numeric(df['KM/LITRO OU LITROS/HORA'], errors='coerce')
    df['MEDIA_CALCULADA'] = df['KM_RODADOS'] / df['LITROS_ESTIMADO']

    # Filtros interativos
    placas = df['PLACA'].dropna().unique()
    placas_sel = st.multiselect("Filtrar por placa", sorted(placas), default=sorted(placas))

    datas = df['DIA'].dropna().unique()
    datas_sel = st.multiselect("Filtrar por data", sorted(datas), default=sorted(datas))

    df_filt = df[df['PLACA'].isin(placas_sel) & df['DIA'].isin(datas_sel)]

    st.markdown("### 🔍 Visão Geral")
    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Total Abastecimentos", len(df_filt))
    col2.metric("Total R$", f"{df_filt['VALOR EMISSAO'].sum():,.2f}")
    col3.metric("KM Rodados", f"{df_filt['KM_RODADOS'].sum():,.1f}")
    col4.metric("Média Geral KM/L", f"{df_filt['MEDIA_CALCULADA'].mean():.2f}")

    st.markdown("### 📋 Detalhes por Placa e Dia")
    resumo = df_filt.groupby(['DIA', 'PLACA']).agg(
        TOTAL_REAIS=('VALOR EMISSAO', 'sum'),
        KM_RODADOS=('KM_RODADOS', 'sum'),
        LITROS_ESTIMADO=('LITROS_ESTIMADO', 'sum'),
        MEDIA_KM_L=('MEDIA_CALCULADA', 'mean'),
        ABASTECIMENTOS=('CODIGO TRANSACAO', 'count')
    ).reset_index()

    st.dataframe(resumo.style.background_gradient(cmap="RdYlGn", subset=['MEDIA_KM_L']))

    st.download_button("📥 Baixar resumo diário", resumo.to_csv(index=False).encode(), "resumo_diario.csv", "text/csv")
else:
    st.info("Envie o arquivo de relatório para começar.")
