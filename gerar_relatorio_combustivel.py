
import pandas as pd
import os
from openpyxl import Workbook
from openpyxl.styles import PatternFill
from openpyxl.utils.dataframe import dataframe_to_rows

# Funções auxiliares
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

def gerar_analise_mensal(df):
    analise = df.groupby('MES_ANO').agg(
        QTD_TRANSACOES=('CODIGO TRANSACAO', 'count'),
        TOTAL_VALOR=('VALOR EMISSAO', 'sum'),
        TOTAL_DIESEL=('LITROS', lambda x: x[df.loc[x.index, 'COMBUSTIVEL_TIPO'] == 'DIESEL'].sum()),
        TOTAL_GASOLINA=('LITROS', lambda x: x[df.loc[x.index, 'COMBUSTIVEL_TIPO'] == 'GASOLINA'].sum()),
        TOTAL_ARLA=('LITROS', lambda x: x[df.loc[x.index, 'COMBUSTIVEL_TIPO'] == 'ARLA'].sum())
    ).reset_index().sort_values(by='MES_ANO')

    if len(analise) >= 2:
        last = analise.iloc[-1, 1:]
        prev = analise.iloc[-2, 1:]
        diff = ['DIFERENÇA'] + (last - prev).tolist()
        perc = ['%'] + (((last - prev) / prev) * 100).tolist()
        analise = pd.concat([analise, pd.DataFrame([diff, perc], columns=analise.columns)], ignore_index=True)
    return analise

def gerar_resumo_empresa_mes(df):
    return df.groupby(['EMPRESA', 'MES_ANO']).agg(
        TOTAL_REAIS=('VALOR EMISSAO', 'sum'),
        DIESEL=('LITROS', lambda x: x[df.loc[x.index, 'COMBUSTIVEL_TIPO'] == 'DIESEL'].sum()),
        GASOLINA=('LITROS', lambda x: x[df.loc[x.index, 'COMBUSTIVEL_TIPO'] == 'GASOLINA'].sum()),
        ARLA=('LITROS', lambda x: x[df.loc[x.index, 'COMBUSTIVEL_TIPO'] == 'ARLA'].sum()),
        TRANSACOES=('CODIGO TRANSACAO', 'count')
    ).reset_index()

def gerar_relatorio(pasta_entrada, caminho_saida):
    arquivos = [f for f in os.listdir(pasta_entrada) if f.endswith(".xlsx")]
    df_list = []
    for arquivo in arquivos:
        caminho = os.path.join(pasta_entrada, arquivo)
        df = pd.read_excel(caminho)
        df['EMPRESA_ARQUIVO'] = os.path.splitext(arquivo)[0]
        df_list.append(df)
    df_total = pd.concat(df_list, ignore_index=True)

    df_total['EMPRESA'] = df_total['NOME REDUZIDO'].apply(identificar_empresa)
    df_total['DATA TRANSACAO'] = pd.to_datetime(df_total['DATA TRANSACAO'], errors='coerce')
    df_total['MES_ANO'] = df_total['DATA TRANSACAO'].dt.to_period("M").astype(str)
    df_total['DIA'] = df_total['DATA TRANSACAO'].dt.day
    df_total['COMBUSTIVEL_TIPO'] = df_total['TIPO COMBUSTIVEL'].apply(classificar_combustivel)

    df_q1 = df_total[df_total['DIA'] <= 15]
    df_q2 = df_total[df_total['DIA'] > 15]

    analise_mensal = gerar_analise_mensal(df_total)
    analise_q1 = gerar_analise_mensal(df_q1)
    analise_q2 = gerar_analise_mensal(df_q2)

    resumo_geral = gerar_resumo_empresa_mes(df_total)
    resumo_q1 = gerar_resumo_empresa_mes(df_q1)
    resumo_q2 = gerar_resumo_empresa_mes(df_q2)

    with pd.ExcelWriter(caminho_saida, engine='openpyxl') as writer:
        analise_mensal.to_excel(writer, sheet_name='ANÁLISE MENSAL - COMBUSTÍVEL', index=False)
        analise_q1.to_excel(writer, sheet_name='ANÁLISE 1ª QUINZENA', index=False)
        analise_q2.to_excel(writer, sheet_name='ANÁLISE 2ª QUINZENA', index=False)
        resumo_geral.to_excel(writer, sheet_name='RESUMO GERAL', index=False)
        resumo_q1.to_excel(writer, sheet_name='1ª QUINZENA', index=False)
        resumo_q2.to_excel(writer, sheet_name='2ª QUINZENA', index=False)

        for empresa in df_total['EMPRESA'].unique():
            df_emp = df_total[df_total['EMPRESA'] == empresa].copy()
            df_emp['MES_ANO'] = df_emp['DATA TRANSACAO'].dt.to_period("M").astype(str)
            df_emp = df_emp[['DATA TRANSACAO', 'MES_ANO', 'PLACA', 'TIPO COMBUSTIVEL', 'LITROS', 'VALOR EMISSAO']]
            df_emp.to_excel(writer, sheet_name=empresa, index=False)

if __name__ == "__main__":
    pasta_entrada = r"C:\ConsumoCombustiveis\MENSAL"
    saida = os.path.join(pasta_entrada, "RELATORIO_COMBUSTIVEL_MENSAL.xlsx")
    gerar_relatorio(pasta_entrada, saida)
