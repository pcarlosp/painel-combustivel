
import os
import pandas as pd
from datetime import datetime
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import numbers

# CONFIGURAÇÕES
CAMINHO_PLANILHAS = r'C:\ConsumoCombustiveis\Planilhas'
ARQUIVO_MODELO = 'modelo a ser seguido.xlsx'

# Data de hoje
hoje = datetime.today().date()

# Coletar arquivos válidos
arquivos = [
    f for f in os.listdir(CAMINHO_PLANILHAS)
    if f.endswith('.xlsx') and
       ARQUIVO_MODELO not in f and
       not f.startswith('Relatorio_')
]

# Carregar dados
df_lista = []
for nome in arquivos:
    caminho = os.path.join(CAMINHO_PLANILHAS, nome)
    try:
        df = pd.read_excel(caminho)
        df_lista.append(df)
    except Exception as e:
        print(f"Erro ao ler {nome}: {e}")

df_total = pd.concat(df_lista, ignore_index=True)
df_total['DATA TRANSACAO'] = pd.to_datetime(df_total['DATA TRANSACAO'], errors='coerce')
df_total.columns = df_total.columns.str.strip().str.upper()

# Filtrar até hoje
df_validos = df_total[df_total['DATA TRANSACAO'].dt.date <= hoje]

# Gerar dados consolidados
consolidados = []
for (placa, combustivel), grupo in df_validos.groupby(['PLACA', 'TIPO COMBUSTIVEL']):
    grupo = grupo.sort_values('DATA TRANSACAO')
    if len(grupo) < 2:
        continue
    hodometro_ini = grupo['HODOMETRO OU HORIMETRO'].iloc[0]
    hodometro_fim = grupo['HODOMETRO OU HORIMETRO'].iloc[-1]
    litros = grupo['LITROS'].sum()
    km_rodado = hodometro_fim - hodometro_ini
    km_l = round(km_rodado / litros, 2) if litros > 0 else None

    linha_base = grupo.iloc[-1].copy()
    linha_base['LITROS'] = litros
    linha_base['KM/L'] = km_l
    consolidados.append(linha_base)

# Aplicar colunas do modelo
modelo = pd.read_excel(os.path.join(CAMINHO_PLANILHAS, ARQUIVO_MODELO))
df_resultado = pd.DataFrame(consolidados)
df_final = df_resultado[modelo.columns]

# Exportar
wb = Workbook()
ws = wb.active
ws.title = f"Consolidado {hoje.strftime('%d-%m')}"

for r in dataframe_to_rows(df_final, index=False, header=True):
    ws.append(r)

col_kml = list(df_final.columns).index("KM/L") + 1
for row in ws.iter_rows(min_row=2, min_col=col_kml, max_col=col_kml):
    for cell in row:
        cell.number_format = numbers.FORMAT_NUMBER_00

nome_arquivo = f"Relatorio_Consumo_Acumulado_{hoje.strftime('%Y%m%d')}.xlsx"
caminho_saida = os.path.join(CAMINHO_PLANILHAS, nome_arquivo)
wb.save(caminho_saida)
print(f"Relatório gerado: {caminho_saida}")
