import pandas as pd
import warnings
import sys  # sys.exit() para funcionar como um 'EXIT SCRIPT;' em Qlik
import customtkinter as ctk

warnings.filterwarnings("ignore")


# Função para ler o arquivo e informações complementares
def read_scope_file(file_full_path: str):

    # Colunas a serem lidas no arquivo (essenciais)
    colunas = ['PN', 'ECODE', 'QTY']

    # Leitura da tabela e Filtros
    scope = pd.read_excel(file_full_path, usecols=colunas)
    scope_filtered = scope.query("QTY > 0").copy()

    # Formatação
    scope_filtered.loc[:, 'ECODE'] = scope_filtered['ECODE'].astype(int)

    # Busca de informações complementares (Leadtime, ECCN, etc)
    # leadtime_source = r'\\sjkfs05\vss\GMT\40. Stock Efficiency\J - Operational Efficiency\006 - Srcfiles\003 - SAP\marcsa.txt'
    leadtime_source = r'C:\Users\prsarau\Documents\Arquivos de Trabalho\(Diego Sodre) Build-Up Plan Analyzer\marcsa.txt'

    # Colunas a serem lidas
    columns = ['Material', ' PEP']

    # Lendo a base de Leadtimes
    leadtimes = pd.read_csv(leadtime_source, usecols=columns, encoding='latin', skiprows=3, sep='|', low_memory=False)

    # Removendo nulos
    leadtimes = leadtimes.dropna()

    # Renomeando colunas
    leadtimes.rename(columns={'Material': 'ECODE', ' PEP': 'LEADTIME'}, inplace=True)

    # Vinculando o Leadtime aos Materiais
    bup_scope = scope_filtered.merge(leadtimes, on='ECODE', how='left')

    # Convertendo as colunas numéricas de float para int
    bup_scope['ECODE'] = bup_scope['ECODE'].astype(int)
    bup_scope['QTY'] = bup_scope['QTY'].astype(int)
    bup_scope['LEADTIME'] = bup_scope['LEADTIME'].astype(int)

    # Ordenando pelo Leadtime descending
    bup_scope = bup_scope.sort_values('LEADTIME', ascending=False)

    # Renomeando as colunas
    bup_scope['Ecode'] = bup_scope['ECODE']
    bup_scope['Qty'] = bup_scope['QTY']
    bup_scope['Leadtime'] = bup_scope['LEADTIME']

    del bup_scope['ECODE']
    del bup_scope['QTY']
    del bup_scope['LEADTIME']

    # Maiores Leadtimes
    highest_leadimes = bup_scope.nlargest(3, 'Leadtime')

    return bup_scope, highest_leadimes


def generate_histogram():
    pass
