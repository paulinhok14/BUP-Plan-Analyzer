import pandas as pd
import warnings
import matplotlib.pyplot as plt
from PIL import Image
import customtkinter as ctk


warnings.filterwarnings("ignore")

# Lista de Scenarios
scenarios_list = []


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

    return bup_scope


def generate_histogram(bup_scope):  # Gera o Histograma e retorna uma Figura e os maiores Leadtimes

    # DF com os maiores Leadtimes
    highest_leadimes = bup_scope.nlargest(3, 'Leadtime').to_string(index=False)

    # Criando uma figura e eixos para inserir o gráfico
    fig, ax = plt.subplots(figsize=(8, 4))
    fig.patch.set_facecolor("None")

    # Criando Histograma, e salvando as informações em variáveis de controle
    n, bins, patches = ax.hist(bup_scope['Leadtime'], bins=20, edgecolor='k', linewidth=0.7, alpha=0.9)

    # Configuração do Histograma
    ax.set_xlabel('Leadtime (in days)')
    ax.set_ylabel('Materials Count')
    ax.set_title('Leadtime Histogram')

    # Inserindo a contagem em cada barra
    for count, bar in zip(n, patches):
        x = bar.get_x() + bar.get_width() / 2
        y = bar.get_height()
        ax.text(x, y, f'{int(count)}', ha='center', va='bottom', fontdict={
            'family': 'open sans',
            'size': 9
        })

    # Salvando a figura com fundo transparente, para depois carregá-la.
    fig.savefig('histogram.png', transparent=True)

    # Carregando para um objeto Image
    histogram_image = ctk.CTkImage(Image.open('histogram.png'),
                                   dark_image=Image.open('histogram.png'),
                                   size=(660, 330))

    return histogram_image, highest_leadimes


# Função para criação dos Scenarios
def create_scenario() -> None:

    global scenarios_list

    scenario = {}

    if not scenarios_list:

        # Caso não tenha nenhum elemento na lista, segue o procedimento normal para cadastro.

        print("\nCONTRACT INFORMATION FULLFILMENT:\n")
        print("--------------------------------\n")
