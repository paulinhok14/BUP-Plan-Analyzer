import pandas as pd
import warnings
import matplotlib.pyplot as plt
from matplotlib.ticker import FuncFormatter
from PIL import Image
from io import BytesIO
import os
import customtkinter as ctk
from tkinter import messagebox

warnings.filterwarnings("ignore")

active_user = os.getlogin()

# Lista de Scenarios
scenarios_list = []

excel_icon_path = r'excel_transparent.png'
export_output_path = fr'C:\Users\{active_user}\Downloads\bup_scenarios_data.xlsx'

# Declarando as variáveis que irão armazenar temporariamente os valores prévios de Cenários já cadastrados,
# para caso o usuário queira reaproveitar os parâmetros contratuais do Cenário.
t0_previous_value, hyp_t0_previous_value, acft_delivery_start_previous_value, material_delivery_start_previous_value\
    , material_delivery_end_previous_value = None, None, None, None, None


# Função para ler o arquivo e informações complementares
def read_scope_file(file_full_path: str):
    # Colunas a serem lidas no arquivo (essenciais)
    colunas = ['PN', 'ECODE', 'QTY', 'EIS', 'SPC']

    # Fonte das informações complementares
    # leadtime_source = r'\\sjkfs05\vss\GMT\40. Stock Efficiency\J - Operational Efficiency\006 - Srcfiles\003 - SAP\marcsa.txt'
    leadtime_source = r'marcsa.txt'
    # ecode_data_path = r'\\egmap20038\Databases\DB_Ecode-Data.txt'
    ecode_data_path = r'DB_Ecode-Data.txt'

    # Leitura da tabela e Filtros
    scope = pd.read_excel(file_full_path, usecols=colunas)
    scope_filtered = scope.query("QTY > 0").copy()

    # Formatação
    scope_filtered.loc[:, 'ECODE'] = scope_filtered['ECODE'].astype(int)
    scope_filtered['EIS'] = scope_filtered['EIS'].fillna('')

    # -------------- Busca de informações complementares (Leadtime, ECCN, Acq Cost, Reparabilidade, etc) -------------

    # Colunas a serem lidas fonte SAP
    sap_source_columns = ['Material', ' PEP']
    # Lendo a base de Leadtimes
    leadtimes = pd.read_csv(leadtime_source, usecols=sap_source_columns, encoding='latin', skiprows=3, sep='|', low_memory=False)
    # Removendo nulos
    leadtimes = leadtimes.dropna()
    # Renomeando colunas
    leadtimes.rename(columns={'Material': 'ECODE', ' PEP': 'LEADTIME'}, inplace=True)
    # Vinculando o Leadtime aos Materiais
    bup_scope = scope_filtered.merge(leadtimes, on='ECODE', how='left')

    # Colunas a serem lidas Ecode Data
    ecode_data_columns = ['ECODE', 'ACQCOST']

    # Pegando para cada Ecode o índice do registro que tem o maior Acq Cost (premissa p/ duplicados)
    ecode_data = pd.read_csv(ecode_data_path, usecols=ecode_data_columns).drop_duplicates()
    ecode_data_max_acqcost = ecode_data.groupby('ECODE')['ACQCOST'].idxmax()
    ecode_data_filtered = ecode_data.loc[ecode_data_max_acqcost].reset_index(drop=True)

    # Fazendo a regra do Tipo de Material (Repairable/Expendable)
    bup_scope['SPC'] = bup_scope['SPC'].apply(
        lambda x: 'Repairable' if x in [2, 6] else 'Expendable'
    )

    # Convertendo as colunas numéricas de float para int
    bup_scope['ECODE'] = bup_scope['ECODE'].astype(int)
    bup_scope['QTY'] = bup_scope['QTY'].astype(int)
    bup_scope['LEADTIME'] = bup_scope['LEADTIME'].fillna(127).astype(int)  # Leadtime default 127
    # Garantindo que Acq Cost seja flutuante
    ecode_data_filtered['ACQCOST'] = ecode_data_filtered['ACQCOST'].str.replace(',', '.').astype(float)

    # Fazendo join das informações do Ecode Data
    # Acq Cost
    bup_scope = bup_scope.merge(ecode_data_filtered[['ECODE', 'ACQCOST']], how='left', on='ECODE')

    # Ordenando pelo Leadtime descending
    bup_scope = bup_scope.sort_values('LEADTIME', ascending=False).reset_index(drop=True)

    # Renomeando as colunas
    bup_scope.rename(columns={'ECODE': 'Ecode', 'QTY': 'Qty', 'LEADTIME': 'Leadtime',
                              'EIS': 'EIS Critical', 'ACQCOST': 'Acq Cost'}, inplace=True)

    return bup_scope


def generate_dispersion_chart(bup_scope):
    # Função para formatar os valores do eixo y em milhares
    def format_acq_cost(value, _):
        return f'US$ {value / 1000:.0f}k'

    # Tamanho da imagem
    width, height = 600, 220
    fig, ax = plt.subplots(figsize=(width / 100, height / 100))
    expendable_items = bup_scope[bup_scope['SPC'] == 'Expendable']
    repairable_items = bup_scope[bup_scope['SPC'] == 'Repairable']

    # Plotando os pontos Expendable
    ax.scatter(expendable_items['Leadtime'], expendable_items['Acq Cost'], color='orange', label='Expendables')
    # Plotando os pontos Repairable
    ax.scatter(repairable_items['Leadtime'], repairable_items['Acq Cost'], color='purple', label='Repairables')

    # Adicionando rótulos e legendas
    ax.set_xlabel('Leadtime')
    ax.set_ylabel('Acq Cost')
    ax.set_title('Dispersion Acq Cost x Leadtime', fontsize=10)
    ax.legend(fontsize=9, framealpha=0.6)
    plt.grid(True)
    # Configurando o formato personalizado para o eixo y
    ax.yaxis.set_major_formatter(FuncFormatter(format_acq_cost))

    # --------------- TRANSFORMANDO EM UMA IMAGEM PARA SER EXIBIDA ---------------

    # Salvando a figura matplotlib em um objeto BytesIO (memória), para não ter que salvar em um arquivo de imagem
    tmp_img_dispersion_chart = BytesIO()
    fig.savefig(tmp_img_dispersion_chart, format='png', transparent=True)
    tmp_img_dispersion_chart.seek(0)

    # Carregando a imagem do gráfico para um objeto Image que irá ser retornado pela função
    dispersion_chart = Image.open(tmp_img_dispersion_chart)

    return dispersion_chart


def generate_histogram(bup_scope):  # Gera o Histograma e retorna uma Figura e os maiores Leadtimes

    # DF com os maiores Leadtimes
    highest_leadimes = bup_scope.nlargest(3, 'Leadtime').to_string(index=False)

    # Tamanho da imagem
    width, height = 600, 220

    # Criando uma figura e eixos para inserir o gráfico
    fig, ax = plt.subplots(figsize=(width / 100, height / 100))
    fig.patch.set_facecolor("None")

    # Criando Histograma, e salvando as informações em variáveis de controle
    n, bins, patches = ax.hist(bup_scope['Leadtime'], bins=20, edgecolor='k', linewidth=0.7, alpha=0.9)

    # Configuração do Histograma
    ax.set_xlabel('Leadtime (in days)')
    ax.set_ylabel('Materials Count')
    ax.set_title('Leadtime Histogram', fontsize=10)
    # Ajustando o limite do eixo y ( a maior barra estava escapando)
    ax.set_ylim(0, max(n) + 50)  # Adicionando uma margem para acomodar a contagem no topo

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
                                   size=(600, 220))

    return histogram_image, highest_leadimes


def create_scenario(scenario_window, bup_scope, efficient_curve_window, hypothetical_curve_window, lbl_pending_scenario) -> None:
    global scenarios_list

    scenario = {}

    # Variável CTk que irá armazenar a contagem de Scenarios. Será util para implementar um tracking com callback
    # para exibição ou ocultamento de componentes
    var_scenarios_count = ctk.IntVar()

    # Imagem com o ícone Excel
    excel_icon = ctk.CTkImage(light_image=Image.open(excel_icon_path),
                      dark_image=Image.open(excel_icon_path),
                      size=(30, 30))

    # Função executada ao exportar Dados
    def export_data():
        try:
            bup_scope.to_excel(export_output_path, index=False)
            messagebox.showinfo(title="Success!", message=str("Excel sheet was exported to: " + export_output_path))
        except Exception as ex:
            messagebox.showinfo(title="Error!", message=str(ex) + "\n\n Please make sure that the Excel file is "
                                                                  "closed and you have access to the Downloads folder.")

    # Botão de exportar dados
    btn_export_data = ctk.CTkButton(efficient_curve_window, text="Export to Excel",
                                    font=ctk.CTkFont('open sans', size=10, weight='bold'),
                                    image=excel_icon, compound="top", fg_color="transparent",
                                    text_color="#000000", hover=False, border_spacing=1,
                                    command=export_data)

    # Função que irá ser chamada para avaliar a variável de controle e Exibir/Ocultar botão de Exportar Dados
    def export_data_button(scenarios_count):

        if scenarios_count.get() == 1:
            # Exibir o botão de Exportar para Excel e ocultar a mensagem de Criação de Scenario
            btn_export_data.place(relx=0.92, rely=0.94, anchor=ctk.CENTER)
            lbl_pending_scenario.place_forget()
        else:
            pass

    # Fazendo o tracing da variável e chamando a função toda vez que a variável for alterada
    var_scenarios_count.trace_add("write", callback=lambda *args: export_data_button(var_scenarios_count))

    # Se a lista de Cenários contiver algum já cadastrado, é oferecido ao usuário a opção de utilizar os valores
    # de Contractual Conditions do primeiro cenário, alterando apenas os parâmetros de Procurement Length

    if scenarios_list:
        # Função para abrir a caixa de Diálogo questionando ao usuário
        def open_confirm_dialog():
            confirm_window = ctk.CTkToplevel(scenario_window)
            confirm_window.title("Warning!")
            confirm_window.resizable(width=False, height=False)

            # Geometria da Tela de Diálogo
            cw_width = 400
            cw_height = 170
            conf_window_width = confirm_window.winfo_screenwidth()  # Width of the screen
            conf_window_height = confirm_window.winfo_screenheight()  # Height of the screen
            # Calcula o X e Y inicial a posicionar a tela
            cw_x = int((conf_window_width / 2) - (cw_width / 2))
            cw_y = int((conf_window_height / 2) - (cw_height / 2))
            confirm_window.geometry('%dx%d+%d+%d' % (cw_width, cw_height, cw_x, cw_y))

            # Setando(captando) o foco para as janelas, de forma subsequente
            confirm_window.grab_set()

            # Elementos da Tela de Diálogo

            lbl_question = ctk.CTkLabel(confirm_window, text="There is already a registered Scenario. Do you want to "
                                                             "use previously Contractual Conditions information?",
                                        font=ctk.CTkFont('open sans', size=13, weight='bold'),
                                        width=300, wraplength=350
                                        )
            lbl_question.pack(pady=(20, 0))

            # Função que, ao clicar no botão YES, ele usa os valores do primeiro Scenario cadastrado e desabilita os Entry
            def use_previous_scenario_values():
                # Usando de forma global as variáveis de controle para armazenar os valores previamente cadastrados
                global t0_previous_value, hyp_t0_previous_value, acft_delivery_start_previous_value, material_delivery_start_previous_value \
                    , material_delivery_end_previous_value

                t0_previous_value = ctk.StringVar(value=
                                                  scenarios_list[0]['t0'].strftime("%d/%m/%Y")
                                                  )
                entry_t0.configure(textvariable=t0_previous_value)
                entry_t0.configure(state="disabled")

                hyp_t0_previous_value = ctk.StringVar(value=
                                                      scenarios_list[0]['hyp_t0_start'])
                entry_hyp_pln_start.configure(textvariable=hyp_t0_previous_value)
                entry_hyp_pln_start.configure(state="disabled")

                acft_delivery_start_previous_value = ctk.StringVar(
                    value=scenarios_list[0]['acft_delivery_start'].strftime("%d/%m/%Y")
                )
                entry_acft_delivery_start.configure(textvariable=acft_delivery_start_previous_value)
                entry_acft_delivery_start.configure(state="disabled")

                material_delivery_start_previous_value = ctk.StringVar(
                    value=scenarios_list[0]['material_delivery_start'])
                entry_material_delivery_start.configure(textvariable=material_delivery_start_previous_value)
                entry_material_delivery_start.configure(state="disabled")

                material_delivery_end_previous_value = ctk.StringVar(value=scenarios_list[0]['material_delivery_end'])
                entry_material_delivery_end.configure(textvariable=material_delivery_end_previous_value)
                entry_material_delivery_end.configure(state="disabled")

            def nullify_previous_scenario_variables():
                """ Função que tornam vazias as variáveis de controle que armazenam os cenários pré-cadastrados, pois
                o usuário optou por não utilizar o cenário ja cadastrado
                """

                # Usando de forma global as variáveis
                global t0_previous_value, hyp_t0_previous_value, acft_delivery_start_previous_value, material_delivery_start_previous_value \
                    , material_delivery_end_previous_value

                t0_previous_value, hyp_t0_previous_value, acft_delivery_start_previous_value, material_delivery_start_previous_value \
                    , material_delivery_end_previous_value = None, None, None, None, None

            # Botão YES
            btn_yes = ctk.CTkButton(confirm_window, text='Yes', command=lambda: (
                use_previous_scenario_values(),
                confirm_window.destroy(),
                scenario_window.lift(),
                scenario_window.grab_set()
            ),
                                    font=ctk.CTkFont('open sans', size=12, weight='bold'),
                                    bg_color="#ebebeb", fg_color="#009898", hover_color="#006464",
                                    width=100, height=30, corner_radius=30, cursor="hand2"
                                    )
            btn_yes.place(relx=0.3, rely=0.8, anchor=ctk.CENTER)

            # Botão NO
            btn_no = ctk.CTkButton(confirm_window, text='No', command=lambda: (confirm_window.destroy(),
                                                                               nullify_previous_scenario_variables(),
                                                                               scenario_window.lift(),
                                                                               scenario_window.grab_set()),
                                   font=ctk.CTkFont('open sans', size=12, weight='bold'),
                                   bg_color="#ebebeb", fg_color="#ff0000", hover_color="#af0003",
                                   width=100, height=30, corner_radius=30, cursor="hand2"
                                   )
            btn_no.place(relx=0.7, rely=0.8, anchor=ctk.CENTER)

        # Abrindo a caixa de diálogo caso já tenha um Scenario cadastrado
        open_confirm_dialog()

    # ----------------- CONTRACTUAL CONDITIONS -----------------

    # Frame interno Contractual Conditions
    contractual_conditions_frame = ctk.CTkFrame(scenario_window, width=440, corner_radius=20)
    contractual_conditions_frame.pack(pady=(15, 0), expand=False)

    # Label título Contractual Conditions
    lbl_contractual_conditions = ctk.CTkLabel(contractual_conditions_frame, text="Contractual Conditions",
                                              font=ctk.CTkFont('open sans', size=14, weight='bold')
                                              )
    lbl_contractual_conditions.grid(row=0, column=0, columnspan=4, sticky="n", pady=(10, 10))

    # --- t0 ---
    lbl_t0 = ctk.CTkLabel(contractual_conditions_frame, text="T0 Date (*)",
                          font=ctk.CTkFont('open sans', size=11, weight="bold")
                          )
    lbl_t0.grid(row=1, column=0, sticky="w", padx=(12, 0))
    entry_t0 = ctk.CTkEntry(contractual_conditions_frame, width=160,
                            placeholder_text="Format: DD/MM/YYYY")
    entry_t0.grid(row=2, column=0, padx=(10, 0), sticky="w")

    # --- Hypothetical Planning Start Date (an integer added to t0, in months, Ex: 3 for t0+3) ---
    lbl_hyp_pln_start = ctk.CTkLabel(contractual_conditions_frame, text="t0+X",
                          font=ctk.CTkFont('open sans', size=11, weight="bold")
                                     )
    lbl_hyp_pln_start.grid(row=1, column=1, sticky="w")
    entry_hyp_pln_start = ctk.CTkEntry(contractual_conditions_frame, width=35,
                            placeholder_text="3")
    entry_hyp_pln_start.grid(row=2, column=1, sticky="w")

    # --- Aircraft Delivery Start ---
    lbl_acft_delivery_start = ctk.CTkLabel(contractual_conditions_frame, text="Aircraft Delivery Start (*)",
                                           font=ctk.CTkFont('open sans', size=11, weight='bold')
                                           )
    lbl_acft_delivery_start.grid(row=1, column=2, columnspan=2, padx=(0, 12), sticky="e")
    entry_acft_delivery_start = ctk.CTkEntry(contractual_conditions_frame, width=200,
                                             placeholder_text="Format: DD/MM/YYYY")
    entry_acft_delivery_start.grid(row=2, column=2, columnspan=2, padx=(0, 10), sticky="e")

    # --- Material Delivery Start ---
    lbl_material_delivery_start = ctk.CTkLabel(contractual_conditions_frame, text="Material Delivery Start (*)",
                                               font=ctk.CTkFont('open sans', size=11, weight='bold')
                                               )
    lbl_material_delivery_start.grid(row=3, column=0, columnspan=2, sticky="w", padx=12)
    entry_material_delivery_start = ctk.CTkEntry(contractual_conditions_frame, width=200,
                                                 placeholder_text="Nº of months (ex: 38 for T+38)")
    entry_material_delivery_start.grid(row=4, column=0, columnspan=2, padx=10, sticky="w", pady=(0, 20))

    # --- Material Delivery End ---
    lbl_material_delivery_end = ctk.CTkLabel(contractual_conditions_frame, text="Material Delivery End (*)",
                                             font=ctk.CTkFont('open sans', size=11, weight='bold')
                                             )
    lbl_material_delivery_end.grid(row=3, column=2, columnspan=2, sticky="e", padx=12)
    entry_material_delivery_end = ctk.CTkEntry(contractual_conditions_frame, width=200,
                                               placeholder_text="Nº of months (ex: 43 for T+43)")
    entry_material_delivery_end.grid(row=4, column=2, columnspan=2, padx=10, sticky="e", pady=(0, 20))

    # ----------------- PROCUREMENT LENGTH -----------------

    # Frame interno Procurement Length
    procurement_length_frame = ctk.CTkFrame(scenario_window, width=440, corner_radius=20)
    procurement_length_frame.pack(pady=(15, 0), expand=False)

    # Label título Procurement Length
    lbl_procurement_length = ctk.CTkLabel(procurement_length_frame, text="Procurement Length",
                                          font=ctk.CTkFont('open sans', size=14, weight='bold')
                                          )
    lbl_procurement_length.grid(row=0, columnspan=2, sticky="n", pady=(10, 10))

    # --- PR Release and Approval VSS ---
    lbl_pr_release_approval_vss = ctk.CTkLabel(procurement_length_frame, text="PR Release & Approval VSS (days)",
                                               font=ctk.CTkFont('open sans', size=11, weight='bold')
                                               )
    lbl_pr_release_approval_vss.grid(row=1, column=0, sticky="w", padx=12)
    entry_pr_release_approval_vss = ctk.CTkEntry(procurement_length_frame, width=200,
                                                 placeholder_text="Default: 5 days")
    entry_pr_release_approval_vss.grid(row=2, column=0, padx=10, sticky="w")

    # --- PO Commercial Condition ---
    lbl_po_commercial_condition = ctk.CTkLabel(procurement_length_frame, text="PO Commercial Condition (days)",
                                               font=ctk.CTkFont('open sans', size=11, weight='bold')
                                               )
    lbl_po_commercial_condition.grid(row=1, column=1, padx=12, sticky="e")
    entry_po_commercial_condition = ctk.CTkEntry(procurement_length_frame, width=200,
                                                 placeholder_text="Default: 30 days")
    entry_po_commercial_condition.grid(row=2, column=1, padx=10, sticky="e")

    # --- PO Conversion ---
    lbl_po_conversion = ctk.CTkLabel(procurement_length_frame, text="PO Conversion (days)",
                                     font=ctk.CTkFont('open sans', size=11, weight='bold')
                                     )
    lbl_po_conversion.grid(row=3, column=0, sticky="w", padx=12)
    entry_po_conversion = ctk.CTkEntry(procurement_length_frame, width=200,
                                       placeholder_text="Default: 30 days")
    entry_po_conversion.grid(row=4, column=0, padx=10, sticky="w")

    # --- Export License ---
    lbl_export_license = ctk.CTkLabel(procurement_length_frame, text="Export License [Defense Only]",
                                      font=ctk.CTkFont('open sans', size=11, weight='bold')
                                      )
    lbl_export_license.grid(row=3, column=1, padx=12, sticky="e")
    entry_export_license = ctk.CTkEntry(procurement_length_frame, width=200,
                                        placeholder_text="Default: 0 days")
    entry_export_license.grid(row=4, column=1, padx=10, sticky="e")

    # --- Buffer ---
    lbl_buffer = ctk.CTkLabel(procurement_length_frame, text="Buffer (days)",
                              font=ctk.CTkFont('open sans', size=11, weight='bold')
                              )
    lbl_buffer.grid(row=5, column=0, sticky="w", padx=12)
    entry_buffer = ctk.CTkEntry(procurement_length_frame, width=200,
                                placeholder_text="Default: 60 days")
    entry_buffer.grid(row=6, column=0, padx=10, sticky="w", pady=(0, 20))

    # --- Outbound Logistic ---
    lbl_outbound_logistic = ctk.CTkLabel(procurement_length_frame, text="Outbound Logistic (days)",
                                         font=ctk.CTkFont('open sans', size=11, weight='bold')
                                         )
    lbl_outbound_logistic.grid(row=5, column=1, padx=12, sticky="e")
    entry_outbound_logistic = ctk.CTkEntry(procurement_length_frame, width=200,
                                           placeholder_text="Default: 30 days")
    entry_outbound_logistic.grid(row=6, column=1, padx=10, sticky="e", pady=(0, 20))

    # ----------------- LABEL -----------------

    lbl_required_infornation = ctk.CTkLabel(scenario_window, text="(*) Required Information",
                                            font=ctk.CTkFont('open sans', size=10, weight='bold'),
                                            text_color="#ff0000")
    lbl_required_infornation.pack(padx=(20, 0), anchor="w")

    # ----------------- BOTÕES DE INTERAÇÃO -----------------

    # Função para retornar os valores digitados pelo usuário nos Entry, com tratativa para Defaults
    def get_entry_values():

        # --------- Contractual Conditions ---------

        # Nestas condições, primeiro há uma verificação se o usuário optou por usar informações já cadastradas

        # t0

        if t0_previous_value is None:

            # Verifica se o valor é uma data válida antes de atribuir à variável
            date_t0_value = pd.to_datetime(entry_t0.get(), format='%d/%m/%Y', errors='coerce')
            if not pd.isna(date_t0_value):
                scenario['t0'] = date_t0_value
            else:
                messagebox.showerror("Error",
                                     "Invalid date. Please enter a valid date format for T0.")
                return
        else:
            scenario['t0'] = pd.to_datetime(t0_previous_value.get(), format='%d/%m/%Y', errors='coerce')

        # t0 + X: Integer (default: 3) que será somado ao t0 para indicar data hipotética de início das compras de materiais

        if hyp_t0_previous_value is None:

            try:
                if entry_hyp_pln_start.get().strip() != "":
                    scenario['hyp_t0_start'] = int(entry_hyp_pln_start.get())
                else:
                    scenario['hyp_t0_start'] = int(3)  # Default: 3
            except ValueError:
                messagebox.showerror("Error",
                                     "Invalid character. Please enter a valid number (months, in integer) for Hypothetical T0 Start (t0+x).")
                return

        else:
            scenario['hyp_t0_start'] = int(hyp_t0_previous_value.get())

        # Aircraft Delivery Start

        if acft_delivery_start_previous_value is None:

            # Verifica se o valor é uma data válida antes de atribuir à variável
            date_acft_delivery_value = pd.to_datetime(entry_acft_delivery_start.get(), format='%d/%m/%Y', errors='coerce')
            if not pd.isna(date_acft_delivery_value):
                scenario['acft_delivery_start'] = date_acft_delivery_value
            else:
                messagebox.showerror("Error",
                                     "Invalid date. Please enter a valid date format for Aircraft Delivery Start.")
                return
        else:
            scenario['acft_delivery_start'] = pd.to_datetime(acft_delivery_start_previous_value.get(), format='%d/%m/%Y', errors='coerce')

        # Material Delivery Start

        if material_delivery_start_previous_value is None:

            try:
                scenario['material_delivery_start'] = int(entry_material_delivery_start.get())
            except ValueError:
                messagebox.showerror("Error",
                                     "Invalid character. Please enter a valid number for Material Delivery Start.")
                return
        else:
            scenario['material_delivery_start'] = int(material_delivery_start_previous_value.get())

        # Material Delivery End

        if material_delivery_end_previous_value is None:

            try:
                scenario['material_delivery_end'] = int(entry_material_delivery_end.get())
            except ValueError:
                messagebox.showerror("Error",
                                     "Invalid character. Please enter a valid number for Material Delivery End.")
                return
        else:
            scenario['material_delivery_end'] = int(material_delivery_end_previous_value.get())

        # --------- Procurement Length ---------
        # Atribuição dos valores Default no try/except caso o usuário não preencha nada

        # --- PR Release and Approval VSS ---
        try:
            if entry_pr_release_approval_vss.get().strip() != "":
                scenario['pr_release_approval_vss'] = int(entry_pr_release_approval_vss.get())
            else:
                scenario['pr_release_approval_vss'] = 5
        except ValueError:
            messagebox.showerror("Error",
                                 "Invalid character. Please enter a valid number for PR Release and Approval VSS.")
            return

        # --- PO Commercial Condition ---
        try:
            if entry_po_commercial_condition.get().strip() != "":
                scenario['po_commercial_condition'] = int(entry_po_commercial_condition.get())
            else:
                scenario['po_commercial_condition'] = 30
        except ValueError:
            messagebox.showerror("Error",
                                 "Invalid character. Please enter a valid number for PO Commercial Condition.")
            return

        # --- PO Conversion ---
        try:
            if entry_po_conversion.get().strip() != "":
                scenario['po_conversion'] = int(entry_po_conversion.get())
            else:
                scenario['po_conversion'] = 30
        except ValueError:
            messagebox.showerror("Error",
                                 "Invalid character. Please enter a valid number for PO Conversion.")
            return

        # --- Export License ---
        try:
            if entry_export_license.get().strip() != "":
                scenario['export_license'] = int(entry_export_license.get())
            else:
                scenario['export_license'] = 0
        except ValueError:
            messagebox.showerror("Error",
                                 "Invalid character. Please enter a valid number for Export License.")
            return

        # --- Buffer ---
        try:
            if entry_buffer.get().strip() != "":
                scenario['buffer'] = int(entry_buffer.get())
            else:
                scenario['buffer'] = 60
        except ValueError:
            messagebox.showerror("Error",
                                 "Invalid character. Please enter a valid number for Buffer.")
            return

        # --- Outbound Logistic ---
        try:
            if entry_outbound_logistic.get().strip() != "":
                scenario['outbound_logistic'] = int(entry_outbound_logistic.get())
            else:
                scenario['outbound_logistic'] = 30
        except ValueError:
            messagebox.showerror("Error",
                                 "Invalid character. Please enter a valid number for Outbound Logistic.")
            return

        # Incluindo o cenário na lista de Dicts global, e fechando a tela
        scenarios_list.append(scenario)
        scenario_window.destroy()

        # Somando 1 à IntVar com a contagem de Scenarios
        var_scenarios_count.set(var_scenarios_count.get() + 1)

        # ----------- Chamada da Função de Geração do Gráfico -----------

        # Chamando a função para gerar o gráfico de Efficient Build-Up. O retorno da função é o gráfico na figura (objeto Image),
        # Além dos DataFrames/Variáveis construídos na função, como retorno a serem usados no gráfico Hipotético
        bup_efficient_chart, df_scope_with_scenarios, scenario_dataframes = \
            generate_efficient_curve_buildup_chart(bup_scope, scenarios_list)

        # Carregando para um objeto Image do CTk
        img_bup_efficient_chart = ctk.CTkImage(bup_efficient_chart,
                                    dark_image=bup_efficient_chart,
                                    size=(580, 380))

        # Gráfico de Build-Up Efficient Curve - inputando CTkImage no Label e posicionando na tela
        ctk.CTkLabel(efficient_curve_window, image=img_bup_efficient_chart,
                     text="").place(relx=0.5, rely=0.43, anchor=ctk.CENTER)

        # Chamando a função para gerar o gráfico de Hypothetical Build-Up.
        bup_hypothetical_chart = generate_hypothetical_curve_buildup_chart(df_scope_with_scenarios, scenario_dataframes)

        # Carregando para um objeto Image do CTk
        img_bup_hypothetical_chart = ctk.CTkImage(bup_hypothetical_chart,
                                    dark_image=bup_hypothetical_chart,
                                    size=(580, 380))

        # Gráfico de Build-Up Hypothetical Curve - inputando CTkImage no Label e posicionando na tela
        ctk.CTkLabel(hypothetical_curve_window, image=img_bup_hypothetical_chart,
                    text="").place(relx=0.5, rely=0.43, anchor=ctk.CENTER)

    # Botão OK
    btn_ok = ctk.CTkButton(scenario_window, text='OK', command=get_entry_values,
                           font=ctk.CTkFont('open sans', size=12, weight='bold'),
                           bg_color="#ebebeb", fg_color="#009898", hover_color="#006464",
                           width=100, height=30, corner_radius=30, cursor="hand2"
                           ).place(relx=0.3, rely=0.92, anchor=ctk.CENTER)

    # Botão Cancelar
    btn_cancel = ctk.CTkButton(scenario_window, text='Cancel', command=scenario_window.destroy,
                               font=ctk.CTkFont('open sans', size=12, weight='bold'),
                               bg_color="#ebebeb", fg_color="#ff0000", hover_color="#af0003",
                               width=100, height=30, corner_radius=30, cursor="hand2"
                               ).place(relx=0.7, rely=0.92, anchor=ctk.CENTER)


def generate_efficient_curve_buildup_chart(bup_scope, scenarios):

    # --------------- PREPARAÇÃO DOS DADOS ---------------

    # Lista para armazenar as combinações de Cenários e Escopo
    combinations = []

    for _, row in bup_scope.iterrows():
        # Percorrendo cada elemento da lista de dicionários (cada elemento um Cenário em scenarios_list)
        for index, scenario in enumerate(scenarios):
            # Combinando os valores do dataframe (escopo) com os valores do dicionário (cenário)
            comb = {**row, 'Scenario': index, **scenario}
            combinations.append(comb)

    # Criando um novo dataframe com as combinações de Escopo e Cenários juntos
    df_scope_with_scenarios = pd.DataFrame(combinations).sort_values(by='Scenario').reset_index(drop=True)
    df_scope_with_scenarios['avg_month_diff'] = ((df_scope_with_scenarios['material_delivery_end']
                                                 - df_scope_with_scenarios['material_delivery_start']) / 2).astype(int)

    # Procurement Length - OBS: Chegará um momento que terei que criar aqui a lógica para Export License
    df_scope_with_scenarios['PN Procurement Length'] = df_scope_with_scenarios[['Leadtime', 'pr_release_approval_vss',
                                                                                'po_commercial_condition',
                                                                                'po_conversion', 'export_license',
                                                                                'buffer', 'outbound_logistic']].sum(axis=1)

    # Gerando a data média (do intervalo de entrega dos materiais baseada no t0).
    df_scope_with_scenarios['avg_date_between_materials_deadline'] = df_scope_with_scenarios.apply(
        lambda linha: linha['t0'] + pd.DateOffset(months=linha['material_delivery_start']) +
        pd.DateOffset(months=linha['avg_month_diff']), axis=1)

    # Criando colunas de Data para as 2 que vem como inteiro baseadas no t0
    df_scope_with_scenarios['material_delivery_start_date'] = df_scope_with_scenarios.apply(
        lambda linha: linha['t0'] + pd.DateOffset(months=linha['material_delivery_start'])
        , axis=1)

    df_scope_with_scenarios['material_delivery_end_date'] = df_scope_with_scenarios.apply(
        lambda linha: linha['t0'] + pd.DateOffset(months=linha['material_delivery_end'])
        , axis=1)

    # Calculando a data que deveria ser emitida a compra do material, considerando Procurement Length e Delivery Date
    df_scope_with_scenarios['PN Order Date'] = df_scope_with_scenarios.apply(
        lambda linha: linha['avg_date_between_materials_deadline'] - pd.DateOffset(days=linha['PN Procurement Length'])
        , axis=1)

    # Criando a coluna com a data hipotética de início das compras de material, já pensando no gráfico Hipotético
    df_scope_with_scenarios['PN Order Date Hypothetical'] = df_scope_with_scenarios.apply(
        lambda linha: linha['t0'] + pd.DateOffset(months=linha['hyp_t0_start'])
        , axis=1)
    # Criando a coluna de data em que será entregue o material, para gráfico Hipotético
    df_scope_with_scenarios['Delivery Date Hypothetical'] = df_scope_with_scenarios.apply(
        lambda linha: linha['PN Order Date Hypothetical'] + pd.DateOffset(days=linha['PN Procurement Length'])
        , axis=1)

    # Pegando a Maior e Menor Data entre todas as datas possíveis, para delimitar o eixo X do gráfico
    date_columns = ['t0', 'acft_delivery_start', 'material_delivery_start_date', 'material_delivery_end_date',
                    'PN Order Date', 'PN Order Date Hypothetical']

    # O primeiro min/max retorna a menor/maior data por coluna e no final temos então uma lista de mínimos/máximos.
    # O segundo pega o menor da lista criada. Adiciono um mês em cada extremo para não coincidir as linhas verticais
    # dos parâmetros com o limite do eixo do gráfico
    min_date = df_scope_with_scenarios[date_columns].min().min() - pd.DateOffset(months=1)
    max_date = df_scope_with_scenarios[date_columns].max().max() + pd.DateOffset(months=1)

    date_range = pd.date_range(start=min_date, end=max_date, freq='M')
    df_dates = pd.DataFrame({'Date': date_range.strftime('%m/%Y')})

    # ------- Tabela para Criação dos gráficos --------
    # -- Efficient Curve
    # Criando tabela com a contagem agrupada de itens comprados por mês, para cada Scenario
    grouped_counts_eff = df_scope_with_scenarios.groupby([df_scope_with_scenarios['PN Order Date'].dt.to_period('M'), 'Scenario']).size().reset_index(
        name='Ordered Qty')
    grouped_counts_eff['PN Order Date'] = grouped_counts_eff['PN Order Date'].dt.strftime('%m/%Y')
    # -- Hypothetical Curve
    # Criando também os campos de Delivered Qty (Hipotético)
    grouped_counts_hyp = df_scope_with_scenarios.groupby([df_scope_with_scenarios['Delivery Date Hypothetical'].dt.to_period('M'),
                                                          'Scenario']).size().reset_index(name='Delivered Qty Hyp')
    grouped_counts_hyp['Delivery Date Hypothetical'] = grouped_counts_hyp['Delivery Date Hypothetical'].dt.strftime('%m/%Y')
    # -- Consolidando
    grouped_counts = grouped_counts_eff.merge(grouped_counts_hyp,
                                              left_on=['PN Order Date', 'Scenario'],
                                              right_on=['Delivery Date Hypothetical', 'Scenario'],
                                              how='outer')
    # -- Ajustando o DataFrame final
    grouped_counts['Date'] = grouped_counts['PN Order Date'].fillna(grouped_counts['Delivery Date Hypothetical'])
    grouped_counts = grouped_counts.drop(['PN Order Date', 'Delivery Date Hypothetical'], axis=1)

    # Passando a informação de Ordered Qty e Delivered Qty agrupada por mês e por Scenario para o DF com o Range de datas
    final_df_scenarios = df_dates.merge(grouped_counts, left_on='Date', right_on='Date', how='left')
    final_df_scenarios['Scenario'] = final_df_scenarios['Scenario'].fillna(-1).astype(int)
    final_df_scenarios['Ordered Qty'] = final_df_scenarios['Ordered Qty'].fillna(0)
    final_df_scenarios['Delivered Qty Hyp'] = final_df_scenarios['Delivered Qty Hyp'].fillna(0)

    # Para cada scenario, calculando a Quantidade Acumulada para plotar
    for scenario in final_df_scenarios['Scenario'].unique():
        # Filtrando o df para o Scenario atual
        scenario_df = final_df_scenarios[final_df_scenarios['Scenario'] == scenario]

        # Calculando a quantidade acumulada
        final_df_scenarios.loc[scenario_df.index, 'Accum. Ordered Qty (Eff)'] = scenario_df['Ordered Qty'].cumsum()
        final_df_scenarios.loc[scenario_df.index, 'Accum. Delivered Qty (Hyp)'] = scenario_df['Delivered Qty Hyp'].cumsum()

    # Criando um dicionário para armazenar os DataFrames de cenários
    scenario_dataframes = {}

    # Criando um DF para cada Scenario
    for scenario in final_df_scenarios['Scenario'].unique():
        if scenario != -1:  # Scenario -1 é apenas indicativo de nulidade, não é um cenário real
            tmp_filtered_scenario_final_df = final_df_scenarios[final_df_scenarios['Scenario'] == scenario]
            scenario_df = df_dates.merge(tmp_filtered_scenario_final_df[
                                             ['Date', 'Scenario', 'Ordered Qty', 'Delivered Qty Hyp',
                                              'Accum. Ordered Qty (Eff)', 'Accum. Delivered Qty (Hyp)']],
                                         left_on='Date', right_on='Date', how='left')
            # Preenchendo nulos
            scenario_df['Scenario'] = scenario_df['Scenario'].fillna(scenario).astype(int)
            scenario_df['Ordered Qty'] = scenario_df['Ordered Qty'].fillna(0)
            scenario_df['Delivered Qty Hyp'] = scenario_df['Delivered Qty Hyp'].fillna(0)

            # Preenchendo os meses vazios para Accumulated Qty (tanto Eff quanto Hyp) com base no último registro
            scenario_df['Accum. Ordered Qty (Eff)'].fillna(method='ffill', inplace=True)
            scenario_df['Accum. Delivered Qty (Hyp)'].fillna(method='ffill', inplace=True)

            # Armazenando o DataFrame no dicionário com o nome do cenário
            scenario_dataframes[f'Scenario_{int(scenario)}'] = scenario_df

    # --------------- GERAÇÃO DO GRÁFICO ---------------

    # Lista com as cores, para que cada Scenario tenha uma cor específica e facilite a diferenciação
    colors_array = ['blue', 'orange', 'black', 'green', 'purple']

    # Tamanho da imagem
    width, height = 580, 380

    # Criando uma figura e eixos para inserir o gráfico
    figura, eixos = plt.subplots(figsize=(width / 100, height / 100))

    # Plotando a linha para cada Scenario do dicionário
    for index, (scenario_name, scenario_df) in enumerate(scenario_dataframes.items()):
        eixos.plot(scenario_df['Date'], scenario_df['Accum. Ordered Qty (Eff)'], label=f'Scen. {index}', color=colors_array[index])
        # Configurando o eixo
        plt.xticks(scenario_df.index[::3], scenario_df['Date'][::3], rotation=45, ha='right')

        # Obtendo a data t0 para o Scenario atual e convertendo para o formato MM/YYYY
        t0_date = pd.to_datetime(df_scope_with_scenarios.loc[df_scope_with_scenarios['Scenario'] == index, 't0'].values[0])
        t0_date = t0_date.strftime('%m/%Y')
        # Adicionando uma linha vertical em t0
        eixos.axvline(x=t0_date, linestyle='--', color=colors_array[index], label=f't0: Scen. {index}')

        # Obtendo a data acft_delivery_start para o Scenario atual e convertendo para o formato MM/YYYY
        acft_delivery_start_date = pd.to_datetime(df_scope_with_scenarios.loc[df_scope_with_scenarios['Scenario'] == index, 'acft_delivery_start'].values[0])
        acft_delivery_start_date = acft_delivery_start_date.strftime('%m/%Y')
        # Adicionando uma linha vertical em acft_delivery_start
        eixos.axvline(x=acft_delivery_start_date, linestyle='dotted', color=colors_array[index], label=f'Acft Delivery Start: Scen. {index}')

        # Adicionando uma faixa de entrega dos materiais entre as data Início e Fim
        material_delivery_start_date = pd.to_datetime(df_scope_with_scenarios.loc[df_scope_with_scenarios['Scenario'] == index, 'material_delivery_start_date'].values[0])
        material_delivery_start_date = material_delivery_start_date.strftime('%m/%Y')
        material_delivery_end_date = pd.to_datetime(df_scope_with_scenarios.loc[df_scope_with_scenarios['Scenario'] == index, 'material_delivery_end_date'].values[0])
        material_delivery_end_date = material_delivery_end_date.strftime('%m/%Y')

        eixos.axvspan(material_delivery_start_date, material_delivery_end_date, alpha=0.5, color=colors_array[index])

    # Configuração do Gráfico
    eixos.set_ylabel('Materials Ordered Qty (Accumulated)')
    eixos.set_title('Efficient Curve: Build-Up Forecast')
    eixos.grid(True)

    # Ajustando espaçamento dos eixos para não cortar os rótulos
    plt.subplots_adjust(left=0.15, right=0.9, bottom=0.2, top=0.9)

    # Legenda
    eixos.legend(loc='upper left', fontsize=7, framealpha=0.8)

    # --------------- TRANSFORMANDO EM UMA IMAGEM PARA SER EXIBIDA ---------------

    # Salvando a figura matplotlib em um objeto BytesIO (memória), para não ter que salvar em um arquivo de imagem
    tmp_img_bup_chart = BytesIO()
    figura.savefig(tmp_img_bup_chart, format='png', transparent=True)
    tmp_img_bup_chart.seek(0)

    # Carregando a imagem do gráfico para um objeto Image que irá ser retornado pela função
    bup_chart = Image.open(tmp_img_bup_chart)

    return bup_chart, df_scope_with_scenarios, scenario_dataframes


def generate_hypothetical_curve_buildup_chart(df_scope_with_scenarios, scenario_dataframes):
    """
    Function that creates the Hypothetycal Curve BuildUp Chart.
    :param scenario_dataframes: Dictionary with all scenarios dataframes.
    :param df_scope_with_scenarios: Created DataFrame on Efficient Curve Build-Up construction. Combinations Scope/Scenarios.
    :return: Returns an Image object
    """

    # --------------- GERAÇÃO DO GRÁFICO ---------------

    # Lista com as cores, para que cada Scenario tenha uma cor específica e facilite a diferenciação
    colors_array = ['blue', 'orange', 'black', 'green', 'purple']

    # Tamanho da imagem
    width, height = 580, 380

    # Criando uma figura e eixos para inserir o gráfico
    figura, eixos = plt.subplots(figsize=(width / 100, height / 100))

    # Plotando a linha para cada Scenario do dicionário
    for index, (scenario_name, scenario_df) in enumerate(scenario_dataframes.items()):
        eixos.plot(scenario_df['Date'], scenario_df['Accum. Delivered Qty (Hyp)'], label=f'Scen. {index}',
                   color=colors_array[index])
        # Configurando o eixo
        plt.xticks(scenario_df.index[::3], scenario_df['Date'][::3], rotation=45, ha='right')

        # Obtendo a data t0 para o Scenario atual e convertendo para o formato MM/YYYY
        t0_date = pd.to_datetime(
            df_scope_with_scenarios.loc[df_scope_with_scenarios['Scenario'] == index, 't0'].values[0])
        t0_date = t0_date.strftime('%m/%Y')
        # Adicionando uma linha vertical em t0
        eixos.axvline(x=t0_date, linestyle='--', color=colors_array[index], label=f't0: Scen. {index}')

        # Obtendo a data acft_delivery_start para o Scenario atual e convertendo para o formato MM/YYYY
        acft_delivery_start_date = pd.to_datetime(
            df_scope_with_scenarios.loc[df_scope_with_scenarios['Scenario'] == index, 'acft_delivery_start'].values[0])
        acft_delivery_start_date = acft_delivery_start_date.strftime('%m/%Y')
        # Adicionando uma linha vertical em acft_delivery_start
        eixos.axvline(x=acft_delivery_start_date, linestyle='dotted', color=colors_array[index],
                      label=f'Acft Delivery Start: Scen. {index}')

        # Adicionando uma faixa de entrega dos materiais entre as data Início e Fim
        material_delivery_start_date = pd.to_datetime(df_scope_with_scenarios.loc[df_scope_with_scenarios[
                                                                                      'Scenario'] == index, 'material_delivery_start_date'].values[
                                                          0])
        material_delivery_start_date = material_delivery_start_date.strftime('%m/%Y')
        material_delivery_end_date = pd.to_datetime(df_scope_with_scenarios.loc[df_scope_with_scenarios[
                                                                                    'Scenario'] == index, 'material_delivery_end_date'].values[
                                                        0])
        material_delivery_end_date = material_delivery_end_date.strftime('%m/%Y')

        eixos.axvspan(material_delivery_start_date, material_delivery_end_date, alpha=0.5, color=colors_array[index])

        # Adicionando uma anotação no ponto em que o Build-Up é concluído (todos os itens entregues)
        index_max_acc_qty = scenario_df['Accum. Delivered Qty (Hyp)'].idxmax()
        x_max = scenario_df.loc[index_max_acc_qty, 'Date']
        y_max = scenario_df.loc[index_max_acc_qty, 'Accum. Delivered Qty (Hyp)']
        plt.scatter(x_max, y_max, color=colors_array[index], marker='o', label=f'BUP Conclusion: {x_max}')

    # Configuração do Gráfico
    eixos.set_ylabel('Materials Delivered Qty (Accumulated)')
    eixos.set_title('Hypothetical Curve: Build-Up Forecast')
    eixos.grid(True)

    # Ajustando espaçamento dos eixos para não cortar os rótulos
    plt.subplots_adjust(left=0.15, right=0.9, bottom=0.2, top=0.9)

    # Legenda
    eixos.legend(loc='upper left', fontsize=7, framealpha=0.8)

    # --------------- TRANSFORMANDO EM UMA IMAGEM PARA SER EXIBIDA ---------------

    # Salvando a figura matplotlib em um objeto BytesIO (memória), para não ter que salvar em um arquivo de imagem
    tmp_img_hypothetical_chart = BytesIO()
    figura.savefig(tmp_img_hypothetical_chart, format='png', transparent=True)
    tmp_img_hypothetical_chart.seek(0)

    # Carregando a imagem do gráfico para um objeto Image que irá ser retornado pela função
    bup_hypothetical_chart = Image.open(tmp_img_hypothetical_chart)

    return bup_hypothetical_chart
