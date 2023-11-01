import pandas as pd
import warnings
import matplotlib.pyplot as plt
from PIL import Image
import customtkinter as ctk
from tkinter import messagebox
from datetime import datetime

warnings.filterwarnings("ignore")

# Lista de Scenarios
scenarios_list = []

# Declarando as variáveis que irão armazenar temporariamente os valores prévios de Cenários já cadastrados,
# para caso o usuário queira reaproveitar os parâmetros contratuais do Cenário.
t0_previous_value, acft_delivery_start_previous_value, material_delivery_start_previous_value\
    , material_delivery_end_previous_value = None, None, None, None


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


def create_scenario(scenario_window, bup_scope) -> None:
    global scenarios_list

    scenario = {}

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
                global t0_previous_value, acft_delivery_start_previous_value, material_delivery_start_previous_value \
                    , material_delivery_end_previous_value

                t0_previous_value = ctk.StringVar(value=
                                                  scenarios_list[0]['t0'].strftime("%d/%m/%Y")
                                                  )
                entry_t0.configure(textvariable=t0_previous_value)
                entry_t0.configure(state="disabled")

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
                global t0_previous_value, acft_delivery_start_previous_value, material_delivery_start_previous_value \
                    , material_delivery_end_previous_value

                t0_previous_value, acft_delivery_start_previous_value, material_delivery_start_previous_value \
                    , material_delivery_end_previous_value = None, None, None, None

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
    lbl_contractual_conditions.grid(row=0, columnspan=2, sticky="n", pady=(10, 10))

    # --- t0 ---
    lbl_t0 = ctk.CTkLabel(contractual_conditions_frame, text="T0 Date (*)",
                          font=ctk.CTkFont('open sans', size=11, weight="bold")
                          )
    lbl_t0.grid(row=1, column=0, sticky="w", padx=12)
    entry_t0 = ctk.CTkEntry(contractual_conditions_frame, width=200,
                            placeholder_text="Format: DD/MM/YYYY")
    entry_t0.grid(row=2, column=0, padx=10, sticky="w")

    # --- Aircraft Delivery Start ---
    lbl_acft_delivery_start = ctk.CTkLabel(contractual_conditions_frame, text="Aircraft Delivery Start (*)",
                                           font=ctk.CTkFont('open sans', size=11, weight='bold')
                                           )
    lbl_acft_delivery_start.grid(row=1, column=1, padx=12, sticky="e")
    entry_acft_delivery_start = ctk.CTkEntry(contractual_conditions_frame, width=200,
                                             placeholder_text="Format: DD/MM/YYYY")
    entry_acft_delivery_start.grid(row=2, column=1, padx=10, sticky="e")

    # --- Material Delivery Start ---
    lbl_material_delivery_start = ctk.CTkLabel(contractual_conditions_frame, text="Material Delivery Start (*)",
                                               font=ctk.CTkFont('open sans', size=11, weight='bold')
                                               )
    lbl_material_delivery_start.grid(row=3, column=0, sticky="w", padx=12)
    entry_material_delivery_start = ctk.CTkEntry(contractual_conditions_frame, width=200,
                                                 placeholder_text="Nº of months (ex: 38 for T+38)")
    entry_material_delivery_start.grid(row=4, column=0, padx=10, sticky="w", pady=(0, 20))

    # --- Material Delivery End ---
    lbl_material_delivery_end = ctk.CTkLabel(contractual_conditions_frame, text="Material Delivery End (*)",
                                             font=ctk.CTkFont('open sans', size=11, weight='bold')
                                             )
    lbl_material_delivery_end.grid(row=3, column=1, sticky="e", padx=12)
    entry_material_delivery_end = ctk.CTkEntry(contractual_conditions_frame, width=200,
                                               placeholder_text="Nº of months (ex: 43 for T+43)")
    entry_material_delivery_end.grid(row=4, column=1, padx=10, sticky="e", pady=(0, 20))

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

        # Chamando a função para gerar o gráfico de Build-Up
        generate_buildup_chart(bup_scope, scenarios_list)

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


def generate_buildup_chart(bup_scope, scenarios):

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

    # Pegando a Maior e Menor Data entre todas as datas possíveis, para delimitar o eixo X do gráfico
    date_columns = ['t0', 'acft_delivery_start', 'material_delivery_start_date', 'material_delivery_end_date',
                    'PN Order Date']

    # O primeiro min/max retorna a menor/maior data por coluna e no final temos então uma lista de mínimos/máximos.
    # O segundo pega o menor da lista criada.
    min_date = df_scope_with_scenarios[date_columns].min().min()
    max_date = df_scope_with_scenarios[date_columns].max().max()

    date_range = pd.date_range(start=min_date, end=max_date, freq='M')
    df_dates = pd.DataFrame({'Date': date_range.strftime('%m/%Y')})

    # # Para cada Scenario, criando um df separado com todas as datas e respectivas ordens.
    # for index, scenario in enumerate(scenarios):
    #     print(index)
    #     print(scenario)

    # Criando tabela com a contagem agrupada de itens comprados por mês, para cada Scenario
    grouped_counts = df_scope_with_scenarios.groupby([df_scope_with_scenarios['PN Order Date'].dt.to_period('M'), 'Scenario']).size().reset_index(
        name='Ordered Qty')
    # Formatando a Data YYYY-MM da tabela atual para o modelo de apresentação: MM/YYYY
    grouped_counts['PN Order Date'] = grouped_counts['PN Order Date'].dt.strftime('%m/%Y')

    # Passando a informação de Ordered Qty agrupada por mês e por Scenario para o DF com o Range de datas
    final_df_scenarios = df_dates.merge(grouped_counts, left_on='Date', right_on='PN Order Date', how='left')
    final_df_scenarios['Scenario'] = final_df_scenarios['Scenario'].fillna(-1)
    final_df_scenarios['Scenario'] = final_df_scenarios['Scenario'].astype(int)
    final_df_scenarios['Ordered Qty'] = final_df_scenarios['Ordered Qty'].fillna(0)

    print(final_df_scenarios)


    #df_scope_with_scenarios.to_excel("scope with scenarios.xlsx")

    # Criando um DF para cada Scenario




    # --------------- GERAÇÃO DO GRÁFICO ---------------
