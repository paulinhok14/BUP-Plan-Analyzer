import pandas as pd
import warnings
import matplotlib.pyplot as plt
from PIL import Image
import customtkinter as ctk
from tkinter import messagebox


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
def create_scenario(scenario_window) -> None:

    global scenarios_list

    scenario = {}

    if not scenarios_list:

        # Caso não tenha nenhum elemento na lista, segue o procedimento normal para cadastro.

        enabled_contractual_conditions_entries = False

        # ----------------- CONTRACTUAL CONDITIONS -----------------

        lbl_contractual_conditions = ctk.CTkLabel(scenario_window, text="Contractual Conditions",
                                                  font=ctk.CTkFont('open sans', size=14, weight='bold')
                                                  ).pack(pady=10)

        # t0
        lbl_t0 = ctk.CTkLabel(scenario_window, text="T0 Date (*)",
                              font=ctk.CTkFont('open sans', size=11, weight="bold")
                              )
        lbl_t0.pack(anchor="w", padx=35)
        entry_t0 = ctk.CTkEntry(scenario_window, width=350,
                                placeholder_text="Format: DD/MM/YYYY")
        entry_t0.pack()

        # Aircraft Delivery Start
        lbl_acft_delivery_start = ctk.CTkLabel(scenario_window, text="Aircraft Delivery Start (*)",
                                            font=ctk.CTkFont('open sans', size=11, weight='bold')
                                               )
        lbl_acft_delivery_start.pack(anchor="w", padx=35)
        entry_acft_delivery_start = ctk.CTkEntry(scenario_window, width=350,
                                                 placeholder_text="Format: DD/MM/YYYY")
        entry_acft_delivery_start.pack()

        # Material Delivery Start
        lbl_material_delivery_start = ctk.CTkLabel(scenario_window, text="Material Delivery Start (*)",
                                            font=ctk.CTkFont('open sans', size=11, weight='bold')
                                                   ).pack(anchor="w", padx=35)
        entry_material_delivery_start = ctk.CTkEntry(scenario_window, width=350,
                                                 placeholder_text="Format: DD/MM/YYYY").pack()

        # Material Delivery End
        lbl_material_delivery_end = ctk.CTkLabel(scenario_window, text="Material Delivery End (*)",
                                            font=ctk.CTkFont('open sans', size=11, weight='bold')
                                                 ).pack(anchor="w", padx=35)
        entry_material_delivery_end = ctk.CTkEntry(scenario_window, width=350,
                                                 placeholder_text="Format: DD/MM/YYYY").pack()

        # ----------------- PROCUREMENT LENGTH -----------------

        lbl_procurement_length = ctk.CTkLabel(scenario_window, text="Procurement Length",
                                              font=ctk.CTkFont('open sans', size=14, weight='bold')
                                              ).pack(pady=10)

        # PR Release and Approval VSS
        lbl_pr_release_approval_vss = ctk.CTkLabel(scenario_window, text="PR Release and Approval VSS (in days)",
                                                 font=ctk.CTkFont('open sans', size=11, weight='bold')
                                                 ).pack(anchor="w", padx=35)
        entry_pr_release_approval_vss = ctk.CTkEntry(scenario_window, width=350,
                                                   placeholder_text="Default: 5 days").pack()

        # PO Commercial Condition
        lbl_po_commercial_condition = ctk.CTkLabel(scenario_window, text="PO Commercial Condition (in days)",
                                                   font=ctk.CTkFont('open sans', size=11, weight='bold')
                                                   ).pack(anchor="w", padx=35)
        entry_po_commercial_condition = ctk.CTkEntry(scenario_window, width=350,
                                                     placeholder_text="Default: 30 days").pack()

        # PO Conversion
        lbl_po_conversion = ctk.CTkLabel(scenario_window, text="PO Conversion (in days)",
                                         font=ctk.CTkFont('open sans', size=11, weight='bold')
                                         ).pack(anchor="w", padx=35)
        entry_po_conversion = ctk.CTkEntry(scenario_window, width=350,
                                           placeholder_text="Default: 30 days").pack()

        # Export License
        lbl_export_license = ctk.CTkLabel(scenario_window, text="Export License (in days) [For Controlled Items Only.]",
                                         font=ctk.CTkFont('open sans', size=11, weight='bold')
                                          ).pack(anchor="w", padx=35)
        entry_export_license = ctk.CTkEntry(scenario_window, width=350,
                                           placeholder_text="Default: 0 days").pack()

        # ----------------- BOTÕES DE INTERAÇÃO -----------------

        # Botão OK
        btn_ok = ctk.CTkButton(scenario_window, text='OK',
                            font=ctk.CTkFont('open sans', size=12, weight='bold'),
                            bg_color="#ebebeb", fg_color="#009898", hover_color="#006464",
                            width=60, height=30, corner_radius=30
                            ).place(relx=0.3, rely=0.95, anchor=ctk.CENTER)

        # Botão Cancelar
        btn_cancel = ctk.CTkButton(scenario_window, text='Cancel',
                                   font=ctk.CTkFont('open sans', size=12, weight='bold'),
                                   bg_color="#ebebeb", fg_color="#ff0000", hover_color="#af0003",
                                   width=100, height=30, corner_radius=30
                                   ).place(relx=0.7, rely=0.95, anchor=ctk.CENTER)


def create_scenario_test(scenario_window, bup_scope) -> None:

    global scenarios_list

    scenario = {}

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

    # ----------------- BOTÕES DE INTERAÇÃO -----------------

    # Função para retornar os valores digitados pelo usuário nos Entry, com tratativa para Defaults
    def get_entry_values():

        # --------- Contractual Conditions ---------

        # t0
        # Verifica se o valor é uma data válida antes de atribuir à variável
        date_t0_value = pd.to_datetime(entry_t0.get(), format='%d/%m/%Y', errors='coerce')
        if not pd.isna(date_t0_value):
            scenario['t0'] = date_t0_value
        else:
            messagebox.showerror("Error",
                                 "Invalid date. Please enter a valid date format for T0.")
            return

        # Aircraft Delivery Start
        # Verifica se o valor é uma data válida antes de atribuir à variável
        date_acft_delivery_value = pd.to_datetime(entry_acft_delivery_start.get(), format='%d/%m/%Y', errors='coerce')
        if not pd.isna(date_acft_delivery_value):
            scenario['acft_delivery_start'] = date_acft_delivery_value
        else:
            messagebox.showerror("Error",
                                 "Invalid date. Please enter a valid date format for Aircraft Delivery Start.")
            return

        # Material Delivery Start
        try:
            scenario['material_delivery_start'] = int(entry_material_delivery_start.get())
        except ValueError:
            messagebox.showerror("Error",
                                 "Invalid character. Please enter a valid number for Material Delivery Start.")
            return

        # Material Delivery End
        try:
            scenario['material_delivery_end'] = int(entry_material_delivery_end.get())
        except ValueError:
            messagebox.showerror("Error",
                                 "Invalid character. Please enter a valid number for Material Delivery End.")
            return

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
                           width=100, height=30, corner_radius=30
                           ).place(relx=0.3, rely=0.92, anchor=ctk.CENTER)

    # Botão Cancelar
    btn_cancel = ctk.CTkButton(scenario_window, text='Cancel', command=scenario_window.destroy,
                               font=ctk.CTkFont('open sans', size=12, weight='bold'),
                               bg_color="#ebebeb", fg_color="#ff0000", hover_color="#af0003",
                               width=100, height=30, corner_radius=30
                               ).place(relx=0.7, rely=0.92, anchor=ctk.CENTER)


def generate_buildup_chart(bup_scope, scenarios):

    # Lista para armazenar as combinações de Cenários e Escopo
    combinations = []

    for _, row in bup_scope.iterrows():
        # Percorrendo cada elemento da lista de dicionários (cada elemento um Cenário em scenarios_list)
        for index, scenario in enumerate(scenarios):
            # Combinando os valores do dataframe (escopo) com os valores do dicionário (cenário)
            comb = {**row, 'Scenario': index, **scenario}
            combinations.append(comb)

    # Criando um novo dataframe com as combinações de Escopo e Cenários juntos
    df_scope_with_cenarios = pd.DataFrame(combinations).sort_values(by='Scenario').reset_index()
    df_scope_with_cenarios['avg_month_diff'] = ((df_scope_with_cenarios['material_delivery_end']
                                                 - df_scope_with_cenarios['material_delivery_start']) / 2).astype(int)
