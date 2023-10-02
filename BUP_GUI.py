from tkinter import filedialog  # Para abrir arquivo no Explorar
import customtkinter as ctk
import tkinter as tk
from PIL import Image
import os
from tkinter import ttk
from pandastable import Table, TableModel  # Para exibição de df pandas em objeto tinker

import bup_plan_analyzer as bup  # Arquivo-fonte com as funções do programa


logo_path = r'C:\Users\prsarau\PycharmProjects\BUP Plan Analyzer\logo.png'
title_path = r'C:\Users\prsarau\PycharmProjects\BUP Plan Analyzer\titulo_bup_analyzer.png'
img_open_file_path = r'C:\Users\prsarau\PycharmProjects\BUP Plan Analyzer\browse_icon_transp.png'


main_screen = ctk.CTk()  # Tela Principal

# Configurações da tela principal
main_screen.title("Build-Up Plan Analyzer")
main_screen.resizable(width=False, height=False)
main_screen._set_appearance_mode("dark")
main_screen.iconbitmap("bup_icon.ico")


# Geometria da tela Principal - Centralizando
ms_width = 700
ms_height = 550
screen_width = main_screen.winfo_screenwidth()  # Width of the screen
screen_height = main_screen.winfo_screenheight()  # Height of the screen
# Calcula o X e Y inicial a posicionar a tela
x = (screen_width/2) - (ms_width/2)
y = (screen_height/2) - (ms_height/2)
main_screen.geometry('%dx%d+%d+%d' % (ms_width, ms_height, x, y))


def select_file():  # Função para selecionar o arquivo do Escopo
    file_path = filedialog.askopenfilename()
    return file_path


def create_new_window(title: str):  # Função para criar nova janela
    # Lendo o arquivo de Escopo antes de criar a tela
    full_file_path = select_file()
    bup_scope = bup.read_scope_file(full_file_path)

    # Criando janela
    new_window = ctk.CTkToplevel(main_screen)

    # Configurações da nova tela
    new_window.title(title)
    new_window.resizable(width=False, height=False)
    new_window.iconbitmap("bup_icon.ico")

    # Geometria da Nova Tela
    nw_width = 700
    nw_height = 550
    window_width = new_window.winfo_screenwidth()  # Width of the screen
    window_height = new_window.winfo_screenheight()  # Height of the screen
    # Calcula o X e Y inicial a posicionar a tela
    nw_x = (window_width / 2) - (nw_width / 2)
    nw_y = (window_height / 2) - (nw_height / 2)
    new_window.geometry('%dx%d+%d+%d' % (nw_width, nw_height, nw_x, nw_y))

    # Ocultando a tela principal
    main_screen.iconify()

    # TabView - Elementos da tela secundária: Abas
    tbvmenu = ctk.CTkTabview(new_window, width=650, height=520, corner_radius=25,
                             segmented_button_fg_color="#009898", segmented_button_unselected_color="#009898",
                             segmented_button_selected_color="#006464")

    tbvmenu.pack()
    tbvmenu.add("Scope")
    tbvmenu.add("Leadtime Histogram")
    tbvmenu.add("Scenarios")

    # Aba 1 - Exibição da Tabela (TreeView)

    # Criando e posicionando Scrollbar
    scrollbar = ttk.Scrollbar(tbvmenu.tab("Scope"))
    scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

    treeview_bup_scope = ttk.Treeview(tbvmenu.tab("Scope"), yscrollcommand=scrollbar.set,
                                      columns=list(bup_scope.columns), show="headings",
                                      height=19)

    scrollbar.config(command=treeview_bup_scope.yview)

    # Nomeando as colunas
    for coluna in list(bup_scope.columns):
        treeview_bup_scope.column(coluna, minwidth=140, width=140, anchor="center")
        treeview_bup_scope.heading(coluna, text=coluna)

    # Inserindo as linhas na tabela
    for index, linha in bup_scope.iterrows():
        treeview_bup_scope.insert("", "end", values=(linha[0], linha[1], linha[2], linha[3]))

    treeview_bup_scope.pack()

    # Label com as informações: nome do arquivo e quantidade de registros
    lbl_file_name = ctk.CTkLabel(tbvmenu.tab("Scope"), text="File: " + os.path.basename(full_file_path),
                                 font=ctk.CTkFont('open sans', size=10, weight='bold'))
    lbl_file_name.pack(side="left", padx=5)

    lbl_rows_count = ctk.CTkLabel(tbvmenu.tab("Scope"), text="Rows: " + str(bup_scope.shape[0]),
                                  font=ctk.CTkFont('open sans', size=10, weight='bold'))
    lbl_rows_count.pack(side="right", padx=10)

    # Aba 2 - Histograma de Leadtimes

    # Gerando o Histograma e salvando a imagem do gráfico
    histogram_image, highest_leadimes = bup.generate_histogram(bup_scope)

    # Label com os maiores Leadtimes
    ctk.CTkLabel(tbvmenu.tab("Leadtime Histogram"),
                 text=highest_leadimes).pack()

    # Gráfico de Histograma - inputando no Label e posicionando na tela
    ctk.CTkLabel(tbvmenu.tab("Leadtime Histogram"), image=histogram_image,
                 text="").place(relx=0.5, rely=0.6, anchor=ctk.CENTER)

    # Aba 3 - Scenarios

    # Label com a instrução de criar um Scenario
    lbl_pending_scenario = ctk.CTkLabel(tbvmenu.tab("Scenarios"),
                                        text="Please create a Scenario in order to generate Build-Up chart.",
                                        font=ctk.CTkFont('open sans', size=16, weight='bold', slant='italic'))
    lbl_pending_scenario.place(rely=0.5, relx=0.07)

    # Função que cria a tela para Adicionar um Novo Scenario
    def open_form_add_new_scenario():
        # Criando janela
        scenario_window = ctk.CTkToplevel(tbvmenu.tab("Scenarios"))

        # Configurações da nova tela
        scenario_window.title("Add New Scenario")
        scenario_window.resizable(width=False, height=False)

        # Geometria da Nova Tela
        sw_width = 480
        sw_height = 550
        total_window_width = scenario_window.winfo_screenwidth()  # Width of the screen
        total_window_height = new_window.winfo_screenheight()  # Height of the screen
        # Calcula o X e Y inicial a posicionar a tela
        sw_x = (total_window_width / 2) - (sw_width / 2)
        sw_y = (total_window_height / 2) - (sw_height / 2)
        scenario_window.geometry('%dx%d+%d+%d' % (sw_width, sw_height, sw_x, sw_y))

        # Setando(captando) o foco para a janela
        scenario_window.grab_set()

        # Função que cria os elementos e interage com a Lista de Scenarios
        bup.create_scenario_test(scenario_window, bup_scope)

    # Botão Create Scenario
    btn_create_scenario = ctk.CTkButton(tbvmenu.tab("Scenarios"), text='Create Scenario',
                                  command=lambda: (lbl_pending_scenario.place_forget(), open_form_add_new_scenario()),
                                  font=ctk.CTkFont('open sans', size=12, weight='bold'),
                                  bg_color="#dbdbdb", fg_color="#009898", hover_color="#006464",
                                  width=200, height=30, corner_radius=30
                                        ).place(relx=0.5, rely=0.95, anchor=ctk.CENTER)


# Ícone botão Search File CTKImage
img_open_file = ctk.CTkImage(light_image=Image.open(img_open_file_path),
                             dark_image=Image.open(img_open_file_path),
                             size=(20, 20))

# Botão Search File
btnSearchFile = ctk.CTkButton(master=main_screen, text='Search Scope File',
                              command=lambda: (create_new_window("Build-Up Plan Analyzer")),
                              font=ctk.CTkFont('open sans', size=13, weight='bold'),
                              bg_color="#242424", fg_color="#009898", hover_color="#006464",
                              width=250, height=45, corner_radius=30,
                              image=img_open_file, compound="right"
                              ).place(relx=0.5, rely=0.82, anchor=ctk.CENTER)

# Logo objeto CTkImage
image_logo = ctk.CTkImage(light_image=Image.open(logo_path),
                      dark_image=Image.open(logo_path),
                      size=(270, 235))

# Logo inputada na Label
lblLogo = ctk.CTkLabel(main_screen, image=image_logo, text="",
                    bg_color="#242424").place(relx=0.5, rely=0.45, anchor=ctk.CENTER)

# Título objeto CTkImage
image_title = ctk.CTkImage(light_image=Image.open(title_path),
                      dark_image=Image.open(title_path),
                      size=(620, 110))

# Título inputado na Label
lblMainTitle = ctk.CTkLabel(main_screen, image=image_title,
                         text="",
                         bg_color="#242424").place(relx=0.5, rely=0.1, anchor=ctk.CENTER)


# main_screen.iconify() e forget(), deiconify()
main_screen.mainloop()  # Loop de execução Tela Principal

# rodar dps que tiver ativado venv: pip freeze > requirements.txt
