from tkinter import filedialog  # In order to open file with Explorer
import customtkinter as ctk
from PIL import Image
import os
from tksheet import Sheet

import bup_plan_analyzer as bup  # Source file with program functions

logo_path = r'src\images\logo.png'
title_path = r'src\images\titulo_bup_analyzer.png'
img_open_file_path = r'src\images\browse_icon_transp.png'
main_screen_icon = r'src\images\bup_icon.ico'


def main():

    main_screen = ctk.CTk()  # Main Screen

    # Main Screen settings
    main_screen.title("Build-Up Plan Analyzer")
    main_screen.resizable(width=False, height=False)
    main_screen._set_appearance_mode("dark")
    main_screen.iconbitmap(main_screen_icon)

    # Main Scren geometry - Centralizing
    ms_width = 700
    ms_height = 600
    screen_width = main_screen.winfo_screenwidth()  # Width of the screen
    screen_height = main_screen.winfo_screenheight()  # Height of the screen
    # Calculates the initial X and Y to position the screen
    x = (screen_width/2) - (ms_width/2)
    y = (screen_height/2) - (ms_height/2)
    main_screen.geometry('%dx%d+%d+%d' % (ms_width, ms_height, x, y))

    def select_file():  # Function to select the Scope file
        file_path = filedialog.askopenfilename()
        return file_path

    def create_new_window(title: str):  # Function to create new window
        # Configures the Execution warning label
        lbl_loading.configure(text="Please wait while complementary data is fetched...")

        # Reading the Scope file before creating the window
        # Selecting Scope file
        full_file_path = select_file()

        # Executing the function that reads complementary info and organizes DataFrame
        bup_scope = bup.read_scope_file(full_file_path)

        # Resets the Execution label to empty text
        lbl_loading.configure(text="")

        # Creating window
        new_window = ctk.CTkToplevel(main_screen)

        # New window settings
        new_window.title(title)
        new_window.resizable(width=False, height=False)
        new_window.iconbitmap(main_screen_icon)

        # New window geometry
        nw_width = 700
        nw_height = 600
        window_width = new_window.winfo_screenwidth()  # Width of the screen
        window_height = new_window.winfo_screenheight()  # Height of the screen
        # Calculates the initial X and Y to position the screen
        nw_x = (window_width / 2) - (nw_width / 2)
        nw_y = (window_height / 2) - (nw_height / 2)
        new_window.geometry('%dx%d+%d+%d' % (nw_width, nw_height, nw_x, nw_y))

        # TabView - Secondary screen elements: Tabs
        tbvmenu = ctk.CTkTabview(new_window, width=650, height=570, corner_radius=20,
                                 segmented_button_fg_color="#009898", segmented_button_unselected_color="#009898",
                                 segmented_button_selected_color="#006464")

        tbvmenu.pack()
        tbvmenu.add("Scope")
        tbvmenu.add("Leadtime Analysis")
        tbvmenu.add("Scenarios")

        # Tab 1 - Scope
        bup_cost = "List Value: US$ " + str("{:.2f}".format(
            bup_scope.apply(lambda linha: linha['Acq Cost'] * linha['Qty'], axis=1).sum() / 1000000
        )) + " MM"

        label_cost = ctk.CTkLabel(tbvmenu.tab("Scope"), text=bup_cost,
                                  font=ctk.CTkFont('open sans', size=14, weight='bold'))
        label_cost.pack(anchor="e")

        # Creating Sheet object to display dataframe
        sheet = Sheet(tbvmenu.tab("Scope"), data=bup_scope.values.tolist(), headers=bup_scope.columns.tolist())
        sheet.pack(fill="both", expand=True)

        # Label with information: file name and rows number
        lbl_file_name = ctk.CTkLabel(tbvmenu.tab("Scope"), text="File: " + os.path.basename(full_file_path),
                                     font=ctk.CTkFont('open sans', size=10, weight='bold'))
        lbl_file_name.pack(side="left", padx=5)

        lbl_rows_count = ctk.CTkLabel(tbvmenu.tab("Scope"), text="Rows: " + str(bup_scope.shape[0]),
                                      font=ctk.CTkFont('open sans', size=10, weight='bold'))
        lbl_rows_count.pack(side="right", padx=10)

        # Tab 2 - Leadtime Analysis

        # Generating the Acq Cost x Leadtime scatterplot (dispersion chart)
        dispersion_chart = bup.generate_dispersion_chart(bup_scope)

        # Loading the function return into a CTk Image object
        img_dispersion_chart = ctk.CTkImage(dispersion_chart,
                                     dark_image=dispersion_chart,
                                     size=(600, 220))

        # Dispersion Chart - inputting the label and positioning it on the screen
        ctk.CTkLabel(tbvmenu.tab("Leadtime Analysis"), image=img_dispersion_chart,
                     text="").pack(pady=20, anchor="e")

        # Generating the Histogram and saving the chart image
        histogram_image, highest_leadimes = bup.generate_histogram(bup_scope)

        # Histogram Chart - inputting the Label and positioning it on the screen
        ctk.CTkLabel(tbvmenu.tab("Leadtime Analysis"), image=histogram_image,
                     text="").pack()

        # Tab 3 - Scenarios

        # TabView - Segregation between Efficient Curve & Hypothetical Curve charts
        tbv_curve_charts = ctk.CTkTabview(tbvmenu.tab("Scenarios"), width=620, height=490, corner_radius=15,
                                 segmented_button_fg_color="#009898", segmented_button_unselected_color="#009898",
                                 segmented_button_selected_color="#006464")
        tbv_curve_charts.pack()
        tbv_curve_charts.add("Efficient Curve")
        tbv_curve_charts.add("Hypothetical Curve")

        # Label with the instruction to create a Scenario
        lbl_pending_scenario = ctk.CTkLabel(tbvmenu.tab("Scenarios"),
                                            text="Please create a Scenario in order to generate Build-Up chart.",
                                            font=ctk.CTkFont('open sans', size=16, weight='bold', slant='italic'),
                                            fg_color='#cfcfcf')
        lbl_pending_scenario.place(rely=0.5, relx=0.07)

        # Function that creates the window to Add a New Scenario
        def open_form_add_new_scenario():
            # Creating window
            scenario_window = ctk.CTkToplevel(tbvmenu.tab("Scenarios"))
            efficient_curve_window = tbv_curve_charts.tab("Efficient Curve")
            hypothetical_curve_window = tbv_curve_charts.tab("Hypothetical Curve")

            # Window settings
            scenario_window.title("Add New Scenario")
            scenario_window.resizable(width=False, height=False)

            # Window geometry
            sw_width = 480
            sw_height = 550
            total_window_width = scenario_window.winfo_screenwidth()  # Width of the screen
            total_window_height = new_window.winfo_screenheight()  # Height of the screen
            # Calculates the initial X and Y to position the screen
            sw_x = (total_window_width / 2) - (sw_width / 2)
            sw_y = (total_window_height / 2) - (sw_height / 2)
            scenario_window.geometry('%dx%d+%d+%d' % (sw_width, sw_height, sw_x, sw_y))

            # Setting (capturing) focus to the window
            scenario_window.grab_set()

            # Function that creates the window elements and interacts with the Scenario List
            bup.create_scenario(scenario_window, bup_scope, efficient_curve_window, hypothetical_curve_window, lbl_pending_scenario)

        # Create Scenario button
        btn_create_scenario = ctk.CTkButton(tbvmenu.tab("Scenarios"), text='Create Scenario',
                                      command=lambda: (open_form_add_new_scenario()),
                                      font=ctk.CTkFont('open sans', size=12, weight='bold'),
                                      bg_color="#cfcfcf", fg_color="#009898", hover_color="#006464",
                                      width=200, height=30, corner_radius=30
                                            ).place(relx=0.5, rely=0.93, anchor=ctk.CENTER)

        # Hiding Main Screen
        main_screen.iconify()
        # main_screen.withdraw()

    # Search File CTKImage button icon
    img_open_file = ctk.CTkImage(light_image=Image.open(img_open_file_path),
                                 dark_image=Image.open(img_open_file_path),
                                 size=(20, 20))

    # Search File button
    btnSearchFile = ctk.CTkButton(master=main_screen, text='Search Scope File',
                                  command=lambda: (create_new_window("Build-Up Plan Analyzer")),
                                  font=ctk.CTkFont('open sans', size=13, weight='bold'),
                                  bg_color="#242424", fg_color="#009898", hover_color="#006464",
                                  width=250, height=45, corner_radius=30,
                                  image=img_open_file, compound="right", cursor="hand2"
                                  ).place(relx=0.5, rely=0.82, anchor=ctk.CENTER)

    # Loading label - will be displayed while the file and related information are being read
    lbl_loading = ctk.CTkLabel(master=main_screen, text='', fg_color='#242424', bg_color='#242424',
                               font=ctk.CTkFont('open sans', size=13, weight='bold'), text_color='#ffff00')
    lbl_loading.place(relx=0.5, rely=0.7, anchor=ctk.CENTER)

    # Logo object CTkImage
    image_logo = ctk.CTkImage(light_image=Image.open(logo_path),
                          dark_image=Image.open(logo_path),
                          size=(270, 235))

    # Logo put into the Label
    lblLogo = ctk.CTkLabel(main_screen, image=image_logo, text="",
                        bg_color="#242424").place(relx=0.5, rely=0.45, anchor=ctk.CENTER)

    # Title CTkImage object
    image_title = ctk.CTkImage(light_image=Image.open(title_path),
                          dark_image=Image.open(title_path),
                          size=(620, 110))

    # Title put into the Label
    lblMainTitle = ctk.CTkLabel(main_screen, image=image_title,
                             text="",
                             bg_color="#242424").place(relx=0.5, rely=0.1, anchor=ctk.CENTER)

    main_screen.mainloop()  # Main Screen running loop


if __name__ == "__main__":
    main()
