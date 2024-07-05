from tkinter import filedialog  # In order to open file with Explorer
import customtkinter as ctk
from PIL import Image
import os
import pandas as pd
from tksheet import Sheet
from tkinter import messagebox
import webbrowser

import bup_plan_analyzer as bup  # Source file with program functions

# App Settings
active_user = os.getlogin()
ctk.set_appearance_mode("light")  # Modes: system (default), light, dark
ctk.set_default_color_theme("blue")  # Themes: blue (default), dark-blue, green
# Version
app_version: str = 'Release: ' + 'v1.2'

# Reference Variables
logo_path = r'src\images\logo.png'
title_path = r'src\images\titulo_bup_analyzer.png'
img_open_file_path = r'src\images\browse_icon_transp.png'
help_button_path = r'src\images\help.png'
main_screen_icon = r'src\images\bup_icon.ico'
download_icon_path = r'src\images\download-green-arrow.png'
excel_icon_path = r'src\images\excel_transparent.png'
export_output_path = fr'C:\Users\{active_user}\Downloads'
readme_url = r'https://github.com/paulinhok14/BUP-Plan-Analyzer/blob/master/README.md'


def main():

    main_screen = ctk.CTk()  # Main Screen

    # Main Screen settings
    main_screen.title("Build-Up Plan Analyzer")
    main_screen.resizable(width=False, height=False)
    main_screen._set_appearance_mode("dark")
    # main_screen.configure(fg_color="#242424")
    main_screen.iconbitmap(main_screen_icon)
    main_screen.protocol("WM_DELETE_WINDOW", lambda: main_screen.quit())

    # Main Screen geometry - Centralizing
    ms_width = 700
    ms_height = 600
    screen_width = main_screen.winfo_screenwidth()  # Width of the screen
    screen_height = main_screen.winfo_screenheight()  # Height of the screen
    # Calculates the initial X and Y to position the screen
    x = (screen_width/2) - (ms_width/2)
    y = (screen_height/2) - (ms_height/2)
    main_screen.geometry('%dx%d+%d+%d' % (ms_width, ms_height, x, y))

    def select_file() -> str:  # Function to select the Scope file and return its path
        file_path = filedialog.askopenfilename()
        return file_path

    def create_new_window(title: str):  # Function to create new window
        # Configures the Execution warning label
        lbl_loading.configure(text="Please wait while complementary data is fetched...")

        # CTk variable that will store the Scenario count. It will be useful to implement tracking with callback
        # function monitoring it. To show or hide components
        var_scenarios_count = ctk.IntVar()

        # Reading the Scope file before creating the window
        full_file_path = select_file()  # Selecting Scope file

        # Executing the function that reads complementary info and organizes DataFrame
        bup_scope = bup.read_scope_file(full_file_path)

        # Resets the Execution label to empty text
        lbl_loading.configure(text="")

        # Creating window
        new_window = ctk.CTkToplevel(main_screen, fg_color='#ebebeb')

        # Function that will handle window closing
        def on_closing_main2() -> None:
            new_window.withdraw()
            main_screen.deiconify()
            # Reseting IntVar with Scenario counts
            var_scenarios_count.set(0)

        # New window settings
        new_window.title(title)
        new_window.resizable(width=False, height=False)
        new_window.iconbitmap(main_screen_icon)
        new_window.protocol("WM_DELETE_WINDOW", on_closing_main2)

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
                                 segmented_button_selected_color="#006464", bg_color='#ebebeb', fg_color='#dbdbdb')

        tbvmenu.pack()
        tbvmenu.add("Scope")
        tbvmenu.add("Leadtime Analysis")
        tbvmenu.add("Scenarios")

        @bup.function_timer
        def export_data(xl_spreadsheet: str) -> None:
            """ Function that will run when user click on "Export to Excel" button.
            param xl_spreadsheet: It contains the specific spreadsheet that will be exported.
            Possibilities: 'efficient_chart', 'hypothetical_chart' and 'scope'.
            All used variables is consumed from bup_plan_analyzer.py.
            """
            '''
            # Switch-case depending on which spreadsheet should be exported
            match xl_spreadsheet:
                case 'efficient_chart':
                    print('é o eficiente')
                    # try:
                    #     full_path = export_output_path + r'\efficient_chart_table.xlsx'

                case 'hypothetical_chart':
                    print('é o hipotetico')
                case 'scope':
                    print('é o escopo')
            '''
            try:
                full_path = export_output_path + r'\bup_scenarios_data.xlsx'
                with pd.ExcelWriter(full_path) as writer:
                    bup.df_scope_with_scenarios.to_excel(writer, sheet_name='BUP Scope with Scenarios', index=False)

                    consolidated_scenarios_df = pd.DataFrame()
                    for scenario, df in bup.scenario_dataframes.items():
                        consolidated_scenarios_df = pd.concat([consolidated_scenarios_df[0], df], ignore_index=True)

                    consolidated_scenarios_df.to_excel(writer, sheet_name='Scenarios Build-Up', index=False)
                messagebox.showinfo(title="Success!", message=str("Excel sheet was exported to: " + full_path))

            except Exception as ex:
                messagebox.showinfo(title="Error!", message=str(ex) + "\n\nPlease make sure that the Excel file is "
                                                                      "closed and you have access to the Downloads folder.")

        # Tab 1 - Scope
        bup_cost = "List Value: US$ " + str("{:.2f}".format(
            bup_scope.apply(lambda linha: linha['Acq Cost'] * linha['Qty'], axis=1).sum() / 1000000
        )) + " MM"

        label_cost = ctk.CTkLabel(tbvmenu.tab("Scope"), text=bup_cost,
                                  font=ctk.CTkFont('open sans', size=14, weight='bold'),
                                  text_color='#000000')
        label_cost.pack(anchor="e")

        # Image with Excel icon for Label
        img_xl_icon_label = ctk.CTkImage(light_image=Image.open(excel_icon_path),
                                  dark_image=Image.open(excel_icon_path),
                                  size=(20, 20))

        # Export to Excel Label
        lbl_export_xl = ctk.CTkLabel(tbvmenu.tab("Scope"),
                                     text='Export to Excel ',
                                     font=ctk.CTkFont('open sans', size=12, underline=True),
                                     text_color='#25a848',
                                     image=img_xl_icon_label,
                                     compound="right"
                                     )
        lbl_export_xl.place(rely=0, relx=0)
        lbl_export_xl.bind(sequence='<Button-1>', command=lambda _: export_data('scope'))

        # Creating Sheet object to display dataframe
        sheet = Sheet(tbvmenu.tab("Scope"), data=bup_scope.values.tolist(), headers=bup_scope.columns.tolist())
        sheet.pack(fill="both", expand=True)

        # Label with information: file name and rows number
        lbl_file_name = ctk.CTkLabel(tbvmenu.tab("Scope"), text="File: " + os.path.basename(full_file_path),
                                     font=ctk.CTkFont('open sans', size=10, weight='bold'),
                                     text_color='#000000')
        lbl_file_name.pack(side="left", padx=5)

        lbl_rows_count = ctk.CTkLabel(tbvmenu.tab("Scope"), text="Rows: " + str(bup_scope.shape[0]),
                                      font=ctk.CTkFont('open sans', size=10, weight='bold'),
                                      text_color='#000000')
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
                                          segmented_button_fg_color="#009898",
                                          segmented_button_unselected_color="#009898",
                                          segmented_button_selected_color="#006464",
                                          bg_color='#dbdbdb', fg_color='#cfcfcf')
        tbv_curve_charts.pack()
        tbv_curve_charts.add("Efficient Curve")
        tbv_curve_charts.add("Hypothetical Curve")

        # Image with Excel icon
        excel_icon = ctk.CTkImage(light_image=Image.open(excel_icon_path),
                                  dark_image=Image.open(excel_icon_path),
                                  size=(30, 30))

        # Image with Download Icon
        download_icon = ctk.CTkImage(light_image=Image.open(download_icon_path),
                                  dark_image=Image.open(download_icon_path),
                                  size=(34, 34))


        # Function that will be called to evaluate the control variable and Show/Hide buttons (Export Date/Save Image)
        def callback_func_scenario_add(scenarios_count) -> None:

            if scenarios_count.get() == 1:
                # Show the Export to Excel/Save Image buttons and hide the Scenario Creation message
                btn_export_data_eff.place(relx=0.92, rely=0.93, anchor=ctk.CENTER)
                btn_export_data_hyp.place(relx=0.92, rely=0.93, anchor=ctk.CENTER)
                btn_save_image_eff.place(relx=0.08, rely=0.93, anchor=ctk.CENTER)
                btn_save_image_hyp.place(relx=0.08, rely=0.93, anchor=ctk.CENTER)
                lbl_pending_scenario.place_forget()
            else:
                pass

        # Tracing the variable and calling the respective functions every time the variable changes
        var_scenarios_count.trace_add("write", callback=lambda *args: callback_func_scenario_add(var_scenarios_count))

        # Button that saves Efficient Chart Image
        btn_save_image_eff = ctk.CTkButton(tbv_curve_charts.tab("Efficient Curve"), text="Save Image",
                                           font=ctk.CTkFont('open sans', size=10, weight='bold'),
                                           image=download_icon, compound="top", fg_color="transparent",
                                           text_color="#000000", hover=False, border_spacing=0,
                                           command=lambda: (
                                               bup.save_chart_image(bup.img_eff_chart,
                                                                    export_output_path,
                                                                    "bup_efficient_chart.png"))
                                           )

        # Button that saves Hypothetical Chart Image
        btn_save_image_hyp = ctk.CTkButton(tbv_curve_charts.tab("Hypothetical Curve"), text="Save Image",
                                           font=ctk.CTkFont('open sans', size=10, weight='bold'),
                                           image=download_icon, compound="top", fg_color="transparent",
                                           text_color="#000000", hover=False, border_spacing=0,
                                           command=lambda: (
                                               bup.save_chart_image(bup.img_hyp_chart,
                                                                    export_output_path,
                                                                    "bup_hypothetical_chart.png"))
                                           )

        # Export data button - Efficient
        btn_export_data_eff = ctk.CTkButton(tbv_curve_charts.tab("Efficient Curve"), text="Export to Excel",
                                        font=ctk.CTkFont('open sans', size=10, weight='bold'),
                                        image=excel_icon, compound="top", fg_color="transparent",
                                        text_color="#000000", hover=False, border_spacing=1,
                                        command=lambda: (export_data('efficient_chart')))

        # Export data button - Hypothetical
        btn_export_data_hyp = ctk.CTkButton(tbv_curve_charts.tab("Hypothetical Curve"), text="Export to Excel",
                                        font=ctk.CTkFont('open sans', size=10, weight='bold'),
                                        image=excel_icon, compound="top", fg_color="transparent",
                                        text_color="#000000", hover=False, border_spacing=1,
                                        command=lambda: (export_data('hypothetical_chart')))

        # Label with the instruction to create a Scenario
        lbl_pending_scenario = ctk.CTkLabel(tbvmenu.tab("Scenarios"),
                                            text="Please create a Scenario in order to generate Build-Up chart.",
                                            font=ctk.CTkFont('open sans', size=16, weight='bold', slant='italic'),
                                            fg_color='#cfcfcf',
                                            bg_color='#cfcfcf',
                                            text_color='#000000')
        lbl_pending_scenario.place(rely=0.5, relx=0.5, anchor=ctk.CENTER)

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
            bup.create_scenario(scenario_window, var_scenarios_count, bup_scope, efficient_curve_window, hypothetical_curve_window)

        # Create Scenario button
        btn_create_scenario = ctk.CTkButton(tbvmenu.tab("Scenarios"), text='Create Scenario',
                                      command=lambda: (open_form_add_new_scenario()),
                                      font=ctk.CTkFont('open sans', size=12, weight='bold'),
                                      bg_color="#cfcfcf", fg_color="#009898", hover_color="#006464",
                                      width=200, height=30, corner_radius=30
                                            )
        btn_create_scenario.place(relx=0.5, rely=0.93, anchor=ctk.CENTER)

        # Hiding Main Screen
        main_screen.withdraw()

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
                                  )
    btnSearchFile.place(relx=0.5, rely=0.82, anchor=ctk.CENTER)

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
                        bg_color="#242424")
    lblLogo.place(relx=0.5, rely=0.45, anchor=ctk.CENTER)

    # Title CTkImage object
    image_title = ctk.CTkImage(light_image=Image.open(title_path),
                          dark_image=Image.open(title_path),
                          size=(620, 110))

    # Title put into the Label
    lblMainTitle = ctk.CTkLabel(main_screen, image=image_title,
                             text="",
                             bg_color="#242424")
    lblMainTitle.place(relx=0.5, rely=0.1, anchor=ctk.CENTER)

    # Version Label
    lbl_version = ctk.CTkLabel(main_screen,
                               text=app_version,
                               fg_color='#242424',
                               bg_color='#242424',
                               font=ctk.CTkFont('open sans', size=13, weight='bold'),
                               text_color='#ffffff')
    lbl_version.place(relx=0.02, rely=0.93)

    # "Help button" CTkImage object
    image_help = ctk.CTkImage(light_image=Image.open(help_button_path),
                              dark_image=Image.open(help_button_path),
                              size=(60, 60))
    # "Help button"
    btn_help = ctk.CTkButton(main_screen, image=image_help,
                                    text="", width=60, height=60,
                                    bg_color="#242424",
                                    fg_color='#242424',
                                    hover=False,
                                    command=lambda: webbrowser.open(readme_url)
                                    )
    btn_help.place(relx=0.885, rely=0.87)

    # MRSE button
    lbl_mrse = ctk.CTkLabel(main_screen, text="© MRSE 2024",
                            font=ctk.CTkFont('open sans', size=13),
                            text_color='#ffffff',
                            fg_color='#242424',
                            bg_color='#242424'
                            )
    lbl_mrse.place(relx=0.45, rely=0.93)

    main_screen.mainloop()  # Main Screen running loop


if __name__ == "__main__":
    main()
