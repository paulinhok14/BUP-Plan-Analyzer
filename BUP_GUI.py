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
app_path = os.path.dirname(os.path.realpath(__file__))
readme_url = r'https://github.com/paulinhok14/BUP-Plan-Analyzer/blob/master/README.md'
# readme_path = os.path.join(app_path, 'README.md')
# TO DO: Make README.md open local file using default web browser (or specific), independent of default ".md" openers from user OS

# These variables will keep the Last Acq Cost canvas showed, so that I can handle it on CTkSwitch toggle (AcqCost/Parts)
last_acq_cost_canvas_eff, last_acq_cost_canvas_hyp = None, None

def main():

    main_screen = ctk.CTk()  # Main Screen

    # Main Screen settings
    main_screen.title("Build-Up Plan Analyzer")
    main_screen.resizable(width=False, height=False)
    main_screen._set_appearance_mode("dark")
    main_screen.iconbitmap(main_screen_icon)
    main_screen.protocol("WM_DELETE_WINDOW", lambda: main_screen.quit())

    # Main Screen geometry - Centralizing
    ms_width = 700
    ms_height = 630
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

        # Another CTk variable. Now StringVar in order to handle with Acq Cost/Parts Qty charts exhibition. It starts with 'Parts (#)' and can be 'Acq Cost (US$)'
        # Each Window (Eff/Hyp) needs its particular StringVar, so as not to be necessary to bound/link both CTkSwitch toggles
        chart_mode_selection_eff = ctk.StringVar(value='Parts (#)')
        chart_mode_selection_hyp = ctk.StringVar(value='Parts (#)')

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
        nw_height = 630
        window_width = new_window.winfo_screenwidth()  # Width of the screen
        window_height = new_window.winfo_screenheight()  # Height of the screen
        # Calculates the initial X and Y to position the screen
        nw_x = (window_width / 2) - (nw_width / 2)
        nw_y = (window_height / 2) - (nw_height / 2)
        new_window.geometry('%dx%d+%d+%d' % (nw_width, nw_height, nw_x, nw_y))

        # TabView - Secondary screen elements: Tabs
        tbvmenu = ctk.CTkTabview(new_window, width=650, height=600, corner_radius=20,
                                 segmented_button_fg_color="#009898", segmented_button_unselected_color="#009898",
                                 segmented_button_selected_color="#006464", bg_color='#ebebeb', fg_color='#dbdbdb')

        tbvmenu.pack()
        tbvmenu.add("Scope")
        tbvmenu.add("Leadtime Analysis")
        tbvmenu.add("Scenarios")
        tbvmenu.add("Stock Analysis")

        @bup.function_timer
        def export_data(xl_spreadsheet: str) -> None:
            """ Function that will run when user click on "Export to Excel" button.
            param xl_spreadsheet: It contains the specific spreadsheet that will be exported.
            Possibilities: 'efficient_chart', 'hypothetical_chart' and 'scope'.
            All used variables is consumed from bup_plan_analyzer.py.
            """

            # Switch-case depending on which spreadsheet should be exported
            match xl_spreadsheet:
                case 'efficient_chart':
                    try:
                        full_path = export_output_path + r'\BUP_Eficient_Chart_Data.xlsx'
                        with pd.ExcelWriter(full_path) as writer:

                            eff_consolidated_scenarios = pd.DataFrame()
                            for scenario_name, scenario_df_list in bup.scenario_dataframes.items():
                                eff_consolidated_scenarios = pd.concat([eff_consolidated_scenarios, scenario_df_list[0]],
                                                                       ignore_index=True)

                            eff_consolidated_scenarios.to_excel(writer, sheet_name='Efficient Chart Data', index=False)
                        messagebox.showinfo(title="Success!", message=str("Excel sheet was exported to: " + full_path))

                    except Exception as ex:
                        messagebox.showinfo(title="Error!",
                                            message=str(ex) + "\n\nPlease make sure that the Excel file is "
                                                              "closed and you have access to the Downloads folder.")

                case 'hypothetical_chart':
                    try:
                        full_path = export_output_path + r'\BUP_Hypothetical_Chart_Data.xlsx'
                        with pd.ExcelWriter(full_path) as writer:

                            hyp_consolidated_scenarios = pd.DataFrame()
                            for scenario_name, scenario_df_list in bup.scenario_dataframes.items():
                                hyp_consolidated_scenarios = pd.concat(
                                    [hyp_consolidated_scenarios, scenario_df_list[1]],
                                    ignore_index=True)

                            hyp_consolidated_scenarios.to_excel(writer, sheet_name='Hypothetical Chart Data', index=False)
                        messagebox.showinfo(title="Success!", message=str("Excel sheet was exported to: " + full_path))

                    except Exception as ex:
                        messagebox.showinfo(title="Error!",
                                            message=str(ex) + "\n\nPlease make sure that the Excel file is "
                                                              "closed and you have access to the Downloads folder.")

                case 'scope':
                    '''
                    If "bup.df_scope_with_scenarios" is a DataFrame, it will give preference to export Scope with Scenarios info.
                    Else, only Scope information will be exported.
                    '''
                    if not isinstance(bup.df_scope_with_scenarios, pd.DataFrame):
                        try:
                            full_path = export_output_path + r'\BUP_Scope_Data.xlsx'
                            with pd.ExcelWriter(full_path) as writer:
                                bup_scope.to_excel(writer, sheet_name='Scope', index=False)
                            messagebox.showinfo(title="Success!", message=str("Excel sheet was exported to: " + full_path))
                        except Exception as ex:
                            messagebox.showinfo(title="Error!",
                                                message=str(ex) + "\n\nPlease make sure that the Excel file is "
                                                                  "closed and you have access to the Downloads folder.")
                    else:
                        try:
                            full_path = export_output_path + r'\BUP_Scope_with_Scenarios_Data.xlsx'
                            with pd.ExcelWriter(full_path) as writer:
                                bup.df_scope_with_scenarios.to_excel(writer, sheet_name='Scope with Scenarios', index=False)
                            messagebox.showinfo(title="Success!", message=str("Excel sheet was exported to: " + full_path))
                        except Exception as ex:
                            messagebox.showinfo(title="Error!",
                                                message=str(ex) + "\n\nPlease make sure that the Excel file is "
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
                                     compound="right",
                                     cursor="hand2"
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

        # Inserting Acq Cost x Leadtime Scatter and returning CTk Image object
        dispersion_image = bup.generate_dispersion_chart(bup_scope, tbvmenu.tab("Leadtime Analysis"))

        # Inserting Histogram Chart and returning CTk Image object
        histogram_image = bup.generate_histogram(bup_scope, tbvmenu.tab("Leadtime Analysis"))

        # Tab 3 - Scenarios

        # TabView - Segregation between Efficient Curve & Hypothetical Curve charts
        tbv_curve_charts = ctk.CTkTabview(tbvmenu.tab("Scenarios"), width=620, height=520, corner_radius=15,
                                          segmented_button_fg_color="#009898",
                                          segmented_button_unselected_color="#009898",
                                          segmented_button_selected_color="#006464",
                                          bg_color='#dbdbdb', fg_color='#cfcfcf')
        tbv_curve_charts.pack()
        tbv_curve_charts.add("Efficient Curve")
        tbv_curve_charts.add("Hypothetical Curve")
        tbv_curve_charts.add("Cost Avoidance")

        # Image with Excel icon
        excel_icon = ctk.CTkImage(light_image=Image.open(excel_icon_path),
                                  dark_image=Image.open(excel_icon_path),
                                  size=(30, 30))

        # Image with Download Icon
        download_icon = ctk.CTkImage(light_image=Image.open(download_icon_path),
                                  dark_image=Image.open(download_icon_path),
                                  size=(34, 34))

        def toggle_chart_mode_selection(str_var: ctk.StringVar, chart_name: str) -> None:
            '''
            This function handles CTkSwitch toggling, for both windows (Eff/Hyp). It receives StringVar that is being changed, and set the components
            such as a CTkLabel to obey the new Text and Color values. This would trigger the action to change the Chart being shown.
            :param str_var: The ctk.StringVar in which 'Parts'/'Acq Cost' are being stored.
            :param chart_name: String containing the name of the Window that is being changed. 'eff' or 'hyp'
            '''
            current_value = str_var.get()

            if current_value == 'Parts (#)':
                str_var.set('Acq Cost (US$)')

                # Updating Label and Switch properties for Acq Cost Selection
                if chart_name == 'eff':
                    swt_toggle_parts_acqcost_eff.configure(button_color='#004b00', progress_color='green')
                    lbl_chart_mode_eff.configure(text=str_var.get(), text_color='green')
                else:
                    swt_toggle_parts_acqcost_hyp.configure(button_color='#004b00', progress_color='green')
                    lbl_chart_mode_hyp.configure(text=str_var.get(), text_color='green')
            else:
                str_var.set('Parts (#)')

                # Updating Label and Switch properties for Parts Selection
                if chart_name == 'eff':
                    swt_toggle_parts_acqcost_eff.configure(button_color='orange', fg_color='#ad7102')
                    lbl_chart_mode_eff.configure(text=str_var.get(), text_color='#ad7102')
                else:
                    swt_toggle_parts_acqcost_hyp.configure(button_color='orange', fg_color='#ad7102')
                    lbl_chart_mode_hyp.configure(text=str_var.get(), text_color='#ad7102')

        # Switch in order to toggle between Parts/Acq Cost chart visualization
        swt_toggle_parts_acqcost_eff = ctk.CTkSwitch(tbv_curve_charts.tab("Efficient Curve"),
                                                     text="",
                                                     command=lambda: (toggle_chart_mode_selection(chart_mode_selection_eff, 'eff'),
                                                                      callback_func_chart_mode_toggle(chart_mode_selection_eff, 'eff',
                                                                                                      [bup.canvas_eff, bup.canvas_hyp, bup.canvas_list_acqcost_eff, bup.canvas_list_acqcost_hyp])),
                                                     width=45, height=12, button_color='orange', fg_color='#ad7102',
                                                     )
        swt_toggle_parts_acqcost_hyp = ctk.CTkSwitch(tbv_curve_charts.tab("Hypothetical Curve"),
                                                     text="",
                                                     command=lambda: (toggle_chart_mode_selection(chart_mode_selection_hyp, 'hyp'),
                                                                      callback_func_chart_mode_toggle(chart_mode_selection_hyp, 'hyp',
                                                                                                      [bup.canvas_eff, bup.canvas_hyp, bup.canvas_list_acqcost_eff, bup.canvas_list_acqcost_hyp])),
                                                     width=45, height=12, button_color='orange', fg_color='#ad7102'
                                                     )

        # Labels to show current selected StringVar value: AcqCost/Parts
        lbl_chart_mode_eff = ctk.CTkLabel(tbv_curve_charts.tab("Efficient Curve"),
                                          text=chart_mode_selection_eff.get(),
                                          font=ctk.CTkFont('open sans', size=10, weight='bold'),
                                          text_color='#ad7102')
        lbl_chart_mode_hyp = ctk.CTkLabel(tbv_curve_charts.tab("Hypothetical Curve"),
                                          text=chart_mode_selection_hyp.get(),
                                          font=ctk.CTkFont('open sans', size=10, weight='bold'),
                                          text_color='#ad7102')


        # ComboBoxes in order to handle the Acq Cost for different Scenarios. Each scenario produces a unique Acq Cost Chart.
        cbx_selected_scenario_eff = ctk.CTkComboBox(tbv_curve_charts.tab("Efficient Curve"),
                                                    height=20, width=130,
                                                    font=ctk.CTkFont('open sans', size=10, weight='bold'),
                                                    dropdown_font=ctk.CTkFont('open sans', size=10, weight='bold'),
                                                    state='readonly',
                                                    command=lambda value: (callback_func_chart_mode_toggle(chart_mode_selection_eff, 'eff',
                                                                                                      [bup.canvas_eff, bup.canvas_hyp, bup.canvas_list_acqcost_eff, bup.canvas_list_acqcost_hyp],
                                                                                                      cbx_triggered=True,
                                                                                                      cbx_selected=value)))
        cbx_selected_scenario_hyp = ctk.CTkComboBox(tbv_curve_charts.tab("Hypothetical Curve"),
                                                    height=20, width=130,
                                                    font=ctk.CTkFont('open sans', size=10, weight='bold'),
                                                    dropdown_font=ctk.CTkFont('open sans', size=10, weight='bold'),
                                                    state='readonly',
                                                    command=lambda value: (callback_func_chart_mode_toggle(chart_mode_selection_hyp, 'hyp',
                                                                                                      [bup.canvas_eff, bup.canvas_hyp, bup.canvas_list_acqcost_eff, bup.canvas_list_acqcost_hyp],
                                                                                                      cbx_triggered=True,
                                                                                                      cbx_selected=value)))


        # Function that will be called to evaluate the control variable and Show/Hide buttons (Export Date/Save Image)
        def callback_func_scenario_add(scenarios_count) -> None:

            global last_acq_cost_canvas_eff, last_acq_cost_canvas_hyp

            # When a Scenario is created, the ComboBox values will also be updated
            cbx_selected_scenario_eff.configure(values=list(bup.scenario_dataframes.keys()))
            cbx_selected_scenario_hyp.configure(values=list(bup.scenario_dataframes.keys()))


            if scenarios_count.get() == 1:
                # Show the Export to Excel/Save Image buttons and hide the Scenario Creation message
                btn_export_data_eff.place(relx=0.92, rely=0.93, anchor=ctk.CENTER)
                btn_export_data_hyp.place(relx=0.92, rely=0.93, anchor=ctk.CENTER)
                btn_save_image_eff.place(relx=0.08, rely=0.93, anchor=ctk.CENTER)
                btn_save_image_hyp.place(relx=0.08, rely=0.93, anchor=ctk.CENTER)
                swt_toggle_parts_acqcost_eff.place(relx=0.94, rely=0.02, anchor=ctk.CENTER)
                swt_toggle_parts_acqcost_hyp.place(relx=0.94, rely=0.02, anchor=ctk.CENTER)
                lbl_chart_mode_eff.place(relx=0.88, rely=0.02, anchor="e")
                lbl_chart_mode_hyp.place(relx=0.88, rely=0.02, anchor="e")
                lbl_pending_scenario.place_forget()

                # Only in the first time, the ComboBox 'placeholder' will have the first Scenario
                cbx_selected_scenario_eff.set(list(bup.scenario_dataframes.keys())[0])
                cbx_selected_scenario_hyp.set(list(bup.scenario_dataframes.keys())[0])

                # Also only when a first Scenario is created, I assign this first scenario charts to Acq Cost last showed charts
                last_acq_cost_canvas_eff = bup.canvas_list_acqcost_eff[0]
                last_acq_cost_canvas_hyp = bup.canvas_list_acqcost_hyp[0]

            else:
                pass


        def callback_func_chart_mode_toggle(chart_mode: ctk.StringVar, chart_name: str, mpl_canvas_list: list, cbx_triggered:bool=False, cbx_selected:str=None) -> None:
            '''
            This is the function that will handle the Charts in different mode exhibition.
            It will be called whenever a CTkSwitch is toggled, and based on the variable, it decides what Chart to show (Parts/Acq Cost)
            :param chart_mode: StringVar with the Chart Mode that is being triggered
            :param chart_name: Window being toggled (eff/hyp)
            :param mpl_canvas_list: List with the Chart canvas to be manipulated
            :param cbx_triggered: This arg will be True only when being called by ComboBox switching Acq Cost charts between scenarios.
            mpl_canvas_list indexs:
            0: Efficient Chart - Parts
            1: Hypothetical Chart - Parts
            2: List with Efficient Charts - Acq Cost, for each Scenario created.
            3: List with Hypothetical Charts - Acq Cost, for each Scenario created.
            '''

            global last_acq_cost_canvas_eff, last_acq_cost_canvas_hyp

            parts_eff_chart = mpl_canvas_list[0]
            parts_hyp_chart = mpl_canvas_list[1]
            list_acqcost_eff_charts = mpl_canvas_list[2]
            list_acqcost_hyp_charts = mpl_canvas_list[3]

            if cbx_triggered == False:

                match chart_name:
                    # Efficient Window: CTkSwitch toggled
                    case 'eff':
                        if chart_mode.get() == 'Acq Cost (US$)':
                            # Remove Parts Chart
                            parts_eff_chart.get_tk_widget().place_forget()
                            # Insert Scenario ComboBox and Acq Cost Chart
                            cbx_selected_scenario_eff.place(relx=0.13, rely=0.02, anchor=ctk.CENTER)
                            last_acq_cost_canvas_eff.get_tk_widget().place(relx=0.5, rely=0.46, anchor=ctk.CENTER)
                        else:
                            # Insert Parts Chart
                            parts_eff_chart.get_tk_widget().place(relx=0.5, rely=0.46, anchor=ctk.CENTER)
                            # Remove Scenario ComboBox and Acq Cost Chart
                            cbx_selected_scenario_eff.place_forget()
                            last_acq_cost_canvas_eff.get_tk_widget().place_forget()

                    # Hypothetical Window: CTkSwitch toggled
                    case 'hyp':
                        if chart_mode.get() == 'Acq Cost (US$)':
                            # Remove Parts Chart
                            parts_hyp_chart.get_tk_widget().place_forget()
                            # Insert Scenario ComboBox and Acq Cost Chart
                            cbx_selected_scenario_hyp.place(relx=0.13, rely=0.02, anchor=ctk.CENTER)
                            last_acq_cost_canvas_hyp.get_tk_widget().place(relx=0.5, rely=0.46, anchor=ctk.CENTER)
                        else:
                            # Insert Parts Chart
                            parts_hyp_chart.get_tk_widget().place(relx=0.5, rely=0.46, anchor=ctk.CENTER)
                            # Remove Scenario ComboBox and Acq Cost Chart
                            cbx_selected_scenario_hyp.place_forget()
                            last_acq_cost_canvas_hyp.get_tk_widget().place_forget()
                        
            elif cbx_triggered == True:
                # Getting the index to search on the list, based on the last char of value, ex: 0 for 'Scenario_0'
                chart_index = int(cbx_selected[-1])

                # Efficient Screen Switching
                if chart_name == 'eff':
                    last_acq_cost_canvas_eff.get_tk_widget().place_forget()
                    list_acqcost_eff_charts[chart_index].get_tk_widget().place(relx=0.5, rely=0.46, anchor=ctk.CENTER)
                    last_acq_cost_canvas_eff = list_acqcost_eff_charts[chart_index]

                # Hypothetical Screen Switching
                elif chart_name == 'hyp':
                    print('hyp',last_acq_cost_canvas_hyp) #test
                    last_acq_cost_canvas_hyp.get_tk_widget().place_forget()
                    list_acqcost_hyp_charts[chart_index].get_tk_widget().place(relx=0.5, rely=0.46, anchor=ctk.CENTER)
                    last_acq_cost_canvas_hyp =  list_acqcost_hyp_charts[chart_index]


        # Tracing Scenario creation variable and calling the respective functions every time the variable changes
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
            cost_avoidance_window = tbv_curve_charts.tab("Cost Avoidance")

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

            '''
            Function that handles Scenarios creation, creates the window elements and interacts with the Scenario List.
            It will also return the Canvas for each Chart, enabling toggle function to handle the exhibition
            '''
            bup.create_scenario(scenario_window, var_scenarios_count, bup_scope, efficient_curve_window, hypothetical_curve_window, cost_avoidance_window)

        # Create Scenario button
        btn_create_scenario = ctk.CTkButton(tbvmenu.tab("Scenarios"), text='Create Scenario',
                                      command=lambda: (open_form_add_new_scenario()),
                                      font=ctk.CTkFont('open sans', size=12, weight='bold'),
                                      bg_color="#cfcfcf", fg_color="#009898", hover_color="#006464",
                                      width=200, height=30, corner_radius=30
                                            )
        btn_create_scenario.place(relx=0.5, rely=0.93, anchor=ctk.CENTER)

        # Tab 4 - Stock Analysis

        btn_read_stock_data = ctk.CTkButton(tbvmenu.tab("Stock Analysis"), text='Read Stock Data',
                                      command=lambda: print('Stock read!'),
                                      font=ctk.CTkFont('open sans', size=12, weight='bold'),
                                      bg_color="#dbdbdb", fg_color="#009898", hover_color="#006464",
                                      width=200, height=30, corner_radius=30
                                      )
        btn_read_stock_data.pack()

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
                                    # command=lambda: webbrowser.open(f'file:///{readme_path}', new=2)
                                    command=lambda: webbrowser.open(readme_url, new=2)
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
