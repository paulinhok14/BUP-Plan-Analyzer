import pandas as pd
import warnings
import matplotlib.pyplot as plt
from matplotlib.ticker import FuncFormatter
import matplotlib.lines as mlines
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import numpy as np
from PIL import Image
from io import BytesIO
import json
import customtkinter as ctk
from tkinter import messagebox
import time
import logging
import mplcursors as mpc

warnings.filterwarnings("ignore")


# Scenarios list
scenarios_list = []

# Declaring the variables that will temporarily store the previous values of already registered Scenarios,
# in case the user wants to reuse the Contractual parameters of the Scenario.
t0_previous_value, hyp_t0_previous_value, acft_delivery_start_previous_value, material_delivery_start_previous_value\
    , material_delivery_end_previous_value = None, None, None, None, None

# Defining global variables that will store Efficient and Hypothetical chart Images when running internal function
img_eff_chart, img_hyp_chart = None, None

# Defining global variables that will store pandas DataFrames that will be exported in main file
# Creating also a dictionary to store scenario DataFrames
df_scope_with_scenarios, scenario_dataframes = None, {}

# Log Configs
open('execution_info.log', 'w').close()  # Clean log file before system execution
log_format = "%(asctime)s: %(levelname)s: %(message)s"
logging.basicConfig(level=logging.INFO,
                    filename='execution_info.log',
                    format=log_format)

# Setting matplotlib log level higher in order to ignore INFO logs on file
plt.set_loglevel('WARNING')


# Decorator function that calculates how long each function of the system takes to execute
def function_timer(func):
    def wrapper(*args):
        start_time = time.time()
        result = func(*args)
        exec_time = time.time() - start_time
        logging.info(f"Function '{func.__name__}' took {round(exec_time, 2)} seconds to run.")
        return result
    return wrapper


@function_timer
def read_scope_file(file_full_path: str) -> pd.DataFrame():
    # Function that reads scope file and complementary info

    # Resetting the list of scenarios every time the e-mail is read
    global scenarios_list
    scenarios_list = []

    # Columns to read from the Scope file (essential)
    colunas = ['PN', 'ECODE', 'QTY', 'EIS', 'SPC']

    # Complementary info Source
    # leadtime_source = r'\\sjkfs05\vss\GMT\40. Stock Efficiency\J - Operational Efficiency\006 - Srcfiles\003 - SAP\marcsa.txt'
    leadtime_source = r'marcsa.txt'
    # ecode_data_path = r'\\egmap20038-new\Databases\DB_Ecode-Data.txt'
    ecode_data_path = r'DB_Ecode-Data.txt'

    # File reading
    scope = pd.read_excel(file_full_path, usecols=colunas)

    # Filters
    scope.loc[:, 'QTY'] = scope['QTY'].fillna(0).astype(int)
    scope_filtered = scope.query("QTY > 0").copy()

    # Formatting
    scope_filtered.loc[:, 'ECODE'] = scope_filtered['ECODE'].fillna(0).astype(int)
    scope_filtered['EIS'] = scope_filtered['EIS'].fillna('')

    # -------------- Fetch for complementary info (Leadtime, ECCN, Acq Cost, Repairability, etc) -------------

    # Columns to read from SAP
    sap_source_columns = ['Material(MATNR)', 'PrzEntrPrev.(PLIFZ)']
    # Reading leadtime database (SAP)
    leadtimes = pd.read_csv(leadtime_source, usecols=sap_source_columns, encoding='latin', sep='|', low_memory=False)
    # Removing nulls
    leadtimes = leadtimes.dropna()
    # Renaming columns
    leadtimes.rename(columns={'Material(MATNR)': 'ECODE', 'PrzEntrPrev.(PLIFZ)': 'LEADTIME'}, inplace=True)
    # Making sure ECODES are int. OBS: replacing "!" char on some Materials
    leadtimes['ECODE'] = leadtimes['ECODE'].str.replace('!', '').astype(int)


    # Joining leadtimes to materials
    bup_scope = scope_filtered.merge(leadtimes, on='ECODE', how='left')

    # Columns to read from Ecode Data
    ecode_data_columns = ['ECODE', 'ACQCOST', 'ENGDESC']

    # Getting for each Ecode the index of the record that has the highest Acq Cost (premise for duplicates)
    ecode_data = pd.read_csv(ecode_data_path, usecols=ecode_data_columns).drop_duplicates()
    ecode_data_max_acqcost = ecode_data.groupby('ECODE')['ACQCOST'].idxmax()
    ecode_data_filtered = ecode_data.loc[ecode_data_max_acqcost].reset_index(drop=True)

    # Making Material Type rule (Repairable/Expendable)
    bup_scope['SPC'] = bup_scope['SPC'].apply(
        lambda x: 'Repairable' if x in [2, 6] else 'Expendable'
    )

    # Converting numeric columns from float to int
    bup_scope['ECODE'] = bup_scope['ECODE'].astype(int)
    bup_scope['QTY'] = bup_scope['QTY'].astype(int)
    bup_scope['LEADTIME'] = bup_scope['LEADTIME'].fillna(127).astype(int)  # Leadtime default 127
    # Ensuring that Acq Cost is floating
    ecode_data_filtered['ACQCOST'] = ecode_data_filtered['ACQCOST'].str.replace(',', '.').astype(float)

    # Joining Ecode Data info
    # Acq Cost
    bup_scope = bup_scope.merge(ecode_data_filtered[['ECODE', 'ACQCOST', 'ENGDESC']], how='left', on='ECODE')

    # Ordering by Leadtime descending
    bup_scope = bup_scope.sort_values('LEADTIME', ascending=False).reset_index(drop=True)

    # Renaming columns
    bup_scope.rename(columns={'ECODE': 'Ecode', 'QTY': 'Qty', 'LEADTIME': 'Leadtime',
                              'EIS': 'EIS Critical', 'ACQCOST': 'Acq Cost', 'ENGDESC': 'Description'}, inplace=True)

    # Reordering columns
    columns_order = ['PN', 'Ecode', 'Description', 'Qty', 'SPC', 'Leadtime', 'Acq Cost', 'EIS Critical']
    bup_scope = bup_scope.reindex(columns_order, axis=1)

    return bup_scope


@function_timer
def generate_dispersion_chart(bup_scope: pd.DataFrame, root: ctk.CTkFrame):
    # This function receives 'bup_scope' paramter as a pandas DataFrame and creates the chart.

    # Function to format y-axis values in thousands
    def format_acq_cost(value, _):
        return f'US$ {value / 1000:.0f}k'

    # Image Size
    width, height = 600, 220
    fig, ax = plt.subplots(figsize=(width / 100, height / 100), layout='constrained')  # Layout property that handles "cutting" axes labels
    # Keeping background transparent
    fig.patch.set_facecolor("None")
    fig.patch.set_alpha(0)
    ax.set_facecolor('None')

    # Scattering Items
    conditional_colors = np.where(bup_scope['SPC'] == 'Expendable', 'orange', 'purple')
    scatter = ax.scatter(bup_scope['Leadtime'], bup_scope['Acq Cost'], color=conditional_colors, marker='o', alpha=0.7)

    # Custom Labels Patches
    expendable_patch = mlines.Line2D([], [], marker='o', color='orange', markersize=6, lw=0, label='Expendable')
    repairable_patch = mlines.Line2D([], [], marker='o', color='purple', markersize=6, lw=0, label='Repairable')

    # Adding labels and legends
    ax.set_xlabel('Leadtime', loc='right')
    ax.set_ylabel('Acq Cost')
    ax.set_title('Dispersion Acq Cost x Leadtime', fontsize=10)
    ax.legend(handles=[expendable_patch, repairable_patch], fontsize=9, framealpha=0.6)
    plt.grid(True)
    # Setting personalized format to y-axis
    ax.yaxis.set_major_formatter(FuncFormatter(format_acq_cost))

    # Inserting chart into Canvas
    canvas_dispersion = FigureCanvasTkAgg(fig, master=root)
    canvas_dispersion.draw()
    # Configuring Canvas background
    canvas_dispersion.get_tk_widget().configure(background='#dbdbdb',)
    canvas_dispersion.get_tk_widget().pack(fill=ctk.BOTH, expand=True)

    # Annotation function to connect with mplcursors
    def set_annotations(sel):
        sel.annotation.set_text(
            'PN: ' + str(bup_scope['PN'][sel.target.index]) + "\n" +
            'Acq Cost: US$ ' + f"{round(bup_scope['Acq Cost'][sel.target.index]):,}" + "\n" +
            'Leadtime: ' + str(bup_scope['Leadtime'][sel.target.index]) + "\n" +
            'EIS Critical: ' + str(bup_scope['EIS Critical'][sel.target.index])
        )

    # Inserting Hover with mplcursors
    mpc.cursor(scatter, hover=True).connect('add', lambda sel: set_annotations(sel))

    # --------------- Turning it into an Image to be displayed  ---------------

    # Saving the matplotlib figure to a BytesIO object (memory), so as not to have to save to an image file
    tmp_img_dispersion_chart = BytesIO()
    fig.savefig(tmp_img_dispersion_chart, format='png', transparent=True)
    tmp_img_dispersion_chart.seek(0)

    # Loading the chart image into an Image object that will be returned by the function
    dispersion_image = Image.open(tmp_img_dispersion_chart)

    return dispersion_image


@function_timer
def generate_histogram(bup_scope: pd.DataFrame, root: ctk.CTkFrame):
    '''
    :param bup_scope: DataFrame with scope and additional information
    :param root: CTk Frame in which chart will be inserted
    :return: histogram_image: Chart CTkImage object in order to export
    '''

    # Image size
    width, height = 600, 220

    # Creating figure and axes to insert the chart
    fig, ax = plt.subplots(figsize=(width / 100, height / 100), layout='constrained')  # Layout property that handles "cutting" axes labels
    # Keeping background transparent
    fig.patch.set_facecolor("None")
    fig.patch.set_alpha(0)
    ax.set_facecolor('None')

    # Creating Histogram and saving the information in control variables
    n, bins, patches = ax.hist(bup_scope['Leadtime'], bins=20, edgecolor='k', linewidth=0.7, alpha=0.9)

    # Histogram settings
    ax.set_ylabel('Materials Count')
    ax.set_xlabel('Leadtime', loc='right')
    ax.set_title('Leadtime Histogram (in days)', fontsize=10)
    # Adjusting the y-axis limit (the largest bar was transcending the upper limit)
    ax.set_ylim(0, max(n) + 50)  # Adding a margin to accommodate the count at the top of the bar

    # Inserting the count into each bar
    for count, bar in zip(n, patches):
        x = bar.get_x() + bar.get_width() / 2
        y = bar.get_height()
        ax.text(x, y, f'{int(count)}', ha='center', va='bottom', fontdict={
            'family': 'open sans',
            'size': 9
        })

    # Annotation function to connect with mplcursors
    def set_annotations(sel):
        sel.annotation.set_text(
            'Parts Count: ' + str(round(sel.target[1])) + "\n" +
            'Lower Limit: ' + str(round(bins[sel.target.index])) + "\n" +
            'Upper Limit: ' + str(round(bins[sel.target.index + 1]))
        )
        pass

    # Inserting Hover with mplcursors
    mpc.cursor(patches, hover=True).connect('add', lambda sel: set_annotations(sel))

    # Inserting chart into Canvas
    canvas_histogram = FigureCanvasTkAgg(fig, master=root)
    canvas_histogram.draw()
    # Configuring Canvas background
    canvas_histogram.get_tk_widget().configure(background='#dbdbdb')
    canvas_histogram.get_tk_widget().pack(fill=ctk.BOTH, expand=True)

    # --- Saving the chart image in BytesIO() (memory) so it is not necessary to save as a file ---
    tmp_img_histogram_chart = BytesIO()
    fig.savefig(tmp_img_histogram_chart, format='png', transparent=True)
    tmp_img_histogram_chart.seek(0)

    # Keeping the image in an Image object
    histogram_chart = Image.open(tmp_img_histogram_chart)

    # Loading into a CTk Image object
    histogram_image = ctk.CTkImage(histogram_chart,
                                   dark_image=histogram_chart,
                                   size=(600, 220))

    return histogram_image


@function_timer
def create_scenario(scenario_window, var_scenarios_count, bup_scope, efficient_curve_window, hypothetical_curve_window) -> None:
    global scenarios_list

    scenario = {}

    # If the Scenarios list contains at least 1 already registered, the user is offered the option of using the
    # values of Contractual Conditions from the first scenario, changing only the Procurement Length parameters

    if scenarios_list:
        # Function to open the Dialog box offering for the user the possibility of reuse
        def open_confirm_dialog() -> None:
            confirm_window = ctk.CTkToplevel(scenario_window, fg_color='#ebebeb')
            confirm_window.title("Warning!")
            confirm_window.resizable(width=False, height=False)

            # Dialog Box Geometry
            cw_width = 400
            cw_height = 170
            conf_window_width = confirm_window.winfo_screenwidth()  # Width of the screen
            conf_window_height = confirm_window.winfo_screenheight()  # Height of the screen
            # Calculates the initial X and Y to position the screen
            cw_x = int((conf_window_width / 2) - (cw_width / 2))
            cw_y = int((conf_window_height / 2) - (cw_height / 2))
            confirm_window.geometry('%dx%d+%d+%d' % (cw_width, cw_height, cw_x, cw_y))

            # Setting (capturing) the windows focus
            confirm_window.grab_set()

            # Dialog Box Elements
            lbl_question = ctk.CTkLabel(confirm_window, text="There is already a registered Scenario. Do you want to "
                                                             "use previously Contractual Conditions information?",
                                        font=ctk.CTkFont('open sans', size=13, weight='bold'),
                                        width=300, wraplength=350
                                        )
            lbl_question.pack(pady=(20, 0))

            # Function that, when clicking the YES button, uses the values of the first registered
            # Scenario and disables the Entries.
            def use_previous_scenario_values() -> None:
                # Using control variables globally to store previously registered values
                global t0_previous_value, hyp_t0_previous_value, acft_delivery_start_previous_value, material_delivery_start_previous_value\
                    , material_delivery_end_previous_value

                t0_previous_value = ctk.StringVar(value=scenarios_list[0]['t0'].strftime("%d/%m/%Y"))
                entry_t0.configure(textvariable=t0_previous_value)
                entry_t0.configure(state="disabled")

                hyp_t0_previous_value = ctk.StringVar(value=scenarios_list[0]['hyp_t0_start'])
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

            def nullify_previous_scenario_variables() -> None:
                """ Function that makes empty the control variables that store the pre-registered scenarios, as
                the user chose not to use the already registered scenario
                """

                # Using variables globally
                global t0_previous_value, hyp_t0_previous_value, acft_delivery_start_previous_value, material_delivery_start_previous_value \
                    , material_delivery_end_previous_value

                t0_previous_value, hyp_t0_previous_value, acft_delivery_start_previous_value, material_delivery_start_previous_value \
                    , material_delivery_end_previous_value = None, None, None, None, None

            # YES button
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

            # NO button
            btn_no = ctk.CTkButton(confirm_window, text='No', command=lambda: (confirm_window.destroy(),
                                                                               nullify_previous_scenario_variables(),
                                                                               scenario_window.lift(),
                                                                               scenario_window.grab_set()),
                                   font=ctk.CTkFont('open sans', size=12, weight='bold'),
                                   bg_color="#ebebeb", fg_color="#ff0000", hover_color="#af0003",
                                   width=100, height=30, corner_radius=30, cursor="hand2"
                                   )
            btn_no.place(relx=0.7, rely=0.8, anchor=ctk.CENTER)

        # Opening the dialog box if you already have a Scenario registered
        open_confirm_dialog()

    # ----------------- CONTRACTUAL CONDITIONS -----------------

    # Internal Frame Contractual Conditions
    contractual_conditions_frame = ctk.CTkFrame(scenario_window, width=440, corner_radius=20)
    contractual_conditions_frame.pack(pady=(15, 0), expand=False)

    # Label title Contractual Conditions
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

    # Internal frame Procurement Length
    procurement_length_frame = ctk.CTkFrame(scenario_window, width=440, corner_radius=20)
    procurement_length_frame.pack(pady=(15, 0), expand=False)

    # Label title Procurement Length
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

    # ----------------- Label: (*) Required Information -----------------

    lbl_required_infornation = ctk.CTkLabel(scenario_window, text="(*) Required Information",
                                            font=ctk.CTkFont('open sans', size=10, weight='bold'),
                                            text_color="#ff0000")
    lbl_required_infornation.pack(padx=(20, 0), anchor="w")

    # ----------------- Interaction Buttons -----------------

    # Function to return the values entered by user in the Entry, handling Defaults. It also saves both chart
    # Images on global scope variables.
    def get_entry_values():
        global img_eff_chart, img_hyp_chart, df_scope_with_scenarios

        # --------- Contractual Conditions ---------

        # In these conditions, there is first a check whether the user has chosen to use information already registered

        # t0
        if t0_previous_value is None:

            # Checks if the value is a valid date before assigning to the variable
            date_t0_value = pd.to_datetime(entry_t0.get(), format='%d/%m/%Y', errors='coerce')
            if not pd.isna(date_t0_value):
                scenario['t0'] = date_t0_value
            else:
                messagebox.showerror("Error",
                                     "Invalid date. Please enter a valid date format for T0.")
                return
        else:
            scenario['t0'] = pd.to_datetime(t0_previous_value.get(), format='%d/%m/%Y', errors='coerce')

        # t0+X: Int (default: 3) which will be added to t0 to indicate a hypothetical starting date for purchasing materials
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

            # Checks if the value is a valid date before assigning to the variable
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
        # Assignment of Default values in try/except if the user does not fill in anything

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

        # Including the scenario in the global Dicts list and closing the screen
        scenarios_list.append(scenario)
        scenario_window.destroy()

        # Adding 1 to IntVar with the Scenarios count
        var_scenarios_count.set(var_scenarios_count.get() + 1)

        # ----------- Calling the chart generation function -----------

        # Calling the function to generate the Efficient Build-Up chart. The return of the function is the chart in a
        # figure (Image object), in addition to the DataFrames/Variables created in the function, as a return to be used
        # in the Hypothetical chart
        bup_eff_chart_whitebg, bup_eff_chart, df_scope_with_scenarios, scenario_dataframes = \
            generate_efficient_curve_buildup_chart(bup_scope, scenarios_list, efficient_curve_window)

        # Loading into a CTk Image object
        img_bup_efficient_chart = ctk.CTkImage(bup_eff_chart,
                                    dark_image=bup_eff_chart,
                                    size=(580, 370))

        # Calling the function to generate Hypothetical Build-Up chart.
        bup_hyp_chart_whitebg, bup_hyp_chart = generate_hypothetical_curve_buildup_chart(df_scope_with_scenarios, scenario_dataframes)

        # Loading into a CTk Image object
        img_bup_hypothetical_chart = ctk.CTkImage(bup_hyp_chart,
                                    dark_image=bup_hyp_chart,
                                    size=(580, 370))

        # Hypothetical Curve Build-Up Chart - inputting CTkImage in the Label and positioning it on the screen
        ctk.CTkLabel(hypothetical_curve_window, image=img_bup_hypothetical_chart,
                    text="").place(relx=0.5, rely=0.43, anchor=ctk.CENTER)

        # Saving both charts Image on global scope variables
        img_eff_chart, img_hyp_chart = bup_eff_chart_whitebg, bup_hyp_chart_whitebg

    # OK button
    btn_ok = ctk.CTkButton(scenario_window, text='OK', command=get_entry_values,
                           font=ctk.CTkFont('open sans', size=12, weight='bold'),
                           bg_color="#ebebeb", fg_color="#009898", hover_color="#006464",
                           width=100, height=30, corner_radius=30, cursor="hand2"
                           )
    btn_ok.place(relx=0.3, rely=0.92, anchor=ctk.CENTER)

    # Cancel button
    btn_cancel = ctk.CTkButton(scenario_window, text='Cancel', command=scenario_window.destroy,
                               font=ctk.CTkFont('open sans', size=12, weight='bold'),
                               bg_color="#ebebeb", fg_color="#ff0000", hover_color="#af0003",
                               width=100, height=30, corner_radius=30, cursor="hand2"
                               )
    btn_cancel.place(relx=0.7, rely=0.92, anchor=ctk.CENTER)


@function_timer
def generate_efficient_curve_buildup_chart(bup_scope: pd.DataFrame, scenarios, root: ctk.CTkFrame):
    '''
    :param bup_scope: DataFrame with Scope and Scenarios
    :param scenarios: List with all created Scenarios
    :param root: CTkFrame in which the Chart will be displayed
    :return: CTkImage objects in order to be exported/downloaded. Chart will be plotted in this function with FigureCanvasTkAgg
    '''
    # --------------- Data Processing ---------------

    # List to store combinations of Scenarios and Scope
    combinations = []

    global scenario_dataframes

    for _, row in bup_scope.iterrows():
        # Going through each element of the dictionaries list (each element is a Scenario in scenarios_list)
        for index, scenario in enumerate(scenarios):
            # Combining dataframe values (scope) with dictionary values (scenario)
            comb = {**row, 'Scenario': index, **scenario}
            combinations.append(comb)

    # Creating a new dataframe with the Scope and Scenario combinations together
    df_scope_with_scenarios = pd.DataFrame(combinations).sort_values(by='Scenario').reset_index(drop=True)
    df_scope_with_scenarios['avg_month_diff'] = ((df_scope_with_scenarios['material_delivery_end']
                                                 - df_scope_with_scenarios['material_delivery_start']) / 2).astype(int)

    # Procurement Length - NOTE: There will come a time when I will have to create the logic for Export License here
    df_scope_with_scenarios['PN Procurement Length'] = df_scope_with_scenarios[['Leadtime', 'pr_release_approval_vss',
                                                                                'po_commercial_condition',
                                                                                'po_conversion', 'export_license',
                                                                                'buffer', 'outbound_logistic']].sum(axis=1)

    # Generating the average date (of the materials delivery interval based on t0).
    df_scope_with_scenarios['avg_date_between_materials_deadline'] = df_scope_with_scenarios.apply(
        lambda linha: linha['t0'] + pd.DateOffset(months=linha['material_delivery_start']) +
        pd.DateOffset(months=linha['avg_month_diff']), axis=1)

    # Creating Date columns for the 2 that come as integers based on t0
    df_scope_with_scenarios['material_delivery_start_date'] = df_scope_with_scenarios.apply(
        lambda linha: linha['t0'] + pd.DateOffset(months=linha['material_delivery_start'])
        , axis=1)

    df_scope_with_scenarios['material_delivery_end_date'] = df_scope_with_scenarios.apply(
        lambda linha: linha['t0'] + pd.DateOffset(months=linha['material_delivery_end'])
        , axis=1)

    # Calculating the date in which the material should be purchased, considering Procurement Length and Delivery Date
    df_scope_with_scenarios['PN Order Date'] = df_scope_with_scenarios.apply(
        lambda linha: linha['avg_date_between_materials_deadline'] - pd.DateOffset(days=linha['PN Procurement Length'])
        , axis=1)

    # Creating the column with the hypothetical start date of material purchases
    df_scope_with_scenarios['PN Order Date Hypothetical'] = df_scope_with_scenarios.apply(
        lambda linha: linha['t0'] + pd.DateOffset(months=linha['hyp_t0_start'])
        , axis=1)
    # Creating the date column on which the material will be delivered, for the Hypothetical chart
    df_scope_with_scenarios['Delivery Date Hypothetical'] = df_scope_with_scenarios.apply(
        lambda linha: linha['PN Order Date Hypothetical'] + pd.DateOffset(days=linha['PN Procurement Length'])
        , axis=1)

    '''
    Getting the Maximum and Minimum Date among all possible dates, to delimit the chart's X axis. Efficient and Hypothetical curve have different min/max dates,
    so it is necessary to create two different timelines.
    '''
    date_columns_eff = ['t0', 'acft_delivery_start', 'material_delivery_start_date', 'material_delivery_end_date', 'PN Order Date']
    date_clumns_hyp = ['t0', 'acft_delivery_start', 'material_delivery_start_date', 'material_delivery_end_date', 'Delivery Date Hypothetical']

    '''
    The first min/max returns the min/max date per column and at the end we then have a list of minimums/maximums
    The second takes the minimum one from the created list. I add a month at each extreme so that the
    parameters vertical lines do not coincide of with chart axis limit line 
    '''
    min_date_eff = df_scope_with_scenarios[date_columns_eff].min().min() - pd.DateOffset(months=1)
    max_date_eff = df_scope_with_scenarios[date_columns_eff].max().max() + pd.DateOffset(months=1)
    min_date_hyp = df_scope_with_scenarios[date_clumns_hyp].min().min() - pd.DateOffset(months=1)
    max_date_hyp = df_scope_with_scenarios[date_clumns_hyp].max().max() + pd.DateOffset(months=1)

    date_range_eff = pd.date_range(start=min_date_eff, end=max_date_eff, freq='M')
    date_range_hyp = pd.date_range(start=min_date_hyp, end=max_date_hyp, freq='M')

    df_dates_eff = pd.DataFrame({'Date': date_range_eff.strftime('%m/%Y')})
    df_dates_hyp = pd.DataFrame({'Date': date_range_hyp.strftime('%m/%Y')})

    # ------- Table for Chart generation --------
    # -- Efficient Curve
    # Creating a table with the grouped count of items purchased per month, for each Scenario
    grouped_counts_eff = df_scope_with_scenarios.groupby([df_scope_with_scenarios['PN Order Date'].dt.to_period('M'), 'Scenario']).size().reset_index(
        name='Ordered Qty')
    grouped_counts_eff['PN Order Date'] = grouped_counts_eff['PN Order Date'].dt.strftime('%m/%Y')
    # -- Hypothetical Curve
    # Also creating the Delivered Qty fields (Hypothetical)
    grouped_counts_hyp = df_scope_with_scenarios.groupby([df_scope_with_scenarios['Delivery Date Hypothetical'].dt.to_period('M'),
                                                          'Scenario']).size().reset_index(name='Delivered Qty Hyp')
    grouped_counts_hyp['Delivery Date Hypothetical'] = grouped_counts_hyp['Delivery Date Hypothetical'].dt.strftime('%m/%Y')


    # Passing the Ordered Qty and Delivered Qty info grouped by month and by Scenario to the DF with the Dates Range. Each Hyp/Eff has a dates_range_df
    # Eff
    final_df_scenarios_eff = df_dates_eff.merge(grouped_counts_eff, left_on='Date', right_on='PN Order Date', how='left')
    final_df_scenarios_eff['Scenario'] = final_df_scenarios_eff['Scenario'].fillna(-1).astype(int)
    final_df_scenarios_eff['Ordered Qty'] = final_df_scenarios_eff['Ordered Qty'].fillna(0)
    # Hyp
    final_df_scenarios_hyp = df_dates_hyp.merge(grouped_counts_hyp, left_on='Date', right_on='Delivery Date Hypothetical', how='left')
    final_df_scenarios_hyp['Scenario'] = final_df_scenarios_hyp['Scenario'].fillna(-1).astype(int)
    final_df_scenarios_hyp['Delivered Qty Hyp'] = final_df_scenarios_hyp['Delivered Qty Hyp'].fillna(0)

    # Eff - For each scenario, calculating the Accumulated Quantity to plot
    for scenario in final_df_scenarios_eff['Scenario'].unique():
        # Filtering the df for the current Scenario
        scenario_df = final_df_scenarios_eff[final_df_scenarios_eff['Scenario'] == scenario]
        # Calculating the accumulated quantity
        final_df_scenarios_eff.loc[scenario_df.index, 'Accum. Ordered Qty (Eff)'] = scenario_df['Ordered Qty'].cumsum()

    # Hyp - For each scenario, calculating the Accumulated Quantity to plot
    for scenario in final_df_scenarios_hyp['Scenario'].unique():
        # Filtering the df for the current Scenario
        scenario_df = final_df_scenarios_hyp[final_df_scenarios_hyp['Scenario'] == scenario]
        # Calculating the accumulated quantity
        final_df_scenarios_hyp.loc[scenario_df.index, 'Accum. Delivered Qty (Hyp)'] = scenario_df['Delivered Qty Hyp'].cumsum()

    # Eff - Creating a DF for each Scenario
    for scenario in final_df_scenarios_eff['Scenario'].unique():
        if scenario != -1:  # Scenario -1 is only indicative of nullity, it is not a real scenario

            '''
            Creating the list associated to each Scenario in the dict. This list should have 2 elements.
            First one is Efficient DF and Second one is Hypothetical DF
            '''
            scenario_dataframes[f'Scenario_{int(scenario)}'] = []

            tmp_filtered_scenario_final_df = final_df_scenarios_eff[final_df_scenarios_eff['Scenario'] == scenario]
            scenario_df = df_dates_eff.merge(tmp_filtered_scenario_final_df[
                                             ['Date', 'Scenario', 'Ordered Qty','Accum. Ordered Qty (Eff)']
                                             ],
                                         left_on='Date', right_on='Date', how='left')
            # Filling in nulls
            scenario_df['Scenario'] = scenario_df['Scenario'].fillna(scenario).astype(int)
            scenario_df['Ordered Qty'] = scenario_df['Ordered Qty'].fillna(0)

            # Filling empty months for Accumulated Qty based on last record
            scenario_df['Accum. Ordered Qty (Eff)'].fillna(method='ffill', inplace=True)

            # Storing the DataFrame in the dictionary with the scenario name
            scenario_dataframes[f'Scenario_{int(scenario)}'].append(scenario_df)

    # Hyp - Creating a DF for each Scenario
    for scenario in final_df_scenarios_hyp['Scenario'].unique():
        if scenario != -1:  # Scenario -1 is only indicative of nullity, it is not a real scenario

            tmp_filtered_scenario_final_df = final_df_scenarios_hyp[final_df_scenarios_hyp['Scenario'] == scenario]
            scenario_df = df_dates_hyp.merge(tmp_filtered_scenario_final_df[
                                             ['Date', 'Scenario', 'Delivered Qty Hyp','Accum. Delivered Qty (Hyp)']
                                             ],
                                         left_on='Date', right_on='Date', how='left')
            # Filling in nulls
            scenario_df['Scenario'] = scenario_df['Scenario'].fillna(scenario).astype(int)
            scenario_df['Delivered Qty Hyp'] = scenario_df['Delivered Qty Hyp'].fillna(0)

            # Filling empty months for Accumulated Qty based on last record
            scenario_df['Accum. Delivered Qty (Hyp)'].fillna(method='ffill', inplace=True)

            # Storing the DataFrame in the dictionary with the scenario name
            scenario_dataframes[f'Scenario_{int(scenario)}'].append(scenario_df)

    # --------------- Chart Generation ---------------

    # Colors list, so that each Scenario has a specific color and facilitates differentiation
    colors_array = ['blue', 'orange', 'black', 'green', 'purple']

    # Image size
    width, height = 580, 370

    # Creating a figure and axes to insert the chart
    fig, ax = plt.subplots(figsize=(width / 100, height / 100))
    # Keeping background transparent
    fig.patch.set_facecolor("None")
    fig.patch.set_alpha(0)
    ax.set_facecolor('None')

    # Eff - Plotting the line for each Scenario in the dictionary
    for index, (scenario_name, scenario_df_list) in enumerate(scenario_dataframes.items()):
        axs = ax.plot(scenario_df_list[0]['Date'], scenario_df_list[0]['Accum. Ordered Qty (Eff)'], label=f'Scen. {index}', color=colors_array[index])
        # Configuring the axis
        plt.xticks(scenario_df_list[0].index[::3], scenario_df_list[0]['Date'][::3], rotation=45, ha='right')

        # Getting the t0 date for the current Scenario and converting it to MM/YYYY format
        t0_date = pd.to_datetime(df_scope_with_scenarios.loc[df_scope_with_scenarios['Scenario'] == index, 't0'].values[0])
        t0_date = t0_date.strftime('%m/%Y')
        # Adding a vertical line at t0
        ax.axvline(x=t0_date, linestyle='--', color=colors_array[index], label=f't0: Scen. {index}')

        # Getting the acft_delivery_start date for the current Scenario and converting it to MM/YYYY format
        acft_delivery_start_date = pd.to_datetime(df_scope_with_scenarios.loc[df_scope_with_scenarios['Scenario'] == index, 'acft_delivery_start'].values[0])
        acft_delivery_start_date = acft_delivery_start_date.strftime('%m/%Y')
        # Adding a vertical line in acft_delivery_start
        ax.axvline(x=acft_delivery_start_date, linestyle='dotted', color=colors_array[index], label=f'Acft Delivery Start: Scen. {index}')

        # Adding a material delivery range between the Start and End dates
        material_delivery_start_date = pd.to_datetime(df_scope_with_scenarios.loc[df_scope_with_scenarios['Scenario'] == index, 'material_delivery_start_date'].values[0])
        material_delivery_start_date = material_delivery_start_date.strftime('%m/%Y')
        material_delivery_end_date = pd.to_datetime(df_scope_with_scenarios.loc[df_scope_with_scenarios['Scenario'] == index, 'material_delivery_end_date'].values[0])
        material_delivery_end_date = material_delivery_end_date.strftime('%m/%Y')

        ax.axvspan(material_delivery_start_date, material_delivery_end_date, alpha=0.5, color=colors_array[index])

        # Adding a note at the point where Build-Up planning should start (date when first order is released)
        filter_dates_with_order = scenario_df_list[0]['Ordered Qty'] != 0
        dates_with_order = scenario_df_list[0][filter_dates_with_order]
        index_first_order = dates_with_order['Ordered Qty'].idxmin()
        x_first_order = scenario_df_list[0].loc[index_first_order, 'Date']
        y_first_order = scenario_df_list[0].loc[index_first_order, 'Accum. Ordered Qty (Eff)']
        ax.scatter(x_first_order, y_first_order, color=colors_array[index], marker='o', label=f'Planning Start: {x_first_order}')

        # Adding a caretdown (not labeling) in BUP finish date (avg between End and Start material delivery date)
        # x_bup_finished = scenario_df.loc[0, 'avg_date_between_materials_deadline']
        y_bup_finished = scenario_df_list[0]['Accum. Ordered Qty (Eff)'].max()
        # Getting the avg_date_between_materials_deadline date for the current Scenario and converting it to MM/YYYY
        x_bup_finished = pd.to_datetime(df_scope_with_scenarios.loc[df_scope_with_scenarios['Scenario'] == index,
                                                                    'avg_date_between_materials_deadline']
                                        .values[0]).strftime('%m/%Y')
        ax.scatter(x_bup_finished, y_bup_finished, color=colors_array[index], marker=7, label=None)

    # Chart settings
    ax.set_ylabel('Materials Ordered Qty (Accumulated)')
    ax.set_title('Efficient Curve: Build-Up Forecast')
    ax.grid(True)
    ax.tick_params(axis='both', labelsize=9)  # Adjusting labels size

    # Adjusting axis spacing to avoid cutting off labels
    plt.subplots_adjust(left=0.15, right=0.9, bottom=0.2, top=0.9)

    # Legend
    ax.legend(loc='upper left', fontsize=7, framealpha=0.8)

    # Annotation function to connect with mplcursors
    def set_annotations(sel):
        sel.annotation.set_text(
            'Ordered Qty: ' + str(round(sel.target[1])) + "\n" +
            'Date ' + str(scenario_df_list[0]['Date'][round(sel.target.index)]) + "\n"
        )
    # Inserting Hover with mplcursors
    mpc.cursor(axs, hover=True).connect('add', lambda sel: set_annotations(sel))

    # Inserting chart into Canvas
    canvas_eff = FigureCanvasTkAgg(fig, master=root)
    canvas_eff.draw()
    # Configuring Canvas background
    canvas_eff.get_tk_widget().configure(background='#cfcfcf')
    # canvas_eff.get_tk_widget().pack(fill=ctk.BOTH, expand=True)
    canvas_eff.get_tk_widget().place(relx=0.5, rely=0.43, anchor=ctk.CENTER)

    # --------------- Turning it into an Image to be displayed ---------------

    # Saving the matplotlib figure to a BytesIO object (memory), so it is not necessary to save it in an image file
    tmp_img_eff_chart = BytesIO()
    fig.savefig(tmp_img_eff_chart, format='png', transparent=True)
    tmp_img_eff_chart.seek(0)

    # Loading the chart image into an Image object that will be returned by the function
    bup_eff_chart = Image.open(tmp_img_eff_chart)

    # It is necessary to save a chart Image with white background. Transparent is to plot. White to save as a file.
    tmp_img_eff_chart_whitebg = BytesIO()
    fig.savefig(tmp_img_eff_chart_whitebg, format='png', transparent=False)
    tmp_img_eff_chart_whitebg.seek(0)

    # Loading the chart image into an Image object that will be returned by the function
    bup_eff_chart_whitebg = Image.open(tmp_img_eff_chart_whitebg)

    return bup_eff_chart_whitebg, bup_eff_chart, df_scope_with_scenarios, scenario_dataframes


@function_timer
def generate_hypothetical_curve_buildup_chart(df_scope_with_scenarios, scenario_dataframes):
    """
    Function that creates the Hypothetycal Curve BuildUp Chart.
    param df_scope_with_scenarios: Created DataFrame on Efficient Curve Build-Up construction. Combinations Scope/Scenarios.
    param scenario_dataframes: Dictionary with all scenarios dataframes. Each scenario has a list with 2 DF elements. Efficient and Hypothetical, respectively.
    return: Returns an Image object
    """

    # --------------- Chart Generation ---------------

    # List of colors, so that each Scenario has a specific color and facilitates differentiation
    colors_array = ['blue', 'orange', 'black', 'green', 'purple']

    # Image Size
    width, height = 580, 370

    # Creating a figure and axes to insert the chart
    figura, eixos = plt.subplots(figsize=(width / 100, height / 100))

    # Plotting the line for each Scenario in the dictionary
    for index, (scenario_name, scenario_df_list) in enumerate(scenario_dataframes.items()):
        eixos.plot(scenario_df_list[1]['Date'], scenario_df_list[1]['Accum. Delivered Qty (Hyp)'], label=f'Scen. {index}',
                   color=colors_array[index])
        # Configuring the axis
        plt.xticks(scenario_df_list[1].index[::3], scenario_df_list[1]['Date'][::3], rotation=45, ha='right')

        # Getting the t0 date for the current Scenario and converting it to MM/YYYY format
        t0_date = pd.to_datetime(
            df_scope_with_scenarios.loc[df_scope_with_scenarios['Scenario'] == index, 't0'].values[0])
        t0_date = t0_date.strftime('%m/%Y')
        # Adding a vertical line at t0
        eixos.axvline(x=t0_date, linestyle='--', color=colors_array[index], label=f't0: Scen. {index}')

        # Getting the acft_delivery_start date for the current Scenario and converting it to MM/YYYY format
        acft_delivery_start_date = pd.to_datetime(
            df_scope_with_scenarios.loc[df_scope_with_scenarios['Scenario'] == index, 'acft_delivery_start'].values[0])
        acft_delivery_start_date = acft_delivery_start_date.strftime('%m/%Y')
        # Adding a vertical line in acft_delivery_start
        eixos.axvline(x=acft_delivery_start_date, linestyle='dotted', color=colors_array[index],
                      label=f'Acft Delivery Start: Scen. {index}')

        # Adding a material delivery range between the Start and End dates
        material_delivery_start_date = pd.to_datetime(df_scope_with_scenarios.loc[
                                                          df_scope_with_scenarios['Scenario'] == index,
                                                          'material_delivery_start_date'].values[0])
        material_delivery_start_date = material_delivery_start_date.strftime('%m/%Y')
        material_delivery_end_date = pd.to_datetime(df_scope_with_scenarios.loc[
                                                        df_scope_with_scenarios['Scenario'] == index,
                                                        'material_delivery_end_date'].values[0])
        material_delivery_end_date = material_delivery_end_date.strftime('%m/%Y')

        eixos.axvspan(material_delivery_start_date, material_delivery_end_date, alpha=0.5, color=colors_array[index])

        # Adding a note at the point where the Build-Up is completed (all items delivered)
        index_max_acc_qty = scenario_df_list[1]['Accum. Delivered Qty (Hyp)'].idxmax()
        x_max = scenario_df_list[1].loc[index_max_acc_qty, 'Date']
        y_max = scenario_df_list[1].loc[index_max_acc_qty, 'Accum. Delivered Qty (Hyp)']
        plt.scatter(x_max, y_max, color=colors_array[index], marker='o', label=f'BUP Conclusion: {x_max}')

    # Chart Settings
    eixos.set_ylabel('Materials Delivered Qty (Accumulated)')
    eixos.tick_params(axis='both', labelsize=9)  # Adjusting labels size
    eixos.set_title('Hypothetical Curve: Build-Up Forecast')
    eixos.grid(True)

    # Adjusting axis spacing to avoid cutting off labels
    plt.subplots_adjust(left=0.15, right=0.9, bottom=0.2, top=0.9)

    # Legend
    eixos.legend(loc='upper left', fontsize=7, framealpha=0.8)

    # --------------- Turning it into an Image to be displayed ---------------

    # Saving the matplotlib figure to a BytesIO object (memory), so as not to have to save an image file
    tmp_img_hyp_chart = BytesIO()
    figura.savefig(tmp_img_hyp_chart, format='png', transparent=True)
    tmp_img_hyp_chart.seek(0)

    # Loading the chart image into an Image object that will be returned by the function
    bup_hyp_chart = Image.open(tmp_img_hyp_chart)

    # It is necessary to save a chart Image with white background. Transparent is to plot. White to save as a file.
    tmp_img_hyp_chart_whitebg = BytesIO()
    figura.savefig(tmp_img_hyp_chart_whitebg, format='png', transparent=False)
    tmp_img_hyp_chart_whitebg.seek(0)

    # Loading the chart image into Image objects that will be returned by the function
    bup_hyp_chart_whitebg = Image.open(tmp_img_hyp_chart_whitebg)

    return bup_hyp_chart_whitebg, bup_hyp_chart


@function_timer
def save_chart_image(chart: Image, output_path: str, filename: str) -> None:
    # Function that receveis as argument the chart Image object, as well as output path and filename, and saves it.

    try:
        chart.save(output_path + '\\' + filename)
        messagebox.showinfo(title="Success!", message=str("Image was exported to: " + output_path + '\\' + filename))
    except Exception as ex:
        messagebox.showinfo(title="Error!", message=str(ex) + "\n\nPlease make sure that the Image file is "
                                                              "closed and you have access to the local Downloads folder.")
