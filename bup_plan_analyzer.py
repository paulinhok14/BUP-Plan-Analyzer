# Data Wrangling
import pandas as pd, numpy as np
from dateutil.relativedelta import relativedelta
from datetime import datetime

# Data Viz
import matplotlib.pyplot as plt
from matplotlib.ticker import FuncFormatter
from matplotlib.dates import DateFormatter
import matplotlib.lines as mlines
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import mplcursors as mpc

# System
import warnings, time, logging

# Extras
from PIL import Image
from io import BytesIO
import customtkinter as ctk
import tkinter as tk
from tkinter import messagebox


warnings.filterwarnings("ignore")

'''
** VARIABLE REFERENCE DOC **:

scenario_dataframes (dict) has a list for each scenario, ex:
scenario_dataframes['Scenario_0'] = []
scenario_dataframes['Scenario_1'] = []

In which, each element of this list, by order, is a specific pandas dataframe:

for (scenario_name, scenario_df_list) in scenario_dataframes.items():
1- scenario_df_list[0]: Efficient Chart DataFrame (Parts)
2- scenario_df_list[1]: Hypothetical Chart DataFrame (Parts)
3- scenario_df_list[2]: Efficient Chart DataFrame (Acq Cost)
4- scenario_df_list[3]: Hypothetical Chart DataFrame (Acq Cost)
5- scenario_df_list[4]: (Cost Avoidance) - Efficient and Hypothetical Accum. Acq Cost Chart DataFrame. Both plots should have same reference Date (Order/Delivery)
'''

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

# Global variables to store FigureCanvasTkAgg objects to be toggled in SwitchButton. Changing Build-Up curves from Parts/AcqCost and also Cost Avoidance
canvas_eff, canvas_hyp, canvas_list_acqcost_eff, canvas_list_acqcost_hyp, canvas_list_cost_avoidance = None, None, [], [], []

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
def read_scope_file(file_full_path: str, load_mode: str) -> pd.DataFrame:
    # Function that reads scope file and complementary info

    # Defines local or network path for reading Ecode Data and MARCSA (SAP, Leadtimes) - Complementary Info source
    if load_mode == 'local':
        ecode_data_path = r'DB_Ecode-Data.txt'
        marcsa_path = r'marcsa.txt'
    elif load_mode == 'network':
        ecode_data_path = r'\\egmap20038-new\Databases\DB_Ecode-Data.txt'
        marcsa_path = r'\\sjkfs05\vss\GMT\40. Stock Efficiency\J - Operational Efficiency\006 - Srcfiles\003 - SAP\marcsa.txt'
    else:
        pass
        # TO DO: RETURN ANY KIND OF BREAK EXPLAINING LOAD_MODE NOT IDENTIFIED. ALLOWED MODES: 'LOCAL' AND 'NETWORK'

    # Resetting the list of scenarios every time the e-mail is read
    global scenarios_list
    scenarios_list = []

    # Columns to read from the Scope file (essential)
    colunas = ['PN', 'ECODE', 'QTY', 'EIS', 'SPC']

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
    leadtimes = pd.read_csv(marcsa_path, usecols=sap_source_columns, encoding='latin', sep='|', low_memory=False)
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
    width, height = 600, 235
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
    canvas_dispersion.get_tk_widget().pack(fill=ctk.BOTH, expand=True, pady=(0,10))

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
    width, height = 600, 235

    # Creating figure and axes to insert the chart
    fig, ax = plt.subplots(figsize=(width / 100, height / 100), layout='constrained')  # Layout property that handles "cutting" axes labels
    # Keeping background transparent
    fig.patch.set_facecolor("None")
    fig.patch.set_alpha(0)
    ax.set_facecolor('None')

    # Inserting Average and Standard Deviation of Leadtimes - Before histogram plottation
    avg_leadtimes = bup_scope['Leadtime'].mean()
    sd_leadtimes = bup_scope['Leadtime'].std()
    ax.axvline(x=avg_leadtimes, linestyle='--', color='black', label=f'Average: {round(avg_leadtimes)}')
    ax.axvspan(avg_leadtimes - sd_leadtimes, avg_leadtimes + sd_leadtimes, alpha=0.4, color='#fccf03',
               label=f'Std: {round(sd_leadtimes)}', hatch='/', edgecolor='black')

    # Creating Histogram and saving the information in control variables
    n, bins, patches = ax.hist(bup_scope['Leadtime'], bins=20, edgecolor='k', color='#1fa9a4', linewidth=0.7, alpha=0.9)

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

    # Legend
    ax.legend(fontsize=9, framealpha=0.6)

    # Annotation function to connect with mplcursors
    def set_annotations(sel):
        sel.annotation.set_text(
            'Parts Count: ' + str(round(sel.target[1])) + "\n" +
            'Lower Limit: ' + str(round(bins[sel.target.index])) + "\n" +
            'Upper Limit: ' + str(round(bins[sel.target.index + 1]))
        )

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
def create_scenario(scenario_window: ctk.CTkFrame, var_scenarios_count: ctk.IntVar, bup_scope: pd.DataFrame,
                    efficient_curve_window: ctk.CTkFrame, hypothetical_curve_window: ctk.CTkFrame, cost_avoidance_window: ctk.CTkFrame, 
                    batches_curve_window: ctk.CTkFrame, bup_cost: float):
    '''
    This is the function that handles Scenarios creating. Here will be created the Scenarios creation window, and also will be
    the function that calls all other functions that executes subsequently after creating a Scenario. That is:
    [generate_efficient_curve_buildup_chart(), generate_hypothetical_curve_buildup_chart(), generate_acqcost_curve(), generate_cost_avoidance_screen()]
    :param scenario_window:
    :param var_scenarios_count:
    :param bup_scope:
    :param efficient_curve_window:
    :param hypothetical_curve_window:
    :return:
    '''
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

    # ----------------- Batch Config -----------------

    # Internal frame Procurement Length
    batch_frame = ctk.CTkFrame(scenario_window, width=440, corner_radius=20)
    batch_frame.pack(pady=(15, 0), expand=False)

     # Label title Batch
    lbl_batch = ctk.CTkLabel(batch_frame, text="Batch Settings",
                                          font=ctk.CTkFont('open sans', size=14, weight='bold')
                                          )
    lbl_batch.grid(row=0, columnspan=2, sticky="n", pady=(10, 10))

    # --- Batches Settings ---

    # Batches Qty
    lbl_qty_batches = ctk.CTkLabel(batch_frame, text="Batches Quantity",
                                               font=ctk.CTkFont('open sans', size=11, weight='bold')
                                               )
    lbl_qty_batches.grid(row=1, column=0, sticky="w", padx=12)

    entry_batches = ctk.CTkEntry(batch_frame, width=150)
    entry_batches.configure(placeholder_text='Nº of Batches (ex: 4)')
    entry_batches.configure(state='disabled')
    entry_batches.grid(row=2, column=0, padx=10, sticky="w", pady=(0, 20))

    # Batches Date
    lbl_qty_batches = ctk.CTkLabel(batch_frame, text="Batches Date (comma-separated values)",
                                               font=ctk.CTkFont('open sans', size=11, weight='bold')
                                               )
    lbl_qty_batches.grid(row=1, column=1, sticky="e", padx=12)

    entry_batches_date = ctk.CTkEntry(batch_frame, width=250)
    entry_batches_date.configure(placeholder_text='DD/MM/YYYY , DD/MM/YYYY , etc')
    entry_batches_date.configure(state='disabled')
    entry_batches_date.grid(row=2, column=1, padx=10, sticky="e", pady=(0, 20))

    # Batch Option Switch

    def toggle_batch_mode_selection() -> None:
        '''
        This function handles the Enabled option for Batch information Entry
        '''
        # Get current State
        current_batch_switch_status = entry_batches.cget(attribute_name='state')
        
        # Conditional toggling
        if current_batch_switch_status == 'disabled':
            swt_batches_opt.configure(button_color='#004b00', progress_color='green')
            entry_batches.configure(state='normal', placeholder_text='Nº of Batches (ex: 4)')
            entry_batches_date.configure(state='normal', placeholder_text='DD/MM/YYYY , DD/MM/YYYY , etc')
            
        else:      
            swt_batches_opt.configure(button_color='#7c1f27', fg_color='red')    
            entry_batches.configure(state='disabled', placeholder_text=' ')
            entry_batches_date.configure(state='disabled', placeholder_text=' ')
            
        

    swt_batches_opt = ctk.CTkSwitch(batch_frame,
                                    text="",
                                    width=45, height=12, button_color='#7c1f27', fg_color='red', progress_color='green',
                                    command= toggle_batch_mode_selection
                                    )
    swt_batches_opt.grid(row=0, column=1, sticky='e')


    # ----------------- Label: (*) Required Information -----------------

    lbl_required_infornation = ctk.CTkLabel(scenario_window, text="(*) Required Information",
                                            font=ctk.CTkFont('open sans', size=10, weight='bold'),
                                            text_color="#ff0000")
    lbl_required_infornation.pack(padx=(20, 0), anchor="w")

    # ----------------- Interaction Buttons -----------------

    # Function to return the values entered by user in the Entry, handling Defaults. It also saves both chart
    # Images on global scope variables.
    def get_entry_values():
        '''
        As this is an awaiting function, it has to assign the value directly to the global scope variables.
        In another way, if tried to return directly from 'create_scenario()' (parent function), it would raise a 'non exists' error
        '''
        global img_eff_chart, img_hyp_chart, df_scope_with_scenarios, canvas_eff, canvas_hyp, canvas_acqcost_eff, canvas_acqcost_hyp

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
        
        # ------------- Batch Settings -------------

        # --- Batch Qty ---
        try:
            if entry_batches.get().strip() != "":
                scenario['batches_qty'] = int(entry_batches.get())
            else:
                scenario['batches_qty'] = None
        except ValueError:
            messagebox.showerror("Error",
                                 "Invalid character. Please enter a valid number for Batches Quantity.")
            return
        
        # --- Batch Dates ---
        try:
            if entry_batches_date.get().strip() != "":
                scenario['batches_dates'] = str(entry_batches_date.get())
            else:
                scenario['batches_dates'] = None
        except ValueError:
            messagebox.showerror("Error",
                                 "Invalid text. Please enter Batch Dates in specified format.")
            return

        # Summing up full Procurement Length values
        scenario['full_procurement_length'] = (scenario['pr_release_approval_vss'] + scenario['po_commercial_condition'] + scenario['po_conversion'] + 
                                               scenario['export_license'] + scenario['buffer'] + scenario['outbound_logistic'])
        
        # Including the scenario in the global Dicts list and closing the screen
        scenarios_list.append(scenario)
        scenario_window.destroy()

        # ----------- Calling the chart generation functions -----------

        # Calling the function to generate the Efficient Build-Up chart. The return of the function is the chart in a
        # figure (Image object), in addition to the DataFrames/Variables created in the function, as a return to be used
        # in the Hypothetical chart
        canvas_eff, bup_eff_chart_whitebg, df_scope_with_scenarios, scenario_dataframes, df_dates_eff, df_dates_hyp = generate_efficient_curve_buildup_chart(bup_scope, scenarios_list, 
                                                                                                                                                             efficient_curve_window,
                                                                                                                                                             hypothetical_curve_window)

        # Calling the function to generate Hypothetical Build-Up chart.
        bup_hyp_chart_whitebg, canvas_hyp = generate_hypothetical_curve_buildup_chart(df_scope_with_scenarios, scenario_dataframes, hypothetical_curve_window)

        # Saving both charts Image on global scope variables
        img_eff_chart, img_hyp_chart = bup_eff_chart_whitebg, bup_hyp_chart_whitebg

        # Calling function to generate Cost Avoidance Chart
        generate_cost_avoidance_screen(cost_avoidance_window, scenario_dataframes, scenarios_list, df_scope_with_scenarios, df_dates_eff, df_dates_hyp, bup_cost)

        # Calling function to generate Batches Build-Up chart
        generate_batches_curve(batches_curve_window, scenarios_list, df_scope_with_scenarios)

        # Adding 1 to IntVar with the Scenarios count
        var_scenarios_count.set(var_scenarios_count.get() + 1)


    # OK button
    btn_ok = ctk.CTkButton(scenario_window, text='OK', command=get_entry_values,
                           font=ctk.CTkFont('open sans', size=12, weight='bold'),
                           bg_color="#ebebeb", fg_color="#009898", hover_color="#006464",
                           width=100, height=30, corner_radius=30, cursor="hand2"
                           )
    btn_ok.place(relx=0.3, rely=0.95, anchor=ctk.CENTER)

    # Cancel button
    btn_cancel = ctk.CTkButton(scenario_window, text='Cancel', command=scenario_window.destroy,
                               font=ctk.CTkFont('open sans', size=12, weight='bold'),
                               bg_color="#ebebeb", fg_color="#ff0000", hover_color="#af0003",
                               width=100, height=30, corner_radius=30, cursor="hand2"
                               )
    btn_cancel.place(relx=0.7, rely=0.95, anchor=ctk.CENTER)



@function_timer
def generate_efficient_curve_buildup_chart(bup_scope: pd.DataFrame, scenarios: list, efficient_curve_window: ctk.CTkFrame, hypothetical_curve_window: ctk.CTkFrame):
    '''
    :param bup_scope: DataFrame with Scope and Scenarios
    :param scenarios: List with all created Scenarios
    :param root: CTkFrame in which the Chart will be displayed
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
            Creating the list associated to each Scenario in the dict. This list should have a DataFrame as each element.
            First one is Efficient DF and Second one is Hypothetical DF (parts), Third and Fourth are Eff and Hyp for Acq Cost, subsequently
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
    width, height = 680, 435

    # Creating a figure and axes to insert the chart
    fig, ax = plt.subplots(figsize=(width / 100, height / 100),  layout='constrained')
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
    ax.set_title('Efficient Curve: Build-Up Forecast', color='#ad7102', fontweight='bold')
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
            'Date ' + str(scenario_df_list[0]['Date'][round(sel.target.index)])
        )
    # Inserting Hover with mplcursors
    mpc.cursor(axs, hover=True).connect('add', lambda sel: set_annotations(sel))

    # Inserting chart into Canvas
    canvas_eff = FigureCanvasTkAgg(fig, master=efficient_curve_window)
    canvas_eff.draw()
    # Configuring Canvas background
    canvas_eff.get_tk_widget().configure(background='#cfcfcf')
    canvas_eff.get_tk_widget().place(relx=0.5, rely=0.46, anchor=ctk.CENTER)

    # --------------- Turning it into an Image to be displayed ---------------

    # It is necessary to save a chart Image with white background. Transparent is to plot. White to save as a file.
    tmp_img_eff_chart_whitebg = BytesIO()
    fig.savefig(tmp_img_eff_chart_whitebg, format='png', transparent=False)
    tmp_img_eff_chart_whitebg.seek(0)

    # Loading the chart image into an Image object that will be returned by the function
    bup_eff_chart_whitebg = Image.open(tmp_img_eff_chart_whitebg)

    # ----------- At last, calling function to Generate Charts with Acq Cost ----------- #
    generate_acqcost_curve(df_scope_with_scenarios, df_dates_eff, df_dates_hyp, scenario_dataframes, efficient_curve_window, hypothetical_curve_window)

    return canvas_eff, bup_eff_chart_whitebg, df_scope_with_scenarios, scenario_dataframes, df_dates_eff, df_dates_hyp


@function_timer
def generate_hypothetical_curve_buildup_chart(df_scope_with_scenarios: pd.DataFrame, scenario_dataframes: dict, root: ctk.CTkFrame):
    """
    Function that creates the Hypothetycal Curve BuildUp Chart.
    param df_scope_with_scenarios: Created DataFrame on Efficient Curve Build-Up construction. Combinations Scope/Scenarios.
    param scenario_dataframes: Dictionary with all scenarios dataframes. Each scenario has a list with 2 DF elements. Efficient and Hypothetical, respectively.
    return: Returns an Image object and also the Chart Canvas object (FigureCanvasTkAgg): canvas_hyp
    """

    # --------------- Chart Generation ---------------

    # List of colors, so that each Scenario has a specific color and facilitates differentiation
    colors_array = ['blue', 'orange', 'black', 'green', 'purple']

    # Image Size
    width, height = 680, 435

    # Creating a figure and axes to insert the chart
    fig, ax = plt.subplots(figsize=(width / 100, height / 100), layout='constrained')
    # Keeping background transparent
    fig.patch.set_facecolor("None")
    fig.patch.set_alpha(0)
    ax.set_facecolor('None')

    # Plotting the line for each Scenario in the dictionary
    for index, (scenario_name, scenario_df_list) in enumerate(scenario_dataframes.items()):
        axs = ax.plot(scenario_df_list[1]['Date'], scenario_df_list[1]['Accum. Delivered Qty (Hyp)'], label=f'Scen. {index}',
                   color=colors_array[index])
        # Configuring the axis
        plt.xticks(scenario_df_list[1].index[::3], scenario_df_list[1]['Date'][::3], rotation=45, ha='right')

        # Getting the t0 date for the current Scenario and converting it to MM/YYYY format
        t0_date = pd.to_datetime(
            df_scope_with_scenarios.loc[df_scope_with_scenarios['Scenario'] == index, 't0'].values[0])
        t0_date = t0_date.strftime('%m/%Y')
        # Adding a vertical line at t0
        ax.axvline(x=t0_date, linestyle='--', color=colors_array[index], label=f't0: Scen. {index}')

        # Getting the acft_delivery_start date for the current Scenario and converting it to MM/YYYY format
        acft_delivery_start_date = pd.to_datetime(
            df_scope_with_scenarios.loc[df_scope_with_scenarios['Scenario'] == index, 'acft_delivery_start'].values[0])
        acft_delivery_start_date = acft_delivery_start_date.strftime('%m/%Y')
        # Adding a vertical line in acft_delivery_start
        ax.axvline(x=acft_delivery_start_date, linestyle='dotted', color=colors_array[index],
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

        ax.axvspan(material_delivery_start_date, material_delivery_end_date, alpha=0.5, color=colors_array[index])

        # Adding a note at the point where the Build-Up is completed (all items delivered)
        index_max_acc_qty = scenario_df_list[1]['Accum. Delivered Qty (Hyp)'].idxmax()
        x_max = scenario_df_list[1].loc[index_max_acc_qty, 'Date']
        y_max = scenario_df_list[1].loc[index_max_acc_qty, 'Accum. Delivered Qty (Hyp)']
        plt.scatter(x_max, y_max, color=colors_array[index], marker='o', label=f'BUP Conclusion: {x_max}')

    # Chart Settings
    ax.set_ylabel('Materials Delivered Qty (Accumulated)')
    ax.tick_params(axis='both', labelsize=9)  # Adjusting labels size
    ax.set_title('Hypothetical Curve: Build-Up Forecast', color='#ad7102', fontweight='bold')
    ax.grid(True)

    # Adjusting axis spacing to avoid cutting off labels
    plt.subplots_adjust(left=0.15, right=0.9, bottom=0.2, top=0.9)

    # Legend
    ax.legend(loc='upper left', fontsize=7, framealpha=0.8)

    # Annotation function to connect with mplcursors
    def set_annotations(sel):
        sel.annotation.set_text(
            'Delivered Qty: ' + str(round(sel.target[1])) + "\n" +
            'Date ' + str(scenario_df_list[0]['Date'][round(sel.target.index)])
        )

    # Inserting Hover with mplcursors
    mpc.cursor(axs, hover=True).connect('add', lambda sel: set_annotations(sel))

    # Inserting chart into Canvas
    canvas_hyp = FigureCanvasTkAgg(fig, master=root)
    canvas_hyp.draw()
    # Configuring Canvas background
    canvas_hyp.get_tk_widget().configure(background='#cfcfcf')
    canvas_hyp.get_tk_widget().place(relx=0.5, rely=0.46, anchor=ctk.CENTER)

    # --------------- Turning it into an Image to be displayed ---------------

    # It is necessary to save a chart Image with white background. Transparent is to plot. White to save as a file.
    tmp_img_hyp_chart_whitebg = BytesIO()
    fig.savefig(tmp_img_hyp_chart_whitebg, format='png', transparent=False)
    tmp_img_hyp_chart_whitebg.seek(0)

    # Loading the chart image into Image objects that will be returned by the function
    bup_hyp_chart_whitebg = Image.open(tmp_img_hyp_chart_whitebg)

    return bup_hyp_chart_whitebg, canvas_hyp


@function_timer
def generate_acqcost_curve(df_scope_with_scenarios: pd.DataFrame, df_dates_eff: pd.DataFrame, df_dates_hyp: pd.DataFrame, scenario_dataframes: dict,
                           efficient_curve_window: ctk.CTkFrame, hypothetical_curve_window: ctk.CTkFrame):
    '''
    This function generates Acq Cost charts for both Efficient and Hypothetical curve.
    :param df_scope_with_scenarios: DataFrame with Scope and All Scenarios joined information
    :param df_dates_eff: DataFrame with a 'date' Series, with Min and Max range date for Efficient Chart
    :param df_dates_hyp: DataFrame with a 'date' Series, with Min and Max range date for Hypothetical Chart
    :return: Object FigureCanvasTkAgg, in order to be plotted as soon as the Switches to Acq Cost are toggled. It will be managed by another function.
    '''

    # Global variables to store Acq Cost charts (each Scenario produces a particular Chart)
    global canvas_list_acqcost_eff, canvas_list_acqcost_hyp

    # Hard copy so not to change main df_scope_with_scenarios
    df_acqcost_chart_info = df_scope_with_scenarios.copy()

    # Creating Material Total Cost column (Qty * Acq Cost)
    df_acqcost_chart_info['Total Acq Cost'] = df_acqcost_chart_info['Qty'] * df_acqcost_chart_info['Acq Cost']
    # Creating Month/Year columns for Efficient and Hypothetical charts
    df_acqcost_chart_info['Order Date (Eff)'] = df_acqcost_chart_info['PN Order Date'].dt.to_period('M').dt.strftime('%m/%Y')
    df_acqcost_chart_info['Delivery Date (Hyp)'] = df_acqcost_chart_info['Delivery Date Hypothetical'].dt.strftime('%m/%Y')

    # Creating DFs with Grouped Sum for each Chart (Eff/Hyp)
    grouped_sum_acqcost_eff = df_acqcost_chart_info[['Scenario', 'Order Date (Eff)', 'Total Acq Cost']].groupby(['Scenario', 'Order Date (Eff)']).sum('Total Acq Cost').reset_index()
    grouped_sum_acqcost_hyp = df_acqcost_chart_info[['Scenario', 'Delivery Date (Hyp)', 'Total Acq Cost']].groupby(['Scenario', 'Delivery Date (Hyp)']).sum('Total Acq Cost').reset_index()

    # Merging with Date Range DFs for each Chart (Eff/Hyp)
    grouped_sum_acqcost_eff = df_dates_eff.merge(grouped_sum_acqcost_eff, how='left', left_on='Date', right_on='Order Date (Eff)')
    grouped_sum_acqcost_hyp = df_dates_hyp.merge(grouped_sum_acqcost_hyp, how='left', left_on='Date', right_on='Delivery Date (Hyp)')

    # Making '-1' the 'Scenario' null values. This is a solution to not allow multiple scenarios dfs to disturb a fillna() with specific rules.
    grouped_sum_acqcost_eff['Scenario'] = grouped_sum_acqcost_eff['Scenario'].fillna(-1).astype(int)
    grouped_sum_acqcost_hyp['Scenario'] = grouped_sum_acqcost_hyp['Scenario'].fillna(-1).astype(int)


    # Efficient - For each scenario, calculating the Accumulated Acq Cost to plot
    for scenario in grouped_sum_acqcost_eff['Scenario'].unique():
        # Filtering the df for the current Scenario
        scenario_df = grouped_sum_acqcost_eff[grouped_sum_acqcost_eff['Scenario'] == scenario]
        # Calculating the accumulated $ Volume
        grouped_sum_acqcost_eff.loc[scenario_df.index, 'Accum. Acq Cost'] = scenario_df['Total Acq Cost'].cumsum()

    # Hypothetical - For each scenario, calculating the Accumulated Quantity to plot
    for scenario in grouped_sum_acqcost_hyp['Scenario'].unique():
        # Filtering the df for the current Scenario
        scenario_df = grouped_sum_acqcost_hyp[grouped_sum_acqcost_hyp['Scenario'] == scenario]
        # Calculating the accumulated $ Volume
        grouped_sum_acqcost_hyp.loc[scenario_df.index, 'Accum. Acq Cost'] = scenario_df['Total Acq Cost'].cumsum()


    # Efficient - Creating a DF for each Scenario
    for scenario in grouped_sum_acqcost_eff['Scenario'].unique():
        if scenario != -1:  # Scenario -1 is only indicative of nullity, it is not a real scenario
            '''
            Adding to the List of Dataframes from dict "scenario_dataframs" 2 new DFs.
            First one is already Efficient DF and Second one is Hypothetical DF, for Parts Qty
            Third and Fourth will be Acq Cost DFs for Efficient and Hypothetical, respectively
            '''
            tmp_filtered_scenario_acqcost_df = grouped_sum_acqcost_eff[grouped_sum_acqcost_eff['Scenario'] == scenario]
            scenario_df = df_dates_eff.merge(tmp_filtered_scenario_acqcost_df, left_on='Date', right_on='Date', how='left')
            # Filling in nulls
            scenario_df['Scenario'] = scenario_df['Scenario'].fillna(scenario).astype(int)
            scenario_df['Total Acq Cost'] = scenario_df['Total Acq Cost'].fillna(0)

            # Filling empty months for Accumulated Qty based on last record
            scenario_df['Accum. Acq Cost'].fillna(method='ffill', inplace=True)

            # Using specific Date column for Efficient Chart
            scenario_df['Order Date (Eff)'] = scenario_df['Order Date (Eff)'].fillna(scenario_df['Date'])
            scenario_df = scenario_df.drop('Date', axis=1)

            # Storing the DataFrame in the dictionary with the scenario name
            scenario_dataframes[f'Scenario_{int(scenario)}'].append(scenario_df)

    # Hypothetical - Creating a DF for each Scenario
    for scenario in grouped_sum_acqcost_hyp['Scenario'].unique():
        if scenario != -1:  # Scenario -1 is only indicative of nullity, it is not a real scenario
            '''
            Adding to the List of Dataframes from dict "scenario_dataframs" 2 new DFs.
            First one is already Efficient DF and Second one is Hypothetical DF, for Parts Qty
            Third and Fourth will be Acq Cost DFs for Efficient and Hypothetical, respectively
            '''
            tmp_filtered_scenario_acqcost_df = grouped_sum_acqcost_hyp[grouped_sum_acqcost_hyp['Scenario'] == scenario]
            scenario_df = df_dates_hyp.merge(tmp_filtered_scenario_acqcost_df, left_on='Date', right_on='Date',
                                             how='left')
            # Filling in nulls
            scenario_df['Scenario'] = scenario_df['Scenario'].fillna(scenario).astype(int)
            scenario_df['Total Acq Cost'] = scenario_df['Total Acq Cost'].fillna(0)

            # Filling empty months for Accumulated Qty based on last record
            scenario_df['Accum. Acq Cost'].fillna(method='ffill', inplace=True)

            # Using specific Date column for Hypothetical Chart
            scenario_df['Delivery Date (Hyp)'] = scenario_df['Delivery Date (Hyp)'].fillna(scenario_df['Date'])
            scenario_df = scenario_df.drop('Date', axis=1)

            # Storing the DataFrame in the dictionary with the scenario name
            scenario_dataframes[f'Scenario_{int(scenario)}'].append(scenario_df)


    # ------------------------- Chart Generation ------------------------- #

    # List of colors, so that each Scenario has a specific color and facilitates differentiation
    colors_array = ['blue', 'orange', 'black', 'green', 'purple']

    # Image size
    width, height = 680, 435

    # Creating the chart for each Scenario, separately
    '''
    The Canvas objects should be passed as a list, as each Scenario demands a particular Chart (Canvas Object).
    Everytime that a new Scenario is created, this list is cleared and the object created for each scenario will be appended to list
    '''
    canvas_list_acqcost_eff.clear()
    canvas_list_acqcost_hyp.clear()

    # Efficient - Acq Cost
    for index, (scenario_name, scenario_df_list) in enumerate(scenario_dataframes.items()):
        # Creating a figure and axes to insert the chart
        fig, ax = plt.subplots(figsize=(width / 100, height / 100), layout='constrained')
        # Keeping background transparent
        fig.patch.set_facecolor("None")
        fig.patch.set_alpha(0)
        ax.set_facecolor('None')

        # Bars - Monthly Acq Cost
        bars = ax.bar(scenario_df_list[2]['Order Date (Eff)'], scenario_df_list[2]['Total Acq Cost'],
                      label=f'Scen. {index}',
                      color=colors_array[index])

        # Accumulated Line
        axs = ax.plot(scenario_df_list[2]['Order Date (Eff)'], scenario_df_list[2]['Accum. Acq Cost'],
                      label=f'Scen. {index}',
                      color=colors_array[index])
        # Configuring the axis
        plt.xticks(scenario_df_list[2].index[::3], scenario_df_list[2]['Order Date (Eff)'][::3], rotation=45, ha='right')

        # Getting the t0 date for the current Scenario and converting it to MM/YYYY format
        t0_date = pd.to_datetime(
            df_scope_with_scenarios.loc[df_scope_with_scenarios['Scenario'] == index, 't0'].values[0])
        t0_date = t0_date.strftime('%m/%Y')
        # Adding a vertical line at t0
        ax.axvline(x=t0_date, linestyle='--', color=colors_array[index], label=f't0: Scen. {index}')

        # Getting the acft_delivery_start date for the current Scenario and converting it to MM/YYYY format
        acft_delivery_start_date = pd.to_datetime(
            df_scope_with_scenarios.loc[df_scope_with_scenarios['Scenario'] == index, 'acft_delivery_start'].values[0])
        acft_delivery_start_date = acft_delivery_start_date.strftime('%m/%Y')
        # Adding a vertical line in acft_delivery_start
        ax.axvline(x=acft_delivery_start_date, linestyle='dotted', color=colors_array[index],
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

        ax.axvspan(material_delivery_start_date, material_delivery_end_date, alpha=0.5, color=colors_array[index])

        # Adding a note at the point where the Build-Up is completed (all items delivered)
        index_max_acc_cost = scenario_df_list[2]['Accum. Acq Cost'].idxmax()
        x_max = scenario_df_list[2].loc[index_max_acc_cost, 'Order Date (Eff)']
        y_max = scenario_df_list[2].loc[index_max_acc_cost, 'Accum. Acq Cost']
        plt.scatter(x_max, y_max, color=colors_array[index], marker='o', label=f'BUP Conclusion: {x_max}')

        # Chart Settings
        ax.set_ylabel('Acq Cost (US$) Order Qty')
        ax.tick_params(axis='both', labelsize=9)  # Adjusting labels size
        ax.set_title(f'Efficient Curve (US$): {scenario_name}', color='green', fontweight='bold')
        ax.grid(True)

        # Function to format y-axis (float) to money format in million (US$ X M)
        def y_axis_acqcost_fmt(x, _):
            return f'U$ {x/1e6:.2f}M'
        
        # Setting formatter function for y axis
        ax.yaxis.set_major_formatter(FuncFormatter(y_axis_acqcost_fmt))

        # Inserting chart into Canvas
        canvas_acqcost_eff = FigureCanvasTkAgg(fig, master=efficient_curve_window)
        canvas_acqcost_eff.draw()
        # Configuring Canvas background
        canvas_acqcost_eff.get_tk_widget().configure(background='#cfcfcf')

        # Annotation function to connect with mplcursors - Bars
        def set_annotations_bars_eff(sel):
            sel.annotation.set_text(
                'Date: ' + str(scenario_df_list[2]['Order Date (Eff)'][sel.target.index]) + '\n' +
                'Order Qty: U$' + f"{scenario_df_list[2]['Total Acq Cost'][sel.target.index] / 1e3:.0f}k"        
            )

        # Annotation function to connect with mplcursors - Lines
        def set_annotations_lines_eff(sel):
            order_date = scenario_df_list[2]['Order Date (Eff)'][round(sel.target.index)]
            accum_acq_cost = scenario_df_list[2]['Accum. Acq Cost'][round(sel.target.index)]
            sel.annotation.set_text(
                f'Date: {order_date}\nOrder Qty: U$ {accum_acq_cost / 1e6:.2f}M'
            )

        # Inserting Hover with mplcursors
        # Bars
        mpc.cursor(bars, hover=True).connect('add', lambda sel: set_annotations_bars_eff(sel))
        # Line
        mpc.cursor(axs, hover=True).connect('add', lambda sel: set_annotations_lines_eff(sel))

        # Appending to List
        canvas_list_acqcost_eff.append(canvas_acqcost_eff)


    # Hypothetical - Acq Cost
    for index, (scenario_name, scenario_df_list) in enumerate(scenario_dataframes.items()):
        # Creating a figure and axes to insert the chart
        fig, ax = plt.subplots(figsize=(width / 100, height / 100), layout='constrained')
        # Keeping background transparent
        fig.patch.set_facecolor("None")
        fig.patch.set_alpha(0)
        ax.set_facecolor('None')

        # Bars - Monthly Acq Cost
        bars = ax.bar(scenario_df_list[3]['Delivery Date (Hyp)'], scenario_df_list[3]['Total Acq Cost'],
                      label=f'Scen. {index}',
                      color=colors_array[index])

        # Accumulated Line
        axs = ax.plot(scenario_df_list[3]['Delivery Date (Hyp)'], scenario_df_list[3]['Accum. Acq Cost'],
                      label=f'Scen. {index}',
                      color=colors_array[index])
        # Configuring the axis
        plt.xticks(scenario_df_list[3].index[::3], scenario_df_list[3]['Delivery Date (Hyp)'][::3], rotation=45, ha='right')

        # Getting the t0 date for the current Scenario and converting it to MM/YYYY format
        t0_date = pd.to_datetime(
            df_scope_with_scenarios.loc[df_scope_with_scenarios['Scenario'] == index, 't0'].values[0])
        t0_date = t0_date.strftime('%m/%Y')
        # Adding a vertical line at t0
        ax.axvline(x=t0_date, linestyle='--', color=colors_array[index], label=f't0: Scen. {index}')

        # Getting the acft_delivery_start date for the current Scenario and converting it to MM/YYYY format
        acft_delivery_start_date = pd.to_datetime(
            df_scope_with_scenarios.loc[df_scope_with_scenarios['Scenario'] == index, 'acft_delivery_start'].values[0])
        acft_delivery_start_date = acft_delivery_start_date.strftime('%m/%Y')
        # Adding a vertical line in acft_delivery_start
        ax.axvline(x=acft_delivery_start_date, linestyle='dotted', color=colors_array[index],
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

        ax.axvspan(material_delivery_start_date, material_delivery_end_date, alpha=0.5, color=colors_array[index])

        # Adding a note at the point where the Build-Up is completed (all items delivered)
        index_max_acc_cost = scenario_df_list[3]['Accum. Acq Cost'].idxmax()
        x_max = scenario_df_list[3].loc[index_max_acc_cost, 'Delivery Date (Hyp)']
        y_max = scenario_df_list[3].loc[index_max_acc_cost, 'Accum. Acq Cost']
        plt.scatter(x_max, y_max, color=colors_array[index], marker='o', label=f'BUP Conclusion: {x_max}')

        # Chart Settings
        ax.set_ylabel('Acq Cost (US$) Delivered Qty')
        ax.tick_params(axis='both', labelsize=9)  # Adjusting labels size
        ax.set_title(f'Hypothetical Curve (US$): {scenario_name}', color='green', fontweight='bold')
        ax.grid(True)

        # Function to format y-axis (float) to money format in million (US$ X M)
        def y_axis_acqcost_fmt(x, _):
            return f'U$ {x/1e6:.2f}M'
        
        # Setting formatter function for y axis
        ax.yaxis.set_major_formatter(FuncFormatter(y_axis_acqcost_fmt))

        # Inserting chart into Canvas
        canvas_acqcost_hyp = FigureCanvasTkAgg(fig, master=hypothetical_curve_window)
        canvas_acqcost_hyp.draw()
        # Configuring Canvas background
        canvas_acqcost_hyp.get_tk_widget().configure(background='#cfcfcf')

        # Annotation function to connect with mplcursors - Bars
        def set_annotations_bars_hyp(sel):
            sel.annotation.set_text(
                'Date: ' + str(scenario_df_list[3]['Delivery Date (Hyp)'][sel.target.index]) + '\n' +
                'Delivered Qty: U$' + f"{scenario_df_list[3]['Total Acq Cost'][sel.target.index] / 1e3:.0f}k"        
            )

        # Annotation function to connect with mplcursors - Lines
        def set_annotations_lines_hyp(sel):
            order_date = scenario_df_list[3]['Delivery Date (Hyp)'][round(sel.target.index)]
            accum_acq_cost = scenario_df_list[3]['Accum. Acq Cost'][round(sel.target.index)]
            sel.annotation.set_text(
                f'Date: {order_date}\nDelivered Qty: U$ {accum_acq_cost / 1e6:.2f}M'
            )

        # Inserting Hover with mplcursors
        # Bars
        mpc.cursor(bars, hover=True).connect('add', lambda sel: set_annotations_bars_hyp(sel))
        # Line
        mpc.cursor(axs, hover=True).connect('add', lambda sel: set_annotations_lines_hyp(sel))

        # Appending to List
        canvas_list_acqcost_hyp.append(canvas_acqcost_hyp)




@function_timer
def generate_cost_avoidance_screen(cost_avoidance_screen: ctk.CTkFrame, scenario_dataframes: dict, scenarios_list: list, df_scope_with_scenarios: pd.DataFrame,
                                   df_dates_eff, df_dates_hyp, bup_cost: float):

    # Global variable to store charts Canvas (each Scenario produces a particular Chart)
    global canvas_list_cost_avoidance

    # WACC
    wacc_value = 5.42 # Mock 07/10/24 as analyzed in cost of debt proportion. Cost of equity are 13,41% as beta for ERJ is 1.54, 10-Year Tresury rates are 4.1% and 6% ERP.
    # Equity to Debt ratio is 63/37% so full WACC, considering equity, would be higher as Cost of Equity considers Equity Risk Premium
    doublevar_wacc = ctk.DoubleVar(cost_avoidance_screen, value=wacc_value)
    # Calculating monthly Cost of Capital based on WACC variable (compounded mode)
    monthly_wacc = ((1+(wacc_value/100))**(1/12)-1)*100
    # Daily Cost of Capital
    daily_wacc = ((1+(monthly_wacc/100))**(1/30)-1)*100


    # Colors list, so that each Scenario has a specific color and facilitates differentiation
    colors_array = ['blue', 'orange', 'black', 'green', 'purple']

    # --------------------- GUI Elements ---------------------

    # WACC Elements
    lbl_wacc = ctk.CTkLabel(cost_avoidance_screen, text=r'WACC (% in US$): ',
                            font=ctk.CTkFont('open sans', size=12, weight='bold'),
                            )
    lbl_wacc.place(relx=0.88, rely=0.023, anchor="e")

    entry_wacc = ctk.CTkEntry(cost_avoidance_screen,
                              textvariable=doublevar_wacc,
                              width=50, height=12)
    entry_wacc.place(relx=0.94, rely=0.023, anchor=ctk.CENTER)

    # Slider values settings
    from_value: float = -50
    to_value: float = +50
    steps = 100
    doublevar_operat_eff_variation = ctk.DoubleVar(value=0)

    # Function to format Cost Avoidance Frame values depending on digits qty
    def format_k_m_pattern(x):
        if abs(x) >= 1_000_000:
            formatted_string = f'US$ {abs(x) / 1e6:.2f}M'
        elif abs(x) >= 1_000:
            formatted_string = f'US$ {abs(x) / 1e3:.2f}K'
        else:
            formatted_string = round(abs(x), 2)
        return formatted_string

    # Slider callback function to update Label text
    def update_label(value):
        # If value is positive
        if doublevar_operat_eff_variation.get() >= 0:
            lbl_efficiency_variation_num.configure(text=f'Efficiency Gain: +{doublevar_operat_eff_variation.get()}%',
                                                   text_color='green')
            # Calculating Efficiency Gain and new Procurement Length
            procur_len_simulated = int(355 * (1 - (doublevar_operat_eff_variation.get()/100))) # mock
            procur_len_gain = int(355 * (0 + (doublevar_operat_eff_variation.get()/100))) # mock
            lbl_procur_len_simul_num.configure(text=f'{procur_len_simulated} (-{procur_len_gain})', text_color='green')

            # Calculating Additional Savings based on daily Cost of Capital
            additional_savings = (procur_len_gain * (daily_wacc/100)) * bup_cost
            # Updating Additional Savings/Cost label
            lbl_additional_savings_simul.configure(text=f'{format_k_m_pattern(additional_savings)}', text_color='green')

        # If negative
        else:
            lbl_efficiency_variation_num.configure(text=f'Efficiency Loss: {doublevar_operat_eff_variation.get()}%',
                                                   text_color='red')
            # Calculating Efficiency Loss and new Procurement Length
            procur_len_simulated = int(355 * (1 - (doublevar_operat_eff_variation.get()/100))) # mock
            procur_len_loss = int(355 * (0 + (doublevar_operat_eff_variation.get()/100))) # mock
            lbl_procur_len_simul_num.configure(text=f'{procur_len_simulated} ({procur_len_loss})', text_color='red')

            # Calculating Additional Costs based on daily Cost of Capital
            additional_costs = (procur_len_loss * (daily_wacc/100)) * bup_cost
            # Updating Additional Savings/Cost label
            lbl_additional_savings_simul.configure(text=f'-{format_k_m_pattern(additional_costs)}', text_color='red')

    
    # Operational Efficiency Parameters - Full Supply Chain steps
    lbl_operat_eff = ctk.CTkLabel(cost_avoidance_screen,
                                  text='Procurement Length Simulation',
                                  font=ctk.CTkFont('open sans', size=12, weight='bold')
                                  )
    lbl_operat_eff.place(relx=0.5, rely=0.72, anchor=ctk.CENTER)

    # Slider
    slider = ctk.CTkSlider(cost_avoidance_screen,
                           from_=from_value, to=to_value,
                           number_of_steps=steps,
                           variable=doublevar_operat_eff_variation,
                           command=update_label,
                           progress_color='green',
                           button_color='#009898',
                           button_hover_color='#009898',
                           fg_color='red'
                           )
    
    slider.place(relx=0.5, rely=0.79, anchor=ctk.CENTER)

    # Label indicating percentage changing on Procurement Length parameters (Operational Efficiency Gain/Loss)
    lbl_efficiency_variation_num = ctk.CTkLabel(cost_avoidance_screen,
                                                font=ctk.CTkFont('open sans', size=14, weight='bold'),
                                                text_color='green',
                                                text=f'Efficiency Change: {doublevar_operat_eff_variation.get()}%')
    lbl_efficiency_variation_num.place(relx=0.5, rely=0.845, anchor=ctk.CENTER)
    

    # --------------------- Procurement Length Frame ---------------------
    # Todo: add Full Procurement length for Scenario and the variation (in days) of Efficiency Gain/Loss from slider.
    procur_length_frame = ctk.CTkFrame(cost_avoidance_screen, width=180, height=170,
                                       corner_radius=20)
    procur_length_frame.place(relx=0.15, rely=0.8, anchor=ctk.CENTER)

    # Scenario Procurement Length label
    lbl_scen_procur_length = ctk.CTkLabel(procur_length_frame, text="Procurement Length (days):",
                                          font=ctk.CTkFont('open sans', size=11, weight='bold'))
    lbl_scen_procur_length.place(relx=0.5, rely=0.12, anchor=ctk.CENTER)
    # Procurement Length Fixed Number label
    lbl_scen_procur_length_num = ctk.CTkLabel(procur_length_frame, text="355", #mock
                                          font=ctk.CTkFont('open sans', size=22, weight='bold'))
    lbl_scen_procur_length_num.place(relx=0.5, rely=0.32, anchor=ctk.CENTER)

    # Procurement Length Simulation label
    lbl_procur_len_simulation = ctk.CTkLabel(procur_length_frame, text='Simulation (days):',
                                             font=ctk.CTkFont('open sans', size=11, weight='bold'))
    lbl_procur_len_simulation.place(relx=0.5, rely=0.52, anchor=ctk.CENTER)
    # Procurement Length Simulation number
    lbl_procur_len_simul_num = ctk.CTkLabel(procur_length_frame, text='355', #mock
                                             font=ctk.CTkFont('open sans', size=22, weight='bold'))
    lbl_procur_len_simul_num.place(relx=0.5, rely=0.72, anchor=ctk.CENTER)

    #  --------------------- Cost Avoidance Frame ---------------------
    cost_avoidance_frame = ctk.CTkFrame(cost_avoidance_screen, width=180, height=170,
                                       corner_radius=20)
    cost_avoidance_frame.place(relx=0.85, rely=0.8, anchor=ctk.CENTER)

    # Efficient Curve Savings
    # To do: A new frame (right) to show fixed: Efficient/Hypothetical difference savings, Additional Savings/Costs (red/green) considering WACC and efficiency slider
    lbl_efficient_curve_savings = ctk.CTkLabel(cost_avoidance_frame, text='Efficient Curve Savings:',
                                               font=ctk.CTkFont('open sans', size=11, weight='bold'))
    lbl_efficient_curve_savings.place(relx=0.5, rely=0.12, anchor=ctk.CENTER)

    # Additional Savings
    lbl_add_savings_and_costs = ctk.CTkLabel(cost_avoidance_frame, text='Additional Savings/Costs:',
                                             font=ctk.CTkFont('open sans', size=11, weight='bold'))
    lbl_add_savings_and_costs.place(relx=0.5, rely=0.52, anchor=ctk.CENTER)
    
    # --------------------- Cost Avoidance Chart Generation ---------------------
    consolidated_dates = pd.concat([df_dates_eff, df_dates_hyp], axis=0).drop_duplicates()

    # Creating Cost Avoidance DataFrame for each Scenario
    for scenario, (scenario_name, scenario_df_list) in enumerate(scenario_dataframes.items()):

        # Adding Efficient Accumulated info - Acq Cost
        scenario_df_costavoid = consolidated_dates.merge(right=scenario_df_list[2][['Order Date (Eff)', 'Accum. Acq Cost']], how='left',
                                                         left_on='Date', right_on='Order Date (Eff)')
        # Renaming Accum. Acq Cost column and droping Order Date (Eff) column
        scenario_df_costavoid.rename(columns={'Accum. Acq Cost': 'Accum. Acq Cost (Eff)'}, inplace=True)
        scenario_df_costavoid = scenario_df_costavoid.drop(['Order Date (Eff)'], axis=1)

        # Adding Hypothetical Accumulated info - Acq Cost
        scenario_df_costavoid = scenario_df_costavoid.merge(right=scenario_df_list[3][['Delivery Date (Hyp)', 'Accum. Acq Cost']], how='left',
                                                         left_on='Date', right_on='Delivery Date (Hyp)')
        # Renaming Accum. Acq Cost column and droping Order Date (Eff) column
        scenario_df_costavoid.rename(columns={'Accum. Acq Cost': 'Accum. Acq Cost (Hyp)'}, inplace=True)
        scenario_df_costavoid = scenario_df_costavoid.drop(['Delivery Date (Hyp)'], axis=1)
       

        # test - creating Date column copy as a datetime object in order to compare
        scenario_df_costavoid['Date_dt'] = pd.to_datetime(scenario_df_costavoid['Date'], format='%m/%Y')

        # Storing the DataFrame in the dictionary with the scenario name
        scenario_dataframes[f'Scenario_{int(scenario)}'].append(scenario_df_costavoid)

    # Creating Chart Canvas for each Scenario, separately
    '''
    The Canvas objects should be passed as a list, as each Scenario demands a particular Chart (Canvas Object).
    Everytime that a new Scenario is created, this list is cleared and the object created for each scenario will be appended to list
    '''
    canvas_list_cost_avoidance.clear()

    # Image Size
    width, height = 680, 280

    for index, (scenario_name, scenario_df_list) in enumerate(scenario_dataframes.items()):

        # Creating a figure and axes to insert the chart
        fig, ax = plt.subplots(figsize=(width / 100, height / 100), layout='constrained')
        # Keeping background transparent
        fig.patch.set_facecolor("None")
        fig.patch.set_alpha(0)
        ax.set_facecolor('None')

        # Configuring the axis
        plt.xticks(scenario_df_list[4].index[::3], scenario_df_list[4]['Date'][::3], rotation=45, ha='right')

        # Efficient Accumulated Line - Acq Cost
        eff_axs = ax.plot(scenario_df_list[4]['Date'], scenario_df_list[4]['Accum. Acq Cost (Eff)'], label='Efficient Curve', color=colors_array[index],
                          ls='dashed')
                
        # Getting 't0+X' date for current Scenario iterated
        scenario_pln_start_date = (scenarios_list[index]['t0'] + pd.DateOffset(months=scenarios_list[index]['hyp_t0_start'])).strftime('%m/%Y')
        scenario_pln_months_from_t0 = scenarios_list[index]['hyp_t0_start']
        # Getting Build-Up List Total Cost
        #bup_cost = scenario_df_list[4]['Accum. Acq Cost (Hyp)'].max()
        # Adding Info to pandas dataframe
        scenario_df_list[4]['Acq Amount Hyp'] = 0 # It should not be NaN in order to compare in fill_between() method
        scenario_df_list[4].loc[scenario_df_list[4]['Date'] == scenario_pln_start_date, 'Acq Amount Hyp'] = bup_cost

        # Making every 'Accum. Acq Cost (Eff)' NaN values be 0
        scenario_df_list[4]['Accum. Acq Cost (Eff)'] = scenario_df_list[4]['Accum. Acq Cost (Eff)'].fillna(0)


        # Hypothetical Order Qty (Acq Cost) happens all in 'T0+X' Date. This is the concept of Hypothetical curve. Buying everything all at once, with no cadence, when Planning starts.
        hyp_axs = ax.bar(scenario_df_list[4]['Date'], scenario_df_list[4]['Acq Amount Hyp'], 
                         label=f"Hypothetical t0+{str(scenario_pln_months_from_t0)} Purchase", 
                         color=colors_array[index])      
        
        # Creating Control Variable to manage the area to be filled
        scenario_df_list[4]['Fill Between Ctrl Variable'] = 0
        # First Acq Cost Amount (Efficient Curve)
        eff_first_acq_amount = scenario_df_list[4].loc[scenario_df_list[4]['Accum. Acq Cost (Eff)'] != 0, 'Accum. Acq Cost (Eff)'].iloc[0]


        # Filling Control Variable
        scenario_df_list[4].loc[scenario_df_list[4]['Fill Between Ctrl Variable'] == 0, 
                                'Fill Between Ctrl Variable'] = scenario_df_list[4]['Accum. Acq Cost (Eff)']  # For all 0 values, get the Accumulated Value (Efficient Line)
        
        # If the Hypothetical Purchase happens before Efficient Curve Start, I assign the first Efficient Curve purchase to the same date and fill following 0's
        # with the same value until reaching the Efficient Curve construction
        # scenario_df_list[4].loc[scenario_df_list[4]['Acq Amount Hyp'] != 0, 'Fill Between Ctrl Variable'] = eff_first_acq_amount  # In Hyp Purchase date
        
        # Hypothetical Purchase Date
        hyp_purchase_date = pd.to_datetime(scenario_pln_start_date, format='%m/%Y')
        # Efficient Curve Start Date
        eff_purchase_start_date = scenario_df_list[4].loc[scenario_df_list[4]['Accum. Acq Cost (Eff)'] != 0, 'Date_dt'].iloc[0]
        
        # Above mentioned Conditional. If not True, nothing is done
        if hyp_purchase_date < eff_purchase_start_date:
            # Assigning first Efficient Curve Purchase to the Hypothetical Purchase date
            scenario_df_list[4].loc[scenario_df_list[4]['Date_dt'] == hyp_purchase_date, 'Fill Between Ctrl Variable'] = eff_first_acq_amount
            # Filling forward 0's with this value until reaching Efficient Curve construction
            first_nonzero_idx_ctrl_var = scenario_df_list[4][scenario_df_list[4]['Fill Between Ctrl Variable'] != 0].index[0]
            # Updating Fill Between Ctrl Variable with 0's but only when its after the first value allocated
            scenario_df_list[4].loc[(scenario_df_list[4].index > first_nonzero_idx_ctrl_var) & (scenario_df_list[4]['Fill Between Ctrl Variable'] == 0),
                                    'Fill Between Ctrl Variable'] = eff_first_acq_amount
        else:
            pass

        # Creating Start Date and End Date for Efficient Curve in order to condition 'where' arg on fill_between() method
        start_date = pd.to_datetime(scenario_pln_start_date, format='%m/%Y')

        end_date = pd.to_datetime(
            scenario_df_list[4].loc[scenario_df_list[4]['Accum. Acq Cost (Eff)'] != 0, 'Date'].iloc[-1]
            , format='%m/%Y') # Efficient curve End Date (last month different than 0)
  
        # Filling Cost Avoidance area
        ax.fill_between(x=scenario_df_list[4]['Date'], y1=scenario_df_list[4]['Acq Amount Hyp'].max(), y2=scenario_df_list[4]['Fill Between Ctrl Variable'],
                        where=((scenario_df_list[4]['Date_dt'] >= start_date) & (scenario_df_list[4]['Date_dt']<= end_date)), interpolate=True, 
                        color=colors_array[index], alpha=0.2, hatch='\\', label='Cash Saved')
        
    
        # Chart Settings
        ax.set_ylabel('Acq Cost (US$) Delivered Qty')
        ax.tick_params(axis='both', labelsize=9)  # Adjusting labels size
        ax.set_title(f'Cost Avoidance (Efficient Asset Allocation)', color=colors_array[index], fontweight='bold')
        ax.grid(True)
        ax.legend(loc='lower right', fontsize=7, framealpha=0.8)

        # Function to format y-axis (float) to money format in million (US$ X M)
        def y_axis_acqcost_fmt(x, _):
            return f'U$ {x/1e6:.2f}M'
        
        # Setting formatter function for y axis
        ax.yaxis.set_major_formatter(FuncFormatter(y_axis_acqcost_fmt))

        # Inserting chart into Canvas
        canvas_cost_avoidance = FigureCanvasTkAgg(fig, master=cost_avoidance_screen)
        canvas_cost_avoidance.draw()
        # Configuring Canvas background
        canvas_cost_avoidance.get_tk_widget().configure(background='#cfcfcf')
        canvas_cost_avoidance.get_tk_widget().place(relx=0.5, rely=0.33, anchor=ctk.CENTER)

        # Appending Canvas to List
        canvas_list_cost_avoidance.append(canvas_cost_avoidance)

        # --------------------------------------------------- SAVINGS CALCULATION ---------------------------------------------------

        # Calculating Scenario Savings between Efficient Curve x Hypothetical t0+X purchase

        # Creating Raw Postponed Column (US$) with Efficient Curve
        scenario_df_list[4]['Raw Postponed Amount'] = 0
        # Creating the condition in which the values will be applied (Beginning in the date when Hypothetical Purchase was done)
        condition = scenario_df_list[4]['Date_dt'] >= hyp_purchase_date
        # Subtracting the Hypothetical Purchase Amount (Acq Cost) from Efficient Curve to get the difference on each month
        scenario_df_list[4].loc[condition, 'Raw Postponed Amount'] = bup_cost - scenario_df_list[4].loc[condition, 'Accum. Acq Cost (Eff)']

        # Calculating Monthly Savings
        scenario_df_list[4]['Postponed Savings (US$)'] = scenario_df_list[4]['Raw Postponed Amount'] * (monthly_wacc/100)
        
        # Total Savings Efficient x Hypothetical purchase
        total_savings_eff = round(scenario_df_list[4]['Postponed Savings (US$)'].sum(), 2)


        # Label Efficient Purchase Savings
        lbl_savings_eff = ctk.CTkLabel(cost_avoidance_frame, text=f'{format_k_m_pattern(total_savings_eff)}',
                            font=ctk.CTkFont('open sans', size=22, weight='bold'),
                            text_color='green',
                            )
        # Label Positioning
        lbl_savings_eff.place(relx=0.5, rely=0.32, anchor=ctk.CENTER)

        # Label Dynamic Additional Savings/Cost based on Efficiency Gain/Loss
        lbl_additional_savings_simul = ctk.CTkLabel(cost_avoidance_frame, text='0',
                            font=ctk.CTkFont('open sans', size=22, weight='bold'),
                            text_color='green'
                            )
        lbl_additional_savings_simul.place(relx=0.5, rely=0.72, anchor=ctk.CENTER)



@function_timer
def generate_batches_curve(batches_curve_window: ctk.CTkFrame, scenarios_list: list, df_scope_with_scenarios: pd.DataFrame) :
    '''
    Function that receives the input so as to generate the Build-Up Curve based on batches.
    '''
    df_scope_with_scenarios.to_excel('df_scope_with_scenarios.xlsx')
    # Adding 2 tabs to Batch Charts: Parts Qty & Acq Cost
    # TabView - Batch Charts
    tbv_batch_charts = ctk.CTkTabview(batches_curve_window, width=620, height=470, corner_radius=15,
                                      segmented_button_fg_color="#009898",
                                      segmented_button_unselected_color="#009898",
                                      segmented_button_selected_color="#006464",
                                      bg_color='#cfcfcf', fg_color='#cfcfcf')
    tbv_batch_charts.pack()
    tbv_batch_charts.add('Parts Qty')
    tbv_batch_charts.add('Acq Cost')
    # Naming Frames
    batch_qty_frame = tbv_batch_charts.tab('Parts Qty')
    batch_cost_frame = tbv_batch_charts.tab('Acq Cost')

    # If there's Batch Information, creates the chart, Else: it shows an alert label
    if scenarios_list[0]['batches_dates'] !=  None:

        # Batches List
        batches_dates_list = scenarios_list[0]['batches_dates'].split(',')
        # Convert batches_dates_list to datetime format and sort it ascending
        batches_dates_list = [datetime.strptime(date.strip(), "%d/%m/%Y") for date in batches_dates_list]
        batches_dates_list.sort()

        # Creating DataFrame with specific columns for Batch assignment
        pns_full_procurement_length = df_scope_with_scenarios[['PN', 'Ecode', 'Qty', 'Acq Cost','t0', 'hyp_t0_start', 'PN Procurement Length', 'Delivery Date Hypothetical']]
        pns_full_procurement_length['planning_start_date'] = pns_full_procurement_length.apply(lambda row: row['t0'] + relativedelta(months=row['hyp_t0_start']), axis=1) 

        # Assigning Part Numbers to specific Batches, being tested on ascending order
        def assign_batch(row):
            for i, batch_date in enumerate(batches_dates_list):
                if row['PN Procurement Length'] <= (batch_date - row['planning_start_date']).days:
                    return int(i + 1)
            return 'No Batch Assigned'
        
        pns_full_procurement_length['Batch'] = pns_full_procurement_length.apply(assign_batch, axis=1)

        # Create a new column 'Batch Date' based on the Batch number or 'No Batch Assigned'
        def assign_batch_date(row):
            if row['Batch'] == 'No Batch Assigned':
                return 'No Batch Assigned'
            else:
                # return batches_dates_list[row['Batch'] - 1]
                return batches_dates_list[row['Batch'] - 1].strftime("%d/%m/%Y")

        pns_full_procurement_length['Batch Date'] = pns_full_procurement_length.apply(assign_batch_date, axis=1)
        # Creating Delivery Month Date, so as to be the X-axis of Batches charts
        pns_full_procurement_length['Delivery Month Hyp'] = pns_full_procurement_length['Delivery Date Hypothetical'].dt.to_period('M')
        # Creating Total Part Acq Cost column (Acq Cost * Qty)
        pns_full_procurement_length['Total Part Acq Cost'] = round(pns_full_procurement_length['Acq Cost'] * pns_full_procurement_length['Qty'], 2)

        # Creating Grouped Sum Qty DataFrame to Generate Batches bar chart
        df_grouped_qty_delivery_date = (
            pns_full_procurement_length
            .groupby('Delivery Month Hyp')['PN']
            .nunique()
            .sort_index()
            .reset_index()
            .rename(columns={'PN': 'Distinct PNs Count'})
        )

        # Creating Cumulative Sum Qty to Generate Batches line chart
        df_grouped_qty_delivery_date['Cumulative Sum Qty'] = df_grouped_qty_delivery_date['Distinct PNs Count'].cumsum()

        # Creating Grouped Acq Cost DataFrame to Generate Batches bar chart
        df_grouped_acqcost_delivery_date = (
            pns_full_procurement_length
            .assign(AcqCostQty= lambda x: round(x['Acq Cost'] * x['Qty'], 2))
            .groupby('Delivery Month Hyp')['AcqCostQty']
            .sum()
            .sort_index()
            .reset_index()
            .rename(columns={'AcqCostQty': 'Total Acq Cost'})
        )

        # Creating Cumulative Sum Qty to Generate Acq Cost Batches line chart
        df_grouped_acqcost_delivery_date['Cumulative Acq Cost Qty'] = df_grouped_acqcost_delivery_date['Total Acq Cost'].cumsum()

        # For AXVSpan Batches chart insertion, further transformation is necessary
        df_batches = (
            pns_full_procurement_length
            .groupby(['Batch', 'Batch Date'])['Ecode']
            .nunique()
            .reset_index()
            .rename(columns={'Ecode': 'PNs Qty'})
        )

        # As it is for visualizing purposes only, I remove 'No Batch Assigned' rows
        df_batches = df_batches[df_batches['Batch'] != 'No Batch Assigned'].reset_index(drop=True)

        # Initializing new Batch Start Date column as a variable at first
        batch_start_date = []
        # Now I create the column Batch Start Date (date of Last Batch end or Planning Start)
        for i in range(len(df_batches)):
            # First Batch will have the Planning Date as Start Date
            if i == 0:
                start_date = pns_full_procurement_length.loc[0, 'planning_start_date'].strftime('%d/%m/%Y')
            # Other rows: use previous batch date
            else:
                start_date = df_batches.loc[i - 1, 'Batch Date']

            # Appending to list
            batch_start_date.append(start_date)

        # Creating Batch Start Date column
        df_batches['Batch Start Date'] = batch_start_date
        # Converting to datetime so as to match the other charts format
        df_batches[['Batch Start Date', 'Batch Date']] = df_batches[['Batch Start Date', 'Batch Date']].apply(
            lambda col: pd.to_datetime(col, format='%d/%m/%Y')
        )
        # Making a copy of df_batches without dropping "No Batch Assigned" Parts
        df_batches_full_info = df_batches.copy()
        # Adding Acq Cost (US$) per Batch info
        total_acq_cost_batch = pns_full_procurement_length.groupby('Batch Date')['Total Part Acq Cost'].sum().reset_index()
        total_acq_cost_batch = total_acq_cost_batch[total_acq_cost_batch['Batch Date'] != 'No Batch Assigned']
        # Forcing datetime type on Batch Date
        total_acq_cost_batch['Batch Date'] = pd.to_datetime(total_acq_cost_batch['Batch Date'], format='%d/%m/%Y')
        # Merging with df_batches
        df_batches = df_batches.merge(total_acq_cost_batch, on='Batch Date', how='left')

        # Based on Batches spreasheet, generates Chart Images for Parts Qty batch feature
        def create_qty_batch_chart(df_grouped_qty_delivery_date: pd.DataFrame):
            # Image size
            width, height = 680, 435
            # Creating figure and axes to insert the chart: Batch Line Items
            fig, ax = plt.subplots(figsize=(width / 100, height / 100),
                                   layout='constrained')  # Layout property that handles "cutting" axes labels
            # Keeping background transparent
            fig.patch.set_facecolor("None")
            fig.patch.set_alpha(0)
            ax.set_facecolor('None')

            # Common index (X-Axis) for all charts in the figure
            x_labels = df_grouped_qty_delivery_date['Delivery Month Hyp'].dt.to_timestamp()

            # Colors list, so that each Batch has one distinct axvspan
            colors_array = ['blue', 'orange', 'green', 'black', 'purple', 'gray']

            # For each Batch, create a AXVSpan Chart (Vertical Dintinction of Batches)
            for i, row in df_batches.iterrows():
                ax.axvspan(
                    xmin=row['Batch Start Date'],
                    xmax=row['Batch Date'],
                    ymin=0, ymax=1,
                    color=colors_array[i],
                    alpha=0.3,
                    label=f"B{row['Batch']} ({row['Batch Date'].strftime('%m/%Y')}) - {row['PNs Qty']} PNs"
                )

            # Bar Chart
            bar = ax.bar(x=x_labels,
                   height=df_grouped_qty_delivery_date['Distinct PNs Count'],
                   color='steelblue',
                   width=22,
                   edgecolor='black',
                   linewidth=0.4
                   )

            # Line Chart (Cumulative)
            line = ax.plot(x_labels,
                    df_grouped_qty_delivery_date['Cumulative Sum Qty'],
                    color='black',
                    marker='o',
                    markersize=4)

            # debug
            df_grouped_qty_delivery_date.to_excel('df_grouped_qty_delivery_date.xlsx')

            # Batch Chart Settings
            ax.set_ylabel('PNs Count')
            ax.set_xlabel('Date', loc='right')
            ax.set_title(f'PNs - All Line Items ({df_scope_with_scenarios["Ecode"].count()} PNs)', fontsize=10)
            ax.grid(True)
            plt.legend(loc='upper left', fontsize=8)
            # Rotating X labels
            ax.tick_params(axis='x', rotation=45)
            # Adjusting axis spacing to avoid cutting off labels
            plt.subplots_adjust(left=0.15, right=0.9, bottom=0.2, top=0.9)
            # Formatting x-axis from YYYY-MM to MM/YYYY
            ax.xaxis.set_major_formatter(DateFormatter('%m/%Y'))

            # Inserting chart into Canvas
            canvas_batch_chart_items = FigureCanvasTkAgg(fig, master=batch_qty_frame)
            canvas_batch_chart_items.draw()
            # Configuring Canvas background
            canvas_batch_chart_items.get_tk_widget().configure(background='#cfcfcf')
            canvas_batch_chart_items.get_tk_widget().place(relx=0.5, rely=0.46, anchor=ctk.CENTER)

            # --- Saving the chart image in BytesIO() (memory) so it is not necessary to save as a file ---
            tmp_img_batch_chart_items = BytesIO()
            fig.savefig(tmp_img_batch_chart_items, format='png', transparent=True)
            tmp_img_batch_chart_items.seek(0)

            # Keeping the image in an Image object
            batch_items_chart = Image.open(tmp_img_batch_chart_items)

            # Loading into a CTk Image object
            batch_items_image = ctk.CTkImage(batch_items_chart,
                                           dark_image=batch_items_chart,
                                           size=(600, 220))

            # Annotation functions to connect with mplcursors
            def set_annotations_bar(sel):
                # Extracting month from parameters of Annotation text
                annotation_text = sel.annotation.get_text()
                # Extracting value after x= and before \n (in annotation params)
                if 'x=' in annotation_text:
                    start = annotation_text.find('x=') +2
                    end = annotation_text.find('\n', start)
                    x_value = annotation_text[start:end]
                    # Formatting Data from YYYY-MM to MM/YYYY
                    formatted_date = '/'.join(x_value.split('-')[::-1])
                # Setting text
                sel.annotation.set_text(
                    'Date: ' + str(formatted_date) + "\n" +
                    'PNs Qty: ' + str(round(sel.target[1]))
                )
            def set_annotations_line(sel):
                # Extracting month from parameters of Annotation text
                annotation_text = sel.annotation.get_text()
                # Extracting value after x= and before \n (in annotation params)
                if 'x=' in annotation_text:
                    start = annotation_text.find('x=') + 2
                    end = annotation_text.find('\n', start)
                    x_value = annotation_text[start:end]
                    # Formatting Data from YYYY-MM to MM/YYYY
                    formatted_date = '/'.join(x_value.split('-')[::-1])
                # Setting text
                sel.annotation.set_text(
                    'Date: ' + str(formatted_date) + "\n" +
                    'PNs Qty (Accumulated): ' + str(round(sel.target[1]))
                )

            # Inserting Hover with mplcursors
            mpc.cursor(bar, hover=True).connect('add', lambda sel: set_annotations_bar(sel))
            mpc.cursor(line, hover=True).connect('add', lambda sel: set_annotations_line(sel))


        # Based on Batches spreasheet, generates Chart Images for Acq Cost batch feature
        def create_acqcost_batch_chart(df_grouped_acqcost_delivery_date: pd.DataFrame):
            # Image size
            width, height = 680, 435
            # Creating figure and axes to insert the chart: Batch Line Items
            fig, ax = plt.subplots(figsize=(width / 100, height / 100),
                                   layout='constrained')  # Layout property that handles "cutting" axes labels
            # Keeping background transparent
            fig.patch.set_facecolor("None")
            fig.patch.set_alpha(0)
            ax.set_facecolor('None')

            # Common index (X-Axis) for all charts in the figure
            x_labels = df_grouped_acqcost_delivery_date['Delivery Month Hyp'].dt.to_timestamp()

            # Colors list, so that each Batch has one distinct axvspan
            colors_array = ['blue', 'orange', 'green', 'black', 'purple', 'gray']
            # For each Batch, create a AXVSpan Chart (Vertical Dintinction of Batches)
            for i, row in df_batches.iterrows():
                ax.axvspan(
                    xmin=row['Batch Start Date'],
                    xmax=row['Batch Date'],
                    ymin=0, ymax=1,
                    color=colors_array[i],
                    alpha=0.3,
                    label=f"B{row['Batch']} ({row['Batch Date'].strftime('%m/%Y')}) - US$ {row['Total Part Acq Cost']/1_000_000:.2f} M"
                )

            # Bar Chart
            bar = ax.bar(x=x_labels,
                         height=df_grouped_acqcost_delivery_date['Total Acq Cost'],
                         color='steelblue',
                         width=22,
                         edgecolor='black',
                         linewidth=0.4
                         )

            # Line Chart (Cumulative)
            line = ax.plot(x_labels,
                           df_grouped_acqcost_delivery_date['Cumulative Acq Cost Qty'],
                           color='black',
                           marker='o',
                           markersize=4)

            #debug
            df_grouped_acqcost_delivery_date.to_excel('df_grouped_acqcost_delivery_date.xlsx')
            # Acq Cost Batch Chart Settings
            ax.set_ylabel('Acq Cost (US$)')
            ax.set_xlabel('Date', loc='right')
            ax.set_title(f'PNs - Acq Cost (US$ {max(df_grouped_acqcost_delivery_date["Cumulative Acq Cost Qty"])/1_000_000:.2f} M)',
                         fontsize=10)
            ax.grid(True)
            plt.legend(loc='upper left', fontsize=8)
            # Rotating X labels
            ax.tick_params(axis='x', rotation=45)
            # Adjusting axis spacing to avoid cutting off labels
            plt.subplots_adjust(left=0.15, right=0.9, bottom=0.2, top=0.9)
            # Formatting y-axis to show Millions (US$)
            def millions_formatter(x, pos):
                return f'{x/1_000_000:.2f}M' if x >= 1_000_000 else f'{x/1_000:.2f}K' if x >= 1_000 else f'{x:.0f}'

            ax.yaxis.set_major_formatter(FuncFormatter(millions_formatter))
            # Formatting x-axis from YYYY-MM to MM/YYYY
            ax.xaxis.set_major_formatter(DateFormatter('%m/%Y'))

            # Inserting chart into Canvas
            canvas_batch_chart_items = FigureCanvasTkAgg(fig, master=batch_cost_frame)
            canvas_batch_chart_items.draw()
            # Configuring Canvas background
            canvas_batch_chart_items.get_tk_widget().configure(background='#cfcfcf')
            canvas_batch_chart_items.get_tk_widget().place(relx=0.5, rely=0.46, anchor=ctk.CENTER)

            # --- Saving the chart image in BytesIO() (memory) so it is not necessary to save as a file ---
            tmp_img_batch_chart_items = BytesIO()
            fig.savefig(tmp_img_batch_chart_items, format='png', transparent=True)
            tmp_img_batch_chart_items.seek(0)

            # Keeping the image in an Image object
            batch_items_chart = Image.open(tmp_img_batch_chart_items)

            # Loading into a CTk Image object
            batch_items_image = ctk.CTkImage(batch_items_chart,
                                             dark_image=batch_items_chart,
                                             size=(600, 220))

            # Annotation functions to connect with mplcursors
            def set_annotations_bar(sel):
                # Extracting month from parameters of Annotation text
                annotation_text = sel.annotation.get_text()
                # Extracting value after x= and before \n (in annotation params)
                if 'x=' in annotation_text:
                    start = annotation_text.find('x=') + 2
                    end = annotation_text.find('\n', start)
                    x_value = annotation_text[start:end]
                    # Formatting Data from YYYY-MM to MM/YYYY
                    formatted_date = '/'.join(x_value.split('-')[::-1])
                # Setting text
                sel.annotation.set_text(
                    'Date: ' + str(formatted_date) + "\n" +
                    'Acq Cost (US$): ' + str(round(sel.target[1]/1_000, 2)) + " K"
                )

            def set_annotations_line(sel):
                # Extracting month from parameters of Annotation text
                annotation_text = sel.annotation.get_text()
                # Extracting value after x= and before \n (in annotation params)
                if 'x=' in annotation_text:
                    start = annotation_text.find('x=') + 2
                    end = annotation_text.find('\n', start)
                    x_value = annotation_text[start:end]
                    # Formatting Data from YYYY-MM to MM/YYYY
                    formatted_date = '/'.join(x_value.split('-')[::-1])
                # Setting text
                sel.annotation.set_text(
                    'Date: ' + str(formatted_date) + "\n" +
                    'Acq Cost US$ (Accumulated): ' + str(round(sel.target[1]/1_000_000, 2)) + " M"
                )

            # Inserting Hover with mplcursors
            mpc.cursor(bar, hover=True).connect('add', lambda sel: set_annotations_bar(sel))
            mpc.cursor(line, hover=True).connect('add', lambda sel: set_annotations_line(sel))

        # Calling create_qty_batch_chart() function
        create_qty_batch_chart(df_grouped_qty_delivery_date)
        # Calling create_acqcost_batch_chart() function
        create_acqcost_batch_chart(df_grouped_acqcost_delivery_date)

    else:
        # Label with the instruction to create a Scenario
        lbl_no_batches_curve = ctk.CTkLabel(batches_curve_window,
                                            text="No Batches Curve were defined on Scenario creation.",
                                            font=ctk.CTkFont('open sans', size=16, weight='bold', slant='italic'),
                                            fg_color='#cfcfcf',
                                            bg_color='#cfcfcf',
                                            text_color='#000000')
        lbl_no_batches_curve.place(rely=0.5, relx=0.5, anchor=ctk.CENTER)  
    


@function_timer
def save_chart_image(chart: Image, output_path: str, filename: str) -> None:
    # Function that receveis as argument the chart Image object, as well as output path and filename, and saves it.

    try:
        chart.save(output_path + '\\' + filename)
        messagebox.showinfo(title="Success!", message=str("Image was exported to: " + output_path + '\\' + filename))
    except Exception as ex:
        messagebox.showinfo(title="Error!", message=str(ex) + "\n\nPlease make sure that the Image file is "
                                                              "closed and you have access to the local Downloads folder.")

@function_timer
def read_stock_data():
    pass