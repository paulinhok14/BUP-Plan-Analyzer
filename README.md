# Build-Up Plan Analyzer
###### Version 1.0
### What Is?


This is a desktop application developed with the intent (at first) to provide the possibility for the Engineers to create and compare different Scenarios (with different Dates and Supply Chain parameters assumptions), mainly when dealing with contractual commitments for stock Build-Up in the aeronautical business.

It was developed on a procedural paradigm.
BUP_GUI.py file is the main one, where you can execute the application.
bup_plan_analyzer.py has the main functions of the system.
___
### How it Works

A scope file with the materials list corresponding to the contract is an essential input.
The information that should contain in the scope file (.xlsx Excel format) is as follows:
- 'PN': *text*
- 'ECODE': *numeric*
- 'QTY': *numeric*
- 'SPC': *numeric*
- 'EIS': *boolean* **('X' or not filled)**


#### __**Example:**__ 

![Excel Input Format](docs/excel_input_format.jpg)

**OBS**: It is important that the columns description is exactly the same as described in this documentation.  
The Excel file can have as many tabs as you want, as long as **the first tab contains the contractual scope information**, with columns mentioned above.

![Tabs Excel](docs/tabs_excel.png)

___
### How to Use

As an end-user, you have been granted access through a folder with the application files, with an executable (.exe) one.  
You can make it a desktop shortcut if you want. Run it doubleclicking.

![App Icon](docs/app_icon.png)

The main screen should appear after a while:

![Main Screen](docs/main_screen.jpg)

Click on 'Search Scope File'

![Click Scope File](docs/click_scope_file.jpg)

A window with explorer will appear. Choose your scope file and click on 

![Click Open File](docs/click_open_file.png)

It may take a while since complementary information is fetched at this moment in the application. You can check the average time elapsed for this function in the 'execution_info.log' file, also in the application folder.

![Loading Reading Scope](docs/loading_reading_scope.jpg)
![Log File](docs/log_file.png)

Having it done, another screen should appear, in the first tab called "Scope", where you can see all loaded parts with their respective complementary information.  
You can also make sure if the file loaded is correct, by checking file name, rows quantity and Contract Value (US$):

![Scope Tab](docs/scope_tab.png)

In the second tab, called "Leadtime Analysis", is presented two charts in order to evaluate Leadtime distribution:

![Leadtime Analysis Tab](docs/leadtime_analysis_tab.png)

In the third tab "Scenarios", is

![Scenarios Tab](docs/scenarios_tab.png)




