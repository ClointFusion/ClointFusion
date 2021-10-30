## Welcome to <img src="https://raw.githubusercontent.com/ClointFusion/Image_ICONS_GIFs/main/Cloint-LOGO-New.png" height="30"> , Made in India with &#10084;&#65039; 

<br>

<img src="https://raw.githubusercontent.com/ClointFusion/Image_ICONS_GIFs/main/CCEW.PNG">


## Description

Cloint India Pvt. Ltd - Python functions for Robotic Process Automation shortly `RPA`.

# What is ClointFusion?

ClointFusion is an Indian firm based in Vadodara, Gujarat. ClointFusion is a Python-based RPA platform for developing Software BOTs. Using AI, we're working on Common Man's RPA.

#### Check out Project Status

![PyPI](https://img.shields.io/pypi/v/ClointFusion?label=PyPI%20Version) 
![PyPI - License](https://img.shields.io/pypi/l/ClointFusion?label=License) 
![PyPI - Status](https://img.shields.io/pypi/status/ClointFusion?label=Release%20Status) 
![ClointFusion](https://snyk.io/advisor/python/ClointFusion/badge.svg) 
![PyPI - Downloads](https://img.shields.io/pypi/dm/ClointFusion?label=PyPI%20Downloads) 
![Libraries.io SourceRank](https://img.shields.io/librariesio/sourcerank/pypi/ClointFusion) 
![PyPI - Format](https://img.shields.io/pypi/format/ClointFusion?label=PyPI%20Format) 
![GitHub contributors](https://img.shields.io/github/contributors/ClointFusion/ClointFusion?label=Contributors) 
![GitHub last commit](https://img.shields.io/github/last-commit/ClointFusion/ClointFusion?label=Last%20Commit) 

![GitHub Repo stars](https://img.shields.io/github/stars/ClointFusion/ClointFusion?label=Stars&style=social) 
![Twitter URL](https://img.shields.io/twitter/url?style=social&url=https%3A%2F%2Ftwitter.com%2FClointFusion) 
![YouTube Channel Subscribers](https://img.shields.io/youtube/channel/subscribers/UCIygBtp1y_XEnC71znWEW2w?style=social) 
![Twitter Follow](https://img.shields.io/twitter/follow/ClointFusion?style=social)

## Release Notes

- Click here for <a href="https://github.com/ClointFusion/ClointFusion/blob/master/Release_Notes.txt" target="_blank"> Release Notes</a>

---

# Installation

## ClointFusion is now supported in Windows / Ubuntu / macOS !

1. Please install Python 3.9.7 with 64 bit: <a href="https://www.python.org/downloads" target="_blank"> Python 3.9.7 64 Bit</a>. Windows users may refer to these steps | <a href="https://dev.to/fharookshaik/install-clointfusion-in-windows-operating-system-clointfusion-2dae" target="_blank">Install ClointFusion in Windows Operating System</a>

2. It is recommended to run ClointFusion in a Virtual Environment. Please refer these steps to create one, as per your OS: <a href="https://packaging.python.org/guides/installing-using-pip-and-virtual-environments/#creating-a-virtual-environment" target="_blank">Creating a virtual environment in Windows / Mac / Ubuntu</a>

3. Install ClointFusion by executing this package in command promt (with Admin rights): 

```
pip install -U ClointFusion
```

4. Open a new file in your favorite Python IDE and type: 

```
import ClointFusion as cf
```

***PS: `Ubuntu` users may need to install some additional packages:***

```
sudo apt-get install python3-tk python3-dev
```

## The Voice Guided Tour of ClointFusion Functions | ClointFusion Automated Self-Test

When you import or update to a new version of ClointFusion, you'll be prompted with the `ClointFusion Automated Self-Test`, which highlights all of ClointFusion's 100+ features operating live on your computer while also confirming ClointFusion's compatibility with your PC's settings and configurations.
You will receive an email with a self-test report once you have completed the test successfully. 

## What do ClointFusion have?

ClointFusion offers 

- **More than 100 ready to use functions helpful in building BOTs**

<br>

### DOST - Your friend in automation

`DOST` is an interactive blockly based **`no code`** BOT Builder platform designed and optimized for BOT development using ClointFusion. We believe that automation is not just for programmers, and that a non-technical person can develop a BOT in minutes using DOST. 

#### Advantages of DOST 

- Easy to Use.
- Build BOT in minutes.
- No prior Programming knowledge needed.

#### Usage of DOST

Open your favourite terminal and type `dost`. That's it!


**Build BOT with DOST:** [DOST Website](https://dost.clointfusion.com/)

<br>
<img src="https://github.com/ClointFusion/Image_ICONS_GIFs/blob/main/DOST.gif?raw=true" height="400">

> Checkout more about Blockly by Google [here](https://developers.google.com/blockly)

<br>

### BOL - Your automation voice based assistant

`BOL` is voice based automation assistant designed to execute BOTs build out of ClointFusion without any human computer interaction.

#### Usage of BOL

Open your favourite terminal and type `bol`. Within a moment, a personalised Virtual Assistant will be at your service.

***Note:`BOL` is currently in deveopment stage. More functionalities are yet to be added.***

<br>
<img src="https://github.com/ClointFusion/Image_ICONS_GIFs/blob/main/DOST.gif?raw=true" height="400">

<br>
<br>

### WORK - The Work Hour Monitor

`WORK` is an intelligent application that detects each and every work you do in your PC and displays a detailed work report.

#### Usage of WORK

Open your favourite terminal and type `work`. A detailed work report will be displayed.

***Important NOTE: All the information that is being collected by `WORK` is stored in a securely maintained database in your system.***


<br>
<img src="https://github.com/ClointFusion/Image_ICONS_GIFs/blob/main/DOST.gif?raw=true" height="400">

<br>
<br>


---

## **Now access more than 100 functions (hit ctrl+space in your IDE)**

***TIP: You can find and inspect all of ClointFusion's functions using only one function i.e., `find()`. Just pass the partial name of the function.***

```
cf.find("sort")
cf.find("gui")
```
* ### 6 gui functions, to take any input from user:

| Function | Accepted Parameters | Description |
| :--------: | :----: | :----------- |
| cf.gui_get_any_input_from_user() | msgForUser="Please enter : ", password=False, multi_line=False, mandatory_field=True | Generic function to accept any input (text / numeric) from user using GUI. Returns the value in string format. |
| cf.gui_get_any_file_from_user() | msgForUser="the file : ", Extension_Without_Dot="*" | Generic function to accept file path from user using GUI. Returns the filepath value in string format.Default allows all files. |
| cf.gui_get_consent_from_user() | msgForUser="Continue ?" | Generic function to get consent from user using GUI. Returns the string 'yes' or 'no' |
| cf.gui_get_dropdownlist_values_from_user() | msgForUser=" ", dropdown_list=[], multi_select=True | Generic function to accept one of the drop-down value from user using GUI. Returns all chosen values in list format. |
| cf.gui_get_excel_sheet_header_from_user() | msgForUser=" " | Generic function to accept excel path, sheet name and header from user using GUI. Returns all these values in disctionary format. |
| cf.gui_get_folder_path_from_user() | msgForUser="the folder : " | Generic function to accept folder path from user using GUI. Returns the folderpath value in string format. |

<br>
<img src="https://github.com/ClointFusion/Image_ICONS_GIFs/blob/main/Functions%20Light%20GIFs/gui_function.gif?raw=true" height="400">

----

* ### 4 functions on Mouse Operations:

| Function | Accepted Parameters | Description |
| :--------: | :----: | :----------- |
| cf.mouse_click() | x=" ", y=" ", left_or_right="left", no_of_clicks=1  | Clicks at the given X Y Co-ordinates on the screen using ingle / double / tripple click(s). Optionally copies selected data to clipboard (works for double / triple clicks) |
| cf.mouse_move() | x=" ", y=" " | Moves the cursor to the given X Y Co-ordinates |
| cf.mouse_drag_from_to() | x1=" ", y1=" ", x2=" ",y2=" ", delay=0.5 | Clicks and drags from X1 Y1 co-ordinates to X2 Y2 Co-ordinates on the screen |
| cf.mouse_search_snip_return_coordinates_x_y() | img=" ", wait=180 | Searches the given image on the screen and returns its center of X Y co-ordinates. |

<br>
<img src="https://github.com/ClointFusion/Image_ICONS_GIFs/blob/main/Mouse_Operations.gif?raw=true" height="400">

----

* ### 6 functions on Window Operations (works only in Windows OS):

| Function | Accepted Parameters | Description |
| :--------: | :----: | :----------- |
| cf.window_show_desktop() | None | Minimizes all the applications and shows Desktop. |
| cf.window_get_all_opened_titles_windows() | window_title=" " | Gives the title of all the existing (open) windows. |
| cf.window_activate_and_maximize_windows() | windowName=" " | Activates and maximizes the desired window. |
| cf.window_minimize_windows() | windowName=" " | Activates and minimizes the desired window. | 
| cf.window_close_windows() | windowName=" " | Close the desired window. |
| cf.launch_any_exe_bat_application() | pathOfExeFile=" " | Launches any exe or batch file or excel file etc. |

<br>
<img src="https://github.com/ClointFusion/Image_ICONS_GIFs/blob/main/Functions%20Light%20GIFs/Window Operations.gif?raw=true" height="400">

----
    
* ### 5 functions on Window Objects (works only in Windows OS):

| Function | Accepted Parameters | Description |
| :--------: | :----: | :----------- |
| cf.win_obj_open_app() | title, program_path_with_name, file_path_with_name=" ", backend="uia" | Open any windows application. |
| cf.win_obj_get_all_objects() | main_dlg, save=False, file_name_with_path=" " | Print or Save all the windows object elements of an application. |
| cf.win_obj_mouse_click() | main_dlg,title=" ",  auto_id=" ", control_type=" " | Simulate high level mouse clicks on windows object elements. |
| cf.win_obj_key_press() | main_dlg,write, title=" ", auto_id=" ", control_type=" " | Simulate high level Keypress on windows object elements. |
| cf.win_obj_get_text() | main_dlg, title=" ",  auto_id=" ", control_type=" ", value = False | Read text from windows object element. |

<br>
<img src="https://github.com/ClointFusion/Image_ICONS_GIFs/blob/main/Windows_Object_Operation.gif?raw=true" height="400">

----

* ### 8 functions on Folder Operations:

| Function | Accepted Parameters | Description |
| :--------: | :----: | :----------- |
| cf.folder_read_text_file() | txt_file_path=" " | Reads from a given text file and returns entire contents as a single list |
| cf.folder_write_text_file() | txt_file_path=" ", contents=" " |  Writes given contents to a text file |
| cf.folder_create() | strFolderPath=" " | When you are making leaf directory, if any intermediate-level directory is missing, folder_create() method creates them. |
| cf.folder_create_text_file() | textFolderPath=" ", txtFileName=" " | Creates text file in the given path. |
| cf.folder_get_all_filenames_as_list() | strFolderPath=" ", extension='all' | Get all the files of the given folder in a list. |
| cf.folder_delete_all_files() | fullPathOfTheFolder=" ", file_extension_without_dot="all" | Deletes all the files of the given folder |
| cf.file_rename() | old_file_path='', new_file_name='', ext=False | Renames the given file name to new file name with same extension. |
|cf.file_get_json_details() | path_of_json_file='', section='' | Returns all the details of the given section in a dictionary |

<br>
<img src="https://github.com/ClointFusion/Image_ICONS_GIFs/blob/main/Folder_Operations.gif?raw=true" height="400">

----

* ### 28 functions on Excel Operations:

| Function | Accepted Parameters | Description |
| :--------: | :----: | :----------- |
| cf.excel_get_all_sheet_names() | excelFilePath=" " | Gives you all names of the sheets in the given excel sheet. |
| cf.excel_create_excel_file_in_given_folder() | fullPathToTheFolder=" ", excelFileName=" ", sheet_name="Sheet1" | Creates an excel file in the desired folder with desired filename |
| cf.excel_if_value_exists() | excel_path=" ", sheet_name="Sheet1", header=0, usecols=" ", value=" " | Check if a given value exists in given excel. Returns True / False |
| cf.excel_create_file() | fullPathToTheFile=" ", fileName=" ", sheet_name="Sheet1" | Create a Excel file in fullPathToTheFile with filename. |
| cf.excel_copy_paste_range_from_to_sheet() | excel_path=" ", sheet_name="Sheet1",<br> startCol=0, startRow=0, endCol=0,<br> endRow=0, copiedData=" " | Pastes the copied data in specific range of the given excel sheet. |
| cf.excel_get_row_column_count() | excel_path=" ", sheet_name="Sheet1", header=0 | Gets the row and coloumn count of the provided excel sheet. |
| cf.excel_copy_range_from_sheet() | excel_path=" ", sheet_name="Sheet1", startCol=0, startRow=0, endCol=0, endRow=0 | Copies the specific range from the provided excel sheet and returns copied data as a list |
| cf.excel_split_by_column() | excel_path=" ", sheet_name="Sheet1",<br> header=0, columnName=" " | Splits the excel file by Column Name |
| cf.excel_split_the_file_on_row_count() | excel_path=" ", sheet_name = "Sheet1", rowSplitLimit=" ", outputFolderPath=" ", outputTemplateFileName ="Split" |  Splits the excel file as per given row limit |
| cf.excel_merge_all_files() | input_folder_path=" ", output_folder_path=" " | Merges all the excel files in the given folder |
| cf.excel_drop_columns() | excel_path=" ", sheet_name="Sheet1",<br> header=0, columnsToBeDropped = " " | Drops the desired column from the given excel file |
| cf.excel_sort_columns() | excel_path=" ", sheet_name="Sheet1",<br> header=0, firstColumnToBeSorted=None, secondColumnToBeSorted=None, thirdColumnToBeSorted=None, firstColumnSortType=True, secondColumnSortType=True, thirdColumnSortType=True,<br> view_excel=False | A function which takes excel full path to excel and column names on which sort is to be performed |
| cf.excel_clear_sheet() | excel_path=" ",sheet_name="Sheet1",<br> header=0 |  Clears the contents of given excel files keeping header row intact |
| cf.excel_set_single_cell() | excel_path=" ", sheet_name="Sheet1",<br> header=0, columnName=" ", cellNumber=0, setText=" " | Writes the given text to the desired column/cell number for the given excel file |
| cf.excel_get_single_cell() | excel_path=" ",sheet_name="Sheet1",<br> header=0, columnName=" ",cellNumber=0 | Gets the text from the desired column/cell number of the given excel file |
| cf.excel_remove_duplicates() | excel_path=" ",sheet_name="Sheet1",<br> header=0, columnName=" ", saveResultsInSameExcel=True, which_one_to_keep="first" | Drops the duplicates from the desired Column of the given excel file |
| cf.excel_vlook_up() | filepath_1=" ", sheet_name_1 = "Sheet1",<br> header_1 = 0, filepath_2=" ", sheet_name_2 = "Sheet1",<br> header_2 = 0, Output_path=" ", OutputExcelFileName=" ", match_column_name=" ", how='left', view_excel=False | Performs excel_vlook_up on the given excel files for the desired columns. Possible values for how are "inner","left", "right", "outer" |
| cf.excel_describe_data() | excel_path=" ",sheet_name="Sheet1", header=0, view_excel=False |  Describe statistical data for the given excel |
| cf.excel_change_corrupt_xls_to_xlsx() | xls_file ='',xlsx_file = '', xls_sheet_name='' | Repair corrupt excel file |
| cf.excel_get_all_header_columns() | excel_path=" ",sheet_name="Sheet1",header=0 | Gives you all column header names of the given excel sheet |
| cf.excel_convert_to_image() | excel_file_path=" " | Returns an Image (PNG) of given Excel |
| cf.excel_split_on_user_defined_conditions() | excel_file_path, sheet_name="Sheet1", column_name='', condition_strings=None,output_dir='', view_excel=False | Splits the excel based on user defined row/column conditions |
| cf.excel_apply_format_as_table() | excel_file_path, table_style="TableStyleMedium21", sheet_name="Sheet1" | Applies table format to the used range of the given excel |
| cf.excel_convert_xls_to_xlsx() | xls_file_path='',xlsx_file_path='' | Converts given XLS file to XLSX |
| cf.isNaN() | value | Returns TRUE if a given value is NaN False otherwise |
| cf.convert_csv_to_excel() | csv_path=" ", sep=" " | Function to convert CSV to Excel | 
| cf.excel_sub_routines() | None | Excel VBA Macros called from ClointFusion |
| cf.excel_to_colored_html() | formatted_excel_path=" " | Converts given Excel to HTML preserving the Excel format and saves in same folder as .html |

<br>
<img src="https://github.com/ClointFusion/Image_ICONS_GIFs/blob/main/Excel_Operations.gif?raw=true" height="400">

----

* ### 3 functions on Keyboard Operations:

| Function | Accepted Parameters | Description |
| :--------: | :----: | :----------- |
| cf.key_hit_enter() | write_to_window=" " | Enter key will be pressed once. |
| cf.key_press() | key_1='', key_2='', key_3='', write_to_window=" " | Emulates the given keystrokes. |
| cf.key_write_enter() | text_to_write=" ", write_to_window=" ", delay_after_typing=1, key="e" | Writes/Types the given text and press enter (by default) or tab key. |

<br>
<img src="https://github.com/ClointFusion/Image_ICONS_GIFs/blob/main/KB_Operations.gif?raw=true" height="400">

----

* ### 5 functions on Screenscraping Operations:

| Function | Accepted Parameters | Description |
| :--------: | :----: | :----------- |
|cf.scrape_save_contents_to_notepad() | folderPathToSaveTheNotepad=" ", switch_to_window=" ",X=0, Y=0 | Copy pastes all the available text on the screen to notepad and saves it. |
| cf.scrape_get_contents_by_search_copy_paste() | highlightText=" " | Gets the focus on the screen by searching given text using crtl+f and performs copy/paste of all data. Useful in Citrix applications. This is useful in Citrix applications |
| cf.screen_clear_search() | delay=0.2 | Clears previously found text (crtl+f highlight) |
| cf.search_highlight_tab_enter_open() | searchText=" ", hitEnterKey="Yes", shift_tab='No' | Searches for a text on screen using crtl+f and hits enter. This function is useful in Citrix environment. |
| cf.find_text_on_screen() | searchText=" ", delay=0.1, occurance=1, isSearchToBeCleared=False | Clears previous search and finds the provided text on screen. |

<br>
<img src="https://github.com/ClointFusion/Image_ICONS_GIFs/blob/main/Screen_Scraping.gif?raw=true" height="400">

----

* ### 11 functions on Browser Operations:

| Function | Accepted Parameters | Description |
| :--------: | :----: | :----------- |
| cf.browser_activate() | url=" ", files_download_path='', dummy_browser=True,<br> open_in_background=False, incognito=False,<br> clear_previous_instances=False, profile="Default" | Function to launch browser and start the session. |
| cf.browser_navigate_h() | url=" " | Navigates to Specified URL. |
| cf.browser_write_h() | Value=" ",  User_Visible_Text_Element=" " |  Write a string on the given element. |
| cf.browser_mouse_click_h() | User_Visible_Text_Element=" ", element=" ",<br> double_click=False, right_click=False | Click on the given element. |
|cf.browser_locate_element_h() | selector=" ", get_text=False,<br> multiple_elements=False | Find the element by Xpath, id or css selection. |
| cf.browser_wait_until_h() | text=" ", element="t" | Wait until a specific element is found. |
| cf.browser_refresh_page_h() | None | Refresh the page. |
| cf.browser_quit_h() | None | Close the Helium browser. |
| cf.browser_hit_enter_h() | None | Hits enter KEY using Browser Helium Functions |
| cf.browser_key_press_h() | key_1=" ", key_2=" " | Type text using Browser Helium Functions and press hot keys |
| cf.browser_mouse_hover_h() | User_Visible_Text_Element=" " | Performs a Mouse Hover over the Given User Visible Text Element | 

<br>
<img src="https://github.com/ClointFusion/Image_ICONS_GIFs/blob/main/Functions%20Light%20GIFs/browser_functions.gif?raw=true" height="400">

----

* ### 4 functions on Alert Messages:

| Function | Accepted Parameters | Description |
| :--------: | :----: | :----------- |
| cf.message_counter_down_timer() | strMsg="Calling ClointFusion Function in (seconds)", start_value=5 | Function to show count-down timer. Default is 5 seconds. |
| cf.message_pop_up() | strMsg=" ", delay=3 | Specified message will popup on the screen for a specified duration of time.|
| cf.message_flash() | msg=" ", delay=3 | Specified msg will popup for a specified duration of time with OK button. |
| cf.message_toast() | message,website_url=" ", file_folder_path=" " | Function for displaying Windows 10 Toast Notifications. Pass website URL OR file / folder path that needs to be opened when user clicks on the toast notification. |

----

* ### 3 functions on String Operations: 

| Function | Accepted Parameters | Description |
| :--------: | :----: | :----------- |
| cf.string_remove_special_characters() | inputStr=" " | Removes all the special character. |
| cf.string_extract_only_alphabets() | inputString=" " | Returns only alphabets from given input string |
| cf.string_extract_only_numbers() | inputString=" " | Returns only numbers from given input string |

<br>
<img src="https://github.com/ClointFusion/Image_ICONS_GIFs/blob/main/String_Operations.gif?raw=true" height="400">

----

* ### Loads of miscellaneous functions related to emoji, capture photo, flash (pop-up) messages etc:

| Function | Accepted Parameters | Description |
| :--------: | :----: | :----------- |
| cf.clear_screen() | None | Clears Python Interpreter Terminal Window Screen |
| cf.print_with_magic_color() | strMsg:str=" ", magic:bool=False | Function to color and format terminal output |
| cf.schedule_create_task_windows() | Weekly_Daily="D", week_day="Sun", start_time_hh_mm_24_hr_frmt="11:00" | Schedules (weekly & daily options as of now) the current BOT (.bat) using Windows Task Scheduler. Please call create_batch_file() function before using this function to convert .pyw file to .bat |
| cf.schedule_delete_task_windows() | None | Deletes already scheduled task. Asks user to supply task_name used during scheduling the task. You can also perform this action from Windows Task Scheduler. |
| cf.show_emoji() | strInput=" " | Function which prints Emojis |
| cf.dismantle_code() | strFunctionName=" " | This functions dis-assembles given function and shows you column-by-column summary to explain the output of disassembled bytecode. |
| cf.ON_semi_automatic_mode() | None | This function sets semi_automatic_mode as True => ON |
| cf.OFF_semi_automatic_mode()| None | This function sets semi_automatic_mode as False => OFF |
| cf.email_send_via_desktop_outlook() | toAddress=" ", ccAddress=" ", subject=" ",htmlBody=" ", embedImgPath=" ", attachmentFilePath=" " | Send email using Outlook from Desktop email application |
| cf.download_this_file() | url=" " | Downloads a given url file to BOT output folder or Browser's Download folder |
| cf.pause_program() | seconds="5" | Stops the program for given seconds |
| cf.string_regex() | inputStr=" ", strExpAfter=" ", <br> strExpBefore=" ", intIndex=0 | Regex API service call, to search within a given string data |
| cf.ocr_now() | img_path=" " | Recognize and read the text embedded in images using Google's Tesseract-OCR |
| cf.update_log_excel_file() | message=" " | Given message will be updated in the excel log file of output folder |
| cf.create_batch_file() | application_exe_pyw_file_path=" " | Creates .bat file for the given application / exe or even .pyw BOT developed by you. This is required in Task Scheduler. |

<br>
<img src="https://github.com/ClointFusion/Image_ICONS_GIFs/blob/main/miscallaneous.gif?raw=true" height="400">    


<!-- # ClointFusion's function works in different modes: -->
# ClointFusion's Semi Automatic Mode

1. If you pass all the required parameters, function works silently. So, this is expert (Non-GUI) mode. This mode gives you more control over the function's parameters.
2. If you do not pass any parameter, GUI would pop-up asking you the required parameters. Next time, when you run the BOT, based upon your configuration, which you get to choose at the beginning of BOT run:
    -  If `Semi-Automatic mode` is OFF, GUI would pop-up again, showing you the previous entries, allowing you to modify the parameters.
    -  If `Semi-Automatic mode` in ON, BOT works silently taking your previous GUI entries.
    - Toggle `Semi-Automatic mode` by using the following command

    ```
    cf.ON_semi_automatic_mode   # To turn ON semi automatic mode
    cf.OFF_semi_automatic_mode  # To turn OFF semi automatic mode
    ```

3. GUI Mode is for beginners. Anytime, if you are not getting how to use the function, just call an empty function (without parameters) and GUI would pop-up asking you for required parameters.


<br>    
<img src="https://github.com/ClointFusion/Image_ICONS_GIFs/blob/main/Auto_Semi_Auto.gif?raw=true" height="400">    

# BOTS made out of ClointFusion

### Outlook Email BOT implemented using ClointFusion

<img src="https://github.com/ClointFusion/Image_ICONS_GIFs/blob/main/Functions%20Light%20GIFs/Gmail_and_Outlook_BOT.gif?raw=true" height="400">

<br>

# We love your contribution

Contribute to us by giving a star, writing articles on `ClointFusion`, giving comments, reporting bugs, bug fixes, feature enhancements, adding documentation, and many other ways. 


## Invitation to our Monthly Branded Hackathon

We also invite everyone to take part in our monthly branded event, the `ClointFusion Hackathon`, and stand a chance to work with us.

Checkout our Hackathon Website for more details here: [ClointFusion Hackathon](https://sites.google.com/view/clointfusion-hackathon
)

<br>

## Date &#10084;&#65039; with ClointFusion

This an initiative for fast track entry into our growing workforce. For more details, please visit: [Date with ClointFusion](https://lnkd.in/gh_r9YB)


## Acknowledgements

We sincerely thanks to all it's dependent packages for the great contribution, which made `ClointFusion` possible!

Please find all the dependencies [here](https://openbase.com/python/ClointFusion/dependencies) 
<!-- 
<a href="https://openbase.com/python/ClointFusion/dependencies" target="blank">ClointFusion thanks all its dependent packages for the great contribution, which has made ClointFusion possible !</a> -->

## Credits

#### ReadMe File Maintainer 
fharookshaik, Intern @ ClointFusion. Incase of any queries reach him on 

<a href="https://www.linkedin.com/in/fharook-shaik-7a757b181/" target="_blank"><img src="https://img.shields.io/badge/linkedin-%230077B5.svg?style=for-the-badge&logo=linkedin&logoColor=white" alt="LinkedIn"></a> &nbsp;
<a href="https://github.com/fharookshaik" target="_blank"><img src="https://img.shields.io/badge/github-%23121011.svg?style=for-the-badge&logo=github&logoColor=white" alt="GitHub"></a> &nbsp;  


## Need help in Building BOTS?

Write us by clicking below<br>
<div align='left'>
<a href="mailto:ClointFusion@cloint.com" target="_blank"><img src="https://img.shields.io/badge/Gmail-D14836?style=for-the-badge&logo=gmail&logoColor=white" alt="Gmail"></a> &nbsp;
</div>

---

