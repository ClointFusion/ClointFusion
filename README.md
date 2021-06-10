## Welcome to <img src="https://raw.githubusercontent.com/ClointFusion/Image_ICONS_GIFs/main/Cloint-LOGO.png" height="30"> , Made in India with &#10084;&#65039; 

</br>

<img src="https://raw.githubusercontent.com/ClointFusion/Image_ICONS_GIFs/main/CCE.PNG">

# Description: 
Cloint India Pvt. Ltd - Python functions for Automation (RPA)

# What is ClointFusion ?
ClointFusion is a company registered at Vadodara, Gujarat, India. ClointFusion is our home-grown Python based RPA platform for Software BOT development. We are working towards Common Man's RPA using AI. 

![PyPI](https://img.shields.io/pypi/v/ClointFusion?label=PyPI%20Version) ![PyPI - License](https://img.shields.io/pypi/l/ClointFusion?label=License) ![PyPI - Status](https://img.shields.io/pypi/status/ClointFusion?label=Release%20Status)             ![ClointFusion](https://snyk.io/advisor/python/ClointFusion/badge.svg) ![PyPI - Downloads](https://img.shields.io/pypi/dm/ClointFusion?label=PyPI%20Downloads) ![Libraries.io SourceRank](https://img.shields.io/librariesio/sourcerank/pypi/ClointFusion) ![PyPI - Format](https://img.shields.io/pypi/format/ClointFusion?label=PyPI%20Format) ![GitHub contributors](https://img.shields.io/github/contributors/ClointFusion/ClointFusion?label=Contributors) ![GitHub last commit](https://img.shields.io/github/last-commit/ClointFusion/ClointFusion?label=Last%20Commit) 

![GitHub Repo stars](https://img.shields.io/github/stars/ClointFusion/ClointFusion?label=Stars&style=social) ![Twitter URL](https://img.shields.io/twitter/url?style=social&url=https%3A%2F%2Ftwitter.com%2FClointFusion) ![YouTube Channel Subscribers](https://img.shields.io/youtube/channel/subscribers/UCIygBtp1y_XEnC71znWEW2w?style=social) ![Twitter Follow](https://img.shields.io/twitter/follow/ClointFusion?style=social)

# Test Drive ClointFusion on Google Colabs
<a href="https://colab.research.google.com/github/ClointFusion/ClointFusion/blob/master/ClointFusion_Labs.ipynb" target="_blank"><img src="https://colab.research.google.com/assets/colab-badge.svg" alt="Open In Colab\"/></a>

## Click here for <a href="https://github.com/ClointFusion/ClointFusion/blob/master/Release_Notes.txt" target="_blank"> Release Notes</a>

# Installation on your local computer

# ClointFusion is now supported in Windows / Ubuntu / macOS !

1. Please install Python 3.9.5 with 64 bit: <a href="https://www.python.org/downloads" target="_blank"> Python 3.9.5 64 Bit</a>

    Windows users may refer to these steps : <a href="https://dev.to/fharookshaik/install-clointfusion-in-windows-operating-system-clointfusion-2dae" target="_blank">Install ClointFusion in Windows Operating System</a>

2. It is recommended to run ClointFusion in a Virtual Environment.
Please refer these steps to create one, as per your OS: <a href="https://packaging.python.org/guides/installing-using-pip-and-virtual-environments/#creating-a-virtual-environment" target="_blank">Creating a virtual environment in Windows / Mac / Ubuntu</a>

3. Install ClointFusion by executing this package in command promt (with Admin rights): 

    ## pip install --upgrade ClointFusion
4. Open a new file in your favorite Python IDE and type: 

    ## import ClointFusion as cf

PS: Ubuntu users: May need to install some additional packages:
1) sudo apt-get install python3-tk python3-dev
2) sudo apt-get install -y fonts-symbola
3) sudo apt-get install scrot
4) sudo apt-get install libcairo2-dev libjpeg-dev libgif-dev
5) sudo apt-get install libgirepository1.0-dev
6) sudo apt-get install python3-apt
7) sudo apt-get install  python3-xlib

---
# ClointFusion First Run Setup: 
First time, when you import ClointFusion, you would be prompted to run ClointFusions's Automated Selftest, to check whether all functions of ClointFusion are compatible with your computer settings & configurations. 
You would receive an email with self-test report.

---

## Now access more than 130 functions (hit ctrl+space in your IDE)

* ## 6 gui functions, to take any input from user:

    cf.gui_get_any_file_from_user() : Generic function to accept file path from user using GUI. Returns the filepath value in string format.Default allows all files.

    cf.gui_get_consent_from_user() : Generic function to get consent from user using GUI. Returns the string 'yes' or 'no'

    cf.gui_get_dropdownlist_values_from_user() :  Generic function to accept one of the drop-down value from user using GUI. Returns all chosen values in list format.

    cf.gui_get_excel_sheet_header_from_user() : Generic function to accept excel path, sheet name and header from user using GUI. Returns all these values in disctionary format.

    cf.gui_get_folder_path_from_user() : Generic function to accept folder path from user using GUI. Returns the folderpath value in string format.

    cf.gui_get_any_input_from_user() : Generic function to accept any input (text / numeric) from user using GUI. Returns the value in string format.


* ## 8 functions on Mouse operations:

    cf.mouse_click() : Clicks at the given X Y Co-ordinates on the screen using ingle / double / tripple click(s). Optionally copies selected data to clipboard (works for double / triple clicks)

    cf.mouse_move() : Moves the cursor to the given X Y Co-ordinates

    cf.mouse_get_color_by_position() : Gets the color by X Y co-ordinates of the screen

    cf.mouse_drag_from_to() : Clicks and drags from X1 Y1 co-ordinates to X2 Y2 Co-ordinates on the screen

    cf.mouse_search_snip_return_coordinates_x_y() : Searches the given image on the screen and returns its center of X Y co-ordinates.

    cf.mouse_search_snips_return_coordinates_x_y() : Searches the given set of images on the screen and returns its center of X Y co-ordinates of FIRST OCCURANCE

    cf.mouse_search_snip_return_coordinates_box() : Searches the given image on the screen and returns the 4 bounds co-ordinates (x,y,w,h)

    cf.mouse_find_highlight_click() : Searches the given text on the screen, highlights and clicks it

* ## 5 functions on Window operations (works only in Windows OS):

    cf.window_show_desktop() : Minimizes all the applications and shows Desktop.

    cf.window_get_all_opened_titles_windows() : Gives the title of all the existing (open) windows.

    cf.window_activate_and_maximize_windows() : Activates and maximizes the desired window.

    cf.window_minimize_windows() :  Activates and minimizes the desired window.

    cf.window_close_windows() :  Close the desired window.

* ## 5 functions on Window Objects (works only in Windows OS):

    cf.win_obj_open_app() : Open any windows application.

    cf.win_obj_get_all_objects() : Print or Save all the windows object elements of an application.

    cf.win_obj_mouse_click() : Simulate high level mouse clicks on windows object elements.

    cf.win_obj_key_press() : Simulate high level Keypress on windows object elements.

    cf.win_obj_get_text() : Read text from windows object element.

* ## 2 functions on File operations:
    cf.file_rename() : Renames the given file name to new file name with same extension

    cf.file_get_json_details() : Returns all the details of the given section in a dictionary

* ## 6 functions on Folder operations:

    cf.folder_read_text_file() : Reads from a given text file and returns entire contents as a single list

    cf.folder_write_text_file() :  Writes given contents to a text file

    cf.folder_create() : while making leaf directory if any intermediate-level directory is missing, folder_create() method will create them all.

    cf.folder_create_text_file() : Creates Text file in the given path.

    cf.folder_get_all_filenames_as_list() : Get all the files of the given folder in a list.

    cf.folder_delete_all_files() : Deletes all the files of the given folder

* ## 21 functions on Excel operations:

    cf.excel_get_all_sheet_names() : Gives you all names of the sheets in the given excel sheet.

    cf.excel_create_cf.excel_file_in_given_folder()

    cf.excel_if_value_exists() : Check if a given value exists in given excel. Returns True / False

    cf.excel_create_file()

    cf.excel_copy_paste_range_from_to_sheet() : Pastes the copied data in specific range of the given excel sheet.

    cf.excel_get_row_column_count() : Gets the row and coloumn count of the provided excel sheet.

    cf.excel_copy_range_from_sheet() : Copies the specific range from the provided excel sheet and returns copied data as a list

    cf.excel_split_by_column() : Splits the excel file by Column Name

    cf.excel_split_the_file_on_row_count() :  Splits the excel file as per given row limit

    cf.excel_merge_all_files() : Merges all the excel files in the given folder

    cf.excel_drop_columns() : Drops the desired column from the given excel file

    cf.excel_sort_columns() : A function which takes excel full path to excel and column names on which sort is to be performed

    cf.excel_clear_sheet() :  Clears the contents of given excel files keeping header row intact

    cf.excel_set_single_cell() : Writes the given text to the desired column/cell number for the given excel file

    cf.excel_get_single_cell() : Gets the text from the desired column/cell number of the given excel file

    cf.excel_remove_duplicates() : Drops the duplicates from the desired Column of the given excel file

    cf.excel_vlook_up() : Performs excel_vlook_up on the given excel files for the desired columns. Possible values for how are "inner","left", "right", "outer"

    cf.excel_draw_charts() : Interactive data visualization function, which accepts excel file, X & Y column. Chart types accepted are bar , scatter , pie , sun , histogram , box  , strip. You can pass color column as well, having a boolean value.

    cf.excel_clean_data() : Cleans our data from lowercase / remove_digits / remove_diacritics / remove_stopwords / remove_whitespace

    cf.excel_describe_data() :  Describe statistical data for the given excel

    cf.excel_drag_drop_pivot_table() : Interactive Drag and Drop Pivot Table Generation
    
    cf.excel_change_corrupt_xls_to_xlsx() : Repair corrupt file to regular file and then convert it to xlsx.

    cf.excel_convert_xls_to_xlsx() : Converts given XLS file to XLSX

    cf.excel_apply_template_format_save_to_new() : Converts given excel to Template Excel

    cf.excel_apply_format_as_table() : Applies table format to the used range of the given excel

    cf.excel_split_based_on_row_conditions_unique() : Splits the excel based on user defined row/column conditions

* ## 3 functions on Keyboard operations:

    cf.key_hit_enter() : Enter key will be pressed once.

    cf.key_press() : Emulates the given keystrokes.

    cf.key_write_enter() : Writes/Types the given text and press enter (by default) or tab key.


* ## 2 functions on Screenscraping operations:

    cf.scrape_save_contents_to_notepad : Copy pastes all the available text on the screen to notepad and saves it.

    cf.scrape_get_contents_by_search_copy_paste : Gets the focus on the screen by searching given text using crtl+f and performs copy/paste of all data. Useful in Citrix applications. This is useful in Citrix applications


* ## 12 functions on Browser operations:

    cf.browser_get_html_text() : Function to get HTML text without tags using Beautiful soup

    cf.browser_get_html_tabular_data_from_website() : Web Scrape HTML Tables : Gets Website Table Data Easily as an Excel using Pandas. Just pass the URL of Website having HTML Tables.

    cf.browser_navigate_h() : Navigates to Specified URL.

    cf.browser_write_h() :  Write a string on the given element.

    cf.browser_mouse_click_h() : Click on the given element.

    cf.browser_mouse_double_click_h() : Doubleclick on the given element.

    cf.browser_locate_element_h() : Find the element by Xpath, id or css selection.

    cf.browser_locate_elements_h() : Find the elements by Xpath, id or css selection.

    cf.browser_wait_until_h() : Wait until a specific element is found.

    cf.browser_refresh_page_h() : Refresh the page.

    cf.browser_quit_h() : Close the Helium browser.

    cf.browser_hit_enter_h() : Hits enter KEY using Browser Helium Functions

    cf.rowser_get_title_h() : Returns the Browser Window Title

    cf.browser_mouse_hover_h() : Performs a Mouse Hover over the given user Visible Text Element

    cf.browser_get_dropdown_options_h() : Returns the available options in the given labelled dropdown.

    cf.browser_select_dropdown_option_h() : Sets the Dropdown option either by label or by xpath with the given value.


* ## 3 functions on Alert Messages:

    cf.message_counter_down_timer() : Function to show count-down timer. Default is 5 seconds.

    cf.message_pop_up() : Specified message will popup on the screen for a specified duration of time.

    cf.message_flash() : Specified msg will popup for a specified duration of time with OK button.


* ## 3 functions on String Operations:

    cf.string_remove_special_characters() : Removes all the special character.

    cf.string_extract_only_alphabets() : Returns only alphabets from given input string

    cf.string_extract_only_numbers() : Returns only numbers from given input string


* ## Loads of miscellaneous functions related to emoji, capture photo, flash (pop-up) messages etc:

    cf.launch_any_exe_bat_application() : Launches any exe or batch file or excel file etc.

    cf.launch_website_h() : Internal function to launch browser.

    cf.schedule_create_task_windows() : Schedules (weekly & daily options as of now) the current BOT (.bat) using Windows Task Scheduler. Please call create_batch_file() function before using this function to convert .pyw file to .bat

    cf.schedule_delete_task_windows() : Deletes already scheduled task. Asks user to supply task_name used during scheduling the task. You can also perform this action from Windows Task Scheduler.

    cf.show_emoji() : Function which prints Emojis

    cf.message_counter_down_timer() : Function to show count-down timer. Default is 5 seconds.

    cf.get_long_lat() : Function takes zip_code as input (int) and returns longitude, latitude, state, city, county. 

    cf.dismantle_code() : This functions dis-assembles given function and shows you column-by-column summary to explain the output of disassembled bytecode.

    cf.ON_semi_automatic_mode() : This function sets semi_automatic_mode as True => ON

    cf.OFF_semi_automatic_mode() : This function sets semi_automatic_mode as False => OFF

    cf.camera_capture_image() : turn ON camera & take photo 

    cf.convert_csv_to_excel() : Function to convert CSV to Excel 

    cf.capture_snip_now() :  Captures the snip and stores in Image Folder of the BOT by giving continous numbering

    cf.take_error_screenshot() : Takes screenshot of an error popup parallely without waiting for the flow of the program. The screenshot will be saved in the log folder for reference.

    cf.find_text_on_screen() : Clears previous search and finds the provided text on screen.

    cf.word_cloud_from_url() : Function to create word cloud from a given website

    cf.isNaN() : Returns TRUE if a given value is NaN False otherwise

    cf.excel_sub_routines() : Excel VBA Macros called from ClointFusion

    cf.email_send_via_desktop_outlook() : Send email using Outlook from Desktop email application

    cf.email_send_gmail_via_api() : Sends gmail using API. User needs to supply his client_secret.json as parameter

    cf.download_this_file() : Downloads a given url file to BOT output folder or Browser's Download folder

    cf.excel_to_colored_html() : Converts given Excel to HTML preserving the Excel format and saves in same folder as .html
    
# ClointFusion's function works in different modes:
1) If you pass all the required parameters, function works silently. So, this is expert (Non-GUI) mode. This mode gives you more control over the function's parameters.

2) If you do not pass any parameter, GUI would pop-up asking you the required parameters. Next time, when you run the BOT, based upon your configuration, which you get to choose at the beginning of BOT run:

        A) If Semi-Automatic mode is OFF, GUI would pop-up again, showing you the previous entries, allowing you to modify the parameters.

        B) If Semi-Automatic mode in ON, BOT works silently taking your previous GUI entries.

    GUI Mode is for beginners. Anytime, if you are not getting how to use the function, just call an empty function (without parameters) and GUI would pop-up asking you for required parameters.

# We love your contribution
Contribute by giving a star / writing article on ClointFusion / feedback / report issues / bug fixes / feature enhancement / add documentation / many more ways as you please..

Participate in our monthly online hackathons & weekly meetups. Click here for more details: https://sites.google.com/view/clointfusion-hackathon

Please visit our GitHub repository: https://github.com/ClointFusion/ClointFusion

# Date &#10084;&#65039; with ClointFusion 
This an initiative for fast track entry into our growing workforce. For more details, please visit: https://lnkd.in/gh_r9YB

# Contact us: 
Drop a mail to ClointFusion@cloint.com