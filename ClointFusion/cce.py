from ClointFusion import loader

# ---------  Libraries Imports | Current Count : 15
from loader import pd
try:
    from loader import pg
except:
    pass
from loader import clointfusion_directory
from loader import clipboard
from loader import re
from loader import op
from loader import os
from loader import time
from loader import shutil
from loader import sys
from loader import datetime
from loader import subprocess
from loader import traceback
from loader import logging
from loader import user_name
from loader import user_email
from loader import webbrowser


# ---------  Variables Imports | Current Count : 8
from loader import batch_file_path
from loader import output_folder_path
from loader import config_folder_path
from loader import img_folder_path
from loader import error_screen_shots_path
from loader import cf_icon_file_path
from loader import cf_logo_file_path
from loader import os_name
from loader import windows_os
from loader import linux_os
from loader import mac_os
from loader import browser

# ---------  GUI Functions ---------

def gui_get_consent_from_user(msgForUser="Continue ?"):    
    """
    Generic function to get consent from user using GUI. Returns the yes or no

    Default Text: "Do you want to "
    """
    return loader.gui_get_consent_from_user(msgForUser)

def gui_get_dropdownlist_values_from_user(msgForUser="",dropdown_list=[],multi_select=True): 
    """
    Generic function to accept one of the drop-down value from user using GUI. Returns all chosen values in list format.

    Default Text: "Please choose the item(s) from "
    """
    return loader.gui_get_dropdownlist_values_from_user(msgForUser,dropdown_list,multi_select)

def gui_get_excel_sheet_header_from_user(msgForUser=""): 
    """
    Generic function to accept excel path, sheet name and header from user using GUI. Returns all these values in disctionary format.

    Default Text: "Please choose the excel "
    """
    return loader.gui_get_excel_sheet_header_from_user(msgForUser)
    
def gui_get_folder_path_from_user(msgForUser="the folder : "):    
    """
    Generic function to accept folder path from user using GUI. Returns the folderpath value in string format.

    Default text: "Please choose "
    """
    return loader.gui_get_folder_path_from_user(msgForUser)

def gui_get_any_input_from_user(msgForUser="the value : ",password=False,multi_line=False,mandatory_field=True):   
    """
    Generic function to accept any input (text / numeric) from user using GUI. Returns the value in string format.
    Please use unique message (key) for each value.

    Default Text: "Please enter "
    """
    return loader.gui_get_any_input_from_user(msgForUser,password,multi_line,mandatory_field)

def gui_get_any_file_from_user(msgForUser="the file : ",Extension_Without_Dot="*"):    
    """
    Generic function to accept file path from user using GUI. Returns the filepath value in string format.Default allows all files i.e *

    Default Text: "Please choose "
    """
    return loader.gui_get_any_file_from_user(msgForUser,Extension_Without_Dot)

# ---------  GUI Functions Ends ---------



# ---------  Message  Functions --------- 

def message_counter_down_timer(strMsg="Calling ClointFusion Function in (seconds)",start_value=5):
    """
    Function to show count-down timer. Default is 5 seconds.
    Ex: message_counter_down_timer()
    """
    return loader.message_counter_down_timer(strMsg,start_value)

def message_pop_up(strMsg="",delay=3):
    """
    Specified message will popup on the screen for a specified duration of time.

    Parameters:
        strMsg  (str) : message to popup.
        delay   (int) : duration of the popup.
    """
    loader.message_pop_up(strMsg,delay)

def message_flash(msg="",delay=3):
    """
    specified msg will popup for a specified duration of time with OK button.

    Parameters:
        msg     (str) : message to popup.
        delay   (int) : duration of the popup.
    """
    loader.message_flash(msg,delay)

def message_toast(message,website_url="", file_folder_path=""):
    """
    Function for displaying Windows 10 Toast Notifications.
    Pass website URL OR file / folder path that needs to be opened when user clicks on the toast notification.
    """
    
    loader.message_toast(message,website_url, file_folder_path)

# ---------  Message  Functions Ends ---------



# ---------  Mouse Functions --------- 
    
def mouse_click(x='', y='', left_or_right="left", no_of_clicks=1):
    """Clicks at the given X Y Co-ordinates on the screen using single / double / triple click(s). Default clicks on current position.

    Args:
        x (int): x-coordinate on screen.
        Eg: 369 or 435, Defaults: ''.
        y (int): y-coordinate on screen.
        Eg: 369 or 435, Defaults: ''.
        left_or_right (str, optional): Which mouse button.
        Eg: right or left, Defaults: left.
        no_of_click (int, optional): Number of times specified mouse button to be clicked.
        Eg: 1 or 2, Max 3. Defaults: 1.

    Returns:
        bool: Whether the function is successful or failed.
    """
    return loader.mouse_click(x, y, left_or_right, no_of_clicks)

def mouse_move(x="",y=""):
    """Moves the cursor to the given X Y Co-ordinates.

    Args:
        x (int): x-coordinate on screen.
        Eg: 369 or 435, Defaults: ''.
        y (int): y-coordinate on screen.
        Eg: 369 or 435, Defaults: ''.

    Returns:
        bool: Whether the function is successful or failed.
    """
    return loader.mouse_move(x,y)

def mouse_drag_from_to(x1="",y1="",x2="",y2="",delay=0.5):
    """Clicks and drags from x1 y1 co-ordinates to x2 y2 Co-ordinates on the screen

    Args:
        x1 or x2 (int): x-coordinate on screen.
        Eg: 369 or 435, Defaults: ''.
        y1 or y2 (int): y-coordinate on screen.
        Eg: 369 or 435, Defaults: ''.
        delay (float, optional): Seconds to wait while performing action. 
        Eg: 1 or 0.5, Defaults to 0.5.

    Returns:
        bool: Whether the function is successful or failed.
    """
    return loader.mouse_drag_from_to(x1,y1,x2,y2,delay)

def mouse_search_snip_return_coordinates_x_y(img="", wait=180):
    """Searches the given image on the screen and returns its center of X Y co-ordinates.

    Args:
        img (str, optional): Path of the image. 
        Eg: D:\Files\Image.png, Defaults to "".
        wait (int, optional): Time you want to wait while program searches for image repeatably.
        Eg: 10 or 100 Defaults to 180.
        
    Returns:
        bool: If function is failed returns False.
        tuple (x, y): Image Center co-ordinates.
    """
    return loader.mouse_search_snip_return_coordinates_x_y(img, wait)

# ---------  Mouse Functions Ends --------- 



# ---------  Keyboard Functions --------- 

def key_press(key_1='', key_2='', key_3='', write_to_window=""):
    """Emulates the given keystrokes.

    Args:
        key_1 (str, optional): Enter the 1st key 
        Eg: ctrl or shift. Defaults to ''.
        key_2 (str, optional): Enter the 2nd key in combination. 
        Eg: alt or A. Defaults to ''.
        key_3 (str, optional): Enter the 3rd key in combination. 
        Eg: del or tab. Defaults to ''.
        write_to_window (str, optional): (Only in Windows) Name of Window you want to activate. Defaults to "".
        
    Supported Keys:
        ['\\t', '\\n', '\\r', ' ', '!', '"', '#', '$', '%', '&', "'", '(',')', '*', '+', ',', '-', '.', '/', 
        '0', '1', '2', '3', '4', '5', '6', '7','8', '9', ':', ';', '<', '=', '>', '?', '@', '[', '\\', ']', '^', '_', '`', 
        'a', 'b', 'c', 'd', 'e','f', 'g', 'h', 'i', 'j', 'k', 'l', 'm', 'n', 'o', 'p', 'q', 'r', 's', 't', 'u', 'v', 'w', 'x', 'y', 'z', 
        '{', '|', '}', '~', 'accept', 'add', 'alt', 'altleft', 'altright', 'apps', 'backspace',
        'browserback', 'browserfavorites', 'browserforward', 'browserhome',
        'browserrefresh', 'browsersearch', 'browserstop', 'capslock', 'clear',
        'convert', 'ctrl', 'ctrlleft', 'ctrlright', 'decimal', 'del', 'delete',
        'divide', 'down', 'end', 'enter', 'esc', 'escape', 'execute', 'f1', 'f10',
        'f11', 'f12', 'f13', 'f14', 'f15', 'f16', 'f17', 'f18', 'f19', 'f2', 'f20',
        'f21', 'f22', 'f23', 'f24', 'f3', 'f4', 'f5', 'f6', 'f7', 'f8', 'f9',
        'final', 'fn', 'hanguel', 'hangul', 'hanja', 'help', 'home', 'insert', 'junja',
        'kana', 'kanji', 'launchapp1', 'launchapp2', 'launchmail',
        'launchmediaselect', 'left', 'modechange', 'multiply', 'nexttrack',
        'nonconvert', 'num0', 'num1', 'num2', 'num3', 'num4', 'num5', 'num6',
        'num7', 'num8', 'num9', 'numlock', 'pagedown', 'pageup', 'pause', 'pgdn',
        'pgup', 'playpause', 'prevtrack', 'print', 'printscreen', 'prntscrn',
        'prtsc', 'prtscr', 'return', 'right', 'scrolllock', 'select', 'separator',
        'shift', 'shiftleft', 'shiftright', 'sleep', 'space', 'stop', 'subtract', 'tab',
        'up', 'volumedown', 'volumemute', 'volumeup', 'win', 'winleft', 'winright', 'yen',
        'command', 'option', 'optionleft', 'optionright']
    
    Returns:
        bool: Whether the function is successful or failed.
    """
    return loader.key_press(key_1, key_2, key_3, write_to_window)
    
def key_write_enter(text_to_write="", write_to_window="", delay_after_typing=1, key="e"):
    """Writes/Types the given text.

    Args:
        text_to_write (str, optional): Text you wanted to type
        Eg: ClointFusion is awesone. Defaults to "".
        write_to_window (str, optional): (Only in Windows) Name of Window you want to activate
        Eg: Notepad. Defaults to "".
        delay_after_typing (int, optional): Seconds in time to wait after entering the text
        Eg: 5. Defaults to 1.
        key (str, optional): Press Enter key after typing.
        Eg: t for tab. Defaults to e

    Returns:
        bool: Whether the function is successful or failed.
    """
    return loader.key_write_enter(text_to_write, write_to_window, delay_after_typing, key)

def key_hit_enter(write_to_window=""):
    """Enter key will be pressed once.

    Args:
        write_to_window (str, optional): (Only in Windows)Name of Window you want to activate.
        Eg: Notepad. Defaults to "".

    Returns:
        bool: Whether the function is successful or failed.
    """
    status = False
    return loader.key_hit_enter(write_to_window)

# --------- Keyboard Functions Ends --------- 



# ---------  Browser Functions --------- 

def browser_activate(url="", files_download_path='', dummy_browser=True, open_in_background=False, incognito=False,
                     clear_previous_instances=False, profile="Default"):
    """Function to launch browser and start the session.

    Args:
        url (str, optional): Website you want to visit. Defaults to "".
        files_download_path (str, optional): Path to which the files need to be downloaded.
        Defaults: ''.
        dummy_browser (bool, optional): If it is false Default profile is opened.
        Defaults: True.
        incognito (bool, optional): Opens the browser in incognito mode.
        Defaults: False.
        clear_previous_instances (bool, optional): If true all the opened chrome instances are closed.
        Defaults: False.
        profile (str, optional): By default it opens the 'Default' profile.
        Eg : Profile 1, Profile 2

    Returns:
        bool: Whether the function is successful or failed.
    """
    return loader.browser_activate(url, files_download_path, dummy_browser, open_in_background, incognito,
                     clear_previous_instances, profile)

def browser_navigate_h(url=""):
    """Navigate through the url after the session is started.

    Args:
        url (str, optional): Url which you want to visit.
        Defaults: "".

    Returns:
        bool: Whether the function is successful or failed.
    """
    return loader.browser_navigate_h(url)

def browser_write_h(Value="", User_Visible_Text_Element=""):
    """Write a string in browser, if User_Visible_Text_Element is given it writes on the given element.

    Args:
        Value (str, optional): String which has be written.
        Defaults: "".
        User_Visible_Text_Element (str, optional): The element which is visible(Like : Sign in).
        Defaults: "".

    Returns:
        bool: Whether the function is successful or failed.
    """
    return loader.browser_write_h(Value, User_Visible_Text_Element)

def browser_mouse_click_h(User_Visible_Text_Element="", element="", double_click=False, right_click=False):
    """Click on the given element.

    Args:
        User_Visible_Text_Element (str, optional): The element which is visible(Like : Sign in).
        Defaults: "".
        element (str, optional): Use locate_element to get element and use to click.
        Defaults: "".
        double_click (bool, optional): True to perform a Double click.
        Defaults: False.
        right_click (bool, optional): True to perform a Right click.
        Defaults: False.

    Returns:
        bool: Whether the function is successful or failed.
    """
    return loader.browser_mouse_click_h(User_Visible_Text_Element, element, double_click, right_click)

def browser_locate_element_h(selector="", get_text=False, multiple_elements=False):
    """Find the element by Xpath, id or css selection.

    Args:
        selector (str, optional): Give Xpath or CSS selector. Defaults to "".
        get_text (bool, optional): Give the text of the element. Defaults to False.
        multiple_elements (bool, optional): True if you want to get all the similar elements with matching selector as list. Defaults to False.

    Returns:
        element         : If only one element
        list of elements: If multiple_elements is True
    """
    return loader.browser_locate_element_h

def browser_wait_until_h(text="", element="t"):
    """Wait until a specific element is found.

    Args:
        text (str, optional): To wait until the string appears on the screen.
        Eg: Export Successfull Completed. Defaults: ""
        element (str, optional): Type of Element Whether its a Text(t) or Button(b).
        Defaults: "t - Text".

    Returns:
        bool: Whether the function is successful or failed.
    """
    return loader.browser_wait_until_h(text, element)

def browser_refresh_page_h():
    """Refresh the current active browser page.

    Returns:
        bool: Whether the function is successful or failed.
    """
    return loader.browser_refresh_page_h()

def browser_hit_enter_h():
    """Hits enter KEY in Browser

    Returns:
        bool: Whether the function is successful or failed.
    """
    return loader.browser_hit_enter_h()

def browser_key_press_h(key_1="", key_2=""):
    """Type text using Browser Helium Functions and press hot keys.

    Args:
        key_1 (str): Keys you want to simulate or string you want to press
        Eg: "tab" or "Murali". Defaults: ""
        key_2 (str, optional): Key you want to simulate with combination to key_1.
        Eg: "shift" or "escape". Defaults: ""

    Returns:
        bool: Whether the function is successful or failed.
    """
    return loader.browser_key_press_h(key_1, key_2)

def browser_mouse_hover_h(User_Visible_Text_Element=""):
    """Performs a Mouse Hover over the Given User Visible Text Element

    Args:
        User_Visible_Text_Element (str, optional): The element which is visible(Like : Sign in).
        Defaults: "".

    Returns:
        bool: Whether the function is successful or failed.
    """
    return loader.browser_mouse_hover_h(User_Visible_Text_Element)

def browser_quit_h():
    """Close the Browser or Browser Automation Session.

    Returns:
        bool: Whether the function is successful or failed.
    """
    return loader.browser_quit_h()

def browser_set_waiting_time(time=10):
    """
    Set the waiting time for the browser. If element is not found in the given time, it will raise an exception.

    Args:
        time ([int]): The time in seconds to wait for the element to be found.
    """
    return loader.browser_set_waiting_time(time)

# ---------  Browser Functions Ends --------- 



# ---------  Folder Functions ---------

def folder_read_text_file(txt_file_path=""):
    """
    Reads from a given text file and returns entire contents as a single list
    """
    return loader.folder_read_text_file(txt_file_path)

def folder_write_text_file(txt_file_path="",contents=""):
    """
    Writes given contents to a text file
    """
    loader.folder_write_text_file(txt_file_path,contents)

def folder_create(strFolderPath=""):
    """
    while making leaf directory if any intermediate-level directory is missing,
    folder_create() method will create them all.

    Parameters:
        folderPath (str) : path to the folder where the folder is to be created.

    For example consider the following path:

    """
    loader.folder_create(strFolderPath)

def folder_create_text_file(textFolderPath="",txtFileName=""):
    """
    Creates Text file in the given path.
    Internally this uses folder_create() method to create folders if the folder/s does not exist.
    automatically adds txt extension if not given in textFilePath.

    Parameters:
        textFilePath (str) : Complete path to the folder with double slashes.
    """
    loader.folder_create_text_file(textFolderPath,txtFileName)

def folder_get_all_filenames_as_list(strFolderPath="",extension='all'):
    """
    Get all the files of the given folder in a list.

    Parameters:
        strFolderPath  (str) : Location of the folder.
        extension      (str) : extention of the file. by default all the files will be listed regarless of the extension.
    
    returns:
        allFilesOfaFolderAsLst (List) : all the file names as a list.
    """
    return loader.folder_get_all_filenames_as_list(strFolderPath,extension)

def folder_delete_all_files(fullPathOfTheFolder="",file_extension_without_dot="all"):  
    """
    Deletes all the files of the given folder

    Parameters:
        fullPathOfTheFolder  (str) : Location of the folder.
        extension            (str) : extention of the file. by default all the files will be deleted inside the given folder 
                                    regarless of the extension.
    returns:
        count (int) : number of files deleted.
    """ 
    return loader.folder_delete_all_files(fullPathOfTheFolder,file_extension_without_dot)

def file_rename(old_file_path='',new_file_name='',ext=False):
    '''
    Renames the given file name to new file name with same extension
    '''
    loader.file_rename(old_file_path,new_file_name,ext)

def file_get_json_details(path_of_json_file='',section=''):
    '''
    Returns all the details of the given section in a dictionary 
    '''
    return loader.file_get_json_details(path_of_json_file,section)

# ---------  Folder Functions Ends ---------



# ---------  Window Operations Functions --------- 

def window_show_desktop():
    """
    Minimizes all the applications and shows Desktop.
    """
    loader.window_show_desktop()

def window_get_all_opened_titles_windows():
    """
    Gives the title of all the existing (open) windows.

    Returns:
        allTitles_lst  (list) : returns all the titles of the window as list.
    """
    return loader.window_get_all_opened_titles_windows()

def window_activate_and_maximize_windows(windowName=""):
    """
    Activates and maximizes the desired window.

    Parameters:
        windowName  (str) : Name of the window to maximize.
    """
    loader.window_activate_and_maximize_windows(windowName)

def window_minimize_windows(windowName=""):
    """
    Activates and minimizes the desired window.

    Parameters:
        windowName  (str) : Name of the window to miniimize.
    """
    loader.window_minimize_windows(windowName)

def window_close_windows(windowName=""):
    """
    Close the desired window.

    Parameters:
        windowName  (str) : Name of the window to close.
    """
    loader.window_close_windows(windowName)

def launch_any_exe_bat_application(pathOfExeFile=""):
    """Launches any exe or batch file or excel file etc.

    Args:
        pathOfExeFile (str, optional): Location of the file with extension 
        Eg: Notepad, TextEdit. Defaults to "".
    """
    return loader.launch_any_exe_bat_application(pathOfExeFile)

# ---------  Window Operations Functions Ends --------- 



# ---------  String Functions --------- 

def string_extract_only_alphabets(inputString=""):
    """
    Returns only alphabets from given input string
    """
    return loader.string_extract_only_alphabets(inputString)

def string_extract_only_numbers(inputString=""):
    """
    Returns only numbers from given input string
    """
    return loader.string_extract_only_numbers(inputString)   

def string_remove_special_characters(inputStr=""):
    """
    Removes all the special character.

    Parameters:
        inputStr  (str) : string for removing all the special character in it.

    Returns :
        outputStr (str) : returns the alphanumeric string.
    """

    return loader.string_remove_special_characters(inputStr)

def string_regex(inputStr="",strExpAfter="",strExpBefore="",intIndex=0):
    """
    Regex API service call, to search within a given string data
    """
    return loader.string_regex(inputStr,strExpAfter,strExpBefore,intIndex)

# ---------  String Functions Ends --------- 



# ---------  Excel Functions --------- 

def excel_get_row_column_count(excel_path="", sheet_name="Sheet1", header=0):
    """
    Gets the row and coloumn count of the provided excel sheet.

    Parameters:
        excel_path  (str) : Full path to the excel file with slashes.
        sheet_name           (str) : by default it is Sheet1.

    Returns:
        row (int) : number of rows
        col (int) : number of coloumns
    """
    return loader.excel_get_row_column_count(excel_path, sheet_name, header)

def excel_copy_range_from_sheet(excel_path="", sheet_name='Sheet1', startCol=0, startRow=0, endCol=0, endRow=0): #*
    """
    Copies the specific range from the provided excel sheet and returns copied data as a list
    Parameters:
        excel_path :"Full path of the excel file with double slashes"
        sheet_name     :"Source sheet name from where contents are to be copied"
        startCol          :"Starting column number (index starts from 1) from where copying starts"
        startRow          :"Starting row number (index starts from 1) from where copying starts"
        endCol            :"Ending column number ex:4 upto where cells to be copied"
        endRow            :"Ending column number ex:5 upto where cells to be copied"

    Returns:
    rangeSelected        : the copied range data
    """
    return loader.excel_copy_range_from_sheet(excel_path, sheet_name, startCol, startRow, endCol, endRow)

def excel_copy_paste_range_from_to_sheet(excel_path="", sheet_name='Sheet1', startCol=0, startRow=0, endCol=0, endRow=0, copiedData=""):#*
    """
    Pastes the copied data in specific range of the given excel sheet.
    """
    return loader.excel_copy_paste_range_from_to_sheet(excel_path, sheet_name, startCol, startRow, endCol, endRow, copiedData)

def excel_split_by_column(excel_path="",sheet_name='Sheet1',header=0,columnName=""):#*
    """
    Splits the excel file by Column Name
    """
    loader.excel_split_by_column(excel_path,sheet_name,header,columnName)

def excel_split_the_file_on_row_count(excel_path="", sheet_name = 'Sheet1', rowSplitLimit="", outputFolderPath="", outputTemplateFileName ="Split"):#*
    """
    Splits the excel file as per given row limit
    """
    return loader.excel_split_the_file_on_row_count(excel_path, sheet_name, rowSplitLimit, outputFolderPath, outputTemplateFileName)

def excel_merge_all_files(input_folder_path="",output_folder_path=""):
    """
    Merges all the excel files in the given folder
    """
    return loader.excel_merge_all_files(input_folder_path,output_folder_path)

def excel_drop_columns(excel_path="", sheet_name='Sheet1', header=0, columnsToBeDropped = ""):
    """
    Drops the desired column from the given excel file
    """
    loader.excel_drop_columns(excel_path, sheet_name, header, columnsToBeDropped)

def excel_sort_columns(excel_path="",sheet_name='Sheet1',header=0,firstColumnToBeSorted=None,secondColumnToBeSorted=None,thirdColumnToBeSorted=None,firstColumnSortType=True,secondColumnSortType=True,thirdColumnSortType=True):#*
    """
    A function which takes excel full path to excel and column names on which sort is to be performed

    """
    return loader.excel_sort_columns(excel_path,sheet_name,header,firstColumnToBeSorted,secondColumnToBeSorted,thirdColumnToBeSorted,firstColumnSortType,secondColumnSortType,thirdColumnSortType)     

def excel_clear_sheet(excel_path="",sheet_name="Sheet1", header=0):
    """
    Clears the contents of given excel files keeping header row intact
    """
    loader.excel_clear_sheet(excel_path,sheet_name, header)

def excel_set_single_cell(excel_path="", sheet_name="Sheet1", header=0, columnName="", cellNumber=0, setText=""): #*
    """
    Writes the given text to the desired column/cell number for the given excel file
    """
    return loader.excel_set_single_cell(excel_path, sheet_name, header, columnName, cellNumber, setText)

def excel_get_single_cell(excel_path="",sheet_name="Sheet1",header=0, columnName="",cellNumber=0): #*
    """
    Gets the text from the desired column/cell number of the given excel file
    """
    return loader.excel_get_single_cell(excel_path,sheet_name,header, columnName,cellNumber)

def excel_remove_duplicates(excel_path="",sheet_name="Sheet1", header=0, columnName="", saveResultsInSameExcel=True, which_one_to_keep="first"): #*
    """
    Drops the duplicates from the desired Column of the given excel file
    """
    return loader.excel_remove_duplicates(excel_path,sheet_name, header, columnName, saveResultsInSameExcel, which_one_to_keep)

def excel_vlook_up(filepath_1="", sheet_name_1 = 'Sheet1', header_1 = 0, filepath_2="", sheet_name_2 = 'Sheet1', header_2 = 0, Output_path="", OutputExcelFileName="", match_column_name="",how='left'):#*
    """
    Performs excel_vlook_up on the given excel files for the desired columns. Possible values for how are "inner","left", "right", "outer"
    """
    return loader.excel_vlook_up(filepath_1, sheet_name_1, header_1, filepath_2, sheet_name_2, header_2, Output_path, OutputExcelFileName, match_column_name,how)

def excel_change_corrupt_xls_to_xlsx(xls_file ='',xlsx_file = '', xls_sheet_name=''): 
    '''
        Repair corrupt file to regular file and then convert it to xlsx.
        status : Done.
    '''
    return loader.excel_change_corrupt_xls_to_xlsx(xls_file,xlsx_file, xls_sheet_name)

def excel_convert_xls_to_xlsx(xls_file_path='',xlsx_file_path=''):
    """
    Converts given XLS file to XLSX
    """
    return loader.excel_convert_xls_to_xlsx(xls_file_path,xlsx_file_path)

def excel_apply_format_as_table(excel_file_path,table_style="TableStyleMedium21",sheet_name='Sheet1'): # range : "A1:AA"
    '''
        Applies table format to the used range of the given excel.
        Just it takes an path and converts it to table here you can change the table style below.
        if you want to change the table style just change the styles by refering excel
    '''
    loader.excel_apply_format_as_table(excel_file_path,table_style,sheet_name)

def excel_split_on_user_defined_conditions(excel_file_path,sheet_name='Sheet1',column_name='',condition_strings=None,output_dir=''):
    '''
        Splits the excel based on user defined row/column conditions
        Just give the column name and row condition which you want split your  excel.
        Give one string or if more conditions the  pass as list it will split the excel based on those conditions and save  them 
        in the given output directory.
        Here if output dir is not there it will create output dir in current directory and save all excels there. 
        If you want unique rows data in different excel files simply don't pass any thing in condition strings
    '''
    loader.excel_split_on_user_defined_conditions(excel_file_path,sheet_name,column_name,condition_strings,output_dir)

def excel_convert_to_image(excel_file_path=""):
    """
    Returns an Image (PNG) of given Excel
    """
    return loader.excel_convert_to_image(excel_file_path)

def excel_create_excel_file_in_given_folder(fullPathToTheFolder="",excelFileName="",sheet_name="Sheet1"):
    """
    Creates an excel file in the desired folder with desired filename

    Internally this uses folder_create() method to create folders if the folder/s does not exist.

    Parameters:
        fullPathToTheFolder (str) : Complete path to the folder with double slashes.
        excelFileName       (str) : File Name of the excel to be created (.xlsx extension will be added automatically.
        sheet_name           (str) : By default it will be "Sheet1".
    
    Returns:
        returns boolean TRUE if the excel file is created
    """
    return loader.excel_create_excel_file_in_given_folder(fullPathToTheFolder,excelFileName,sheet_name)

def excel_if_value_exists(excel_path="",sheet_name='Sheet1',header=0,usecols="",value=""):
    """
    Check if a given value exists in given excel. Returns True / False
    """
    return loader.excel_if_value_exists(excel_path,sheet_name,header,usecols,value)

def excel_create_file(fullPathToTheFile="",fileName="",sheet_name="Sheet1"):
    """
        Create a Excel file in fullPathToTheFile with filename.
    """
    return loader.excel_create_file(fullPathToTheFile,fileName,sheet_name)

def excel_to_colored_html(formatted_excel_path=""):
    """
    Converts given Excel to HTML preserving the Excel format and saves in same folder as .html
    """
    return loader.excel_to_colored_html(formatted_excel_path)

def excel_get_all_sheet_names(excelFilePath=""):
    """
    Gives you all names of the sheets in the given excel sheet.

    Parameters:
        excelFilePath  (str) : Full path to the excel file with slashes.
    
    returns : 
        all the names of the excelsheets as a LIST.
    """
    return loader.excel_get_all_sheet_names(excelFilePath)

def excel_get_all_header_columns(excel_path="",sheet_name="Sheet1",header=0):
    """
    Gives you all column header names of the given excel sheet.
    """
    return loader.excel_get_all_header_columns(excel_path,sheet_name,header)

def excel_describe_data(excel_path="",sheet_name='Sheet1',header=0):
    """
    Describe statistical data for the given excel
    """
    return loader.excel_describe_data(excel_path,sheet_name,header)

def excel_sub_routines():
    """
    Excel VBA Macros called from ClointFusion
    """
    loader.excel_sub_routines()

def convert_csv_to_excel(csv_path="",sep=""):
    """
    Function to convert CSV to Excel 

    Ex: convert_csv_to_excel()
    """
    loader.convert_csv_to_excel(csv_path,sep)

def isNaN(value):
    """
    Returns TRUE if a given value is NaN False otherwise
    """
    return loader.isNaN(value)

# ---------  Excel Functions Ends --------- 



# --------- Windows Objects Functions ---------

def win_obj_open_app(title,program_path_with_name,file_path_with_name="",backend='uia'):  
    """
    Open any windows application
    Parameters : 
        Title - Title of the application window.
        program_path_with_name - The full path of the application
        file_path_with_name - The full path to the file (only if required ex: to open an already saved excel file)
    """
    return loader.win_obj_open_app(title,program_path_with_name,file_path_with_name,backend)
    
def win_obj_get_all_objects(main_dlg,save=False,file_name_with_path=""):
    """
    Print or Save all the windows object elements of an application.
    Parameters : 
        main_dlg - Main Dialogue Handle returned from OpenApp_w() function.
        save - True if you want to save.
        file_name_with_path - new txt file name with path if you want to save)
    """
    loader.win_obj_get_all_objects(main_dlg,save,file_name_with_path)

def win_obj_mouse_click(main_dlg,title="", auto_id="", control_type=""):
    """
    Simulate high level mouse clicks on windows object elements.
    Parameters : 
        main_dlg - Main Dialogue Handle returned from OpenApp_w() function.
        title - Title of the application window.
        auto_id - Automation ID of the windows object element.
        control_type - Control type of the windows object element.
    """
    loader.win_obj_mouse_click(main_dlg,title, auto_id, control_type)

def win_obj_key_press(main_dlg,write,title="", auto_id="", control_type=""):
    """
    Simulate high level Keypress on windows object elements.
    Parameters : 
        main_dlg - Main Dialogue Handle returned from OpenApp_w() function.
        write - text to write.
        title - Title of the application window.
        auto_id - Automation ID of the windows object element.
        control_type - Control type of the windows object element.
    """
    loader.win_obj_key_press(main_dlg,write,title, auto_id, control_type)

def win_obj_get_text(main_dlg,title="", auto_id="", control_type="", value = False):
    """
    Read text from windows object element.
    Parameters : 
        main_dlg - Main Dialogue Handle returned from OpenApp_w() function.
        title - Title of the application window.
        auto_id - Automation ID of the windows object element.
        control_type - Control type of the windows object element.
        Value - True to read  a set of text and false to read another set of text for the same windows object element.
    """
    return loader.win_obj_get_text(main_dlg,title, auto_id, control_type, value)

# --------- Windows Objects Functions Ends ---------



# --------- Screenscraping Functions ---------

def scrape_save_contents_to_notepad(folderPathToSaveTheNotepad="",switch_to_window="",X=0,Y=0):  #"Full path to the folder (with double slashes) where notepad is to be stored"
    """
    Copy pastes all the available text on the screen to notepad and saves it.
    """
    return loader.scrape_save_contents_to_notepad(folderPathToSaveTheNotepad,switch_to_window,X,Y)

def scrape_get_contents_by_search_copy_paste(highlightText=""):
    """
    Gets the focus on the screen by searching given text using crtl+f and performs copy/paste of all data. Useful in Citrix applications
    This is useful in Citrix applications
    """
    return loader.scrape_get_contents_by_search_copy_paste(highlightText)

def screen_clear_search(delay=0.2):
    """
    Clears previously found text (crtl+f highlight)
    """
    loader.screen_clear_search(delay)

def search_highlight_tab_enter_open(searchText="",hitEnterKey="Yes",shift_tab='No'):
    """
    Searches for a text on screen using crtl+f and hits enter.
    This function is useful in Citrix environment
    """
    loader.search_highlight_tab_enter_open(searchText,hitEnterKey,shift_tab)
  
def find_text_on_screen(searchText="",delay=0.1, occurance=1,isSearchToBeCleared=False):
    """
    Clears previous search and finds the provided text on screen.
    """
    loader.find_text_on_screen(searchText,delay, occurance,isSearchToBeCleared)

# --------- Screenscraping Functions Ends ---------



# --------- Schedule Functions ---------

def schedule_create_task_windows(Weekly_Daily="D",week_day="Sun",start_time_hh_mm_24_hr_frmt="11:00"):#*
    """
    Schedules (weekly & daily options as of now) the current BOT (.bat) using Windows Task Scheduler. Please call create_batch_file() function before using this function to convert .pyw file to .bat
    """
    loader.schedule_create_task_windows(Weekly_Daily,week_day,start_time_hh_mm_24_hr_frmt)

def schedule_delete_task_windows():
    """
    Deletes already scheduled task. Asks user to supply task_name used during scheduling the task. You can also perform this action from Windows Task Scheduler.
    """
    loader.schedule_delete_task_windows()

# --------- Schedule Functions Ends ---------



# --------- Email Functions ---------

def email_send_via_desktop_outlook(toAddress="",ccAddress="",subject="",htmlBody="",embedImgPath="",attachmentFilePath=""):
    """
    Send email using Outlook from Desktop email application
    """
    loader.email_send_via_desktop_outlook(toAddress,ccAddress,subject,htmlBody,embedImgPath,attachmentFilePath)

# --------- Email Functions Ends ---------



# --------- Utility Functions ---------

def find(function_partial_name=""):
    # Find and inspect python functions
    loader.find(function_partial_name)

def pause_program(seconds="5"):
    """
    Stops the program for given seconds
    """
    loader.pause_program(seconds)

def show_emoji(strInput=""):
    """
    Function which prints Emojis

    Usage: 
    print(show_emoji('thumbsup'))
    print("OK",show_emoji('thumbsup'))
    Default: thumbsup
    """
    return loader.show_emoji(strInput)

def create_batch_file(application_exe_pyw_file_path=""):
    """
    Creates .bat file for the given application / exe or even .pyw BOT developed by you. This is required in Task Scheduler.
    """
    return loader.create_batch_file(application_exe_pyw_file_path)

def dismantle_code(strFunctionName=""):
    """
    This functions dis-assembles given function and shows you column-by-column summary to explain the output of disassembled bytecode.

    Ex: dismantle_code(show_emoji)
    """
    return loader.dismantle_code(strFunctionName)

def download_this_file(url=""):
    """
    Downloads a given url file to BOT output folder or Browser's Download folder
    """
    return loader.download_this_file(url)

def clear_screen():
    """
    Clears Python Interpreter Terminal Window Screen
    """
    loader.clear_screen()            

def print_with_magic_color(strMsg="", magic=False):
    """
    Prints the message with colored foreground font
    """
    loader.print_with_magic_color(strMsg, magic)

def ocr_now(img_path=""):
    """
    Recognize and “read” the text embedded in images using Google's Tesseract-OCR
    """
    loader.ocr_now(img_path)

# --------- Utility Functions Ends ---------



# --------- Voice Interface ---------

def text_to_speech(audio, show=True, rate=170):
    """
    Text to Speech using Google's Generic API
    Rate is the speed of speech. Default is 150
    Actual default : 200
    """
    return loader.text_to_speech(audio, show, rate)

def speech_to_text():
    """
    Speech to Text using Google's Generic API
    """
    return loader.speech_to_text()

# --------- Voice Interface Ends ---------



# --------- Self-Test and ClointFusion Related Functions | Current Count : 3

def update_log_excel_file(message=""):
    """
    Given message will be updated in the excel log file.

    Parameters:
        message  (str) : message to update.

    Retursn:
        returns a boolean true if updated sucessfully
    """
    return loader.update_log_excel_file(message)

def OFF_semi_automatic_mode():
    """
    This function sets semi_automatic_mode as False => OFF
    """
    loader.OFF_semi_automatic_mode()     

def ON_semi_automatic_mode():
    """
    This function sets semi_automatic_mode as True => ON
    """
    loader.ON_semi_automatic_mode()

# --------- Self-Test and ClointFusion Related Functions | Current Count : 3
